"""PostgreSQL storage, Excel import and report queries for the Matalan carton tracker."""
import hashlib
import json
import math
import os
import pickle
import re
from datetime import date, datetime, timedelta

import openpyxl
import psycopg

TABLES = {"STD": "track_std", "CUS": "track_cus"}
SHEETS = {"STD": "TRACK_STD", "CUS": "TRACK_CUS"}

# Database column -> header text in the tracker sheets, in sheet order.
HEADERS = {
    "po_received_date": "PO Received Date",
    "shipment_date": "Shipment Date",
    "matalan_po": "Matalan Product PO#",
    "supplier_name": "Supplier Name",
    "factory_name": "Factory Name",
    "packaging_supplier": "Packaging Supplier Name",
    "ebiz_customer_ref": "eBiz Customer Reference #",
    "ebiz_so": "eBiz Sales Order #",
    "carton_qty": "Carton QTY",
    "carton_style": "Carton Style",
    "flute": "FLUTE",
    "carton_selection": "Carton Selection",
    "length_mm": "Length(mm)",
    "width_mm": "Width(mm)",
    "height_mm": "Height(mm)",
    "l_mm": "L",
    "w_mm": "W",
    "area_sqm": "Area(sqm)",
    "unit_price": "Carton's Unit Price (USD,EXW w/VAT)",
    "final_price": "Final Price, USD/CTN",
    "total_amount": "Total Amount",
    "note": "NOTE",
}
COLUMNS = list(HEADERS)
# Supplier workbooks still label the prices CNY although they are USD; import accepts both spellings.
OLD_HEADERS = {"unit_price": "Carton's Unit Price (CNY,EXW w/VAT)", "final_price": "Final Price, CNY/CTN"}
DATE_COLS = {"po_received_date", "shipment_date"}
INT_COLS = {"carton_qty"}
REAL_COLS = {"length_mm", "width_mm", "height_mm", "l_mm", "w_mm", "area_sqm",
             "unit_price", "final_price", "total_amount"}
# Columns that identify a tracker row across imports (prices are formulas, so left out).
KEY_COLS = [c for c in COLUMNS if c not in {"area_sqm", "unit_price", "final_price", "total_amount"}]

FIRST_DATE = date(2020, 1, 1)  # earlier dates are typos
LEAD_DAYS = 60  # a day/month swap is suggested only if it puts the PO at most this many days before shipment

# eBiz export column -> header text (first sheet of each vendor's export).
EBIZ_HEADERS = {
    "vendor": "Vendor",
    "vendor_code": "Vendor Code",
    "product": "Product",
    "order_number": "Order Number",
    "qty_shipped": "Qty Shipped",
    "net_price_1000": "Current Net Price per 1000",
    "gross_price_1000": "Current Gross Price per 1000",
    "total_rebate": "Total Rebate",
    "currency": "Currency",
}
EBIZ_REQUIRED = ["vendor", "order_number", "qty_shipped", "total_rebate"]
EBIZ_REAL = {"qty_shipped", "net_price_1000", "gross_price_1000", "total_rebate"}
EBIZ_COLUMNS = list(EBIZ_HEADERS) + ["so_1", "so_2", "vendor_key"]

SEED_RATES = [
    ("STD", "2025-12-01", 0.04, 0.05),
    ("CUS", "2025-12-01", 0.075, 0.10),
    ("STD", "2026-07-01", 0.04, 0.05),
    ("CUS", "2026-07-01", 0.075, 0.05),
]

# Carton Selection -> (Length, Width, Height) in mm, as the tracker workbook's formulas fill them.
# Ratio cartons have a fixed footprint; their Height is typed per row. "Custom" sizes are always typed.
SIZES = {
    "STD": {f"{letter}{n}": (495 if letter == "A" else 595, {"C": 325, "D": 380}.get(letter, 285),
                             {1: 375, 2: 310}.get(n, 260)) for letter in "ABCD" for n in (1, 2, 3)},
    "CUS": {"R-FPA": (495, 285, None), "R-FPB": (595, 285, None),
            "R-FPC": (595, 325, None), "R-FPD": (595, 380, None)},
}
# Blank size: L = (Length + Width) * 2 + a, W = Width + Height + b; Area = L * W * factor / 1e6.
BLANK = {"STD": (11, -16, 1.0), "CUS": (15, -10, 1.025)}
CALC_COLS = ["length_mm", "width_mm", "height_mm", "l_mm", "w_mm", "area_sqm",
             "unit_price", "final_price", "total_amount"]
CALC_TOLERANCE = 0.01  # a saved value this close to the calculation is a match (rounding)
# Shipments before this date (up to Q3 2026) keep their values: blanks are filled, nothing is flagged.
CHECK_FROM = "2026-10-01"
# (carton type, Carton Selection, from, USD price, per "carton" or "sqm"), from the RFQ sheet of the tracker.
SEED_PRICES = [("STD", sel, "2025-12-01", price, "carton") for sel, price in zip(
    SIZES["STD"], [0.97, 0.88, 0.81, 1.10, 0.99, 0.90, 1.22, 1.10, 1.02, 1.39, 1.27, 1.18])] + [
    ("STD", "Custom", "2025-12-01", 0.96, "sqm")] + [
    ("CUS", sel, "2025-12-01", 1.03, "sqm") for sel in [*SIZES["CUS"], "Custom Carton"]]

VIEW_SQL = """
CREATE OR REPLACE VIEW {view} AS
SELECT t.*, r.pacd_pct, r.matalan_pct,
  t.unit_price * t.carton_qty * r.matalan_pct AS matalan_rebate,
  (t.total_amount - t.unit_price * t.carton_qty * r.matalan_pct) * r.pacd_pct AS pacd,
  EXTRACT(YEAR FROM t.shipment_date::date)::int
    + (date_trunc('month', t.shipment_date::date) = DATE '2025-12-01')::int AS fiscal_year,
  CASE
    WHEN t.shipment_date IS NULL THEN NULL
    WHEN date_trunc('month', t.shipment_date::date) = DATE '2025-12-01' THEN 1
    ELSE (EXTRACT(MONTH FROM t.shipment_date::date)::int + 2) / 3
  END AS quarter
FROM {table} t
LEFT JOIN rates r ON r.carton_type = '{kind}' AND r.effective_from = (
  SELECT MAX(effective_from) FROM rates
  WHERE carton_type = '{kind}' AND effective_from::date <= t.shipment_date::date);
"""


def _col_type(col):
    return "INTEGER" if col in INT_COLS else "DOUBLE PRECISION" if col in REAL_COLS else "TEXT"


def _schema():
    cols = ",\n  ".join(f"{c} {_col_type(c)}" for c in COLUMNS)
    sql = [f"CREATE TABLE IF NOT EXISTS {t} (\n  id BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,\n  {cols},\n"
           f"  supplier_key TEXT,\n  source_file TEXT,\n  imported_at TEXT\n);" for t in TABLES.values()]
    sql.append("""CREATE TABLE IF NOT EXISTS rates (
  carton_type TEXT NOT NULL CHECK (carton_type IN ('STD', 'CUS')),
  effective_from TEXT NOT NULL,
  pacd_pct DOUBLE PRECISION NOT NULL,
  matalan_pct DOUBLE PRECISION NOT NULL,
  PRIMARY KEY (carton_type, effective_from)
);""")
    sql.append("""CREATE TABLE IF NOT EXISTS carton_prices (
  carton_type TEXT NOT NULL CHECK (carton_type IN ('STD', 'CUS')),
  carton_selection TEXT NOT NULL,
  effective_from TEXT NOT NULL,
  price DOUBLE PRECISION NOT NULL,
  per TEXT NOT NULL CHECK (per IN ('carton', 'sqm')),
  PRIMARY KEY (carton_type, carton_selection, effective_from)
);""")
    sql.append("""CREATE TABLE IF NOT EXISTS date_fixes (
  row_key TEXT PRIMARY KEY,
  carton_type TEXT NOT NULL,
  packaging_supplier TEXT,
  ebiz_so TEXT,
  carton_qty INTEGER,
  file_po TEXT,
  file_ship TEXT,
  po_received_date TEXT,
  shipment_date TEXT,
  saved_at TEXT NOT NULL
);""")
    ebiz_cols = ",\n  ".join(f"{c} {'DOUBLE PRECISION' if c in EBIZ_REAL else 'TEXT'}" for c in EBIZ_COLUMNS)
    sql.append(f"CREATE TABLE IF NOT EXISTS ebiz (\n  id BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY,\n  {ebiz_cols},\n"
               f"  source_file TEXT,\n  imported_at TEXT\n);")
    sql.append("""CREATE TABLE IF NOT EXISTS persistent_cache (
  cache_key TEXT PRIMARY KEY,
  namespace TEXT NOT NULL,
  payload BYTEA NOT NULL,
  created_at TEXT NOT NULL,
  accessed_at TEXT NOT NULL
);""")
    sql.append("""CREATE TABLE IF NOT EXISTS data_revision (
  singleton INTEGER PRIMARY KEY CHECK (singleton = 1),
  revision INTEGER NOT NULL
);""")
    for kind, table in TABLES.items():
        view = f"v_{kind.lower()}"
        sql.append(VIEW_SQL.format(view=view, table=table, kind=kind))
    return "\n".join(sql)


class Row(tuple):
    """A result row read by position, by column name, or as a dict via dict(row)."""

    def __new__(cls, names, values):
        row = super().__new__(cls, values)
        row._names = names
        return row

    def keys(self):
        return self._names

    def __getitem__(self, key):
        return super().__getitem__(self._names.index(key) if isinstance(key, str) else key)


def _row_factory(cursor):
    names = [c.name for c in cursor.description or ()]
    return lambda values: Row(names, values)


_ready = set()  # database URLs whose schema this process has created


def connect(url=None):
    """Open an autocommit connection; create the schema and seed rates once per process.

    URL: the argument, else env DATABASE_URL. Writes go in `with con.transaction():`.
    """
    url = url or os.environ.get("DATABASE_URL")
    if not url:
        raise RuntimeError("DATABASE_URL is not set.")
    timeout = {} if "connect_timeout" in url else {"connect_timeout": 10}  # a silent server fails fast
    con = psycopg.connect(url, autocommit=True, row_factory=_row_factory, **timeout)
    if url not in _ready:
        with con.transaction():
            con.execute(_schema())
            con.execute("INSERT INTO data_revision VALUES (1, 0) ON CONFLICT (singleton) DO NOTHING")
            if not con.execute("SELECT 1 FROM rates LIMIT 1").fetchone():
                con.cursor().executemany("INSERT INTO rates VALUES (%s, %s, %s, %s)", SEED_RATES)
            if not con.execute("SELECT 1 FROM carton_prices LIMIT 1").fetchone():
                con.cursor().executemany("INSERT INTO carton_prices VALUES (%s, %s, %s, %s, %s)", SEED_PRICES)
        fill_saved(con)  # rows saved before the calculation existed get their blank cells filled
        _ready.add(url)
    return con


CACHE_VERSION = 2  # bump when report content changes for the same data (2: calendar quarters)


def cache_key(namespace, *parts):
    """Stable key for an application-owned persistent cache entry."""
    digest = hashlib.sha256(f"{CACHE_VERSION}:{namespace}".encode())
    for part in parts:
        raw = (bytes(part) if isinstance(part, (bytes, bytearray, memoryview)) else
               json.dumps(part, sort_keys=True, default=str,
                          separators=(",", ":")).encode())
        digest.update(len(raw).to_bytes(8, "big"))
        digest.update(raw)
    return f"{namespace}:{digest.hexdigest()}"


def cache_get(con, key):
    """Return a trusted application cache value, or None when absent/corrupt."""
    row = con.execute("SELECT payload FROM persistent_cache WHERE cache_key = %s", (key,)).fetchone()
    if not row:
        return None
    try:
        value = pickle.loads(row[0])
    except (pickle.PickleError, EOFError, AttributeError, ValueError, TypeError):
        with con.transaction():
            con.execute("DELETE FROM persistent_cache WHERE cache_key = %s", (key,))
        return None
    with con.transaction():
        con.execute("UPDATE persistent_cache SET accessed_at = %s WHERE cache_key = %s",
                    (datetime.now().isoformat(timespec="seconds"), key))
    return value


def cache_set(con, key, namespace, value, max_entries=30):
    """Persist an application-produced value and trim old namespace entries."""
    now = datetime.now().isoformat(timespec="seconds")
    payload = pickle.dumps(value, protocol=pickle.HIGHEST_PROTOCOL)
    with con.transaction():
        con.execute(
            "INSERT INTO persistent_cache VALUES (%s, %s, %s, %s, %s) "
            "ON CONFLICT (cache_key) DO UPDATE SET namespace = EXCLUDED.namespace, "
            "payload = EXCLUDED.payload, created_at = EXCLUDED.created_at, "
            "accessed_at = EXCLUDED.accessed_at",
            (key, namespace, payload, now, now))
        con.execute(
            "DELETE FROM persistent_cache WHERE cache_key IN ("
            "SELECT cache_key FROM persistent_cache WHERE namespace = %s "
            "ORDER BY accessed_at DESC OFFSET %s)",
            (namespace, int(max_entries)),
        )


def _bump_revision(con):
    con.execute("UPDATE data_revision SET revision = revision + 1 WHERE singleton = 1")


def data_revision(con):
    """A compact token that changes after tracker, eBiz, or rate data changes."""
    revision = con.execute(
        "SELECT revision FROM data_revision WHERE singleton = 1").fetchone()[0]
    values = []
    for table in (*TABLES.values(), "ebiz"):
        values.extend(con.execute(
            f"SELECT COUNT(*), COALESCE(MAX(imported_at), '') FROM {table}"
        ).fetchone())
    values.extend(con.execute(
        "SELECT COUNT(*), COALESCE(MAX(effective_from || ':' || pacd_pct::text || ':' || matalan_pct::text), '') "
        "FROM rates"
    ).fetchone())
    return f"{revision}-" + hashlib.sha256(repr(tuple(values)).encode()).hexdigest()[:16]


EXCEL_ERRORS = {"#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!"}


class ImportFileError(ValueError):
    """The workbook can't be imported. The message says why."""


def _norm(header):
    return re.sub(r"\s+", " ", str(header or "")).strip().lower()


def _clean(value):
    if isinstance(value, float) and math.isnan(value):
        return None
    if isinstance(value, str):
        value = value.strip()
        if not value or value in EXCEL_ERRORS:
            return None
    return value


_PUNCT_TO_SPACE = str.maketrans('.,;:"\'!?', " " * 8)


def supplier_key(name):
    """Upper-cased first word of a packaging supplier name, so its spellings merge (port of the VBA FirstWord).

    'M&U  Packaging Ltd.' -> 'M&U'. None when there is no word.
    """
    if name is None:
        return None
    words = name.translate(_PUNCT_TO_SPACE).split()
    return words[0].upper() if words else None


_TEXT_DATE = re.compile(r"(\d{1,2})[/.-](\d{1,2})[/.-](\d{4}|\d{2})")


def _text_date(text):
    """'18/12/2025' -> date(2025, 12, 18). Text dates are read day first, as the factories type them."""
    m = _TEXT_DATE.fullmatch(text.strip())
    if not m:
        return None
    day, month, year = map(int, m.groups())
    try:
        return date(year + 2000 if year < 100 else year, month, day)
    except ValueError:
        return None


def _convert(col, value):
    if value is None:
        return None
    if col in DATE_COLS:
        if isinstance(value, datetime):
            return value.date().isoformat()
        if isinstance(value, date):
            return value.isoformat()
        d = _text_date(value) if isinstance(value, str) else None
        return d.isoformat() if d else None
    if col in INT_COLS or col in REAL_COLS:
        try:
            number = float(value)
        except (TypeError, ValueError):
            return None
        return int(round(number)) if col in INT_COLS else number
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value)


def _shown(value):
    """A date cell as the user sees it: a date as 07-Dec-2025 (no day/month doubt), text as typed."""
    if isinstance(value, datetime):
        value = value.date()
    if isinstance(value, date):
        return f"{value:%d-%b-%Y}"
    return None if value is None else str(value)


def _row_key(kind, raw, row):
    """Hash of the row's input cells, so the same tracker row has the same key in next month's file."""
    ident = [kind] + [_shown(raw[c]) if c in DATE_COLS else row[c] for c in KEY_COLS]
    return hashlib.sha1(json.dumps(ident).encode()).hexdigest()


def _open(src):
    try:
        return openpyxl.load_workbook(src, data_only=True)
    except Exception as e:  # not an xlsx, corrupt or password protected
        raise ImportFileError(f"Can't open the workbook: {e}") from e


def _header_row(ws, marker, where):
    """({normalised header: column index}, row number) of the first of rows 1-10 with a `marker` cell."""
    for n, header_row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
        found = {_norm(h): i for i, h in enumerate(header_row)}
        if _norm(marker) in found:
            return found, n
    raise ImportFileError(f"{where} header row not found (no '{marker}' cell in rows 1-10).")


def _find_columns(sheet, ws):
    """Return (column index per COLUMNS, header row number).

    The header row is the first of rows 1-10 with a 'Shipment Date' cell: row 2 in the
    master tracker (title above), row 1 in the per-supplier exports.
    """
    found, n = _header_row(ws, HEADERS["shipment_date"], f"{sheet}:")
    index = {}
    for col in COLUMNS:
        index[col] = found.get(_norm(HEADERS[col]), found.get(_norm(OLD_HEADERS.get(col, HEADERS[col]))))
        if index[col] is None and col == "unit_price":
            index[col] = next((i for h, i in found.items() if h.startswith("carton's unit price")), None)
        if index[col] is None:
            raise ImportFileError(f"{sheet}: column '{HEADERS[col]}' not found in the header row (row {n}).")
    return index, n


def read_workbook(src):
    """Read the data rows of TRACK_STD and TRACK_CUS from a path or binary file object.

    Returns {"STD": [...], "CUS": [...]}. Each row is a dict of COLUMNS plus "supplier_key",
    "_row" (Excel row number), "_file" (the two date cells as they are in the file) and
    "_key" (identifies the row in later imports, for saved date corrections).
    """
    wb = _open(src)
    names = {n.upper(): n for n in wb.sheetnames}
    for sheet in SHEETS.values():
        if sheet not in names:
            raise ImportFileError(f"Sheet '{sheet}' not found.")
    data = {}
    for kind, sheet in SHEETS.items():
        ws = wb[names[sheet]]
        index, header = _find_columns(sheet, ws)
        rows = []
        for n, values in enumerate(ws.iter_rows(min_row=header + 1, values_only=True), start=header + 1):
            raw = {c: _clean(values[i]) if i < len(values) else None for c, i in index.items()}
            if raw["shipment_date"] is None and raw["carton_qty"] is None:
                continue  # template row: formulas but no input
            row = {c: _convert(c, v) for c, v in raw.items()}
            row["supplier_key"] = supplier_key(row["packaging_supplier"])
            row["_row"] = n
            row["_file"] = {c: raw[c] for c in DATE_COLS}
            row["_key"] = _row_key(kind, raw, row)
            rows.append(row)
        if rows and all(r["total_amount"] is None for r in rows):
            raise ImportFileError(f"{sheet}: the file has no calculated values. "
                                  "Open it in Excel, save, and import again.")
        data[kind] = rows
    return data


def _month(row):
    return row["shipment_date"][:7] if row["shipment_date"] else None


_SAME_KEY = "supplier_key IS NOT DISTINCT FROM %s AND substr(shipment_date, 1, 7) IS NOT DISTINCT FROM %s"


def preview(con, data):
    """Rows per sheet, supplier key and shipment month, and how many DB rows each replaces."""
    out = []
    for kind, rows in data.items():
        groups = {}
        for r in rows:
            groups.setdefault((r["supplier_key"], _month(r)), []).append(r["packaging_supplier"])
        for (key, month), names in sorted(groups.items(), key=lambda g: (str(g[0][0]), str(g[0][1]))):
            replaced = con.execute(f"SELECT COUNT(*) FROM {TABLES[kind]} WHERE {_SAME_KEY}",
                                   (key, month)).fetchone()[0]
            out.append({"Sheet": SHEETS[kind], "Supplier": key,
                        "Spellings in file": " / ".join(sorted(set(names) - {None})),
                        "Shipment month": month or "no valid date", "Rows in file": len(names),
                        "Rows replaced in DB": replaced})
    return out


def _swap(d):
    """d with day and month exchanged, or None if that is not a date."""
    try:
        return d.replace(month=d.day, day=d.month)
    except ValueError:
        return None


def _lead_ok(po, ship):
    return po <= ship <= po + timedelta(days=LEAD_DAYS)


def _with_ship_year(po, ship):
    """po moved to the shipment's year (or the year before) when that puts it shortly before shipment."""
    if ship is None:
        return None
    for year in (ship.year, ship.year - 1):
        try:
            fixed = po.replace(year=year)
        except ValueError:  # 29 February
            continue
        if _lead_ok(fixed, ship):
            return fixed
    return None


def _check(row, today):
    """(problems, suggested PO Received Date, suggested Shipment Date) for one row; no problems if it looks right.

    Factories type dates day first; on a month-first PC Excel stores 06/12 as 12 June and keeps 18/12 as text.
    """
    problems, dates, valid = [], {}, {}
    for col in ("shipment_date", "po_received_date"):
        label, raw = HEADERS[col], row["_file"][col]
        d = date.fromisoformat(row[col]) if row[col] else None
        left_out = "; the row is left out of reports" if col == "shipment_date" else ""
        if raw is None:
            if col == "shipment_date":
                problems.append(f"{label} is empty{left_out}")
        elif d is None:
            problems.append(f"{label} is not a date{left_out}")
        else:
            if isinstance(raw, str):
                problems.append(f"{label} typed as text, read as day/month/year")
            if not FIRST_DATE <= d <= today:
                where = "in the future" if d > today else f"before {FIRST_DATE.year}"
                swapped = _swap(d)
                year_fixed = (_with_ship_year(d, dates["shipment_date"] if valid["shipment_date"] else None)
                              if col == "po_received_date" and d < FIRST_DATE else None)
                if swapped and FIRST_DATE <= swapped <= today:
                    problems.append(f"{label} is {where}; day and month look swapped")
                    d = swapped
                elif year_fixed:
                    problems.append(f"{label} is {where}; the year looks wrong")
                    d = year_fixed
                else:
                    problems.append(f"{label} is {where}")
        dates[col] = d
        valid[col] = d is not None and FIRST_DATE <= d <= today
    po, ship = dates["po_received_date"], dates["shipment_date"]
    if valid["po_received_date"] and valid["shipment_date"] and not _lead_ok(po, ship):
        po_swapped, ship_swapped = _swap(po), _swap(ship)
        if po_swapped and _lead_ok(po_swapped, ship):
            problems.append("PO Received Date: day and month look swapped")
            po = po_swapped
        elif po > ship:
            if ship_swapped and ship_swapped <= today and _lead_ok(po, ship_swapped):
                problems.append("Shipment Date: day and month look swapped")
                ship = ship_swapped
            else:
                problems.append("PO Received Date is after Shipment Date")
    return problems, po, ship


def check_dates(con, files, today=None):
    """Apply saved date corrections, then list the rows whose dates still look wrong.

    `files` is {file name: read_workbook()}. Rows with a saved correction get its dates in place and
    are not listed. Returns (rows corrected from saved corrections, issues). An issue is a dict of
    display fields, the suggested "PO Received Date" and "Shipment Date" (date or None), and "row":
    the row dict apply_dates() updates.
    """
    today = today or date.today()
    saved = {r["row_key"]: r for r in con.execute("SELECT row_key, po_received_date, shipment_date FROM date_fixes")}
    reused, issues = 0, []
    for name, data in files.items():
        for kind, rows in data.items():
            for r in rows:
                fix = saved.get(r["_key"])
                if fix:
                    r["po_received_date"], r["shipment_date"] = fix["po_received_date"], fix["shipment_date"]
                    reused += 1
                    continue
                problems, po, ship = _check(r, today)
                if problems:
                    issues.append({
                        "File": name, "Sheet": SHEETS[kind], "Excel row": r["_row"],
                        "Packaging supplier": r["packaging_supplier"], "eBiz Sales Order #": r["ebiz_so"],
                        "Carton QTY": r["carton_qty"],
                        "PO Received Date in file": _shown(r["_file"]["po_received_date"]),
                        "Shipment Date in file": _shown(r["_file"]["shipment_date"]),
                        "Problem": "; ".join(problems), "PO Received Date": po, "Shipment Date": ship, "row": r})
    return reused, issues


def apply_dates(issues, dates):
    """Give each issue's row its reviewed (PO Received Date, Shipment Date); save_many() remembers them."""
    for issue, (po, ship) in zip(issues, dates):
        row = issue["row"]
        row["po_received_date"] = po.isoformat() if po else None
        row["shipment_date"] = ship.isoformat() if ship else None
        row["_review"] = True


def saved_dates(con):
    """Saved date corrections, newest first. "key" is what forget_dates() takes."""
    def shown(iso):
        return _shown(date.fromisoformat(iso)) if iso else None
    return [{"Sheet": SHEETS[r["carton_type"]], "Packaging supplier": r["packaging_supplier"],
             "eBiz Sales Order #": r["ebiz_so"], "Carton QTY": r["carton_qty"],
             "PO Received Date in file": r["file_po"], "Shipment Date in file": r["file_ship"],
             "PO Received Date": shown(r["po_received_date"]), "Shipment Date": shown(r["shipment_date"]),
             "Saved at": r["saved_at"].replace("T", " ")[:16], "key": r["row_key"]}
            for r in con.execute("SELECT * FROM date_fixes ORDER BY saved_at DESC, row_key")]


def forget_dates(con, keys):
    """Delete saved corrections; those rows are checked again the next time they are imported."""
    with con.transaction():
        con.cursor().executemany("DELETE FROM date_fixes WHERE row_key = %s", [(k,) for k in keys])


def clashes(files):
    """Sheet + supplier + month groups found in more than one file. `files` is {file name: read_workbook()}."""
    seen = {}
    for name, data in files.items():
        for kind, rows in data.items():
            for supplier, month in {(r["supplier_key"], _month(r)) for r in rows}:
                seen.setdefault((kind, supplier, month), []).append(name)
    return [{"Sheet": SHEETS[kind], "Supplier": supplier, "Shipment month": month or "no valid date",
             "Files": " / ".join(names)}
            for (kind, supplier, month), names in sorted(seen.items(), key=str) if len(names) > 1]


_FIX_COLS = ["row_key", "carton_type", "packaging_supplier", "ebiz_so", "carton_qty", "file_po", "file_ship",
             "po_received_date", "shipment_date", "saved_at"]
_SAVE_FIX = (f"INSERT INTO date_fixes VALUES ({', '.join(['%s'] * len(_FIX_COLS))}) ON CONFLICT (row_key) DO UPDATE SET "
             + ", ".join(f"{c} = EXCLUDED.{c}" for c in _FIX_COLS[1:]))


def save_many(con, files):
    """Import several files in one transaction: all of them load, or none do.

    Per (supplier key, shipment month) in a file: delete those DB rows, insert the file's rows.
    Dates reviewed with apply_dates() are saved for the next import of the same rows.
    Raises ValueError if two files share a sheet + supplier + month. Returns {file: rows per sheet}.
    """
    clash = clashes(files)
    if clash:
        raise ValueError("Files overlap: " + "; ".join(
            f"{c['Sheet']} {c['Supplier']} {c['Shipment month']} in {c['Files']}" for c in clash))
    now = datetime.now().isoformat(timespec="seconds")
    cols = COLUMNS + ["supplier_key", "source_file", "imported_at"]
    prices = get_prices(con)
    with con.transaction():
        cur = con.cursor()
        for name, data in files.items():
            for kind, rows in data.items():
                table = TABLES[kind]
                for r in rows:
                    fill_row(kind, r, prices)
                for key in {(r["supplier_key"], _month(r)) for r in rows}:
                    con.execute(f"DELETE FROM {table} WHERE {_SAME_KEY}", key)
                cur.executemany(f"INSERT INTO {table} ({', '.join(cols)}) VALUES ({', '.join(['%s'] * len(cols))})",
                                [[r[c] for c in COLUMNS] + [r["supplier_key"], name, now] for r in rows])
                cur.executemany(_SAVE_FIX,
                                [(r["_key"], kind, r["packaging_supplier"], r["ebiz_so"], r["carton_qty"],
                                  _shown(r["_file"]["po_received_date"]), _shown(r["_file"]["shipment_date"]),
                                  r["po_received_date"], r["shipment_date"], now)
                                 for r in rows if r.get("_review")])
        _bump_revision(con)
    return {name: {kind: len(rows) for kind, rows in data.items()} for name, data in files.items()}


def _so(text):
    """An SO number as text, or None for blank and '0'."""
    text = (text or "").strip()
    return text if text and text != "0" else None


def _so_pair(order_number):
    """Order Number "A|B" -> ("A", "B"); blank or "0" parts become None."""
    parts = (order_number or "").split("|")
    return _so(parts[0]), _so(parts[1]) if len(parts) > 1 else None


def _ebiz_value(col, value):
    if value is None:
        return None
    if col in EBIZ_REAL:
        try:
            return float(value)
        except (TypeError, ValueError):
            return None
    if isinstance(value, float) and value.is_integer():
        value = int(value)  # 7124561.0 -> 7124561
    return str(value).strip()


def read_ebiz(src):
    """Rows of one eBiz export (its first sheet) as dicts of EBIZ_COLUMNS. Raises ImportFileError.

    Order Number "A|B" gives so_1 "A" and so_2 "B"; vendor_key is supplier_key(vendor). Blank rows and rows
    with no Vendor (total lines) are skipped.
    """
    ws = _open(src).worksheets[0]
    found, n = _header_row(ws, EBIZ_HEADERS["order_number"], "eBiz")
    index = {c: found.get(_norm(h)) for c, h in EBIZ_HEADERS.items()}
    for col in EBIZ_REQUIRED:
        if index[col] is None:
            raise ImportFileError(f"eBiz column '{EBIZ_HEADERS[col]}' not found in the header row (row {n}).")
    rows = []
    for values in ws.iter_rows(min_row=n + 1, values_only=True):
        row = {c: _ebiz_value(c, _clean(values[i])) if i is not None and i < len(values) else None
               for c, i in index.items()}
        if row["order_number"] is None and row["qty_shipped"] is None:
            continue
        row["vendor_key"] = supplier_key(row["vendor"])
        if row["vendor_key"] is None:  # a total/footer line: no vendor to report it under
            continue
        row["so_1"], row["so_2"] = _so_pair(row["order_number"])
        rows.append(row)
    return rows


def ebiz_preview(files):
    """Per file and vendor key: spellings, rows, Qty Shipped and Total Rebate. `files` is {name: read_ebiz()}."""
    out = []
    for name, rows in files.items():
        groups = {}
        for r in rows:
            groups.setdefault(r["vendor_key"], []).append(r)
        for key, group in sorted(groups.items(), key=lambda g: str(g[0])):
            out.append({"File": name, "Vendor": key,
                        "Spellings in file": " / ".join(sorted({r["vendor"] for r in group} - {None})),
                        "Rows": len(group), "Qty Shipped": sum(r["qty_shipped"] or 0 for r in group),
                        "Total Rebate": round(sum(r["total_rebate"] or 0 for r in group), 2)})
    return out


def ebiz_clashes(files):
    """Vendor keys found in more than one eBiz file."""
    seen = {}
    for name, rows in files.items():
        for key in {r["vendor_key"] for r in rows}:
            seen.setdefault(key, []).append(name)
    return [{"Vendor": key, "Files": " / ".join(names)}
            for key, names in sorted(seen.items(), key=lambda s: str(s[0])) if len(names) > 1]


def replace_ebiz(con, files):
    """Replace only the uploaded vendors' eBiz rows in one transaction. Returns {file: rows inserted}.

    Raises ValueError, and changes nothing, when two files share a vendor key.
    """
    clash = ebiz_clashes(files)
    if clash:
        raise ValueError("Files overlap: " + "; ".join(f"{c['Vendor']} in {c['Files']}" for c in clash))
    now = datetime.now().isoformat(timespec="seconds")
    cols = EBIZ_COLUMNS + ["source_file", "imported_at"]
    vendors = sorted({row["vendor_key"] for rows in files.values() for row in rows})
    with con.transaction():
        if vendors:
            con.execute("DELETE FROM ebiz WHERE vendor_key = ANY(%s)", (vendors,))
        for name, rows in files.items():
            con.cursor().executemany(f"INSERT INTO ebiz ({', '.join(cols)}) VALUES ({', '.join(['%s'] * len(cols))})",
                            [[r[c] for c in EBIZ_COLUMNS] + [name, now] for r in rows])
        if vendors:
            _bump_revision(con)
    return {name: len(rows) for name, rows in files.items()}


def ebiz_status(con):
    """(row count, source file names, import time) of the current eBiz data, or None when there is none."""
    count, imported_at = con.execute("SELECT COUNT(*), MAX(imported_at) FROM ebiz").fetchone()
    if not count:
        return None
    files = [r[0] for r in con.execute("SELECT DISTINCT source_file FROM ebiz ORDER BY 1")]
    return count, files, imported_at


def ebiz_vendors(con):
    """Vendor keys in the eBiz data, sorted."""
    return [r[0] for r in con.execute(
        'SELECT DISTINCT vendor_key COLLATE "C" FROM ebiz WHERE vendor_key IS NOT NULL ORDER BY 1')]


def _ebiz_where(vendor):
    return ("WHERE vendor_key IS NOT DISTINCT FROM %s", [vendor]) if vendor is not None else ("", [])


def ebiz_rows(con, vendor=None, limit=None, offset=0):
    """eBiz rows for the editor (id, EBIZ_COLUMNS, source_file, imported_at), in id order."""
    where, params = _ebiz_where(vendor)
    paging = ""
    if limit is not None:
        paging = " LIMIT %s OFFSET %s"
        params += [int(limit), int(offset)]
    return [dict(r) for r in con.execute(
        f"SELECT id, {', '.join(EBIZ_COLUMNS)}, source_file, imported_at FROM ebiz {where} ORDER BY id{paging}",
        params)]


def ebiz_row_count(con, vendor=None):
    where, params = _ebiz_where(vendor)
    return con.execute(f"SELECT COUNT(*) FROM ebiz {where}", params).fetchone()[0]


def update_ebiz_rows(con, edited):
    """Save edited eBiz rows in one transaction; vendor_key and SO_1/SO_2 are recomputed. Returns rows saved.

    Raises ValueError, saving nothing, for a blank Vendor, a non-numeric figure or an id that no longer exists.
    """
    prepared = []
    for number, item in enumerate(edited, 1):
        try:
            row_id = int(item.get("id"))
        except (TypeError, ValueError):
            raise ValueError(f"Row {number}: row ID is invalid.") from None
        values = {}
        for col, header in EBIZ_HEADERS.items():
            value = _clean(item.get(col))
            if value is not None and col in EBIZ_REAL:
                try:
                    value = float(value)
                except (TypeError, ValueError):
                    raise ValueError(f"Row {number}: {header} must be a number.") from None
            elif value is not None:
                value = _ebiz_value(col, value)
            values[col] = value
        values["vendor_key"] = supplier_key(values["vendor"])
        if values["vendor_key"] is None:
            raise ValueError(f"Row {number}: Vendor is required.")
        values["so_1"], values["so_2"] = _so_pair(values["order_number"])
        prepared.append((row_id, values))
    cols = list(EBIZ_HEADERS) + ["vendor_key", "so_1", "so_2"]
    with con.transaction():
        for number, (row_id, values) in enumerate(prepared, 1):
            cur = con.execute(f"UPDATE ebiz SET {', '.join(f'{c} = %s' for c in cols)} WHERE id = %s",
                              [values[c] for c in cols] + [row_id])
            if not cur.rowcount:
                raise ValueError(f"Row {number}: row ID {row_id} no longer exists. Refresh the page.")
        if prepared:
            _bump_revision(con)
    return len(prepared)


def delete_ebiz_rows(con, ids):
    """Delete these eBiz row ids; ids already gone are ignored. Returns rows deleted."""
    ids = [int(i) for i in ids]
    if not ids:
        return 0
    with con.transaction():
        count = _delete(con, "ebiz", "id = ANY(%s)", (ids,))
        if count:
            _bump_revision(con)
    return count


def so_order(so):
    """Sort key for SO numbers: numeric ones by value, then text, then None."""
    if so is None:
        return (2, 0, "")
    return (0, int(so), "") if so.isdecimal() else (1, 0, so)


def cross_reference(con, supplier):
    """Tracker vs eBiz for one supplier key over all its tracker rows, any date (port of the old macros).

    Returns a dict:
    - "orders": one dict per tracker SO#, sorted by SO#: so, tracker_qty, tracker_rebate (PACD + Matalan
      rebate), ebiz_qty, ebiz_rebate, via_so2 (an eBiz row matched it through SO_2) and status
      ("Not in eBiz", "Tracker > eBiz", "Tracker < eBiz" or "Match").
    - "rows": {"STD": [...], "CUS": [...]} every tracker row of the supplier (view columns plus "so", the
      SO# or None), sorted by SO#, shipment date, id.
    - "ebiz_only": the supplier's eBiz rows whose order is not a tracker SO#, with "key" (SO_1, else SO_2).
    """
    rows = {}
    for kind in TABLES:
        found = [dict(r) for r in con.execute(f"SELECT * FROM v_{kind.lower()} WHERE supplier_key IS NOT DISTINCT FROM %s",
                                              (supplier,))]
        for r in found:
            r["so"] = _so(r["ebiz_so"])
        rows[kind] = sorted(found, key=lambda r: (so_order(r["so"]), r["shipment_date"] or "", r["id"]))
    orders = {}
    for r in rows["STD"] + rows["CUS"]:
        if r["so"]:
            o = orders.setdefault(r["so"], {"so": r["so"], "tracker_qty": 0, "tracker_rebate": 0.0,
                                            "ebiz_qty": 0, "ebiz_rebate": 0.0, "via_so2": False, "_found": 0})
            o["tracker_qty"] += r["carton_qty"] or 0
            o["tracker_rebate"] += (r["pacd"] or 0) + (r["matalan_rebate"] or 0)
    ebiz_only = []
    for e in (dict(r) for r in con.execute("SELECT * FROM ebiz WHERE vendor_key IS NOT DISTINCT FROM %s ORDER BY id", (supplier,))):
        if e["so_1"] in orders:
            e["key"], via_so2 = e["so_1"], False
        elif e["so_2"] in orders:
            e["key"], via_so2 = e["so_2"], True
        else:
            e["key"] = e["so_1"] or e["so_2"]
            ebiz_only.append(e)
            continue
        o = orders[e["key"]]
        o["ebiz_qty"] += e["qty_shipped"] or 0
        o["ebiz_rebate"] += e["total_rebate"] or 0
        o["via_so2"] = o["via_so2"] or via_so2
        o["_found"] += 1
    for o in orders.values():
        found = o.pop("_found")
        o["status"] = ("Not in eBiz" if not found else "Tracker > eBiz" if o["tracker_qty"] > o["ebiz_qty"]
                       else "Tracker < eBiz" if o["tracker_qty"] < o["ebiz_qty"] else "Match")
    return {"orders": sorted(orders.values(), key=lambda o: so_order(o["so"])), "rows": rows,
            "ebiz_only": ebiz_only}


def cross_reference_rows(xref):
    """Order-level rows shared by the live view and generated Cross Reference workbook."""
    result = [{
        "so": order["so"],
        "tracker_qty": order["tracker_qty"],
        "tracker_rebate": round(order["tracker_rebate"], 2),
        "ebiz_qty": order["ebiz_qty"],
        "ebiz_rebate": round(order["ebiz_rebate"], 2),
        "qty_difference": order["tracker_qty"] - order["ebiz_qty"],
        "status": order["status"],
        "via_so2": order["via_so2"],
    } for order in xref["orders"]]
    no_so = [row for rows in xref["rows"].values() for row in rows if not row["so"]]
    if no_so:
        qty = sum(row["carton_qty"] or 0 for row in no_so)
        rebate = round(sum((row["pacd"] or 0) + (row["matalan_rebate"] or 0)
                           for row in no_so), 2)
        result.append({
            "so": "(No SO#)", "tracker_qty": qty, "tracker_rebate": rebate,
            "ebiz_qty": None, "ebiz_rebate": None, "qty_difference": qty,
            "status": None, "via_so2": False,
        })
    return result


def ebiz_only_orders(xref):
    """Aggregated eBiz orders missing from Tracker, shared by live and generated reports."""
    groups = {}
    for row in xref["ebiz_only"]:
        group = groups.setdefault(row["key"] or "(no order number)", [0, 0.0])
        group[0] += row["qty_shipped"] or 0
        group[1] += row["total_rebate"] or 0
    return [{"order": order, "ebiz_qty": values[0], "ebiz_rebate": round(values[1], 2)}
            for order, values in sorted(groups.items(), key=lambda item: so_order(item[0]))]


def get_rates(con):
    """All rates, STD first, oldest first. Percentages are fractions (0.075 = 7.5%)."""
    return [dict(r) for r in con.execute(
        "SELECT carton_type, effective_from, pacd_pct, matalan_pct FROM rates "
        "ORDER BY carton_type DESC, effective_from")]


def save_rates(con, rates):
    """Replace every rate.

    Raises ValueError, and saves nothing, if a row is invalid or a dated shipment would be left without a rate.
    """
    if not rates:
        raise ValueError("At least one rate is needed.")
    rows, seen = [], set()
    for i, r in enumerate(rates, 1):
        kind = r.get("carton_type")
        if kind not in TABLES:
            raise ValueError(f"Row {i}: carton type must be STD or CUS.")
        try:
            start = date.fromisoformat(str(r.get("effective_from"))).isoformat()
        except ValueError:
            raise ValueError(f"Row {i}: 'From' must be a date (YYYY-MM-DD).") from None
        pcts = (r.get("pacd_pct"), r.get("matalan_pct"))
        if not all(isinstance(p, (int, float)) and 0 <= p <= 1 for p in pcts):
            raise ValueError(f"Row {i}: percentages must be between 0 and 100.")
        if (kind, start) in seen:
            raise ValueError(f"Row {i}: duplicate {kind} rate for {start}.")
        seen.add((kind, start))
        rows.append((kind, start, float(pcts[0]), float(pcts[1])))
    with con.transaction():
        con.execute("DELETE FROM rates")
        con.cursor().executemany("INSERT INTO rates VALUES (%s, %s, %s, %s)", rows)
        n, first = _unrated(con)
        if n:  # raising here rolls the transaction back
            raise ValueError(f"{n} shipment(s) from {first} on would have no rate. "
                             f"Add a rate that starts on or before {first}.")
        _bump_revision(con)


def get_prices(con):
    """All carton prices, STD first, then by Carton Selection and date."""
    return [dict(r) for r in con.execute(
        "SELECT carton_type, carton_selection, effective_from, price, per FROM carton_prices "
        "ORDER BY carton_type DESC, carton_selection, effective_from")]


def save_prices(con, prices):
    """Replace every carton price, then fill saved tracker cells that are blank and can now be calculated.

    Raises ValueError, and saves nothing, if a row is invalid. Returns how many tracker rows were filled.
    """
    rows, seen = [], set()
    for i, p in enumerate(prices, 1):
        kind, sel = p.get("carton_type"), str(p.get("carton_selection") or "").strip()
        if kind not in TABLES:
            raise ValueError(f"Row {i}: carton type must be STD or CUS.")
        if not sel:
            raise ValueError(f"Row {i}: Carton Selection is empty.")
        try:
            start = date.fromisoformat(str(p.get("effective_from"))).isoformat()
        except ValueError:
            raise ValueError(f"Row {i}: 'From' must be a date (YYYY-MM-DD).") from None
        price = p.get("price")
        if not isinstance(price, (int, float)) or not price > 0:
            raise ValueError(f"Row {i}: price must be more than 0.")
        if p.get("per") not in ("carton", "sqm"):
            raise ValueError(f"Row {i}: 'Per' must be carton or sqm.")
        if (kind, sel.casefold(), start) in seen:
            raise ValueError(f"Row {i}: duplicate {kind} {sel} price for {start}.")
        seen.add((kind, sel.casefold(), start))
        rows.append((kind, sel, start, float(price), p["per"]))
    with con.transaction():
        con.execute("DELETE FROM carton_prices")
        con.cursor().executemany("INSERT INTO carton_prices VALUES (%s, %s, %s, %s, %s)", rows)
        _bump_revision(con)
    return fill_saved(con)


def _price(prices, kind, selection, shipment_date):
    """The price row in effect on the shipment date for this carton type and selection, or None."""
    found = [p for p in prices if p["carton_type"] == kind and p["carton_selection"].casefold() == selection.casefold()
             and shipment_date and p["effective_from"] <= shipment_date]
    return max(found, key=lambda p: p["effective_from"], default=None)


def calculate(kind, row, prices):
    """CALC_COLS of one tracker row by the tracker workbook's formulas; None where it can't be calculated.

    Like the workbook, each step reads the row's own earlier value when it has one, else the calculated value.
    """
    selection = str(row.get("carton_selection") or "").strip()
    size = next((v for k, v in SIZES[kind].items() if k.casefold() == selection.casefold()), (None, None, None))
    out = dict.fromkeys(CALC_COLS)
    out.update(zip(CALC_COLS[:3], size))

    def value(col):
        return row[col] if row.get(col) is not None else out[col]

    a, b, factor = BLANK[kind]
    length, width, height = value("length_mm"), value("width_mm"), value("height_mm")
    if length and width:
        out["l_mm"] = (length + width) * 2 + a
    if width and height:
        out["w_mm"] = width + height + b
    if value("l_mm") and value("w_mm"):
        out["area_sqm"] = value("l_mm") * value("w_mm") * factor / 1_000_000
    price = _price(prices, kind, selection, row.get("shipment_date"))
    if price and price["per"] == "carton":
        out["unit_price"] = price["price"]
    elif price and value("area_sqm"):
        out["unit_price"] = value("area_sqm") * price["price"]
    out["final_price"] = value("unit_price")
    if row.get("carton_qty") is not None and value("final_price") is not None:
        out["total_amount"] = row["carton_qty"] * value("final_price")
    return out


def fill_row(kind, row, prices):
    """Fill the row's blank CALC_COLS in place with the calculation.

    Returns [(column, saved value, calculated value)] for saved values that differ from the calculation by more
    than CALC_TOLERANCE, and for cells left blank because they can't be calculated (both values None).
    Rows shipped before CHECK_FROM are only filled, never listed.
    """
    calc = calculate(kind, row, prices)
    found = []
    for col in CALC_COLS:
        if row.get(col) is None:
            row[col] = calc[col]
            if calc[col] is None:
                found.append((col, None, None))
        elif calc[col] is not None and abs(row[col] - calc[col]) > CALC_TOLERANCE:
            found.append((col, row[col], calc[col]))
    return [] if (row.get("shipment_date") or "9999") < CHECK_FROM else found


def _calc_issue(kind, row, col, saved, calc, **where):
    return {**where, "Sheet": SHEETS[kind], "Packaging supplier": row["packaging_supplier"],
            "eBiz Sales Order #": row["ebiz_so"], "Carton Selection": row["carton_selection"],
            "Column": HEADERS[col], "Value": saved, "Calculated": calc,
            "Problem": "Can't calculate" if calc is None else "Differs from calculation"}


def calc_check(con, files):
    """Fill blank calculated cells of uploaded rows in place; list the cells that differ or stay blank."""
    prices, issues = get_prices(con), []
    for name, data in files.items():
        for kind, rows in data.items():
            for r in rows:
                issues += [_calc_issue(kind, r, *found, File=name, **{"Excel row": r["_row"]})
                           for found in fill_row(kind, r, prices)]
    return issues


def _saved_rows(con):
    for kind, table in TABLES.items():
        for r in con.execute(f"SELECT id, {', '.join(COLUMNS)} FROM {table} ORDER BY id"):
            yield kind, table, dict(r)


def fill_saved(con):
    """Fill blank calculated cells of the saved tracker rows. Returns how many rows changed."""
    prices, changed = get_prices(con), 0
    with con.transaction():
        for kind, table, row in list(_saved_rows(con)):
            before = {col: row[col] for col in CALC_COLS}
            fill_row(kind, row, prices)
            cols = [col for col in CALC_COLS if row[col] != before[col]]
            if cols:
                con.execute(f"UPDATE {table} SET {', '.join(f'{c} = %s' for c in cols)} WHERE id = %s",
                            [row[c] for c in cols] + [row["id"]])
                changed += 1
        if changed:
            _bump_revision(con)
    return changed


def saved_calc_issues(con):
    """Saved tracker cells that differ from the calculation or are blank because they can't be calculated."""
    prices = get_prices(con)
    return [_calc_issue(kind, row, *found, **{"Row ID": row["id"]})
            for kind, _table, row in _saved_rows(con) for found in fill_row(kind, row, prices)]


CALENDAR_QUARTERS = {1: "Jan-Mar", 2: "Apr-Jun", 3: "Jul-Sep", 4: "Oct-Dec"}


def quarter_months(fiscal_year, quarter):
    """Months of a fiscal quarter: calendar quarters, except Q1 2026, which starts in December 2025."""
    return "Dec-Mar" if (int(fiscal_year), int(quarter)) == (2026, 1) else CALENDAR_QUARTERS[int(quarter)]


def quarter_label(fiscal_year, quarter):
    return f"Q{quarter} {fiscal_year} ({quarter_months(fiscal_year, quarter)})"


def _where(supplier=None, suppliers=None, fiscal_year=None, quarters=None):
    sql, params = ["shipment_date IS NOT NULL"], []
    if supplier is not None:
        sql.append("supplier_key = %s")
        params.append(supplier)
    if suppliers:
        sql.append(f"supplier_key IN ({', '.join(['%s'] * len(suppliers))})")
        params += list(suppliers)
    if fiscal_year is not None:
        sql.append("fiscal_year = %s")
        params.append(fiscal_year)
    if quarters:
        sql.append(f"quarter IN ({', '.join(['%s'] * len(quarters))})")
        params += list(quarters)
    return " AND ".join(sql), params


def _dated(**filters):
    """SQL source for the dated rows of both views, and its parameters."""
    where, params = _where(**filters)
    return f"(SELECT * FROM v_std WHERE {where} UNION ALL SELECT * FROM v_cus WHERE {where})", params * 2


def rows(con, kind, **filters):
    """Dated rows of v_std or v_cus, sorted by shipment date."""
    where, params = _where(**filters)
    return [dict(r) for r in con.execute(
        f"SELECT * FROM v_{kind.lower()} WHERE {where} ORDER BY shipment_date, id", params)]


def _editable_search(search):
    if not search or not str(search).strip():
        return "", []
    value = f"%{str(search).strip().lower()}%"
    columns = ["id", *COLUMNS, "source_file", "imported_at"]
    return ("(" + " OR ".join(f"LOWER(CAST({col} AS TEXT)) LIKE %s" for col in columns) + ")",
            [value] * len(columns))


def editable_rows(con, kind, supplier=None, limit=None, offset=0, search=None):
    """Raw tracker rows for the web editor, including undated rows.

    Computed report fields are deliberately excluded: rates, rebates, PACD,
    fiscal year and quarter are recalculated by the views after every edit.
    """
    if kind not in TABLES:
        raise ValueError("Carton type must be STD or CUS.")
    clauses, params = [], []
    if supplier is not None:
        clauses.append("supplier_key IS NOT DISTINCT FROM %s")
        params.append(supplier)
    search_sql, search_params = _editable_search(search)
    if search_sql:
        clauses.append(search_sql)
        params += search_params
    where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
    paging = ""
    if limit is not None:
        paging = " LIMIT %s OFFSET %s"
        params += [int(limit), int(offset)]
    return [dict(r) for r in con.execute(
        f"SELECT id, {', '.join(COLUMNS)}, supplier_key, source_file, imported_at "
        f"FROM {TABLES[kind]} {where} "
        f"ORDER BY shipment_date IS NULL, shipment_date, id{paging}", params)]


def editable_row_count(con, kind, supplier=None, search=None):
    """Number of raw rows matching one tracker-editor selection."""
    if kind not in TABLES:
        raise ValueError("Carton type must be STD or CUS.")
    clauses, params = [], []
    if supplier is not None:
        clauses.append("supplier_key IS NOT DISTINCT FROM %s")
        params.append(supplier)
    search_sql, search_params = _editable_search(search)
    if search_sql:
        clauses.append(search_sql)
        params += search_params
    where = f" WHERE {' AND '.join(clauses)}" if clauses else ""
    return con.execute(f"SELECT COUNT(*) FROM {TABLES[kind]}{where}", params).fetchone()[0]


def editable_rows_all(con, supplier=None, limit=None, offset=0, search=None):
    """Raw tracker rows from both carton types for the cross-sheet editor search."""
    selects, params = [], []
    for kind, table in TABLES.items():
        clauses, kind_params = [], []
        if supplier is not None:
            clauses.append("supplier_key IS NOT DISTINCT FROM %s")
            kind_params.append(supplier)
        search_sql, search_params = _editable_search(search)
        if search_sql:
            clauses.append(search_sql)
            kind_params += search_params
        where = f"WHERE {' AND '.join(clauses)}" if clauses else ""
        selects.append(
            f"SELECT '{kind}' AS carton_type, id, {', '.join(COLUMNS)}, supplier_key, source_file, imported_at "
            f"FROM {table} {where}")
        params += kind_params
    paging = ""
    if limit is not None:
        paging = " LIMIT %s OFFSET %s"
        params += [int(limit), int(offset)]
    sql = " UNION ALL ".join(selects)
    return [dict(r) for r in con.execute(
        f"SELECT * FROM ({sql}) AS tracker_rows "
        f"ORDER BY shipment_date IS NULL, shipment_date, carton_type, id{paging}", params)]


def editable_row_count_all(con, supplier=None, search=None):
    """Number of raw tracker rows matching a search across both carton types."""
    total = 0
    for kind in TABLES:
        total += editable_row_count(con, kind, supplier, search)
    return total


def _manual_date(value, row_number, label):
    """Validate a date entered in the tracker editor and return ISO text."""
    if value is None or (isinstance(value, str) and not value.strip()):
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = str(value).strip()
    try:
        return date.fromisoformat(text).isoformat()
    except ValueError:
        parsed = _text_date(text)
        if parsed:
            return parsed.isoformat()
    raise ValueError(
        f"Row {row_number}: {label} must be YYYY-MM-DD or DD/MM/YYYY.")


def update_tracker_rows(con, kind, edited):
    """Save rows from the tracker editor in one transaction.

    ``kind`` fixes the destination table and each mapping must contain an
    existing ``id`` plus the imported columns. Invalid input rolls the whole
    edit back, so a partially edited tracker is never persisted.
    """
    if kind not in TABLES:
        raise ValueError("Carton type must be STD or CUS.")
    table = TABLES[kind]
    prepared, seen, prices = [], set(), get_prices(con)
    for number, item in enumerate(edited, 1):
        try:
            row_id = int(item.get("id"))
        except (TypeError, ValueError):
            raise ValueError(f"Row {number}: row ID is invalid.") from None
        if row_id in seen:
            raise ValueError(f"Row {number}: duplicate row ID {row_id}.")
        seen.add(row_id)
        if not con.execute(f"SELECT 1 FROM {table} WHERE id = %s", (row_id,)).fetchone():
            raise ValueError(f"Row {number}: row ID {row_id} no longer exists. Refresh the page.")

        values = []
        for col in COLUMNS:
            value = _clean(item.get(col))
            if col in DATE_COLS:
                value = _manual_date(value, number, HEADERS[col])
            elif value is None:
                pass
            elif col in INT_COLS or col in REAL_COLS:
                try:
                    value = int(round(float(value))) if col in INT_COLS else float(value)
                except (TypeError, ValueError):
                    raise ValueError(f"Row {number}: {HEADERS[col]} must be a number.") from None
            else:
                value = str(value)
            values.append(value)
        row = dict(zip(COLUMNS, values))
        fill_row(kind, row, prices)
        values = [row[col] for col in COLUMNS]
        key = supplier_key(values[COLUMNS.index("packaging_supplier")])
        prepared.append(values + [key, row_id])

    assignments = ", ".join(f"{col} = %s" for col in COLUMNS)
    now = datetime.now().isoformat(timespec="seconds")
    with con.transaction():
        con.cursor().executemany(
            f"UPDATE {table} SET {assignments}, supplier_key = %s, imported_at = %s WHERE id = %s",
            [values[:-1] + [now, values[-1]] for values in prepared])
        _bump_revision(con)
    return len(prepared)


def _delete(con, table, where, params):
    """DELETE in the caller's transaction; returns the number of rows removed."""
    return con.execute(f"DELETE FROM {table} WHERE {where}", params).rowcount


def delete_tracker_rows(con, kind, ids):
    """Delete these row ids from one tracker sheet. Ids already gone are ignored. Returns rows deleted."""
    if kind not in TABLES:
        raise ValueError("Carton type must be STD or CUS.")
    ids = [int(i) for i in ids]
    if not ids:
        return 0
    with con.transaction():
        count = _delete(con, TABLES[kind], "id = ANY(%s)", (ids,))
        if count:
            _bump_revision(con)
    return count


def _delete_both(con, where, params):
    with con.transaction():
        counts = {kind: _delete(con, table, where, params) for kind, table in TABLES.items()}
        if any(counts.values()):
            _bump_revision(con)
    return counts


def delete_supplier_month(con, supplier, month):
    """Delete one supplier's dated rows shipped in `month` (YYYY-MM) from both sheets. Returns {kind: rows}."""
    return _delete_both(con, "supplier_key IS NOT DISTINCT FROM %s AND substr(shipment_date, 1, 7) = %s",
                        (supplier, month))


def delete_source_file(con, source_file):
    """Delete every tracker row imported from `source_file`, both sheets. Returns {kind: rows}."""
    return _delete_both(con, "source_file = %s", (source_file,))


def tracker_months(con, supplier):
    """Shipment months (YYYY-MM) with dated rows for this supplier, newest first."""
    sql = " UNION ".join(f"SELECT substr(shipment_date, 1, 7) AS month FROM {table} "
                         "WHERE supplier_key IS NOT DISTINCT FROM %s AND shipment_date IS NOT NULL"
                         for table in TABLES.values())
    return [r[0] for r in con.execute(f"SELECT month FROM ({sql}) ORDER BY month DESC", (supplier,) * len(TABLES))]


def tracker_files(con):
    """One dict per imported source file: source_file, STD and CUS row counts, latest imported_at; newest first."""
    sql = " UNION ALL ".join(f"SELECT '{kind}' AS kind, source_file, imported_at FROM {table}"
                             for kind, table in TABLES.items())
    rows = con.execute(
        f"SELECT source_file, COUNT(*) FILTER (WHERE kind = 'STD'), COUNT(*) FILTER (WHERE kind = 'CUS'), "
        f"MAX(imported_at) FROM ({sql}) WHERE source_file IS NOT NULL "
        f"GROUP BY source_file ORDER BY MAX(imported_at) DESC NULLS LAST, source_file")
    return [{"source_file": f, "STD": std, "CUS": cus, "imported_at": at} for f, std, cus, at in rows]


def summary_rows(con, fiscal_year, supplier=None):
    """The generated Summary workbook's supplier/quarter/carton-type rows.

    Values and rate labels deliberately follow the same rules as the workbook,
    making this the shared source for both the live dashboard and Excel output.
    """
    keys = [supplier] if supplier is not None else suppliers(con)
    out = []
    for key in keys:
        totals = {kind: quarter_totals(con, kind, key, fiscal_year) for kind in TABLES}
        rates = {kind: quarter_rates(con, kind, key, fiscal_year) for kind in TABLES}
        quarters = sorted(set(totals["STD"]) | set(totals["CUS"]))
        for quarter in quarters:
            for kind, label in (("STD", "Standard"), ("CUS", "Custom")):
                qty, amount, pacd, rebate = totals[kind].get(quarter, (0, 0, 0, 0))
                if quarter not in rates[kind]:
                    pacd_rate, rebate_rate = "-", "-"
                else:
                    pacd_rate, rebate_rate = (
                        "Mixed" if value is None else value
                        for value in rates[kind][quarter])
                out.append({
                    "supplier_key": key,
                    "packaging_supplier": supplier_name(con, key),
                    "fiscal_year": int(fiscal_year),
                    "quarter": quarter,
                    "quarter_label": f"Q{quarter} ({quarter_months(fiscal_year, quarter)})",
                    "carton_type": label,
                    "pacd_rate": pacd_rate,
                    "rebate_rate": rebate_rate,
                    "cartons": int(qty or 0),
                    "total_amount": round(amount or 0, 2),
                    "pacd": round(pacd or 0, 2),
                    "matalan_rebate": round(rebate or 0, 2),
                })
    return out


def cross_reference_overview(con):
    """One live Tracker-vs-eBiz status summary per packaging supplier."""
    overview = []
    for key in reconciliation_suppliers(con):
        data = cross_reference(con, key)
        statuses = {label: 0 for label in ("Match", "Not in eBiz", "Tracker > eBiz", "Tracker < eBiz")}
        for order in data["orders"]:
            statuses[order["status"]] += 1
        overview.append({
            "supplier_key": key,
            "packaging_supplier": supplier_name(con, key),
            "tracker_orders": len(data["orders"]),
            "matches": statuses["Match"],
            "not_in_ebiz": statuses["Not in eBiz"],
            "quantity_mismatches": statuses["Tracker > eBiz"] + statuses["Tracker < eBiz"],
            "ebiz_only": len(data["ebiz_only"]),
        })
    return overview


def suppliers(con):
    """Supplier keys of the dated rows, sorted."""
    src, params = _dated()
    return [r[0] for r in con.execute(
        f"SELECT DISTINCT supplier_key COLLATE \"C\" FROM {src} WHERE supplier_key IS NOT NULL ORDER BY 1", params)]


def tracker_suppliers(con):
    """Supplier keys from every raw tracker row, including undated rows."""
    sql = " UNION ".join(
        f"SELECT supplier_key FROM {table} WHERE supplier_key IS NOT NULL"
        for table in TABLES.values())
    return [r[0] for r in con.execute(f'SELECT supplier_key FROM ({sql}) ORDER BY supplier_key COLLATE "C"')]


def reconciliation_suppliers(con):
    """Supplier keys found in the tracker or in eBiz, sorted."""
    return sorted(set(tracker_suppliers(con)) | set(ebiz_vendors(con)))


def supplier_name(con, key):
    """Most common Tracker supplier spelling, then eBiz vendor spelling, else the key."""
    src = " UNION ALL ".join(
        f"SELECT packaging_supplier, supplier_key FROM {table}" for table in TABLES.values())
    row = con.execute(
        f"SELECT packaging_supplier FROM ({src}) "
        "WHERE supplier_key IS NOT DISTINCT FROM %s AND packaging_supplier IS NOT NULL "
        'GROUP BY packaging_supplier ORDER BY COUNT(*) DESC, packaging_supplier COLLATE "C" LIMIT 1',
        (key,)).fetchone()
    if row:
        return row[0]
    row = con.execute(
        "SELECT vendor FROM ebiz WHERE vendor_key IS NOT DISTINCT FROM %s AND vendor IS NOT NULL "
        'GROUP BY vendor ORDER BY COUNT(*) DESC, vendor COLLATE "C" LIMIT 1', (key,)).fetchone()
    return row[0] if row else key


def fiscal_years(con, **filters):
    src, params = _dated(**filters)
    return [r[0] for r in con.execute(f"SELECT DISTINCT fiscal_year FROM {src} ORDER BY 1", params)]


def quarters(con, **filters):
    src, params = _dated(**filters)
    return [r[0] for r in con.execute(f"SELECT DISTINCT quarter FROM {src} ORDER BY 1", params)]


def date_range(con, **filters):
    """(first, last) shipment date of the matching rows, or None."""
    src, params = _dated(**filters)
    first, last = con.execute(f"SELECT MIN(shipment_date), MAX(shipment_date) FROM {src}", params).fetchone()
    return (date.fromisoformat(first), date.fromisoformat(last)) if first else None


def _unrated(con):
    """(count, first shipment date) of the dated rows that have no rate."""
    src, params = _dated()
    return con.execute(f"SELECT COUNT(*), MIN(shipment_date) FROM {src} WHERE pacd_pct IS NULL", params).fetchone()


def unrated_count(con):
    """Dated rows with no rate: shipped before the first rate of their carton type."""
    return _unrated(con)[0]


def quarter_totals(con, kind, supplier, fiscal_year):
    """{quarter: (qty, total amount, pacd, matalan rebate)} for one supplier and fiscal year."""
    where, params = _where(supplier=supplier, fiscal_year=fiscal_year)
    return {r[0]: tuple(r[1:]) for r in con.execute(
        f"SELECT quarter, SUM(carton_qty), SUM(total_amount), SUM(pacd), SUM(matalan_rebate) "
        f"FROM v_{kind.lower()} WHERE {where} GROUP BY quarter", params)}


def quarter_rates(con, kind, supplier, fiscal_year):
    """{quarter: (pacd %, matalan %)} for one supplier and fiscal year; None where the rows don't share one rate."""
    where, params = _where(supplier=supplier, fiscal_year=fiscal_year)
    out = {}
    for q, n_pacd, pacd, n_matalan, matalan, unrated in con.execute(
            f"SELECT quarter, COUNT(DISTINCT pacd_pct), MIN(pacd_pct), COUNT(DISTINCT matalan_pct), "
            f"MIN(matalan_pct), COUNT(*) FILTER (WHERE pacd_pct IS NULL) FROM v_{kind.lower()} WHERE {where} GROUP BY quarter", params):
        out[q] = (pacd if n_pacd == 1 and not unrated else None, matalan if n_matalan == 1 and not unrated else None)
    return out


def rate_headers(rows):
    """Headers for the PACD and rebate columns; they show the % when every row has the same rate."""
    def label(key, name):
        pcts = {r[key] for r in rows}
        return f"{name} ({round(pcts.pop() * 100, 4):g}%)" if len(pcts) == 1 and None not in pcts else name
    return label("pacd_pct", "PACD Rebate"), label("matalan_pct", "Matalan Rebate")
