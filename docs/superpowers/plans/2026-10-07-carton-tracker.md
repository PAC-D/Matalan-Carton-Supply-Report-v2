# Matalan Carton Tracker DB Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A local app that imports the `TRACK_STD` and `TRACK_CUS` sheets of a Matalan carton tracker workbook into SQLite and builds the phase-1 tracker reports from the database.

**Architecture:** `core.py` owns the SQLite schema, the rate-aware views, the Excel import and the report queries. `reports.py` turns query results into xlsx (openpyxl) and PDF (reportlab) files, each returned as `(file name, bytes)`. `app.py` is a thin Streamlit UI with three tabs (Import, Rates, Reports). `check.py` is the pytest self-check.

**Tech Stack:** Python 3.14, sqlite3 (stdlib), openpyxl 3.1.5, pandas 3.0.5, reportlab 5.0.0, streamlit 1.61.1, pypdf 6.16.0 (tests only), pytest 9.1.1.

**Spec:** `docs/superpowers/specs/2026-10-07-carton-tracker-design.md`

## Global Constraints

- Project folder (all paths below are relative to it): `C:\Users\Shoaib\OneDrive - PacD\Projects\Matalan\Matalan Carton Supplu Report`. Run every command from this folder in PowerShell.
- No new dependencies. Everything listed in Tech Stack is already installed.
- The folder is not a git repository. There are no commit steps; each task ends with the full test run passing.
- Do not modify `eBiz Report Generator.xlsm`, `eBiz Report Generator_.xlsm` or anything under `Reports/August/`.
- Only PACD and Matalan rebate are calculated. Every other column is stored exactly as Excel cached it (`openpyxl.load_workbook(..., data_only=True)`).
- `matalan_rebate = unit_price × carton_qty × matalan_pct`; `pacd = (total_amount − matalan_rebate) × pacd_pct`.
- Rate for a row = the `rates` row of the same carton type with the latest `effective_from` on or before `shipment_date`.
- Seed rates: STD 2025-12-01 4% / 5%; CUS 2025-12-01 7.5% / 10%; STD 2026-07-01 4% / 5%; CUS 2026-07-01 7.5% / 5%.
- Quarters by shipment month: Dec–Mar = Q1, Apr–Jun = Q2, Jul–Aug = Q3, Sep–Nov = Q4. December belongs to the next fiscal year (2025-12-01 is Q1 2026).
- Re-import rule: for each (packaging supplier, shipment month) in the file, delete those DB rows, then insert the file's rows, in one transaction.
- File names use the tracker's Packaging Supplier value with `\ / : * ? " < > |` removed and trailing dots and spaces stripped.
- Test commands: `python -m pytest check.py -v` (all) or `python -m pytest check.py -v -k <name>` (one).

## File Structure

| File | Responsibility |
|---|---|
| `core.py` | Schema, seed rates, views, Excel import, rates editing, report queries |
| `reports.py` | Summary / Raw Data / Shipment Tracking xlsx, Full Tracker PDF, bundle + save to folder |
| `app.py` | Streamlit UI (Import, Rates, Reports tabs) |
| `check.py` | pytest self-check |
| `run.bat` | Starts the app |
| `samples/Matalan Carton Shipment Tracking_NPS_BD Aug-26.xlsx` | Demo workbook used by the tests |
| `assets/logo_left.png`, `assets/logo_right.png` | PDF header logos |
| `carton.db` | Created on first run (not created by tests) |

---

### Task 1: Database schema, seed rates and rate-aware views

**Files:**
- Create: `core.py`
- Create: `check.py`
- Create: `samples/Matalan Carton Shipment Tracking_NPS_BD Aug-26.xlsx` (copy)

**Interfaces:**
- Consumes: nothing.
- Produces:
  - `core.connect(path=None) -> sqlite3.Connection` (row_factory = `sqlite3.Row`; path falls back to env `CARTON_DB`, then `carton.db` next to `core.py`).
  - `core.TABLES = {"STD": "track_std", "CUS": "track_cus"}`, `core.SHEETS = {"STD": "TRACK_STD", "CUS": "TRACK_CUS"}`.
  - `core.HEADERS: dict[str, str]` (db column -> workbook header text, in sheet order), `core.COLUMNS: list[str]`, `core.DATE_COLS`, `core.INT_COLS`, `core.REAL_COLS`.
  - Views `v_std` and `v_cus`: all table columns plus `pacd_pct`, `matalan_pct`, `matalan_rebate`, `pacd`, `fiscal_year` (INTEGER), `quarter` (INTEGER 1–4).

- [ ] **Step 1: Copy the demo workbook into the project**

```powershell
New-Item -ItemType Directory -Force samples | Out-Null
Copy-Item "C:\Users\Shoaib\Downloads\Matalan Carton Shipment Tracking_NPS_BD Aug-26.xlsx" "samples\"
Get-Item "samples\Matalan Carton Shipment Tracking_NPS_BD Aug-26.xlsx" | Select-Object Name, Length
```

Expected: one file, Length 340752.

- [ ] **Step 2: Write the failing tests**

Create `check.py`:

```python
"""Self-check for the carton tracker. Run: python -m pytest check.py -v"""
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import openpyxl
import pytest

import core

HERE = Path(__file__).parent
DEMO = HERE / "samples" / "Matalan Carton Shipment Tracking_NPS_BD Aug-26.xlsx"
TW = "Transworld Packaging (BD) Ltd."
UG = "Uniglory Packaging Industries Ltd."


def insert(con, kind, **values):
    cols = ", ".join(values)
    marks = ", ".join("?" * len(values))
    with con:
        con.execute(f"INSERT INTO {core.TABLES[kind]} ({cols}) VALUES ({marks})", list(values.values()))


def test_quarter_mapping(tmp_path):
    con = core.connect(tmp_path / "t.db")
    cases = {
        "2025-12-01": (2026, 1), "2026-03-31": (2026, 1), "2026-04-01": (2026, 2),
        "2026-08-31": (2026, 3), "2026-09-01": (2026, 4), "2026-11-30": (2026, 4),
        "2026-12-01": (2027, 1),
    }
    for d in cases:
        insert(con, "STD", shipment_date=d, carton_qty=1, unit_price=1.0, total_amount=1.0)
    got = {r["shipment_date"]: (r["fiscal_year"], r["quarter"]) for r in con.execute("SELECT * FROM v_std")}
    assert got == cases


def test_rate_follows_shipment_date(tmp_path):
    con = core.connect(tmp_path / "t.db")
    insert(con, "CUS", shipment_date="2026-06-30", carton_qty=100, unit_price=2.0, total_amount=200.0)
    insert(con, "CUS", shipment_date="2026-07-01", carton_qty=100, unit_price=2.0, total_amount=200.0)
    insert(con, "STD", shipment_date="2026-07-01", carton_qty=100, unit_price=2.0, total_amount=200.0)
    cus = {r["shipment_date"]: r for r in con.execute("SELECT * FROM v_cus")}
    assert cus["2026-06-30"]["matalan_rebate"] == pytest.approx(20.0)  # 10%
    assert cus["2026-06-30"]["pacd"] == pytest.approx(13.5)  # (200 - 20) * 7.5%
    assert cus["2026-07-01"]["matalan_rebate"] == pytest.approx(10.0)  # 5% from Q3 2026
    assert cus["2026-07-01"]["pacd"] == pytest.approx(14.25)  # (200 - 10) * 7.5%
    std = con.execute("SELECT * FROM v_std").fetchone()
    assert std["matalan_rebate"] == pytest.approx(10.0)  # 5%
    assert std["pacd"] == pytest.approx(7.6)  # (200 - 10) * 4%


def test_rebate_uses_unit_price_and_pacd_uses_total(tmp_path):
    con = core.connect(tmp_path / "t.db")
    insert(con, "STD", shipment_date="2026-08-01", carton_qty=100, unit_price=2.0, total_amount=999.0)
    r = con.execute("SELECT * FROM v_std").fetchone()
    assert r["matalan_rebate"] == pytest.approx(10.0)  # 2.0 * 100 * 5%
    assert r["pacd"] == pytest.approx((999.0 - 10.0) * 0.04)


def test_row_without_date_has_no_rate_or_quarter(tmp_path):
    con = core.connect(tmp_path / "t.db")
    insert(con, "CUS", shipment_date=None, carton_qty=5, unit_price=1.0, total_amount=5.0)
    r = con.execute("SELECT * FROM v_cus").fetchone()
    assert (r["pacd"], r["matalan_rebate"], r["fiscal_year"], r["quarter"]) == (None, None, None, None)


def test_connect_is_repeatable_and_seeds_once(tmp_path):
    core.connect(tmp_path / "t.db")
    con = core.connect(tmp_path / "t.db")
    assert con.execute("SELECT COUNT(*) FROM rates").fetchone()[0] == 4
```

- [ ] **Step 3: Run tests to verify they fail**

Run: `python -m pytest check.py -v`
Expected: collection error `ModuleNotFoundError: No module named 'core'`.

- [ ] **Step 4: Write `core.py`**

```python
"""SQLite storage, Excel import and report queries for the Matalan carton tracker."""
import os
import sqlite3
from pathlib import Path

DEFAULT_DB = Path(__file__).with_name("carton.db")
TABLES = {"STD": "track_std", "CUS": "track_cus"}
SHEETS = {"STD": "TRACK_STD", "CUS": "TRACK_CUS"}

# Database column -> header text in row 2 of the tracker sheets, in sheet order.
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
    "unit_price": "Carton's Unit Price (CNY,EXW w/VAT)",
    "final_price": "Final Price, CNY/CTN",
    "total_amount": "Total Amount",
    "note": "NOTE",
}
COLUMNS = list(HEADERS)
DATE_COLS = {"po_received_date", "shipment_date"}
INT_COLS = {"carton_qty"}
REAL_COLS = {"length_mm", "width_mm", "height_mm", "l_mm", "w_mm", "area_sqm",
             "unit_price", "final_price", "total_amount"}

SEED_RATES = [
    ("STD", "2025-12-01", 0.04, 0.05),
    ("CUS", "2025-12-01", 0.075, 0.10),
    ("STD", "2026-07-01", 0.04, 0.05),
    ("CUS", "2026-07-01", 0.075, 0.05),
]

VIEW_SQL = """
CREATE VIEW {view} AS
SELECT t.*, r.pacd_pct, r.matalan_pct,
  t.unit_price * t.carton_qty * r.matalan_pct AS matalan_rebate,
  (t.total_amount - t.unit_price * t.carton_qty * r.matalan_pct) * r.pacd_pct AS pacd,
  CAST(strftime('%Y', t.shipment_date) AS INTEGER) + (strftime('%m', t.shipment_date) = '12') AS fiscal_year,
  CASE
    WHEN t.shipment_date IS NULL THEN NULL
    WHEN CAST(strftime('%m', t.shipment_date) AS INTEGER) IN (12, 1, 2, 3) THEN 1
    WHEN CAST(strftime('%m', t.shipment_date) AS INTEGER) BETWEEN 4 AND 6 THEN 2
    WHEN CAST(strftime('%m', t.shipment_date) AS INTEGER) IN (7, 8) THEN 3
    ELSE 4
  END AS quarter
FROM {table} t
LEFT JOIN rates r ON r.carton_type = '{kind}' AND r.effective_from = (
  SELECT MAX(effective_from) FROM rates
  WHERE carton_type = '{kind}' AND effective_from <= t.shipment_date);
"""


def _col_type(col):
    return "INTEGER" if col in INT_COLS else "REAL" if col in REAL_COLS else "TEXT"


def _schema():
    cols = ",\n  ".join(f"{c} {_col_type(c)}" for c in COLUMNS)
    sql = [f"CREATE TABLE IF NOT EXISTS {t} (\n  id INTEGER PRIMARY KEY,\n  {cols},\n"
           f"  source_file TEXT,\n  imported_at TEXT\n);" for t in TABLES.values()]
    sql.append("""CREATE TABLE IF NOT EXISTS rates (
  carton_type TEXT NOT NULL CHECK (carton_type IN ('STD', 'CUS')),
  effective_from TEXT NOT NULL,
  pacd_pct REAL NOT NULL,
  matalan_pct REAL NOT NULL,
  PRIMARY KEY (carton_type, effective_from)
);""")
    for kind, table in TABLES.items():
        view = f"v_{kind.lower()}"
        sql.append(f"DROP VIEW IF EXISTS {view};")
        sql.append(VIEW_SQL.format(view=view, table=table, kind=kind))
    return "\n".join(sql)


def connect(path=None):
    """Open the database and create anything missing.

    Path: the argument, else env CARTON_DB, else carton.db next to this file.
    """
    con = sqlite3.connect(path or os.environ.get("CARTON_DB") or DEFAULT_DB)
    con.row_factory = sqlite3.Row
    con.executescript(_schema())
    if not con.execute("SELECT 1 FROM rates LIMIT 1").fetchone():
        with con:
            con.executemany("INSERT INTO rates VALUES (?, ?, ?, ?)", SEED_RATES)
    return con
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `python -m pytest check.py -v`
Expected: 5 passed.

---

### Task 2: Excel import with preview and replace-overlap save

**Files:**
- Modify: `core.py` (add imports at top, append import functions)
- Modify: `check.py` (append helpers and tests)

**Interfaces:**
- Consumes: `core.connect`, `core.HEADERS`, `core.COLUMNS`, `core.TABLES`, `core.SHEETS`, `core.DATE_COLS`, `core.INT_COLS`, `core.REAL_COLS` (Task 1).
- Produces:
  - `core.ImportFileError(ValueError)`.
  - `core.read_workbook(src) -> dict[str, list[dict]]`: keys `"STD"`, `"CUS"`; each row dict has every name in `core.COLUMNS` plus `"_row"` (int, Excel row) and `"_bad_date"` (bool). `src` is a path or a binary file object.
  - `core.preview(con, data) -> list[dict]` with keys `"Sheet"`, `"Packaging supplier"`, `"Shipment month"`, `"Rows in file"`, `"Rows replaced in DB"`.
  - `core.bad_dates(data) -> list[dict]` with keys `"Sheet"`, `"Excel row"`, `"Packaging supplier"`.
  - `core.save(con, data, source_file) -> dict[str, int]` (`{"STD": n, "CUS": m}` rows inserted).

- [ ] **Step 1: Write the failing tests**

Append to `check.py`:

```python
def rec(ship="2026-08-03", sup="Pack A Ltd.", qty=100, price=1.0, **kw):
    """One tracker row for make_xlsx. Keys are core.COLUMNS names."""
    row = {
        "shipment_date": datetime.fromisoformat(ship) if ship else None,
        "packaging_supplier": sup,
        "carton_qty": qty,
        "unit_price": price,
        "final_price": price,
        "total_amount": qty * price if qty and price else None,
    }
    row.update(kw)
    return row


def make_xlsx(path, std=(), cus=(), drop=None, formula_total=False, names=("TRACK_STD", "TRACK_CUS")):
    """Write a minimal tracker workbook: headers in row 2, data from row 3."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in zip(names, (std, cus)):
        ws = wb.create_sheet(name)
        cols = [c for c in core.COLUMNS if c != drop]
        for i, c in enumerate(cols, 1):
            ws.cell(2, i, core.HEADERS[c])
        for r, row in enumerate(rows, 3):
            for i, c in enumerate(cols, 1):
                value = "=1+1" if c == "total_amount" and formula_total else row.get(c)
                ws.cell(r, i, value)
    wb.save(path)
    return path


def test_demo_import_counts_and_totals(tmp_path):
    con = core.connect(tmp_path / "t.db")
    data = core.read_workbook(DEMO)
    assert {k: len(v) for k, v in data.items()} == {"STD": 88, "CUS": 118}
    assert core.save(con, data, DEMO.name) == {"STD": 88, "CUS": 118}

    def sums(view):
        return tuple(con.execute(f"SELECT SUM(carton_qty), SUM(total_amount), SUM(pacd), "
                                 f"SUM(matalan_rebate) FROM {view}").fetchone())

    assert sums("v_std") == pytest.approx((28001, 33646.9152, 1278.5828, 1682.3458), abs=1e-3)
    assert sums("v_cus") == pytest.approx((41529, 37815.6551, 2694.3654, 1890.7828), abs=1e-3)

    core.save(con, core.read_workbook(DEMO), DEMO.name)  # importing again replaces, not duplicates
    assert con.execute("SELECT COUNT(*) FROM track_std").fetchone()[0] == 88
    assert con.execute("SELECT COUNT(*) FROM track_cus").fetchone()[0] == 118


def test_template_rows_are_skipped(tmp_path):
    path = make_xlsx(tmp_path / "a.xlsx", std=[rec(), {"carton_selection": "Custom", "flute": "BC-Flute"}])
    assert len(core.read_workbook(path)["STD"]) == 1


def test_values_are_cleaned(tmp_path):
    path = make_xlsx(tmp_path / "a.xlsx", std=[rec(supplier_name="Lariz Fashion Ltd ", ebiz_so=7124457,
                                                   ebiz_customer_ref="#N/A")])
    row = core.read_workbook(path)["STD"][0]
    assert row["supplier_name"] == "Lariz Fashion Ltd"
    assert row["ebiz_so"] == "7124457"
    assert row["ebiz_customer_ref"] is None
    assert row["shipment_date"] == "2026-08-03"
    assert row["carton_qty"] == 100


def test_sheet_names_match_any_case(tmp_path):
    path = make_xlsx(tmp_path / "a.xlsx", std=[rec()], cus=[rec()], names=("track_std", "Track_Cus"))
    assert {k: len(v) for k, v in core.read_workbook(path).items()} == {"STD": 1, "CUS": 1}


def test_missing_sheet_is_an_error(tmp_path):
    wb = openpyxl.Workbook()
    wb.active.title = "TRACK_STD"
    wb.save(tmp_path / "a.xlsx")
    with pytest.raises(core.ImportFileError, match="TRACK_CUS"):
        core.read_workbook(tmp_path / "a.xlsx")


def test_missing_header_is_an_error(tmp_path):
    path = make_xlsx(tmp_path / "a.xlsx", std=[rec()], drop="carton_qty")
    with pytest.raises(core.ImportFileError, match="Carton QTY"):
        core.read_workbook(path)


def test_file_without_cached_values_is_an_error(tmp_path):
    path = make_xlsx(tmp_path / "a.xlsx", std=[rec()], formula_total=True)
    with pytest.raises(core.ImportFileError, match="no calculated values"):
        core.read_workbook(path)


def test_bad_shipment_date_is_flagged_and_kept(tmp_path):
    con = core.connect(tmp_path / "t.db")
    path = make_xlsx(tmp_path / "a.xlsx", cus=[rec(), rec(ship=None, shipment_date="TBC")])
    data = core.read_workbook(path)
    assert [r["_bad_date"] for r in data["CUS"]] == [False, True]
    assert core.bad_dates(data) == [{"Sheet": "TRACK_CUS", "Excel row": 4, "Packaging supplier": "Pack A Ltd."}]
    core.save(con, data, "a.xlsx")
    assert con.execute("SELECT COUNT(*) FROM v_cus WHERE quarter IS NULL").fetchone()[0] == 1


def test_reimport_replaces_only_overlapping_supplier_months(tmp_path):
    con = core.connect(tmp_path / "t.db")
    first = make_xlsx(tmp_path / "a.xlsx", std=[rec(sup="A"), rec(sup="A"), rec(sup="B"),
                                                rec(ship="2026-07-15", sup="A")])
    core.save(con, core.read_workbook(first), "a.xlsx")
    second = core.read_workbook(make_xlsx(tmp_path / "b.xlsx", std=[rec(sup="A", qty=5)]))
    assert core.preview(con, second) == [{"Sheet": "TRACK_STD", "Packaging supplier": "A",
                                          "Shipment month": "2026-08", "Rows in file": 1,
                                          "Rows replaced in DB": 2}]
    core.save(con, second, "b.xlsx")
    got = sorted((r["packaging_supplier"], r["shipment_date"], r["carton_qty"])
                 for r in con.execute("SELECT * FROM track_std"))
    assert got == [("A", "2026-07-15", 100), ("A", "2026-08-03", 5), ("B", "2026-08-03", 100)]
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest check.py -v`
Expected: the 9 new tests FAIL with `AttributeError: module 'core' has no attribute 'read_workbook'` (or `'ImportFileError'`); the 5 Task 1 tests still pass.

- [ ] **Step 3: Implement the import**

In `core.py`, replace the import block at the top with:

```python
import os
import re
import sqlite3
from collections import Counter
from datetime import date, datetime
from pathlib import Path

import openpyxl
```

Append to `core.py`:

```python
EXCEL_ERRORS = {"#N/A", "#VALUE!", "#REF!", "#DIV/0!", "#NAME?", "#NUM!", "#NULL!"}


class ImportFileError(ValueError):
    """The workbook can't be imported. The message says why."""


def _norm(header):
    return re.sub(r"\s+", " ", str(header or "")).strip().lower()


def _clean(value):
    if isinstance(value, str):
        value = value.strip()
        if not value or value in EXCEL_ERRORS:
            return None
    return value


def _convert(col, value):
    if value is None:
        return None
    if col in DATE_COLS:
        if isinstance(value, datetime):
            return value.date().isoformat()
        return value.isoformat() if isinstance(value, date) else None
    if col in INT_COLS or col in REAL_COLS:
        try:
            number = float(value)
        except (TypeError, ValueError):
            return None
        return int(round(number)) if col in INT_COLS else number
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value)


def _find_columns(sheet, header_row):
    found = {_norm(h): i for i, h in enumerate(header_row)}
    index = {}
    for col in COLUMNS:
        index[col] = found.get(_norm(HEADERS[col]))
        if index[col] is None and col == "unit_price":
            index[col] = next((i for h, i in found.items() if h.startswith("carton's unit price")), None)
        if index[col] is None:
            raise ImportFileError(f"{sheet}: column '{HEADERS[col]}' not found in row 2.")
    return index


def read_workbook(src):
    """Read the data rows of TRACK_STD and TRACK_CUS from a path or binary file object.

    Returns {"STD": [...], "CUS": [...]}. Each row is a dict of COLUMNS plus
    "_row" (Excel row number) and "_bad_date" (Shipment Date present but not a date).
    """
    try:
        wb = openpyxl.load_workbook(src, data_only=True)
    except Exception as e:  # not an xlsx, corrupt or password protected
        raise ImportFileError(f"Can't open the workbook: {e}") from e
    names = {n.upper(): n for n in wb.sheetnames}
    for sheet in SHEETS.values():
        if sheet not in names:
            raise ImportFileError(f"Sheet '{sheet}' not found.")
    data = {}
    for kind, sheet in SHEETS.items():
        ws = wb[names[sheet]]
        index = _find_columns(sheet, next(ws.iter_rows(min_row=2, max_row=2, values_only=True), ()))
        rows = []
        for n, values in enumerate(ws.iter_rows(min_row=3, values_only=True), start=3):
            raw = {c: _clean(values[i]) if i < len(values) else None for c, i in index.items()}
            if raw["shipment_date"] is None and raw["carton_qty"] is None:
                continue  # template row: formulas but no input
            row = {c: _convert(c, v) for c, v in raw.items()}
            row["_row"] = n
            row["_bad_date"] = raw["shipment_date"] is not None and row["shipment_date"] is None
            rows.append(row)
        if rows and all(r["total_amount"] is None for r in rows):
            raise ImportFileError(f"{sheet}: the file has no calculated values. "
                                  "Open it in Excel, save, and import again.")
        data[kind] = rows
    return data


def _month(row):
    return row["shipment_date"][:7] if row["shipment_date"] else None


_SAME_KEY = "packaging_supplier IS ? AND substr(shipment_date, 1, 7) IS ?"


def preview(con, data):
    """Rows per sheet, packaging supplier and shipment month, and how many DB rows each replaces."""
    out = []
    for kind, rows in data.items():
        groups = Counter((r["packaging_supplier"], _month(r)) for r in rows)
        for (supplier, month), n in sorted(groups.items(), key=lambda g: (str(g[0][0]), str(g[0][1]))):
            replaced = con.execute(f"SELECT COUNT(*) FROM {TABLES[kind]} WHERE {_SAME_KEY}",
                                   (supplier, month)).fetchone()[0]
            out.append({"Sheet": SHEETS[kind], "Packaging supplier": supplier,
                        "Shipment month": month or "no valid date", "Rows in file": n,
                        "Rows replaced in DB": replaced})
    return out


def bad_dates(data):
    """Rows whose Shipment Date is filled in but is not a date."""
    return [{"Sheet": SHEETS[kind], "Excel row": r["_row"], "Packaging supplier": r["packaging_supplier"]}
            for kind, rows in data.items() for r in rows if r["_bad_date"]]


def save(con, data, source_file):
    """Per (packaging supplier, shipment month) in the file: delete those DB rows, insert the file's rows.

    Runs in one transaction. Returns the number of rows inserted per sheet.
    """
    now = datetime.now().isoformat(timespec="seconds")
    cols = COLUMNS + ["source_file", "imported_at"]
    with con:
        for kind, rows in data.items():
            table = TABLES[kind]
            for key in {(r["packaging_supplier"], _month(r)) for r in rows}:
                con.execute(f"DELETE FROM {table} WHERE {_SAME_KEY}", key)
            con.executemany(f"INSERT INTO {table} ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
                            [[r[c] for c in COLUMNS] + [source_file, now] for r in rows])
    return {kind: len(rows) for kind, rows in data.items()}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest check.py -v`
Expected: 14 passed.

---

### Task 3: Editing rates

**Files:**
- Modify: `core.py` (append)
- Modify: `check.py` (append)

**Interfaces:**
- Consumes: `core.connect`, `core.TABLES` (Task 1); `insert` test helper (Task 1).
- Produces:
  - `core.get_rates(con) -> list[dict]` with keys `carton_type`, `effective_from`, `pacd_pct`, `matalan_pct` (fractions, e.g. 0.075).
  - `core.save_rates(con, rates: list[dict]) -> None`; raises `ValueError` and saves nothing when a row is invalid.

- [ ] **Step 1: Write the failing tests**

Append to `check.py`:

```python
SEEDED = [
    {"carton_type": "STD", "effective_from": "2025-12-01", "pacd_pct": 0.04, "matalan_pct": 0.05},
    {"carton_type": "STD", "effective_from": "2026-07-01", "pacd_pct": 0.04, "matalan_pct": 0.05},
    {"carton_type": "CUS", "effective_from": "2025-12-01", "pacd_pct": 0.075, "matalan_pct": 0.10},
    {"carton_type": "CUS", "effective_from": "2026-07-01", "pacd_pct": 0.075, "matalan_pct": 0.05},
]


def test_seeded_rates(tmp_path):
    assert core.get_rates(core.connect(tmp_path / "t.db")) == SEEDED


def test_rate_change_applies_to_existing_rows(tmp_path):
    con = core.connect(tmp_path / "t.db")
    insert(con, "CUS", shipment_date="2026-09-10", carton_qty=100, unit_price=2.0, total_amount=200.0)
    rates = core.get_rates(con)
    rates[-1]["matalan_pct"] = 0.08
    core.save_rates(con, rates)
    assert con.execute("SELECT matalan_rebate FROM v_cus").fetchone()[0] == pytest.approx(16.0)


@pytest.mark.parametrize("change, message", [
    ({"carton_type": "XYZ"}, "STD or CUS"),
    ({"effective_from": "None"}, "must be a date"),
    ({"pacd_pct": 1.5}, "between 0 and 100"),
    ({"matalan_pct": float("nan")}, "between 0 and 100"),
    ({"effective_from": "2026-07-01"}, "duplicate"),
])
def test_invalid_rates_are_rejected(tmp_path, change, message):
    con = core.connect(tmp_path / "t.db")
    rates = core.get_rates(con)
    rates[0].update(change)  # rates[0] is STD 2025-12-01; the duplicate case collides with STD 2026-07-01
    with pytest.raises(ValueError, match=message):
        core.save_rates(con, rates)
    assert core.get_rates(con) == SEEDED


def test_empty_rates_are_rejected(tmp_path):
    con = core.connect(tmp_path / "t.db")
    with pytest.raises(ValueError, match="At least one"):
        core.save_rates(con, [])
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest check.py -v`
Expected: the new tests FAIL with `AttributeError: module 'core' has no attribute 'get_rates'`.

- [ ] **Step 3: Implement**

Append to `core.py`:

```python
def get_rates(con):
    """All rates, STD first, oldest first. Percentages are fractions (0.075 = 7.5%)."""
    return [dict(r) for r in con.execute(
        "SELECT carton_type, effective_from, pacd_pct, matalan_pct FROM rates "
        "ORDER BY carton_type DESC, effective_from")]


def save_rates(con, rates):
    """Replace every rate. Raises ValueError, and saves nothing, if a row is invalid."""
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
    with con:
        con.execute("DELETE FROM rates")
        con.executemany("INSERT INTO rates VALUES (?, ?, ?, ?)", rows)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest check.py -v`
Expected: 22 passed.

---

### Task 4: Report queries and the xlsx reports (Summary, Tracker Raw Data, Shipment Tracking export)

**Files:**
- Modify: `core.py` (append queries)
- Create: `reports.py`
- Modify: `check.py` (add `import reports`, append fixture and tests)

**Interfaces:**
- Consumes: `core.connect`, `core.read_workbook`, `core.save`, `core.HEADERS`, `core.COLUMNS`, `core.DATE_COLS`, `core.SHEETS` (Tasks 1–2).
- Produces:
  - `core.QUARTER_MONTHS`, `core.quarter_label(fiscal_year, quarter) -> str` (e.g. `"Q3 2026 (Jul-Aug)"`).
  - `core.rows(con, kind, supplier=None, fiscal_year=None, quarters=None) -> list[dict]` (dated view rows, sorted by shipment_date then id).
  - `core.suppliers(con) -> list[str]`, `core.fiscal_years(con) -> list[int]`, `core.quarters(con, supplier=None, fiscal_year=None) -> list[int]`.
  - `core.date_range(con, supplier=None, fiscal_year=None, quarters=None) -> tuple[date, date] | None`.
  - `core.quarter_totals(con, kind, supplier, fiscal_year) -> dict[int, tuple]` (qty, total, pacd, rebate sums).
  - `core.rate_headers(rows) -> tuple[str, str]` (PACD header, rebate header).
  - `reports.safe_name(s) -> str`, `reports.SUMMARY_PREFIX`, `reports.KINDS`, `reports.TRACKER_COLS`, `reports.RAW_COLS`, `reports._headers(rows, cols)`.
  - `reports.summary_xlsx(con, fiscal_year)`, `reports.raw_data_xlsx(con, supplier, fiscal_year, quarter)`, `reports.shipment_tracking_xlsx(con, supplier)`: each returns `(file name, bytes)` or `None` when there is no data.

- [ ] **Step 1: Write the failing tests**

In `check.py`, add `import reports` on the line after `import core`. Then append:

```python
@pytest.fixture
def demo_con(tmp_path):
    con = core.connect(tmp_path / "demo.db")
    core.save(con, core.read_workbook(DEMO), DEMO.name)
    return con


def test_queries(demo_con):
    assert core.suppliers(demo_con) == [TW, UG]
    assert core.fiscal_years(demo_con) == [2026]
    assert core.quarters(demo_con, supplier=TW, fiscal_year=2026) == [3]
    assert core.date_range(demo_con) == (date(2026, 8, 1), date(2026, 8, 31))
    assert core.date_range(demo_con, fiscal_year=2027) is None
    assert len(core.rows(demo_con, "STD", supplier=TW, fiscal_year=2026, quarters=[3])) == 55
    assert len(core.rows(demo_con, "CUS", supplier=TW)) == 91
    dates = [r["shipment_date"] for r in core.rows(demo_con, "CUS", supplier=UG)]
    assert dates == sorted(dates)
    qty, total, pacd, rebate = core.quarter_totals(demo_con, "STD", TW, 2026)[3]
    assert qty == sum(r["carton_qty"] for r in core.rows(demo_con, "STD", supplier=TW))
    assert core.quarter_label(2026, 1) == "Q1 2026 (Dec-Mar)"


def test_rate_headers():
    one = [{"pacd_pct": 0.075, "matalan_pct": 0.05}] * 2
    mixed = [{"pacd_pct": 0.075, "matalan_pct": 0.05}, {"pacd_pct": 0.075, "matalan_pct": 0.10}]
    assert core.rate_headers(one) == ("7.5% for PACD", "5% Rebates for Matalan")
    assert core.rate_headers(mixed) == ("7.5% for PACD", "Rebates for Matalan")
    assert core.rate_headers([]) == ("PACD", "Rebates for Matalan")


def test_safe_name():
    assert reports.safe_name('A/B: "C". ') == "AB C"
    assert reports.safe_name(TW) == "Transworld Packaging (BD) Ltd"


def test_summary_xlsx(demo_con):
    name, data = reports.summary_xlsx(demo_con, 2026)
    assert name == "Matalan Carton Shipment Tracking_Summary_Aug-Aug 26.xlsx"
    rows = [[c.value for c in r] for r in openpyxl.load_workbook(BytesIO(data)).active.iter_rows()]
    assert rows[0][0] == TW
    assert rows[1] == ["Quarter", "Carton Type", "Carton QTY", "Total Amount", "PACD", "Rebates for Matalan"]
    assert rows[2][:2] == ["Q3", "Standard"] and rows[3][:2] == ["Q3", "Custom"] and rows[4][:2] == ["Q3", "Total"]
    assert rows[4][3] == pytest.approx(rows[2][3] + rows[3][3], abs=1e-9)  # Total adds the rounded values
    grand = [r for r in rows if r[0] == "Grand Total"]
    assert len(grand) == 2
    assert sum(r[2] for r in grand) == 28001 + 41529
    assert sum(r[3] for r in grand) == pytest.approx(33646.9152 + 37815.6551, abs=0.05)
    assert sum(r[5] for r in grand) == pytest.approx(1682.3458 + 1890.7828, abs=0.05)
    assert reports.summary_xlsx(demo_con, 2027) is None


def test_raw_data_xlsx(demo_con):
    name, data = reports.raw_data_xlsx(demo_con, TW, 2026, 3)
    assert name == "Tracker Raw Data_Transworld Packaging (BD) Ltd_Q3_2026.xlsx"
    ws = openpyxl.load_workbook(BytesIO(data)).active
    assert ws.title == "Q3 2026"
    assert ws["A1"].value == "Standard Carton"
    assert [c.value for c in ws[2]][21:24] == ["4% for PACD", "5% Rebates for Matalan", "NOTE"]
    assert ws["A59"].value == "Custom Carton"  # 55 STD rows on 3-57, blank row 58
    assert [c.value for c in ws[60]][21:23] == ["7.5% for PACD", "5% Rebates for Matalan"]
    assert ws.max_row == 151  # 91 CUS rows on 61-151
    assert isinstance(ws["B3"].value, datetime)
    assert reports.raw_data_xlsx(demo_con, TW, 2026, 1) is None


def test_shipment_tracking_xlsx(demo_con):
    name, data = reports.shipment_tracking_xlsx(demo_con, TW)
    assert name == "Matalan Carton Shipment Tracking_NPS_BD Aug-26 to Aug-26 (Transworld).xlsx"
    wb = openpyxl.load_workbook(BytesIO(data))
    assert wb.sheetnames == ["TRACK_STD", "TRACK_CUS"]
    assert wb["TRACK_STD"].max_row == 56 and wb["TRACK_CUS"].max_row == 92
    assert wb["TRACK_STD"]["A1"].value == "PO Received Date"
    assert reports.shipment_tracking_xlsx(demo_con, "Nobody") is None
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `python -m pytest check.py -v`
Expected: collection error `ModuleNotFoundError: No module named 'reports'`.

- [ ] **Step 3: Add the queries to `core.py`**

Append to `core.py`:

```python
QUARTER_MONTHS = {1: "Dec-Mar", 2: "Apr-Jun", 3: "Jul-Aug", 4: "Sep-Nov"}


def quarter_label(fiscal_year, quarter):
    return f"Q{quarter} {fiscal_year} ({QUARTER_MONTHS[quarter]})"


def _where(supplier=None, fiscal_year=None, quarters=None):
    sql, params = ["shipment_date IS NOT NULL"], []
    if supplier is not None:
        sql.append("packaging_supplier = ?")
        params.append(supplier)
    if fiscal_year is not None:
        sql.append("fiscal_year = ?")
        params.append(fiscal_year)
    if quarters:
        sql.append(f"quarter IN ({', '.join('?' * len(quarters))})")
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


def suppliers(con):
    src, params = _dated()
    return [r[0] for r in con.execute(
        f"SELECT DISTINCT packaging_supplier FROM {src} WHERE packaging_supplier IS NOT NULL ORDER BY 1", params)]


def fiscal_years(con):
    src, params = _dated()
    return [r[0] for r in con.execute(f"SELECT DISTINCT fiscal_year FROM {src} ORDER BY 1", params)]


def quarters(con, **filters):
    src, params = _dated(**filters)
    return [r[0] for r in con.execute(f"SELECT DISTINCT quarter FROM {src} ORDER BY 1", params)]


def date_range(con, **filters):
    """(first, last) shipment date of the matching rows, or None."""
    src, params = _dated(**filters)
    first, last = con.execute(f"SELECT MIN(shipment_date), MAX(shipment_date) FROM {src}", params).fetchone()
    return (date.fromisoformat(first), date.fromisoformat(last)) if first else None


def quarter_totals(con, kind, supplier, fiscal_year):
    """{quarter: (qty, total amount, pacd, matalan rebate)} for one supplier and fiscal year."""
    where, params = _where(supplier=supplier, fiscal_year=fiscal_year)
    return {r[0]: tuple(r[1:]) for r in con.execute(
        f"SELECT quarter, SUM(carton_qty), SUM(total_amount), SUM(pacd), SUM(matalan_rebate) "
        f"FROM v_{kind.lower()} WHERE {where} GROUP BY quarter", params)}


def rate_headers(rows):
    """Headers for the PACD and rebate columns; they show the % when every row has the same rate."""
    def label(key, after_pct, plain):
        pcts = {r[key] for r in rows if r[key] is not None}
        return f"{round(pcts.pop() * 100, 4):g}% {after_pct}" if len(pcts) == 1 else plain
    return label("pacd_pct", "for PACD", "PACD"), label("matalan_pct", "Rebates for Matalan", "Rebates for Matalan")
```

- [ ] **Step 4: Create `reports.py` with the xlsx reports**

```python
"""Report files built from the tracker database.

Each builder returns (file name, bytes), or None when the selection has no rows.
"""
from datetime import date
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

import core

KINDS = (("STD", "Standard"), ("CUS", "Custom"))
TRACKER_COLS = [c for c in core.COLUMNS if c != "note"] + ["pacd", "matalan_rebate"]  # sheet columns A-W
RAW_COLS = TRACKER_COLS + ["note"]  # A-X
RAW_WIDTHS = [15.6, 13.4, 18.4, 23.1, 23.1, 28.2, 23.4, 9.0, 10.1, 20.5, 7.9, 14.8,
              11.0, 10.4, 10.8, 6.6, 6.6, 9.5, 12.0, 18.5, 11.8, 12.6, 20.0, 11.2]
XLSX_FORMATS = {
    "po_received_date": "mm-dd-yy",
    "shipment_date": "mm-dd-yy",
    **dict.fromkeys(["length_mm", "width_mm", "height_mm", "l_mm", "w_mm"], "#,##0"),
    "area_sqm": "#,##0.000",
    **dict.fromkeys(["unit_price", "final_price", "total_amount", "pacd", "matalan_rebate"], "$#,##0.00"),
}
SUMMARY_PREFIX = "Matalan Carton Shipment Tracking_Summary_"
SUMMARY_HEADERS = ["Quarter", "Carton Type", "Carton QTY", "Total Amount", "PACD", "Rebates for Matalan"]

BOLD = Font(bold=True)
HEAD_FONT = Font(bold=True, color="1F497D")
HEAD_FILL = PatternFill("solid", fgColor="B4C6E7")
BANNER_FONT = Font(bold=True, color="FFFFFF")
BANNER_FILL = PatternFill("solid", fgColor="4472C4")
TOTAL_FILL = PatternFill("solid", fgColor="DCE6F1")


def safe_name(s):
    """Text usable in a Windows file or folder name."""
    for ch in '\\/:*?"<>|':
        s = s.replace(ch, "")
    return s.strip().rstrip(". ") or "Unknown"


def _headers(rows, cols):
    pacd, rebate = core.rate_headers(rows)
    return [{"pacd": pacd, "matalan_rebate": rebate}.get(c) or core.HEADERS[c] for c in cols]


def _bytes(wb):
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _banner(ws, row, text, width):
    cell = ws.cell(row, 1, text)
    cell.font, cell.fill, cell.alignment = BANNER_FONT, BANNER_FILL, Alignment(horizontal="left")
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=width)


def _write_table(ws, row, rows, cols):
    """Header at `row`, data rows below it. Returns the first free row."""
    for i, h in enumerate(_headers(rows, cols), 1):
        cell = ws.cell(row, i, h)
        cell.font, cell.fill = HEAD_FONT, HEAD_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r, values in enumerate(rows, row + 1):
        for i, col in enumerate(cols, 1):
            value = values[col]
            if col in core.DATE_COLS and value:
                value = date.fromisoformat(value)
            cell = ws.cell(r, i, value)
            if col in XLSX_FORMATS:
                cell.number_format = XLSX_FORMATS[col]
    return row + len(rows) + 1


def _summary_row(ws, row, quarter, label, values, bold=False):
    cells = [ws.cell(row, 1, quarter), ws.cell(row, 2, label)]
    cells += [ws.cell(row, 3 + i, round(v, 2) if i else v) for i, v in enumerate(values)]
    cells[2].number_format = "#,##0"
    for cell in cells[3:]:
        cell.number_format = "#,##0.00"
    if bold:
        for cell in cells:
            cell.font, cell.fill = BOLD, TOTAL_FILL


def summary_xlsx(con, fiscal_year):
    """Per packaging supplier: Standard / Custom / Total per quarter, then a Grand Total."""
    span = core.date_range(con, fiscal_year=fiscal_year)
    if not span:
        return None
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    row = 1
    for supplier in core.suppliers(con):
        totals = {kind: core.quarter_totals(con, kind, supplier, fiscal_year) for kind, _ in KINDS}
        quarters = sorted(set(totals["STD"]) | set(totals["CUS"]))
        if not quarters:
            continue
        ws.cell(row, 1, supplier).font = BOLD
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=6)
        for i, h in enumerate(SUMMARY_HEADERS, 1):
            ws.cell(row + 1, i, h).font = BOLD
        row += 2
        grand = [0, 0, 0, 0]
        for q in quarters:
            q_total = [0, 0, 0, 0]
            for kind, label in KINDS:
                qty, amount, pacd, rebate = totals[kind].get(q, (0, 0, 0, 0))
                values = [int(qty or 0)] + [round(v or 0, 2) for v in (amount, pacd, rebate)]
                _summary_row(ws, row, f"Q{q}", label, values)
                q_total = [a + b for a, b in zip(q_total, values)]
                row += 1
            _summary_row(ws, row, f"Q{q}", "Total", q_total, bold=True)
            grand = [a + b for a, b in zip(grand, q_total)]
            row += 1
        _summary_row(ws, row, "Grand Total", None, grand, bold=True)
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
        row += 2
    _widths(ws, [12, 14, 12, 14, 12, 20])
    first, last = span
    return f"{SUMMARY_PREFIX}{first:%b}-{last:%b} {last:%y}.xlsx", _bytes(wb)


def raw_data_xlsx(con, supplier, fiscal_year, quarter):
    """One quarter for one supplier: a Standard Carton block, then a Custom Carton block."""
    blocks = [(label, core.rows(con, kind, supplier=supplier, fiscal_year=fiscal_year, quarters=[quarter]))
              for kind, label in KINDS]
    if not any(rows for _, rows in blocks):
        return None
    wb = Workbook()
    ws = wb.active
    ws.title = f"Q{quarter} {fiscal_year}"
    row = 1
    for label, rows in blocks:
        if rows:
            _banner(ws, row, f"{label} Carton", len(RAW_COLS))
            row = _write_table(ws, row + 1, rows, RAW_COLS) + 1
    _widths(ws, RAW_WIDTHS)
    return f"Tracker Raw Data_{safe_name(supplier)}_Q{quarter}_{fiscal_year}.xlsx", _bytes(wb)


def shipment_tracking_xlsx(con, supplier):
    """All dated rows of one supplier as TRACK_STD and TRACK_CUS sheets, values only."""
    span = core.date_range(con, supplier=supplier)
    if not span:
        return None
    wb = Workbook()
    wb.remove(wb.active)
    for kind, sheet in core.SHEETS.items():
        ws = wb.create_sheet(sheet)
        _write_table(ws, 1, core.rows(con, kind, supplier=supplier), RAW_COLS)
        _widths(ws, RAW_WIDTHS)
        ws.freeze_panes = "A2"
    first, last = span
    short = safe_name(supplier.split()[0])
    return f"Matalan Carton Shipment Tracking_NPS_BD {first:%b-%y} to {last:%b-%y} ({short}).xlsx", _bytes(wb)
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `python -m pytest check.py -v`
Expected: 28 passed.

- [ ] **Step 6: Eyeball the xlsx output**

```powershell
python -c "import os, core, reports, pathlib; con = core.connect(':memory:'); core.save(con, core.read_workbook('samples/Matalan Carton Shipment Tracking_NPS_BD Aug-26.xlsx'), 'demo'); [pathlib.Path(os.environ['TEMP'], n).write_bytes(d) for n, d in (reports.summary_xlsx(con, 2026), reports.raw_data_xlsx(con, 'Transworld Packaging (BD) Ltd.', 2026, 3))]"
Start-Process "$env:TEMP\Tracker Raw Data_Transworld Packaging (BD) Ltd_Q3_2026.xlsx"
```

Expected: the file opens in Excel with a blue "Standard Carton" banner, light-blue header row, dates as mm-dd-yy and money as `$#,##0.00`, matching the look of `Reports/August/Uniglory Packaging Industries/Tracker Raw Data_Uniglory Packaging Industries_Q3_2026.xlsx`. Close Excel afterwards.

---

### Task 5: Full Tracker PDF (PacD and Matalan)

**Files:**
- Create: `assets/logo_left.png`, `assets/logo_right.png`
- Modify: `reports.py` (add imports, append PDF code)
- Modify: `check.py` (add `import pypdf`, append tests)

**Interfaces:**
- Consumes: `core.rows`, `core.quarter_label`, `reports.safe_name`, `reports.KINDS`, `reports.TRACKER_COLS`, `reports._headers` (Task 4).
- Produces: `reports.full_tracker_pdf(con, supplier, fiscal_year, quarters, brand) -> (str, bytes) | None`; `brand` is `"PacD"` (with PACD column) or `"Matalan"` (without).

- [ ] **Step 1: Extract the logos from the xlsm**

In the xlsm's `Assets` sheet, drawing `LogoLeft` points to `xl/media/image3.png` and `LogoRight` to `xl/media/image4.png`. The xlsm may be open in Excel, so read a copy:

```powershell
New-Item -ItemType Directory -Force assets | Out-Null
Copy-Item "eBiz Report Generator.xlsm" "$env:TEMP\gen_copy.xlsm"
python -c "import os, zipfile; z = zipfile.ZipFile(os.path.join(os.environ['TEMP'], 'gen_copy.xlsm')); open('assets/logo_left.png', 'wb').write(z.read('xl/media/image3.png')); open('assets/logo_right.png', 'wb').write(z.read('xl/media/image4.png'))"
python -c "import struct; [print(n, struct.unpack('>II', open(f'assets/{n}', 'rb').read()[16:24])) for n in ('logo_left.png', 'logo_right.png')]"
Remove-Item "$env:TEMP\gen_copy.xlsm"
```

Expected: `logo_left.png (458, 154)` and `logo_right.png (392, 156)`.

- [ ] **Step 2: Write the failing tests**

In `check.py`, add `import pypdf` on the line after `import pytest`. Append:

```python
def pdf_text(data):
    return "\n".join(p.extract_text() for p in pypdf.PdfReader(BytesIO(data)).pages)


def test_full_tracker_pdf_pacd(demo_con):
    name, data = reports.full_tracker_pdf(demo_con, TW, 2026, [3], "PacD")
    assert name == "Full Tracker Sheet for PacD_Transworld Packaging (BD) Ltd_Q3_2026.pdf"
    text = pdf_text(data)
    total = sum(r["total_amount"] for kind in ("STD", "CUS") for r in core.rows(demo_con, kind, supplier=TW))
    for expected in ("Full Tracker Report (Transworld Packaging (BD) Ltd.)", "Standard Carton",
                     "Q3 2026 (Jul-Aug) Quarter", "Q3 2026 (Jul-Aug) Subtotal", "Standard Carton Total",
                     "Custom Carton Total", "Tracker Grand Total", f"{total:,.2f}", "PACD",
                     "Page 1 of", "Generated:"):
        assert expected in text, expected
    assert len(pypdf.PdfReader(BytesIO(data)).pages[0].images) == 2  # both logos


def test_full_tracker_pdf_matalan_has_no_pacd(demo_con):
    name, data = reports.full_tracker_pdf(demo_con, UG, 2026, [1, 2, 3], "Matalan")
    assert name == "Full Tracker Sheet for Matalan_Uniglory Packaging Industries Ltd_Q1-Q3_2026.pdf"
    text = pdf_text(data)
    assert "PACD" not in text
    assert "Rebates" in text and "Tracker Grand Total" in text


def test_full_tracker_pdf_without_rows_is_none(demo_con):
    assert reports.full_tracker_pdf(demo_con, TW, 2026, [1], "PacD") is None
```

- [ ] **Step 3: Run tests to verify they fail**

Run: `python -m pytest check.py -v`
Expected: the 3 new tests FAIL with `AttributeError: module 'reports' has no attribute 'full_tracker_pdf'`.

- [ ] **Step 4: Implement the PDF**

In `reports.py`, replace the import block with:

```python
from datetime import date, datetime
from io import BytesIO
from itertools import groupby
from pathlib import Path
from xml.sax.saxutils import escape

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

import core
```

Append to `reports.py`:

```python
ASSETS = Path(__file__).with_name("assets")
PAGE = landscape(A4)
MARGIN = 0.3 * inch
NAVY, BLUE = colors.HexColor("#1F497D"), colors.HexColor("#4472C4")
HEAD_BG, QUARTER_BG = colors.HexColor("#B4C6E7"), colors.HexColor("#D9E1F2")
ALT_BG, SUBTOTAL_BG, TOTAL_BG = colors.HexColor("#EBF1FA"), colors.HexColor("#FFF2CC"), colors.HexColor("#DCE6F1")
CELL = ParagraphStyle("cell", fontName="Helvetica", fontSize=5.5, leading=6.5)
HEAD = ParagraphStyle("head", parent=CELL, fontName="Helvetica-Bold", textColor=NAVY, alignment=1)
SUM_COLS = ["carton_qty", "final_price", "total_amount", "pacd", "matalan_rebate"]  # summed like the VBA
BASE_WIDTHS = {
    "sr": 16, "po_received_date": 30, "shipment_date": 30, "matalan_po": 30, "supplier_name": 48,
    "factory_name": 48, "packaging_supplier": 48, "ebiz_customer_ref": 44, "ebiz_so": 30, "carton_qty": 26,
    "carton_style": 40, "flute": 26, "carton_selection": 30, "length_mm": 22, "width_mm": 22, "height_mm": 22,
    "l_mm": 22, "w_mm": 22, "area_sqm": 24, "unit_price": 28, "final_price": 30, "total_amount": 40,
    "pacd": 34, "matalan_rebate": 34,
}
PADDING = [("LEFTPADDING", (0, 0), (-1, -1), 1.5), ("RIGHTPADDING", (0, 0), (-1, -1), 1.5),
           ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]


def _fmt(col, value):
    """PDF cell content: a string for numbers and dates, a wrapping Paragraph for text."""
    if value is None:
        return ""
    if col in core.DATE_COLS:
        d = date.fromisoformat(value)
        return f"{d.month}/{d.day}/{d.year}"
    if col in ("carton_qty", "length_mm", "width_mm", "height_mm", "l_mm", "w_mm"):
        return f"{value:,.0f}"
    if col == "area_sqm":
        return f"{value:.3f}"
    if col == "unit_price":
        return f"{value:.2f}"
    if col in ("final_price", "total_amount", "pacd", "matalan_rebate"):
        return f"$ {value:,.2f}"
    return Paragraph(escape(str(value)), CELL)


def _total_line(text, rows, cols):
    line = [text] + [""] * len(cols)
    for col in SUM_COLS:
        if col in cols:
            line[cols.index(col) + 1] = _fmt(col, sum(r[col] or 0 for r in rows))
    return line


def _block(label, rows, cols, widths):
    """One carton type: banners, header, per-quarter rows with subtotals, block total."""
    qty_i = cols.index("carton_qty") + 1
    data = []
    style = PADDING + [
        ("FONT", (0, 0), (-1, -1), "Helvetica", 5.5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 2), (-1, -1), 0.25, HEAD_BG),
        ("ALIGN", (qty_i, 3), (-1, -1), "RIGHT"),
        ("BACKGROUND", (0, 2), (-1, 2), HEAD_BG),
    ]

    def banner(text, background, size, color=colors.white):
        r = len(data)
        data.append([text] + [""] * len(cols))
        style.extend([("SPAN", (0, r), (-1, r)), ("BACKGROUND", (0, r), (-1, r), background),
                      ("TEXTCOLOR", (0, r), (-1, r), color), ("FONT", (0, r), (-1, r), "Helvetica-Bold", size)])

    def total(text, total_rows, background):
        r = len(data)
        data.append(_total_line(text, total_rows, cols))
        style.extend([("SPAN", (0, r), (qty_i - 1, r)), ("BACKGROUND", (0, r), (-1, r), background),
                      ("FONT", (0, r), (-1, r), "Helvetica-Bold", 5.5), ("ALIGN", (0, r), (-1, r), "RIGHT")])

    banner(f"Full Tracker Table - {label}", NAVY, 9)
    banner(f"{label} Carton", BLUE, 7)
    data.append([Paragraph("Sr No.", HEAD)] + [Paragraph(escape(h), HEAD) for h in _headers(rows, cols)])
    sr = 0
    for (year, quarter), q_rows in groupby(rows, key=lambda r: (r["fiscal_year"], r["quarter"])):
        q_rows = list(q_rows)
        q_label = core.quarter_label(year, quarter)
        banner(f"{q_label} Quarter", QUARTER_BG, 6, NAVY)
        for i, r in enumerate(q_rows):
            sr += 1
            data.append([str(sr)] + [_fmt(c, r[c]) for c in cols])
            if i % 2:
                style.append(("BACKGROUND", (0, len(data) - 1), (-1, len(data) - 1), ALT_BG))
        total(f"{q_label} Subtotal", q_rows, SUBTOTAL_BG)
    total(f"{label} Carton Total", rows, TOTAL_BG)
    return Table(data, colWidths=widths, repeatRows=3, style=TableStyle(style))


class _NumberedCanvas(rl_canvas.Canvas):
    """Canvas that adds 'Page x of y' once the page count is known."""

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._pages = []

    def showPage(self):
        self._pages.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        for page in self._pages:
            self.__dict__.update(page)
            self.setFont("Helvetica", 7)
            self.drawRightString(PAGE[0] - MARGIN, 0.2 * inch, f"Page {self._pageNumber} of {len(self._pages)}")
            super().showPage()
        super().save()


def full_tracker_pdf(con, supplier, fiscal_year, quarters, brand):
    """Full Tracker Sheet for PacD (brand "PacD") or for Matalan (brand "Matalan": no PACD column)."""
    cols = TRACKER_COLS if brand == "PacD" else [c for c in TRACKER_COLS if c != "pacd"]
    blocks = [(label, core.rows(con, kind, supplier=supplier, fiscal_year=fiscal_year, quarters=quarters))
              for kind, label in KINDS]
    all_rows = [r for _, rows in blocks for r in rows]
    if not all_rows:
        return None
    avail = PAGE[0] - 2 * MARGIN
    base = [BASE_WIDTHS["sr"]] + [BASE_WIDTHS[c] for c in cols]
    widths = [w * avail / sum(base) for w in base]
    title = f"Full Tracker Report ({supplier})"
    story = [Table([[title]], colWidths=[avail], style=TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NAVY), ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
        ("FONT", (0, 0), (-1, -1), "Helvetica-Bold", 14), ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)])), Spacer(1, 6)]
    for label, rows in blocks:
        if rows:
            story += [_block(label, rows, cols, widths), Spacer(1, 8)]
    qty_i = cols.index("carton_qty") + 1
    story.append(Table([_total_line("Tracker Grand Total", all_rows, cols)], colWidths=widths,
                       style=TableStyle(PADDING + [
                           ("SPAN", (0, 0), (qty_i - 1, 0)), ("BACKGROUND", (0, 0), (-1, -1), NAVY),
                           ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
                           ("FONT", (0, 0), (-1, -1), "Helvetica-Bold", 6), ("ALIGN", (0, 0), (-1, -1), "RIGHT")])))
    generated = datetime.now().strftime("Generated: %d-%b-%Y %I:%M %p")

    def decorate(c, doc):
        c.saveState()
        for name, right in (("logo_left.png", False), ("logo_right.png", True)):
            path = ASSETS / name
            if path.exists():
                img = ImageReader(str(path))
                iw, ih = img.getSize()
                h = 0.45 * inch
                w = iw * h / ih
                c.drawImage(img, PAGE[0] - MARGIN - w if right else MARGIN, PAGE[1] - 0.3 * inch - h, w, h,
                            mask="auto")
        c.setFont("Helvetica", 7)
        c.drawString(MARGIN, 0.2 * inch, generated)
        c.restoreState()

    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=PAGE, leftMargin=MARGIN, rightMargin=MARGIN,
                            topMargin=0.9 * inch, bottomMargin=0.45 * inch, title=title)
    doc.build(story, onFirstPage=decorate, onLaterPages=decorate, canvasmaker=_NumberedCanvas)
    qs = sorted(quarters)
    span = f"Q{qs[0]}" if len(qs) == 1 else f"Q{qs[0]}-Q{qs[-1]}"
    return f"Full Tracker Sheet for {brand}_{safe_name(supplier)}_{span}_{fiscal_year}.pdf", buf.getvalue()
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `python -m pytest check.py -v`
Expected: 31 passed.

- [ ] **Step 6: Eyeball the PDF against the old one**

```powershell
python -c "import os, core, reports, pathlib; con = core.connect(':memory:'); core.save(con, core.read_workbook('samples/Matalan Carton Shipment Tracking_NPS_BD Aug-26.xlsx'), 'demo'); n, d = reports.full_tracker_pdf(con, 'Uniglory Packaging Industries Ltd.', 2026, [3], 'PacD'); pathlib.Path(os.environ['TEMP'], n).write_bytes(d)"
Start-Process "$env:TEMP\Full Tracker Sheet for PacD_Uniglory Packaging Industries Ltd_Q3_2026.pdf"
Start-Process "Reports\August\Uniglory Packaging Industries\Tracker Sheet for PacD_Uniglory Packaging Industries_Q3_2026.pdf"
```

Expected: both open. Check the new one has the logos at top left and right, a navy title, both blocks, quarter header and subtotal rows, readable wrapped text, no clipped money values, and the footer on every page. If a money column clips, widen its entry in `BASE_WIDTHS` and re-run Step 5.

---

### Task 6: Report bundle and "Save all" to the Reports folder

**Files:**
- Modify: `reports.py` (append)
- Modify: `check.py` (append)

**Interfaces:**
- Consumes: `reports.summary_xlsx`, `reports.full_tracker_pdf`, `reports.raw_data_xlsx`, `reports.shipment_tracking_xlsx`, `reports.safe_name`, `reports.SUMMARY_PREFIX` (Tasks 4–5); `core.date_range` (Task 4).
- Produces:
  - `reports.build_all(con, supplier, fiscal_year, quarters) -> list[tuple[str, bytes]]` (Summary first; reports without data left out).
  - `reports.save_all(con, supplier, fiscal_year, quarters, reports_dir) -> list[Path]`.

- [ ] **Step 1: Write the failing test**

Append to `check.py`:

```python
def test_save_all(demo_con, tmp_path):
    root = tmp_path / "Reports"
    paths = reports.save_all(demo_con, TW, 2026, [3], root)
    folder = "August/Transworld Packaging (BD) Ltd/"
    assert sorted(p.relative_to(root).as_posix() for p in paths) == sorted([
        "August/Matalan Carton Shipment Tracking_Summary_Aug-Aug 26.xlsx",
        folder + "Full Tracker Sheet for PacD_Transworld Packaging (BD) Ltd_Q3_2026.pdf",
        folder + "Full Tracker Sheet for Matalan_Transworld Packaging (BD) Ltd_Q3_2026.pdf",
        folder + "Tracker Raw Data_Transworld Packaging (BD) Ltd_Q3_2026.xlsx",
        folder + "Matalan Carton Shipment Tracking_NPS_BD Aug-26 to Aug-26 (Transworld).xlsx",
    ])
    assert all(p.stat().st_size > 0 for p in paths)
    assert reports.save_all(demo_con, TW, 2026, [1], root) == []
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest check.py -v -k save_all`
Expected: FAIL with `AttributeError: module 'reports' has no attribute 'save_all'`.

- [ ] **Step 3: Implement**

Append to `reports.py`:

```python
def build_all(con, supplier, fiscal_year, quarters):
    """Every phase-1 report for one supplier as (file name, bytes), Summary first."""
    items = [summary_xlsx(con, fiscal_year)]
    items += [full_tracker_pdf(con, supplier, fiscal_year, quarters, brand) for brand in ("PacD", "Matalan")]
    items += [raw_data_xlsx(con, supplier, fiscal_year, q) for q in sorted(quarters)]
    items.append(shipment_tracking_xlsx(con, supplier))
    return [item for item in items if item]


def save_all(con, supplier, fiscal_year, quarters, reports_dir):
    """Write build_all() to <reports_dir>/<Month>/ (Summary) and <reports_dir>/<Month>/<supplier>/ (the rest).

    <Month> is the full month name of the latest shipment in the selection. Existing files are overwritten.
    """
    span = core.date_range(con, supplier=supplier, fiscal_year=fiscal_year, quarters=quarters)
    if not span:
        return []
    month_dir = Path(reports_dir) / f"{span[1]:%B}"
    supplier_dir = month_dir / safe_name(supplier)
    supplier_dir.mkdir(parents=True, exist_ok=True)
    paths = []
    for name, data in build_all(con, supplier, fiscal_year, quarters):
        path = (month_dir if name.startswith(SUMMARY_PREFIX) else supplier_dir) / name
        path.write_bytes(data)
        paths.append(path)
    return paths
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `python -m pytest check.py -v`
Expected: 32 passed.

---

### Task 7: Streamlit app and launcher

**Files:**
- Create: `app.py`
- Create: `run.bat`
- Modify: `check.py` (append)

**Interfaces:**
- Consumes: `core.connect`, `core.read_workbook`, `core.ImportFileError`, `core.preview`, `core.bad_dates`, `core.save`, `core.get_rates`, `core.save_rates`, `core.suppliers`, `core.fiscal_years`, `core.quarters`, `core.quarter_label` (Tasks 1–4); `reports.build_all`, `reports.save_all` (Task 6).
- Produces: the UI. Nothing imports `app.py`.

- [ ] **Step 1: Write the failing smoke test**

Append to `check.py`:

```python
def test_app_smoke(tmp_path, monkeypatch):
    from streamlit.testing.v1 import AppTest
    db = tmp_path / "app.db"
    monkeypatch.setenv("CARTON_DB", str(db))
    core.save(core.connect(db), core.read_workbook(DEMO), DEMO.name)
    at = AppTest.from_file(str(HERE / "app.py"), default_timeout=30).run()
    assert not at.exception
    assert at.selectbox[0].options == [TW, UG]
    assert len(at.tabs) == 3
```

- [ ] **Step 2: Run test to verify it fails**

Run: `python -m pytest check.py -v -k app_smoke`
Expected: FAIL (`app.py` does not exist).

- [ ] **Step 3: Create `app.py`**

```python
"""Matalan carton tracker UI. Start with run.bat (streamlit run app.py)."""
from pathlib import Path

import pandas as pd
import streamlit as st

import core
import reports

REPORTS_DIR = Path(__file__).with_name("Reports")
RATE_COLS = ["carton_type", "effective_from", "pacd_pct", "matalan_pct"]

st.set_page_config(page_title="Matalan Carton Tracker", layout="wide")
con = core.connect()
tab_import, tab_rates, tab_reports = st.tabs(["Import", "Rates", "Reports"])

with tab_import:
    upload = st.file_uploader("Tracker workbook (TRACK_STD and TRACK_CUS sheets)", type=["xlsx", "xlsm"])
    if upload:
        try:
            data = core.read_workbook(upload)
        except core.ImportFileError as e:
            st.error(str(e))
        else:
            st.dataframe(pd.DataFrame(core.preview(con, data)), hide_index=True)
            bad = core.bad_dates(data)
            if bad:
                st.warning(f"{len(bad)} row(s) have no valid Shipment Date. "
                           "They are imported but left out of the reports.")
                st.dataframe(pd.DataFrame(bad), hide_index=True)
            if st.button("Confirm import", type="primary"):
                counts = core.save(con, data, upload.name)
                st.success(f"Imported {counts['STD']} TRACK_STD and {counts['CUS']} TRACK_CUS rows.")

with tab_rates:
    st.caption("A rate applies to shipments on or after its 'From' date, "
               "until the next row of the same carton type.")
    df = pd.DataFrame(core.get_rates(con), columns=RATE_COLS)
    df["effective_from"] = pd.to_datetime(df["effective_from"]).dt.date
    df[["pacd_pct", "matalan_pct"]] *= 100
    edited = st.data_editor(df, num_rows="dynamic", hide_index=True, column_config={
        "carton_type": st.column_config.SelectboxColumn("Carton type", options=["STD", "CUS"], required=True),
        "effective_from": st.column_config.DateColumn("From", required=True),
        "pacd_pct": st.column_config.NumberColumn("PACD %", min_value=0, max_value=100, format="%.2f"),
        "matalan_pct": st.column_config.NumberColumn("Matalan rebate %", min_value=0, max_value=100,
                                                     format="%.2f"),
    })
    if st.button("Save rates"):
        try:
            core.save_rates(con, [{"carton_type": r.carton_type, "effective_from": str(r.effective_from),
                                   "pacd_pct": round(r.pacd_pct / 100, 6),
                                   "matalan_pct": round(r.matalan_pct / 100, 6)}
                                  for r in edited.itertuples()])
        except (ValueError, TypeError) as e:
            st.error(str(e))
        else:
            st.success("Rates saved. Reports use them from now on.")

with tab_reports:
    suppliers = core.suppliers(con)
    if not suppliers:
        st.info("No data yet. Import a tracker file first.")
    else:
        supplier = st.selectbox("Packaging supplier", suppliers)
        years = core.fiscal_years(con)
        year = st.selectbox("Fiscal year", years, index=len(years) - 1)
        available = core.quarters(con, supplier=supplier, fiscal_year=year)
        chosen = st.multiselect("Quarters", available, default=available,
                                format_func=lambda q: core.quarter_label(year, q))
        if not chosen:
            st.warning("Pick at least one quarter.")
        else:
            key = (supplier, year, tuple(chosen))
            build_col, save_col = st.columns(2)
            if build_col.button("Build reports", type="primary"):
                st.session_state.built = (key, reports.build_all(con, supplier, year, chosen))
            if save_col.button("Save all to Reports folder"):
                paths = reports.save_all(con, supplier, year, chosen, REPORTS_DIR)
                st.success("Saved:\n" + "\n".join(f"- {p}" for p in paths))
            built = st.session_state.get("built")
            if built and built[0] == key:
                if not built[1]:
                    st.info("No data for this selection.")
                for name, data in built[1]:
                    st.download_button(name, data, file_name=name, key=name)
```

- [ ] **Step 4: Create `run.bat`**

```bat
@echo off
cd /d "%~dp0"
python -m streamlit run app.py
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `python -m pytest check.py -v`
Expected: 33 passed.

---

### Task 8: End-to-end check in the running app

**Files:** none changed unless a defect is found.

- [ ] **Step 1: Full test run**

Run: `python -m pytest check.py -v`
Expected: 33 passed, 0 failed.

- [ ] **Step 2: Start the app headless and confirm it serves**

```powershell
$p = Start-Process python -ArgumentList "-m", "streamlit", "run", "app.py", "--server.headless", "true", "--server.port", "8501" -PassThru -WindowStyle Hidden
Start-Sleep -Seconds 8
(Invoke-WebRequest http://localhost:8501 -UseBasicParsing).StatusCode
```

Expected: `200`. This creates `carton.db` in the project folder with the seed rates.

- [ ] **Step 3: Walk through the UI**

Open http://localhost:8501 in a browser.
1. Import tab: upload `samples\Matalan Carton Shipment Tracking_NPS_BD Aug-26.xlsx`. The preview shows 4 groups (TRACK_STD and TRACK_CUS × Transworld, Uniglory, all 2026-08) with 55 / 33 / 91 / 27 rows and 0 replaced. Click Confirm import; the message says 88 TRACK_STD and 118 TRACK_CUS rows.
2. Upload the same file again: "Rows replaced in DB" now equals "Rows in file".
3. Rates tab: 4 rows, CUS 2026-07-01 shows 7.50 / 5.00. Change any % to 150 and save: an error appears and nothing changes. Reload the page to discard.
4. Reports tab: pick Transworld, 2026, Q3. Build reports: 5 download buttons. Save all: files appear under `Reports\August\` and `Reports\August\Transworld Packaging (BD) Ltd\`.

- [ ] **Step 4: Stop the app and remove the walkthrough output**

```powershell
Stop-Process -Id $p.Id
Remove-Item -Recurse "Reports\August\Transworld Packaging (BD) Ltd"
Remove-Item "Reports\August\Matalan Carton Shipment Tracking_Summary_Aug-Aug 26.xlsx"
Remove-Item carton.db
```

Expected: `Reports\August` holds only the original files again, and the next real start creates a fresh `carton.db`.

## Not in this plan

- eBiz import and the Rebate, Matalan Rebate, Discrepancy and eBiz Discrepancy reports (phase 2).
- Confirming the re-import rule against the real monthly or cumulative files.
- Checking Q1 and Q2 totals against `Reports/August/Matalan Carton Shipment Tracking_Summary_Dec-June 26.xlsx`. This needs the real files.
