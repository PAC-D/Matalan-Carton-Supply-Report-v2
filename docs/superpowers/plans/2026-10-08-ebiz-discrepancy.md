# eBiz Import, Discrepancy Report and Cross Reference Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Import eBiz exports into a new `ebiz` table (each upload replaces all eBiz data) and build, per packaging supplier, a Discrepancy Report PDF and a Cross Reference xlsx comparing all tracker rows with eBiz.

**Architecture:** `core.py` gains the `ebiz` table, the eBiz reader/replace functions and `cross_reference()`, a pure function that ports the old VBA matching and returns one plain dict. `reports.py` gains `discrepancy_pdf()` (reusing the Full Tracker PDF look) and `cross_reference_xlsx()` (reusing the Summary workbook style), plus a `discrepancy` option on `build_all` / `save_all`. `app.py` gets an eBiz section on the Import tab and a checkbox on the Reports tab.

**Tech Stack:** Python 3.14, sqlite3, openpyxl 3.1.5, pandas 3.0.5, reportlab 5.0.0, streamlit 1.61.1, pypdf (tests), pytest 9.1.1.

**Spec:** `docs/superpowers/specs/2026-10-08-ebiz-discrepancy-design.md` (extends `docs/superpowers/specs/2026-10-07-carton-tracker-design.md`).

## Global Constraints

- Project folder (all paths relative to it): `C:\Users\Shoaib\OneDrive - PacD\Projects\Matalan\Matalan Carton Supplu Report`. Use the **PowerShell** tool; the Bash tool fails on this PC (exit 53).
- No new dependencies.
- Git branch `feat/carton-tracker`. Commit after each task. Never push, amend or force. Every commit message ends with a blank line and `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`. In PowerShell commit with `git commit -q -m @'` … `'@` (closing `'@` at column 0); `git commit -F -` does not work here.
- Do not modify `eBiz Report Generator.xlsm` or `eBiz Report Generator_.xlsm`; copy to the scratchpad before reading. `carton.db` is the user's live database: never write to it from tests or scripts (tests use `tmp_path`).
- Tests: `python -m pytest check.py -q` (all, currently 64 passing; all must keep passing) or `python -m pytest check.py -q -k <name>`.
- Every eBiz upload deletes **all** existing eBiz rows, then inserts the uploaded files, in one transaction.
- Discrepancy and Cross Reference compare **all** tracker rows of the supplier (any date, including rows without a Shipment Date) with all eBiz rows whose `vendor_key` equals the supplier key. Fiscal year and quarter pickers don't apply.
- Matching: an eBiz row resolves to `so_1` if it is a tracker SO#, else `so_2` if it is a tracker SO# ("via SO_2"), else it is eBiz-only (key `so_1`, else `so_2`). `"0"` and blank mean no SO#.
- Tracker rebate per row = `pacd + matalan_rebate` (NULL = 0).
- File names: `Discrepancy Report_<supplier>_<dd-Mmm-yyyy>.pdf`, `Cross Reference_<supplier>_<dd-Mmm-yyyy>.xlsx` (generation date; `<supplier>` = `reports.safe_name(core.supplier_name(...))`).

## File Structure

| File | Change |
|---|---|
| `core.py` | `EBIZ_*` constants, `ebiz` table (SCHEMA_VERSION 3), `read_ebiz`, `ebiz_preview`, `ebiz_clashes`, `replace_ebiz`, `ebiz_status`, `ebiz_vendors`, `so_order`, `cross_reference` |
| `reports.py` | `_title_bar`, `_pdf` (extracted from `full_tracker_pdf`), `_bar`, `_grid`, `discrepancy_pdf`; `_title_block`, `_print_setup` (extracted from `summary_xlsx`), `cross_reference_xlsx`; `discrepancy=` on `build_all` / `save_all` |
| `app.py` | eBiz section on Import tab; checkbox on Reports tab |
| `check.py` | helpers `make_ebiz`, `erow`, fixture `xref_con`, new tests |
| `docs/superpowers/specs/2026-10-08-ebiz-discrepancy-design.md` | Cross Reference is A4 portrait (Task 4) |

---

### Task 1: eBiz table and import functions

**Files:**
- Modify: `core.py`
- Test: `check.py` (append at the end)

**Interfaces:**
- Consumes: existing `core._norm`, `core._clean`, `core.supplier_key`, `core.ImportFileError`, `core.TABLES`.
- Produces:
  - `core.EBIZ_HEADERS: dict[str, str]` (db column → header text), `core.EBIZ_REQUIRED: list[str]`, `core.EBIZ_REAL: set[str]`, `core.EBIZ_COLUMNS: list[str]` (= `EBIZ_HEADERS` keys + `so_1`, `so_2`, `vendor_key`).
  - `core.read_ebiz(src) -> list[dict]` (keys = `EBIZ_COLUMNS`); raises `ImportFileError`.
  - `core.ebiz_preview(files: dict[str, list[dict]]) -> list[dict]` with keys `File`, `Vendor`, `Spellings in file`, `Rows`, `Qty Shipped`, `Total Rebate`.
  - `core.ebiz_clashes(files) -> list[dict]` with keys `Vendor`, `Files`.
  - `core.replace_ebiz(con, files) -> dict[str, int]`; raises `ValueError` on clash.
  - `core.ebiz_status(con) -> (int, list[str], str) | None`.
  - `core.ebiz_vendors(con) -> list[str]`.
  - `core._so(text) -> str | None` (used again in Task 2).

- [ ] **Step 1: Write the failing tests** — append to the end of `check.py`:

```python
def make_ebiz(path, rows, header_row=1, drop=None):
    """Write an eBiz export: headers in header_row, rows (dicts keyed by core.EBIZ_HEADERS names) below."""
    wb = openpyxl.Workbook()
    ws = wb.active
    cols = [c for c in core.EBIZ_HEADERS if c != drop]
    for i, c in enumerate(cols, 1):
        ws.cell(header_row, i, core.EBIZ_HEADERS[c])
    for r, row in enumerate(rows, header_row + 1):
        for i, c in enumerate(cols, 1):
            ws.cell(r, i, row.get(c))
    wb.save(path)
    return path


def erow(order="7001", qty=100, rebate=1.0, vendor="Pack A Ltd.", **kw):
    """One eBiz export row for make_ebiz."""
    return {"vendor": vendor, "vendor_code": "M0001", "product": "MLN-CARTONA1", "order_number": order,
            "qty_shipped": qty, "net_price_1000": 840, "gross_price_1000": 880, "total_rebate": rebate,
            "currency": "USD", **kw}


def test_read_ebiz_splits_order_numbers(tmp_path):
    path = make_ebiz(tmp_path / "e.xlsx", [erow("7001|7002"), erow(7003.0, qty=5), erow(None, qty=None),
                                           erow(" 7004 | ")], header_row=2)
    rows = core.read_ebiz(path)
    assert [(r["order_number"], r["so_1"], r["so_2"]) for r in rows] == [
        ("7001|7002", "7001", "7002"), ("7003", "7003", None), ("7004 |", "7004", None)]
    assert rows[0]["vendor_key"] == "PACK"
    assert rows[1]["qty_shipped"] == 5.0 and rows[1]["currency"] == "USD"


def test_read_ebiz_needs_its_headers(tmp_path):
    with pytest.raises(core.ImportFileError, match="header row not found"):
        core.read_ebiz(make_ebiz(tmp_path / "a.xlsx", [erow()], drop="order_number"))
    with pytest.raises(core.ImportFileError, match="Qty Shipped"):
        core.read_ebiz(make_ebiz(tmp_path / "b.xlsx", [erow()], drop="qty_shipped"))
    assert core.read_ebiz(make_ebiz(tmp_path / "c.xlsx", [erow()], drop="currency"))[0]["currency"] is None


def test_ebiz_preview(tmp_path):
    rows = core.read_ebiz(make_ebiz(tmp_path / "a.xlsx", [erow("1", qty=10, rebate=1.5),
                                                          erow("2", qty=5, rebate=0.25, vendor="PACK A LIMITED")]))
    assert core.ebiz_preview({"a.xlsx": rows}) == [{
        "File": "a.xlsx", "Vendor": "PACK", "Spellings in file": "PACK A LIMITED / Pack A Ltd.", "Rows": 2,
        "Qty Shipped": 15.0, "Total Rebate": 1.75}]


def test_ebiz_upload_replaces_all_ebiz_data(tmp_path):
    con = core.connect(tmp_path / "t.db")
    assert core.ebiz_status(con) is None
    a = {"a.xlsx": core.read_ebiz(make_ebiz(tmp_path / "a.xlsx", [erow("1"), erow("2", vendor="Other Co")]))}
    assert core.replace_ebiz(con, a) == {"a.xlsx": 2}
    assert core.ebiz_vendors(con) == ["OTHER", "PACK"]
    b = {"b.xlsx": core.read_ebiz(make_ebiz(tmp_path / "b.xlsx", [erow("3")]))}
    core.replace_ebiz(con, b)
    assert [r[0] for r in con.execute("SELECT order_number FROM ebiz")] == ["3"]
    count, files, imported_at = core.ebiz_status(con)
    assert (count, files) == (1, ["b.xlsx"]) and imported_at
    clash = {"c.xlsx": b["b.xlsx"], "d.xlsx": b["b.xlsx"]}
    assert core.ebiz_clashes(clash) == [{"Vendor": "PACK", "Files": "c.xlsx / d.xlsx"}]
    with pytest.raises(ValueError, match="PACK"):
        core.replace_ebiz(con, clash)
    assert core.ebiz_status(con)[0] == 1  # a refused upload keeps the old data


def test_old_database_gets_the_ebiz_table(tmp_path):
    con = core.connect(tmp_path / "t.db")
    insert(con, "STD", shipment_date="2026-08-03", carton_qty=1, unit_price=1.0, total_amount=1.0)
    con.execute("DROP TABLE ebiz")
    con.execute("PRAGMA user_version = 2")
    con.close()
    con = core.connect(tmp_path / "t.db")
    assert core.ebiz_status(con) is None
    assert con.execute("SELECT COUNT(*) FROM track_std").fetchone()[0] == 1
```

- [ ] **Step 2: Run the tests to verify they fail**

Run: `python -m pytest check.py -q -k "ebiz"`
Expected: FAIL, `AttributeError: module 'core' has no attribute 'EBIZ_HEADERS'` (or `read_ebiz`).

- [ ] **Step 3: Implement** in `core.py`.

3a. Change `SCHEMA_VERSION = 2` to `SCHEMA_VERSION = 3`.

3b. Below the `LEAD_DAYS = ...` line add:

```python

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
```

3c. In `_schema()`, directly before the line `for kind, table in TABLES.items():` add:

```python
    ebiz_cols = ",\n  ".join(f"{c} {'REAL' if c in EBIZ_REAL else 'TEXT'}" for c in EBIZ_COLUMNS)
    sql.append(f"CREATE TABLE IF NOT EXISTS ebiz (\n  id INTEGER PRIMARY KEY,\n  {ebiz_cols},\n"
               f"  source_file TEXT,\n  imported_at TEXT\n);")
```

3d. Directly above `def get_rates(con):` add:

```python
def _so(text):
    """An SO number as text, or None for blank and '0'."""
    text = (text or "").strip()
    return text if text and text != "0" else None


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

    Order Number "A|B" gives so_1 "A" and so_2 "B"; vendor_key is supplier_key(vendor).
    """
    try:
        wb = openpyxl.load_workbook(src, data_only=True)
    except Exception as e:  # not an xlsx, corrupt or password protected
        raise ImportFileError(f"Can't open the workbook: {e}") from e
    ws = wb.worksheets[0]
    for n, header_row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), start=1):
        found = {_norm(h): i for i, h in enumerate(header_row)}
        if _norm(EBIZ_HEADERS["order_number"]) in found:
            break
    else:
        raise ImportFileError("eBiz header row not found (no 'Order Number' cell in rows 1-10).")
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
        parts = (row["order_number"] or "").split("|")
        row["so_1"] = _so(parts[0])
        row["so_2"] = _so(parts[1]) if len(parts) > 1 else None
        row["vendor_key"] = supplier_key(row["vendor"])
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
    """Delete every eBiz row, then insert the files' rows, in one transaction. Returns {file: rows inserted}.

    Raises ValueError, and changes nothing, when two files share a vendor key.
    """
    clash = ebiz_clashes(files)
    if clash:
        raise ValueError("Files overlap: " + "; ".join(f"{c['Vendor']} in {c['Files']}" for c in clash))
    now = datetime.now().isoformat(timespec="seconds")
    cols = EBIZ_COLUMNS + ["source_file", "imported_at"]
    with con:
        con.execute("DELETE FROM ebiz")
        for name, rows in files.items():
            con.executemany(f"INSERT INTO ebiz ({', '.join(cols)}) VALUES ({', '.join('?' * len(cols))})",
                            [[r[c] for c in EBIZ_COLUMNS] + [name, now] for r in rows])
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
        "SELECT DISTINCT vendor_key FROM ebiz WHERE vendor_key IS NOT NULL ORDER BY 1")]

```

- [ ] **Step 4: Run the tests to verify they pass**

Run: `python -m pytest check.py -q -k "ebiz"` → PASS (5 tests).
Run: `python -m pytest check.py -q` → all pass.

- [ ] **Step 5: Commit**

```powershell
git add core.py check.py
git commit -q -m @'
feat: eBiz table and import (each upload replaces all eBiz data)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
'@
```

---

### Task 2: Cross reference matching

**Files:**
- Modify: `core.py` (add below `ebiz_vendors`)
- Test: `check.py` (append)

**Interfaces:**
- Consumes: `core._so` (Task 1), views `v_std` / `v_cus`, table `ebiz`.
- Produces:
  - `core.so_order(so: str | None) -> tuple` — sort key: numeric SOs by value, then text, then None.
  - `core.cross_reference(con, supplier: str) -> dict` with:
    - `"orders"`: list of dicts `{"so", "tracker_qty", "tracker_rebate", "ebiz_qty", "ebiz_rebate", "via_so2", "status"}`, sorted by `so_order`. `status` ∈ `"Not in eBiz"`, `"Tracker > eBiz"`, `"Tracker < eBiz"`, `"Match"`.
    - `"rows"`: `{"STD": [...], "CUS": [...]}` — every tracker row of the supplier (all `v_std`/`v_cus` columns plus `"so"`: SO# or None), sorted by `(so_order(so), shipment_date or "", id)`.
    - `"ebiz_only"`: eBiz row dicts (all `ebiz` columns plus `"key"`) not matched to a tracker SO#, in `id` order.

- [ ] **Step 1: Write the failing test** — append to `check.py`:

```python
@pytest.fixture
def xref_con(tmp_path):
    """Supplier PACK: tracker SOs 1001-1004 and rows without an SO#; eBiz rows giving every status."""
    con = core.connect(tmp_path / "t.db")
    for so, qty, ship in [("1001", 100, "2026-08-03"), ("1002", 100, "2026-08-04"), ("1003", 50, "2026-08-05"),
                          ("1004", 10, "2026-08-06"), ("1004", 1, None), (None, 5, "2026-08-07"),
                          ("0", 2, "2026-08-08")]:
        insert(con, "STD", shipment_date=ship, packaging_supplier="Pack A Ltd.", supplier_key="PACK", ebiz_so=so,
               carton_qty=qty, unit_price=1.0, total_amount=float(qty))
    rows = core.read_ebiz(make_ebiz(tmp_path / "e.xlsx", [
        erow("1001", qty=100, rebate=8.8), erow("9999|1002", qty=80, rebate=4.0), erow("1003", qty=70, rebate=2.0),
        erow("8888|8889", qty=7, rebate=0.5), erow("1001", qty=3, vendor="Other Co")]))
    core.replace_ebiz(con, {"e.xlsx": rows})
    return con


def test_cross_reference(xref_con):
    x = core.cross_reference(xref_con, "PACK")
    got = [(o["so"], o["tracker_qty"], round(o["tracker_rebate"], 4), o["ebiz_qty"], o["ebiz_rebate"],
            o["status"], o["via_so2"]) for o in x["orders"]]
    assert got == [("1001", 100, 8.8, 100, 8.8, "Match", False),  # rebate 5 + PACD (100 - 5) * 4%
                   ("1002", 100, 8.8, 80, 4.0, "Tracker > eBiz", True),  # matched through SO_2
                   ("1003", 50, 4.4, 70, 2.0, "Tracker < eBiz", False),
                   ("1004", 11, 0.88, 0, 0.0, "Not in eBiz", False)]  # undated row counts, rebate 0
    assert [r["so"] for r in x["rows"]["STD"]] == ["1001", "1002", "1003", "1004", "1004", None, None]
    assert x["rows"]["CUS"] == []
    assert [(e["key"], e["qty_shipped"]) for e in x["ebiz_only"]] == [("8888", 7.0)]  # Other Co ignored
    assert core.cross_reference(xref_con, "NOBODY") == {"orders": [], "rows": {"STD": [], "CUS": []},
                                                        "ebiz_only": []}
    assert sorted(["12", None, "AB", "3"], key=core.so_order) == ["3", "12", "AB", None]
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest check.py -q -k "cross_reference"`
Expected: FAIL, `AttributeError: module 'core' has no attribute 'cross_reference'`.

- [ ] **Step 3: Implement** — in `core.py` below `ebiz_vendors`:

```python
def so_order(so):
    """Sort key for SO numbers: numeric ones by value, then text, then None."""
    if so is None:
        return (2, 0, "")
    return (0, int(so), "") if so.isdigit() else (1, 0, so)


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
        found = [dict(r) for r in con.execute(f"SELECT * FROM v_{kind.lower()} WHERE supplier_key IS ?",
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
    for e in (dict(r) for r in con.execute("SELECT * FROM ebiz WHERE vendor_key IS ? ORDER BY id", (supplier,))):
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
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest check.py -q -k "cross_reference"` → PASS. Then `python -m pytest check.py -q` → all pass.

- [ ] **Step 5: Commit**

```powershell
git add core.py check.py
git commit -q -m @'
feat: cross reference of tracker SO numbers against eBiz

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
'@
```

---

### Task 3: Discrepancy Report PDF

**Files:**
- Modify: `reports.py`
- Test: `check.py` (append)

**Interfaces:**
- Consumes: `core.cross_reference` output (Task 2), `core.EBIZ_HEADERS`, `core.EBIZ_REAL` (Task 1); existing `reports` helpers `_fmt`, `_total_line`, `_headers`, `safe_name`, `_NumberedCanvas`, constants `TRACKER_COLS`, `KINDS`, `BASE_WIDTHS`, `PADDING`, `CELL`, `HEAD`, `NAVY`, `BLUE`, `HEAD_BG`, `ALT_BG`, `TOTAL_BG`, `PAGE`, `MARGIN`, `ASSETS`.
- Produces:
  - `reports._title_bar(title, avail) -> Table`, `reports._pdf(story, title) -> bytes` (landscape A4 with logos, generated time, Page x of y). `full_tracker_pdf` now uses both; its output is unchanged.
  - `reports.discrepancy_pdf(con, supplier, xref) -> (str, bytes)`.

- [ ] **Step 1: Write the failing tests** — append to `check.py`:

```python
DISCREPANCY_TITLES = ("Tracker Orders Not Found in eBiz", "Quantity Mismatch (Tracker > eBiz)",
                      "Quantity Mismatch (Tracker < eBiz)", "Tracker Rows With No SO#",
                      "eBiz Orders Not Found in Tracker")


def test_discrepancy_pdf(xref_con):
    name, data = reports.discrepancy_pdf(xref_con, "PACK", core.cross_reference(xref_con, "PACK"))
    assert name == f"Discrepancy Report_Pack A Ltd_{date.today():%d-%b-%Y}.pdf"
    text = pdf_text(data)
    for expected in ("Discrepancy Report (Pack A Ltd.)", *DISCREPANCY_TITLES, "Standard Carton", "eBiz QTY",
                     "8888|8889", "Page 1 of", "Generated:"):
        assert expected in text, expected
    assert "No discrepancies found." not in text
    assert len(pypdf.PdfReader(BytesIO(data)).pages[0].images) == 2  # both logos


def test_discrepancy_pdf_says_when_a_section_is_empty(tmp_path):
    con = core.connect(tmp_path / "t.db")
    insert(con, "STD", shipment_date="2026-08-03", packaging_supplier="Pack A Ltd.", supplier_key="PACK",
           ebiz_so="1001", carton_qty=100, unit_price=1.0, total_amount=100.0)
    core.replace_ebiz(con, {"e.xlsx": core.read_ebiz(make_ebiz(tmp_path / "e.xlsx", [erow("1001", qty=100)]))})
    _, data = reports.discrepancy_pdf(con, "PACK", core.cross_reference(con, "PACK"))
    text = pdf_text(data)
    assert all(t in text for t in DISCREPANCY_TITLES)
    assert text.count("No discrepancies found.") == 5
```

- [ ] **Step 2: Run to verify they fail**

Run: `python -m pytest check.py -q -k "discrepancy"`
Expected: FAIL, `AttributeError: module 'reports' has no attribute 'discrepancy_pdf'`.

- [ ] **Step 3: Extract the shared PDF helpers.** In `reports.py`, directly above `def full_tracker_pdf(`, add:

```python
def _title_bar(title, avail):
    return Table([[title]], colWidths=[avail], style=TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NAVY), ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
        ("FONT", (0, 0), (-1, -1), "Helvetica-Bold", 14), ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]))


def _pdf(story, title):
    """Landscape A4 PDF bytes with both logos, the generated time and 'Page x of y' on every page."""
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
    return buf.getvalue()
```

Then replace the whole `full_tracker_pdf` function with:

```python
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
    name = core.supplier_name(con, supplier)
    title = f"Full Tracker Report ({name})"
    story = [_title_bar(title, avail), Spacer(1, 6)]
    for label, rows in blocks:
        if rows:
            story += [_block(label, rows, cols, widths), Spacer(1, 8)]
    qty_i = cols.index("carton_qty") + 1
    story.append(Table([_total_line("Tracker Grand Total", all_rows, cols)], colWidths=widths,
                       style=TableStyle(PADDING + [
                           ("SPAN", (0, 0), (qty_i - 1, 0)), ("BACKGROUND", (0, 0), (-1, -1), NAVY),
                           ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
                           ("FONT", (0, 0), (-1, -1), "Helvetica-Bold", 6), ("ALIGN", (0, 0), (-1, -1), "RIGHT")])))
    qs = sorted(quarters)
    span = f"Q{qs[0]}" if len(qs) == 1 else f"Q{qs[0]}-Q{qs[-1]}"
    return f"Full Tracker Sheet for {brand}_{safe_name(name)}_{span}_{fiscal_year}.pdf", _pdf(story, title)
```

Run `python -m pytest check.py -q -k "full_tracker"` → still PASS (refactor only).

- [ ] **Step 4: Add the discrepancy report.** Directly below `full_tracker_pdf` add:

```python
YELLOW = colors.HexColor("#FFFF99")
NOTE_STYLE = ParagraphStyle("note", fontName="Helvetica-Oblique", fontSize=8, textColor=colors.HexColor("#595959"))
DISCREPANCY_SECTIONS = [  # (order status, or None for rows without an SO#; title; bar colour; eBiz QTY column)
    ("Not in eBiz", "Tracker Orders Not Found in eBiz", colors.HexColor("#C00000"), False),
    ("Tracker > eBiz", "Quantity Mismatch (Tracker > eBiz)", colors.HexColor("#C55A11"), True),
    ("Tracker < eBiz", "Quantity Mismatch (Tracker < eBiz)", colors.HexColor("#C55A11"), True),
    (None, "Tracker Rows With No SO#", colors.HexColor("#505050"), False),
]
EBIZ_PDF_COLS = ["vendor", "vendor_code", "product", "order_number", "qty_shipped", "net_price_1000",
                 "gross_price_1000", "total_rebate", "currency"]
EBIZ_PDF_WIDTHS = [16, 90, 36, 60, 60, 30, 34, 34, 34, 24]  # SL, then EBIZ_PDF_COLS


def _bar(text, background, avail):
    return Table([[text]], colWidths=[avail], style=TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), background), ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
        ("FONT", (0, 0), (-1, -1), "Helvetica-Bold", 10), ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))


def _grid(header, lines, widths, total, right_from, highlight=(), banner=None):
    """Bordered table: optional blue banner, header, lines (alternate shading, `highlight` line indexes pale
    yellow) and a total line whose label spans the columns before `right_from`. Columns from `right_from`
    align right."""
    data = []
    style = PADDING + [("FONT", (0, 0), (-1, -1), "Helvetica", 5.5), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]
    if banner:
        data.append([banner] + [""] * (len(widths) - 1))
        style += [("SPAN", (0, 0), (-1, 0)), ("BACKGROUND", (0, 0), (-1, 0), BLUE),
                  ("TEXTCOLOR", (0, 0), (-1, 0), colors.white), ("FONT", (0, 0), (-1, 0), "Helvetica-Bold", 7)]
    head = len(data)
    data.append([Paragraph(escape(h), HEAD) for h in header])
    style += [("BACKGROUND", (0, head), (-1, head), HEAD_BG), ("GRID", (0, head), (-1, -1), 0.25, HEAD_BG),
              ("ALIGN", (right_from, head + 1), (-1, -1), "RIGHT")]
    for i, line in enumerate(lines):
        data.append(line)
        r = len(data) - 1
        if i in highlight:
            style.append(("BACKGROUND", (0, r), (-1, r), YELLOW))
        elif i % 2:
            style.append(("BACKGROUND", (0, r), (-1, r), ALT_BG))
    r = len(data)
    data.append(total)
    style += [("SPAN", (0, r), (right_from - 1, r)), ("BACKGROUND", (0, r), (-1, r), TOTAL_BG),
              ("FONT", (0, r), (-1, r), "Helvetica-Bold", 5.5), ("ALIGN", (0, r), (-1, r), "RIGHT")]
    return Table(data, colWidths=widths, repeatRows=head + 1, style=TableStyle(style))


def _ebiz_fmt(col, value):
    """PDF cell content for an eBiz column."""
    if value is None:
        return ""
    if col == "qty_shipped":
        return f"{value:,.0f}"
    if col in core.EBIZ_REAL:
        return f"{value:,.2f}"
    return Paragraph(escape(str(value)), CELL)


def _no_discrepancies(avail):
    return Table([[Paragraph("No discrepancies found.", NOTE_STYLE)]], colWidths=[avail])


def discrepancy_pdf(con, supplier, xref):
    """Tracker rows that disagree with eBiz, then eBiz rows missing from the tracker, for one supplier.

    `xref` is core.cross_reference(con, supplier). Rows matched through SO_2 are pale yellow.
    """
    name = core.supplier_name(con, supplier)
    title = f"Discrepancy Report ({name})"
    avail = PAGE[0] - 2 * MARGIN
    orders = {o["so"]: o for o in xref["orders"]}
    qty_i = TRACKER_COLS.index("carton_qty") + 1
    story = [_title_bar(title, avail), Spacer(1, 6)]
    for status, heading, color, with_ebiz in DISCREPANCY_SECTIONS:
        story += [_bar(heading, color, avail), Spacer(1, 3)]
        extra = 1 if with_ebiz else 0
        base = [BASE_WIDTHS["sr"]] + [BASE_WIDTHS[c] for c in TRACKER_COLS] + [36] * extra  # 'eBiz QTY' fits
        widths = [w * avail / sum(base) for w in base]
        found = False
        for kind, label in KINDS:
            rows = [r for r in xref["rows"][kind] if (orders[r["so"]]["status"] if r["so"] else None) == status]
            if not rows:
                continue
            found = True
            header = ["SL"] + _headers(rows, TRACKER_COLS) + ["eBiz QTY"] * extra
            lines = [[str(i)] + [_fmt(c, r[c]) for c in TRACKER_COLS]
                     + [f"{orders[r['so']]['ebiz_qty']:,.0f}"] * extra for i, r in enumerate(rows, 1)]
            highlight = {i for i, r in enumerate(rows) if r["so"] and orders[r["so"]]["via_so2"]}
            total = _total_line("Total", rows, TRACKER_COLS) + [""] * extra
            story += [_grid(header, lines, widths, total, qty_i, highlight, f"{label} Carton"), Spacer(1, 6)]
        if not found:
            story += [_no_discrepancies(avail), Spacer(1, 6)]
    story += [_bar("eBiz Orders Not Found in Tracker", NAVY, avail), Spacer(1, 3)]
    ebiz = xref["ebiz_only"]
    if ebiz:
        header = ["SL"] + [core.EBIZ_HEADERS[c] for c in EBIZ_PDF_COLS]
        lines = [[str(i)] + [_ebiz_fmt(c, e[c]) for c in EBIZ_PDF_COLS] for i, e in enumerate(ebiz, 1)]
        highlight = {i for i, e in enumerate(ebiz) if not e["so_1"] and e["so_2"]}
        total = ["Total", "", "", "", "", _ebiz_fmt("qty_shipped", sum(e["qty_shipped"] or 0 for e in ebiz)),
                 "", "", _ebiz_fmt("total_rebate", sum(e["total_rebate"] or 0 for e in ebiz)), ""]
        widths = [w * avail / sum(EBIZ_PDF_WIDTHS) for w in EBIZ_PDF_WIDTHS]
        story.append(_grid(header, lines, widths, total, 5, highlight))
    else:
        story.append(_no_discrepancies(avail))
    return f"Discrepancy Report_{safe_name(name)}_{date.today():%d-%b-%Y}.pdf", _pdf(story, title)
```

- [ ] **Step 5: Run tests**

Run: `python -m pytest check.py -q -k "discrepancy or full_tracker"` → PASS. Then `python -m pytest check.py -q` → all pass.

- [ ] **Step 6: Commit**

```powershell
git add reports.py check.py
git commit -q -m @'
feat: discrepancy report PDF (tracker vs eBiz)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
'@
```

---

### Task 4: Cross Reference workbook

**Files:**
- Modify: `reports.py`, `docs/superpowers/specs/2026-10-08-ebiz-discrepancy-design.md`
- Test: `check.py` (append)

**Interfaces:**
- Consumes: `core.cross_reference` output, `core.so_order` (Task 2); existing Summary helpers in `reports.py`: `_Rows`, `_page_points`, `_cell`, `_span`, `_logos`, `_add`, `_widths`, `_bytes`, `_font`, constants `BODY`, `BODY_BOLD`, `WHITE_BOLD`, `HEADING`, `NAVY_FILL`, `GRAND_FILL`, `ROW_LINE`, `GRAND_LINE`, `HEADING_LINE`, `LEFT`, `CENTER`, `RIGHT`, `QTY_FORMAT`, `MONEY_FORMAT`, `NAVY_HEX`, `GREY_HEX`, `SUMMARY_MARGINS`; test helper `summary_sheet(data)` already in `check.py`.
- Produces:
  - `reports._title_block(ws, rows, ncols, widths, title, subtitle, now)` and `reports._print_setup(ws, ncols, last_row, footer_left, now)`; `summary_xlsx` now uses both (output unchanged).
  - `reports.cross_reference_xlsx(con, supplier, xref) -> (str, bytes)`.

- [ ] **Step 1: Write the failing test** — append to `check.py`:

```python
XREF_HEADER = ["SO#", "Tracker QTY", "Tracker Rebate", "eBiz QTY", "eBiz Rebate", "QTY Difference", "Status"]


def test_cross_reference_xlsx(xref_con):
    name, data = reports.cross_reference_xlsx(xref_con, "PACK", core.cross_reference(xref_con, "PACK"))
    assert name == f"Cross Reference_Pack A Ltd_{date.today():%d-%b-%Y}.xlsx"
    ws, rows = summary_sheet(data)
    assert rows[0][0] == "Cross Reference"
    assert len(ws._images) == 2  # PacD and Matalan logos
    head = rows.index(XREF_HEADER)
    assert rows[head + 1:head + 7] == [
        ["1001", 100, 8.8, 100, 8.8, 0, "Match"],
        ["1002", 100, 8.8, 80, 4.0, 20, "Tracker > eBiz"],
        ["1003", 50, 4.4, 70, 2.0, -20, "Tracker < eBiz"],
        ["1004", 11, 0.88, 0, 0, 11, "Not in eBiz"],
        ["(No SO#)", 7, 0.62, None, None, None, None],
        ["Total", 268, 23.5, 250, 14.8, 18, None]]
    only = rows.index(["Order", "eBiz QTY", "eBiz Rebate", None, None, None, None])
    assert rows[only + 1:only + 3] == [["8888", 7, 0.5, None, None, None, None],
                                       ["Total", 7, 0.5, None, None, None, None]]
    assert ws.page_setup.orientation == "portrait" and ws.print_title_rows
    assert ws.oddFooter.center.text == "Page &P of &N"
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest check.py -q -k "cross_reference_xlsx"`
Expected: FAIL, `AttributeError: module 'reports' has no attribute 'cross_reference_xlsx'`.

- [ ] **Step 3: Extract the Summary title block and print setup.** In `reports.py`, directly above `def summary_xlsx(`, add:

```python
def _title_block(ws, rows, ncols, widths, title, subtitle, now):
    """Rows 1-3: title, subtitle and generated time, centred between the PacD and Matalan logos."""
    _span(ws, rows.add(30), 1, ncols, title, font=_font(18, bold=True, color=NAVY_HEX), align=CENTER)
    _span(ws, rows.add(18), 1, ncols, subtitle, font=_font(11, color=GREY_HEX), align=CENTER)
    _span(ws, rows.add(16), 1, ncols, f"Generated {now:%d-%b-%Y %H:%M}",
          font=_font(9, italic=True, color=GREY_HEX), align=CENTER)
    _logos(ws, widths, top_px=20, height_px=44)


def _print_setup(ws, ncols, last_row, footer_left, now):
    """A4 portrait, one page wide, centred; footer with `footer_left`, 'Page x of y' and the generated date."""
    ws.print_area = f"A1:{get_column_letter(ncols)}{last_row}"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins = PageMargins(**SUMMARY_MARGINS)
    for part, text in ((ws.oddFooter.left, footer_left), (ws.oddFooter.center, "Page &P of &N"),
                       (ws.oddFooter.right, f"Generated {now:%d-%b-%Y}")):
        part.text, part.size, part.color = text, 8, GREY_HEX
```

In `summary_xlsx`, replace these lines:

```python
    _span(ws, rows.add(30), 1, ncols, SUMMARY_TITLE, font=_font(18, bold=True, color=NAVY_HEX), align=CENTER)
    _span(ws, rows.add(18), 1, ncols, f"Fiscal year {fiscal_year}  ·  Shipments {period}",
          font=_font(11, color=GREY_HEX), align=CENTER)
    _span(ws, rows.add(16), 1, ncols, f"Generated {now:%d-%b-%Y %H:%M}",
          font=_font(9, italic=True, color=GREY_HEX), align=CENTER)
    _logos(ws, SUMMARY_WIDTHS, top_px=20, height_px=44)
```

with:

```python
    _title_block(ws, rows, ncols, SUMMARY_WIDTHS, SUMMARY_TITLE, f"Fiscal year {fiscal_year}  ·  Shipments {period}",
                 now)
```

and replace the block from `ws.print_area = ...` through the footer `for` loop (the lines just before `return f"{SUMMARY_PREFIX}...`) with:

```python
    _print_setup(ws, ncols, rows.next - 1, f"{SUMMARY_TITLE}  ·  FY{fiscal_year}", now)
```

Run `python -m pytest check.py -q -k "summary"` → still PASS (refactor only).

- [ ] **Step 4: Add the Cross Reference workbook.** Directly below `summary_xlsx` add:

```python
XREF_HEADERS = ["SO#", "Tracker QTY", "Tracker Rebate", "eBiz QTY", "eBiz Rebate", "QTY Difference", "Status"]
XREF_WIDTHS = [16, 13, 15, 13, 15, 15, 16]
XREF_FORMATS = [None, QTY_FORMAT, MONEY_FORMAT, QTY_FORMAT, MONEY_FORMAT, QTY_FORMAT, None]
XREF_ALIGNS = [LEFT, RIGHT, RIGHT, RIGHT, RIGHT, RIGHT, CENTER]
STATUS_COLORS = {"Not in eBiz": "C00000", "Tracker > eBiz": "C55A11", "Tracker < eBiz": "C55A11", "Match": "548235"}


def _xref_line(ws, row, values, font=BODY, fill=None, border=ROW_LINE):
    """Values from column A, formatted as the Cross Reference columns."""
    for col, (value, fmt, align) in enumerate(zip(values, XREF_FORMATS, XREF_ALIGNS), 1):
        _cell(ws, row, col, value, font=font, fmt=fmt, align=align, fill=fill, border=border)


def _xref_header(ws, row, headers):
    for col, (header, align) in enumerate(zip(headers, XREF_ALIGNS), 1):
        _cell(ws, row, col, header, font=WHITE_BOLD, fill=NAVY_FILL, align=align)


def cross_reference_xlsx(con, supplier, xref):
    """Tracker vs eBiz quantity and rebate per SO#, then eBiz orders missing from the tracker, for one supplier.

    `xref` is core.cross_reference(con, supplier). Laid out like the Summary: A4 portrait, one page wide.
    """
    name = core.supplier_name(con, supplier)
    now = datetime.now()
    wb = Workbook()
    ws = wb.active
    ws.title = "Cross Reference"
    ws.sheet_view.showGridLines = False
    _widths(ws, XREF_WIDTHS)
    ncols = len(XREF_HEADERS)
    rows = _Rows(ws, _page_points(XREF_WIDTHS))
    _title_block(ws, rows, ncols, XREF_WIDTHS, "Cross Reference", f"{name}  ·  all tracker rows vs eBiz data", now)

    rows.add(12)
    _span(ws, rows.add(22), 1, ncols, "Tracker orders", font=HEADING, align=LEFT, border=HEADING_LINE)
    header_row = rows.add(30)
    _xref_header(ws, header_row, XREF_HEADERS)
    ws.print_title_rows = f"{header_row}:{header_row}"
    total = [0, 0, 0, 0]
    for o in xref["orders"]:
        values = [o["tracker_qty"], round(o["tracker_rebate"], 2), o["ebiz_qty"], round(o["ebiz_rebate"], 2)]
        row = rows.add(16)
        _xref_line(ws, row, [o["so"]] + values + [values[0] - values[2], o["status"]])
        ws.cell(row, 7).font = _font(bold=True, color=STATUS_COLORS[o["status"]])
        total = _add(total, values)
    no_so = [r for kind_rows in xref["rows"].values() for r in kind_rows if not r["so"]]
    if no_so:
        qty = sum(r["carton_qty"] or 0 for r in no_so)
        rebate = round(sum((r["pacd"] or 0) + (r["matalan_rebate"] or 0) for r in no_so), 2)
        _xref_line(ws, rows.add(16), ["(No SO#)", qty, rebate, None, None, None, None])
        total = _add(total, [qty, rebate, 0, 0])
    _xref_line(ws, rows.add(18), ["Total"] + total + [round(total[0] - total[2], 2), None],
               BODY_BOLD, GRAND_FILL, GRAND_LINE)

    groups = {}
    for e in xref["ebiz_only"]:
        group = groups.setdefault(e["key"] or "(no order number)", [0, 0])
        group[0] += e["qty_shipped"] or 0
        group[1] += e["total_rebate"] or 0
    rows.keep_together(12 + 22 + 30 + 16 * min(len(groups), 5) + 18)
    rows.add(12)
    _span(ws, rows.add(22), 1, ncols, "eBiz orders not in tracker", font=HEADING, align=LEFT, border=HEADING_LINE)
    _xref_header(ws, rows.add(30), ["Order", "eBiz QTY", "eBiz Rebate"])
    if groups:
        for order in sorted(groups, key=core.so_order):
            qty, rebate = groups[order]
            _xref_line(ws, rows.add(16), [order, qty, round(rebate, 2)])
        _xref_line(ws, rows.add(18), ["Total", sum(g[0] for g in groups.values()),
                                      round(sum(g[1] for g in groups.values()), 2)], BODY_BOLD, GRAND_FILL, GRAND_LINE)
    else:
        _span(ws, rows.add(16), 1, 3, "None", font=_font(italic=True, color=GREY_HEX), align=LEFT)

    _print_setup(ws, ncols, rows.next - 1, f"Cross Reference  ·  {name}", now)
    return f"Cross Reference_{safe_name(name)}_{now:%d-%b-%Y}.xlsx", _bytes(wb)
```

- [ ] **Step 5: Update the spec.** In `docs/superpowers/specs/2026-10-08-ebiz-discrepancy-design.md`, section "Cross Reference (xlsx)", change `A4 landscape fit to one page wide` to `A4 portrait fit to one page wide, header row repeated on each printed page` (seven narrow columns fit portrait).

- [ ] **Step 6: Run tests**

Run: `python -m pytest check.py -q -k "cross_reference or summary"` → PASS. Then `python -m pytest check.py -q` → all pass.

- [ ] **Step 7: Commit**

```powershell
git add reports.py check.py docs/superpowers/specs/2026-10-08-ebiz-discrepancy-design.md
git commit -q -m @'
feat: cross reference workbook (tracker vs eBiz per SO number)

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
'@
```

---

### Task 5: Discrepancy option in build_all / save_all

**Files:**
- Modify: `reports.py` (`build_all`, `save_all`)
- Test: `check.py` (append)

**Interfaces:**
- Consumes: `core.ebiz_vendors`, `core.cross_reference`, `reports.discrepancy_pdf`, `reports.cross_reference_xlsx`; test constants `SUMMARY`, `TW_FILES`, `UG_FILES`, `TW`, `TW_KEY`, `UG_KEY`, fixture `demo_con`.
- Produces: `reports.build_all(con, suppliers, fiscal_year, quarters, discrepancy=False)` and `reports.save_all(con, suppliers, fiscal_year, quarters, reports_dir, discrepancy=False)`. With `discrepancy=True`, each picked supplier that has eBiz rows gets `<folder>/Discrepancy Report_<folder>_<date>.pdf` and `<folder>/Cross Reference_<folder>_<date>.xlsx` after its tracker files, even when it has no rows in the selected quarters.

- [ ] **Step 1: Write the failing test** — append to `check.py`:

```python
def test_build_all_adds_discrepancy_files_for_suppliers_with_ebiz(demo_con, tmp_path):
    so = demo_con.execute("SELECT ebiz_so FROM track_std WHERE supplier_key = ? AND ebiz_so IS NOT NULL",
                          (TW_KEY,)).fetchone()[0]
    core.replace_ebiz(demo_con, {"e.xlsx": core.read_ebiz(make_ebiz(tmp_path / "e.xlsx", [erow(so, vendor=TW)]))})
    folder = "Transworld Packaging (BD) Ltd"
    today = f"{date.today():%d-%b-%Y}"
    disc = [f"{folder}/Discrepancy Report_{folder}_{today}.pdf", f"{folder}/Cross Reference_{folder}_{today}.xlsx"]
    items = reports.build_all(demo_con, [TW_KEY, UG_KEY], 2026, [3], discrepancy=True)
    assert [p for p, _ in items] == [SUMMARY] + TW_FILES + disc + UG_FILES  # Uniglory has no eBiz rows
    assert [p for p, _ in reports.build_all(demo_con, [TW_KEY, UG_KEY], 2026, [3])] == [SUMMARY] + TW_FILES + UG_FILES
    # all tracker rows are compared, so a supplier with nothing in the chosen quarters still gets them
    assert [p for p, _ in reports.build_all(demo_con, [TW_KEY], 2026, [1], discrepancy=True)] == [SUMMARY] + disc
    root = tmp_path / "Reports"
    paths = reports.save_all(demo_con, [TW_KEY], 2026, [3], root, discrepancy=True)
    assert sorted(p.relative_to(root).as_posix() for p in paths) == sorted(
        f"August/{p}" for p in [SUMMARY] + TW_FILES + disc)
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest check.py -q -k "adds_discrepancy"`
Expected: FAIL, `TypeError: build_all() got an unexpected keyword argument 'discrepancy'`.

- [ ] **Step 3: Implement.** Replace `build_all` and `save_all` in `reports.py` with:

```python
def build_all(con, suppliers, fiscal_year, quarters, discrepancy=False):
    """Every report for the chosen suppliers as (relative path, bytes).

    The Summary comes first, then each supplier's files under a folder named after the supplier. Tracker
    reports need rows in the selection. With `discrepancy`, suppliers that have eBiz rows also get the
    Discrepancy Report and the Cross Reference, which compare all their tracker rows (selection ignored).
    """
    items = [summary_xlsx(con, fiscal_year)]
    with_ebiz = set(core.ebiz_vendors(con)) if discrepancy else set()
    for key in suppliers:
        files = []
        if core.date_range(con, supplier=key, fiscal_year=fiscal_year, quarters=quarters):
            files += [full_tracker_pdf(con, key, fiscal_year, quarters, brand) for brand in ("PacD", "Matalan")]
            files += [raw_data_xlsx(con, key, fiscal_year, q) for q in sorted(quarters)]
            files.append(shipment_tracking_xlsx(con, key))
        if key in with_ebiz:
            xref = core.cross_reference(con, key)
            files += [discrepancy_pdf(con, key, xref), cross_reference_xlsx(con, key, xref)]
        folder = safe_name(core.supplier_name(con, key))
        items += [(f"{folder}/{name}", data) for name, data in filter(None, files)]
    return [item for item in items if item]
```

```python
def save_all(con, suppliers, fiscal_year, quarters, reports_dir, discrepancy=False):
    """Write build_all() under <reports_dir>/<Month>/. Existing files are overwritten."""
    month = month_name(con, suppliers, fiscal_year, quarters)
    if not month:
        return []
    paths = []
    for rel, data in build_all(con, suppliers, fiscal_year, quarters, discrepancy):
        path = Path(reports_dir) / month / rel
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(data)
        paths.append(path)
    return paths
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest check.py -q -k "build or save_all"` → PASS. Then `python -m pytest check.py -q` → all pass.

- [ ] **Step 5: Commit**

```powershell
git add reports.py check.py
git commit -q -m @'
feat: build discrepancy and cross reference files with the other reports

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
'@
```

---

### Task 6: App — eBiz section and Reports checkbox

**Files:**
- Modify: `app.py`
- Test: `check.py` (append one test; edit the uploader stub in `test_app_date_check_imports_and_remembers_suggested_dates`)

**Interfaces:**
- Consumes: `core.read_ebiz`, `core.ebiz_preview`, `core.ebiz_clashes`, `core.replace_ebiz`, `core.ebiz_status`, `core.suppliers`; `reports.build_all(..., discrepancy=)`, `reports.save_all(..., discrepancy=)`; test helpers `FakeUpload`, `make_ebiz`, `erow`, `DEMO`, `TW`.
- Produces: Import tab uploader with `key="ebiz_upload"`, button "Replace all eBiz data"; Reports tab checkbox "Discrepancy Report + Cross Reference" (key `disc-True` / `disc-False`).

- [ ] **Step 1: Write the failing test and fix the existing stub.**

In `test_app_date_check_imports_and_remembers_suggested_dates`, change

```python
    monkeypatch.setattr(st, "file_uploader", lambda *a, **k: [FakeUpload(path)])
```

to

```python
    monkeypatch.setattr(st, "file_uploader", lambda *a, **k: [] if k.get("key") == "ebiz_upload" else [FakeUpload(path)])
```

Append to `check.py`:

```python
def test_app_replaces_ebiz_data(tmp_path, monkeypatch):
    import streamlit as st
    from streamlit.testing.v1 import AppTest
    db = tmp_path / "app.db"
    monkeypatch.setenv("CARTON_DB", str(db))
    core.save(core.connect(db), core.read_workbook(DEMO), DEMO.name)
    path = make_ebiz(tmp_path / "ebiz.xlsx", [erow("1", vendor=TW), erow("2", vendor="Nobody Ltd")])
    monkeypatch.setattr(st, "file_uploader",
                        lambda *a, **k: [FakeUpload(path)] if k.get("key") == "ebiz_upload" else [])
    at = AppTest.from_file(str(HERE / "app.py"), default_timeout=30).run()
    assert not at.exception
    assert any("NOBODY" in w.value for w in at.warning)  # vendor without tracker rows
    checkbox = next(c for c in at.checkbox if c.label == "Discrepancy Report + Cross Reference")
    assert checkbox.disabled and not checkbox.value  # no eBiz data yet
    next(b for b in at.button if b.label == "Replace all eBiz data").click().run()
    assert not at.exception
    assert core.ebiz_status(core.connect(db))[0] == 2
    checkbox = next(c for c in at.checkbox if c.label == "Discrepancy Report + Cross Reference")
    assert checkbox.value and not checkbox.disabled
```

- [ ] **Step 2: Run to verify it fails**

Run: `python -m pytest check.py -q -k "app_replaces_ebiz"`
Expected: FAIL (no warning / `StopIteration` for the checkbox or button).

- [ ] **Step 3: Implement in `app.py`.**

3a. Below the existing `read_upload` function add:

```python
@st.cache_data(max_entries=20, show_spinner="Reading eBiz file...")
def read_ebiz_upload(content):
    """read_ebiz() once per file content."""
    return core.read_ebiz(BytesIO(content))
```

3b. At the end of the `with tab_import:` block (after the "Saved date corrections" expander, still indented under `with tab_import:`), add:

```python
    st.divider()
    st.subheader("eBiz data")
    status = core.ebiz_status(con)
    st.caption(f"Current eBiz data: {status[0]} rows from {len(status[1])} file(s), loaded "
               f"{status[2].replace('T', ' ')[:16]}." if status else "No eBiz data yet.")
    ebiz_uploads = st.file_uploader("eBiz exports (one file per vendor; upload them all together)",
                                    type=["xlsx", "xlsm"], accept_multiple_files=True, key="ebiz_upload")
    ebiz_files = {}
    for upload in ebiz_uploads or []:
        try:
            ebiz_files[upload.name] = read_ebiz_upload(upload.getvalue())
        except core.ImportFileError as e:
            st.error(f"{upload.name}: {e} This file will not be imported.")
    if ebiz_files:
        st.dataframe(pd.DataFrame(core.ebiz_preview(ebiz_files)), hide_index=True)
        unknown = sorted({r["vendor_key"] for rows in ebiz_files.values() for r in rows}
                         - set(core.suppliers(con)), key=str)
        if unknown:
            st.warning(f"No tracker rows for {', '.join(map(str, unknown))}. "
                       "They can't be reported until their tracker is imported.")
        ebiz_clash = core.ebiz_clashes(ebiz_files)
        if ebiz_clash:
            st.error("These files contain the same vendor. Remove one of them before importing.")
            st.dataframe(pd.DataFrame(ebiz_clash), hide_index=True)
        st.caption(f"Replacing deletes all {status[0] if status else 0} current eBiz rows first.")
        if st.button("Replace all eBiz data", type="primary", disabled=bool(ebiz_clash)):
            counts = core.replace_ebiz(con, ebiz_files)
            st.session_state.pop("built", None)
            st.success(f"eBiz data replaced: {sum(counts.values())} rows from {len(counts)} file(s).")
```

3c. In the Reports tab, replace

```python
                key = (tuple(picked), year, tuple(chosen))
```

with

```python
                has_ebiz = core.ebiz_status(con) is not None
                with_disc = st.checkbox("Discrepancy Report + Cross Reference", value=has_ebiz,
                                        disabled=not has_ebiz, key=f"disc-{has_ebiz}",
                                        help="Compares all tracker rows of each supplier with the eBiz data; "
                                             "fiscal year and quarters don't apply.")
                key = (tuple(picked), year, tuple(chosen), with_disc)
```

and change the two report calls to pass the option:

```python
                        items = reports.build_all(con, picked, year, chosen, discrepancy=with_disc)
```

```python
                            paths = reports.save_all(con, picked, year, chosen, REPORTS_DIR, discrepancy=with_disc)
```

- [ ] **Step 4: Run tests**

Run: `python -m pytest check.py -q -k "app"` → PASS. Then `python -m pytest check.py -q` → all pass.

- [ ] **Step 5: Browser check (controller).** Start the `carton-app-scratch-db` preview from `.claude/launch.json` (port 8502, database `%TEMP%\carton_ui_test.db`, never `carton.db`). Upload a small tracker made with `check.make_xlsx` and an eBiz file made with `check.make_ebiz` through the file inputs (inject with JavaScript `DataTransfer`, as done before). Confirm: the eBiz preview, the warning for a vendor without tracker rows, "Replace all eBiz data", the Reports checkbox turning on, and "Build reports" listing the Discrepancy Report and Cross Reference. Take a screenshot, then stop the preview and delete the scratch database.

- [ ] **Step 6: Commit**

```powershell
git add app.py check.py
git commit -q -m @'
feat: eBiz upload on Import tab and discrepancy option on Reports tab

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
'@
```

---

### Task 7: Real-data check against the old generator

**Files:** none changed (scratchpad script only).

**Interfaces:**
- Consumes: `core.read_workbook`, `core.save`, `core.read_ebiz`, `core.replace_ebiz`, `core.cross_reference`.

- [ ] **Step 1: Write the check script** at `C:\Users\Shoaib\AppData\Local\Temp\claude\C--Users-Shoaib-OneDrive---PacD-Projects-Matalan-Matalan-Carton-Supplu-Report\92a98c42-9236-410d-a184-da4bd348270f\scratchpad\xref_real.py` (the session scratchpad; never inside `Reports/` or the project):

```python
"""Compare core.cross_reference with the old generator's Cross Reference sheet (Uniglory, 15 Sep)."""
import shutil, sys, tempfile
from pathlib import Path
import openpyxl
root = Path(r"C:\Users\Shoaib\OneDrive - PacD\Projects\Matalan\Matalan Carton Supplu Report")
sys.path.insert(0, str(root))
import core
sys.stdout.reconfigure(encoding="utf-8")
tmp = Path(tempfile.mkdtemp())
shutil.copy(root / "eBiz Report Generator.xlsm", tmp / "gen.xlsm")  # never open the original
gen = openpyxl.load_workbook(tmp / "gen.xlsm", data_only=True)
out = openpyxl.Workbook()
for row in gen["eBiz data"].iter_rows(values_only=True):
    out.active.append(row)
out.save(tmp / "ebiz_uniglory.xlsx")  # read_ebiz reads the first sheet
con = core.connect(tmp / "t.db")
tracker = root / "Reports" / "August" / "Uniglory Packaging Industries Ltd" / \
    "Matalan Carton Shipment Tracking_NPS_BD Dec-25 to Aug-26 (Uniglory).xlsx"
core.save(con, core.read_workbook(tracker), tracker.name)
core.replace_ebiz(con, {"ebiz_uniglory.xlsx": core.read_ebiz(tmp / "ebiz_uniglory.xlsx")})
x = core.cross_reference(con, "UNIGLORY")
old, old_only = {}, {}
for a, b, _, d, _, _, g, h, _ in gen["Cross Reference"].iter_rows(min_row=2, max_col=9, values_only=True):
    if a not in (None, "") and a != "(No Supplier Order)":
        old[str(a)] = (b or 0, d or 0)
    if g not in (None, ""):
        old_only[str(g)] = h or 0
new = {o["so"]: (o["tracker_qty"], o["ebiz_qty"]) for o in x["orders"]}
diff = {k: (old.get(k), new.get(k)) for k in set(old) | set(new) if old.get(k) != new.get(k)}
print("tracker SO#s old/new:", len(old), len(new), "| differences:", len(diff))
for k in sorted(diff, key=core.so_order)[:20]:
    print("  ", k, "old", diff[k][0], "new", diff[k][1])
new_only = {}
for e in x["ebiz_only"]:
    new_only[e["key"]] = new_only.get(e["key"], 0) + (e["qty_shipped"] or 0)
print("eBiz-only equal:", old_only == new_only, len(old_only), len(new_only))
print("statuses:", {s: sum(o["status"] == s for o in x["orders"])
                    for s in ("Not in eBiz", "Tracker > eBiz", "Tracker < eBiz", "Match")})
```

- [ ] **Step 2: Run it**

Run: `python "C:\Users\Shoaib\AppData\Local\Temp\claude\C--Users-Shoaib-OneDrive---PacD-Projects-Matalan-Matalan-Carton-Supplu-Report\92a98c42-9236-410d-a184-da4bd348270f\scratchpad\xref_real.py"`
Expected: `differences: 0` and `eBiz-only equal: True`. If there are differences, do not change matching code to force a match: print the tracker rows of the first differing SO# from both sources, find the cause (for example the tracker changed after 15 Sep) and report it to the user.

- [ ] **Step 3: Render one real Discrepancy Report and Cross Reference** into the scratchpad with `reports.discrepancy_pdf` / `reports.cross_reference_xlsx` on the same temporary database, open the PDF with the Read tool and the xlsx via Excel COM → PDF (copy, read-only, `ExportAsFixedFormat`), and check layout: section bars, yellow SO_2 rows, totals, logos, page numbers, no clipped columns. Fix layout issues in `reports.py` with a test where one applies, rerun `python -m pytest check.py -q`, and commit:

```powershell
git add reports.py check.py
git commit -q -m @'
fix: discrepancy and cross reference layout on real data

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>
'@
```

(Skip the commit if nothing changed.)
