# Reports and Data Management Update Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** New quarter rule, PACD Rebate / Matalan Rate renames, no "Streamlit" tab title, eBiz data and tracker delete tools, and statement-style PDFs with matching Excel twins.

**Architecture:** Quarters and fiscal years stay computed in the `v_std`/`v_cus` views, so one SQL change moves every report. New writes live in `core.py` as small transactional functions that bump `data_revision`; pages only call them. PDF restyling stays inside `reports.py`, sharing one header/footer/table style across Rebate, Full Tracker and Discrepancy.

**Tech Stack:** Python 3.14, Streamlit 1.61.1, psycopg 3.3.6 / PostgreSQL 18, reportlab 5, openpyxl 3.1.5, pytest + streamlit AppTest.

**Spec:** `docs/superpowers/specs/2026-11-09-reports-and-data-management-design.md`

## Global Constraints

- No report column is removed, merged or hidden, except **L** and **W**, which are dropped from the **PDFs only**; Excel outputs keep every column.
- Money prints in USD with `$` (`$1,234.56`); quantities `1,234`; dates `01 Aug 2026` in PDFs.
- The four report figures (cartons, total amount, PACD Rebate, Matalan Rebate) have equal size and colour.
- Q1 2026 = Dec 2025 – Mar 2026; from then on calendar quarters; months before Dec 2025 use calendar quarters.
- Labels: `PACD Rebate`, `Matalan Rate`, `PACD Rebate (4%)`, `Matalan Rebate (5%)`; "PACD Rate" and "Matalan rebate" unchanged.
- Every core write (edit, delete) runs in one `con.transaction()` and calls `core._bump_revision(con)`; pages call `web_ui.clear_built_reports()` after it.
- Tests run against local PostgreSQL (`TEST_DATABASE_URL`, default `postgresql://postgres:postgres@localhost:5432/postgres`; start with `docker start carton-pg`).
- Commits end with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`; git identity for this repo is already set.

## Review Focus

1. **Stale delete:** a second tab deletes rows the first tab still shows; deleting them again must report 0 and not error. Test in Task 4.
2. **Undated tracker rows** (no shipment date) must not be touched by "supplier + month" deletes and must still be deletable by row or by file. Test in Task 4.
3. **eBiz Order Number edited to blank:** the row stays, with SO_1/SO_2 both empty, and shows up as eBiz-only. Test in Task 5.
4. **Quarter edges:** Nov 2025 → (2025, 4); Dec 2026 → (2026, 4), not 2027. Test in Task 1.
5. **Report caches from before the quarter change** must not be served after deploy. Test in Task 1 (cache keys change with `CACHE_VERSION`).

---

### Task 1: Quarter rule, year-aware labels, cache version

**Files:**
- Modify: `core.py` (`VIEW_SQL`, `QUARTER_MONTHS` → `quarter_months`, `quarter_label`, `summary_rows`, `cache_key`)
- Modify: `reports.py` (`summary_xlsx` block labels and quarter note)
- Modify: `pages/1_Dashboard.py` (chart `order`)
- Test: `check.py`

**Interfaces:**
- Produces: `core.quarter_months(fiscal_year: int, quarter: int) -> str` (`"Dec-Mar"` only for (2026, 1); else `"Jan-Mar"`, `"Apr-Jun"`, `"Jul-Sep"`, `"Oct-Dec"`). `core.quarter_label(fy, q) -> str` unchanged in signature. `core.CACHE_VERSION = 2`, mixed into every `core.cache_key`.

- [ ] **Step 1: Update and add tests**

In `test_quarter_mapping` replace `cases` with:
```python
cases = {
    "2025-11-30": (2025, 4), "2025-12-01": (2026, 1), "2026-03-31": (2026, 1), "2026-04-01": (2026, 2),
    "2026-08-31": (2026, 3), "2026-09-30": (2026, 3), "2026-10-01": (2026, 4), "2026-12-31": (2026, 4),
    "2027-01-01": (2027, 1),
}
```
Add:
```python
def test_quarter_labels_follow_the_year():
    assert core.quarter_label(2026, 1) == "Q1 2026 (Dec-Mar)"
    assert core.quarter_label(2026, 3) == "Q3 2026 (Jul-Sep)"
    assert core.quarter_label(2026, 4) == "Q4 2026 (Oct-Dec)"
    assert core.quarter_label(2027, 1) == "Q1 2027 (Jan-Mar)"

def test_cache_keys_change_with_the_cache_version(monkeypatch):
    before = core.cache_key("x", 1)
    monkeypatch.setattr(core, "CACHE_VERSION", core.CACHE_VERSION + 1)
    assert core.cache_key("x", 1) != before
```
Replace every `Q3 (Jul-Aug)` with `Q3 (Jul-Sep)` and `Q3 2026 (Jul-Aug)` with `Q3 2026 (Jul-Sep)` in existing assertions.

- [ ] **Step 2: Run** `python -m pytest check.py -q -k "quarter or cache_keys or summary or full_tracker"` — Expected: FAIL on the new mappings and labels.

- [ ] **Step 3: Implement**
- `VIEW_SQL`: `fiscal_year` = `EXTRACT(YEAR ...)::int + (date_trunc('month', t.shipment_date::date) = DATE '2025-12-01')::int`; `quarter` = `NULL` when no date, `1` for December 2025, else `((EXTRACT(MONTH ...)::int + 2) / 3)`.
- `quarter_months` as in Interfaces; `quarter_label` and `summary_rows` use it; delete `QUARTER_MONTHS`.
- `reports.summary_xlsx`: block label `f"Q{q} ({core.quarter_months(fiscal_year, q)})"`; the note lists the quarters of that fiscal year from `quarter_months`.
- Dashboard `order`: built with `core.quarter_months(year, q)` for the selected year.
- `cache_key`: hash `CACHE_VERSION` before the namespace.

- [ ] **Step 4: Run** the full suite `python -m pytest check.py -q -p no:cacheprovider` — Expected: all pass.

- [ ] **Step 5: Commit** `git commit -m "Use calendar quarters after the Dec 2025 – Mar 2026 first quarter."`

---

### Task 2: Renames

**Files:**
- Modify: `core.py` (`rate_headers`), `reports.py` (`SUMMARY_HEADERS`, Summary note text), `pages/1_Dashboard.py`, `pages/7_Discrepancy_Report.py`, `pages/8_Rebate_Report.py`
- Test: `check.py`

**Interfaces:**
- Produces: `core.rate_headers(rows) -> (str, str)` returning `("PACD Rebate (4%)", "Matalan Rebate (5%)")` for one shared rate, `("PACD Rebate", "Matalan Rebate")` when mixed/unrated/empty.

- [ ] **Step 1: Update tests**
```python
def test_rate_headers():
    ...  # same fixtures as today
    assert core.rate_headers(one) == ("PACD Rebate (7.5%)", "Matalan Rebate (5%)")
    assert core.rate_headers(mixed) == ("PACD Rebate (7.5%)", "Matalan Rebate")
    assert core.rate_headers(unrated) == ("PACD Rebate", "Matalan Rebate")
    assert core.rate_headers([]) == ("PACD Rebate", "Matalan Rebate")

def test_pages_use_the_new_rebate_labels():
    for path in (HERE / "pages").glob("*.py"):
        text = path.read_text(encoding="utf-8")
        assert '"Rebate rate"' not in text and '"PACD"' not in text, path.name
```
Update Raw Data assertions to `["PACD Rebate (4%)", "Matalan Rebate (5%)", "NOTE"]` / `["PACD Rebate (7.5%)", "Matalan Rebate (5%)"]`, Summary headers to `[..., "PACD Rate", "Matalan Rate", ..., "PACD Rebate", "Matalan Rebate"]`, and PDF text checks from `"PACD"` / `"Rebates"` / `"for Matalan"` to `"PACD Rebate"` / `"Matalan Rebate"`. The Matalan PDF check stays `"PACD" not in text`.

- [ ] **Step 2: Run** `python -m pytest check.py -q -k "rate_headers or labels or summary or raw_data or pdf"` — Expected: FAIL.

- [ ] **Step 3: Implement** the label table from spec section 2 in the listed files (page column renames, KPI labels, chart tooltips; Summary note says "Matalan Rate").

- [ ] **Step 4: Run** the full suite — Expected: all pass.

- [ ] **Step 5: Commit** `git commit -m "Rename PACD to PACD Rebate and Rebate rate to Matalan Rate."`

---

### Task 3: Browser tab title

**Files:** Modify `Dockerfile`; Test `check.py`

- [ ] **Step 1: Test**
```python
def test_docker_build_replaces_the_streamlit_tab_title():
    text = (HERE / "Dockerfile").read_text(encoding="utf-8")
    assert "s|<title>Streamlit</title>|<title>Matalan Carton Supply Report</title>|" in text
```
- [ ] **Step 2: Run** it — Expected: FAIL.
- [ ] **Step 3: Implement** the `RUN sed -i ...` line from spec section 3, after `RUN pip install -r requirements.txt`.
- [ ] **Step 4: Verify** test passes, then `docker build -q -t matalan-carton-report .` and `docker run --rm --entrypoint sh matalan-carton-report -c "grep -o '<title>[^<]*' $(python -c 'import streamlit,os;print(os.path.dirname(streamlit.__file__))')/static/index.html"`. Expected: `<title>Matalan Carton Supply Report`. (Run the inner `python -c` inside the container: wrap the whole command in single quotes for `sh -c`.)
- [ ] **Step 5: Commit** `git commit -m "Show the app name in the browser tab while Streamlit loads."`

---

### Task 4: Deleting tracker data

**Files:** Modify `core.py`, `pages/3_Tracker.py`; Test `check.py`

**Interfaces:**
- Produces:
  - `core.delete_tracker_rows(con, kind: str, ids: list[int]) -> int`
  - `core.delete_supplier_month(con, supplier: str, month: str) -> dict[str, int]` (`month` `"YYYY-MM"`; returns `{"STD": n, "CUS": n}`)
  - `core.delete_source_file(con, source_file: str) -> dict[str, int]`
  - `core.tracker_files(con) -> list[dict]` with keys `source_file`, `STD`, `CUS`, `imported_at` (latest), newest first
  - `core.tracker_months(con, supplier: str) -> list[str]` (`YYYY-MM` with dated rows, newest first)
- All three deletes: one transaction, `_bump_revision` only when something was deleted; `kind` validated like `editable_rows`.

- [ ] **Step 1: Tests** (use `insert` and `make_xlsx` helpers)
```python
def test_delete_selected_tracker_rows_and_a_stale_repeat():
    con = core.connect()
    for qty in (1, 2, 3):
        insert(con, "STD", shipment_date="2026-08-03", supplier_key="PACK", carton_qty=qty)
    ids = [r["id"] for r in core.editable_rows(con, "STD", "PACK")][:2]
    revision = core.data_revision(con)
    assert core.delete_tracker_rows(con, "STD", ids) == 2
    assert [r["carton_qty"] for r in core.editable_rows(con, "STD", "PACK")] == [3]
    assert core.data_revision(con) != revision
    assert core.delete_tracker_rows(con, "STD", ids) == 0  # already gone in another tab

def test_delete_supplier_month_clears_both_sheets_but_not_undated_rows():
    con = core.connect()
    insert(con, "STD", shipment_date="2026-08-03", supplier_key="PACK", carton_qty=1)
    insert(con, "CUS", shipment_date="2026-08-20", supplier_key="PACK", carton_qty=2)
    insert(con, "STD", shipment_date="2026-07-31", supplier_key="PACK", carton_qty=3)
    insert(con, "STD", shipment_date="2026-08-03", supplier_key="OTHER", carton_qty=4)
    insert(con, "STD", shipment_date=None, supplier_key="PACK", carton_qty=5)
    assert core.tracker_months(con, "PACK") == ["2026-08", "2026-07"]
    assert core.delete_supplier_month(con, "PACK", "2026-08") == {"STD": 1, "CUS": 1}
    left = sorted(r["carton_qty"] for k in core.TABLES for r in core.editable_rows(con, k))
    assert left == [3, 4, 5]

def test_delete_source_file(tmp_path):
    con = core.connect()
    core.save_many(con, {"a.xlsx": core.read_workbook(make_xlsx(tmp_path / "a.xlsx", std=[rec()], cus=[rec()])),
                         "b.xlsx": core.read_workbook(make_xlsx(tmp_path / "b.xlsx", std=[rec(sup="Other Ltd")]))})
    assert [(f["source_file"], f["STD"], f["CUS"]) for f in core.tracker_files(con)] in (
        [("a.xlsx", 1, 1), ("b.xlsx", 1, 0)], [("b.xlsx", 1, 0), ("a.xlsx", 1, 1)])
    assert core.delete_source_file(con, "a.xlsx") == {"STD": 1, "CUS": 1}
    assert [f["source_file"] for f in core.tracker_files(con)] == ["b.xlsx"]

def test_tracker_page_offers_the_delete_tools(xref_con):
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(HERE / "pages" / "3_Tracker.py"), default_timeout=60).run()
    assert not at.exception
    assert at.radio(key="tracker_delete_mode").options == ["Selected rows", "Supplier and month", "Imported file"]
```
- [ ] **Step 2: Run** `-k "delete or offers_the_delete"` — Expected: FAIL.
- [ ] **Step 3: Implement** the core functions, then on the Tracker page:
  - add a leading `Delete` checkbox column to the editor frame (`st.column_config.CheckboxColumn`, default False, not in `core.COLUMNS`, ignored by the save diff);
  - under the editor, `st.expander("Delete tracker data")` with `st.radio("Delete", [...], key="tracker_delete_mode", horizontal=True)`:
    - **Selected rows:** count the ticked rows; button "Delete N selected rows".
    - **Supplier and month:** month select from `tracker_months` for the chosen supplier; shows "N Standard and M Custom rows".
    - **Imported file:** select from `tracker_files` labelled `"{file} · {STD} STD · {CUS} CUS · {imported_at[:16]}"`.
  - each button opens an `@st.dialog("Delete tracker data")` repeating exactly what goes, the line "This cannot be undone here. RDS keeps 7 days of backups.", and a primary "Delete" button; on confirm: core call, `clear_built_reports()`, toast with the count, `st.rerun()`.
- [ ] **Step 4: Run** the full suite — Expected: all pass.
- [ ] **Step 5: Commit** `git commit -m "Delete tracker rows by selection, supplier month or imported file."`

---

### Task 5: eBiz data page

**Files:** Create `pages/9_Ebiz_Data.py`; Modify `core.py`, `app.py`; Test `check.py`

**Interfaces:**
- Produces:
  - `core.ebiz_rows(con, vendor: str | None = None, limit: int | None = None, offset: int = 0) -> list[dict]` (columns `id`, `EBIZ_COLUMNS`, `source_file`, `imported_at`; ordered by id)
  - `core.ebiz_row_count(con, vendor: str | None = None) -> int`
  - `core.update_ebiz_rows(con, edited: list[dict]) -> int` (each dict has `id` plus `EBIZ_HEADERS` keys; recomputes `vendor_key`, `so_1`, `so_2` with the import rules; `ValueError` "Row N: Vendor is required." / "Row N: <header> must be a number." / unknown id)
  - `core.delete_ebiz_rows(con, ids: list[int]) -> int`
- Consumes: `core.supplier_key`, `core._so`, `core._clean`, Task 4's dialog/checkbox pattern.

- [ ] **Step 1: Tests**
```python
def test_ebiz_edit_recomputes_keys_and_keeps_rows_with_a_blank_order(tmp_path):
    con = core.connect()
    core.replace_ebiz(con, {"e.xlsx": core.read_ebiz(make_ebiz(tmp_path / "e.xlsx", [erow("7001"), erow("7002")]))})
    first, second = core.ebiz_rows(con)
    revision = core.data_revision(con)
    assert core.update_ebiz_rows(con, [{**first, "vendor": "Other Co", "order_number": "8001|8002"},
                                       {**second, "order_number": None}]) == 2
    first, second = core.ebiz_rows(con)
    assert (first["vendor_key"], first["so_1"], first["so_2"]) == ("OTHER", "8001", "8002")
    assert (second["so_1"], second["so_2"]) == (None, None)
    assert core.data_revision(con) != revision
    with pytest.raises(ValueError, match="Vendor is required"):
        core.update_ebiz_rows(con, [{**first, "vendor": " "}])
    assert core.delete_ebiz_rows(con, [second["id"]]) == 1
    assert [r["id"] for r in core.ebiz_rows(con)] == [first["id"]]
    assert core.delete_ebiz_rows(con, [second["id"]]) == 0

def test_ebiz_data_page(xref_con):
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(str(HERE / "pages" / "9_Ebiz_Data.py"), default_timeout=60).run()
    assert not at.exception
    assert at.selectbox[0].label == "Vendor"
    assert "Save eBiz changes" in [b.label for b in at.button]
```
- [ ] **Step 2: Run** `-k ebiz_edit or ebiz_data_page` — Expected: FAIL.
- [ ] **Step 3: Implement** the core functions (`_bump_revision` on write), the page (vendor select from `ebiz_vendors` with "All vendors" first, 250-row pages, editor with `Delete` checkbox, hidden derived columns, read-only id/source/imported, "Save eBiz changes", delete dialog as in Task 4), and register it in `app.py` under "Manage" after Tracker data: `st.Page("pages/9_Ebiz_Data.py", title="eBiz data", icon=":material/receipt_long:", url_path="ebiz-data")`. Empty state: "No eBiz data yet. Import vendor exports on the Data import page."
- [ ] **Step 4: Run** the full suite — Expected: all pass.
- [ ] **Step 5: Commit** `git commit -m "Add an eBiz data page to view, edit and delete eBiz rows."`

---

### Task 6: Statement-style PDFs

**Files:** Modify `reports.py`; Test `check.py`

**Interfaces:**
- Produces: `reports.PDF_COLS = [c for c in TRACKER_COLS if c not in ("l_mm", "w_mm")]`; tracker and discrepancy PDFs use it (Matalan brand also drops `pacd`). Public function names and return values unchanged.
- Consumes: `core.quarter_label`, `core.quarter_months`, `core.rate_headers` (Tasks 1–2).

- [ ] **Step 1: Tests**
```python
def test_statement_pdf_layout(demo_con):
    name, data = reports.rebate_report_pdf(demo_con, TW_KEY, 2026, [3])
    text = pdf_text(data)
    for expected in ("Rebate statement", "Transworld Packaging (BD) Ltd.", "Q3 2026 · 1 Jul – 30 Sep 2026",
                     "All tracker rows", "Cartons shipped", "Total amount", "PACD Rebate", "Matalan Rebate",
                     "Sheet 1 of", "01 Aug 2026", "Length (mm)"):
        assert expected in text, expected
    assert "l_mm" not in reports.PDF_COLS and "w_mm" not in reports.PDF_COLS
    assert "1,851" not in text  # the first row's L value: the L column is gone from the PDF
    assert "8/1/2026" not in text
    assert len(pypdf.PdfReader(BytesIO(data)).pages[0].images) == 2

def test_statement_pdf_variants(demo_con, xref_con):
    so_text = pdf_text(reports.rebate_report_pdf(xref_con, "PACK", 2026, [3], so_only=True)[1])
    assert "Rows with eBiz SO#" in so_text
    matalan = pdf_text(reports.full_tracker_pdf(demo_con, TW_KEY, 2026, [3], "Matalan")[1])
    assert "PACD" not in matalan and "Full tracker" in matalan
    disc = pdf_text(reports.discrepancy_pdf(xref_con, "PACK", core.cross_reference(xref_con, "PACK"))[1])
    assert "All shipments · eBiz data loaded" in disc and "Discrepancy report" in disc
```
Update existing PDF tests whose strings changed (title bars, `Page 1 of` → `Sheet 1 of`, `Generated:` stays in the footer).
- [ ] **Step 2: Run** `-k "pdf"` — Expected: FAIL.
- [ ] **Step 3: Implement** spec section 6 in `reports.py`:
  - palette constants `INK #1D1C1A`, `INK2 #57534D`, `INK3 #8B857C`, `HAIR` = ink at 12% alpha, `KRAFT #F5F0E6`, `NAVY #1F497D`;
  - `_pdf(story, title, footer_left)`: logos 0.25 in high, no title bar; footer left text, generated time centred, right a double-outlined box "Sheet N of M" (draw in `_NumberedCanvas.save`);
  - `_statement(kind, supplier_name, period, row_line, figures, meta)` flowable: shipping-mark box (2 pt + 0.75 pt outlines, 3 pt apart) on the left, figures as equal label/value pairs on the right;
  - period from the selected quarters: first day of the first quarter to last day of the last (Q1 2026 starts 1 Dec 2025);
  - tables: no header fill, 0.75 pt ink rule under the header, hairline row rules, quarter rows navy text, subtotal/total rows kraft + bold, per-column widths, `01 Aug 2026` dates, `$` money, headers from `_headers` with readable spacing (`Length (mm)`, `Width (mm)`, `Height (mm)`, `Area (sqm)`).
  - Discrepancy: statement figures are the five section counts; period line "All shipments · eBiz data loaded <date>" (date from `core.ebiz_status`).
- [ ] **Step 4: Render check:** generate the Transworld Rebate PDF and render page 1 to PNG with `pypdfium2` (already in `.venv`, not in requirements); compare with the approved mockup (statement block, equal figures, no L/W). Then run the full suite. Expected: all pass.
- [ ] **Step 5: Commit** `git commit -m "Restyle PDFs as rebate statements with a shipping-mark header."`

---

### Task 7: Excel twins get the statement header

**Files:** Modify `reports.py`; Test `check.py`

- [ ] **Step 1: Test**
```python
def test_excel_twins_have_the_statement_header_and_keep_l_and_w(demo_con):
    _, data = reports.rebate_report_xlsx(demo_con, TW_KEY, 2026, [3])
    rows = list(openpyxl.load_workbook(BytesIO(data)).active.iter_rows(values_only=True))
    cells = [c for r in rows for c in r if c is not None]
    for expected in ("Rebate statement", "Transworld Packaging (BD) Ltd.", "Q3 2026 · 1 Jul – 30 Sep 2026",
                     "Cartons shipped", "PACD Rebate", "Matalan Rebate"):
        assert expected in cells, expected
    header = next(r for r in rows if "Carton QTY" in r)
    assert "L" in header and "W" in header
```
- [ ] **Step 2: Run** it — Expected: FAIL.
- [ ] **Step 3: Implement** a shared `_xlsx_statement(ws, kind, supplier_name, period, row_line, figures) -> int` (returns first free row) used by the three twins; table header styling = no fill, bottom border ink; subtotals kraft fill. Keep `xlsx_total` lookups working (labels in column A unchanged).
- [ ] **Step 4: Run** the full suite — Expected: all pass.
- [ ] **Step 5: Commit** `git commit -m "Give the Excel reports the same statement header as the PDFs."`

---

### Task 8: Docs, final review, deploy

**Files:** Modify `README.md`, `CHANGELOG.md`

- [ ] **Step 1:** README: pages table gains eBiz data; Tracker data mentions the delete tools; quarter rule sentence; renamed labels. CHANGELOG Unreleased: one line per spec item.
- [ ] **Step 2:** Full suite green; whole-branch review (executing-plans final review); fixes with RED→GREEN tests.
- [ ] **Step 3:** With the user's go: push `main`, then on the server `cd /opt/matalan-carton && git pull && docker compose up -d --build`; verify health `ok`, tracker row count unchanged (5,919 at last check), tab title, and `https://matalan.pac-d.com/carton-supply-report/` returns 401.
