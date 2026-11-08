# Reports and data management update

**Date:** 2026-11-09
**Status:** Design approved in conversation; awaiting spec review

## Goal

Six changes requested on 2026-11-09:

1. View, edit and delete eBiz data.
2. Delete old tracker data (rows, a supplier's month, or a whole imported file).
3. The browser tab must not say "Streamlit".
4. The exported PDFs look professional: a rebate statement in the style approved in the mockup.
5. Rename "PACD" to "PACD Rebate" and "Rebate rate" to "Matalan Rate".
6. New quarter rule: Q1 2026 is Dec 2025 – Mar 2026; every quarter after that is a calendar quarter.

Success: each change is visible on the live site at `https://matalan.pac-d.com/carton-supply-report/`,
existing data (5,919 tracker rows, eBiz rows, rates, date corrections) is preserved, and the full test
suite passes.

## Fixed rules (from the user)

- No report column is removed, merged or hidden, except **L** and **W**, which are dropped from the **PDFs
  only**. Excel outputs keep every column, L and W included.
- Money (Final Price, Total Amount, PACD Rebate, Matalan Rebate) is USD and prints with `$`.
- The four report figures (cartons, total amount, PACD Rebate, Matalan Rebate) get equal visual weight.

## 1. Quarters

Fiscal year and quarter come from the shipment date in the `v_std` / `v_cus` views (`core.VIEW_SQL`).

| Shipment month | Fiscal year | Quarter |
|---|---|---|
| Dec 2025 | 2026 | 1 |
| Jan – Mar Y | Y | 1 |
| Apr – Jun Y | Y | 2 |
| Jul – Sep Y | Y | 3 |
| Oct – Dec Y (Y ≥ 2026) | Y | 4 |
| Any month before Dec 2025 | calendar year | calendar quarter |

- `fiscal_year` = calendar year, plus 1 only for December 2025.
- `quarter` = 1 for December 2025, otherwise `ceil(month / 3)`.
- Labels depend on the year: `core.quarter_months(fiscal_year, quarter)` returns `Dec-Mar` for (2026, 1),
  otherwise `Jan-Mar`, `Apr-Jun`, `Jul-Sep`, `Oct-Dec`. It replaces the `QUARTER_MONTHS` dict everywhere it is
  used: `core.quarter_label`, `core.summary_rows` (`quarter_label` field), `reports.summary_xlsx` (block labels
  and the quarter note under the Summary), and the Dashboard chart order.
- `core.quarter_label(2026, 1)` = `Q1 2026 (Dec-Mar)`; `core.quarter_label(2026, 4)` = `Q4 2026 (Oct-Dec)`;
  `core.quarter_label(2027, 1)` = `Q1 2027 (Jan-Mar)`.
- No data migration: the views recalculate on every query, so stored rows move to their new quarter at once.
  Report caches are keyed on data and must also include the quarter rule; bump every cache namespace version.

## 2. Renames

Applied on every page, PDF and Excel output.

| Today | New |
|---|---|
| `PACD` (column, KPI, chart and summary labels) | `PACD Rebate` |
| `Rebate rate` / `Rebate Rate` | `Matalan Rate` |
| `4% for PACD` (rate-aware report header) | `PACD Rebate (4%)` |
| `5% Rebates for Matalan` | `Matalan Rebate (5%)` |
| Mixed-rate fallbacks `PACD` / `Rebates for Matalan` | `PACD Rebate` / `Matalan Rebate` |

`core.rate_headers` produces the two rate-aware headers. "PACD rate" stays "PACD Rate"; "Matalan rebate" stays.

## 3. Browser tab title

Streamlit's static `index.html` ships `<title>Streamlit</title>`, shown until the app script runs. The
Dockerfile replaces it after `pip install`:

```
RUN sed -i 's|<title>Streamlit</title>|<title>Matalan Carton Supply Report</title>|' \
    "$(python -c 'import streamlit, os; print(os.path.dirname(streamlit.__file__))')/static/index.html"
```

The build fails if the file is missing (`sed` on a missing path exits non-zero), so a future Streamlit upgrade
cannot silently skip it. A test asserts the Dockerfile contains this line.

## 4. eBiz data page

New page **Manage → eBiz data** (`pages/9_Ebiz_Data.py`, `url_path="ebiz-data"`).

- Vendor filter (vendor keys from `core.ebiz_vendors`, display names from `core.supplier_name`), 250-row pages
  like Tracker data.
- `st.data_editor` with the `EBIZ_HEADERS` columns editable; `id`, `source_file`, `imported_at` read-only;
  `vendor_key`, `so_1`, `so_2` hidden and derived.
- **Save:** `core.update_ebiz_rows(con, edited) -> int`. One transaction. Validates numbers like
  `update_tracker_rows`; recomputes `vendor_key = supplier_key(vendor)` and `so_1`/`so_2` from Order Number with
  the import rule; a row whose Vendor becomes blank is rejected ("Vendor is required").
- **Delete:** a selection column plus "Delete selected rows" → confirm dialog showing the count →
  `core.delete_ebiz_rows(con, ids) -> int`.
- Every write calls `core._bump_revision` and `web_ui.clear_built_reports()`.

## 5. Deleting tracker data

On **Manage → Tracker data**, below the editor, an expander "Delete tracker data" with three modes. Each shows
the exact rows that will go (count per sheet, supplier, months) and needs a confirm click.

| Mode | Function |
|---|---|
| Selected rows (selection column in the editor page) | `core.delete_tracker_rows(con, kind, ids) -> int` |
| Supplier + shipment month (both sheets) | `core.delete_supplier_month(con, supplier, month) -> dict[str, int]` (`month` = `YYYY-MM`) |
| Imported file | `core.delete_source_file(con, source_file) -> dict[str, int]` |

- `core.tracker_files(con) -> list[dict]`: source file, rows per sheet, last imported time, for the file picker.
- `core.tracker_months(con, supplier) -> list[str]`: shipment months with rows for that supplier, newest first.
- Each delete is one transaction, calls `_bump_revision`, and the page clears built reports.
- Saved date corrections (`date_fixes`) are kept: they only apply if the same row is imported again.
- Deletes are permanent; the confirm text says so and that RDS keeps 7-day backups.

## 6. PDF and Excel style (the approved "statement" mockup)

Applies to every PDF: Rebate (all rows and with-SO#), Full Tracker (PacD, Matalan), Discrepancy.

**Page (A4 landscape)**
- Logos on one line, height 0.25 in (today 0.45 in); PACD left, Matalan right.
- No full-width title bar and no stacked banners.

**Statement block (page 1)**
- Left: the **shipping-mark box**: a double outline (2 pt outer, 0.75 pt inner, 3 pt apart). Inside: report kind
  (e.g. "Rebate statement", "Full tracker", "Discrepancy report"), supplier name, period
  (`Q3 2026 · 1 Jul – 30 Sep 2026`, from the first and last day of the selected quarters), and the row type
  ("All tracker rows" / "Rows with eBiz SO#") with the shipment count.
- Right: four figures in equal size and colour (ink, not navy): Cartons shipped, Total amount,
  PACD Rebate · rate, Matalan Rebate · rate (rate shows "Mixed" when rows differ). The Matalan brand PDF omits
  the PACD Rebate figure, as it omits the column. A meta line below: rows per carton type and generation time.
- Discrepancy PDF: it compares every tracker row, so the period line reads "All shipments · eBiz data loaded
  <date>" and the row-type line is omitted; the right side shows the five section counts (equal weight) instead
  of money figures.

**Tables**
- Header text in secondary ink, normal weight, one 0.75 pt ink rule below; no header fill.
- Rows separated by hairlines (ink at 12%); no zebra striping.
- Quarter heading row: navy text, no fill. Section heading: "Standard carton" + shipment count.
- Subtotal and total rows: kraft fill `#F5F0E6`, bold, 0.75 pt ink rule below.
- Column widths per column (names and references wide, sizes narrow); long references wrap anywhere.
- Dates `01 Aug 2026`. Money `$1,234.56`. Quantities `1,234`.
- Header labels: the renamed labels from section 2 and readable spacing (`Length (mm)`, not `Length(mm)`).

**Palette**
- Ink `#1D1C1A`, secondary `#57534D`, muted `#8B857C`, hairline ink at 12%, kraft `#F5F0E6`, navy `#1F497D`
  (quarter headings only). No other fills.

**Footer (every page)**
- Left: "<report kind> · <supplier> · <period>"; middle: generated time; right: a small shipping-mark box with
  "Sheet N of M".

**Excel twins** (`rebate_report_xlsx`, `full_tracker_xlsx`, `discrepancy_xlsx`) get the same statement header
(figures as values in cells), the same header/subtotal styling and date/money formats, and keep **all**
columns, L and W included. Raw Data, Shipment Tracking, Summary and Cross Reference workbooks only get the
renames and quarter labels.

## Testing

- Quarter mapping: Dec 2025 → (2026, 1); Mar 2026 → (2026, 1); Sep 2026 → (2026, 3); Oct/Dec 2026 → (2026, 4);
  Jan 2027 → (2027, 1); labels for (2026, 1), (2026, 4), (2027, 1).
- Rate headers and page labels use the new names; no "Rebate rate" or bare "PACD" header left in pages.
- Dockerfile contains the title patch.
- eBiz: edit recomputes vendor key and SO_1/SO_2; blank vendor rejected; delete removes only chosen ids; both
  bump `data_revision`.
- Tracker deletes: each mode removes exactly the described rows on both sheets and nothing else; counts returned
  match; revision bumps.
- PDFs: L and W headers absent; statement texts (kind, supplier, period, figures, "Sheet 1 of") present; logo
  images still 2 on page 1; Matalan PDF has no "PACD Rebate". Excel twins keep L and W.
- Existing suites (imports, reports, pages, leak) stay green.

## Out of scope

- Undo for deletes (RDS backups only).
- Editing derived columns (eBiz vendor key, SO_1/SO_2; tracker supplier key, rates, rebates).
- Restyling Raw Data, Shipment Tracking, Summary and Cross Reference workbooks beyond renames and quarter labels.
