# eBiz Import, Discrepancy Report and Cross Reference — Design

Date: 2026-10-08. Extends `2026-10-07-carton-tracker-design.md` (phase 2, first part).

## Purpose

Compare the carton tracker in the database with eBiz exports, per packaging supplier,
and produce the two outputs the old `eBiz Report Generator.xlsm` made from its
`CrossReference` and `DiscrepancyReport` macros:

- **Discrepancy Report** (PDF): tracker orders missing from eBiz, quantity mismatches,
  tracker rows without an SO#, and eBiz orders missing from the tracker.
- **Cross Reference** (xlsx): tracker vs eBiz quantity and rebate per Sales Order.

Out of scope: Rebate Report and Matalan Rebate (later).

## User decisions

| Topic | Decision |
|---|---|
| eBiz files | One export file per vendor. Several files are uploaded together. |
| Replacement | Every eBiz upload deletes all existing eBiz rows, then inserts the uploaded files. |
| Period | Compare **all** tracker rows of the supplier (any date, including rows without a Shipment Date) with all of its eBiz rows. eBiz has no dates. |
| Format | Discrepancy Report as one PDF (includes the eBiz-only section); Cross Reference as xlsx. |
| Matching | Same rules as the old macros (below). |

## eBiz source file

Sample (old generator, sheet `eBiz data`, an Excel table): header in row 1, data from row 2.

| Header | Example |
|---|---|
| Vendor | `Uniglory Packaging Industries ` |
| Vendor Code | `M0068` |
| Product | `MLN-CARTONA2` |
| Order Number | `7124561\|7124562` (or a single number) |
| Qty Shipped | `236` |
| Current Net Price per 1000 | `840` |
| Current Gross Price per 1000 | `880` |
| Total Rebate | `9.44` |
| Currency | `USD` |

The old generator added `SO_1` and `SO_2` columns; if present they are ignored and
recomputed.

## Data model

### Table `ebiz`

| Column | Type | Source |
|---|---|---|
| id | INTEGER PK | |
| vendor | TEXT | Vendor |
| vendor_code | TEXT | Vendor Code |
| product | TEXT | Product |
| order_number | TEXT | Order Number |
| qty_shipped | REAL | Qty Shipped |
| net_price_1000 | REAL | Current Net Price per 1000 |
| gross_price_1000 | REAL | Current Gross Price per 1000 |
| total_rebate | REAL | Total Rebate |
| currency | TEXT | Currency |
| so_1, so_2 | TEXT | Order Number split on `\|`, trimmed; blank → NULL |
| vendor_key | TEXT | `core.supplier_key(vendor)` (upper-cased first word) |
| source_file | TEXT | upload file name |
| imported_at | TEXT | ISO timestamp |

`SCHEMA_VERSION` 3 adds the table (`CREATE TABLE IF NOT EXISTS`; existing data kept).

SO numbers are normalised like the tracker's `ebiz_so`: a whole-number float becomes
its integer text (`7124561.0` → `"7124561"`); text is trimmed. `"0"` and blank mean
"no SO#" on both sides.

## Import (Import tab, section "eBiz data")

1. Multi-file uploader (`.xlsx`, `.xlsm`), separate from the tracker uploader.
2. Each file: first sheet; header row = first of rows 1–10 with an `Order Number`
   cell (headers normalised like the tracker's). Required: Vendor, Order Number,
   Qty Shipped, Total Rebate. Optional columns are stored when present, else NULL.
   Rows with blank Order Number and blank Qty Shipped are skipped, and so are rows
   with a blank Vendor (e.g. total lines). A file that fails is listed with its error
   and left out.
3. Preview per file and vendor key: vendor spelling(s), rows, total Qty Shipped,
   total rebate.
4. Checks:
   - Two files with the same vendor key: blocked, as with tracker clashes.
   - A vendor key with no tracker rows: warning ("can't be reported until its
     tracker is imported").
5. Status line above the uploader: current eBiz data (rows, files, import time) or
   "No eBiz data yet".
6. Button **Replace all eBiz data**. Its caption states that the current N eBiz rows
   are deleted. On click, one transaction: `DELETE FROM ebiz`, insert all rows.
   Reports built before are cleared from the session (`built`), as on tracker import.

## Matching (`core.cross_reference(con, supplier_key)`)

Port of the old macros:

1. Tracker rows: every row of `v_std` and `v_cus` with this `supplier_key`, any date.
   Tracker rebate per row = `pacd + matalan_rebate` (NULL counts as 0, e.g. undated
   rows). This is the old tool's "Total Rebate (Tracker)" (columns V + W).
2. Tracker SO set: distinct non-blank `ebiz_so` of those rows.
3. eBiz rows: `vendor_key` = supplier key. Each row resolves to a key:
   - `so_1` if it is in the tracker SO set; else
   - `so_2` if it is in the tracker SO set (row is "matched via SO_2"); else
   - eBiz-only, keyed by `so_1`, else `so_2`.
4. Per tracker SO#: tracker QTY, tracker rebate, eBiz QTY and eBiz rebate (sums of the
   eBiz rows resolved to it). Status:
   - **Not in eBiz**: no eBiz rows resolved to it.
   - **Tracker > eBiz** / **Tracker < eBiz**: quantities differ.
   - **Match**: equal quantities.
5. Tracker rows without an SO#: one group (QTY, rebate).
6. eBiz-only: eBiz rows whose key is not a tracker SO#, and their per-key totals.

Returns a plain structure (dicts/lists) used by both reports. No result is stored.

## Reports

Built for every picked supplier that has eBiz rows. They ignore the fiscal-year and
quarter pickers (all rows are compared). The Reports tab gets a checkbox
**Discrepancy Report + Cross Reference**, on by default when eBiz data exists and
disabled otherwise. Files go into each supplier's folder, the zip and "Save all".

### Discrepancy Report (PDF)

`Discrepancy Report_<supplier>_<dd-Mmm-yyyy>.pdf` (generation date).

- Landscape A4, same look as the Full Tracker PDFs: navy title bar
  `Discrepancy Report (<supplier>)`, logos, generated date, `Page x of y`.
- Sections, each with a coloured section bar:
  1. Tracker Orders Not Found in eBiz
  2. Quantity Mismatch (Tracker > eBiz)
  3. Quantity Mismatch (Tracker < eBiz)
  4. Tracker Rows With No SO#
  5. eBiz Orders Not Found in Tracker
- Sections 1–4 list tracker rows with the Full Tracker for PacD columns, in a
  Standard Carton block and a Custom Carton block, each with a total line
  (Carton QTY, Total Amount, PACD, Matalan rebate); sections 2–3 add an
  **eBiz QTY** column (the SO#'s eBiz total). Rows are sorted by SO#, then
  shipment date.
- Section 5 lists eBiz rows: SL, Vendor, Vendor Code, Product, Order Number,
  Qty Shipped, Net/1000, Gross/1000, Total Rebate, Currency, with a total line.
- Highlight (pale yellow): tracker rows whose SO# was matched via SO_2; eBiz rows
  with blank SO_1 and a non-blank SO_2.
- A section with nothing to show prints "No discrepancies found."

### Cross Reference (xlsx)

`Cross Reference_<supplier>_<dd-Mmm-yyyy>.xlsx`, styled like the Summary workbook
(title block with logos, navy header rows, thin row lines, shaded total, A4
portrait fit to one page wide, header row repeated on each printed page, no gridlines, `Page x of y` footer).

- Table 1 "Tracker orders": SO#, Tracker QTY, Tracker Rebate, eBiz QTY, eBiz Rebate,
  QTY Difference (tracker − eBiz), Status. Sorted by SO# (numeric when all digits).
  Then a "(No SO#)" line when there are such rows, then a Total line.
- Table 2 "eBiz orders not in tracker" is on its own second sheet, "eBiz not in
  tracker" (print titles repeat per sheet), with its own title block and print
  header: Order, eBiz QTY, eBiz Rebate, Total line; "None" when empty.
- Zero shows as `-`; quantities `#,##0`, money `#,##0.00`.

## Files

| File | Change |
|---|---|
| `core.py` | `ebiz` table, `read_ebiz`, `ebiz_preview`, `ebiz_clashes`, `replace_ebiz`, `ebiz_status`, `ebiz_vendors`, `cross_reference` |
| `reports.py` | `discrepancy_pdf`, `cross_reference_xlsx`; `build_all` / `save_all` gain the option |
| `app.py` | eBiz section on Import tab; checkbox on Reports tab |
| `check.py` | tests below |

No new dependencies.

## Testing (`check.py`)

Synthetic eBiz workbooks (`make_ebiz` helper):

1. Import reads header row 1 or 2, splits `A|B`, normalises `7124561.0`, skips blank
   rows, rejects a file without `Order Number`.
2. Replace: a second upload leaves only the new rows; a failing batch leaves the old
   rows.
3. Two files with the same vendor key clash.
4. Matching: SO_1 match, SO_2 fallback (flagged), eBiz-only, Not in eBiz, Tracker >
   eBiz, Tracker < eBiz, Match, no-SO# group, undated tracker rows included,
   tracker rebate = PACD + Matalan rebate.
5. Reports: PDF text contains each section title and "No discrepancies found." for
   empty ones; xlsx table values and totals; `build_all` adds both files only for
   suppliers with eBiz rows and only when the option is on.
6. App: eBiz upload preview and replace (stubbed uploader), Reports checkbox.

Real-data check: import the Uniglory eBiz rows from the old generator (copy) and the
Uniglory tracker (Dec-25 to Aug-26, as already imported); per-SO quantities must
equal the generator's `Cross Reference` sheet, which was built from the same data on
15 Sep. Rebates differ only for Q3 Custom rows (5% rebate).
