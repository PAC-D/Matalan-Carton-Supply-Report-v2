# Matalan Carton Tracker DB — Design (Phase 1)

Date: 2026-10-07
Status: draft, awaiting user review

## Purpose

Replace the tracker-only part of `eBiz Report Generator.xlsm` with a small local app.
The app imports the `TRACK_STD` and `TRACK_CUS` sheets of a Matalan Carton Shipment
Tracking workbook into a SQLite database, keeps the data across imports, and
generates the tracker reports from the database.

## Scope

In scope (phase 1):

- Import `TRACK_STD` and `TRACK_CUS` into two tables.
- Editable, date-effective PACD % and Matalan rebate % per carton type.
- Reports: Summary, Full Tracker Sheet for PacD, Full Tracker Sheet for Matalan,
  Tracker Raw Data, Shipment Tracking export.

Out of scope (phase 1):

- eBiz data import and the 4 reports that need it (Rebate Report, Matalan Rebate,
  Discrepancy Report, eBiz Discrepancy). Planned as phase 2.
- The per-quarter "Tracker Sheet for PacD / Matalan" PDFs (user decision: skip).
  A Full Tracker report with one quarter selected gives the same content.
- Login, multi-user access, hosting.
- Editing individual rows in the UI, charts.

## Decisions (from brainstorming)

| Topic | Decision |
|---|---|
| Users / platform | One user, this PC. Python + SQLite + Streamlit on localhost. |
| Input sheets | Only `TRACK_STD` and `TRACK_CUS`. Other sheets in the workbook are ignored. |
| Calculated columns | Only PACD and Matalan rebate are calculated by the app. All other columns are stored as the values Excel cached when the file was saved. |
| Rates | Date-effective, by Shipment Date. See "Rates" below. |
| Re-import rule | Provisional: replace overlap (see "Import"). To be confirmed when the real files arrive. |
| eBiz data | Phase 2. |

## Source file facts (from the demo `Matalan Carton Shipment Tracking_NPS_BD Aug-26.xlsx`)

- Row 1 holds `SUBTOTAL` formulas, row 2 holds headers, data starts on row 3.
- Both sheets have the same columns A–X (TRACK_CUS has an extra empty column Y).
- Data rows: TRACK_STD 88 (rows 3–90), TRACK_CUS 118 (rows 3–120).
- Template rows below the data have formulas but no input values
  (419 in TRACK_STD, 323 in TRACK_CUS). They must be skipped.
- Some price formulas link to other workbooks (`[1]`–`[4]`). The cached values are
  correct, which is why the app stores cached values instead of re-deriving prices.
- Text values can have trailing spaces (`'Lariz Fashion Ltd '`).
- Column C (Matalan Product PO#) is empty in every demo row.
- The sheets may have an active AutoFilter. The import reads every row regardless.

## Data model

SQLite file `carton.db` in the project folder.

### Tables `track_std` and `track_cus` (identical columns)

| Column | Type | Source header (row 2, normalised) |
|---|---|---|
| id | INTEGER PRIMARY KEY | — |
| po_received_date | TEXT (ISO date) | `po received date` |
| shipment_date | TEXT (ISO date) | `shipment date` |
| matalan_po | TEXT | `matalan product po#` |
| supplier_name | TEXT | `supplier name` |
| factory_name | TEXT | `factory name` |
| packaging_supplier | TEXT | `packaging supplier name` |
| ebiz_customer_ref | TEXT | `ebiz customer reference #` |
| ebiz_so | TEXT | `ebiz sales order #` |
| carton_qty | INTEGER | `carton qty` |
| carton_style | TEXT | `carton style` |
| flute | TEXT | `flute` |
| carton_selection | TEXT | `carton selection` |
| length_mm | REAL | `length(mm)` |
| width_mm | REAL | `width(mm)` |
| height_mm | REAL | `height(mm)` |
| l_mm | REAL | `l` |
| w_mm | REAL | `w` |
| area_sqm | REAL | `area(sqm)` |
| unit_price | REAL | starts with `carton's unit price` |
| final_price | REAL | `final price, cny/ctn` |
| total_amount | REAL | `total amount` |
| note | TEXT | `note` |
| source_file | TEXT | file name of the import |
| imported_at | TEXT | import timestamp |
| supplier_key | TEXT | derived at import (see "Supplier identity") |

### Supplier identity

The tracker spells each packaging supplier several ways (for example
`M&U Packaging Ltd`, `M&U  Packaging Ltd.`, `Uniglory Packging Industries Ltd.`).
The old VBA tool merged them by first word, so the app does too:

- `supplier_key` = the first word of Packaging Supplier, upper-cased, with
  `. , ; : " ' ! ?` removed (`M&U`, `TRANSWORLD`, `UNIGLORY`). It is the supplier in
  the re-import key, every report filter, the Summary blocks and the Reports page.
- Display name = the most common raw spelling of that key. It is used in report
  titles, file names and folders.
- `packaging_supplier` keeps the raw text of each row, so exports mirror the tracker.

Two suppliers whose names start with the same word would merge. That doesn't
happen in the current data; an alias table would replace the first-word rule if
it ever does.

Header normalisation: lower case, every run of whitespace (including line breaks)
collapsed to one space, trimmed. The PACD and rebate columns of the file are not
imported.

### Table `rates`

| Column | Type |
|---|---|
| carton_type | TEXT, `STD` or `CUS` |
| effective_from | TEXT (ISO date) |
| pacd_pct | REAL (0.04 = 4%) |
| matalan_pct | REAL |

Primary key: (carton_type, effective_from).

Seed rows:

| carton_type | effective_from | pacd_pct | matalan_pct |
|---|---|---|---|
| STD | 2025-12-01 (Q1 2026) | 0.04 | 0.05 |
| CUS | 2025-12-01 (Q1 2026) | 0.075 | 0.10 |
| STD | 2026-07-01 (Q3 2026) | 0.04 | 0.05 |
| CUS | 2026-07-01 (Q3 2026) | 0.075 | 0.05 |

The Q3 2026 Custom rebate of 5% is a confirmed change. The demo file and the
August Summary use 10% for Q3 Custom; the app's figures will differ from them
on purpose.

### Views `v_std` and `v_cus`

Each view returns all table columns plus:

- `pacd_pct`, `matalan_pct`: from the `rates` row of the same carton type with the
  latest `effective_from` on or before `shipment_date`. NULL when the row has no
  shipment date.
- `matalan_rebate = unit_price × carton_qty × matalan_pct`
  (same as the Excel formula `=S*I*rate`).
- `pacd = (total_amount − matalan_rebate) × pacd_pct`
  (same as the Excel formula `=((T*I)-W)*rate`).
- `fiscal_year`: calendar year of the shipment date, plus 1 when the month is December.
- `quarter`: 1 for Dec–Mar, 2 for Apr–Jun, 3 for Jul–Aug, 4 for Sep–Nov.

Rate changes made on the Rates page take effect in every report immediately,
because nothing derived from a rate is stored.

## Import

1. User uploads one or more `.xlsx` files on the Import page. Each file is read on
   its own; a file that fails steps 2–5 is listed with its error and left out.
   If two files share a sheet + supplier key + shipment month, the clash is listed
   and the import is blocked until one file is removed. Valid files are saved in
   one transaction, each row keeping its own source file name.
2. Open with openpyxl, `data_only=True`. Find sheets `TRACK_STD` and `TRACK_CUS`
   (case-insensitive). Missing sheet: stop with an error naming the sheet.
3. Find the header row: the first of rows 1–10 with a `Shipment Date` cell (row 2
   in the master tracker, which has a title row; row 1 in the per-supplier exports).
   No such row, or a missing required header: stop with an error naming the sheet.
   Map columns by normalised header (table above).
4. Read every row below the header row to the sheet's max row.
   - Skip a row when both Shipment Date and Carton QTY are blank (template row).
   - Trim text. Treat `#N/A` and other Excel error strings as blank.
   - Dates: Excel datetime to ISO date. Text dates (`18/12/2025`) are read day
     first. Other text is stored as NULL.
5. Guard: if a sheet has data rows but `total_amount` is blank in all of them,
   stop with "The file has no calculated values. Open it in Excel, save, and import
   again." (openpyxl only sees values Excel cached on save.)
6. Date check (see below), then preview: row counts per sheet × supplier key ×
   shipment month, the spellings merged into each key and the number of rows that
   will be replaced. The preview follows the dates chosen in the date check.
7. On confirm, in one transaction, per sheet: for each
   (supplier_key, shipment month) present in the file, delete the existing
   rows with that key, then insert the file's rows. Rows without a valid date use
   the key (supplier_key, no date).
8. After the save, warn when any dated row falls before the first rate (it has no
   PACD or rebate). Saving rates is refused if it would leave a dated row without
   a rate.

Rows without a valid date are stored but left out of all quarter-based reports.

### Date check

The factories type dates day first, often on PCs set to month-first: Excel stores
`06/12/2025` as 12 June and keeps `18/12/2025` as text. The source files can't be
fixed, so the app warns and lets the user correct dates before import.

Each row is checked (today = import date, `LEAD_DAYS` = 60):

| Problem | Suggested date |
|---|---|
| Shipment Date empty or not a date (row left out of reports) | none |
| Date typed as text | read day first |
| Date in the future | day/month swapped, if that is not in the future |
| PO Received Date before 2020 (year lost, e.g. 1900) | same day/month in the shipment's year or the year before, if 0–60 days before shipment |
| PO Received Date not 0–60 days before Shipment Date | PO day/month swapped, if that fits |
| PO Received Date after Shipment Date | Shipment day/month swapped, if that fits; else none |

A long lead time alone (PO more than 60 days before shipment, no swap fits) is not
listed. Listed rows appear in an editable table (file values, problem, suggested
dates); the user changes any date and confirms the import.

Reviewed dates are saved in `date_fixes`, keyed by a hash of the row's input cells
(sheet + every column except the price formulas, with the date cells as in the
file). When the same row comes back in a later file it gets the saved dates and is
not listed again; a row whose cells changed is checked afresh. An expander lists
saved corrections and can forget them, so their rows are checked again at the next
import. `SCHEMA_VERSION` 2 adds the table to existing databases.

## Reports

Reports page inputs: suppliers (multi-select by key, shown with display names, all
selected by default), fiscal year, quarters (multi-select). One build makes the
Summary plus each selected supplier's reports in a folder named after the supplier
(suppliers with no rows in the selection are skipped). Downloads come as one zip
with that layout, with the single files in a collapsed list. "Save all" uses one
`<Month>` folder for the whole selection.
Each report has a download button. "Save all" writes reports 2–5 for the chosen
supplier into `Reports/<Month>/<Supplier>/` and the Summary into `Reports/<Month>/`,
where `<Month>` is the full month name of the latest shipment date in the selection.
Existing files with the same name are overwritten.

Rows are sorted by Shipment Date in every report. Rows without an eBiz SO# are
included.

### 1. Summary (xlsx)

- All packaging suppliers, chosen fiscal year, laid out as a print-ready report
  (A4 portrait, one page wide, no gridlines, footer with title, `Page x of y` and
  generated date).
- Title block: PacD logo left, Matalan logo right, title `Carton Shipment Summary`,
  fiscal year and shipment period, generated date/time.
- Overview: one line per supplier (its grand total) and an `All suppliers` total.
- Then one section per supplier: heading; header row `Quarter | Carton Type |
  PACD Rate | Rebate Rate | Carton QTY | Total Amount | PACD | Matalan Rebate`; for
  each quarter with data a Standard and a Custom line (quarter cell merged over
  both), a shaded quarter Total line; then a Grand Total line. The rate columns
  show the rate used (`Mixed` if the quarter's rows used more than one, `-` if the
  type has no rows). A section is moved to the next page rather than split.
- Notes at the end: quarter months and the PACD / rebate formulas.
- Standard and Custom values are rounded to 2 decimals. Total and Grand Total add
  the rounded values (same as the current Summary workbook). Zero shows as `-`.
- File: `Matalan Carton Shipment Tracking_Summary_<first month>-<last month> <yy>.xlsx`.

### 2. Full Tracker Sheet for PacD (PDF)

- One supplier, the selected quarters.
- A4 landscape. Left and right logos from the `Assets` sheet of the xlsm
  (extracted once into `assets/`). Title `Full Tracker Report (<supplier>)`.
- Standard block, then Custom block. Each block: banner, header row
  (repeated on every page), a quarter header row and a quarter subtotal row for each
  quarter, then `Standard Carton Total` / `Custom Carton Total`.
  Last row: `Tracker Grand Total`.
- Columns: Sr No., then source columns A–W (NOTE is not shown), with PACD and
  rebate calculated by the app.
- Formats: money `$ #,##0.00`; area 3 decimals; dimensions whole numbers.
- PACD and rebate headers show the % (for example `5% Rebates for Matalan`) when
  every selected quarter uses the same rate for that block; otherwise they show the
  plain name.
- Footer: left `Generated: dd-Mmm-yyyy hh:mm AM/PM`, right `Page x of y`.
- File: `Full Tracker Sheet for PacD_<supplier>_<Qa>-<Qb>_<year>.pdf`, or
  `_<Qn>_<year>.pdf` for one quarter.

### 3. Full Tracker Sheet for Matalan (PDF)

Same as report 2 without the PACD column. File name uses `Matalan` in place of `PacD`.

### 4. Tracker Raw Data (xlsx, one per quarter)

- One sheet named `Q<n> <year>`, with a `Standard Carton` block and a
  `Custom Carton` block. Each block has a banner row, the source header row and the
  data rows, with PACD and rebate as calculated values.
- Layout and colours follow the existing files in `Reports/August/*/Tracker Raw Data_*.xlsx`
  (checked during implementation).
- File: `Tracker Raw Data_<supplier>_Q<n>_<year>.xlsx`.

### 5. Shipment Tracking export (xlsx)

- One supplier, all its dated rows. Two sheets, `TRACK_STD` and `TRACK_CUS`, values
  only, header in row 1, PACD and rebate as calculated values.
- File: `Matalan Carton Shipment Tracking_NPS_BD <Mmm-yy> to <Mmm-yy> (<first word of supplier>).xlsx`,
  months from the first and last shipment dates.

Supplier names in file names and folders are the supplier's display name (the most
common Packaging Supplier spelling, for example `Transworld Packaging (BD) Ltd.`),
with characters that are not allowed in Windows file names removed. This differs
from the old tool, which used the truncated eBiz supplier name.

## UI (Streamlit, three tabs)

- Import: file uploader, date check table, preview, Confirm button, result message,
  saved date corrections.
- Rates: editable table of the `rates` table (add, edit, delete rows), Save button.
  Validation: carton_type is STD or CUS, date is valid, % between 0 and 1.
- Reports: inputs and buttons described above.

Start: `run.bat` runs `streamlit run app.py` from the project folder, bound to
localhost only.

## Files

| File | Purpose |
|---|---|
| `app.py` | Streamlit UI |
| `core.py` | Schema, seed rates, import, report queries |
| `reports.py` | PDF (reportlab) and xlsx (openpyxl) output |
| `check.py` | Assert-based self-check |
| `run.bat` | Starts the app |
| `assets/logo_left.png`, `assets/logo_right.png` | Report logos |

No new dependencies: Python 3.14, openpyxl, pandas, reportlab and streamlit are
already installed; SQLite is in the standard library.

## Error handling

- Import errors (missing sheet, missing header, no cached values) stop the import
  before any database change and show the reason.
- The import runs in one transaction; a failure rolls back.
- A report with no rows for the selection shows "No data for this selection" and
  writes no file.

## Testing (`check.py`)

Uses a temporary database and the demo file.

1. Import gives 88 STD rows and 118 CUS rows.
2. STD sums: QTY 28,001; Total 33,646.9152; PACD 1,278.5828; rebate 1,682.3458
   (equal to the file, because STD rates did not change).
3. CUS sums: QTY 41,529; Total 37,815.6551; rebate 1,890.7828 (5%);
   PACD 2,694.3654.
4. Quarter mapping: 2025-12-01 is Q1 2026, 2026-03-31 Q1, 2026-04-01 Q2,
   2026-08-31 Q3, 2026-09-01 Q4, 2026-11-30 Q4, 2026-12-01 Q1 2027.
5. Importing the same file again leaves 88 and 118 rows.
6. Each report builds without error and the PDF and xlsx totals equal the view sums.

When the real files arrive: import them and compare Q1 and Q2 totals with
`Reports/August/Matalan Carton Shipment Tracking_Summary_Dec-June 26.xlsx`
(Q1 and Q2 rates are unchanged, so they must match to the cent).

## Open items

- Re-import rule: confirm "replace overlap" once the real monthly or cumulative
  files are available.
- Logos (resolved): in the xlsm, `Assets` sheet drawing `LogoLeft` is
  `xl/media/image3.png` (458×154) and `LogoRight` is `xl/media/image4.png` (392×156).
