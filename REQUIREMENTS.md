# Requirements — Matalan Carton Supply Report

| | |
|---|---|
| **Document version** | 3.3 |
| **Product version** | 3.0.0 (in development) |
| **Status** | In development |
| **Last updated** | 2026-10-22 |
| **Owner** | PAC-D |
| **Source** | `docs/superpowers/specs/2026-10-07-carton-tracker-design.md`, `docs/superpowers/specs/2026-10-08-ebiz-discrepancy-design.md` |

## 1. Purpose

Provide a responsive, hosted-capable web application that replaces the tracker-reporting and eBiz reconciliation parts of `eBiz Report Generator.xlsm`. The application imports and edits Matalan carton shipment tracker data in SQLite, applies date-effective rates, reconciles tracker Sales Orders against eBiz exports, and produces print-ready Excel and PDF reports.

## 2. Scope

### In scope

- Import of `TRACK_STD` and `TRACK_CUS` sheets from tracker workbooks.
- Date validation and remembered date corrections.
- Date-effective PACD and Matalan rebate rates.
- Supplier identity merging.
- Tracker reports: Summary, Full Tracker for PacD, Full Tracker for Matalan, Tracker Raw Data, and Shipment Tracking.
- eBiz export import with full replacement on each upload.
- Tracker-to-eBiz matching by Sales Order.
- Discrepancy Report PDF and Cross Reference workbook.
- Report bundle download and saving to a local `Reports/` folder.
- Local schema creation and automatic schema upgrades.
- Responsive multipage website with a dashboard and branded PACD/Matalan navigation.
- Manual editing of imported tracker rows with transactional validation.
- Single-instance Docker deployment with persistent database and report volumes.

### Out of scope

- Rebate Report and Matalan Rebate report (remaining phase 2 items).
- Migrating data automatically from the old `.xlsm` workbook.
- Application-managed login, user accounts, roles, or horizontal scaling.
- Scheduled jobs, emailing, or an API.
- Manual matching of tracker rows to eBiz orders.
- Automatic correction of source workbooks.

## 3. Users and operating environment

| Item | Requirement |
|---|---|
| Primary user | PAC-D operations users through a browser |
| Operating system | Windows 10/11 for local development; Linux/Docker for hosting |
| Python | 3.14 (tested with 3.14.6) |
| Runtime packages | `openpyxl` 3.1.5, `pandas` 3.0.5, `reportlab` 5.0.0, `streamlit` 1.61.1 |
| Test packages | `pypdf` 6.16.0, `pytest` 9.1.1 |
| Database | SQLite via the Python standard library; file `carton.db` beside the app |
| Network | Localhost during development; hosted behind an authenticated HTTPS reverse proxy in production |

## 4. Definitions

| Term | Meaning |
|---|---|
| Tracker workbook | An Excel file containing `TRACK_STD` and `TRACK_CUS` shipment sheets |
| `STD` / `CUS` | Standard and Custom carton types |
| Supplier key | Upper-cased first word of the Packaging Supplier name, with `. , ; : " ' ! ?` removed |
| Display name | Most common raw Packaging Supplier spelling for a supplier key |
| PACD | `(Total Amount − Matalan Rebate) × PACD rate` |
| Matalan rebate | `Unit Price × Carton QTY × Matalan rebate rate` |
| Tracker rebate | `PACD + Matalan rebate` (used in eBiz reconciliation) |
| Fiscal year | Calendar year of the Shipment Date, plus one when the month is December |
| Quarter | Q1 = Dec–Mar, Q2 = Apr–Jun, Q3 = Jul–Aug, Q4 = Sep–Nov |
| eBiz | Vendor shipment/rebate export used for the discrepancy comparison |
| SO# | eBiz Sales Order number from the tracker or vendor export |
| Dated row | A tracker row with a valid Shipment Date; only dated rows appear in quarter reports |

## 5. Functional requirements

### 5.1 Database and startup

| ID | Requirement |
|---|---|
| FR-DB-01 | On first start, the system shall create `carton.db` with the current schema (version 5) and seed the default rates. |
| FR-DB-02 | On start, the system shall upgrade an older database schema without deleting existing tracker, rate, date-correction, or eBiz data where those tables exist. |
| FR-DB-03 | The system shall store tracker data in `track_std` and `track_cus`, rates in `rates`, saved date corrections in `date_fixes`, and eBiz data in `ebiz`. |
| FR-DB-04 | The system shall expose rate-aware views `v_std` and `v_cus` with PACD, Matalan rebate, fiscal year, and quarter calculated at query time. |
| FR-DB-05 | The database path shall default to `carton.db` beside the app and may be overridden by the `CARTON_DB` environment variable. |

### 5.2 Tracker import

| ID | Requirement |
|---|---|
| FR-IMP-01 | The system shall accept one or more `.xlsx` or `.xlsm` tracker workbooks in a single import. |
| FR-IMP-02 | The system shall locate the header row within rows 1–10 by finding a `Shipment Date` cell. |
| FR-IMP-03 | The system shall require both `TRACK_STD` and `TRACK_CUS` sheets, matching sheet names case-insensitively, and shall report the missing sheet. |
| FR-IMP-04 | The system shall map columns by normalised header text and shall reject a file that lacks a required column, naming the missing column and header row. |
| FR-IMP-05 | The system shall read rows below the header and skip template rows where both Shipment Date and Carton QTY are blank. |
| FR-IMP-06 | The system shall trim text, treat Excel error strings such as `#N/A` as blank, and convert Excel dates and day-first text dates to ISO dates. |
| FR-IMP-07 | The system shall reject a sheet that has data rows but no cached `Total Amount` values, with guidance to open and save the file in Excel. |
| FR-IMP-08 | The system shall derive a supplier key for each row and keep the raw Packaging Supplier spelling. |
| FR-IMP-09 | The system shall preview rows per sheet, supplier key, and shipment month, the spellings merged into each key, and the number of database rows that will be replaced. |
| FR-IMP-10 | The system shall block an import when two files contain the same sheet, supplier key, and shipment month. |
| FR-IMP-11 | On confirmation, the system shall import all valid files in one transaction: for each sheet, supplier key, and shipment month, delete the existing group and insert the file's rows. |
| FR-IMP-12 | If any part of the import fails, the system shall roll back the whole batch and leave the database unchanged. |
| FR-IMP-13 | Rows without a valid Shipment Date shall be stored but excluded from all quarter-based reports and report filters. |
| FR-IMP-14 | After import, the system shall warn when dated rows exist before the first applicable rate. |

### 5.3 Date review

| ID | Requirement |
|---|---|
| FR-DATE-01 | The system shall list rows whose dates are empty, not dates, future dates, before 2020, or typed as text, with the file value, the problem, and a suggested date where one can be derived. |
| FR-DATE-02 | The system shall detect a likely day/month swap on a future date or on a PO Received Date outside the 0–60 day lead window, and suggest the swapped date when valid. |
| FR-DATE-03 | The system shall detect a PO Received Date that appears to have lost its year and suggest the same day/month in the shipment year or the year before when that fits the lead window. |
| FR-DATE-04 | The system shall detect a PO Received Date after the Shipment Date and suggest a shipment day/month swap when that fits. |
| FR-DATE-05 | The system shall present listed rows in an editable table so the user can change either date before confirming the import. |
| FR-DATE-06 | The system shall save reviewed dates against a hash of the row's input cells and apply them automatically when the same row returns in a later import. |
| FR-DATE-07 | The system shall list saved corrections and allow the user to forget selected corrections, after which those rows are checked again on the next import. |

### 5.3A Tracker editing

| ID | Requirement |
|---|---|
| FR-EDIT-01 | The Tracker data page shall allow the user to select `TRACK_STD` or `TRACK_CUS` and a packaging supplier. |
| FR-EDIT-02 | Every imported source column shall be editable; the database row ID, source file and last-change time shall be read-only. |
| FR-EDIT-03 | PO Received Date and Shipment Date shall accept `YYYY-MM-DD` and day-first `DD/MM/YYYY` manual input. |
| FR-EDIT-04 | Invalid dates or numeric values shall reject the complete save and leave every selected row unchanged. |
| FR-EDIT-05 | Changing a Packaging Supplier name shall recalculate its supplier key. Calculated rates, rebates, PACD, fiscal year and quarter shall update automatically through the database views. |
| FR-EDIT-06 | The editor shall paginate at no more than 250 rows per browser view. |

### 5.4 Rates

| ID | Requirement |
|---|---|
| FR-RATE-01 | The system shall store rates by carton type and effective date, with PACD and Matalan percentages as fractions. |
| FR-RATE-02 | A rate shall apply to a shipment from its effective date until the next rate of the same carton type; the applicable rate is the latest effective date on or before the Shipment Date. |
| FR-RATE-03 | The system shall seed STD 4% / 5% and CUS 7.5% / 10% from 2025-12-01, and STD 4% / 5% and CUS 7.5% / 5% from 2026-07-01. |
| FR-RATE-04 | The Rates page shall allow adding, editing, and deleting rate rows and saving the complete table. |
| FR-RATE-05 | The system shall reject invalid carton types, invalid dates, percentages outside 0–100, duplicate type/date pairs, and an empty rate table. |
| FR-RATE-06 | The system shall reject a rate save that would leave any dated shipment without a rate, and shall save nothing in that case. |
| FR-RATE-07 | Rate changes shall take effect in reports immediately because PACD and rebate are calculated from the current rates at query time. |

### 5.5 Tracker reports

| ID | Requirement |
|---|---|
| FR-REP-01 | The Build reports page shall allow selecting one or more suppliers, a fiscal year, and one or more quarters. |
| FR-REP-02 | The system shall build the Summary workbook for **all suppliers** in the chosen fiscal year, independent of the supplier selection. |
| FR-REP-03 | The Summary shall show a title block with logos, fiscal year, shipment period, generated date/time, an overview line per supplier, an all-suppliers total, per-supplier quarter sections with Standard and Custom lines, quarter subtotals, grand totals, notes, and page numbering. |
| FR-REP-04 | The Summary shall be laid out for A4 portrait, one page wide, without gridlines, and shall not split a supplier section across pages. |
| FR-REP-05 | The system shall build a Full Tracker PDF for PacD for each selected supplier with data in the selection, including logos, a Standard and Custom block, quarter subtotals, block totals, a grand total, repeated headers, and page numbering. |
| FR-REP-06 | The system shall build a Full Tracker PDF for Matalan with the same content and layout as the PacD version, without the PACD column. |
| FR-REP-07 | The Full Tracker PDFs shall show a rate in the PACD or rebate header when all selected rows share one rate, otherwise the plain column name. |
| FR-REP-08 | The system shall build one Tracker Raw Data workbook per selected quarter for each selected supplier, with separate Standard and Custom carton blocks. |
| FR-REP-09 | The system shall build a Shipment Tracking workbook for each selected supplier with `TRACK_STD` and `TRACK_CUS` sheets containing all dated rows for that supplier. |
| FR-REP-10 | Report rows shall be sorted by Shipment Date, and rows without an eBiz SO# shall still be included. |
| FR-REP-11 | A selection with no rows shall show “No data for this selection” and shall not write an empty report. |
| FR-REP-12 | Supplier names in filenames and folders shall use the display name with characters invalid in Windows filenames removed and trailing dots/spaces stripped. |

### 5.6 eBiz import

| ID | Requirement |
|---|---|
| FR-EBIZ-01 | The system shall accept multiple `.xlsx` or `.xlsm` eBiz export files in one upload, one file per vendor. |
| FR-EBIZ-02 | The system shall read the first sheet and locate the header row within rows 1–10 by finding an `Order Number` cell. |
| FR-EBIZ-03 | The system shall require Vendor, Order Number, Qty Shipped, and Total Rebate, and store Vendor Code, Product, Net Price per 1000, Gross Price per 1000, and Currency when present. |
| FR-EBIZ-04 | The system shall skip blank rows, rows with no Vendor (including total/footer lines), and rows with neither an Order Number nor a Qty Shipped. |
| FR-EBIZ-05 | The system shall split `Order Number` on `|` into `SO_1` and `SO_2`, trim the values, normalise whole-number floats to integer text, and treat blank and `0` as no SO#. |
| FR-EBIZ-06 | The system shall derive `vendor_key` with the same supplier key rule used by tracker rows and shall show vendor spellings in the preview. |
| FR-EBIZ-07 | The system shall preview rows, total Qty Shipped, and total rebate per file and vendor key before import. |
| FR-EBIZ-08 | The system shall block an upload when two files contain the same vendor key. |
| FR-EBIZ-09 | The system shall warn when an uploaded vendor has no tracker rows and when vendors in the current data are absent from the new upload. |
| FR-EBIZ-10 | On **Replace all eBiz data**, the system shall delete every existing eBiz row and insert all uploaded rows in one transaction; a failure shall leave the previous eBiz data intact. |
| FR-EBIZ-11 | The system shall show the current eBiz row count, source files, and import time, or “No eBiz data yet”. |

### 5.7 Reconciliation

| ID | Requirement |
|---|---|
| FR-XREF-01 | Reconciliation shall use every tracker row of the supplier, including undated rows, and every eBiz row whose vendor key matches the supplier. |
| FR-XREF-02 | The system shall group tracker rows by non-blank eBiz Sales Order number; rows without an SO# shall form one separate group. |
| FR-XREF-03 | Each eBiz row shall resolve to `SO_1` when `SO_1` is a tracker SO#, otherwise to `SO_2` when `SO_2` is a tracker SO# (flagged as matched via SO_2), otherwise to an eBiz-only group keyed by `SO_1` or `SO_2`. |
| FR-XREF-04 | For each tracker SO#, the system shall sum tracker quantity and tracker rebate (PACD + Matalan rebate, NULL counted as zero), and eBiz quantity and eBiz rebate. |
| FR-XREF-05 | Each SO# shall be classified as `Not in eBiz`, `Tracker > eBiz`, `Tracker < eBiz`, or `Match`. |
| FR-XREF-06 | eBiz orders that do not resolve to a tracker SO# shall be reported as eBiz-only with their per-order quantity and rebate totals. |
| FR-XREF-07 | Reconciliation results shall not be persisted; they shall be recalculated for each report build. |

### 5.8 eBiz reports

| ID | Requirement |
|---|---|
| FR-DISC-01 | The Build reports page shall provide a **Discrepancy report and cross reference** option, enabled only when eBiz data exists. |
| FR-DISC-02 | The eBiz reports shall be built for each selected supplier that has eBiz rows, regardless of the fiscal year and quarter filters. |
| FR-DISC-03 | The Discrepancy Report PDF shall use the Full Tracker visual style: landscape A4, title bar, logos, generated date, and page numbering. |
| FR-DISC-04 | The Discrepancy Report shall contain the five sections: tracker orders not found in eBiz; quantity mismatch tracker > eBiz; quantity mismatch tracker < eBiz; tracker rows with no SO#; eBiz orders not found in tracker. |
| FR-DISC-05 | Sections 1–4 shall show tracker rows in Standard and Custom blocks with quantity, amount, PACD, and rebate totals; sections 2–3 shall add the eBiz quantity for the SO#. |
| FR-DISC-06 | Section 5 shall list eBiz rows with vendor, vendor code, product, order number, quantity, net and gross price per 1000, total rebate, and currency, with a total line. |
| FR-DISC-07 | The Discrepancy Report shall highlight tracker rows matched via SO_2 and eBiz rows with blank SO_1 and a non-blank SO_2. |
| FR-DISC-08 | An empty section shall print “No discrepancies found.” |
| FR-DISC-09 | The Cross Reference workbook shall contain a tracker-orders sheet with SO#, tracker quantity and rebate, eBiz quantity and rebate, quantity difference, and status, sorted by SO#, followed by a no-SO# line where applicable and a total line. |
| FR-DISC-10 | The Cross Reference workbook shall contain a second sheet, “eBiz not in tracker”, with order, eBiz quantity, and eBiz rebate, plus a total line or “None”. |
| FR-DISC-11 | The Cross Reference shall use the Summary-style title block, navy headers, A4 portrait fit-to-width print setup, repeated print titles, no gridlines, and page numbering. |
| FR-DISC-12 | Dedicated Cross Reference and Discrepancy Report pages shall recalculate from the complete current Tracker and eBiz datasets whenever they are viewed. |
| FR-DISC-13 | The live Cross Reference shall show Tracker/eBiz quantities and rebates, quantity differences, match status, SO_2 matches, no-SO totals, and aggregated eBiz-only orders. |
| FR-DISC-14 | The live Discrepancy Report shall show the same five sections and row-level fields as the generated PDF and allow the current PDF to be prepared for download. |

### 5.9 Report delivery and UI

| ID | Requirement |
|---|---|
| FR-UI-01 | The UI shall provide grouped website navigation for Dashboard, Build reports, Cross reference, Discrepancy report, Rebate report, Tracker data, Rates, and Data import pages. |
| FR-UI-02 | The Data import page shall contain the tracker uploader, date review, preview, confirm action, result message, saved corrections, and eBiz upload section. |
| FR-UI-03 | The Rates page shall contain the editable rate table and save action with validation errors shown to the user. |
| FR-UI-04 | The Build reports page shall contain supplier, fiscal year, and quarter selectors; the discrepancy option; **Build reports**; **Save to Reports folder**; and per-file download buttons. |
| FR-UI-05 | **Build reports** shall produce a zip whose layout is Summary at the root and each supplier's files inside a folder named after the supplier. |
| FR-UI-06 | **Save all** shall write reports to `Reports/<Month>/` for the Summary and `Reports/<Month>/<Supplier>/` for supplier reports, overwriting files with the same name. |
| FR-UI-07 | `<Month>` shall be the full month name of the latest Shipment Date in the selection. |
| FR-UI-08 | The launcher `run.bat` shall create/reuse `.venv`, install pinned requirements and start Streamlit from the project folder on localhost. |
| FR-UI-09 | The system shall show actionable error messages for import errors, report errors, and missing data. |
| FR-UI-10 | The website shall use both supplied logos and remain usable on desktop, laptop, tablet, and phone widths. |
| FR-UI-11 | The Dashboard shall provide fiscal-year and packaging-supplier filters and show total cartons, Standard/Custom carton totals, quarterly packaging-supplier volumes, the detailed Summary rows, and Tracker-to-eBiz reconciliation. |
| FR-UI-12 | The repository shall include a non-root Docker image, persistent SQLite volume, report volume, health check, and nginx reverse-proxy template. |
| FR-UI-13 | Dashboard Summary values and the generated Summary Excel workbook shall use the same calculation source so their supplier, quarter, carton-type, rate and amount values remain consistent. |
| FR-UI-14 | The Reports navigation group shall include dedicated Cross Reference, Discrepancy Report and Rebate Report pages. |
| FR-UI-15 | The live Rebate Report shall filter by packaging supplier, fiscal year and quarter; show Standard and Custom detail rows and totals; and allow the current PDF to be prepared for download. |

## 6. Data requirements

### 6.1 Database schema (version 5)

| Object | Purpose |
|---|---|
| `track_std` | Standard carton tracker rows, 23 source columns plus supplier key, source file, and import time |
| `track_cus` | Custom carton tracker rows, same structure as `track_std` |
| `rates` | `carton_type`, `effective_from`, `pacd_pct`, `matalan_pct`; primary key `(carton_type, effective_from)` |
| `date_fixes` | Reviewed date corrections keyed by the row hash, with display fields and save time |
| `ebiz` | Vendor export rows, optional columns, `so_1`, `so_2`, `vendor_key`, source file, and import time |
| `v_std` / `v_cus` | Tracker rows joined to the applicable rate, with PACD, Matalan rebate, fiscal year, and quarter |

### 6.2 Calculation rules

- Matalan rebate = `unit_price × carton_qty × matalan_pct`
- PACD = `(total_amount − matalan_rebate) × pacd_pct`
- Tracker rebate for reconciliation = `pacd + matalan_rebate`
- A dated row with no applicable rate has NULL PACD and rebate; the app warns and refuses rate saves that would create this state.
- Fiscal year and quarter are derived from the Shipment Date:
  - Q1: December–March (December belongs to the next fiscal year)
  - Q2: April–June
  - Q3: July–August
  - Q4: September–November

### 6.3 Source data assumptions

- Tracker sheets contain a header row within rows 1–10 and data below it.
- Formula cells contain values cached by Excel on save; the app does not recalculate workbook formulas.
- Text dates are interpreted day-first.
- The tracker may contain template rows with formulas but no Shipment Date or Carton QTY; these are skipped.
- eBiz exports contain one vendor each and may have a header row within rows 1–10.
- eBiz `Order Number` may contain one SO or two SOs separated by `|`.

### 6.4 Retention and backup

- The database persists between sessions.
- Source file names and import timestamps are retained on tracker and eBiz rows for traceability.
- The user is responsible for backing up `carton.db`; generated files under `Reports/` can be rebuilt from the database.

## 7. Non-functional requirements

| ID | Requirement |
|---|---|
| NFR-01 | The same application shall run locally for development and as one persistent hosted container. |
| NFR-02 | Local Streamlit shall bind to localhost; the container shall bind internally to `0.0.0.0` and be published only to the host loopback interface for reverse proxying. |
| NFR-03 | Tracker and eBiz data shall remain in the configured SQLite database; the app shall not send data to an external application service. |
| NFR-04 | Tracker imports and eBiz replacement shall be atomic: a failure shall not leave a partially updated database. |
| NFR-05 | Schema upgrades shall preserve existing data and run automatically at startup. |
| NFR-06 | Reports shall be generated on demand from current database values; rate changes shall not require re-importing tracker rows. |
| NFR-07 | Generated filenames shall be safe on Windows and shall identify the supplier and period where applicable. |
| NFR-08 | Excel reports shall be print-ready (page setup, margins, headers/footers, repeated titles) and PDFs shall include page numbering and generation timestamps. |
| NFR-09 | The implementation shall use SQLite from the Python standard library and shall not require a database server or ORM. |
| NFR-10 | Runtime and test dependencies shall be pinned in `requirements.txt`. |
| NFR-11 | The automated test suite shall cover database logic, tracker import, date review, rates, live Summary calculations, report generation, eBiz import, reconciliation, tracker editing, and UI smoke tests. |
| NFR-12 | The application shall run on Windows 10/11 with Python 3.14 and the pinned dependencies. |
| NFR-13 | File and folder creation shall not overwrite unrelated files; only same-named generated reports are replaced. |
| NFR-14 | The project shall be distributed under the Mozilla Public License 2.0. |
| NFR-15 | Primary workflows and navigation shall be responsive at desktop, tablet and phone widths. |
| NFR-16 | A production deployment shall terminate TLS and enforce authentication before exposing the app to the internet. |
| NFR-17 | The SQLite deployment shall run as one application instance; horizontal scaling requires a database migration. |

## 8. Acceptance criteria

| ID | Criterion |
|---|---|
| AC-01 | Importing the demo workbook yields 88 `TRACK_STD` rows and 118 `TRACK_CUS` rows. |
| AC-02 | Demo totals match the specification: STD QTY 28,001; Total 33,646.9152; PACD 1,278.5828; rebate 1,682.3458. CUS QTY 41,529; Total 37,815.6551; rebate 1,890.7828 (5%); PACD 2,694.3654. |
| AC-03 | Re-importing the same workbook leaves 88 and 118 rows, with no duplicates. |
| AC-04 | Quarter mapping is correct for December 2025 through December 2026, including December belonging to the next fiscal year. |
| AC-05 | Every report builder produces a file whose totals equal the corresponding database view sums. |
| AC-06 | Date review suggests swapped dates, remembers reviewed values, reapplies them on re-import, and rechecks changed rows. |
| AC-07 | A failed multi-file import leaves the database unchanged; overlapping files are blocked. |
| AC-08 | An eBiz upload replaces all existing eBiz rows, and a failed batch leaves the previous data intact. |
| AC-09 | Reconciliation covers SO_1 match, SO_2 fallback, eBiz-only, Not in eBiz, tracker > eBiz, tracker < eBiz, Match, no-SO# rows, and undated tracker rows. |
| AC-10 | The Discrepancy Report contains every section title and the empty-section message; the Cross Reference values and totals match the reconciliation structure. |
| AC-11 | `python -m pytest check.py -q` passes all automated checks. |
| AC-12 | The app starts on localhost and all eight Dashboard, report, Tracker, Rates, and Import pages load without errors. |
| AC-13 | A user can edit a PO, date and quantity in Tracker data; valid changes persist and an invalid row rolls back the complete save. |
| AC-14 | The running app returns HTTP 200 from the Streamlit health endpoint. |
| AC-15 | Upgrading a version-4 database removes obsolete assignment-management tables without deleting tracker, rate, date-correction, or eBiz data. |
| AC-16 | Dashboard Summary rows equal the supplier/quarter/carton-type values used by the generated Summary workbook. |
| AC-17 | The Cross Reference, Discrepancy Report and Rebate Report pages load without error and display current database values for the selected supplier and period. |

## 9. Assumptions and constraints

- The user has Excel available to open and save source workbooks so formula values are cached.
- The tracker is the source of truth for carton quantities and prices; eBiz is the comparison source.
- eBiz exports have no shipment dates, so reconciliation is not date-filtered.
- Supplier merging by first word is sufficient for the current data; two suppliers sharing a first word would merge.
- Rates are maintained by the user; the seeded Q3 2026 Custom rebate of 5% intentionally differs from older workbooks that used 10%.
- Report layout is tuned to the current tracker columns A–X and A4 paper sizes.
- The app does not recalculate prices or other workbook formulas.

## 10. Traceability

| Area | Primary implementation | Primary checks |
|---|---|---|
| Schema, views, rates | `core.py` | `test_quarter_mapping`, `test_rate_follows_shipment_date`, `test_seeded_rates`, `test_rate_change_applies_to_existing_rows` |
| Tracker import | `core.py` | `test_demo_import_counts_and_totals`, `test_reimport_replaces_only_overlapping_supplier_months`, `test_failed_batch_leaves_database_unchanged` |
| Date review | `core.py`, `pages/5_Import.py` | `test_dates_that_look_right_are_not_listed`, `test_reviewed_dates_are_imported_and_reused_next_time`, `test_forgotten_correction_is_checked_again` |
| Tracker editing | `core.py`, `pages/3_Tracker.py` | `test_tracker_rows_can_be_manually_edited`, `test_invalid_tracker_edit_rolls_back` |
| Live Summary dashboard | `core.py`, `pages/1_Dashboard.py`, `reports.py` | `test_live_summary_rows_match_report_totals`, `test_cross_reference_overview_drives_live_dashboard`, `test_app_smoke` |
| Live report pages | `core.py`, `pages/6_Cross_Reference.py`, `pages/7_Discrepancy_Report.py`, `pages/8_Rebate_Report.py` | `test_live_report_pages`, `test_cross_reference`, `test_rebate_report_pdf` |
| Tracker reports | `reports.py` | `test_summary_xlsx`, `test_full_tracker_pdf_pacd`, `test_raw_data_xlsx`, `test_shipment_tracking_xlsx`, `test_save_all` |
| eBiz import | `core.py`, `pages/5_Import.py` | `test_read_ebiz_splits_order_numbers`, `test_ebiz_upload_replaces_all_ebiz_data`, `test_app_replaces_ebiz_data` |
| Reconciliation | `core.py` | `test_cross_reference` |
| eBiz reports | `reports.py` | `test_discrepancy_pdf`, `test_cross_reference_xlsx`, `test_build_all_adds_discrepancy_files_for_suppliers_with_ebiz` |
| Website UI | `app.py`, `web_ui.py`, `pages/` | `test_app_smoke`, `test_app_builds_reports_for_every_selected_supplier`, `test_app_no_data_for_selection` |

## 11. Document history

| Version | Date | Change |
|---|---|---|
| 3.3 | 2026-10-22 | Added dedicated live Cross Reference, Discrepancy Report and Rebate Report pages with matching report downloads |
| 3.2 | 2026-10-22 | Refocused the website on the editable Tracker and a live dashboard matching the generated Summary and reconciliation reports; removed assignment management |
| 3.1 | 2026-10-21 | Packaging-supplier catalog and factory-assignment management requirements |
| 3.0 | 2026-10-20 | Responsive hosted website, dashboard, multipage navigation, editable tracker and Docker deployment requirements |
| 2.0 | 2026-10-19 | Requirements for the released v2.0.0, including eBiz reconciliation, reports, and acceptance criteria |
| 1.0 | 2026-10-07 | Initial tracker, rates, date review, and tracker report requirements (source specification) |
