# Changelog

All notable changes to the Matalan Carton Supply Report are documented here.

The format follows [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and this project uses [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

- Responsive multipage website with grouped navigation for Dashboard, report building/live views, Tracker data, Rates, and Data import.
- Branded PACD/Matalan page shell using the existing left and right logo assets.
- Dashboard with live Summary-report KPIs, quarterly packaging-supplier volumes, Standard/Custom breakdown and Tracker-to-eBiz reconciliation.
- Dedicated live Cross Reference, Discrepancy Report and Rebate Report pages under Reports, including report-aligned filters, totals, detail tables and downloads.
- Editable tracker page for all imported `TRACK_STD` and `TRACK_CUS` source fields, with supplier filters and 250-row pagination.
- Transactional tracker-edit validation, including manual `YYYY-MM-DD` and `DD/MM/YYYY` dates and recalculation of supplier keys.
- Dockerfile, Docker Compose persistence, health check, and nginx reverse-proxy template for hosted deployment.
- `CODEX.md` architecture and deployment handoff log.

### Changed

- `app.py` is now the website router; feature UI is separated into page modules.
- Dashboard spacing, KPI density and chart sizing were redesigned for a compact modern layout without overlapping graph content.
- Dashboard filters now select fiscal year and packaging suppliers; all dashboard Summary values use the same aggregation source as the generated Summary Excel workbook. `assets/favicon.png` is the browser tab icon.
- The fixed brand bar now includes an explicit responsive layout spacer, uses a larger Matalan logo, and cannot overlap dashboard content; the footer displays the current year and recurSID attribution dynamically.
- The attribution is now a single compact chip fixed to the bottom-right of the viewport. It is rendered with the page header rather than as a footer section, reserves no document space, and page content scrolls behind it.
- `run.bat` now creates a virtual environment and installs requirements automatically before launch.
- Local and hosted report/database paths can be configured with `REPORTS_DIR` and `CARTON_DB`.
- SQLite schema v5 removes the out-of-scope assignment-management tables while preserving tracker, rate, date-correction and eBiz data.
- Test coverage includes schema migration, shared live-Summary calculations, reconciliation totals and page smoke coverage.

### Verification

- 89 automated checks pass (`python -m pytest check.py -q`).
- All eight website entry points load without exceptions.
- A live local server returned HTTP 200 and `ok` from `/_stcore/health`.

## [2.0.0] - 2026-10-19

The eBiz reconciliation release. Phase 2 features were completed on 2026-10-08 and published as v2.0.0 with the project documentation.

### Added

- eBiz export import for `.xlsx` and `.xlsm` files, with header detection in rows 1–10, required-column validation, optional columns, and vendor spellings merged into vendor keys.
- Sales Order splitting for `Order Number` values such as `7124561|7124562`, into `SO_1` and `SO_2`, with numeric values normalised to text.
- eBiz upload preview with rows, quantity, and total rebate per vendor, plus duplicate-vendor detection across files.
- Atomic **Replace all eBiz data** action: every existing eBiz row is deleted and the uploaded files are inserted in one transaction.
- Cross-reference matching of tracker rows against eBiz per Sales Order, including SO_1 matches, SO_2 fallback matches, eBiz-only orders, tracker rows without an SO#, and the statuses `Not in eBiz`, `Tracker > eBiz`, `Tracker < eBiz`, and `Match`.
- Discrepancy Report PDF with five sections: tracker orders not found in eBiz, tracker > eBiz, tracker < eBiz, tracker rows with no SO#, and eBiz orders not found in tracker.
- Cross Reference workbook with a tracker-vs-eBiz sheet per SO# and a separate sheet for eBiz orders missing from the tracker.
- Reports tab option to include the Discrepancy Report and Cross Reference in the build, zip bundle, and **Save all** output.
- Database schema version 3, adding the `ebiz` table while preserving existing tracker, rate, and date-correction data.
- Project documentation: `README.md`, `REQUIREMENTS.md`, `RELEASES.md`, `CHANGELOG.md`, and pinned dependencies in `requirements.txt`.

### Changed

- eBiz comparison deliberately ignores the fiscal-year and quarter filters because eBiz rows have no dates; all tracker rows for each selected supplier are compared.
- Discrepancy and Cross Reference reports are generated only for selected suppliers that have eBiz rows.
- Reports layout was adjusted for real tracker data.

### Fixed

- Vendor-less eBiz total and footer rows are skipped instead of being imported as data.
- Discrepancy Report and Cross Reference pagination, section bars, and table layout on real data.
- Discrepancy Report sections with no findings now print “No discrepancies found.”

### Verification

- 78 automated checks pass (`python -m pytest check.py -q`).

## [1.0.0] - 2026-10-07

The initial tracker reporting release.

### Added

- SQLite storage for `TRACK_STD` and `TRACK_CUS` tracker rows, with schema versioning and automatic creation of `carton.db`.
- Excel import with header-row detection, data cleaning, Excel error handling, date conversion, template-row skipping, and a cached-value guard.
- Multi-file tracker import with a preview of rows per sheet, supplier, and shipment month, plus replacement counts.
- Duplicate-file detection for the same sheet, supplier, and shipment month.
- Transactional replace-overlap import: for each supplier and shipment month, existing rows are replaced by the imported rows.
- Supplier identity rule that merges spelling variants by the upper-cased first word of the packaging supplier name.
- Date sanity checks for future dates, dates before 2020, text dates, day/month swaps, lost years, and PO dates outside the expected lead time.
- Editable suggested date corrections, saved in the database and reused automatically when the same rows are imported again, with an option to forget corrections.
- Date-effective PACD and Matalan rebate rates per carton type, editable on the Rates tab.
- Validation that rate rows use `STD`/`CUS`, valid dates, percentages between 0 and 100, unique effective dates, and that every dated shipment keeps a rate.
- Rate-aware views `v_std` and `v_cus` that calculate PACD, Matalan rebate, fiscal year, and quarter.
- Summary workbook with logos, supplier overview, quarter-by-carton-type sections, subtotals, grand totals, notes, and print setup.
- Full Tracker PDFs for PacD and Matalan with logos, repeated headers, quarter subtotals, grand totals, and page numbering.
- Tracker Raw Data workbooks, one per selected quarter, with Standard and Custom carton blocks.
- Shipment Tracking export with `TRACK_STD` and `TRACK_CUS` sheets for all dated rows of a supplier.
- Report zip bundle and **Save all** to `Reports/<Month>/<Supplier>/`.
- Windows launcher `run.bat`, bound to localhost.

### Changed

- Summary workbook changed from a basic table to a professional, print-ready report layout.
- Import supports several files at once, and reports can be built for several suppliers at once.

### Fixed

- Header row is found within rows 1–10 instead of assuming row 2.
- Supplier spelling variants merge correctly on import and re-import.
- Downloads are cleared when inputs change, preventing stale report files.
- Selections with no data show a clear message instead of writing empty reports.
- Missing rates are guarded on import and when saving rates.
- Startup is hardened for schema upgrades and existing databases.
