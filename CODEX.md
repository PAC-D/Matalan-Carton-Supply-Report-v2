# Matalan Carton Supply Report — development log

## Current architecture

- Responsive multipage Streamlit website.
- SQLite persistence (`CARTON_DB`, default `carton.db`).
- Excel/PDF report engine remains in `reports.py`.
- Docker deployment runs one application instance behind nginx at
  `/matalan-carton-report/`.

## Pages

1. Dashboard — live Summary-report KPIs, quarterly supplier totals, STD/CUS breakdown and Tracker-to-eBiz reconciliation.
2. Build reports — fiscal year/quarter filters and ZIP/single-file downloads.
3. Tracker data — editable `TRACK_STD` and `TRACK_CUS` database rows.
4. Rates — date-effective PACD and Matalan percentages.
5. Data import — tracker date review and all-vendor eBiz replacement.

## Important decisions

- Kept Streamlit because the Primark reference implementation uses Streamlit
  and the existing Python report engine is mature and tested.
- Local development requires no separate database server.
- Production currently targets one container with a persistent SQLite volume.
  Before horizontal scaling or high-concurrency editing, migrate persistence
  to PostgreSQL and add application authentication/roles.
- Tracker edits are transactional. Dates accept `YYYY-MM-DD` and
  `DD/MM/YYYY`; report-only fields are recalculated and cannot be edited.
- Dashboard and Summary Excel output consume the same `summary_rows` query so
  the live visualization stays consistent with generated report totals.

## Next deployment inputs

- Final hostname and whether the app is hosted at a domain root or URL prefix.
- Authentication method (nginx basic auth, SSO, or application accounts).
- Backup destination and retention policy for the Docker `carton-data` volume.
