# Matalan Carton Supply Report

A responsive PACD/Matalan website for importing carton shipment trackers,
correcting supplier data, maintaining rebate rates, reconciling eBiz exports,
and creating print-ready Excel and PDF reports.

The application is implemented with Streamlit, following the architecture and
working pattern of the Primark Packaging Supplier Allocation website. It runs
locally for development and as a persistent Docker service behind a web server
for hosting.

## Website pages

| Section | Page | Purpose |
|---|---|---|
| Overview | **Dashboard** | Live Summary-report totals, quarterly supplier volumes, STD/CUS breakdown, and Tracker-to-eBiz reconciliation |
| Reports | **Build reports** | Supplier/fiscal-period filters and ZIP or individual report downloads |
| Reports | **Cross reference** | Live order-level Tracker/eBiz quantities, rebates, differences, match status, and eBiz-only orders |
| Reports | **Discrepancy report** | Live versions of all five discrepancy sections with matching PDF and Excel downloads |
| Reports | **Rebate report** | Live Standard/Custom carton details and rebate totals, for all tracker rows or only rows with an eBiz SO#, with PDF and Excel downloads |
| Manage | **Tracker data** | Edit imported `TRACK_STD` and `TRACK_CUS` rows, including dates, POs, quantities and prices |
| Manage | **Rates** | Maintain date-effective PACD and Matalan rebate percentages |
| Manage | **Data import** | Import tracker workbooks, review suspicious dates and replace eBiz vendor data |

The layout is responsive for desktop, laptop, tablet and phone. The PACD logo
appears in navigation and the Matalan wordmark appears in the top bar.

## Core behavior

- Imports `TRACK_STD` and `TRACK_CUS` from one or more `.xlsx`/`.xlsm` files.
- Reviews suspicious dates before import and remembers confirmed corrections.
- Provides a database-backed tracker editor after import. Dates accept
  `YYYY-MM-DD` or `DD/MM/YYYY`; every edited row is validated and saved in one
  transaction.
- Shows a live dashboard calculated from the same supplier, quarter, rate and
  carton-type rows used to generate the Summary Excel workbook.
- Shows Cross Reference, Discrepancy Report and Rebate Report data as dedicated
  live website pages, with the same calculations used by their downloads.
- Maintains date-effective PACD and Matalan rebate rates.
- Imports all current eBiz vendor exports as one atomic replacement.
- Creates Summary, Full Tracker, Raw Data, Shipment Tracking, Discrepancy and
  Cross Reference reports. Every PDF report also comes as an Excel workbook
  with the same rows and totals.
- Stores application data in PostgreSQL (`DATABASE_URL`). Tables and views are
  created and rates seeded on first connect.

## Requirements

- Python 3.14+
- Packages in `requirements.txt`
- PostgreSQL 16+ (local development uses `postgres:18` in Docker Desktop)
- Windows 10/11 for `run.bat`; Linux can use `run.sh` or the manual/Docker
  commands

## Run locally

### Windows one-click

Start a local PostgreSQL once (Docker Desktop must be running):

```powershell
docker run -d --name carton-pg -e POSTGRES_PASSWORD=postgres -p 5432:5432 postgres:18
```

After a reboot, `docker start carton-pg` brings it back with its data.

Double-click `run.bat`. On the first run it creates `.venv`, installs the
requirements and opens the website at <http://localhost:8501>. Unless
`DATABASE_URL` is already set, it uses the local `carton-pg` database.

### Linux one-click

Start PostgreSQL using the Docker command above, then run `./run.sh`. On the
first run it creates `.venv`, installs the requirements and starts the website
at <http://localhost:8501>. Unless `DATABASE_URL` is already set, it uses the
local `carton-pg` database.

### Manual setup

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
$env:DATABASE_URL = "postgresql://postgres:postgres@localhost:5432/postgres"
python -m streamlit run app.py --server.address localhost --server.port 8501
```

Open <http://localhost:8501>.

## Production (Lightsail + RDS)

Production runs at <https://matalan.pac-d.com/carton-supply-report/> on the
Lightsail server `ubuntu_pacd`, with its data in the RDS PostgreSQL instance
`database-1` (database `matalan_carton`, login `matalan_carton_app`).

On the server, the app lives in `/opt/matalan-carton`:

- `.env` (mode 600, never committed) holds `DATABASE_URL`; see `.env.example`.
- `certs/global-bundle.pem` is the RDS CA bundle used for `sslmode=verify-full`:
  `curl -fsSo certs/global-bundle.pem https://truststore.pki.rds.amazonaws.com/global/global-bundle.pem`
- The container is published on `127.0.0.1:8502` only (8501 is the Primark
  allocation app). Host nginx serves it over
  HTTPS with basic auth, using the locations in
  `deploy/nginx-matalan-carton-report.conf` and the password file
  `/etc/nginx/.htpasswd-carton-supply-report`.

Deploy an update:

```bash
cd /opt/matalan-carton
git pull
sudo docker compose up -d --build
```

Saved reports are written to `/opt/matalan-carton/Reports/`.

## Data workflow

### Import tracker data

1. Open **Data import**.
2. Upload one or more tracker workbooks.
3. Review any flagged dates and edit the proposed values.
4. Review the file/supplier/month preview.
5. Select **Confirm tracker import**.

Existing rows for the same sheet, supplier and shipment month are replaced.
Rows without a valid Shipment Date remain stored but do not appear in
quarter-based reports.

### Correct imported tracker rows

1. Open **Tracker data**.
2. Select `TRACK_STD` or `TRACK_CUS` and a packaging supplier.
3. Edit cells directly in the table.
4. Select **Save tracker changes**.

Row IDs and source metadata are read-only. Rate, rebate, PACD, fiscal year and
quarter are calculated fields and update automatically after edits.

### Import eBiz data

Upload one export per vendor under **Data import**. Replacement deletes all
current eBiz rows and inserts the uploaded files in one transaction, so every
current vendor file must be uploaded together.

### Build reports

Open **Build reports**, select suppliers, fiscal year and quarters, then build
a downloadable ZIP or save files under `Reports/<Month>/`.

### Review live reports

- Open **Cross reference** and select a packaging supplier to review every
  Tracker order against eBiz, including no-SO and eBiz-only totals.
- Open **Discrepancy report** to review its five exception sections and prepare
  the corresponding PDF and Excel files.
- Open **Rebate report** to filter by supplier, fiscal year and quarter, choose
  **All tracker rows** or **Rows with eBiz SO#** (leaves out rows whose eBiz
  Sales Order # is blank or 0), review Standard and Custom calculations, and
  prepare the corresponding PDF and Excel files.

## Database and backups

| Setting | Purpose | Default |
|---|---|---|
| `DATABASE_URL` | PostgreSQL connection URL | none; required (`run.bat` sets a local default) |
| `REPORTS_DIR` | Server-side saved-report directory | `Reports/` beside the app |

Production data is backed up by RDS automated backups (7-day retention on
`database-1`). `.env`, `certs/` and `Reports/` are excluded from Git.

The database contains parsed tracker rows, parsed eBiz rows, saved date
corrections and rebate rates. It records source workbook names and import times,
but it does not embed copies of the original Excel workbooks. Keep those
workbooks separately if the source files must also be retained.

Each page run opens one connection and closes it when the page finishes. Add
connection pooling (`psycopg_pool`) and application roles before scaling to
many concurrent users.

## Design decisions

- Streamlit, because the Primark reference implementation uses it and the
  Python report engine is mature and tested.
- Tracker edits are transactional; report-only fields are recalculated by the
  database views and cannot be edited.
- The Dashboard and Summary workbook share `core.summary_rows`; the live Cross
  Reference, Discrepancy and Rebate pages share their aggregations and views
  with the generated reports, so live and printed totals always agree.

## Open deployment items

- The `database-1` security group allows port 5432 from `0.0.0.0/0` for the
  Primark app; access is protected only by password and TLS. Narrow it when the
  Primark app's network path is known.
- Authentication is nginx basic auth; SSO or application accounts are not set up.

## Testing

```powershell
python -m pytest check.py -q
```

The suite needs the local `carton-pg` database (or set `TEST_DATABASE_URL`).
It drops and recreates the `public` schema before every test, so never point it
at production. It covers schema/rates, tracker and eBiz imports, date
correction, manual tracker editing, report generation, page smoke tests, and
that pages release their database connections.

## Project structure

```text
app.py              Website navigation
pages/              Dashboard, report views/builders, tracker editor, rates and imports
web_ui.py           Responsive PACD/Matalan theme and shared page chrome
core.py             PostgreSQL schema, imports, edits, queries and reconciliation
reports.py          Excel/PDF report builders
assets/             PACD and Matalan logos
.streamlit/         Website theme and upload configuration
Dockerfile          Hosted application image
docker-compose.yml  Single-instance deployment (reads .env, mounts certs/ and Reports/)
deploy/             Reverse-proxy configuration
check.py            Automated test suite
```

## License

Mozilla Public License 2.0. See `LICENSE`.
