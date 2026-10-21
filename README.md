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
- Maintains date-effective PACD and Matalan rebate rates.
- Imports all current eBiz vendor exports as one atomic replacement.
- Creates Summary, Full Tracker, Raw Data, Shipment Tracking, Discrepancy and
  Cross Reference reports.
- Stores application data in SQLite. A fresh database is created and seeded on
  first start.

## Requirements

- Python 3.14+
- Packages in `requirements.txt`
- Windows 10/11 for `run.bat`; Linux is supported through the manual or Docker
  commands
- No external database or API keys for the current single-instance setup

## Run locally

### Windows one-click

Double-click `run.bat`. On the first run it creates `.venv`, installs the
requirements and opens the website at <http://localhost:8501>. The launcher
always uses the existing `carton.db` in this project folder, so previously
imported tracker data is available without importing the workbook again.

### Manual setup

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m streamlit run app.py --server.address localhost --server.port 8501
```

Open <http://localhost:8501>.

The first start creates `carton.db` next to `app.py`. To keep development data
elsewhere:

```powershell
$env:CARTON_DB = "$PWD\data\carton-dev.db"
python -m streamlit run app.py
```

## Run with Docker

Create the environment file once:

```powershell
Copy-Item .env.example .env
docker compose up -d --build
```

Open <http://localhost:8501/matalan-carton-report/>.

Docker stores SQLite data in the named `carton-data` volume and writes saved
reports to the host `Reports/` folder. Do not scale this configuration to
multiple containers; SQLite is intended here for one application instance.

The database is excluded from Git and the Docker image. A new Docker volume
therefore starts empty. To carry the current local data to a new server, start
the container once, stop the service, copy `carton.db` into its data volume,
fix ownership, and start it again:

```bash
docker compose up -d --build
docker compose stop matalan-carton-report
docker cp ./carton.db matalan-carton-report:/data/carton.db
docker compose run --rm --user root matalan-carton-report chown 1000:1000 /data/carton.db
docker compose start matalan-carton-report
```

Alternatively, leave the new volume empty and import the tracker and eBiz
workbooks through the **Data import** page.

For production, put nginx or another TLS reverse proxy in front of the
container. A route template is provided at
`deploy/nginx-matalan-carton-report.conf`. It expects an nginx basic-auth file
at `/etc/nginx/.htpasswd-matalan-carton-report`; create that file before
enabling the route. You can replace basic auth with platform SSO if preferred.

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

## Database and backups

| Setting | Purpose | Default |
|---|---|---|
| `CARTON_DB` | SQLite database path | `carton.db` beside `app.py` |
| `REPORTS_DIR` | Server-side saved-report directory | `Reports/` beside the app |

Back up the database while the app is stopped, or use SQLite's online backup
API. `carton.db`, `.env`, `data/` and `Reports/` are excluded from Git.

The SQLite file contains parsed tracker rows, parsed eBiz rows, saved date
corrections and rebate rates. It records source workbook names and import times,
but it does not embed copies of the original Excel workbooks. Keep those
workbooks separately if the source files must also be retained.

For multiple app containers or sustained concurrent editing, migrate the
persistence layer to PostgreSQL before scaling.

## Testing

```powershell
python -m pytest check.py -q
```

The suite covers schema/rates, tracker and eBiz imports, date correction,
manual tracker editing, report generation and page smoke tests.

## Project structure

```text
app.py              Website navigation
pages/              Dashboard, reports, tracker editor, rates and imports
web_ui.py           Responsive PACD/Matalan theme and shared page chrome
core.py             SQLite schema, imports, edits, queries and reconciliation
reports.py          Excel/PDF report builders
assets/             PACD and Matalan logos
.streamlit/         Website theme and upload configuration
Dockerfile          Hosted application image
docker-compose.yml  Persistent single-instance deployment
deploy/             Reverse-proxy configuration
CODEX.md            Architecture decisions and future deployment inputs
check.py            Automated test suite
```

## License

Mozilla Public License 2.0. See `LICENSE`.
