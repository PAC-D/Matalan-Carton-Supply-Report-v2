# Cloud Deployment (Lightsail + RDS PostgreSQL) Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Serve the app at `https://matalan.pac-d.com/carton-supply-report/` with its data in RDS PostgreSQL `database-1`.

**Architecture:** `core.py` swaps `sqlite3` for psycopg 3, keeping every public function name and signature. Tests run against a local `postgres:18` container. The existing Docker image runs on Lightsail `ubuntu_pacd` behind host nginx (TLS + basic auth) and connects to RDS over verified TLS.

**Tech Stack:** Python 3.14, Streamlit 1.61.1, psycopg 3 (`psycopg[binary]`), PostgreSQL 18, Docker Compose, nginx, certbot.

**Spec:** `docs/superpowers/specs/2026-11-01-cloud-deployment-design.md`

## Global Constraints

- PostgreSQL only; no SQLite code path remains (`import sqlite3` gone from `core.py`).
- One config variable: `DATABASE_URL`. It replaces `CARTON_DB` everywhere (code, tests, Dockerfile, compose, `.env.example`, `run.bat`, README).
- Tests read `TEST_DATABASE_URL`, default `postgresql://postgres:postgres@localhost:5432/postgres`; unreachable DB = test failure, never skip.
- URL prefix: `carton-supply-report` (Streamlit `--server.baseUrlPath`, Docker healthcheck, nginx location).
- Production `DATABASE_URL` form: `postgresql://matalan_carton_app:<password>@database-1.ch4qkgwmwuqn.ap-south-1.rds.amazonaws.com:5432/matalan_carton?sslmode=verify-full&sslrootcert=/app/certs/global-bundle.pem`.
- Dates stay ISO `TEXT` columns; Python callers keep receiving `str` dates.
- SQL text with bound parameters must not contain a literal `%` other than placeholders (psycopg treats `%` as a placeholder marker).
- No secret (DB master password, app password, basic-auth password) is typed by Claude or printed in chat or committed. `.env` and `certs/` stay git-ignored.
- Do not modify security group `sg-0abf806525ce33662` or anything belonging to the Primark app.

## Review Focus

1. **Unreachable or misconfigured database** (RDS down, wrong password, missing `DATABASE_URL`): a page should show the existing "could not be loaded" panel, not a blank screen or raw traceback. Test in Task 3.
2. **`Decimal` or numeric leaking from PostgreSQL** (`EXTRACT`, `SUM` over the wrong type): mixing `Decimal` with `float` raises `TypeError` in totals and reports. Types pinned in Task 1.
3. **Connection leak per Streamlit rerun**: `db.t4g.micro` allows roughly 80 connections and the Primark app shares them; page runs must not leave connections open. Test in Task 3.
4. **Rollback on failed writes**: a failed rate save, tracker edit or import batch must leave the database unchanged. The existing rollback tests carry over in Task 2 and Task 3 and must pass unchanged.
5. **Supplier names with quotes or `%`** (`M&U`, apostrophes): values travel only as bound parameters, never interpolated into SQL text. Test in Task 2.

---

### Task 1: PostgreSQL connection, schema and test harness

**Files:**
- Modify: `requirements.txt` (add `psycopg[binary]` under Runtime, pinned to the newest psycopg 3 release on PyPI that ships a Python 3.14 binary wheel)
- Modify: `core.py:1-160` (imports, `VIEW_SQL`, `_schema`, `connect`)
- Modify: `check.py:1-135` (harness, `insert`, schema tests)

**Interfaces:**
- Produces: `core.connect(url: str | None = None) -> psycopg.Connection` (autocommit; url argument, else env `DATABASE_URL`; raises `RuntimeError("DATABASE_URL is not set.")` when neither exists). Rows support `r["col"]`, `r[0]`, `dict(r)`, `tuple(r)`.
- Produces: `core._ready: set[str]` — URLs whose schema was ensured in this process; tests clear it.
- Produces: check.py autouse fixture `pg` that resets the test database and sets `DATABASE_URL`; tests call `core.connect()` with no argument.
- Write blocks everywhere use `with con.transaction():`; bulk inserts use `con.cursor().executemany(sql, rows)`.

- [ ] **Step 1: Start local PostgreSQL**

Run: `docker run -d --name carton-pg -e POSTGRES_PASSWORD=postgres -p 5432:5432 postgres:18`
Expected: container id printed; `docker exec carton-pg pg_isready` prints `accepting connections`.

- [ ] **Step 2: Rewrite the test harness in `check.py`**

- Add an autouse fixture `pg(monkeypatch)`: connect to `TEST_DATABASE_URL` (default above) with autocommit, run `DROP SCHEMA public CASCADE; CREATE SCHEMA public;`, close, `core._ready.clear()`, `monkeypatch.setenv("DATABASE_URL", url)`, yield url.
- Replace every `core.connect(tmp_path / "<x>.db")` and `core.connect(db)` / `core.connect(path)` with `core.connect()`; every `monkeypatch.setenv("CARTON_DB", ...)` line is deleted (the fixture already set `DATABASE_URL`); unused `db = tmp_path / "app.db"` lines go.
- `insert()` uses `%s` placeholders and `with con.transaction():`. Every other `?` placeholder in check.py SQL becomes `%s`, every `with con:` becomes `with con.transaction():`.
- Delete `test_schema_v5_upgrade_removes_assignment_tables_without_losing_report_data`, `test_old_database_gets_the_corrections_table`, `test_old_database_gets_the_ebiz_table`.
- Change `test_connect_does_not_rewrite_a_current_schema` to call `core.connect()` twice with `_schema` monkeypatched to fail before the second call (unchanged intent).
- Add:

```python
def test_view_numbers_are_python_ints_and_floats():
    con = core.connect()
    insert(con, "STD", shipment_date="2026-08-01", carton_qty=100, unit_price=2.0, total_amount=200.0)
    r = con.execute("SELECT * FROM v_std").fetchone()
    assert type(r["fiscal_year"]) is int and type(r["quarter"]) is int
    assert type(r["pacd"]) is float and type(r["matalan_rebate"]) is float
    totals = con.execute("SELECT SUM(carton_qty), SUM(total_amount) FROM v_std").fetchone()
    assert type(totals[0]) is int and type(totals[1]) is float

def test_rows_act_like_sqlite_rows():
    con = core.connect()
    r = con.execute("SELECT 1 AS a, 'x' AS b").fetchone()
    assert (r[0], r["b"], dict(r), tuple(r)) == (1, "x", {"a": 1, "b": "x"}, (1, "x"))

def test_connect_without_url_is_a_clear_error(monkeypatch):
    monkeypatch.delenv("DATABASE_URL")
    with pytest.raises(RuntimeError, match="DATABASE_URL is not set"):
        core.connect()
```

- [ ] **Step 3: Run the schema tests to see them fail**

Run: `python -m pytest check.py -q -k "quarter_mapping or rate_follows or rebate_uses or without_date or connect or view_numbers or act_like" -p no:cacheprovider`
Expected: FAIL (`core.connect` still opens SQLite / `_ready` missing).

- [ ] **Step 4: Port connection and schema in `core.py`**

- `connect(url=None)`: `psycopg.connect(url or os.environ.get("DATABASE_URL"), autocommit=True, row_factory=_row_factory)`; if no URL, raise the `RuntimeError` above. If the URL is not in `_ready`: run `_schema()` and seed rates in one `con.transaction()`, then add the URL to `_ready`.
- `_row_factory(cursor)`: returns a maker building `Row`, a `tuple` subclass holding column names, with `keys()` and `__getitem__` accepting `str` (column lookup) or `int`/slice (tuple behaviour).
- `_schema()`: drop `PRAGMA`/version logic and legacy `DROP TABLE` lines. Tables: `id BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY`, `REAL` -> `DOUBLE PRECISION`, `INTEGER` and `TEXT` unchanged. Views via `CREATE OR REPLACE VIEW` (drop `DROP VIEW IF EXISTS`).
- `VIEW_SQL` (keep column names/order, so `SELECT *` users are unchanged):
  - `fiscal_year`: `EXTRACT(YEAR FROM t.shipment_date::date)::int + (EXTRACT(MONTH FROM t.shipment_date::date) = 12)::int`
  - `quarter`: same `CASE` as today on `EXTRACT(MONTH FROM t.shipment_date::date)::int`, `NULL` when `shipment_date IS NULL`.
  - rate join: `r.effective_from = (SELECT MAX(effective_from) FROM rates WHERE carton_type = '{kind}' AND effective_from::date <= t.shipment_date::date)`.
- Delete `SCHEMA_VERSION` and `DEFAULT_DB`.

- [ ] **Step 5: Run the schema tests to see them pass**

Run: same command as Step 3.
Expected: all selected tests PASS.

- [ ] **Step 6: Commit**

```bash
git add requirements.txt core.py check.py
git commit -m "Move the carton database connection and schema to PostgreSQL."
```

---

### Task 2: Port import, date-correction and eBiz queries

**Files:**
- Modify: `core.py` — `preview`, `check_dates`, `saved_dates`, `forget_dates`, `save_many`, `replace_ebiz`, `ebiz_status`, `ebiz_vendors`, `_SAME_KEY`
- Test: `check.py` (existing import, date and eBiz tests)

**Interfaces:**
- Consumes: Task 1 `connect`, `Row`, `con.transaction()`, `con.cursor().executemany`.
- Produces: unchanged signatures of every function listed above.

- [ ] **Step 1: Add the parameter-safety test**

```python
def test_supplier_names_with_quotes_and_percent_import_safely(tmp_path):
    con = core.connect()
    name = "M&U 100% O'Neil Pack Ltd."
    make_xlsx(tmp_path / "q.xlsx", std=[rec(sup=name)])
    data = core.read_workbook(tmp_path / "q.xlsx")
    assert core.preview(con, data)[0]["Supplier"] == "M&U"
    core.save_many(con, {"q.xlsx": data})
    row = con.execute("SELECT packaging_supplier, supplier_key FROM track_std").fetchone()
    assert tuple(row) == (name, "M&U")
```

- [ ] **Step 2: Run import, date and eBiz tests to see failures**

Run: `python -m pytest check.py -q -k "(import or date or reimport or spelling or files or batch or ebiz or correction or quotes) and not app" -p no:cacheprovider`
Expected: FAIL (SQLite `?` placeholders, `IS ?`, `INSERT OR REPLACE`, `executemany` on connection, `rowid`).

- [ ] **Step 3: Port the queries**

- `_SAME_KEY = "supplier_key IS NOT DISTINCT FROM %s AND substr(shipment_date, 1, 7) IS NOT DISTINCT FROM %s"`.
- All `?` -> `%s`; `', '.join('?' * n)` -> `', '.join(['%s'] * n)`.
- `save_many`, `replace_ebiz`, `forget_dates`: `with con.transaction():` and `con.cursor().executemany(...)`.
- `date_fixes` write: `INSERT INTO date_fixes VALUES (%s × 10) ON CONFLICT (row_key) DO UPDATE SET` every non-key column `= EXCLUDED.<column>`.
- `saved_dates`: `ORDER BY saved_at DESC, row_key`.

- [ ] **Step 4: Run the same tests to see them pass**

Run: same command as Step 2.
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add core.py check.py
git commit -m "Port tracker and eBiz imports and date corrections to PostgreSQL."
```

---

### Task 3: Port report queries, edits and pages; prove no connection leak

**Files:**
- Modify: `core.py` — `_where`, `_dated`, `rows`, `editable_rows`, `editable_row_count`, `update_tracker_rows`, `save_rates`, `_unrated`, `cross_reference`, `supplier_name`, `tracker_suppliers`, `quarter_totals`, `quarter_rates`, `fiscal_years`, `quarters`, `date_range`, `get_rates`
- Modify: `web_ui.py` (`page` wrapper closes nothing new; only touched if Step 1's error test needs it)
- Modify: `pages/*.py` only where a page holds a connection beyond `main()` (none expected)
- Test: `check.py`

**Interfaces:**
- Consumes: Task 1 and Task 2 interfaces.
- Produces: unchanged signatures; the full suite green on PostgreSQL.

- [ ] **Step 1: Add the outage and leak tests**

```python
def test_page_shows_error_panel_when_database_is_unreachable(monkeypatch):
    monkeypatch.setenv("DATABASE_URL", "postgresql://nobody:wrong@127.0.0.1:1/none?connect_timeout=2")
    at = AppTest.from_file("pages/1_Dashboard.py", default_timeout=30).run()
    assert any("could not be loaded" in e.value for e in at.error)

def test_page_runs_leave_no_open_connections(xref_con):
    def open_connections():
        return xref_con.execute(
            "SELECT COUNT(*) FROM pg_stat_activity WHERE datname = current_database() AND pid <> pg_backend_pid()"
        ).fetchone()[0]
    before = open_connections()
    for page in ("pages/1_Dashboard.py", "pages/6_Cross_Reference.py", "pages/7_Discrepancy_Report.py"):
        AppTest.from_file(page, default_timeout=60).run()
    gc.collect()
    assert open_connections() == before
```

(Use the same `AppTest` import the existing page smoke tests use; add `import gc`.)

- [ ] **Step 2: Run the full suite to see the remaining failures**

Run: `python -m pytest check.py -q -p no:cacheprovider`
Expected: FAIL in report, rates, tracker-edit, cross-reference and page tests.

- [ ] **Step 3: Port the queries**

- `supplier_key IS ?` -> `IS NOT DISTINCT FROM %s` (`editable_rows`, `editable_row_count`, `cross_reference` tracker and eBiz lookups, `supplier_name` both queries).
- `?` -> `%s` everywhere, including `LIMIT %s OFFSET %s`.
- `quarter_rates`: `SUM(pacd_pct IS NULL)` -> `COUNT(*) FILTER (WHERE pacd_pct IS NULL)`.
- `update_tracker_rows`, `save_rates`: `with con.transaction():` and cursor `executemany`; the `ValueError` raised inside `save_rates` still rolls back.
- If the leak test fails, close connections deterministically: in `web_ui.page`, close any `psycopg.Connection` the page returns or wrap `main()` so the page's `con` is closed in `finally`. Pick the smallest change that makes the test pass and keep `core.connect()` call sites unchanged.

- [ ] **Step 4: Run the full suite**

Run: `python -m pytest check.py -q -p no:cacheprovider`
Expected: `92 passed` (previous 89, minus 3 deleted migration tests, plus 6 new).

- [ ] **Step 5: Confirm SQLite is gone**

Run: `git grep -n "sqlite\|CARTON_DB\|PRAGMA\|rowid" -- core.py check.py pages web_ui.py reports.py`
Expected: no output.

- [ ] **Step 6: Commit**

```bash
git add core.py check.py web_ui.py pages
git commit -m "Port report queries, rates and tracker edits to PostgreSQL."
```

---

### Task 4: Container, local run and docs for the new prefix and database

**Files:**
- Modify: `Dockerfile` (drop `CARTON_DB` env; `--server.baseUrlPath=carton-supply-report`; healthcheck URL `http://localhost:8501/carton-supply-report/_stcore/health`)
- Modify: `docker-compose.yml` (drop `CARTON_DB` and the `carton-data` volume; `env_file: .env`; add read-only mount `./certs:/app/certs:ro`; keep `127.0.0.1:8501:8501`, `./Reports:/app/Reports`, `restart: unless-stopped`)
- Modify: `.env.example` (only `DATABASE_URL=` in the production form from Global Constraints with `<password>` placeholder, and `REPORTS_DIR=/app/Reports`)
- Modify: `.gitignore` and `.dockerignore` (add `certs/`)
- Modify: `run.bat` (if `DATABASE_URL` is unset, set it to `postgresql://postgres:postgres@localhost:5432/postgres` and print a line saying to start `carton-pg`; drop `CARTON_DB`)
- Modify: `deploy/nginx-matalan-carton-report.conf` (locations `= /carton-supply-report` redirect and `/carton-supply-report/`; htpasswd file `/etc/nginx/.htpasswd-carton-supply-report`)
- Modify: `README.md` (Run locally: start `carton-pg`; Run with Docker / production: `.env`, certs, prefix URL, update command `git pull && docker compose up -d --build`; Database and backups: RDS with 7-day backups replaces SQLite text; Testing: `TEST_DATABASE_URL`)
- Modify: `CHANGELOG.md` (Unreleased: PostgreSQL storage, new URL prefix, `CARTON_DB` replaced by `DATABASE_URL`)

**Interfaces:**
- Consumes: `DATABASE_URL` contract from Task 1.
- Produces: an image that serves `/carton-supply-report/` on 127.0.0.1:8501 given a `.env` and `certs/global-bundle.pem`.

- [ ] **Step 1: Make the file changes listed above**

- [ ] **Step 2: Build and run the image against local PostgreSQL**

Run (PowerShell): `docker build -t matalan-carton-report .` then
`docker run --rm -d --name carton-app -p 127.0.0.1:8501:8501 -e DATABASE_URL=postgresql://postgres:postgres@host.docker.internal:5432/postgres matalan-carton-report`
Expected: after ~20 s, `curl http://localhost:8501/carton-supply-report/_stcore/health` prints `ok`; `docker inspect --format "{{.State.Health.Status}}" carton-app` prints `healthy`.

- [ ] **Step 3: Load the page in the browser pane**

Open `http://localhost:8501/carton-supply-report/`; expected: Dashboard renders "No tracker data yet". Then `docker stop carton-app`.

- [ ] **Step 4: Validate compose**

Run: `Copy-Item .env.example .env.test; docker compose --env-file .env.test config --quiet; Remove-Item .env.test`
Expected: no output, exit 0.

- [ ] **Step 5: Commit and push**

```bash
git add Dockerfile docker-compose.yml .env.example .gitignore .dockerignore run.bat deploy README.md CHANGELOG.md
git commit -m "Serve the app at /carton-supply-report with DATABASE_URL configuration."
git push origin main
```

---

### Task 5: Deploy to Lightsail and RDS (operational, done with the user)

No product code. Each step names who acts; Claude never types a password.

- [ ] **Step 1: Server access and read-only check (Claude)**

Run locally: `aws lightsail get-instance-access-details --instance-name ubuntu_pacd --protocol ssh` and use the temporary key and `ubuntu` user to SSH. On the server: `nginx -v; docker --version; docker compose version; certbot --version; sudo ss -ltnp | grep -E ':(80|443|3000|8501)\b'; sudo nginx -T | grep -n "server_name\|listen\|location"`.
Expected: report findings to the user; if a package is missing, ask before `sudo apt install -y nginx docker.io docker-compose-v2 certbot python3-certbot-nginx postgresql-client`.

- [ ] **Step 2: Code and certificate (Claude)**

`sudo git clone https://github.com/PAC-D/Matalan-Carton-Supply-Report-v2 /opt/matalan-carton`, `sudo chown -R ubuntu:ubuntu /opt/matalan-carton`, then in that directory `mkdir -p certs Reports` and `curl -fsSo certs/global-bundle.pem https://truststore.pki.rds.amazonaws.com/global/global-bundle.pem`.
Expected: `git -C /opt/matalan-carton log -1 --oneline` shows the Task 4 commit; `certs/global-bundle.pem` is non-empty.

- [ ] **Step 3: App secret, database and app login (Claude writes, user runs psql)**

Claude, in `/opt/matalan-carton` with `umask 077`: generate the app password with `openssl rand -hex 24` into a shell variable (never echoed) and write `.env` containing `DATABASE_URL` (Global Constraints form) and `REPORTS_DIR=/app/Reports`; `chmod 600 .env`. Claude writes `/tmp/create-carton-db.sql` (no secrets in it):
`CREATE ROLE matalan_carton_app LOGIN PASSWORD :'app_pw'; GRANT matalan_carton_app TO postgres; CREATE DATABASE matalan_carton OWNER matalan_carton_app; REVOKE CONNECT ON DATABASE matalan_carton FROM PUBLIC;`
The user runs, and types the master password at the prompt:
`psql "host=database-1.ch4qkgwmwuqn.ap-south-1.rds.amazonaws.com user=postgres dbname=postgres sslmode=verify-full sslrootcert=/opt/matalan-carton/certs/global-bundle.pem" -v app_pw="$(sed -n 's/.*matalan_carton_app:\([^@]*\)@.*/\1/p' /opt/matalan-carton/.env)" -f /tmp/create-carton-db.sql`
Expected: `CREATE ROLE`, `GRANT ROLE`, `CREATE DATABASE`, `REVOKE`; then `rm /tmp/create-carton-db.sql`.

- [ ] **Step 4: Start the app (Claude)**

`cd /opt/matalan-carton && sudo docker compose up -d --build`
Expected: within 60 s `sudo docker compose ps` shows `healthy`; `curl -s localhost:8501/carton-supply-report/_stcore/health` prints `ok`; `sudo docker compose logs --tail 50` has no traceback.

- [ ] **Step 5: nginx and TLS (Claude edits, user sets the password)**

If `matalan.pac-d.com` has no HTTPS server block, run `sudo certbot --nginx -d matalan.pac-d.com` (user confirms the email prompt). Insert the two locations from `deploy/nginx-matalan-carton-report.conf` into that HTTPS block. The user runs `sudo htpasswd -c /etc/nginx/.htpasswd-carton-supply-report <username>` and types the password. Then `sudo nginx -t && sudo systemctl reload nginx`.
Expected: `nginx -t` prints `syntax is ok` and `test is successful`; `curl -sI https://matalan.pac-d.com/carton-supply-report/` returns `401` without credentials.

- [ ] **Step 6: End-to-end check (user + Claude)**

User signs in at `https://matalan.pac-d.com/carton-supply-report/`, imports `samples/Matalan Carton Shipment Tracking_NPS_BD Aug-26.xlsx` on Data import.
Expected: Dashboard shows fiscal year 2026 and TRANSWORLD/UNIGLORY totals; `psql` as `matalan_carton_app` shows `SELECT COUNT(*) FROM track_std` = 88 and `track_cus` = 118.

Rollback (only on user request): `sudo docker compose down`, remove the two nginx locations, reload nginx. Dropping `matalan_carton` is a separate step the user confirms.
