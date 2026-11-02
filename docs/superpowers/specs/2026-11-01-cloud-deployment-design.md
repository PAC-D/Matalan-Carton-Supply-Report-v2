# Cloud deployment: Lightsail + RDS PostgreSQL

**Date:** 2026-11-01
**Status:** Approved design, awaiting spec review

## Goal

Run the Matalan Carton Supply Report in the cloud at
`https://matalan.pac-d.com/carton-supply-report/`, with its data in the AWS RDS
PostgreSQL instance `database-1` instead of a local SQLite file.

Success means: the site loads at that URL behind basic auth, tracker and eBiz
imports write to RDS, every live page and generated report matches what the
SQLite version produced, and the full test suite passes against PostgreSQL.

## Decisions

| Topic | Decision |
|---|---|
| Database | PostgreSQL only (psycopg 3). SQLite support is removed. |
| Initial data | Start empty. Rates are seeded; tracker and eBiz data are re-imported through Data import. |
| Network to DB | Public RDS endpoint over TLS, the same way the existing Primark app connects. |
| Local dev + tests | `postgres:18` in Docker Desktop on the developer PC. |
| Hosting | Docker Compose on Lightsail `ubuntu_pacd`, behind host nginx. |

## Current environment (checked 2026-11-01)

- `database-1`: RDS PostgreSQL 18.6, `db.t4g.micro`, ap-south-1a, default VPC
  `vpc-0d2c29bafd157a455`, publicly accessible, storage encrypted, 7-day backups.
  Endpoint `database-1.ch4qkgwmwuqn.ap-south-1.rds.amazonaws.com:5432`, master user `postgres`.
- Security group `sg-0abf806525ce33662` allows 5432 from `0.0.0.0/0`
  ("Primark allocation app - password + forced TLS").
- Lightsail `ubuntu_pacd`: Ubuntu 24.04, static IP 65.2.228.79, ports 22/80/443/3000 open,
  VPC peering off. `matalan.pac-d.com` resolves to this IP.

## Architecture

```
browser --HTTPS--> nginx on ubuntu_pacd (matalan.pac-d.com, TLS, basic auth)
                     /carton-supply-report/ --> 127.0.0.1:8501 Docker container (Streamlit)
                                                  --TLS 5432--> database-1 / database matalan_carton
```

- The container runs the existing Dockerfile with `restart: unless-stopped` and
  publishes only `127.0.0.1:8501`.
- Streamlit `--server.baseUrlPath` becomes `carton-supply-report`; the Docker
  healthcheck path follows.
- The app uses its own database `matalan_carton` and login `matalan_carton_app`,
  which has no privileges on other databases.
- Configuration is one variable, `DATABASE_URL`, in `/opt/matalan-carton/.env`
  (mode 600, never committed). It replaces `CARTON_DB`. The URL uses
  `sslmode=verify-full&sslrootcert=/app/certs/global-bundle.pem` with the RDS CA bundle.
- "Save to Reports folder" keeps writing to `./Reports` on the server via the
  existing Compose bind mount.

## PostgreSQL port of core.py

Public function names and signatures stay the same, so pages and `reports.py`
need only minimal changes.

**Connection**
- `core.connect()` reads `DATABASE_URL` and opens an autocommit connection.
- Every `with con:` write block becomes `with con.transaction():`, keeping
  all-or-nothing imports, rate saves (including the "no rate" rollback) and tracker edits.
- `executemany` calls move to a cursor (`con.cursor().executemany`).
- A small row factory returns rows that support both `r["col"]` and `r[0]`,
  matching how `sqlite3.Row` is used today.
- Pages keep calling `core.connect()` per run; the connection closes when the
  page function returns. Ceiling: one TLS connect per rerun; switch to
  `psycopg_pool` if latency becomes noticeable.

**Schema**
- Created once per process: `CREATE TABLE IF NOT EXISTS` for all tables,
  `CREATE OR REPLACE VIEW` for `v_std`/`v_cus`, then seed rates if `rates` is empty.
- `PRAGMA user_version`, the SQLite upgrade path and the legacy table drops are removed.
- Types: `REAL` to `DOUBLE PRECISION`; `carton_qty` stays `INTEGER` (so `SUM`
  returns int, not `Decimal`); `id` columns become
  `BIGINT GENERATED ALWAYS AS IDENTITY PRIMARY KEY`.
- Dates stay ISO `TEXT` columns, unchanged for all Python callers. The views cast
  to `date` where they need date arithmetic.

**SQL rewrites**
- Placeholders `?` to `%s`.
- `IS ?` to `IS NOT DISTINCT FROM %s`.
- `INSERT OR REPLACE INTO date_fixes` to `INSERT ... ON CONFLICT (row_key) DO UPDATE`.
- View fiscal year/quarter: `strftime` to `EXTRACT(YEAR|MONTH FROM shipment_date::date)`;
  fiscal year is `year + (month = 12)::int`.
- Rate lookup in the views compares `effective_from::date <= shipment_date::date`.
- `SUM(pacd_pct IS NULL)` to `COUNT(*) FILTER (WHERE pacd_pct IS NULL)`.
- `ORDER BY saved_at DESC, rowid` to `ORDER BY saved_at DESC, row_key`.

## Tests

- `check.py` runs against PostgreSQL from `TEST_DATABASE_URL`
  (default `postgresql://postgres:postgres@localhost:5432/postgres`).
- A fixture drops and recreates schema `public` before each test.
- Tests that created temporary SQLite files use that fixture instead; tests that
  set `CARTON_DB` set `DATABASE_URL`.
- The SQLite schema-migration tests are deleted. All other assertions stay.
- If the test database is unreachable, the suite fails; it does not skip.
- Local start: `docker run -d --name carton-pg -e POSTGRES_PASSWORD=postgres -p 5432:5432 postgres:18`.

## Deployment steps (one time)

1. **Server check (read-only):** SSH with a temporary Lightsail key; confirm nginx,
   Docker, certbot, the `matalan.pac-d.com` server block, and what uses port 3000.
   Install missing packages with `apt`.
2. **Database (user runs):** a `psql` script, run on the server with the master
   password typed by the user, creates database `matalan_carton` and login
   `matalan_carton_app`, grants that role to `postgres` (required on PostgreSQL 16+
   before `CREATE DATABASE ... OWNER`), makes the login the database owner, and
   revokes `CONNECT` on that database from `PUBLIC`.
3. **Secret:** generate the app password on the server, write it into
   `/opt/matalan-carton/.env` (mode 600) and the `ALTER ROLE` statement without
   printing it. Download `global-bundle.pem` to `/opt/matalan-carton/certs/`.
4. **Code:** clone the public repo `https://github.com/PAC-D/Matalan-Carton-Supply-Report-v2`
   (branch `main`) to `/opt/matalan-carton` over HTTPS (no deploy key needed);
   `docker compose up -d --build`.
5. **nginx:** add the `/carton-supply-report/` location (from
   `deploy/nginx-matalan-carton-report.conf`, path updated) to the
   `matalan.pac-d.com` HTTPS server block; the user creates the basic-auth user
   with `htpasswd`; `nginx -t`; reload. If HTTPS is not yet set up for the
   domain, run `certbot --nginx -d matalan.pac-d.com` first.
6. **Verify:** health check through nginx, open the site, import one tracker
   workbook, confirm Dashboard totals.

**Later updates:** `git pull && docker compose up -d --build` (documented in README).

## Security

- App port bound to 127.0.0.1 only; public traffic only via nginx HTTPS + basic auth.
- DB connection verifies the RDS certificate (`verify-full`).
- App login is limited to its own database.
- Secrets exist only in the server `.env`; no password is typed by Claude or shown in chat.

## Rollback

`docker compose down` and remove the nginx location; nothing else on the server
changes. Dropping `matalan_carton` is a separate manual step, done only on the
user's confirmation.

## Out of scope

- Tightening `sg-0abf806525ce33662` (currently `0.0.0.0/0`; shared with the Primark app).
- Lightsail VPC peering, CI/CD, SSO, migrating existing local SQLite data.
