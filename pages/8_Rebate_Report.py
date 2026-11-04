import pandas as pd
import streamlit as st

import core
import reports
from web_ui import header, kpi_grid, page, search_frame, table_search

REPORT_COLUMNS = [column for column in core.COLUMNS if column != "note"]


def rebate_frame(rows):
    frame = pd.DataFrame(rows)
    if frame.empty:
        return frame
    shown = REPORT_COLUMNS + ["pacd_pct", "matalan_pct", "pacd", "matalan_rebate"]
    frame = frame[shown].rename(columns={
        **core.HEADERS,
        "pacd_pct": "PACD rate",
        "matalan_pct": "Rebate rate",
        "pacd": "PACD",
        "matalan_rebate": "Matalan rebate",
    })
    rates = ["PACD rate", "Rebate rate"]
    frame[rates] = frame[rates].map(lambda value: f"{value * 100:g}%" if pd.notna(value) else "-")
    return frame


def show_rebate_rows(frame, key):
    if frame.empty:
        st.info("No rows match this carton type and period.")
        return
    frame = search_frame(frame, table_search("Search rebate rows", f"{key}_search"))
    if frame.empty:
        st.info("No rebate rows match your search.")
        return
    st.dataframe(
        frame, hide_index=True, height=min(620, 38 + 35 * len(frame)),
        column_config={
            "Shipment Date": st.column_config.TextColumn(pinned=True),
            "Carton QTY": st.column_config.NumberColumn(format="localized"),
            "Total Amount": st.column_config.NumberColumn(format="$ %.2f"),
            "PACD": st.column_config.NumberColumn(format="$ %.2f"),
            "Matalan rebate": st.column_config.NumberColumn(format="$ %.2f"),
        },
        key=key,
    )


@page
def main():
    con = core.connect()
    suppliers = core.suppliers(con)
    header(
        "Rebate report",
        "View detailed Standard and Custom carton rebate calculations directly from the live Tracker database.",
        "Available suppliers",
        f"{len(suppliers):,}",
    )
    if not suppliers:
        st.info("No dated Tracker data is available. Import a tracker workbook first.")
        return

    with st.container(border=True, key="card_rebate_filters"):
        supplier = st.selectbox(
            "Packaging supplier", suppliers,
            format_func={key: core.supplier_name(con, key) for key in suppliers}.get,
            key="rebate_supplier",
        )
        years = core.fiscal_years(con, supplier=supplier)
        year = st.selectbox("Fiscal year", years, index=len(years) - 1, key="rebate_year")
        available = core.quarters(con, supplier=supplier, fiscal_year=year)
        quarters = st.pills(
            "Quarters", available, default=available, selection_mode="multi",
            format_func=lambda quarter: core.quarter_label(year, quarter),
            key="rebate_quarters",
        )
        so_only = st.radio(
            "Rows", ["All tracker rows", "Rows with eBiz SO#"], horizontal=True, key="rebate_rows",
            help="Rows with eBiz SO# leaves out rows whose eBiz Sales Order # is blank or 0.",
        ) == "Rows with eBiz SO#"
    if not quarters:
        st.warning("Choose at least one quarter.")
        return

    standard, custom = ([row for row in core.rows(con, kind, supplier=supplier, fiscal_year=year, quarters=quarters)
                         if not so_only or core._so(row["ebiz_so"])] for kind in ("STD", "CUS"))
    all_rows = standard + custom
    total_qty = sum(row["carton_qty"] or 0 for row in all_rows)
    total_amount = sum(row["total_amount"] or 0 for row in all_rows)
    pacd = sum(row["pacd"] or 0 for row in all_rows)
    rebate = sum(row["matalan_rebate"] or 0 for row in all_rows)
    kpi_grid([
        {"label": "Tracker rows", "value": f"{len(all_rows):,}"},
        {"label": "Carton quantity", "value": f"{total_qty:,.0f}"},
        {"label": "Total amount", "value": f"$ {total_amount:,.2f}"},
        {"label": "PACD", "value": f"$ {pacd:,.2f}"},
        {"label": "Matalan rebate", "value": f"$ {rebate:,.2f}"},
    ])

    standard_tab, custom_tab = st.tabs([
        f"Standard carton ({len(standard):,})", f"Custom carton ({len(custom):,})",
    ])
    with standard_tab:
        show_rebate_rows(rebate_frame(standard), "rebate_standard")
    with custom_tab:
        show_rebate_rows(rebate_frame(custom), "rebate_custom")

    if not all_rows:
        return
    tracker_revision = max((row["imported_at"] or "" for row in all_rows), default="")
    build_key = (supplier, year, tuple(sorted(quarters)), so_only, tracker_revision,
                 len(all_rows), total_qty, round(total_amount, 2), round(pacd + rebate, 2))
    artifact_key = core.cache_key("rebate-report-v2", build_key, core.data_revision(con))  # v2: [PDF, Excel]
    if "live_rebate_pdf" not in st.session_state:
        cached = core.cache_get(con, artifact_key)
        if cached is not None:
            st.session_state["live_rebate_pdf"] = (build_key, cached)
    if st.button("Prepare Rebate Report", icon=":material/description:"):
        with st.spinner("Preparing rebate PDF and Excel…"):
            cached = core.cache_get(con, artifact_key)
            if cached is None:
                cached = [reports.rebate_report_pdf(con, supplier, year, quarters, so_only=so_only),
                          reports.rebate_report_xlsx(con, supplier, year, quarters, so_only=so_only)]
                core.cache_set(con, artifact_key, "rebate-report-v2", cached, max_entries=12)
            st.session_state["live_rebate_pdf"] = (build_key, cached)
    built = st.session_state.get("live_rebate_pdf")
    if built and built[0] == build_key:
        (pdf_name, pdf), (xlsx_name, xlsx) = built[1]
        with st.container(horizontal=True):
            st.download_button("Download Rebate Report PDF", pdf, file_name=pdf_name, mime="application/pdf",
                               icon=":material/download:", key="rebate_pdf_download")
            st.download_button("Download Rebate Report Excel", xlsx, file_name=xlsx_name,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", icon=":material/download:", key="rebate_xlsx_download")


main()
