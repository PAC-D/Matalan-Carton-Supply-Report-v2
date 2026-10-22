import pandas as pd
import streamlit as st

import core
import reports
from web_ui import header, kpi_grid, page

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
    frame["PACD rate"] = frame["PACD rate"].map(
        lambda value: f"{value * 100:g}%" if pd.notna(value) else "-")
    frame["Rebate rate"] = frame["Rebate rate"].map(
        lambda value: f"{value * 100:g}%" if pd.notna(value) else "-")
    return frame


def show_rebate_rows(frame, key):
    if frame.empty:
        st.info("No rows match this carton type and period.")
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


@page("Rebate report")
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
            format_func=lambda key: core.supplier_name(con, key),
            key="rebate_supplier",
        )
        years = sorted({row[0] for row in con.execute(
            "SELECT DISTINCT fiscal_year FROM ("
            "SELECT fiscal_year, supplier_key FROM v_std UNION ALL "
            "SELECT fiscal_year, supplier_key FROM v_cus) "
            "WHERE supplier_key = ? AND fiscal_year IS NOT NULL", (supplier,))})
        year = st.selectbox("Fiscal year", years, index=len(years) - 1, key="rebate_year")
        available = core.quarters(con, supplier=supplier, fiscal_year=year)
        quarters = st.pills(
            "Quarters", available, default=available, selection_mode="multi",
            format_func=lambda quarter: core.quarter_label(year, quarter),
            key="rebate_quarters",
        )
    if not quarters:
        st.warning("Choose at least one quarter.")
        return

    standard = core.rows(con, "STD", supplier=supplier, fiscal_year=year, quarters=quarters)
    custom = core.rows(con, "CUS", supplier=supplier, fiscal_year=year, quarters=quarters)
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

    tracker_revision = max((row["imported_at"] or "" for row in all_rows), default="")
    build_key = (supplier, year, tuple(sorted(quarters)), tracker_revision,
                 len(all_rows), total_qty, round(total_amount, 2), round(pacd + rebate, 2))
    if st.button("Prepare Rebate Report PDF", icon=":material/picture_as_pdf:"):
        with st.spinner("Building PDF..."):
            st.session_state["live_rebate_pdf"] = (
                build_key, *reports.rebate_report_pdf(con, supplier, year, quarters))
    built = st.session_state.get("live_rebate_pdf")
    if built and built[0] == build_key:
        _, filename, data = built
        st.download_button(
            "Download Rebate Report PDF", data, file_name=filename,
            mime="application/pdf", icon=":material/download:", key="rebate_pdf_download",
        )


main()
