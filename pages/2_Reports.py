import os
from pathlib import Path

import streamlit as st

import core
import reports
from web_ui import header, page

REPORTS_DIR = Path(os.environ.get(
    "REPORTS_DIR", Path(__file__).resolve().parents[1] / "Reports"))


@st.fragment
def report_controls():
    con = core.connect()
    suppliers = core.suppliers(con)
    names = {key: core.supplier_name(con, key) for key in suppliers}
    with st.container(border=True, key="card_report_filters"):
        st.subheader("Report selection")
        picked = st.multiselect("Packaging suppliers", suppliers, default=suppliers,
                                format_func=names.get, key="report_suppliers")
        years = core.fiscal_years(con)
        year = st.selectbox("Fiscal year", years, index=len(years) - 1)
        available = core.quarters(con, suppliers=picked, fiscal_year=year) if picked else []
        if not picked:
            st.warning("Pick at least one supplier.")
            return
        if not available:
            st.info("No data for this selection.")
            return
        chosen = st.pills("Quarters", available, default=available, selection_mode="multi",
                          format_func=lambda q: core.quarter_label(year, q))
        if not chosen:
            st.warning("Pick at least one quarter.")
            return
        has_ebiz = core.ebiz_status(con) is not None
        with_disc = st.checkbox("Discrepancy report and cross reference", value=has_ebiz,
                                disabled=not has_ebiz,
                                key="report_discrepancy",
                                help="Compares all tracker rows with eBiz; period filters do not apply.")

    key = (tuple(picked), year, tuple(chosen), with_disc)
    with st.container(horizontal=True):
        build = st.button("Build reports", type="primary", icon=":material/package_2:")
        save = st.button("Save to Reports folder", icon=":material/save:")
    if build:
        with st.spinner("Building reports…"):
            items = reports.build_all(con, picked, year, chosen, discrepancy=with_disc)
            month = reports.month_name(con, picked, year, chosen)
            st.session_state["built_reports"] = (
                key, items, reports.to_zip(items), f"Matalan Carton Reports_{month} {year}.zip")
    if save:
        try:
            with st.spinner("Building reports…"):
                paths = reports.save_all(con, picked, year, chosen, REPORTS_DIR,
                                         discrepancy=with_disc)
        except OSError as exc:
            st.error(f"Could not save: {exc}. Close any open report files and try again.")
        else:
            st.toast(f"Saved {len(paths)} report files", icon=":material/check_circle:")

    built = st.session_state.get("built_reports")
    if built and built[0] == key:
        _, items, zipped, zip_name = built
        if not items:
            st.info("No data for this selection.")
        else:
            with st.container(border=True, key="card_report_downloads"):
                st.subheader("Downloads")
                st.download_button(f"Download all ({len(items)} files)", zipped,
                                   file_name=zip_name, type="primary",
                                   icon=":material/download:")
                with st.expander("Download individual files"):
                    for path, data in items:
                        st.download_button(path, data, file_name=Path(path).name, key=path,
                                           icon=":material/download:")


@page("Build reports")
def main():
    con = core.connect()
    suppliers = core.suppliers(con)
    header("Build reports", "Create branded Excel and PDF reports for the selected period.",
           "Available suppliers", str(len(suppliers)))
    if not suppliers:
        st.info("No dated tracker data yet. Import a tracker workbook first.")
        return
    report_controls()


main()
