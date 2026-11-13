import os
from pathlib import Path

import streamlit as st

import core
import reports
from web_ui import page, reports_header

REPORTS_DIR = Path(os.environ.get(
    "REPORTS_DIR", Path(__file__).resolve().parents[1] / "Reports"))
REPORT_CACHE_VERSION = "built-reports-v3"


@st.fragment
@page
def report_controls():
    con = core.connect()
    suppliers = core.suppliers(con)
    names = {key: core.supplier_name(con, key) for key in suppliers}
    st.html("<div class='mc-reports-section-label'>Report setup</div>")
    with st.container(border=True, key="reports_panel_selection"):
        has_ebiz = core.ebiz_status(con) is not None
        with st.container(
                key="reports_selection_header", horizontal=True,
                horizontal_alignment="right", vertical_alignment="center", gap="small"):
            st.subheader("Report selection")
            with_disc = st.checkbox(
                "Discrepancy report and cross reference", value=has_ebiz,
                disabled=not has_ebiz, key="report_discrepancy",
                help="Compares all tracker rows with eBiz; period filters do not apply.",
            )
        years = core.fiscal_years(con)
        picked = st.multiselect(
            "Packaging suppliers", suppliers, default=suppliers,
            format_func=names.get, key="report_suppliers",
            placeholder="Select packaging suppliers",
        )
        if not picked:
            st.warning("Pick at least one supplier.")
            return

        year_column, quarter_column = st.columns([1, 4.4], gap="small")
        with year_column:
            year = st.selectbox(
                "Fiscal year", years, index=len(years) - 1,
                key="report_fiscal_year",
            )
        available = core.quarters(con, suppliers=picked, fiscal_year=year) if picked else []
        if not available:
            st.info("No data for this selection.")
            return
        with quarter_column:
            chosen = st.pills(
                "Quarters", available, default=available, selection_mode="multi",
                format_func=lambda q: core.quarter_label(year, q), width="content",
            )
        if not chosen:
            st.warning("Pick at least one quarter.")
            return

    # Keep the builder version in both keys so a code-only report layout change cannot reuse an old
    # database artifact or an already prepared download from this browser session.
    key = (tuple(picked), year, tuple(chosen), with_disc, REPORT_CACHE_VERSION)
    artifact_key = core.cache_key(REPORT_CACHE_VERSION, key, core.data_revision(con))
    if "built_reports" not in st.session_state:
        cached = core.cache_get(con, artifact_key)
        if cached is not None:
            st.session_state["built_reports"] = (key, *cached)

    built = st.session_state.get("built_reports")
    ready = bool(built and built[0] == key and built[1])
    if ready:
        _, items, zipped, zip_name = built
    else:
        items, zipped, zip_name = [], b"", "reports.zip"

    st.html("<div class='mc-reports-section-label'>Prepared files</div>")
    with st.container(border=True, key="reports_panel_downloads"):
        with st.container(
                key="reports_file_header", horizontal=True,
                horizontal_alignment="right", vertical_alignment="center", gap="small"):
            st.subheader("Individual files")
            build = st.button(
                "Build reports", type="primary", icon=":material/package_2:",
                key="report_build_button", width="content",
            )
            st.download_button(
                f"Download all ({len(items)} files)" if ready else "Download all",
                zipped, file_name=zip_name, type="secondary", icon=":material/download:",
                key="report_download_all", width="content", disabled=not ready,
            )
            save = st.button(
                "Save to Reports folder", icon=":material/save:",
                key="report_save_button", width="content",
            )
        if ready:
            with st.expander("Download individual files"):
                with st.container(key="reports_individual_files"):
                    for path, data in items:
                        st.download_button(
                            path, data, file_name=Path(path).name, key=path,
                            type="secondary", icon=":material/download:", width="content",
                        )
        else:
            st.caption("Build the current selection to prepare downloads.")
    if build:
        with st.spinner("Building reports…"):
            cached = core.cache_get(con, artifact_key)
            if cached is None:
                items = reports.build_all(con, picked, year, chosen, discrepancy=with_disc)
                month = reports.month_name(con, picked, year, chosen)
                cached = (items, reports.to_zip(items),
                          f"Matalan Carton Reports_{month} {year}.zip")
                core.cache_set(con, artifact_key, REPORT_CACHE_VERSION, cached, max_entries=12)
            st.session_state["built_reports"] = (key, *cached)
        st.rerun(scope="fragment")
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
    if built and built[0] == key and not built[1]:
        st.info("No data for this selection.")


@page
def main():
    con = core.connect()
    suppliers = core.suppliers(con)
    reports_header(len(suppliers))
    if not suppliers:
        st.info("No dated tracker data yet. Import a tracker workbook first.")
        return
    report_controls()


main()
