import pandas as pd
import streamlit as st

import core
import reports
from web_ui import (
    cross_reference_header,
    dashboard_interactive_metric_grid,
    page,
    search_frame,
    table_search,
)


def tracker_orders_table(rows):
    frame = pd.DataFrame(rows).rename(columns={
        "so": "SO#",
        "tracker_qty": "Tracker QTY",
        "tracker_rebate": "Tracker Rebate",
        "ebiz_qty": "eBiz QTY",
        "ebiz_rebate": "eBiz Rebate",
        "qty_difference": "QTY Difference",
        "status": "Status",
        "via_so2": "Matched via SO_2",
    })
    frame = search_frame(
        frame,
        table_search("Search tracker orders", "cross_reference_order_search"),
    )
    if frame.empty:
        st.info("No tracker orders match your search.")
        return
    st.dataframe(
        frame,
        hide_index=True,
        height=min(560, 38 + 35 * len(frame)),
        width="stretch",
        column_config={
            "SO#": st.column_config.TextColumn(pinned=True),
            "Tracker QTY": st.column_config.NumberColumn(format="localized"),
            "Tracker Rebate": st.column_config.NumberColumn(format="$ %.2f"),
            "eBiz QTY": st.column_config.NumberColumn(format="localized"),
            "eBiz Rebate": st.column_config.NumberColumn(format="$ %.2f"),
            "QTY Difference": st.column_config.NumberColumn(format="localized"),
            "Matched via SO_2": st.column_config.CheckboxColumn(),
        },
        key="cross_reference_orders_table",
    )


def ebiz_only_table(rows):
    frame = pd.DataFrame(rows).rename(columns={
        "order": "Order",
        "ebiz_qty": "eBiz QTY",
        "ebiz_rebate": "eBiz Rebate",
    })
    frame = search_frame(
        frame,
        table_search("Search eBiz-only orders", "cross_reference_ebiz_search"),
    )
    if frame.empty:
        st.info("No eBiz-only orders match your search.")
        return
    st.dataframe(
        frame,
        hide_index=True,
        height=min(420, 38 + 35 * len(frame)),
        width="stretch",
        column_config={
            "Order": st.column_config.TextColumn(pinned=True),
            "eBiz QTY": st.column_config.NumberColumn(format="localized"),
            "eBiz Rebate": st.column_config.NumberColumn(format="$ %.2f"),
        },
        key="cross_reference_ebiz_only_table",
    )


@st.fragment
@page
def reconciliation_view():
    con = core.connect()
    suppliers = core.reconciliation_suppliers(con)
    status = core.ebiz_status(con)
    if not suppliers:
        st.info("No Tracker or eBiz supplier data is available. Import data first.")
        return
    if not status:
        st.warning("No eBiz data is loaded. Tracker orders will appear as not found in eBiz.")

    names = {key: core.supplier_name(con, key) for key in suppliers}
    with st.container(key="cross_reference_filter_bar"):
        supplier = st.selectbox(
            "Packaging supplier",
            suppliers,
            format_func=names.get,
            key="cross_reference_supplier",
            width="stretch",
        )

    xref = core.cross_reference(con, supplier)
    rows = core.cross_reference_rows(xref)
    ebiz_only = core.ebiz_only_orders(xref)
    matched = sum(row["status"] == "Match" for row in rows)
    missing = sum(row["status"] == "Not in eBiz" for row in rows)
    mismatches = sum(
        row["status"] in ("Tracker > eBiz", "Tracker < eBiz") for row in rows
    )
    no_so = sum(
        1 for kind_rows in xref["rows"].values() for row in kind_rows if not row["so"]
    )

    active = dashboard_interactive_metric_grid([
        {
            "id": "all", "label": "Tracker orders", "value": f"{len(rows):,}",
            "note": "Complete supplier view", "tone": "navy",
        },
        {
            "id": "match", "label": "Matched orders", "value": f"{matched:,}",
            "note": "Tracker and eBiz aligned", "tone": "navy-mid",
        },
        {
            "id": "missing", "label": "Missing from eBiz", "value": f"{missing:,}",
            "note": "Tracker orders not found", "tone": "red",
        },
        {
            "id": "mismatch", "label": "QTY mismatches", "value": f"{mismatches:,}",
            "note": "Different shipped quantities", "tone": "red-mid",
        },
        {
            "id": "ebiz_only", "label": "eBiz-only orders", "value": f"{len(ebiz_only):,}",
            "note": "Not found in Tracker", "tone": "navy-soft",
        },
        {
            "id": "no_so", "label": "Rows without SO#", "value": f"{no_so:,}",
            "note": "Tracker rows needing review", "tone": "red-soft",
        },
    ], "cross_reference", columns=6)

    wanted = {
        "match": {"Match"},
        "missing": {"Not in eBiz"},
        "mismatch": {"Tracker > eBiz", "Tracker < eBiz"},
        "no_so": {None},
    }.get(active)
    shown_rows = [row for row in rows if wanted is None or row["status"] in wanted]
    if active == "ebiz_only":
        shown_rows = []

    st.html("<div class='mc-cross-reference-section-label'>Order reconciliation</div>")
    if active != "ebiz_only":
        with st.container(border=True, key="cross_reference_panel_tracker_orders"):
            st.subheader("Tracker orders")
            if shown_rows:
                tracker_orders_table(shown_rows)
            else:
                st.info("No Tracker orders match the active card.")

    if active in {"all", "ebiz_only"}:
        with st.container(border=True, key="cross_reference_panel_ebiz_only"):
            st.subheader("eBiz orders not found in Tracker")
            if ebiz_only:
                ebiz_only_table(ebiz_only)
            else:
                st.success("No eBiz-only orders for this supplier.")

    artifact_key = core.cache_key(
        "cross-reference-xlsx-v1", supplier, core.data_revision(con)
    )
    artifact = core.cache_get(con, artifact_key)
    if artifact is None:
        artifact = reports.cross_reference_xlsx(con, supplier, xref)
        core.cache_set(
            con, artifact_key, "cross-reference-xlsx-v1", artifact, max_entries=12
        )
    filename, data = artifact
    with st.container(horizontal=True, horizontal_alignment="right"):
        st.download_button(
            "Download Cross Reference workbook",
            data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            icon=":material/download:",
            key="cross_reference_download",
            width="content",
        )


@page
def main():
    con = core.connect()
    cross_reference_header(core.ebiz_row_count(con))
    reconciliation_view()


main()
