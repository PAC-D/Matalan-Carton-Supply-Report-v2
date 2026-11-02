import pandas as pd
import streamlit as st

import core
import reports
from web_ui import header, interactive_kpi_grid, page, search_frame, table_search


def reconciliation_suppliers(con):
    return sorted(set(core.tracker_suppliers(con)) | set(core.ebiz_vendors(con)))


@page("Cross reference")
def main():
    con = core.connect()
    suppliers = reconciliation_suppliers(con)
    status = core.ebiz_status(con)
    header(
        "Cross reference",
        "Compare every Tracker sales order with the current eBiz data, using the same live calculations as the generated report.",
        "eBiz rows",
        f"{status[0]:,}" if status else "Not loaded",
    )
    if not suppliers:
        st.info("No Tracker or eBiz supplier data is available. Import data first.")
        return
    if not status:
        st.warning("No eBiz data is loaded. Tracker orders will appear as not found in eBiz.")

    supplier = st.selectbox(
        "Packaging supplier", suppliers,
        format_func=lambda key: core.supplier_name(con, key),
        key="cross_reference_supplier",
    )
    xref = core.cross_reference(con, supplier)
    rows = core.cross_reference_rows(xref)
    ebiz_only = core.ebiz_only_orders(xref)
    matched = sum(row["status"] == "Match" for row in rows)
    missing = sum(row["status"] == "Not in eBiz" for row in rows)
    mismatches = sum(row["status"] in ("Tracker > eBiz", "Tracker < eBiz") for row in rows)
    no_so = sum(1 for kind_rows in xref["rows"].values() for row in kind_rows if not row["so"])
    active = interactive_kpi_grid([
        {"id": "all", "label": "All tracker orders", "value": f"{len(rows):,}"},
        {"id": "match", "label": "Matched orders", "value": f"{matched:,}"},
        {"id": "missing", "label": "Missing from eBiz", "value": f"{missing:,}"},
        {"id": "mismatch", "label": "Quantity mismatches", "value": f"{mismatches:,}"},
        {"id": "ebiz_only", "label": "eBiz-only orders", "value": f"{len(ebiz_only):,}"},
        {"id": "no_so", "label": "Tracker rows without SO#", "value": f"{no_so:,}"},
    ], "cross_reference")

    wanted = {
        "match": {"Match"}, "missing": {"Not in eBiz"},
        "mismatch": {"Tracker > eBiz", "Tracker < eBiz"}, "no_so": {None},
    }.get(active)
    shown_rows = [row for row in rows if wanted is None or row["status"] in wanted]
    if active == "ebiz_only":
        shown_rows = []

    if shown_rows:
        with st.container(border=True, key="card_cross_reference_orders"):
            st.subheader("Tracker orders")
            frame = pd.DataFrame(shown_rows).rename(columns={
                "so": "SO#", "tracker_qty": "Tracker QTY",
                "tracker_rebate": "Tracker Rebate", "ebiz_qty": "eBiz QTY",
                "ebiz_rebate": "eBiz Rebate", "qty_difference": "QTY Difference",
                "status": "Status", "via_so2": "Matched via SO_2",
            })
            frame = search_frame(frame, table_search(
                "Search tracker orders", "cross_reference_order_search"))
            if frame.empty:
                st.info("No tracker orders match your search.")
            else:
                st.dataframe(
                frame, hide_index=True, height=min(620, 38 + 35 * len(frame)),
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
    else:
        st.info("No Tracker orders match the active card.")

    with st.container(border=True, key="card_cross_reference_ebiz_only"):
        st.subheader("eBiz orders not found in Tracker")
        if ebiz_only:
            frame = pd.DataFrame(ebiz_only).rename(columns={
                "order": "Order", "ebiz_qty": "eBiz QTY", "ebiz_rebate": "eBiz Rebate",
            })
            frame = search_frame(frame, table_search(
                "Search eBiz-only orders", "cross_reference_ebiz_search"))
            if frame.empty:
                st.info("No eBiz-only orders match your search.")
            else:
                st.dataframe(
                frame, hide_index=True,
                column_config={
                    "Order": st.column_config.TextColumn(pinned=True),
                    "eBiz QTY": st.column_config.NumberColumn(format="localized"),
                    "eBiz Rebate": st.column_config.NumberColumn(format="$ %.2f"),
                },
                    key="cross_reference_ebiz_only_table",
                )
        else:
            st.success("No eBiz-only orders for this supplier.")

    artifact_key = core.cache_key(
        "cross-reference-xlsx-v1", supplier, core.data_revision(con))
    artifact = core.cache_get(con, artifact_key)
    if artifact is None:
        artifact = reports.cross_reference_xlsx(con, supplier, xref)
        core.cache_set(con, artifact_key, "cross-reference-xlsx-v1", artifact, max_entries=12)
    filename, data = artifact
    st.download_button(
        "Download Cross Reference workbook", data, file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        icon=":material/download:", key="cross_reference_download",
    )


main()
