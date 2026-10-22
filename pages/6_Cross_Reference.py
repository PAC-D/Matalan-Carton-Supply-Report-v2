import pandas as pd
import streamlit as st

import core
import reports
from web_ui import header, kpi_grid, page


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
    kpi_grid([
        {"label": "Matched orders", "value": f"{matched:,}", "tone": "ok"},
        {"label": "Missing from eBiz", "value": f"{missing:,}",
         "tone": "warn" if missing else "ok"},
        {"label": "Quantity mismatches", "value": f"{mismatches:,}",
         "tone": "warn" if mismatches else "ok"},
        {"label": "eBiz-only orders", "value": f"{len(ebiz_only):,}",
         "tone": "warn" if ebiz_only else "ok"},
        {"label": "Tracker rows without SO#", "value": f"{no_so:,}",
         "tone": "warn" if no_so else "ok"},
    ])

    if rows:
        with st.container(border=True, key="card_cross_reference_orders"):
            st.subheader("Tracker orders")
            frame = pd.DataFrame(rows).rename(columns={
                "so": "SO#", "tracker_qty": "Tracker QTY",
                "tracker_rebate": "Tracker Rebate", "ebiz_qty": "eBiz QTY",
                "ebiz_rebate": "eBiz Rebate", "qty_difference": "QTY Difference",
                "status": "Status", "via_so2": "Matched via SO_2",
            })
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
        st.info("No Tracker orders exist for this packaging supplier.")

    with st.container(border=True, key="card_cross_reference_ebiz_only"):
        st.subheader("eBiz orders not found in Tracker")
        if ebiz_only:
            frame = pd.DataFrame(ebiz_only).rename(columns={
                "order": "Order", "ebiz_qty": "eBiz QTY", "ebiz_rebate": "eBiz Rebate",
            })
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

    filename, data = reports.cross_reference_xlsx(con, supplier, xref)
    st.download_button(
        "Download Cross Reference workbook", data, file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        icon=":material/download:", key="cross_reference_download",
    )


main()
