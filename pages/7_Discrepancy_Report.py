import pandas as pd
import streamlit as st

import core
import reports
from web_ui import (
    dashboard_interactive_metric_grid,
    download_started,
    operation_status,
    page,
    report_page_header,
    search_frame,
    table_search,
)

TRACKER_COLUMNS = [column for column in core.COLUMNS if column != "note"] + ["pacd", "matalan_rebate"]


def tracker_frame(rows, carton_type, orders, include_ebiz=False):
    records = []
    for row in rows:
        item = {"carton_type": carton_type, **{column: row[column] for column in TRACKER_COLUMNS}}
        item["via_so2"] = bool(row["so"] and orders[row["so"]]["via_so2"])
        if include_ebiz:
            item["ebiz_qty"] = orders[row["so"]]["ebiz_qty"]
        records.append(item)
    frame = pd.DataFrame(records)
    if frame.empty:
        return frame
    labels = {"carton_type": "Carton type", **core.HEADERS,
              "pacd": "PACD Rebate", "matalan_rebate": "Matalan rebate", "ebiz_qty": "eBiz QTY",
              "via_so2": "Matched via SO_2"}
    return frame.rename(columns=labels)


def show_tracker_rows(frame, key):
    if frame.empty:
        st.info("No discrepancies found.")
        return
    frame = search_frame(frame, table_search("Search discrepancy rows", f"{key}_search"))
    if frame.empty:
        st.info("No discrepancy rows match your search.")
        return
    st.dataframe(
        frame, hide_index=True, height=min(620, 38 + 35 * len(frame)),
        column_config={
            "Carton type": st.column_config.TextColumn(pinned=True),
            "Carton QTY": st.column_config.NumberColumn(format="localized"),
            "eBiz QTY": st.column_config.NumberColumn(format="localized"),
            "Total Amount": st.column_config.NumberColumn(format="$ %.2f"),
            "PACD Rebate": st.column_config.NumberColumn(format="$ %.2f"),
            "Matalan rebate": st.column_config.NumberColumn(format="$ %.2f"),
            "Matched via SO_2": st.column_config.CheckboxColumn(),
        },
        key=key,
    )


@page
def main():
    con = core.connect()
    suppliers = core.reconciliation_suppliers(con)
    status = core.ebiz_status(con)
    report_page_header(
        "Discrepancy report",
        "eBiz<br>rows",
        f"{core.ebiz_row_count(con):,}",
    )
    if not suppliers:
        st.info("No Tracker or eBiz supplier data is available. Import data first.")
        return
    if not status:
        st.warning("No eBiz data is loaded. Tracker orders will appear as not found in eBiz.")

    names = {key: core.supplier_name(con, key) for key in suppliers}
    duplicates = {name for name in names.values()
                  if sum(value == name for value in names.values()) > 1}
    supplier_options = {
        f"{name} ({key})" if name in duplicates else name: key
        for key, name in names.items()
    }
    with st.container(key="discrepancy_filter_bar"):
        supplier_label = st.selectbox(
            "Packaging supplier",
            list(supplier_options),
            key="discrepancy_supplier",
            width="stretch",
        )
    supplier = supplier_options[supplier_label]
    xref = core.cross_reference(con, supplier)
    orders = {order["so"]: order for order in xref["orders"]}
    order_counts = {label: sum(order["status"] == label for order in xref["orders"])
                    for label in ("Not in eBiz", "Tracker > eBiz", "Tracker < eBiz")}
    no_so = sum(1 for rows in xref["rows"].values() for row in rows if not row["so"])
    active = dashboard_interactive_metric_grid([
        {"id": "missing", "label": "Missing orders",
         "value": f"{order_counts['Not in eBiz']:,}",
         "note": "Tracker orders not in eBiz", "tone": "matte-navy-deep"},
        {"id": "greater", "label": "Tracker > eBiz",
         "value": f"{order_counts['Tracker > eBiz']:,}",
         "note": "Higher Tracker quantity", "tone": "matte-crimson-deep"},
        {"id": "lower", "label": "Tracker < eBiz",
         "value": f"{order_counts['Tracker < eBiz']:,}",
         "note": "Lower Tracker quantity", "tone": "matte-navy-mid"},
        {"id": "no_so", "label": "Rows without SO#", "value": f"{no_so:,}",
         "note": "Tracker rows needing review", "tone": "matte-crimson-mid"},
        {"id": "ebiz_only", "label": "eBiz-only rows",
         "value": f"{len(xref['ebiz_only']):,}",
         "note": "Not found in Tracker", "tone": "matte-navy-soft"},
    ], "discrepancy", default="missing", columns=5)

    section = {
        "missing": ("Not in eBiz", "Tracker orders not found in eBiz"),
        "greater": ("Tracker > eBiz", "Tracker quantity is greater"),
        "lower": ("Tracker < eBiz", "Tracker quantity is lower"),
        "no_so": (None, "Tracker rows without SO#"),
    }
    if active == "ebiz_only":
        st.subheader("eBiz orders not found in Tracker")
        if xref["ebiz_only"]:
            frame = pd.DataFrame(xref["ebiz_only"])
            shown = list(core.EBIZ_HEADERS)
            frame = frame[shown].rename(columns=core.EBIZ_HEADERS)
            frame = search_frame(frame, table_search(
                "Search eBiz-only rows", "discrepancy_ebiz_only_search"))
            if frame.empty:
                st.info("No eBiz-only rows match your search.")
            else:
                st.dataframe(
                frame, hide_index=True, height=min(620, 38 + 35 * len(frame)),
                column_config={
                    "Vendor": st.column_config.TextColumn(pinned=True),
                    "Qty Shipped": st.column_config.NumberColumn(format="localized"),
                    "Current Net Price per 1000": st.column_config.NumberColumn(format="$ %.2f"),
                    "Current Gross Price per 1000": st.column_config.NumberColumn(format="$ %.2f"),
                    "Total Rebate": st.column_config.NumberColumn(format="$ %.2f"),
                },
                    key="discrepancy_ebiz_only",
                )
        else:
            st.info("No discrepancies found.")
    else:
        wanted, title = section[active]
        st.subheader(title)
        selected = []
        for kind, carton_type in (("STD", "Standard"), ("CUS", "Custom")):
            rows = [row for row in xref["rows"][kind]
                    if (orders[row["so"]]["status"] if row["so"] else None) == wanted]
            selected.append(tracker_frame(
                rows, carton_type, orders,
                include_ebiz=wanted in ("Tracker > eBiz", "Tracker < eBiz"),
            ))
        frames = [frame for frame in selected if not frame.empty]
        show_tracker_rows(pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(),
                          f"discrepancy_{active}")

    tracker_rows = [row for rows in xref["rows"].values() for row in rows]
    tracker_revision = max((row["imported_at"] or "" for row in tracker_rows), default="")
    tracker_totals = (
        sum(row["carton_qty"] or 0 for row in tracker_rows),
        round(sum((row["pacd"] or 0) + (row["matalan_rebate"] or 0)
                  for row in tracker_rows), 2),
    )
    build_key = (supplier, status[2] if status else None, tracker_revision, tracker_totals)
    artifact_key = core.cache_key("discrepancy-report-v2", build_key, core.data_revision(con))  # v2: [PDF, Excel]
    if "live_discrepancy_pdf" not in st.session_state:
        cached = core.cache_get(con, artifact_key)
        if cached is not None:
            st.session_state["live_discrepancy_pdf"] = (build_key, cached)
    if st.button("Prepare Discrepancy Report", icon=":material/description:"):
        with operation_status(
            "Preparing discrepancy PDF and Excel…", "Discrepancy files ready"
        ):
            cached = core.cache_get(con, artifact_key)
            if cached is None:
                cached = [reports.discrepancy_pdf(con, supplier, xref), reports.discrepancy_xlsx(con, supplier, xref)]
                core.cache_set(con, artifact_key, "discrepancy-report-v2", cached, max_entries=12)
            st.session_state["live_discrepancy_pdf"] = (build_key, cached)
    built = st.session_state.get("live_discrepancy_pdf")
    if built and built[0] == build_key:
        (pdf_name, pdf), (xlsx_name, xlsx) = built[1]
        with st.container(horizontal=True):
            st.download_button("Download Discrepancy Report PDF", pdf, file_name=pdf_name, mime="application/pdf",
                               icon=":material/download:", key="discrepancy_pdf_download",
                               on_click=download_started, args=(pdf_name,))
            st.download_button("Download Discrepancy Report Excel", xlsx, file_name=xlsx_name,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               icon=":material/download:", key="discrepancy_xlsx_download",
                               on_click=download_started, args=(xlsx_name,))


main()
