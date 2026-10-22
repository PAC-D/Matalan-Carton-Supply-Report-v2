import pandas as pd
import streamlit as st

import core
import reports
from web_ui import header, kpi_grid, page

TRACKER_COLUMNS = [column for column in core.COLUMNS if column != "note"] + ["pacd", "matalan_rebate"]
SECTIONS = [
    ("Not in eBiz", "Tracker orders not found in eBiz"),
    ("Tracker > eBiz", "Tracker quantity is greater"),
    ("Tracker < eBiz", "Tracker quantity is lower"),
    (None, "Tracker rows without SO#"),
]


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
              "pacd": "PACD", "matalan_rebate": "Matalan rebate", "ebiz_qty": "eBiz QTY",
              "via_so2": "Matched via SO_2"}
    return frame.rename(columns=labels)


def show_tracker_rows(frame, key):
    if frame.empty:
        st.info("No discrepancies found.")
        return
    st.dataframe(
        frame, hide_index=True, height=min(620, 38 + 35 * len(frame)),
        column_config={
            "Carton type": st.column_config.TextColumn(pinned=True),
            "Carton QTY": st.column_config.NumberColumn(format="localized"),
            "eBiz QTY": st.column_config.NumberColumn(format="localized"),
            "Total Amount": st.column_config.NumberColumn(format="$ %.2f"),
            "PACD": st.column_config.NumberColumn(format="$ %.2f"),
            "Matalan rebate": st.column_config.NumberColumn(format="$ %.2f"),
            "Matched via SO_2": st.column_config.CheckboxColumn(),
        },
        key=key,
    )


@page("Discrepancy report")
def main():
    con = core.connect()
    suppliers = sorted(set(core.tracker_suppliers(con)) | set(core.ebiz_vendors(con)))
    status = core.ebiz_status(con)
    header(
        "Discrepancy report",
        "Review the five live exception sections used by the generated Discrepancy Report PDF.",
        "eBiz status",
        f"{status[0]:,} rows loaded" if status else "Not loaded",
    )
    if not suppliers:
        st.info("No Tracker or eBiz supplier data is available. Import data first.")
        return
    if not status:
        st.warning("No eBiz data is loaded. Tracker orders will appear as not found in eBiz.")

    supplier = st.selectbox(
        "Packaging supplier", suppliers,
        format_func=lambda key: core.supplier_name(con, key),
        key="discrepancy_supplier",
    )
    xref = core.cross_reference(con, supplier)
    orders = {order["so"]: order for order in xref["orders"]}
    order_counts = {label: sum(order["status"] == label for order in xref["orders"])
                    for label in ("Not in eBiz", "Tracker > eBiz", "Tracker < eBiz")}
    no_so = sum(1 for rows in xref["rows"].values() for row in rows if not row["so"])
    kpi_grid([
        {"label": "Missing orders", "value": f"{order_counts['Not in eBiz']:,}",
         "tone": "warn" if order_counts["Not in eBiz"] else "ok"},
        {"label": "Tracker > eBiz", "value": f"{order_counts['Tracker > eBiz']:,}",
         "tone": "warn" if order_counts["Tracker > eBiz"] else "ok"},
        {"label": "Tracker < eBiz", "value": f"{order_counts['Tracker < eBiz']:,}",
         "tone": "warn" if order_counts["Tracker < eBiz"] else "ok"},
        {"label": "Rows without SO#", "value": f"{no_so:,}",
         "tone": "warn" if no_so else "ok"},
        {"label": "eBiz-only rows", "value": f"{len(xref['ebiz_only']):,}",
         "tone": "warn" if xref["ebiz_only"] else "ok"},
    ])

    labels = [title for _, title in SECTIONS] + ["eBiz orders not found in Tracker"]
    tabs = st.tabs(labels)
    for tab, (wanted, title) in zip(tabs[:4], SECTIONS):
        with tab:
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
                              f"discrepancy_{wanted or 'no_so'}")

    with tabs[4]:
        if xref["ebiz_only"]:
            frame = pd.DataFrame(xref["ebiz_only"])
            shown = list(core.EBIZ_HEADERS)
            frame = frame[shown].rename(columns=core.EBIZ_HEADERS)
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

    tracker_rows = [row for rows in xref["rows"].values() for row in rows]
    tracker_revision = max((row["imported_at"] or "" for row in tracker_rows), default="")
    tracker_totals = (
        sum(row["carton_qty"] or 0 for row in tracker_rows),
        round(sum((row["pacd"] or 0) + (row["matalan_rebate"] or 0)
                  for row in tracker_rows), 2),
    )
    build_key = (supplier, status[2] if status else None, tracker_revision, tracker_totals)
    if st.button("Prepare Discrepancy Report PDF", icon=":material/picture_as_pdf:"):
        with st.spinner("Building PDF..."):
            st.session_state["live_discrepancy_pdf"] = (
                build_key, *reports.discrepancy_pdf(con, supplier, xref))
    built = st.session_state.get("live_discrepancy_pdf")
    if built and built[0] == build_key:
        _, filename, data = built
        st.download_button(
            "Download Discrepancy Report PDF", data, file_name=filename,
            mime="application/pdf", icon=":material/download:",
            key="discrepancy_pdf_download",
        )


main()
