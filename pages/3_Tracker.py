import pandas as pd
import streamlit as st

import core
from web_ui import header, page

PAGE_SIZE = 250


def same_value(left, right):
    """Treat the different empty values produced by pandas as equivalent."""
    left_missing, right_missing = bool(pd.isna(left)), bool(pd.isna(right))
    return left_missing and right_missing if left_missing or right_missing else left == right


@page("Tracker data")
def main():
    con = core.connect()
    suppliers = core.tracker_suppliers(con)
    total = sum(con.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
                for table in core.TABLES.values())
    header("Tracker data",
           "Review and manually correct imported tracker values. Report calculations update immediately.",
           "Stored rows", f"{total:,}")
    if not suppliers:
        st.info("No tracker rows yet. Import a workbook on the Data import page.")
        return

    with st.container(border=True, key="card_tracker_filters"):
        a, b = st.columns(2)
        kind = a.segmented_control("Tracker sheet", ["STD", "CUS"], default="STD",
                                   format_func=lambda value: core.SHEETS[value])
        supplier = b.selectbox("Packaging supplier", suppliers,
                               format_func=lambda key: core.supplier_name(con, key))

    count = core.editable_row_count(con, kind, supplier)
    pages = max(1, (count + PAGE_SIZE - 1) // PAGE_SIZE)
    selected_page = 1
    if pages > 1:
        selected_page = st.selectbox(
            "Page", range(1, pages + 1),
            format_func=lambda number: f"Page {number} of {pages}", width=180)
    rows = core.editable_rows(con, kind, supplier, PAGE_SIZE,
                              (selected_page - 1) * PAGE_SIZE)
    if not rows:
        st.info("No rows match this sheet and supplier.")
        return

    frame = pd.DataFrame(rows)
    shown = ["id", *core.COLUMNS, "source_file", "imported_at"]
    frame = frame[shown]
    config = {
        "id": st.column_config.NumberColumn("Row ID", disabled=True, pinned=True),
        "source_file": st.column_config.TextColumn("Source file", disabled=True),
        "imported_at": st.column_config.TextColumn("Last changed", disabled=True),
        "po_received_date": st.column_config.TextColumn(
            core.HEADERS["po_received_date"], help="YYYY-MM-DD or DD/MM/YYYY"),
        "shipment_date": st.column_config.TextColumn(
            core.HEADERS["shipment_date"], help="YYYY-MM-DD or DD/MM/YYYY"),
    }
    for col in core.COLUMNS:
        if col not in config:
            config[col] = (st.column_config.NumberColumn(core.HEADERS[col])
                           if col in core.INT_COLS | core.REAL_COLS
                           else st.column_config.TextColumn(core.HEADERS[col]))

    first = (selected_page - 1) * PAGE_SIZE + 1
    last = first + len(rows) - 1
    st.caption(f"Editing rows {first:,}–{last:,} of {count:,} in {core.SHEETS[kind]} for "
               f"{core.supplier_name(con, supplier)}. "
               "Dates accept YYYY-MM-DD or DD/MM/YYYY. Changes are saved together.")
    edited = st.data_editor(frame, key=f"tracker_editor_{kind}_{supplier}_{selected_page}", hide_index=True,
                            num_rows="fixed", height=560, column_config=config)
    if st.button("Save tracker changes", type="primary", icon=":material/save:"):
        changed = []
        for index in frame.index:
            different = any(
                not same_value(frame.at[index, col], edited.at[index, col])
                for col in core.COLUMNS)
            if different:
                changed.append(edited.loc[index].to_dict())
        if not changed:
            st.info("No tracker changes to save.")
            return
        try:
            count = core.update_tracker_rows(con, kind, changed)
        except ValueError as exc:
            st.error(str(exc))
        else:
            st.toast(f"Saved {count:,} tracker rows", icon=":material/check_circle:")
            st.rerun()


main()
