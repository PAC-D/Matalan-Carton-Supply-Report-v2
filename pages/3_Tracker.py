import pandas as pd
import streamlit as st

import core
from web_ui import (
    clear_built_reports,
    operation_status,
    page,
    report_page_header,
    table_search,
)

PAGE_SIZE = 250
DELETE_MODES = ["Selected rows", "Supplier and month", "Imported file"]


def same_value(left, right):
    """Treat the different empty values produced by pandas as equivalent."""
    left_missing, right_missing = bool(pd.isna(left)), bool(pd.isna(right))
    return left_missing and right_missing if left_missing or right_missing else left == right


def cancel_delete():
    st.session_state.pop("tracker_delete_pending", None)


@st.dialog("Delete tracker data", on_dismiss=cancel_delete)
def confirm_delete(what, action):
    """Ask once more, then run a grouped-row, supplier-month, or imported-file deletion.

    The pending request lives in session state, so the dialog survives full reruns until it is used or dismissed.
    """
    st.write(f"Delete {what}?")
    st.caption("This cannot be undone here. RDS keeps 7 days of backups.")
    if st.button("Delete", type="primary", icon=":material/delete:"):
        with operation_status("Deleting tracker data…", "Tracker data deleted"):
            con = core.connect()  # its own connection: nothing here keeps a page connection alive
            if action[0] == "rows":
                deleted = sum(
                    core.delete_tracker_rows(con, kind, ids)
                    for kind, ids in action[1].items()
                )
            elif action[0] == "month":
                deleted = sum(core.delete_supplier_month(con, action[1], action[2]).values())
            else:
                deleted = sum(core.delete_source_file(con, action[1]).values())
            clear_built_reports()
            cancel_delete()
            st.session_state["tracker_deleted"] = deleted
        st.rerun()


def ask(what, action):
    st.session_state["tracker_delete_pending"] = (what, action)


def delete_tools(con, kind, supplier, name, edited, combined=False):
    tools(con, kind, supplier, name, edited, combined)
    pending = st.session_state.get("tracker_delete_pending")
    if pending:
        confirm_delete(*pending)


def tools(con, kind, supplier, name, edited, combined=False):
    with st.expander("Delete tracker data", icon=":material/delete:"):
        mode = st.radio("Delete", DELETE_MODES, key="tracker_delete_mode", horizontal=True)
        if mode == "Selected rows":
            selected = {}
            if edited is not None:
                for i in edited.index:
                    if edited.at[i, "select"]:
                        row_kind = edited.at[i, "carton_type"] if combined else kind
                        selected.setdefault(row_kind, []).append(int(edited.at[i, "id"]))
            count = sum(len(ids) for ids in selected.values())
            st.caption("Tick rows in the Select column of the table above.")
            if st.button(f"Delete {count:,} selected rows", disabled=not count, key="tracker_delete_rows"):
                detail = " and ".join(f"{len(ids):,} {core.SHEETS[row_kind]}" for row_kind, ids in selected.items())
                ask(f"{detail} rows of {name}", ("rows", selected))
        elif mode == "Supplier and month":
            months = core.tracker_months(con, supplier)
            if not months:
                st.info(f"{name} has no dated rows.")
                return
            month = st.selectbox("Shipment month", months, key="tracker_delete_month")
            counts = {k: sum(1 for r in core.editable_rows(con, k, supplier) if (r["shipment_date"] or "")[:7] == month)
                      for k in core.TABLES}
            what = f"{counts['STD']:,} Standard and {counts['CUS']:,} Custom rows of {name} shipped in {month}"
            st.caption(what[0].upper() + what[1:] + ". Rows without a shipment date are kept.")
            if st.button("Delete this month", key="tracker_delete_month_button"):
                ask(what, ("month", supplier, month))
        else:
            files = core.tracker_files(con)
            if not files:
                st.info("No imported files are recorded.")
                return
            labels = {f"{f['source_file']} · {f['STD']:,} STD · {f['CUS']:,} CUS · {(f['imported_at'] or '')[:16]}": f
                      for f in files}
            picked = labels[st.selectbox("Imported file", list(labels), key="tracker_delete_file")]
            what = (f"all {picked['STD']:,} Standard and {picked['CUS']:,} Custom rows imported from "
                    f"{picked['source_file']}")
            if st.button("Delete this file's rows", key="tracker_delete_file_button"):
                ask(what, ("file", picked["source_file"]))


@page
def main():
    con = core.connect()
    deleted = st.session_state.pop("tracker_deleted", None)
    if deleted is not None:
        st.toast(f"Deleted {deleted:,} tracker rows", icon=":material/check_circle:")
    suppliers = core.tracker_suppliers(con)
    total = sum(core.editable_row_count(con, kind) for kind in core.TABLES)
    report_page_header("Tracker data", "Stored<br>rows", f"{total:,}")
    if not suppliers:
        st.info("No tracker rows yet. Import a workbook on the Data import page.")
        return

    with st.container(key="manage_filter_tracker"):
        a, b = st.columns(2)
        kind = a.segmented_control("Tracker sheet", ["STD", "CUS"], default="STD",
                                   format_func=lambda value: core.SHEETS[value])
        names = {key: core.supplier_name(con, key) for key in suppliers}
        supplier = b.selectbox("Packaging supplier", suppliers, format_func=names.get)
        search = table_search("Search tracker rows", "tracker_search",
                              "Search SO, PO, factory, carton style, dates, quantities…")

    combined = bool(search and search.strip())
    count = (core.editable_row_count_all(con, supplier, search=search) if combined
             else core.editable_row_count(con, kind, supplier, search=search))
    pages = max(1, (count + PAGE_SIZE - 1) // PAGE_SIZE)
    selected_page = 1
    if pages > 1:
        selected_page = st.selectbox(
            "Page", range(1, pages + 1),
            format_func=lambda number: f"Page {number} of {pages}", width=180)
    rows = (core.editable_rows_all(con, supplier, PAGE_SIZE, (selected_page - 1) * PAGE_SIZE, search=search)
            if combined else
            core.editable_rows(con, kind, supplier, PAGE_SIZE, (selected_page - 1) * PAGE_SIZE, search=search))
    if not rows:
        st.info("No rows match this supplier and search." if combined
                else "No rows match this sheet and supplier.")
        delete_tools(con, kind, supplier, names[supplier], None, combined)
        return

    frame = pd.DataFrame(rows)
    shown = (["carton_type"] if combined else []) + ["id", *core.COLUMNS, "source_file"]
    frame = frame[shown]
    frame.insert(0, "select", False)
    config = {
        "select": st.column_config.CheckboxColumn("Select", default=False, pinned=True),
        "id": st.column_config.NumberColumn("Row ID", disabled=True, pinned=True),
        "source_file": st.column_config.TextColumn("Source file", disabled=True),
        "po_received_date": st.column_config.TextColumn(
            core.HEADERS["po_received_date"], help="YYYY-MM-DD or DD/MM/YYYY"),
        "shipment_date": st.column_config.TextColumn(
            core.HEADERS["shipment_date"], help="YYYY-MM-DD or DD/MM/YYYY"),
    }
    if combined:
        config["carton_type"] = st.column_config.TextColumn("Carton type", disabled=True, pinned=True)
    for col in core.COLUMNS:
        if col not in config:
            config[col] = (st.column_config.NumberColumn(core.HEADERS[col])
                           if col in core.INT_COLS | core.REAL_COLS
                           else st.column_config.TextColumn(core.HEADERS[col]))

    first = (selected_page - 1) * PAGE_SIZE + 1
    last = first + len(rows) - 1
    scope = "Standard and Custom" if combined else core.SHEETS[kind]
    with st.container(key="manage_panel_tracker_data"):
        st.caption(f"Editing rows {first:,}–{last:,} of {count:,} in {scope} for "
                   f"{core.supplier_name(con, supplier)}. "
                   "Dates accept YYYY-MM-DD or DD/MM/YYYY. Changes are saved together.")
        revision = core.data_revision(con)  # ticks and edits are kept by row position: reset them when rows change
        edited = st.data_editor(
            frame,
            key=f"tracker_editor_{kind}_{supplier}_{selected_page}_{revision}",
            hide_index=True,
            num_rows="fixed",
            height=560,
            column_config=config,
        )
        delete_tools(con, kind, supplier, names[supplier], edited, combined)
    if st.button("Save tracker changes", type="primary", icon=":material/save:",
                 key="manage_primary_tracker_save"):
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
        grouped = {}
        for item in changed:
            grouped.setdefault(item.get("carton_type", kind), []).append(item)
        try:
            with operation_status("Saving tracker changes…", "Tracker changes saved"):
                count = sum(core.update_tracker_rows(con, row_kind, items)
                            for row_kind, items in grouped.items())
                clear_built_reports()
        except ValueError as exc:
            st.error(str(exc))
        else:
            st.toast(f"Saved {count:,} tracker rows", icon=":material/check_circle:")
            st.rerun()


main()
