import pandas as pd
import streamlit as st

import core
from web_ui import clear_built_reports, header, page

PAGE_SIZE = 250
ALL = "All vendors"


def same_value(left, right):
    """Treat the different empty values produced by pandas as equivalent."""
    left_missing, right_missing = bool(pd.isna(left)), bool(pd.isna(right))
    return left_missing and right_missing if left_missing or right_missing else left == right


def cancel_delete():
    st.session_state.pop("ebiz_delete_pending", None)


@st.dialog("Delete eBiz rows", on_dismiss=cancel_delete)
def confirm_delete(ids):
    """Ask once more; the pending ids live in session state until used or dismissed."""
    st.write(f"Delete {len(ids):,} eBiz rows?")
    st.caption("This cannot be undone here. RDS keeps 7 days of backups. Re-importing the vendor exports "
               "replaces all eBiz data.")
    if st.button("Delete", type="primary", icon=":material/delete:"):
        deleted = core.delete_ebiz_rows(core.connect(), ids)  # its own connection, like the tracker dialog
        clear_built_reports()
        cancel_delete()
        st.session_state["ebiz_deleted"] = deleted
        st.rerun()


@page
def main():
    con = core.connect()
    deleted = st.session_state.pop("ebiz_deleted", None)
    if deleted is not None:
        st.toast(f"Deleted {deleted:,} eBiz rows", icon=":material/check_circle:")
    total = core.ebiz_row_count(con)
    header("eBiz data", "Review, correct or remove the imported eBiz vendor rows. Reconciliation updates immediately.",
           "Stored rows", f"{total:,}")
    if not total:
        st.info("No eBiz data yet. Import vendor exports on the Data import page.")
        return

    vendors = core.ebiz_vendors(con)
    names = {key: core.supplier_name(con, key) for key in vendors}
    with st.container(border=True, key="card_ebiz_filters"):
        vendor = st.selectbox("Vendor", [ALL, *vendors], format_func=lambda key: names.get(key, key))
    vendor = None if vendor == ALL else vendor

    count = core.ebiz_row_count(con, vendor)
    pages = max(1, (count + PAGE_SIZE - 1) // PAGE_SIZE)
    selected_page = 1
    if pages > 1:
        selected_page = st.selectbox("Page", range(1, pages + 1),
                                     format_func=lambda number: f"Page {number} of {pages}", width=180)
    rows = core.ebiz_rows(con, vendor, PAGE_SIZE, (selected_page - 1) * PAGE_SIZE)

    frame = pd.DataFrame(rows)[["id", *core.EBIZ_HEADERS, "source_file", "imported_at"]]
    frame.insert(0, "delete", False)
    config = {
        "delete": st.column_config.CheckboxColumn("Delete", default=False, pinned=True),
        "id": st.column_config.NumberColumn("Row ID", disabled=True, pinned=True),
        "source_file": st.column_config.TextColumn("Source file", disabled=True),
        "imported_at": st.column_config.TextColumn("Imported", disabled=True),
        **{col: (st.column_config.NumberColumn(label) if col in core.EBIZ_REAL
                 else st.column_config.TextColumn(label))
           for col, label in core.EBIZ_HEADERS.items()},
    }
    first = (selected_page - 1) * PAGE_SIZE + 1
    st.caption(f"Editing rows {first:,}–{first + len(rows) - 1:,} of {count:,}. SO_1 and SO_2 are taken from "
               "Order Number (\"A|B\") when you save.")
    revision = core.data_revision(con)  # ticks and edits are kept by row position: reset them when rows change
    edited = st.data_editor(frame, key=f"ebiz_editor_{vendor}_{selected_page}_{revision}", hide_index=True,
                            num_rows="fixed", height=560, column_config=config)

    ids = [int(edited.at[i, "id"]) for i in edited.index if edited.at[i, "delete"]]
    with st.container(horizontal=True):
        save = st.button("Save eBiz changes", type="primary", icon=":material/save:")
        if st.button(f"Delete {len(ids):,} selected rows", disabled=not ids, icon=":material/delete:"):
            st.session_state["ebiz_delete_pending"] = ids
    pending = st.session_state.get("ebiz_delete_pending")
    if pending:
        confirm_delete(pending)
    if save:
        changed = [edited.loc[i].to_dict() for i in frame.index
                   if any(not same_value(frame.at[i, col], edited.at[i, col]) for col in core.EBIZ_HEADERS)]
        if not changed:
            st.info("No eBiz changes to save.")
            return
        try:
            saved = core.update_ebiz_rows(con, changed)
        except ValueError as exc:
            st.error(str(exc))
        else:
            clear_built_reports()
            st.toast(f"Saved {saved:,} eBiz rows", icon=":material/check_circle:")
            st.rerun()


main()
