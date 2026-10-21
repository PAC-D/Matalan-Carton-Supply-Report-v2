import streamlit as st

import core
from web_ui import header, page


def supplier_row(con, supplier):
    supplier_id = supplier["id"]
    with st.container(key=f"manage_row_supplier_{supplier_id}"):
        row = st.container(horizontal=True, vertical_alignment="center")
        new_name = row.text_input(
            "Supplier name",
            value=supplier["name"],
            key=f"supplier_name_{supplier_id}",
            label_visibility="collapsed",
        )
        changed = bool(new_name.strip() and new_name.strip() != supplier["name"])
        if row.button("Save", key=f"save_supplier_{supplier_id}", disabled=not changed):
            try:
                core.rename_packaging_supplier(con, supplier_id, new_name)
            except ValueError as exc:
                st.error(str(exc))
            else:
                st.toast(f"Renamed to {new_name.strip()}", icon=":material/check_circle:")
                st.rerun()
        if row.button(
            "Delete", key=f"delete_supplier_{supplier_id}", type="tertiary",
            icon=":material/delete:",
        ):
            st.session_state[f"confirm_supplier_{supplier_id}"] = True

        if st.session_state.get(f"confirm_supplier_{supplier_id}"):
            confirm = st.container(horizontal=True, vertical_alignment="center")
            confirm.markdown(f"Delete **{supplier['name']}**? This cannot be undone.")
            if confirm.button("Delete", key=f"confirm_delete_supplier_{supplier_id}", type="primary"):
                try:
                    core.delete_packaging_supplier(con, supplier_id)
                except ValueError as exc:
                    st.error(str(exc))
                else:
                    st.session_state.pop(f"confirm_supplier_{supplier_id}", None)
                    st.toast(f"Deleted {supplier['name']}", icon=":material/delete:")
                    st.rerun()
            if confirm.button("Cancel", key=f"cancel_delete_supplier_{supplier_id}"):
                st.session_state.pop(f"confirm_supplier_{supplier_id}", None)
                st.rerun()


@page("Packaging suppliers")
def main():
    con = core.connect()
    suppliers = core.list_packaging_suppliers(con)
    header(
        "Packaging suppliers",
        "Maintain the packaging suppliers available for factory assignments.",
        "Available suppliers",
        f"{len(suppliers):,}",
    )

    with st.container(border=True, key="card_add_supplier"):
        st.subheader("Add packaging supplier")
        with st.form("add_packaging_supplier", clear_on_submit=True, border=False):
            row = st.container(horizontal=True, vertical_alignment="bottom")
            name = row.text_input(
                "Supplier name", placeholder="e.g. Reflex Packaging Ltd.")
            submitted = row.form_submit_button(
                "Add supplier", type="primary", icon=":material/add:")
        if submitted:
            try:
                core.add_packaging_supplier(con, name)
            except ValueError as exc:
                st.error(str(exc))
            else:
                st.toast(f"Added {name.strip()}", icon=":material/check_circle:")
                st.rerun()

    suppliers = core.list_packaging_suppliers(con)
    with st.container(border=True, key="card_supplier_list"):
        st.subheader("Suppliers")
        if not suppliers:
            st.info("No packaging suppliers yet. Add the first one above.")
            return
        st.caption(
            f"{len(suppliers):,} supplier{'s' if len(suppliers) != 1 else ''}. "
            "Edit a name and select Save to rename it. A supplier can only be "
            "deleted when no factory is assigned to it."
        )
        for supplier in suppliers:
            supplier_row(con, supplier)


main()
