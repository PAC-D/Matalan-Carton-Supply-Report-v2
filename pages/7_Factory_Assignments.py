import streamlit as st

import core
from web_ui import header, kpi_grid, page

PAGE_SIZE = 50
ALL = "All"


def assignment_exchange(con):
    with st.expander("Import or export assignments (Excel)", icon=":material/sync_alt:"):
        left, right = st.columns(2, gap="large")
        with left:
            st.markdown("**Export**")
            st.caption("Download every factory and its current packaging-supplier selections.")
            st.download_button(
                "Download assignments",
                data=core.export_factory_assignments(con),
                file_name="Matalan Factory Assignments.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                icon=":material/download:",
            )
        with right:
            st.markdown("**Import**")
            upload = st.file_uploader(
                "Assignments workbook", type="xlsx", key="factory_assignment_upload")
            st.caption(
                "Use the exported Factory and Packaging Suppliers columns. "
                "Separate multiple suppliers with |, comma, or semicolon."
            )
            if upload is not None and st.button("Apply import", type="primary"):
                try:
                    count = core.import_factory_assignments(con, upload.getvalue())
                except ValueError as exc:
                    st.error(str(exc))
                else:
                    st.toast(f"Imported {count:,} factory assignments", icon=":material/check_circle:")
                    st.rerun()


def factory_row(con, factory, suppliers, widths):
    supplier_names = [supplier["name"] for supplier in suppliers]
    id_by_name = {supplier["name"]: supplier["id"] for supplier in suppliers}
    with st.container(key=f"manage_row_factory_{factory['id']}"):
        name_col, assignment_col, save_col = st.columns(widths, vertical_alignment="center")
        name_col.write(factory["name"])
        if not factory["assigned_ids"]:
            name_col.caption("Unassigned")
        selected = assignment_col.multiselect(
            "Packaging suppliers",
            supplier_names,
            default=factory["assigned_names"],
            key=f"factory_suppliers_{factory['id']}",
            placeholder="Select packaging suppliers",
            label_visibility="collapsed",
        )
        if save_col.button("Save", key=f"save_factory_{factory['id']}", width="stretch"):
            try:
                core.set_factory_assignments(
                    con, factory["id"], [id_by_name[name] for name in selected])
            except ValueError as exc:
                st.error(str(exc))
            else:
                message = (f"Saved {factory['name']}" if selected
                           else f"Cleared assignments for {factory['name']}")
                st.toast(message, icon=":material/check_circle:")
                st.rerun()


@page("Factory assignments")
def main():
    con = core.connect()
    factories = core.list_assignment_factories(con)
    suppliers = core.list_packaging_suppliers(con)
    assigned = sum(bool(factory["assigned_ids"]) for factory in factories)
    header(
        "Factory assignments",
        "Choose the packaging suppliers available to each factory.",
        "Assignment coverage",
        f"{assigned:,} of {len(factories):,}",
    )
    kpi_grid([
        {"label": "Factories", "value": f"{len(factories):,}"},
        {"label": "Assigned", "value": f"{assigned:,}"},
        {"label": "Unassigned", "value": f"{len(factories) - assigned:,}",
         "tone": "ok" if assigned == len(factories) else "warn"},
        {"label": "Packaging suppliers", "value": f"{len(suppliers):,}",
         "note": "Available to assign"},
    ])

    if not factories:
        st.info("No factories are available yet. Import tracker data first.")
        return
    if not suppliers:
        st.warning("Add a packaging supplier before assigning factories.")
        return

    assignment_exchange(con)

    with st.container(border=True, key="card_assignment_filters"):
        factory_col, supplier_col, status_col = st.columns([2, 2, 1])
        factory_filter = factory_col.selectbox(
            "Factory", [ALL] + [factory["name"] for factory in factories])
        supplier_filter = supplier_col.selectbox(
            "Packaging supplier", [ALL] + [supplier["name"] for supplier in suppliers])
        status_filter = status_col.selectbox("Status", [ALL, "Assigned", "Unassigned"])

    filtered = factories
    if factory_filter != ALL:
        filtered = [factory for factory in filtered if factory["name"] == factory_filter]
    if supplier_filter != ALL:
        filtered = [factory for factory in filtered
                    if supplier_filter in factory["assigned_names"]]
    if status_filter == "Assigned":
        filtered = [factory for factory in filtered if factory["assigned_ids"]]
    elif status_filter == "Unassigned":
        filtered = [factory for factory in filtered if not factory["assigned_ids"]]
    filtered = sorted(
        filtered, key=lambda factory: (bool(factory["assigned_ids"]), factory["name"].casefold()))

    if not filtered:
        st.info("No factories match the selected filters.")
        return

    page_count = max(1, (len(filtered) + PAGE_SIZE - 1) // PAGE_SIZE)
    list_bar = st.container(horizontal=True, vertical_alignment="center")
    status = list_bar.empty()
    selected_page = 1
    if page_count > 1:
        selected_page = list_bar.selectbox(
            "Page", range(1, page_count + 1),
            format_func=lambda value: f"Page {value} of {page_count}",
            label_visibility="collapsed", width=160)
    first = (selected_page - 1) * PAGE_SIZE
    last = min(first + PAGE_SIZE, len(filtered))
    status.caption(f"Showing {first + 1:,}–{last:,} of {len(filtered):,} factories")

    widths = [3, 5, 1]
    with st.container(border=True, key="card_assignment_list"):
        heading_name, heading_suppliers, _ = st.columns(widths)
        heading_name.markdown("**Factory**")
        heading_suppliers.markdown("**Packaging suppliers**")
        for factory in filtered[first:last]:
            factory_row(con, factory, suppliers, widths)


main()
