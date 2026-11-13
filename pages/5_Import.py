from io import BytesIO

import pandas as pd
import streamlit as st

import core
from web_ui import clear_built_reports, operation_status, page, report_page_header

DATE_EDITS = {
    "PO Received Date": "PO Received Date to import",
    "Shipment Date": "Shipment Date to import",
}


def read_cached_upload(con, content, kind):
    """Parse once and keep the validated result in the persistent PostgreSQL cache."""
    namespace = f"upload-{kind}-v1"
    key = core.cache_key(namespace, content)
    cached = core.cache_get(con, key)
    if cached is not None:
        return cached, True
    parsed = (core.read_workbook(BytesIO(content)) if kind == "tracker"
              else core.read_ebiz(BytesIO(content)))
    core.cache_set(con, key, namespace, parsed, max_entries=20)
    return parsed, False


def as_date(value):
    return None if pd.isna(value) else pd.Timestamp(value).date()


def _read_uploads(con, uploads, kind):
    files, failures = {}, 0
    if not uploads:
        return files
    label = "tracker" if kind == "tracker" else "eBiz"
    total = len(uploads)
    with st.status(
        f"Validating {total} {label} file{'s' if total != 1 else ''}…",
        expanded=True,
        state="running",
        type="compact",
    ) as validation:
        progress = st.progress(0, text="Preparing validation…")
        for index, upload in enumerate(uploads, 1):
            progress.progress(
                (index - 1) / total,
                text=f"Checking {upload.name} · file {index} of {total}",
            )
            try:
                files[upload.name], cached = read_cached_upload(con, upload.getvalue(), kind)
                st.write(
                    f":material/check_circle: **{upload.name}** — "
                    f"{'restored from cache' if cached else 'validated successfully'}"
                )
            except core.ImportFileError as exc:
                failures += 1
                st.error(f"{upload.name}: {exc} This file will not be imported.")
            progress.progress(index / total, text=f"Checked {index} of {total} files")
        progress.empty()
        if failures:
            validation.update(
                label=f"Validation finished · {len(files)} ready · {failures} failed",
                state="error",
                expanded=True,
            )
        else:
            validation.update(
                label=f"{len(files)} {label} file{'s' if len(files) != 1 else ''} ready",
                state="complete",
                expanded=False,
            )
    return files


def _highlight(table):
    """Red cells: a value that differs from the calculation, or a blank that can't be calculated."""
    style = pd.DataFrame("", index=table.index, columns=table.columns)
    style["Value"] = "background-color: rgba(220, 60, 40, 0.22)"
    return style


def calc_table(issues, message, first, last=None):
    """Calculation check: saved values that differ from the workbook formulas, kept as they are."""
    if not issues:
        return
    st.warning(f"{len(issues):,} {message}")
    table = pd.DataFrame(issues)
    order = first + ["Column", "Value", "Calculated", "Problem", "Carton Selection",
                     "Packaging supplier", "eBiz Sales Order #", last]
    st.dataframe(table[[c for c in order if c in table]].style.apply(_highlight, axis=None)
                 .format({"Value": "{:,.4f}", "Calculated": "{:,.4f}"}, na_rep="—"), hide_index=True)


def saved_calc_check(con):
    issues = core.saved_calc_issues(con)
    if issues:
        with st.expander(f"Calculation check of saved tracker rows ({len(issues):,})", icon=":material/calculate:"):
            calc_table(issues, "saved cell(s) differ from the calculation or can't be calculated. "
                       "Edit them on the Tracker page or in the next import.", ["Sheet", "Row ID"])


def tracker_import(con, uploads):
    if not uploads:
        return
    with st.container(key="manage_panel_tracker_import"):
        st.subheader("Tracker validation")
        st.caption("Files are checked before any Tracker rows are changed.")
        files = _read_uploads(con, uploads, "tracker")
        if files:
            reused, issues = core.check_dates(con, files)
            if reused:
                st.info(f"{reused} row(s) used a previously saved date correction.")
            if issues:
                st.warning(f"{len(issues)} row(s) have dates that need review. Edit the proposed values below.")
                table = pd.DataFrame([{k: v for k, v in item.items() if k != "row"}
                                      for item in issues])
                for col in DATE_EDITS:
                    table[col] = pd.to_datetime(table[col])
                edited = st.data_editor(
                    table, key="date_review_editor", hide_index=True,
                    disabled=[c for c in table.columns if c not in DATE_EDITS],
                    column_order=["Sheet", "Excel row", "PO Received Date in file",
                                  "PO Received Date", "Shipment Date in file", "Shipment Date",
                                  "Problem", "Packaging supplier", "eBiz Sales Order #",
                                  "Carton QTY", "File"],
                    column_config={
                        "Problem": st.column_config.TextColumn(width="large"),
                        **{col: st.column_config.DateColumn(label, format="DD-MMM-YYYY")
                           for col, label in DATE_EDITS.items()},
                    })
                core.apply_dates(issues, zip(map(as_date, edited["PO Received Date"]),
                                             map(as_date, edited["Shipment Date"])))

            calc_table(core.calc_check(con, files),
                       "cell(s) in the uploaded files differ from the calculation or can't be calculated. "
                       "Blank cells that can be calculated are filled in on import.",
                       ["Sheet", "Excel row"], "File")

            st.dataframe(pd.DataFrame([{"File": name, **row}
                                       for name, data in files.items()
                                       for row in core.preview(con, data)]), hide_index=True)
            clash = core.clashes(files)
            if clash:
                st.error("These files contain the same sheet, supplier and month. Remove one file.")
                st.dataframe(pd.DataFrame(clash), hide_index=True)
            if st.button("Confirm tracker import", type="primary", disabled=bool(clash),
                         icon=":material/upload:", key="manage_primary_tracker_import"):
                try:
                    with operation_status("Importing tracker data…", "Tracker import complete"):
                        counts = core.save_many(con, files)
                        total = sum(sum(c.values()) for c in counts.values())
                        clear_built_reports()
                except Exception as exc:
                    st.error(f"Tracker import failed. No partial changes were saved. {exc}")
                else:
                    st.success(f"Import finished: {total:,} tracker rows are saved.")
                    st.toast(f"Imported {total:,} tracker rows", icon=":material/check_circle:")
                    if core.unrated_count(con):
                        st.warning("Some shipments are earlier than the first rate. Add an earlier rate.")


def saved_date_corrections(con):
    saved = core.saved_dates(con)
    if saved:
        with st.expander(f"Saved date corrections ({len(saved)})", icon=":material/event_repeat:"):
            st.caption("These corrections are reused when the same source row is imported again.")
            picked = st.dataframe(pd.DataFrame(saved).drop(columns="key"), hide_index=True,
                                  on_select="rerun", selection_mode="multi-row",
                                  key=f"saved_dates_{hash(tuple(item['key'] for item in saved))}")
            if st.button("Forget selected corrections", disabled=not picked.selection.rows,
                         icon=":material/delete:"):
                with operation_status("Removing date corrections…", "Date corrections removed"):
                    core.forget_dates(con, [saved[i]["key"] for i in picked.selection.rows])
                st.rerun()


def ebiz_import(con, uploads):
    if not uploads:
        return
    with st.container(key="manage_panel_ebiz_import"):
        st.subheader("eBiz validation")
        st.caption("Files are checked before any vendor rows are replaced.")
        files = _read_uploads(con, uploads, "ebiz")
        if not files:
            return
        st.dataframe(pd.DataFrame(core.ebiz_preview(files)), hide_index=True)
        uploaded = {row["vendor_key"] for rows in files.values() for row in rows}
        unknown = sorted(uploaded - set(core.suppliers(con)), key=str)
        if unknown:
            st.warning(f"No tracker rows exist for: {', '.join(map(str, unknown))}.")
        clash = core.ebiz_clashes(files)
        if clash:
            st.error("These files contain the same vendor. Remove one file.")
            st.dataframe(pd.DataFrame(clash), hide_index=True)
        st.caption("Only the uploaded vendors will be replaced. All other vendors will be kept unchanged.")
        if st.button("Update eBiz data", type="primary", disabled=bool(clash),
                     icon=":material/sync:", key="manage_primary_ebiz_import"):
            try:
                with operation_status("Updating eBiz data…", "eBiz update complete"):
                    counts = core.replace_ebiz(con, files)
                    total = sum(counts.values())
                    clear_built_reports()
            except Exception as exc:
                st.error(f"eBiz import failed. Existing eBiz data is unchanged. {exc}")
            else:
                st.session_state["import_flash"] = (
                    f"Update finished: {total:,} uploaded eBiz rows are saved.")
                st.rerun()


def upload_sources(con):
    """Render compact upload choices; validation and database changes stay separate."""
    ebiz_status = core.ebiz_status(con)
    tracker_files = core.tracker_files(con)
    tracker_rows = sum(item["STD"] + item["CUS"] for item in tracker_files)
    tracker_loaded = max(
        (item["imported_at"] or "" for item in tracker_files), default=""
    ).replace("T", " ")[:16]
    with st.container(key="manage_import_sources"):
        tracker_column, ebiz_column = st.columns(2, gap="small")
        with tracker_column:
            st.subheader("Tracker workbooks")
            if tracker_files:
                st.caption(
                    f"Current data: {tracker_rows:,} rows from {len(tracker_files)} file(s) · "
                    f"Last loaded: {tracker_loaded or 'Unknown'}"
                )
            else:
                st.caption("No Tracker data loaded · XLSX or XLSM · multiple files allowed")
            tracker_uploads = st.file_uploader(
                "Tracker workbooks",
                type=["xlsx", "xlsm"],
                accept_multiple_files=True,
                key="tracker_upload",
                label_visibility="collapsed",
            )
        with ebiz_column:
            st.subheader("eBiz exports")
            if ebiz_status:
                st.caption(
                    f"Current data: {ebiz_status[0]:,} rows from {len(ebiz_status[1])} file(s) · "
                    f"Last loaded: {ebiz_status[2].replace('T', ' ')[:16]}"
                )
            else:
                st.caption("No eBiz data loaded · XLSX or XLSM · multiple files allowed")
            ebiz_uploads = st.file_uploader(
                "eBiz exports",
                type=["xlsx", "xlsm"],
                accept_multiple_files=True,
                key="ebiz_upload",
                label_visibility="collapsed",
            )
    return tracker_uploads, ebiz_uploads


@page
def main():
    con = core.connect()
    tracker_count = sum(core.editable_row_count(con, kind) for kind in core.TABLES)
    report_page_header("Data import", "Tracker<br>rows", f"{tracker_count:,}")
    if message := st.session_state.pop("import_flash", None):
        st.success(message, icon=":material/check_circle:")
        st.toast(message, icon=":material/check_circle:")
    tracker_uploads, ebiz_uploads = upload_sources(con)
    tracker_import(con, tracker_uploads)
    saved_date_corrections(con)
    saved_calc_check(con)
    ebiz_import(con, ebiz_uploads)


main()
