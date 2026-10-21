from io import BytesIO

import pandas as pd
import streamlit as st

import core
from web_ui import header, page

DATE_EDITS = {
    "PO Received Date": "PO Received Date to import",
    "Shipment Date": "Shipment Date to import",
}


@st.cache_data(max_entries=20, show_spinner="Reading tracker workbook…")
def read_upload(content):
    return core.read_workbook(BytesIO(content))


@st.cache_data(max_entries=20, show_spinner="Reading eBiz export…")
def read_ebiz_upload(content):
    return core.read_ebiz(BytesIO(content))


def as_date(value):
    return None if pd.isna(value) else pd.Timestamp(value).date()


def tracker_import(con):
    with st.container(border=True, key="card_tracker_import"):
        st.subheader("Tracker workbooks")
        st.caption("Upload one or more .xlsx or .xlsm files containing TRACK_STD and TRACK_CUS.")
        uploads = st.file_uploader("Tracker workbooks", type=["xlsx", "xlsm"],
                                   accept_multiple_files=True, key="tracker_upload",
                                   label_visibility="collapsed")
        files = {}
        for upload in uploads or []:
            try:
                files[upload.name] = read_upload(upload.getvalue())
            except core.ImportFileError as exc:
                st.error(f"{upload.name}: {exc} This file will not be imported.")
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

            st.dataframe(pd.DataFrame([{"File": name, **row}
                                       for name, data in files.items()
                                       for row in core.preview(con, data)]), hide_index=True)
            clash = core.clashes(files)
            if clash:
                st.error("These files contain the same sheet, supplier and month. Remove one file.")
                st.dataframe(pd.DataFrame(clash), hide_index=True)
            if st.button("Confirm tracker import", type="primary", disabled=bool(clash),
                         icon=":material/upload:"):
                counts = core.save_many(con, files)
                st.session_state.pop("built_reports", None)
                total = sum(sum(c.values()) for c in counts.values())
                st.toast(f"Imported {total:,} tracker rows", icon=":material/check_circle:")
                if core.unrated_count(con):
                    st.warning("Some shipments are earlier than the first rate. Add an earlier rate.")

    saved = core.saved_dates(con)
    if saved:
        with st.expander(f"Saved date corrections ({len(saved)})", icon=":material/event_repeat:"):
            st.caption("These corrections are reused when the same source row is imported again.")
            picked = st.dataframe(pd.DataFrame(saved).drop(columns="key"), hide_index=True,
                                  on_select="rerun", selection_mode="multi-row",
                                  key=f"saved_dates_{hash(tuple(item['key'] for item in saved))}")
            if st.button("Forget selected corrections", disabled=not picked.selection.rows,
                         icon=":material/delete:"):
                core.forget_dates(con, [saved[i]["key"] for i in picked.selection.rows])
                st.rerun()


def ebiz_import(con):
    status = core.ebiz_status(con)
    with st.container(border=True, key="card_ebiz_import"):
        st.subheader("eBiz exports")
        if status:
            st.caption(f"Current data: {status[0]:,} rows from {len(status[1])} file(s), loaded "
                       f"{status[2].replace('T', ' ')[:16]}.")
        else:
            st.caption("No eBiz data yet. Upload every vendor together; replacement is all-or-nothing.")
        uploads = st.file_uploader("eBiz exports", type=["xlsx", "xlsm"],
                                   accept_multiple_files=True, key="ebiz_upload",
                                   label_visibility="collapsed")
        files = {}
        for upload in uploads or []:
            try:
                files[upload.name] = read_ebiz_upload(upload.getvalue())
            except core.ImportFileError as exc:
                st.error(f"{upload.name}: {exc} This file will not be imported.")
        if not files:
            return
        st.dataframe(pd.DataFrame(core.ebiz_preview(files)), hide_index=True)
        uploaded = {row["vendor_key"] for rows in files.values() for row in rows}
        unknown = sorted(uploaded - set(core.suppliers(con)), key=str)
        if unknown:
            st.warning(f"No tracker rows exist for: {', '.join(map(str, unknown))}.")
        dropped = sorted(set(core.ebiz_vendors(con)) - uploaded)
        if dropped:
            st.warning(f"Current eBiz data will be removed for: {', '.join(dropped)}.")
        clash = core.ebiz_clashes(files)
        if clash:
            st.error("These files contain the same vendor. Remove one file.")
            st.dataframe(pd.DataFrame(clash), hide_index=True)
        st.caption(f"Replacement deletes all {status[0] if status else 0:,} current eBiz rows first.")
        if st.button("Replace all eBiz data", type="primary", disabled=bool(clash),
                     icon=":material/sync:"):
            counts = core.replace_ebiz(con, files)
            st.session_state.pop("built_reports", None)
            st.toast(f"Loaded {sum(counts.values()):,} eBiz rows", icon=":material/check_circle:")
            st.rerun()


@page("Data import")
def main():
    con = core.connect()
    tracker_count = sum(con.execute(f"SELECT COUNT(*) FROM {table}").fetchone()[0]
                        for table in core.TABLES.values())
    header("Data import", "Validate and import tracker workbooks and eBiz vendor exports.",
           "Tracker rows", f"{tracker_count:,}")
    tracker_import(con)
    ebiz_import(con)


main()
