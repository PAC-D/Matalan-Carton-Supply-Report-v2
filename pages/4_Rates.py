import pandas as pd
import streamlit as st

import core
from web_ui import clear_built_reports, operation_status, page, report_page_header

RATE_COLS = ["carton_type", "effective_from", "pacd_pct", "matalan_pct"]


@page
def main():
    con = core.connect()
    rates = core.get_rates(con)
    report_page_header("Rates", "Rate<br>rows", f"{len(rates):,}")
    with st.container(key="manage_panel_rates"):
        st.subheader("Rebate rates")
        st.caption("Each rate applies from its From date until the next rate for the same carton type.")
        frame = pd.DataFrame(rates, columns=RATE_COLS)
        frame["effective_from"] = pd.to_datetime(frame["effective_from"]).dt.date
        frame[["pacd_pct", "matalan_pct"]] *= 100
        edited = st.data_editor(frame, key="rate_editor", num_rows="dynamic", hide_index=True,
                                column_config={
                                    "carton_type": st.column_config.SelectboxColumn(
                                        "Carton type", options=["STD", "CUS"], required=True),
                                    "effective_from": st.column_config.DateColumn("From", required=True),
                                    "pacd_pct": st.column_config.NumberColumn(
                                        "PACD %", min_value=0, max_value=100, format="%.2f"),
                                    "matalan_pct": st.column_config.NumberColumn(
                                        "Matalan rebate %", min_value=0, max_value=100, format="%.2f"),
                                })
        if st.button("Save rates", type="primary", icon=":material/save:",
                     key="manage_primary_rates_save"):
            try:
                with operation_status("Saving rebate rates…", "Rebate rates saved"):
                    core.save_rates(con, [{
                        "carton_type": row.carton_type,
                        "effective_from": str(row.effective_from)[:10],
                        "pacd_pct": round(row.pacd_pct / 100, 6),
                        "matalan_pct": round(row.matalan_pct / 100, 6),
                    } for row in edited.itertuples()])
                    clear_built_reports()
            except (ValueError, TypeError) as exc:
                st.error(str(exc))
            else:
                st.toast("Rates saved", icon=":material/check_circle:")
    carton_prices(con)


PRICE_COLS = ["carton_type", "carton_selection", "effective_from", "price", "per"]


def carton_prices(con):
    with st.container(key="manage_panel_prices"):
        st.subheader("Carton prices")
        st.caption("USD prices used to fill blank price cells of tracker rows. A price applies from its From date "
                   "until the next price for the same carton type and Carton Selection. "
                   "Per carton: the carton's unit price. Per sqm: multiplied by the carton's Area(sqm).")
        frame = pd.DataFrame(core.get_prices(con), columns=PRICE_COLS)
        frame["effective_from"] = pd.to_datetime(frame["effective_from"]).dt.date
        edited = st.data_editor(frame, key="price_editor", num_rows="dynamic", hide_index=True,
                                column_config={
                                    "carton_type": st.column_config.SelectboxColumn(
                                        "Carton type", options=["STD", "CUS"], required=True),
                                    "carton_selection": st.column_config.TextColumn(
                                        "Carton Selection", required=True),
                                    "effective_from": st.column_config.DateColumn("From", required=True),
                                    "price": st.column_config.NumberColumn(
                                        "Price (USD)", min_value=0, format="%.4f"),
                                    "per": st.column_config.SelectboxColumn(
                                        "Per", options=["carton", "sqm"], required=True),
                                })
        if st.button("Save carton prices", type="primary", icon=":material/save:",
                     key="manage_primary_prices_save"):
            try:
                with operation_status("Saving carton prices…", "Carton prices saved"):
                    filled = core.save_prices(con, [{
                        "carton_type": row.carton_type,
                        "carton_selection": row.carton_selection,
                        "effective_from": str(row.effective_from)[:10],
                        "price": None if pd.isna(row.price) else float(row.price),
                        "per": row.per,
                    } for row in edited.itertuples()])
                    clear_built_reports()
            except (ValueError, TypeError) as exc:
                st.error(str(exc))
            else:
                st.toast(f"Carton prices saved; {filled:,} tracker row(s) filled",
                         icon=":material/check_circle:")


main()
