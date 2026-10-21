import pandas as pd
import streamlit as st

import core
from web_ui import header, page

RATE_COLS = ["carton_type", "effective_from", "pacd_pct", "matalan_pct"]


@page("Rates")
def main():
    con = core.connect()
    rates = core.get_rates(con)
    header("Rates", "Maintain date-effective PACD and Matalan rebate percentages.",
           "Rate rows", str(len(rates)))
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
    if st.button("Save rates", type="primary", icon=":material/save:"):
        try:
            core.save_rates(con, [{
                "carton_type": row.carton_type,
                "effective_from": str(row.effective_from)[:10],
                "pacd_pct": round(row.pacd_pct / 100, 6),
                "matalan_pct": round(row.matalan_pct / 100, 6),
            } for row in edited.itertuples()])
        except (ValueError, TypeError) as exc:
            st.error(str(exc))
        else:
            st.session_state.pop("built_reports", None)
            st.toast("Rates saved", icon=":material/check_circle:")


main()
