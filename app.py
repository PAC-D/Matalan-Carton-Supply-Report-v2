"""Matalan Carton Supply Report website entry point."""
import streamlit as st

from web_ui import APP_NAME, FAVICON, PACD_LOGO, app_chrome

st.set_page_config(page_title=APP_NAME, page_icon=FAVICON, layout="wide")
st.logo(PACD_LOGO, size="medium", icon_image=FAVICON)
app_chrome()

pages = {
    "Overview": [
        st.Page("pages/1_Dashboard.py", title="Dashboard",
                icon=":material/space_dashboard:", default=True),
    ],
    "Reports": [
        st.Page("pages/2_Reports.py", title="Build reports",
                icon=":material/description:", url_path="reports"),
        st.Page("pages/6_Cross_Reference.py", title="Cross reference",
                icon=":material/compare_arrows:", url_path="cross-reference"),
        st.Page("pages/7_Discrepancy_Report.py", title="Discrepancy report",
                icon=":material/difference:", url_path="discrepancy-report"),
        st.Page("pages/8_Rebate_Report.py", title="Rebate report",
                icon=":material/payments:", url_path="rebate-report"),
    ],
    "Manage": [
        st.Page("pages/3_Tracker.py", title="Tracker data",
                icon=":material/edit_note:", url_path="tracker"),
        st.Page("pages/4_Rates.py", title="Rates",
                icon=":material/percent:", url_path="rates"),
        st.Page("pages/5_Import.py", title="Data import",
                icon=":material/upload_file:", url_path="import"),
    ],
}

st.navigation(pages).run()
