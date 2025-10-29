import streamlit as st
from streamlit import session_state as ss
import streamlit_authenticator as stauth
from datetime import datetime, time, date, timedelta
from zoneinfo import ZoneInfo

from modules.data import get_invoices_for_xero, convert_df_for_download
import modules.google_store as gs

st.set_page_config(
    layout="wide",
    page_title="ServiceTitan Reports"
    )

st.title(f"Plumber Commissions")

CONFIG_FILENAME = 'st_auth_config.yaml'
config = gs.load_yaml_from_gcs(CONFIG_FILENAME)

authenticator = stauth.Authenticate(
    credentials = config['credentials']
)

authenticator.login(location='main')

if ss["authentication_status"]:

    st.subheader("Work in progress")
    # time_now = datetime.now(ZoneInfo("Australia/Sydney"))
    # today = time_now.date()
    # yesterday = today - timedelta(days=1)
    # with st.container(horizontal=True):
    #     with st.container(width=200):
    #         date_range = st.date_input(
    #             "Date filter:",
    #             (yesterday, today),
    #             date(2025,1,1),
    #             today,
    #             format="DD/MM/YYYY",
    #         )
    #     with st.container(width=200):
    #         state = st.selectbox(
    #             "State:",
    #             [
    #                 "NSW", 
    #                 "WA",
    #             ]
    #         )

    # if "confirmed_range" not in ss:
    #     ss.confirmed_range = None

    # if len(date_range) == 2:
    #     ss.confirmed_range = tuple(date_range)

    # start_date = ss.confirmed_range[0]
    # end_date = ss.confirmed_range[1]

    # if "invoice_data" not in ss:
    #     ss.invoice_data = None
    #     ss.start_date = "NO_DATE_SELECTED"
    #     ss.end_date = None

    # with st.container(horizontal=True):
    #     if st.button("Fetch Invoice Data", key="invoice_data_button"):
    #         ss.invoice_data = get_invoices_for_xero(state, start_date, end_date)
    #         ss.start_date = start_date
    #         ss.end_date = end_date

    #     st.download_button(
    #         label="Download Invoices",
    #         data=convert_df_for_download(ss.invoice_data),
    #         file_name=f"invoices_{ss.start_date}-{ss.end_date}.csv",
    #         mime="text/csv",
    #         icon=":material/download:",
    #     )

    #     if st.button("Clear Data", key="clear_data_button"):
    #         ss.invoice_data = None
    #         ss.start_date = "NO_DATE_SELECTED"
    #         ss.end_date = None
    # st.write("Current Data:")
    # st.dataframe(ss.invoice_data)

elif ss["authentication_status"] is False:
    st.error('Go to home page to log in.')
elif ss["authentication_status"] is None:
    st.warning('Go to home page to log in.')