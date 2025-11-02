import streamlit as st
from streamlit import session_state as ss
import streamlit_authenticator as stauth
from datetime import datetime, time, date, timedelta
from zoneinfo import ZoneInfo

from modules.data import convert_df_for_download
import modules.google_store as gs

st.set_page_config(
    layout="wide",
    page_title="ServiceTitan Reports"
    )

st.title(f"Doc Check Checker")

CONFIG_FILENAME = 'st_auth_config.yaml'
config = gs.load_yaml_from_gcs(CONFIG_FILENAME)

authenticator = stauth.Authenticate(
    credentials = config['credentials']
)

authenticator.login(location='main')

if ss["authentication_status"]:

    st.subheader("Work in progress")
    # time_now = datetime.now(ZoneInfo("Australia/Sydney"))
    # with st.container(width=200):
    #     state = st.selectbox(
    #         "State:",
    #         [
    #             "NSW", 
    #             "WA",
    #             "VIC", 
    #             "QLD",
    #         ]
    #     )

    # if "doc_check_data" not in ss:
    #     ss.doc_check_data = None

    # with st.container(horizontal=True):
    #     if st.button("Fetch Commission Data", key="commission_data_button"):
    #         ss.doc_check_data = get_doc_check_checker_data(state)

    #     st.download_button(
    #         label="Download Doc Check Data",
    #         data=convert_df_for_download(ss.doc_check_data),
    #         file_name=f"doc_check_data.csv",
    #         mime="text/csv",
    #         icon=":material/download:",
    #     )

    #     if st.button("Clear Data", key="clear_data_button"):
    #         ss.doc_check_data = None
    # st.write("Current Data:")
    # st.dataframe(ss.doc_check_data)


elif ss["authentication_status"] is False:
    st.error('Go to home page to log in.')
elif ss["authentication_status"] is None:
    st.warning('Go to home page to log in.')