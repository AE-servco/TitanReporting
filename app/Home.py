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

st.title(f"ServiceTitan Reports")

CONFIG_FILENAME = 'st_auth_config.yaml'
config = gs.load_yaml_from_gcs(CONFIG_FILENAME)

authenticator = stauth.Authenticate(
    credentials = config['credentials']
)

authenticator.login(location='main')

if ss["authentication_status"]:
    st.write("Welcome! Please use the navigation on the left to find the reports you need.")

elif ss["authentication_status"] is False:
    st.error('Username/password is incorrect')
elif ss["authentication_status"] is None:
    st.warning('Please enter your username and password')