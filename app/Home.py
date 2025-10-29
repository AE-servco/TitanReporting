import streamlit as st
from datetime import datetime, time, date, timedelta
from zoneinfo import ZoneInfo

from modules.data import get_invoices_for_xero, convert_df_for_download


st.set_page_config(
    layout="wide",
    page_title="ServiceTitan Reports"
    )

st.title(f"ServiceTitan Reports")

time_now = datetime.now(ZoneInfo("Australia/Sydney"))
today = time_now.date()
yesterday = today - timedelta(days=1)

date_range = st.sidebar.date_input(
    "Date filter:",
    (yesterday, today),
    date(2025,1,1),
    today,
    format="DD/MM/YYYY",
)

state = st.sidebar.radio(
    "State:",
    [
        "NSW", 
        "WA",
    ]
)

if "confirmed_range" not in st.session_state:
    st.session_state.confirmed_range = None

if len(date_range) == 2:
    st.session_state.confirmed_range = tuple(date_range)

start_date = st.session_state.confirmed_range[0]
end_date = st.session_state.confirmed_range[1]

if "invoice_data" not in st.session_state:
    st.session_state.invoice_data = None
    st.session_state.start_date = "NO_DATE_SELECTED"
    st.session_state.end_date = None

if st.button("Fetch Invoice Data", key="invoice_data_button"):
    st.session_state.invoice_data = get_invoices_for_xero(state, start_date, end_date)
    st.session_state.start_date = start_date
    st.session_state.end_date = end_date

st.download_button(
    label="Download Invoices",
    data=convert_df_for_download(st.session_state.invoice_data),
    file_name=f"invoices_{st.session_state.start_date}-{st.session_state.end_date}.csv",
    mime="text/csv",
    icon=":material/download:",
)

st.dataframe(st.session_state.invoice_data)