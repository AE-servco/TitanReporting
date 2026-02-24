from __future__ import annotations

from datetime import datetime, date, timedelta
from typing import Dict, List, Set, Tuple, Optional, Any, Iterable
import json
from google.cloud import secretmanager
import time
from zoneinfo import ZoneInfo

import streamlit as st
import streamlit_authenticator as stauth
from concurrent.futures import ThreadPoolExecutor, Future, as_completed

from servicetitan_api_client import ServiceTitanClient
import modules.google_store as gs
import modules.helpers as helpers
import modules.templates as templates
import modules.fetching as fetch
import modules.tasks as tasks

###############################################################################
# Filter warnings
###############################################################################
import warnings
warnings.filterwarnings("ignore", message=".*cookie_manager.*")

###############################################################################
# Configuration and helpers
###############################################################################

TENANTS = [
    "alphabravo"
    "foxtrotwhiskey", 
    "sierradelta",
    "victortango",
    "echozulu",
    "mikeecho",
    "bravogolf",
    "alphabravo", 
]

###############################################################################
# Streamlit app logic
###############################################################################

def main() -> None:

    st.set_page_config(page_title="Compliance Doc Checker Lite", layout="wide")
    st.markdown(
        """
        <style>
            div[data-testid="stColumn"] {
                height: 30% !important;
            }
            .stMainBlockContainer {
                padding-top: 2rem;
                padding-bottom: 0rem;
            }
        </style>
        """,
            unsafe_allow_html=True,
        )

    st.subheader("Doc Checker Lite")

    templates.authenticate_app('st_auth_config_payroll_doc_checker.yaml')

    if st.session_state["authentication_status"]:
        with st.spinner("Initialising..."):
            doc_check_criteria = helpers.get_doc_check_criteria()

            # Initialise session state collections
            if "jobs" not in st.session_state or ("curr_page" in st.session_state and st.session_state.curr_page != 'lite_checker'):
                st.session_state.jobs: List[Dict[str, Any]] = []
            if "curr_page" not in st.session_state or st.session_state.curr_page != 'lite_checker':
                st.session_state.curr_page = 'lite_checker'
            if "clients" not in st.session_state:
                st.session_state.clients: Dict[str, Any] = {}
            if "employee_lists" not in st.session_state:
                st.session_state.employee_lists: Dict[str, Any] = {}
            if "current_tenant" not in st.session_state:
                st.session_state.current_tenant: str = ""
            if "app_guid" not in st.session_state:
                st.session_state.app_guid = helpers.get_secret('st_servco_integrations_guid', project_id="prestigious-gcp")
            
        templates.filters_lite()

        # Display the current job if available
        if st.session_state.jobs:
            job = st.session_state.jobs[0]
            job_id = str(job.get("id"))
            job_num = str(job.get("jobNumber"))

            with st.container(horizontal_alignment="center", gap=None):
                # with st.container(horizontal=True, horizontal_alignment="center"):
                #     st.link_button(f"**Job {job_num}**", f"https://{st.session_state.current_tenant}.eh.go.servicetitan.com/#/Job/Index/{job_id}", type='tertiary')

                # Form for the current job
                templates.doc_check_form(job_num, job, None, doc_check_criteria, exdata_key='docchecks_live')


        else:
            st.info("Choose a date range and click 'Fetch Jobs' to begin.")
    elif st.session_state["authentication_status"] is False:
        st.error('Please log in.')
    elif st.session_state["authentication_status"] is None:
        st.warning('Please log in.')
    else:
        st.warning('Please log in.')#


if __name__ == "__main__":
    main()