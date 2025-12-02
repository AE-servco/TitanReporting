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
# Configuration and helpers
###############################################################################

ATTACHMENT_DOWNLOADER_URL = 'https://attachment-downloader-901775793617.australia-southeast1.run.app/'
# ATTACHMENT_DOWNLOADER_URL = 'http://0.0.0.0:8000'

SIGNED_URL_TTL = 900

TENANTS = [
    "foxtrotwhiskey", 
    "sierradelta",
    "victortango",
    "echozulu",
    "mikeecho",
    "bravogolf",
]

###############################################################################
# Streamlit app logic
###############################################################################

def main() -> None:


    st.set_page_config(page_title="ServiceTitan Job Browser", layout="wide")
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

    st.sidebar.title("Doc Checks Payroll")

    templates.authenticate_app('st_auth_config_plumber_commissions.yaml')

    
    if st.session_state["authentication_status"]:
        with st.spinner("Initialising..."):
            doc_check_criteria = helpers.get_doc_check_criteria()

            # Initialise session state collections
            if "jobs" not in st.session_state:
                st.session_state.jobs: List[Dict[str, Any]] = []
            if "clients" not in st.session_state:
                st.session_state.clients: Dict[str, Any] = {}
            if "employee_lists" not in st.session_state:
                st.session_state.employee_lists: Dict[str, Any] = {}
            if "current_tenant" not in st.session_state:
                st.session_state.current_tenant: str = ""
            if "current_index" not in st.session_state:
                st.session_state.current_index: int = 0
            # if "prefetched" not in st.session_state:
            #     # Cache of prefetched attachments keyed by job ID.  Each entry
            #     # contains a dictionary with ``imgs`` and ``pdfs`` lists.
            #     st.session_state.prefetched: Dict[str, Dict[str, List[Tuple[str, Any]]]] = {}
            if "prev_img_size" not in st.session_state:
                st.session_state.prev_img_size: int = 2
            if "app_guid" not in st.session_state:
                st.session_state.app_guid = helpers.get_secret('ST_servco_integrations_guid')
            if "jobs_queued" not in st.session_state:
                st.session_state.jobs_queued: Dict = {}

        # st.write(st.session_state.jobs_queued)
        templates.sidebar_filters()

        with st.sidebar:
            st.markdown("---")

        # Process completed prefetch futures and update prefetched cache
        # helpers.process_completed_prefetches()

        # Display the current job if available
        if st.session_state.jobs:
            # client = st.session_state.clients.get(st.session_state.current_tenant)
            idx = st.session_state.current_index
            # job, job_id, job_num = templates.job_nav_buttons(idx)
            job = st.session_state.jobs[idx]
            job_id = str(job.get("id"))
            job_num = str(job.get("jobNumber"))
            idx = st.session_state.current_index

            with st.container(horizontal_alignment="center", gap=None):
                with st.container(horizontal=True, horizontal_alignment="center"):
                    templates.nav_button('prev')
                    st.link_button(f"**Job {job_num}**", f"https://{st.session_state.current_tenant}.eh.go.servicetitan.com/#/Job/Index/{job_id}", type='tertiary')
                    templates.nav_button('next')
                st.text(f"({idx + 1} of {len(st.session_state.jobs)})")
                with st.form("jobindexselector", width=200):
                    index_selected = st.number_input("Go to image number:", min_value=1, max_value=len(st.session_state.jobs), value=1)
                    index_selector_submit = st.form_submit_button("Go")
                    if index_selector_submit:
                        st.session_state.current_index = index_selected-1
                        fetch.schedule_prefetches(st.session_state.clients[st.session_state.current_tenant])
                        st.rerun()

            job_info, attachments = st.columns([1,4])

            with job_info:
                # prefill_holder = st.text("")
                templates.show_job_info(job)

            with attachments:
                job_attachment_status, error_msg, last_update_time = fetch.get_job_status(job_id, st.session_state.clients['supabase'], st.session_state.current_tenant)
                try:
                    last_update_time = datetime.fromisoformat(last_update_time).replace(tzinfo=ZoneInfo("Australia/Sydney"))
                    update_time_diff = datetime.now() - last_update_time
                except TypeError:
                    # This is broken, always TypeError
                    update_time_diff = timedelta(seconds=0)

                # st.write(update_time_diff)
                # if job_attachment_status == 2 and update_time_diff < timedelta(seconds=30):
                if job_attachment_status == 2:# and update_time_diff < timedelta(seconds=(SIGNED_URL_TTL-100)):
                    attachments_response = fetch.get_attachments_supabase(job_id, st.session_state.clients['supabase'], st.session_state.current_tenant)

                    imgs = [att for att in attachments_response if att['type'] == 'img']
                    pdfs = [att for att in attachments_response if att['type'] == 'pdf']
                    # vids = [att for att in attachments_response if att['type'] == 'vid']
                    # other = [att for att in attachments_response if att['type'] == 'oth']

                elif job_attachment_status == 1:
                    with st.spinner("Downloading attachments. Refreshing in 2 seconds..."):
                        # st.write('status = 1')
                        time.sleep(2)
                        st.rerun()
                elif job_attachment_status == -1:
                    st.write(error_msg)
                    st.write("Please reload the page. If you keep seeing this error, please inform Albie (send screenshot of error message if possible).")
                    imgs = None
                    pdfs = None
                else:
                    # If not already prefetched, download synchronously all attachments
                    with st.spinner("Downloading attachments. Refreshing in 5 seconds..."):
                        # st.write('status = else')
                        fetch.request_job_download(job_id, st.session_state.current_tenant, ATTACHMENT_DOWNLOADER_URL, force_refresh=True)
                        time.sleep(5)
                        st.rerun()

                # Display attachments in tabs: one for images and one for other docs
                tab_images, tab_docs = st.tabs(["Images", "Other Documents"])

                # Show other documents (e.g., PDFs)
                with tab_docs:
                    if pdfs:
                        templates.show_pdfs(pdfs, 900)
                    else:
                        st.info("No PDFs for this job.")

                # Sidebar form for the current job
                with st.sidebar:
                    templates.doc_check_form(job_num, job, pdfs, doc_check_criteria, exdata_key='docchecks_live')

                # Show images
                with tab_images:
                    if imgs:
                        with st.spinner("Loading images..."):
                            imgs.sort(key=lambda img: img['file_date'])
                            templates.show_images(imgs,900)
                    else:
                        st.info("No image attachments for this job.")

            


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