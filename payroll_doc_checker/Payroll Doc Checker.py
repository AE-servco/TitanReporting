from __future__ import annotations

from datetime import date, timedelta
from typing import Dict, List, Set, Tuple, Optional, Any, Iterable
import json
from google.cloud import secretmanager
import time

import streamlit as st
import streamlit_authenticator as stauth
from concurrent.futures import ThreadPoolExecutor, Future, as_completed

from servicetitan_api_client import ServiceTitanClient
import modules.google_store as gs
import modules.helpers as helpers
import modules.templates as templates

###############################################################################
# Configuration and helpers
###############################################################################

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
    st.title("ServiceTitan Job Image Browser")

    templates.authenticate_app('st_auth_config_plumber_commissions.yaml')

    if st.session_state["authentication_status"]:
        with st.spinner("Initialising..."):
            doc_check_criteria = helpers.get_doc_check_criteria()

            # Initialise session state collections
            if "jobs" not in st.session_state:
                st.session_state.jobs: List[Dict[str, Any]] = []
            if "clients" not in st.session_state:
                st.session_state.clients: Dict[str, Any] = {tenant: helpers.get_client(tenant) for tenant in TENANTS}
            if "employee_lists" not in st.session_state:
                st.session_state.employee_lists: Dict[str, Any] = {tenant: helpers.get_all_employee_ids(st.session_state.clients.get(tenant)) for tenant in TENANTS}
            if "current_tenant" not in st.session_state:
                st.session_state.current_tenant: str = ""
            if "current_index" not in st.session_state:
                st.session_state.current_index: int = 0
            if "prefetched" not in st.session_state:
                # Cache of prefetched attachments keyed by job ID.  Each entry
                # contains a dictionary with ``imgs`` and ``pdfs`` lists.
                st.session_state.prefetched: Dict[str, Dict[str, List[Tuple[str, Any]]]] = {}
            if "prefetch_futures" not in st.session_state:
                st.session_state.prefetch_futures: Dict[str, Future] = {}
            if "app_guid" not in st.session_state:
                st.session_state.app_guid = helpers.get_secret('ST_servco_integrations_guid')
            if "prefill_txt" not in st.session_state:
                st.session_state.prefill_txt: str = ""

        templates.sidebar_filters()

        with st.sidebar:
            st.markdown("---")

        # Process completed prefetch futures and update prefetched cache
        helpers.process_completed_prefetches()

        # Display the current job if available
        if st.session_state.jobs:
            client = st.session_state.clients.get(st.session_state.current_tenant)
            idx = st.session_state.current_index
            job, job_id, job_num = templates.job_nav_buttons(idx)
            idx = st.session_state.current_index

            # Display job details and images
            st.write(f"**Viewing job {job_num} ({idx + 1} of {len(st.session_state.jobs)})**")
            st.link_button("Go to job on ServiceTitan", f"https://{st.session_state.current_tenant}.eh.go.servicetitan.com/#/Job/Index/{job_id}")
            prefill_holder = st.text("")

            templates.show_job_info(job)

            attachments = st.session_state.prefetched.get(job_id)
            if attachments is None:
                # If not already prefetched, download synchronously all attachments
                with st.spinner("Downloading attachments..."):
                    attachments = helpers.download_attachments_for_job(job_id, client)
                st.session_state.prefetched[job_id] = attachments

            # Display attachments in tabs: one for images and one for other docs
            tab_images, tab_docs = st.tabs(["Images", "Other Documents"])

            # Show images
            with tab_images:
                imgs = attachments.get("imgs", [])
                imgs.sort(key=lambda img: img[1])
                if imgs:
                    templates.show_images(imgs)
                else:
                    st.info("No image attachments for this job.")

            # Show other documents (e.g., PDFs)
            with tab_docs:
                pdfs = attachments.get("pdfs", [])
                templates.show_pdfs(pdfs)
            

            # Sidebar form for the current job
            with st.sidebar:
                templates.doc_check_form(job_num, job_id, attachments, doc_check_criteria)
                prefill_holder.text(st.session_state.prefill_txt)

        else:
            st.info("Enter a date range and click 'Fetch Jobs' to begin.")
    elif st.session_state["authentication_status"] is False:
        st.error('Please log in.')
    elif st.session_state["authentication_status"] is None:
        st.warning('Please log in.')
    else:
        st.warning('Please log in.')


if __name__ == "__main__":
    main()