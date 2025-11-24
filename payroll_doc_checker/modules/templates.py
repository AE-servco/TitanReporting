import streamlit as st
import streamlit_authenticator as stauth

import modules.google_store as gs
import modules.helpers as helpers
from datetime import date, timedelta
import json
import pandas as pd
import base64
from streamlit_pdf_viewer import pdf_viewer

import modules.fetching as fetch

def authenticate_app(config_file):
    config = gs.load_yaml_from_gcs(config_file)

    authenticator = stauth.Authenticate(
        credentials = config['credentials']
    )

    authenticator.login(location='main')

def sidebar_filters():
    # Sidebar controls for date range and filters
    with st.sidebar.form(key=f"filter_form"):
        st.header("Job filters")
        tenant_filter = st.selectbox(
            "ServiceTitan Tenant",
            [
                "FoxtrotWhiskey (NSW)", 
                "SierraDelta (WA)",
                "VictorTango (VIC)",
                "EchoZulu (QLD)",
                "MikeEcho (VIC new)",
                "BravoGolf (QLD new)",
            ]
        )
        today = date.today()
        default_start = today - timedelta(days=7)
        start_date = st.date_input("Start date", value=default_start, format="DD/MM/YYYY")
        end_date = st.date_input("End date", value=today, format="DD/MM/YYYY")
        custom_job_id = st.text_input(
            "Job ID Search", placeholder="Manual search for job", help="Job ID is different to the job number. ID is the number at the end of the URL of the job's page in ServiceTitan. This overrides any date filters and will show only the job specified (if it exists)."
        )
        job_status_filter = st.multiselect(
            "Job statuses to include (leave empty for all)",
            ['Scheduled', 'Dispatched', 'InProgress', 'Hold', 'Completed', 'Canceled'],
            default=["Completed"]
        )
        filter_unsucessful = st.checkbox("Exclude unsuccessful jobs", value=True)
        fetch_jobs_button = st.form_submit_button("Fetch Jobs", type="primary")

    # When the fetch button is pressed, call the API and reset state
    if fetch_jobs_button:
        helpers.fetch_jobs_button_call(tenant_filter, start_date, end_date, job_status_filter, filter_unsucessful, custom_job_id)

def job_nav_buttons(idx):
    client = st.session_state.clients.get(st.session_state.current_tenant)
    if idx < 0:
        idx = 0
    if idx >= len(st.session_state.jobs):
        idx = len(st.session_state.jobs) - 1
    job = st.session_state.jobs[idx]
    job_id = str(job.get("id"))
    job_num = str(job.get("jobNumber"))

    # Navigation buttons
    col_prev, col_next = st.columns([1, 1], width=250)
    with col_prev:
        if st.button("Previous Job"):
            if st.session_state.current_index > 0:
                st.session_state.current_index -= 1
                fetch.schedule_prefetches(client)
                st.rerun()
    with col_next:
        if st.button("Next Job"):
            if st.session_state.current_index < len(st.session_state.jobs) - 1:
                st.session_state.current_index += 1
                fetch.schedule_prefetches(client)
                st.rerun()
    return job, job_id, job_num

def show_job_info(job):
    client = st.session_state.clients.get(st.session_state.current_tenant)
    # st.write(job)
    job_amt = job['total']
    # st.write(job)
    st.write(f"Job total: ${job_amt}")
    table_df = pd.DataFrame({
        "First Appointment": [
            job['first_appt_num'], 
            client.format_local(client.from_utc_string(job['first_appt_start']), fmt="%H:%M, %d/%m/%Y"), 
            client.format_local(client.from_utc_string(job['first_appt_end']), fmt="%H:%M, %d/%m/%Y"), 
            client.format_local(client.from_utc_string(job['first_appt_arrival_start']), fmt="%H:%M, %d/%m/%Y"), 
            client.format_local(client.from_utc_string(job['first_appt_arrival_end']), fmt="%H:%M, %d/%m/%Y")],
        "Last Appointment": [
            job['last_appt_num'], 
            client.format_local(client.from_utc_string(job['last_appt_start']), fmt="%H:%M, %d/%m/%Y"), 
            client.format_local(client.from_utc_string(job['last_appt_end']), fmt="%H:%M, %d/%m/%Y"), 
            client.format_local(client.from_utc_string(job['last_appt_arrival_start']), fmt="%H:%M, %d/%m/%Y"),
            client.format_local(client.from_utc_string(job['last_appt_arrival_end']), fmt="%H:%M, %d/%m/%Y")
        ]
    }, index=['Appointment #', 'Recorded Start time', 'Recorded End time', 'Arrival window start', 'Arrival window end'])
    st.table(table_df)
    # st.write(job)

@st.fragment
def show_images(imgs):
    img_size = st.slider(
        "Image Size:",
        min_value=1,
        max_value=10,
        value=3,
        step=1,
        width=200
        )
    
    client = st.session_state.clients.get(st.session_state.current_tenant)

    with st.container(horizontal=True, height=1000):
        for filename, file_date, file_by, signed_url in imgs:
            if signed_url:
                data = gs.fetch_from_signed_url(signed_url)
                st.image(data, caption=f'{st.session_state.employee_lists.get(st.session_state.current_tenant).get(file_by)} at {client.format_local(file_date, fmt="%H:%M on %d/%m/%Y")}', width=img_size * 100)
            else:
                st.write(filename)

def show_pdfs(pdfs):
    # Provide a search box to filter document names
    search_query = st.text_input("Search document names", key=f"search_pdfs")
    filtered_pdfs = pdfs
    if search_query:
        query_lower = search_query.lower()
        filtered_pdfs = [(fname, file_date, file_by, signed_url) for fname, file_date, file_by, signed_url in pdfs if query_lower in fname.lower()]
    if filtered_pdfs:
        for fname, file_date, file_by, signed_url in filtered_pdfs:
            with st.expander(fname):
                if signed_url:
                    data = gs.fetch_from_signed_url(signed_url)
                    st.download_button(
                        label=f"Download",
                        data=data,
                        file_name=fname,
                        mime="application/octet-stream"
                    )
                    with st.container(height=1000):
                        pdf_viewer(data, key=fname)
    else:
        st.info("No documents match your search.")

def doc_check_form(job_num, job, attachments, doc_check_criteria, exdata_key='docchecks_testing'):
    with st.form(key=f"doccheck_{job_num}"):
        client = st.session_state.clients.get(st.session_state.current_tenant)
        st.subheader(f"Job {job_num} Doc Check")
        checks = {}

        initial_checks = job.get("tmp_doccheck_bits", fetch.get_job_external_data(job, exdata_key)) # get tmp bits from prev submission if they exist, otherwise fetch from job's external data.
        if not initial_checks.get("qs", False):
            initial_checks['qs'] = helpers.pre_fill_quote_signed_check(attachments.get("pdfs", []))
        if not initial_checks.get("is", False):
            initial_checks['is'] = helpers.pre_fill_invoice_signed_check(attachments.get("pdfs", []))
        if initial_checks['is'] and initial_checks['qs']:
            st.session_state.prefill_txt = "Prefilled: Quote signed and Invoice signed"
        elif initial_checks['is']:
            st.session_state.prefill_txt = "Prefilled: Invoice signed"
        elif initial_checks['qs']:
            st.session_state.prefill_txt = "Prefilled: Quote signed"
        for check_code, check in doc_check_criteria.items():
            default = bool(initial_checks.get(check_code, False))
            checks[check_code] = int(st.checkbox(check, key=f"{job_num}_{check_code}", value=default))

        submitted = st.form_submit_button("Submit")
        if submitted:
            encoded = json.dumps(checks)
            external_data_payload = {
                "externalData": {
                    "patchMode": "Replace",
                    "applicationGuid": client.app_guid,
                    "externalData": [{"key": exdata_key, "value": encoded}]
                }
            }

            patch_url = client.build_url('jpm', 'jobs', resource_id=job['id'])
            try:
                client.patch(patch_url, json=external_data_payload)
                st.success("Form submitted successfully")

                job['tmp_doccheck_bits'] = checks # add to job so that when returning to job's doc check page, they stay filled as they were. This is needed because the job data is not re-fetched on "next" or "prev" buttons.
            except Exception as e:
                st.error(f"Failed to submit form: {e}")

