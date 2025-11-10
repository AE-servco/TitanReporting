import streamlit as st
import streamlit_authenticator as stauth
import modules.google_store as gs
import modules.helpers as helpers
from datetime import date, timedelta
import json


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
        start_date = st.date_input("Start date", value=default_start)
        end_date = st.date_input("End date", value=today)
        custom_job_id = st.text_input(
            "Job ID Search", placeholder="Manual search for job", help="Job ID is different to the job number. ID is the number at the end of the URL of the job's page in ServiceTitan"
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

def doc_check_form(job_num, job_id, attachments, doc_check_criteria):
    with st.form(key=f"doccheck_{job_num}"):
        client = st.session_state.clients.get(st.session_state.current_tenant)
        st.subheader(f"Job {job_num} Doc Check")
        checks = {}
        # initial_bits = get_job_external_data(job_id, client, st.session_state.app_guid)
        initial_checks = helpers.get_job_external_data(job_id, client, st.session_state.app_guid)
        if not initial_checks.get("qs", False):
            initial_checks['qs'] = helpers.pre_fill_quote_signed_check(attachments.get("pdfs", []))
        if not initial_checks.get("is", False):
            initial_checks['is'] = helpers.pre_fill_invoice_signed_check(attachments.get("pdfs", []))
        if initial_checks['is'] and initial_checks['qs']:
            st.session_state.prefill_txt = "Prefilled: Quote signed and Invoice signed"
            # print(st.session_state.prefill_txt)
        elif initial_checks['is']:
            st.session_state.prefill_txt = "Prefilled: Invoice signed"
            # print(st.session_state.prefill_txt)
        elif initial_checks['qs']:
            st.session_state.prefill_txt = "Prefilled: Quote signed"
            # print(st.session_state.prefill_txt)
        for check_code, check in doc_check_criteria.items():
            default = bool(initial_checks.get(check_code, False))
            checks[check_code] = int(st.checkbox(check, key=f"{job_num}_{check_code}", value=default))

        submitted = st.form_submit_button("Submit")
        if submitted:
            encoded = json.dumps(checks)
            print(encoded)
            external_data_payload = {
                "externalData": {
                    "patchMode": "Replace",
                    "applicationGuid": st.session_state.app_guid,
                    "externalData": [{"key": "docchecks", "value": encoded}]
                }
            }

            patch_url = client.build_url('jpm', 'jobs', resource_id=job_id)
            try:
                client.patch(patch_url, json=external_data_payload)
                st.success("Form submitted successfully")
            except Exception as e:
                st.error(f"Failed to submit form: {e}")


@st.fragment
def show_images(imgs, num_columns=3):
    img_size = st.slider(
        "Image Size:",
        min_value=1,
        max_value=10,
        value=3,
        step=1,
        width=200
        )
    # cols = st.columns(num_columns)
    with st.container(horizontal=True):
        for idx_img, (filename, file_date, file_by, data) in enumerate(imgs):
        # with cols[idx_img % 3]:
            if data:
                st.image(data, caption=f'{file_by} at {file_date}', width=img_size * 100)
            else:
                st.write(filename)



