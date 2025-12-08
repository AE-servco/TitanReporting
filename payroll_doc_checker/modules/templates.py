import streamlit as st
import streamlit_authenticator as stauth

import modules.google_store as gs
import modules.helpers as helpers
from datetime import date, timedelta, datetime, timezone
from zoneinfo import ZoneInfo
import json
import pandas as pd
import base64
from streamlit_pdf_viewer import pdf_viewer
import base64
from bidict import bidict


import modules.fetching as fetch

def authenticate_app(config_file):
    config = gs.load_yaml_from_gcs(config_file)

    authenticator = stauth.Authenticate(
        credentials = config['credentials']
    )

    authenticator.login(location='main')

def sidebar_filters():

    default_tenant_map = {
        'ghadeer': 0, # NSW
        'nick': 1, # VIC
        'tim': 2, # QLD
        'lachlan': 3, # WA
    }

    # Sidebar controls for date range and filters
    with st.sidebar.form(key=f"filter_form"):
        st.header("Job filters")

        tenant_filter = st.selectbox(
            "ServiceTitan Tenant",
            [
                "FoxtrotWhiskey (NSW)",
                "MikeEcho (VIC)",
                "BravoGolf (QLD)",
                "SierraDelta (WA)",
                "EchoZulu (old QLD)",
                "VictorTango (old VIC)",
            ],
            index = default_tenant_map[st.session_state.username] if st.session_state.username in default_tenant_map else 0
        )
        today = date.today()
        default_start = today - timedelta(days=1)
        start_date = st.date_input("Start date", value=default_start, format="DD/MM/YYYY")
        end_date = st.date_input("End date", value=default_start, format="DD/MM/YYYY")
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

        doc_check_crits = helpers.get_doc_check_criteria()
        doc_check_filter = st.multiselect(
            "Completed doc checks",
            list(doc_check_crits.values())
        )
        doc_check_filter = doc_check_crits.inv[doc_check_filter]

    # When the fetch button is pressed, call the API and reset state
    if fetch_jobs_button:
        helpers.fetch_jobs_button_call(tenant_filter, start_date, end_date, job_status_filter, filter_unsucessful, custom_job_id, doc_check_filter)

def nav_button(dir):
    client = st.session_state.clients.get(st.session_state.current_tenant)
    if dir=='next':
        if st.button("**>**", key='next_button', type='tertiary'):
            if st.session_state.current_index < len(st.session_state.jobs) - 1:
                st.session_state.current_index += 1
            fetch.schedule_prefetches(client)
            st.rerun()
    elif dir=='prev':
        if st.button("**<**", key='prev_button', type='tertiary'):
            if st.session_state.current_index > 0:
                st.session_state.current_index -= 1
            fetch.schedule_prefetches(client)
            st.rerun()
    else:
        st.button("Does nothing", key='button_that_does_nothing')

def show_job_info(job):
    client = st.session_state.clients.get(st.session_state.current_tenant)
    job_amt = job['total']
    job_status = job.get("jobStatus", "Not available.")
    inv_data = job.get("invoice_data", {})
    inv_desc = inv_data.get("summary", "Not available.")
    inv_subtotal = inv_data.get("subtotal", "Not available.")
    inv_bal = inv_data.get("balance", "Not available.")
    inv_amt_paid = inv_data.get("amt_paid", "Not available.")
    project_id = job.get("projectId")

    payment_data = job.get("payment_data", [])

    payment_colors = {
        'Credit Card': 'blue',
        'AMEX': 'blue',
        'EFT/Bank Transfer': 'yellow',
        'Cash': 'green',
    }
    
    if project_id:
        st.link_button(f"**Project (Click me)**", f"https://{st.session_state.current_tenant}.eh.go.servicetitan.com/#/project/{project_id}", type='tertiary')
        other_jobs_in_proj = job.get('other_in_proj')
        if other_jobs_in_proj:
            other_job_count = 1
            for other_job in other_jobs_in_proj:
                if other_job != job['id']:
                    st.link_button(f"Other Job {other_job_count} (Click me)", f"https://{st.session_state.current_tenant}.eh.go.servicetitan.com/#/Job/Index/{other_job}", type='tertiary')
                    other_job_count += 1
    st.write(f"**Job Status**")
    st.write(job_status)
    st.write(f"**Invoice summary**")
    for item in inv_desc.split("|"):
        st.write(item)
    st.write(f"**Invoice subtotal**")
    st.write(f"${inv_subtotal:.2f}")
    st.write(f"**Invoice balance**")
    st.write(f"${inv_bal:.2f}")
    st.write(f"**Amount paid**")
    st.write(f"${inv_amt_paid:.2f}")
    st.write(f"**Job total**")
    st.write(f"${job_amt:.2f}")
    # payments_str = ', '.join([st.badge(f"{p['payment_type']} {p['payment_amt']}", color="blue") for p in payment_data])
    st.write(f"**Payments made**")
    for p in payment_data:
        st.badge(f"{p['payment_type']} ${p['payment_amt']}", color=payment_colors.get(p['payment_type'], 'grey'))
    
    if job['first_appt_num'] != job['last_appt_num']:
        st.write("**First Appointment**")
        st.write(f"{job['first_appt_num']}") 
        st.write('Started at ' + client.format_local(client.from_utc_string(job['first_appt_start']), fmt="%H:%M, %d/%m/%Y")) 
        st.write('Ended at ' + client.format_local(client.from_utc_string(job['first_appt_end']), fmt="%H:%M, %d/%m/%Y"))
    st.write("**Last Appointment**")
    st.write(f"{job['last_appt_num']}") 
    st.write('Started at ' + client.format_local(client.from_utc_string(job['last_appt_start']), fmt="%H:%M, %d/%m/%Y")) 
    st.write('Ended at ' + client.format_local(client.from_utc_string(job['last_appt_end']), fmt="%H:%M, %d/%m/%Y"))


@st.fragment
def show_images(imgs, container_height=1000):
    img_size = st.slider(
        "Image Size:",
        min_value=1,
        max_value=10,
        value=st.session_state.prev_img_size,
        step=1,
        width=200
        )
    st.session_state.prev_img_size = img_size

    def image_bytes_to_base64(image_bytes: bytes) -> str:
        return base64.b64encode(image_bytes).decode("utf-8")
    
    def display_base64_image(b64_str: str, caption, width=300):
        st.markdown(
            f"""
            <img src="data:image/png;base64,{b64_str}" style="width: {width}px;"/>
            <p>{caption}</p>
            """,
            unsafe_allow_html=True
        )

    client = st.session_state.clients.get(st.session_state.current_tenant)
    with st.container(horizontal=True, height=container_height, border=False):
        for img in imgs:
            if img.get('url'):
                try:
                    data = gs.fetch_from_signed_url(img.get('url'))
                    # st.image(data, caption=f'{st.session_state.employee_lists.get(st.session_state.current_tenant).get(int(img.get("file_by")))} at {client.st_date_to_local(img.get("file_date"), fmt="%H:%M on %d/%m/%Y")}', width=img_size * 100)
                    data_b64 = image_bytes_to_base64(data)
                    caption=f'{st.session_state.employee_lists.get(st.session_state.current_tenant).get(int(img.get("file_by")))} at {client.st_date_to_local(img.get("file_date"), fmt="%H:%M on %d/%m/%Y")}'
                    display_base64_image(data_b64, caption, width=img_size * 100)
                except:
                    st.write(f"Error fetching images, trying again..")
                    # st.write("trying again..")
                    try:
                        data = gs.fetch_from_signed_url(img.get('url'))
                        st.image(data, caption=f'{st.session_state.employee_lists.get(st.session_state.current_tenant).get(int(img.get("file_by")))} at {client.st_date_to_local(img.get("file_date"), fmt="%H:%M on %d/%m/%Y")}', width=img_size * 100)
                    except:
                        print(f"ERROR: IMAGE FETCHING url = {img.get('url')}")
                        st.write(f"Errored again, please try again later or go to the job in ServiceTitan.")
            else:
                st.write("Missing image URL")

def show_pdfs(pdfs, container_height=1000):
    # Provide a search box to filter document names
    # search_query = st.text_input("Search document names", key=f"search_pdfs")
    # filtered_pdfs = pdfs
    # if search_query:
    #     query_lower = search_query.lower()
    #     filtered_pdfs = [(fname, file_date, file_by, signed_url) for fname, file_date, file_by, signed_url in pdfs if query_lower in fname.lower()]
    for pdf in pdfs:
        fname = pdf.get('file_name')
        url = pdf.get('url')
        with st.expander(fname):
            if url:
                try:
                    data = gs.fetch_from_signed_url(url)
                    st.download_button(
                        label=f"Download",
                        data=data,
                        file_name=fname,
                        mime="application/octet-stream"
                    )
                    with st.container(height=container_height):
                        pdf_viewer(data, 
                                key=fname,
                                zoom_level=1.25,
                                width="100%")
                except:
                    print(f'ERROR: PDF FETCHING url = {url}')
                    with st.container(height=container_height):
                        st.write("Error fetching PDF, please try again later or go to the job in ServiceTitan.")

def doc_check_form(job_num, job, pdfs, doc_check_criteria, exdata_key='docchecks_live'):
    with st.form(key=f"doccheck_{job_num}"):
        client = st.session_state.clients.get(st.session_state.current_tenant)
        st.subheader(f"Job {job_num} Doc Check")
        checks = {}

        initial_checks = job.get("tmp_doccheck_bits", fetch.get_job_external_data(job, exdata_key)) # get tmp bits from prev submission if they exist, otherwise fetch from job's external data.
        if not initial_checks.get("qs", False):
            initial_checks['qs'] = helpers.pre_fill_quote_signed_check(pdfs)
        if not initial_checks.get("is", False):
            initial_checks['is'] = helpers.pre_fill_invoice_signed_check(pdfs)
        if not initial_checks.get("ie", False):
            initial_checks['ie'] = helpers.pre_fill_invoice_emailed_check(job)
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

