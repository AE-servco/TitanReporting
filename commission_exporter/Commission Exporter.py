# streamlit_app.py
import streamlit as st
from streamlit import session_state as ss
import streamlit_authenticator as stauth
import datetime as dt
import pandas as pd

import modules.google_store as gs
from modules.excel_builder import build_workbook 
import modules.helpers as helpers
import modules.templates as templates
import modules.data_formatting as format
import modules.data_fetching as fetching

st.title("Weekly Commission Sheets (per technician)")

CONFIG_FILENAME = 'st_auth_config_invoice_exporter.yaml'
config = gs.load_yaml_from_gcs(CONFIG_FILENAME)

authenticator = stauth.Authenticate(
    credentials = config['credentials']
)

authenticator.login(location='main')

if ss["authentication_status"]:
    today = dt.date.today()
    # default to last Mon–Sun
    last_monday = today - dt.timedelta(days=today.weekday() + 7)
    last_sunday = last_monday + dt.timedelta(days=6)

    with st.form("date_select"):
        start_date = st.date_input("Start date", value=last_monday)
        end_date = st.date_input("End date", value=last_sunday)

        submitted = st.form_submit_button("Fetch and build workbook")
        
    if submitted:
        with st.spinner("Loading..."):
            ss.client = helpers.get_client('foxtrotwhiskey')
            with st.spinner("Fetching employee info..."):
                employee_map = helpers.get_all_employee_ids(ss.client)

            with st.spinner("Fetching tags..."):
                tenant_tags = fetching.fetch_tag_types(ss.client)

            with st.spinner("Fetching jobs..."):
                jobs = fetching.fetch_jobs(start_date, end_date, ss.client)

            with st.spinner("Fetching appointments..."):
                # TODO: add logic to fetch more appts if job appt num != num output for that job here
                appt_assmnts = fetching.fetch_appt_assmnts(start_date, end_date, ss.client)

            with st.spinner("Fetching invoices..."):
                invoice_ids = format.get_invoice_ids(jobs)

                invoices = fetching.fetch_invoices(invoice_ids, ss.client)
            with st.spinner("Fetching payments..."):
                payments = fetching.fetch_payments(invoice_ids, ss.client)
            
            with st.spinner("Formatting data..."):
                appt_assmnts = [format.format_appt_assmt(appt) for appt in appt_assmnts]
                appt_assmnts_by_job, num_appts_per_job = format.group_appt_assmnts_by_job(appt_assmnts)
                for job in jobs:
                    job['appt_techs'] = set(appt_assmnts_by_job.get(job['id'], []))
                    job['num_of_appts_in_mem'] = num_appts_per_job.get(job['id'], 0)

                jobs_w_nones = [format.format_job(job, ss.client, tenant_tags, exdata_key='docchecks_testing') for job in jobs]
                jobs = [job for job in jobs_w_nones if job is not None]
                invoices = [format.format_invoice(invoice) for invoice in invoices]
                payments = helpers.flatten_list([format.format_payment(payment) for payment in payments])

                jobs_df = pd.DataFrame(jobs)
                invoices_df = pd.DataFrame(invoices)
                payments_df = pd.DataFrame(payments)
                # payments_grouped = payments_df.groupby('invoiceId', as_index=False).agg(','.join)
                payments_grouped = payments_df.groupby('invoiceId', as_index=False).agg(lambda x: ', '.join(sorted(list(set(x)))))

                # st.dataframe(payments_grouped)

            with st.spinner("Merging data..."):
                merged = pd.merge(pd.merge(jobs_df, invoices_df, on='invoiceId', how='left'), payments_grouped, on='invoiceId', how='left')
                job_records = merged.to_dict(orient='records')
            # st.dataframe(merged)
            
            with st.spinner("Separating by technician..."):
                # group by tech name
                jobs_by_tech = format.group_jobs_by_tech(job_records, employee_map)

            with st.spinner("Building spreadsheet..."):
                excel_bytes = build_workbook(
                    jobs_by_tech=jobs_by_tech,
                )
        
        templates.show_download_button(excel_bytes, f"commissions_{start_date}_{end_date}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        # st.download_button(
        #     "Download Excel (all technicians)",
        #     data=excel_bytes,
        #     file_name=f"commissions_{start_date}_{end_date}.xlsx",
        #     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        # )

        # st.write(jobs_by_tech)

elif ss["authentication_status"] is False:
    st.error('Please log in.')
elif ss["authentication_status"] is None:
    st.warning('Please log in.')