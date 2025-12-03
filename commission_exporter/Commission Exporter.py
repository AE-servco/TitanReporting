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
import modules.lookup_tables as lookup

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
        tenant = st.selectbox(
            "Select tenant",
            lookup.get_tenants().keys()
        )
        # start_date = st.date_input("Start date", value=last_monday)
        end_date = st.date_input("Week ending", value=last_sunday)
        start_date = end_date - dt.timedelta(days=6)

        submitted = st.form_submit_button("Fetch and build workbook")
        
    if submitted:
        tenant_code = lookup.get_tenants()[tenant]
        with st.spinner("Loading..."):
            ss.client = helpers.get_client(tenant_code)
            with st.spinner("Fetching employee info..."):
                employee_map = helpers.get_all_employee_ids(ss.client)
                # employee_map = {}


            with st.spinner("Fetching tags..."):
                tenant_tags = fetching.fetch_tag_types(ss.client)
                # tenant_tags = []

            with st.spinner("Fetching appointments..."):
                appts = fetching.fetch_appts(ss.client, start_date, end_date)
                first_appts = format.get_first_appts(appts)
                job_ids = format.get_job_ids(first_appts)
                appt_ids = [appt['id'] for appt in appts]

                # st.write(appts)
                # st.write('----------')
                # st.write(first_appts)

            with st.spinner("Fetching appointment assignments..."):
                appt_assmnts = fetching.fetch_appt_assmnts(ss.client, appt_ids)
                # appt_assmnts = []

            with st.spinner("Fetching jobs..."):
                jobs = fetching.fetch_jobs(ss.client, job_id_ls=job_ids)
                invoice_ids = format.get_invoice_ids(jobs)

            with st.spinner("Fetching estimates..."):
                estimates = fetching.fetch_estimates(start_date, end_date, ss.client)
                # estimates = []

            with st.spinner("Fetching invoices..."):
                invoices = fetching.fetch_invoices(invoice_ids, ss.client)
                # invoices = []

            with st.spinner("Fetching payments..."):
                payments = fetching.fetch_payments(invoice_ids, ss.client)
                # payments = []
            
            with st.spinner("Formatting data..."):
                appt_assmnts = [format.format_appt_assmt(appt) for appt in appt_assmnts]
                appt_assmnts_by_job, num_appts_per_job = format.group_appt_assmnts_by_job(appt_assmnts)
                first_appts_by_id = format.extract_id_to_key(first_appts, 'jobId')
                for job in jobs:
                    job['appt_techs'] = set(appt_assmnts_by_job.get(job['id'], []))
                    job['num_of_appts_in_mem'] = num_appts_per_job.get(job['id'], 0)
                    job['first_appt'] = first_appts_by_id.get(job['id'], {})
                jobs_w_nones = [format.format_job(job, ss.client, tenant_tags, exdata_key='docchecks_testing') for job in jobs]
                jobs = [job for job in jobs_w_nones if job is not None]
                invoices = [format.format_invoice(invoice) for invoice in invoices]
                payments = helpers.flatten_list([format.format_payment(payment, ss.client) for payment in payments])
                open_estimates = [e for e in [format.format_estimate(est, sold=False) for est in estimates] if e is not None]
                sold_estimates = [e for e in [format.format_estimate(est, sold=True) for est in estimates] if e is not None]

                jobs_df = pd.DataFrame(jobs)
                # st.dataframe(jobs_df)
                invoices_df = pd.DataFrame(invoices)
                payments_df = pd.DataFrame(payments)
                # payments_grouped = payments_df.groupby('invoiceId', as_index=False).agg(','.join)
                payments_grouped = payments_df.groupby('invoiceId', as_index=False).agg(lambda x: ', '.join(sorted(list(set(x)))))
                open_estimates_df = pd.DataFrame(open_estimates)
                sold_estimates_df = pd.DataFrame(sold_estimates)
                open_estimates_grouped = open_estimates_df.groupby('job_id', as_index=False).agg({'est_subtotal': 'sum'})
                sold_estimates_grouped = sold_estimates_df.groupby('job_id', as_index=False).agg({'est_subtotal': 'sum'})

                # st.dataframe(invoices_df)
                # st.dataframe(open_estimates_df)
                # st.dataframe(sold_estimates_df)
                # st.dataframe(open_estimates_grouped)
                # st.dataframe(sold_estimates_grouped)

            with st.spinner("Merging data..."):
                merged = helpers.merge_dfs([jobs_df, invoices_df, payments_grouped], on='invoiceId', how='left')
                merged = helpers.merge_dfs([merged, open_estimates_grouped], on='job_id')
                merged = merged.rename(columns={'est_subtotal': 'open_est_subtotal'})
                merged = helpers.merge_dfs([merged, sold_estimates_grouped], on='job_id')
                merged = merged.rename(columns={'est_subtotal': 'sold_est_subtotal'})
                merged = merged.sort_values(by='first_appt_start_dt')
                # merged = pd.merge(pd.merge(jobs_df, invoices_df, on='invoiceId', how='left'), payments_grouped, on='invoiceId', how='left')
                job_records = merged.to_dict(orient='records')
                for job in job_records:
                    job['payments_in_time'] = helpers.check_payment_dates(job, end_date)
            # st.write(job_records)
            # st.dataframe(merged)
            
            with st.spinner("Separating by technician..."):
                # group by tech name
                jobs_by_tech = format.group_jobs_by_tech(job_records, employee_map)

            with st.spinner("Building spreadsheet..."):
                excel_bytes = build_workbook(
                    jobs_by_tech=jobs_by_tech,
                    end_date=end_date
                )
        
        templates.show_download_button(excel_bytes, f"commissions_{tenant_code}_{start_date}_{end_date}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
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