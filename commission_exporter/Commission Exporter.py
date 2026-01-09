# streamlit_app.py
import streamlit as st
from streamlit import session_state as ss
import streamlit_authenticator as stauth
import datetime as dt
import pandas as pd
import calendar
from pprint import pprint

import modules.google_store as gs

from modules.excel_builder import build_workbook 
from modules.excel_templates import CommissionSpreadSheetExporter

import modules.helpers as helpers
import modules.templates as templates
import modules.data_formatting as format
import modules.data_fetching as fetching
import modules.lookup_tables as lookup

###############################################################################
# Filter warnings
###############################################################################
import warnings
warnings.filterwarnings("ignore", message=".*cookie_manager.*")

st.title("Commission Spreadsheet Exporter")

CONFIG_FILENAME = 'st_auth_config_commission_exporter.yaml'
config = gs.load_yaml_from_gcs(CONFIG_FILENAME)

authenticator = stauth.Authenticate(
    credentials = config['credentials']
)

authenticator.login(location='main')

if ss["authentication_status"]:

    if "spreadsheets" not in ss:
        ss.spreadsheets = {}

    st.write('If there is a yellow box above talking about a Cookie Manager, please disregard. It does not affect the functionality of the app.')
    st.write('Please select a timeframe, tenant, and date to filter by, then click "Fetch and build workbook". This will gather all the relevant data and produce the "Download Spreadsheet" button after a short wait. Click this to get the spreadsheet.')

    today = dt.date.today()
    # default to last Mon–Sun
    last_monday = today - dt.timedelta(days=today.weekday() + 7)
    last_sunday = last_monday + dt.timedelta(days=6)

    timeframe = st.selectbox(
        "Select timeframe",
        [
            'Monthly',
            'Weekly'
        ],
        
    )

    with st.form("date_select"):
        state = st.selectbox(
            "Select State",
            lookup.get_tenant_from_state().keys()
        )
        # end_date = st.date_input("Week ending", value=last_sunday)
        # start_date = end_date - dt.timedelta(days=6)
        # submitted = st.form_submit_button("Fetch and build workbook")

        if timeframe == 'Weekly':
            end_date = st.date_input("Week ending", value=last_sunday)
            start_date = end_date - dt.timedelta(days=6)
            submitted = st.form_submit_button("Fetch and build workbook")

        elif timeframe == 'Monthly':
            mon_abbr_to_num = {name: num for num, name in enumerate(calendar.month_abbr) if num}
            year = st.selectbox(
                            "Year",
                            [
                                2025,
                                2026
                            ],
                            index=1
                        )
            month = st.selectbox(
                            "Month",
                            list(mon_abbr_to_num.keys()),
                            index=today.month-2 if today.month > 1 else 0
                        )
            end_date = helpers.get_last_day_of_month_datetime(year, mon_abbr_to_num[month])
            start_date = dt.date(year, mon_abbr_to_num[month], 1)
            submitted = st.form_submit_button("Fetch and build workbook")
        else: 
            st.write("Select week or month above")
            

        
    if submitted:
        tenant_codes = lookup.get_tenant_from_state(state)
        # print(tenant_code)
        data_present = False # flag for checking if data is present. Only here because refactoring for merging states, not the best way to do it I know.
        job_records = []
        employee_map_total = {}
        for tenant_code in tenant_codes:
            print(f"Running {tenant_code}")
            with st.spinner("Loading..."):
                ss.client = helpers.get_client(tenant_code)
                with st.spinner("Fetching employee info..."):
                    employee_map = helpers.get_all_employee_ids(ss.client)
                    pprint("employee_map")
                    pprint(employee_map)
                    employee_map_total.update(employee_map)


                with st.spinner("Fetching tags..."):
                    tenant_tags = fetching.fetch_tag_types(ss.client)
                    # tenant_tags = []
                    # pprint("tenant_tags")
                    # pprint(tenant_tags)

                with st.spinner("Fetching appointments..."):
                    appts = fetching.fetch_appts(ss.client, start_date, end_date)
                    pprint(f"len(appts) == {len(appts)}")
                if len(appts) == 0:
                    continue
                
                with st.spinner("Fetching appointments..."):
                    first_appts = format.get_first_appts(appts)
                    pprint(f"len(first_appts) == {len(first_appts)}")
                    if len(first_appts) == 0:
                        continue
                    data_present = True # Change to true if any data present.
                    
                    job_ids = format.get_job_ids(first_appts)
                    pprint(f"len(job_ids) == {len(job_ids)}")
                    appt_ids = [appt['id'] for appt in appts]

                with st.spinner("Fetching appointment assignments..."):
                    appt_assmnts = fetching.fetch_appt_assmnts(ss.client, appt_ids)
                    # appt_assmnts = []

                with st.spinner("Fetching jobs..."):
                    jobs = fetching.fetch_jobs(ss.client, job_id_ls=job_ids)
                    pprint(f"len(jobs) == {len(jobs)}")
                    invoice_ids = format.get_invoice_ids(jobs)
                    # st.write(len(jobs))
                    # st.write(jobs)
                    # st.write(len(invoice_ids))
                    # st.write(invoice_ids)

                with st.spinner("Fetching estimates..."):
                    estimates = fetching.fetch_estimates(start_date, end_date, ss.client)
                    # st.write(len(estimates))
                    # st.write(estimates)
                    # estimates = []

                with st.spinner("Fetching invoices..."):
                    invoices = fetching.fetch_invoices(invoice_ids, ss.client)
                    # st.write(len(invoices))
                    # st.write(pd.DataFrame(invoices))
                    # invoices = []

                with st.spinner("Fetching payments..."):
                    payments = fetching.fetch_payments(invoice_ids, ss.client)
                    # st.write(len(payments))
                    # st.write('payments list')
                    # st.write(pd.DataFrame(payments))
                    # # payments = []
                    # st.write(invoice_ids)
                
                with st.spinner("Formatting data..."):
                    appt_assmnts = [format.format_appt_assmt(appt) for appt in appt_assmnts]
                    appt_assmnts_by_job, num_appts_per_job = format.group_appt_assmnts_by_job(appt_assmnts)
                    first_appts_by_id = format.extract_id_to_key(first_appts, 'jobId')
                    for job in jobs:
                        job['appt_techs'] = set(appt_assmnts_by_job.get(job['id'], []))
                        job['num_of_appts_in_mem'] = num_appts_per_job.get(job['id'], 0)
                        job['first_appt'] = first_appts_by_id.get(job['id'], {})
                    jobs_w_nones = [format.format_job(job, ss.client, tenant_tags, exdata_key='docchecks_live') for job in jobs]
                    jobs = [job for job in jobs_w_nones if job is not None]
                    if len(jobs) == 0:
                        continue
                    invoices = [format.format_invoice(invoice) for invoice in invoices]
                    payments = helpers.flatten_list([format.format_payment(payment, ss.client) for payment in payments])
                    open_estimates = [e for e in [format.format_estimate(est, sold=False) for est in estimates] if e is not None]
                    sold_estimates = [e for e in [format.format_estimate(est, sold=True) for est in estimates] if e is not None]

                    jobs_df = pd.DataFrame(jobs)
                    
                    if len(invoices) == 0:
                        invoices_df = pd.DataFrame(columns=['invoiceId'])
                    else:
                        invoices_df = pd.DataFrame(invoices)
                    if len(payments) == 0:
                        payments_df = pd.DataFrame(columns=['invoiceId'])
                    else:
                        payments_df = pd.DataFrame(payments)
                    # pprint("invoices_df.head():")
                    # pprint(invoices_df.head())
                    # pprint("payments_df.head():")
                    # pprint(payments_df.head())
                    payments_grouped = payments_df.groupby('invoiceId', as_index=False).agg(lambda x: ', '.join(sorted(list(set(x)))))
                    
                    open_estimates_df = pd.DataFrame(open_estimates)
                    if open_estimates_df.empty:
                        open_estimates_df = pd.DataFrame(columns=['job_id', 'est_subtotal'])
                        
                    sold_estimates_df = pd.DataFrame(sold_estimates)
                    if sold_estimates_df.empty:
                        sold_estimates_df = pd.DataFrame(columns=['job_id', 'est_subtotal'])
                        
                    open_estimates_grouped = open_estimates_df.groupby('job_id', as_index=False).agg({'est_subtotal': 'sum'})
                    sold_estimates_grouped = sold_estimates_df.groupby('job_id', as_index=False).agg({'est_subtotal': 'sum'})

                    # st.dataframe(jobs_df)
                    # st.dataframe(invoices_df)
                    # st.dataframe(open_estimates_df)
                    # st.dataframe(sold_estimates_df)
                    # st.dataframe(open_estimates_grouped)
                    # st.dataframe(sold_estimates_grouped)
                    # st.dataframe(payments_df)
                    # st.dataframe(payments_grouped)

                with st.spinner("Merging data..."):
                    merged = helpers.merge_dfs([jobs_df, invoices_df, payments_grouped], on='invoiceId', how='left')
                    merged = helpers.merge_dfs([merged, open_estimates_grouped], on='job_id')
                    if 'est_subtotal' in merged.columns:
                        merged = merged.rename(columns={'est_subtotal': 'open_est_subtotal'})
                        
                    merged = helpers.merge_dfs([merged, sold_estimates_grouped], on='job_id')
                    if 'est_subtotal' in merged.columns:
                        merged = merged.rename(columns={'est_subtotal': 'sold_est_subtotal'})
                    
                    if 'first_appt_start_dt' in merged.columns:
                        merged = merged.sort_values(by='first_appt_start_dt')
                    # merged = pd.merge(pd.merge(jobs_df, invoices_df, on='invoiceId', how='left'), payments_grouped, on='invoiceId', how='left')
                    job_records_tmp = merged.to_dict(orient='records')
                    for job in job_records_tmp:
                        job['payments_in_time'] = helpers.check_payment_dates(job, end_date)
                    
                    # print(job_records)
                    # print(job_records_tmp)
                    job_records.extend(job_records_tmp)
                # st.write(job_records)
                # st.dataframe(merged)
        if data_present:    
            with st.spinner("Separating by technician..."):
                # group by tech name
                relevant_holidays = helpers.get_holidays(state)
                # print(relevant_holidays)
                job_records.sort(key = lambda x: x["first_appt_start_str"])
                jobs_by_tech = format.group_jobs_by_tech(job_records, employee_map_total, end_date, relevant_holidays)

            with st.spinner("Building spreadsheet..."):
                builder = CommissionSpreadSheetExporter(jobs_by_tech, end_date, timeframe=timeframe.lower(), col_offset=1, holidays=relevant_holidays)
                
                excel_bytes = builder.build_workbook()
                pprint(f"Workbook built for {state}, {start_date} to {end_date}.")
            # st.write(jobs_by_tech)
                
            ss.spreadsheets[f"commissions_{state}_{start_date}_{end_date}.xlsx"] = excel_bytes
        
            # templates.show_download_button(excel_bytes, f"commissions_{tenant_code}_{start_date}_{end_date}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.write("No data available.")

    for spreadsheet_name, spreadsheet_data in ss.spreadsheets.items():
        templates.show_download_button(spreadsheet_data, spreadsheet_name, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif ss["authentication_status"] is False:
    st.error('Please log in.')
elif ss["authentication_status"] is None:
    st.warning('Please log in.')