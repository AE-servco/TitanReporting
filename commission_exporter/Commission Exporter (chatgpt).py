# streamlit_app.py
import streamlit as st
from streamlit import session_state as ss
import datetime as dt
import pandas as pd

from io import BytesIO
# from your_module import ServiceTitanClient
from modules.excel_builder import build_commission_workbook 
import modules.helpers as helpers
import modules.templates as templates

st.title("Weekly Commission Sheets (per technician)")

today = dt.date.today()
# default to last Mon–Sun
last_monday = today - dt.timedelta(days=today.weekday() + 7)
last_sunday = last_monday + dt.timedelta(days=6)

start_date = st.date_input("Start date", value=last_monday)
end_date = st.date_input("End date", value=last_sunday)

if st.button("Fetch and build workbook"):
    # 1) fetch jobs from ST (replace this with your real client)
    # client = ...
    # jobs = client.get_jobs_created_between(...)
    # for demo let's say jobs is a list of dicts like in your API

    # you will have real UTC datetimes here
    # start_dt = dt.datetime.combine(start_date, dt.time.min).astimezone(dt.timezone.utc)
    # end_dt = dt.datetime.combine(end_date, dt.time.max).astimezone(dt.timezone.utc)
    # jobs = client.get_jobs_created_between(start_dt, end_dt)

    ss.client = helpers.get_client('foxtrotwhiskey')
    with st.spinner("Fetching employee info..."):
        employee_map = helpers.get_all_employee_ids(ss.client)

    with st.spinner("Fetching jobs..."):
        jobs = helpers.fetch_jobs(last_monday, last_sunday, ss.client)

    with st.spinner("Fetching invoices..."):
        invoice_ids = helpers.get_invoice_ids(jobs)

        invoices = helpers.fetch_invoices(invoice_ids, ss.client)
    with st.spinner("Fetching payments..."):
        payments = helpers.fetch_payments(invoice_ids, ss.client)
    
    with st.spinner("Formatting data..."):
        jobs_w_nones = [helpers.format_job(job, ss.client) for job in jobs]
        jobs = [job for job in jobs_w_nones if job is not None]
        invoices = [helpers.format_invoice(invoice) for invoice in invoices]
        payments = helpers.flatten_list([helpers.format_payment(payment) for payment in payments])

        jobs_df = pd.DataFrame(jobs)
        invoices_df = pd.DataFrame(invoices)
        payments_df = pd.DataFrame(payments)
        # payments_grouped = payments_df.groupby('invoiceId', as_index=False).agg(','.join)
        payments_grouped = payments_df.groupby('invoiceId', as_index=False).agg(lambda x: ', '.join(sorted(list(set(x)))))

    with st.spinner("Merging data..."):
        merged = pd.merge(pd.merge(jobs_df, invoices_df, on='invoiceId', how='left'), payments_grouped, on='invoiceId', how='left')
        job_records = merged.to_dict(orient='records')
    st.dataframe(merged)
    
    with st.spinner("Separating by technician..."):
        # group by tech name
        jobs_by_tech: dict[str, list[dict]] = {}
        for j in job_records:
            tid = j.get("Sold By")
            if not tid:
                continue
            if tid == 'No data - unsuccessful' or tid == '-1':
                name = tid
            elif ',' in tid:
                name = f"Tech {tid}"
            else:
                name = employee_map.get(int(tid), f"Tech {tid}")
            jobs_by_tech.setdefault(name, []).append(j)

    # # if you have specific schemes per tech, specify here:
    # commission_by_tech = {
    #     "Alice Smith": "5-or-10",
    #     "Bob Lee": "flat-10-after-threshold",
    # }

    # excel_bytes = build_commission_workbook(
    #     jobs_by_tech=jobs_by_tech,
    #     week_ending=end_date,
    # )

    # st.download_button(
    #     "Download Excel (all technicians)",
    #     data=excel_bytes,
    #     file_name=f"commissions_{start_date}_{end_date}.xlsx",
    #     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    # )
