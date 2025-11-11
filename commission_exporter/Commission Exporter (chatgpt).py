# streamlit_app.py
import streamlit as st
from streamlit import session_state as ss
import datetime as dt

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
    jobs = helpers.fetch_jobs(last_monday, last_sunday, ss.client)

    invoice_ids = helpers.get_invoice_ids(jobs)

    invoices = helpers.fetch_invoices(invoice_ids, ss.client)
    payments = helpers.fetch_payments(invoice_ids, ss.client)

    invoices = [helpers.format_invoice(invoice) for invoice in invoices]
    # ---- stub ----
    # jobs = [
    #     {
    #         "id": 123,
    #         "createdOn": start_date.isoformat() + "T00:00:00Z",
    #         "status": "Completed",
    #         "subtotal": 5000,
    #         "location": {"city": "Sydney"},
    #         "primaryTechnicianId": 1,
    #         "summary": "Hot water install",
    #         "customerName": "Smith",
    #         "externalData": {"photos": "yes", "invoiceEmailed": "yes"},
    #     },
    #     {
    #         "id": 124,
    #         "createdOn": end_date.isoformat() + "T00:00:00Z",
    #         "paidOn": end_date.isoformat() + "T00:00:00Z",
    #         "completedOn": end_date.isoformat() + "T00:00:00Z",
    #         "status": "Completed",
    #         "subtotal": 22000,
    #         "location": {"city": "Sydney"},
    #         "primaryTechnicianId": 1,
    #         "summary": "Drain work",
    #         "customerName": "Jones",
    #         "externalData": {"photos": "yes", "invoiceEmailed": "yes"},
    #     },
    #     {
    #         "id": 200,
    #         "createdOn": start_date.isoformat() + "T00:00:00Z",
    #         "status": "Completed",
    #         "subtotal": 9000,
    #         "location": {"city": "Parramatta"},
    #         "primaryTechnicianId": 2,
    #         "summary": "AC Service",
    #         "customerName": "Brown",
    #         "externalData": {"photos": "yes"},
    #     },
    #     {
    #         "id": 300,
    #         "createdOn": end_date.isoformat() + "T00:00:00Z",
    #         "status": "Completed",
    #         "subtotal": 10000,
    #         "location": {"city": "Melbourne"},
    #         "primaryTechnicianId": 2,
    #         "summary": "AC Service",
    #         "customerName": "Brown",
    #         "externalData": {"photos": "yes"},
    #     },
    #     {
    #         "id": 400,
    #         "createdOn": start_date.isoformat() + "T00:00:00Z",
    #         "paidOn": start_date.isoformat() + "T00:00:00Z",
    #         "status": "Completed",
    #         "subtotal": 9000,
    #         "location": {"city": "Parramatta"},
    #         "primaryTechnicianId": 2,
    #         "summary": "AC Service",
    #         "customerName": "Brown",
    #         "externalData": {"photos": "yes"},
    #     },
    # ]
    # # ---- end stub ----

    # you probably already have a technician map from ST
    tech_id_to_name = {
        1: "Alice Smith",
        2: "Bob Lee",
    }

    # group by tech name
    jobs_by_tech: dict[str, list[dict]] = {}
    for j in jobs:
        tid = j.get("primaryTechnicianId")
        if not tid:
            continue
        name = tech_id_to_name.get(tid, f"Tech {tid}")
        jobs_by_tech.setdefault(name, []).append(j)

    # if you have specific schemes per tech, specify here:
    commission_by_tech = {
        "Alice Smith": "5-or-10",
        "Bob Lee": "flat-10-after-threshold",
    }

    excel_bytes = build_commission_workbook(
        jobs_by_tech=jobs_by_tech,
        week_ending=end_date,
        commission_by_tech=commission_by_tech,
    )

    st.download_button(
        "Download Excel (all technicians)",
        data=excel_bytes,
        file_name=f"commissions_{start_date}_{end_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
