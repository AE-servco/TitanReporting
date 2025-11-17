import streamlit as st
from streamlit import session_state as ss
import streamlit_authenticator as stauth
from datetime import datetime, time, date, timedelta
from zoneinfo import ZoneInfo
from PIL import Image
from io import BytesIO

import modules.data as data
import modules.photos as photos

import modules.google_store as gs

st.set_page_config(
    layout="wide",
    page_title="ServiceTitan Reports"
    )

st.title(f"Doc Check Checker")

CONFIG_FILENAME = 'st_auth_config.yaml'
config = gs.load_yaml_from_gcs(CONFIG_FILENAME)

authenticator = stauth.Authenticate(
    credentials = config['credentials']
)

authenticator.login(location='main')
if 'next_batch_key' not in ss:
    ss.next_batch_key=False

@st.fragment
def job_data_template():
    with st.container(horizontal=True):
        with st.container(width=200):
            state = st.selectbox(
                "State:",
                [
                    "NSW", 
                    "VIC", 
                    "WA",
                    "QLD",
                ]
            )
        with st.container(width=200):
            img_size = st.slider(
                "Image Size:",
                min_value=1,
                max_value=10,
                value=3,
                step=1
            )
        with st.container(width=200):
            job_id = int(st.text_input(
                "Job Id:",
                "143785244"
                # "143302553"
                # "89013379"
            ))
        with st.container(width=200):
            if st.button("Clear session state"):
                data.clear_ss()
        with st.container(width=200):
            if st.button("Next Batch"):
                ss.next_batch_key=True
                print(ss.next_batch_key)

    if state == "VIC":
        state = 'VIC_old'
    if state == "QLD":
        state = 'QLD_old'

    data.check_and_update_ss_for_data_service(state)

    if ss.tech_ids == None or ss.state != state:
        ss.tech_ids = data.get_all_employee_ids(state)
        ss.state = state
    # st.write('techids',ss.tech_ids)
    
    if ss.job_id == None or ss.job_id != job_id:
        ss.job_id = job_id
        with st.spinner("Loading Photos..."):
            if ss.next_batch_key:
                print('loading new')
                ss.imgs = ss.next_batch
            else:
                print('loading current')
                ss.imgs = photos.get_photos_from_job(job_id, state)

    with st.container(horizontal=True, height=1000):
        for img_tup in ss.imgs:
            with st.container(width=img_size * 100):
                # st.image(img_tup[0], width=500)
                # if str(img_tup[1]) in ss.tech_ids:
                #     name = ss.tech_ids[str(img_tup[1])]
                # else:
                #     name = 'Unknown'
                # img = Image.open(BytesIO(img_tup))
                st.image(img_tup, width=img_size * 100)
                # st.write(len(img_tup), img_tup[:50])
                # st.image(img_tup[0], width=img_size * 100)
                # st.write(f"{name} at {data.format_ST_date(img_tup[2], "%H:%M, %d/%m")}")
    
    
    next_attachments = ss[f'st_data_service_{state}'].get_api_data('forms', f'jobs/{143302553}/attachments', options=None, version=2)
    img_types = ('jpg', 'jpeg', 'png')
    next_img_attachments = [a for a in next_attachments if a['fileName'].endswith(img_types)]
    print("Starting prefetch...")
    print(len(next_img_attachments))
    photos.start_prefetch("next_batch", next_img_attachments, state)


def doc_check_template():
    with st.form("form1"):
        before_img = st.checkbox(
            "Before Photo",
        )
        after_img = st.checkbox(
            "After Photo",
        )
        receipt_img = st.checkbox(
            "Receipt Photo",
        )
        submitted = st.form_submit_button("Submit")
        if submitted:
            export_data = {
                "doc_check": str(int(before_img)) + str(int(after_img)) + str(int(receipt_img))
            }
            data.update_job_external_data(ss.job_id, ss.state, export_data)
            st.write("before_img", before_img, "after_img", after_img, "receipt_img", receipt_img)

if ss["authentication_status"]:
    if 'app_guid' not in ss:
        ss.app_guid = data.get_secret("ST_servco_integrations_guid")
    if 'state' not in ss:
        ss.state = None
    if 'job_id' not in ss:
        ss.job_id = None
    if 'tech_ids' not in ss:
        ss.tech_ids = None
    if 'imgs' not in ss:
        ss.imgs = None

    with st.sidebar:
        doc_check_template()
        
    st.subheader("=======Work in progress=======")

    time_now = datetime.now(ZoneInfo("Australia/Sydney"))

    job_data_template()

elif ss["authentication_status"] is False:
    data.clear_ss()
    st.error('Go to home page to log in.')
elif ss["authentication_status"] is None:
    data.clear_ss()
    st.warning('Go to home page to log in.')