import pandas as pd
from google.cloud import secretmanager
from datetime import datetime, time, date, timedelta
from zoneinfo import ZoneInfo
from io import BytesIO
from PIL import Image
import pytz
import holidays
from concurrent.futures import ThreadPoolExecutor
import threading
from streamlit import session_state as ss

import servicepytan as sp
import modules.data as data
from servicetitan_api_client import ServiceTitanClient

def get_secret(secret_id, project_id="servco1", version_id="latest"):
    client = secretmanager.SecretManagerServiceClient()
    name = f"projects/{project_id}/secrets/{secret_id}/versions/{version_id}"
    response = client.access_secret_version(request={"name": name})
    secret_payload = response.payload.data.decode("UTF-8")
    return secret_payload

def state_codes():
    codes = {
        'NSW_old': 'alphabravo',
        'VIC_old': 'victortango',
        'QLD_old': 'echozulu',
        'NSW': 'foxtrotwhiskey',
        'WA': 'sierradelta',
        'QLD': 'bravogolf',
    }
    return codes

def get_STClient(state):
    state_code = state_codes()[state]
    client = ServiceTitanClient(app_key=get_secret("ST_app_key_tester"), tenant=get_secret(f"ST_tenant_id_{state_code}"), client_id=get_secret(f"ST_client_id_{state_code}"), client_secret=get_secret(f"ST_client_secret_{state_code}"), environment="production")

    return client

def get_data_service(state):
    state_code = state_codes()[state]
    st_conn = sp.auth.servicepytan_connect(app_key=get_secret("ST_app_key_tester"), tenant_id=get_secret(f"ST_tenant_id_{state_code}"), client_id=get_secret(f"ST_client_id_{state_code}"), 
    client_secret=get_secret(f"ST_client_secret_{state_code}"), timezone="Australia/Sydney")
    st_data_service = sp.DataService(conn=st_conn)

    return st_data_service


def fetch_many_photos(attachment_ids, state, max_workers=10):
    print('Fetching many photos')
    with ThreadPoolExecutor(max_workers=max_workers) as ex:
        return list(ex.map(ss[f'st_data_service_{state}'].get_attachment, attachment_ids))
    
def prefetch_batch_into_session(key: str, attachment_ids, state):
    """Runs in a separate thread."""
    try:
        print("Fetching inside prefetch_batch_into_session..")
        imgs = fetch_many_photos(attachment_ids, state)
        print("Fetched inside prefetch_batch_into_session.")
        ss[key] = imgs
        print(ss[key])
    except Exception as e:
        # optional: store the error
        ss[key] = {"error": str(e)}

def start_prefetch(key: str, attachment_ids, state):
    """Kick off a background thread that itself uses a thread pool."""
    if key in ss:
        return  # already prefetched or in progress
    t = threading.Thread(
        target=prefetch_batch_into_session,
        args=(key, attachment_ids, state),
        daemon=True,
    )
    t.start()
    
def get_photos_from_job(job_id, state):
    print(f"Getting Photos for {job_id}, {state}")
    images = []

    print(f"Checking data service for {job_id}, {state}")
    data.check_and_update_ss_for_data_service(state)

    print(f"Attachments start for {job_id}, {state}")
    s = datetime.now()
    attachments = ss[f'st_data_service_{state}'].get_api_data('forms', f'jobs/{job_id}/attachments', options=None, version=2)
    
    img_types = ('jpg', 'jpeg', 'png')

    img_attachments = [a for a in attachments if a['fileName'].endswith(img_types)]
    e = datetime.now()
    print(f"Attachments end for {job_id}, {state}, took {e-s} sec")


    print(f"Looping through attachments start for {job_id}, {state}")
    s = datetime.now()
    images = fetch_many_photos([a['id'] for a in img_attachments], state)
    # for attachment in attachments:
    #     print(f"single attachment start for {job_id}, {state}")
    #     s1 = datetime.now()
    #     if attachment['fileName'].endswith(img_types):
    #         img_bytes = ss[f'st_data_service_{state}'].get_attachment(attachment['id'])
    #         # img = Image.open(BytesIO(img_bytes))
    #         images.append((img_bytes, attachment['createdById'], attachment['createdOn']))
    #     e1 = datetime.now()
    # print(f"single attachments end for {job_id}, {state}, took {e1-s1} sec")
    e = datetime.now()
    print(f"Looping through attachments end for {job_id}, {state}, took {e-s} sec")

    return images