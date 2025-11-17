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

def get_data_service(state):
    state_code = state_codes()[state]
    st_conn = sp.auth.servicepytan_connect(app_key=get_secret("ST_app_key_tester"), tenant_id=get_secret(f"ST_tenant_id_{state_code}"), client_id=get_secret(f"ST_client_id_{state_code}"), 
    client_secret=get_secret(f"ST_client_secret_{state_code}"), timezone="Australia/Sydney")
    st_data_service = sp.DataService(conn=st_conn)

    return st_data_service