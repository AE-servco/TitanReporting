import streamlit as st
import streamlit_authenticator as stauth
import modules.google_store as gs
import modules.helpers as helpers
from datetime import date, timedelta
import json
import pandas as pd

def authenticate_app(config_file):
    config = gs.load_yaml_from_gcs(config_file)

    authenticator = stauth.Authenticate(
        credentials = config['credentials']
    )

    authenticator.login(location='main')


