import json
import yaml
from google.cloud import storage
from typing import Optional
import requests
from datetime import timedelta
import google.auth
import os

from google.auth.transport import requests as grequests

BUCKET_NAME = "service_titan_reporter_data"

def is_running_in_cloud_run():
    """Checks if the application is running in Google Cloud Run."""
    return bool(os.environ.get('K_SERVICE'))

def _get_bucket():
    client = storage.Client()
    return client.bucket(BUCKET_NAME)

def load_file(filename):
    """Load JSON config from GCS."""
    bucket = _get_bucket()
    blob = bucket.blob(filename)

    if not blob.exists():
        return {}  

    data = blob.download_as_text()
    return json.loads(data)

def save_file(json_dict: dict, blobname):
    """Save JSON config back to GCS."""
    bucket = _get_bucket()
    blob = bucket.blob(blobname)
    blob.upload_from_string(
        json.dumps(json_dict, indent=2),
        content_type="application/json"
    )

def load_yaml_from_gcs(blob_name):
    """Load YAML config from GCS"""
    bucket = _get_bucket()
    blob = bucket.blob(blob_name)

    if not blob.exists():
        return {}  

    data = blob.download_as_text()

    return yaml.safe_load(data)

def save_yaml_to_gcs(data, blobname):
    """Save YAML config back to GCS."""
    yaml_text = yaml.safe_dump(data, sort_keys=False)

    bucket = _get_bucket()
    blob = bucket.blob(blobname)
    blob.upload_from_string(yaml_text, content_type="text/yaml")

def upload_bytes_to_gcs_signed(
    data: bytes,
    bucket_name: str,
    blob_name: str,
    content_type: Optional[str] = None,
    expires_in_seconds: int = 10800,
) -> str:
    """
    Upload raw bytes to Google Cloud Storage and return a signed URL.

    Args:
        data: Raw bytes to upload.
        bucket_name: Name of the GCS bucket.
        blob_name: Path/name of the file inside the bucket.
        content_type: Optional MIME type (e.g., "image/jpeg").
        expires_in_seconds: How long the signed URL should remain valid (default 3 hours).

    Returns:
        A signed URL for downloading the uploaded object.
    """

    client = storage.Client()
    bucket = client.bucket(bucket_name)
    blob = bucket.blob(blob_name)

    # Upload bytes
    blob.upload_from_string(
        data,
        content_type=content_type
    )
    credentials, project_id = google.auth.default()
    if is_running_in_cloud_run():
        r = grequests.Request()
        credentials.refresh(r)
    # print(credentials)
    # Create signed URL (GET)
    if hasattr(credentials, "service_account_email"):
        url = blob.generate_signed_url(
            version="v4",
            expiration=timedelta(seconds=expires_in_seconds),
            method="GET",
            response_disposition=f'inline; filename="{blob_name}"',
            service_account_email=credentials.service_account_email,
            access_token=credentials.token,
        )
        return url
    return None

def fetch_from_signed_url(url: str) -> bytes:
    """
    Download bytes from a GCS signed URL.
    """
    response = requests.get(url)
    response.raise_for_status()
    return response.content