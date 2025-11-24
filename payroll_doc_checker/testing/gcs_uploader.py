from typing import Optional
from datetime import timedelta
from google.cloud import storage
from servicetitan_api_client import ServiceTitanClient
from google.cloud import secretmanager
import google.auth

from google.auth.transport import requests

def get_secret(secret_id, project_id="servco1", version_id="latest"):
    client = secretmanager.SecretManagerServiceClient()
    name = f"projects/{project_id}/secrets/{secret_id}/versions/{version_id}"
    response = client.access_secret_version(request={"name": name})
    secret_payload = response.payload.data.decode("UTF-8")
    return secret_payload

def upload_bytes_to_gcs(
    data: bytes,
    bucket_name: str,
    blob_name: str,
    content_type: Optional[str] = None,
) -> str:
    """
    Upload raw bytes to a Google Cloud Storage object.

    Args:
        data: The raw bytes to upload.
        bucket_name: Name of the target GCS bucket.
        blob_name: Path/name of the object in the bucket, e.g. "images/foo.jpg".
        content_type: Optional MIME type, e.g. "image/jpeg" or "application/pdf".

    Returns:
        The public GCS URI of the uploaded object (gs://bucket/blob).
    """
    # Assumes GOOGLE_APPLICATION_CREDENTIALS env var is set OR you're on GCP with default creds.
    client = storage.Client()
    bucket = client.bucket(bucket_name)
    blob = bucket.blob(blob_name)

    # Upload the bytes
    blob.upload_from_string(data, content_type=content_type)

    return f"gs://{bucket_name}/{blob_name}"

def upload_bytes_to_gcs_signed(
    data: bytes,
    bucket_name: str,
    blob_name: str,
    content_type: Optional[str] = None,
    expires_in_seconds: int = 3600,
) -> str:
    """
    Upload raw bytes to Google Cloud Storage and return a signed URL.

    Args:
        data: Raw bytes to upload.
        bucket_name: Name of the GCS bucket.
        blob_name: Path/name of the file inside the bucket.
        content_type: Optional MIME type (e.g., "image/jpeg").
        expires_in_seconds: How long the signed URL should remain valid (default 1 hour).

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

    # Create signed URL (GET)
    url = blob.generate_signed_url(
        version="v4",
        expiration=timedelta(seconds=expires_in_seconds),
        method="GET",
        response_disposition=f'inline; filename="{blob_name}"'
    )

    return url

def delete_all_in_bucket(bucket_name: str) -> None:
    """
    Deletes ALL objects inside a Google Cloud Storage bucket.

    Args:
        bucket_name: Name of the GCS bucket.
    """

    client = storage.Client()
    bucket = client.bucket(bucket_name)

    blobs = client.list_blobs(bucket)

    count = 0
    for blob in blobs:
        blob.delete()
        count += 1

    print(f"Deleted {count} objects from bucket '{bucket_name}'.")

if __name__ == "__main__":
    # tenant = 'foxtrotwhiskey'
    # at_id = 143899860
    # client = ServiceTitanClient(
    #             app_key=get_secret("ST_app_key_tester"), 
    #             app_guid=get_secret("ST_servco_integrations_guid"), 
    #             tenant=get_secret(f"ST_tenant_id_{tenant}"), 
    #             client_id=get_secret(f"ST_client_id_{tenant}"), 
    #             client_secret=get_secret(f"ST_client_secret_{tenant}"), 
    #             environment="production"
    #         )
    # url = client.build_url(
    #     'forms', 'jobs/attachment', resource_id=at_id 
    # )
    # data = client.get(url)
    # print(upload_bytes_to_gcs_signed(data=data, bucket_name='doc-check-attachments', blob_name='test_attachment'))
    # delete_all_in_bucket('doc-check-attachments')
    credentials, project_id = google.auth.default()
    # r = requests.Request()
    # credentials.refresh(r)
    print(credentials.service_account_email)