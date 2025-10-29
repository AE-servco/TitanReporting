import json
import yaml
from google.cloud import storage

BUCKET_NAME = "service_titan_reporter_data"

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