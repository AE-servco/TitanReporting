from google.cloud import secretmanager
from supabase import create_client, Client
from typing import Dict, List, Set, Tuple, Optional, Any, Iterable
from servicetitan_api_client import ServiceTitanClient

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}

def get_secret(secret_id, project_id="prestigious-gcp", version_id="latest"):
    client = secretmanager.SecretManagerServiceClient()
    name = f"projects/{project_id}/secrets/{secret_id}/versions/{version_id}"
    response = client.access_secret_version(request={"name": name})
    secret_payload = response.payload.data.decode("UTF-8")
    return secret_payload

def get_supabase() -> Client:
    url: str = get_secret("supabase_url")
    key: str = get_secret("supabase_secret_key")
    # key: str = get_secret("supabase_key")
    return create_client(url, key)

def get_st_client(tenant) -> ServiceTitanClient:
    client = ServiceTitanClient(
        app_key=get_secret("st_app_key_tester"), 
        app_guid=get_secret("st_servco_integrations_guid"), 
        tenant=get_secret(f"st_tenant_id_{tenant}"), 
        client_id=get_secret(f"st_client_id_{tenant}"), 
        client_secret=get_secret(f"st_client_secret_{tenant}"), 
        environment="production"
    )
    return client

def group_attachments_by_type(
    attachments: List[Dict[str, Any]],
    extension_map: Optional[Dict[str, Set[str]]] = None,
) -> Dict[str, List[int]]:
    """
    Group attachments into categories based on file extension.

    :param attachments: A list of attachment dicts; each should have at least
                        'id' and 'fileName'.
    :param extension_map: Optional mapping of category names to sets of file
                          extensions (including the dot). If omitted, images
                          go under 'imgs' and PDFs under 'pdfs'.
    :return: A dict where each key is a category name and the value is a list
             of attachment IDs that match that category.

    Example of custom map:
    custom_map = {
            "imgs": {".jpg", ".png"},
            "pdfs": {".pdf"},
            "videos": {".mp4", ".mov"},
            "spreadsheets": {".xls", ".xlsx", ".csv"},
    }
    """
    if extension_map is None:
        extension_map = {
            "imgs": set(IMAGE_EXTENSIONS),
            "pdfs": {".pdf"},
            "videos": {".mp4", ".mov"},
        }

    result = {category: [] for category in extension_map}
    for att in attachments:
        file_name = att.get("fileName")
        att_id = att.get("id")
        file_date = att.get("createdOn")
        file_by = att.get("createdById")
        if not file_name or att_id is None:
            continue

        # Extract the lower‑cased extension with a leading dot
        ext = file_name.lower().rsplit(".", 1)[-1] if "." in file_name else ""
        ext_with_dot = f".{ext}" if ext else ""
        for category, exts in extension_map.items():
            if ext_with_dot in exts:
                result[category].append((file_name, int(att_id), file_date, file_by))
                break  # stop after the first matching category
    return result

def get_attachment_type(filename):
    extension_map = {
            "img": set(IMAGE_EXTENSIONS),
            "pdf": {".pdf"},
            "vid": {".mp4", ".mov"},
        }
    
    ext = filename.lower().rsplit(".", 1)[-1] if "." in filename else ""
    ext_with_dot = f".{ext}" if ext else ""
    for category, exts in extension_map.items():
            if ext_with_dot in exts:
                return category
    return "oth" # if doesn't match, return other