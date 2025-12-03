from __future__ import annotations

from datetime import datetime

from servicetitan_api_client import ServiceTitanClient
from typing import Dict, List, Set, Tuple, Optional, Any, Iterable
from concurrent.futures import ThreadPoolExecutor, Future, as_completed
from supabase import create_client, Client

if __name__ == '__main__':
    import google_store as gs
    import helpers as helpers
    from helpers import get_supabase, get_st_client

else:
    import modules.google_store as gs
    import modules.helpers as helpers
    from modules.helpers import get_supabase, get_st_client


def fetch_job_attachments(job_id: str, _client: ServiceTitanClient) -> List[Dict[str, Any]]:
    """Retrieve attachment metadata for the given job ID.

    This function calls the attachments listing endpoint for a job and
    returns a list of attachment objects (dictionaries).  Each
    attachment is expected to contain an ``id`` and ``fileName`` key.
    """
    attachments_url = _client.build_url("forms", f"jobs/{job_id}/attachments", version=2)
    attachments = _client.get_all(attachments_url)
    return attachments


def fetch_attachment_bytes(attachment_id: int, _client: ServiceTitanClient) -> bytes:
    """Download an attachment and return its raw bytes.

    The ``attachment_id`` is passed to the URL builder to form
    ``forms/v2/tenant/{tenant}/jobs/attachment/{id}``, which should
    return the binary content.
    """
    url = _client.build_url("forms", "jobs/attachment", resource_id=attachment_id)
    return _client.get(url)

def download_attachment_to_gcs(attachment_id: int, _client: ServiceTitanClient, gcs_bucket: str, gcs_blob: str) -> str:
    """Downloads an attachment to gcs and return a signed url to access
    """
    data = fetch_attachment_bytes(attachment_id, _client)
    url = gs.upload_bytes_to_gcs_signed(data, gcs_bucket, gcs_blob)
    # print("signed_url:")
    # print(url)
    return url

def download_attachments_for_job(job_id: str, client: ServiceTitanClient, sb_client: Client):
    """Download all attachments for a job and group them by type.

    This helper performs two API calls: one to list attachments and a
    second to download each attachment.  It groups attachments into
    ``imgs`` and ``pdfs`` by default (using the same extension map as
    :func:`group_attachments_by_type`).  Each entry in the returned
    dictionary is a list of ``(filename, data)`` tuples, where ``data``
    is the raw bytes of the attachment.  If no attachments exist for a
    category, the list will be empty.
    """

    GCS_BUCKET = 'prestigious-doc-check-attachments'
    # GCS_BUCKET = 'doc-check-attachments'

    attachments = fetch_job_attachments(job_id, client)

    count_urls = 0
    
    max_workers = min(8, len(attachments)) or 1
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        future_map: Dict[Future[bytes], Tuple] = {}
        # for category, att_id, filename, file_date, file_by in tasks:
        for att in attachments:
            file_name = att.get("fileName")
            att_id = att.get("id")
            file_date = att.get("createdOn")
            file_by = att.get("createdById")
            if not file_name or att_id is None:
                continue
            fut = pool.submit(download_attachment_to_gcs, att_id, client, GCS_BUCKET, f'{client.tenant}/{job_id}/{file_name}')
            future_map[fut] = (file_name, file_date, file_by, att_id)
        for fut in as_completed(future_map):
            file_name, file_date, file_by, att_id = future_map[fut]
            # try:
            # If error is raised, will be dealt with in main
            signed_url = fut.result()
            count_urls += 1
            # except Exception as e:
            #     # Set job as errored status here ??
            #     print(f"EXCEPTION: {e}")
            #     signed_url = 'error'
            print(f"uploading {att_id}...")
            response = (
                sb_client.table("gcs_attachments")
                .upsert({
                    "job_id": job_id, 
                    "type": helpers.get_attachment_type(file_name), 
                    "gcs_uploaded": datetime.now().isoformat(), 
                    "url": signed_url, 
                    "tenant": client.tenant,
                    "file_date": file_date, 
                    "file_name": file_name, 
                    "file_by": file_by, 
                    "attachment_id": att_id
                },
                on_conflict='attachment_id')
                .execute()
            # handle response ??
            ) 
            print(f"uploaded {att_id}.")
    return count_urls
    # grouped_meta = helpers.group_attachments_by_type(attachments)
    # result: Dict[str, List[Tuple[str, Any]]] = {key: [] for key in grouped_meta}
    # # Download images and PDFs.  Use a ThreadPoolExecutor to parallelise
    # # downloads across all attachments regardless of type.
    # tasks: List[Tuple[str, int, str]] = []  # (category, att_id, filename)
    # for category, items in grouped_meta.items():
    #     for filename, att_id, file_date, file_by in items:
    #         tasks.append((category, att_id, filename, file_date, file_by))
    # if not tasks:
    #     return result
    # max_workers = min(8, len(tasks)) or 1
    # with ThreadPoolExecutor(max_workers=max_workers) as pool:
    #     future_map: Dict[Future[bytes], Tuple[str, str]] = {}
    #     for category, att_id, filename, file_date, file_by in tasks:
    #         fut = pool.submit(download_attachment_to_gcs, att_id, client, GCS_BUCKET, f'{job_id}/{filename}')
    #         future_map[fut] = (category, filename, file_date, file_by)
    #     for fut in as_completed(future_map):
    #         category, filename, file_date, file_by = future_map[fut]
    #         try:
    #             signed_url = fut.result()
    #         except Exception as e:
    #             print(f"EXCEPTION: {e}")
    #             signed_url = None
    #         result[category].append((filename, client.from_utc_string(file_date), file_by, signed_url))
    # return result

if __name__ == '__main__':
    print(download_attachments_for_job(143554308, get_st_client('bravogolf'), get_supabase()))



