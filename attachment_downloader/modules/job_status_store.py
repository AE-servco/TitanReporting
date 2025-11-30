# job_status_store.py
from typing import Optional
from supabase import Client
from datetime import datetime
# Replace this with real DB code (SQLAlchemy, Supabase, etc.)


def get_job_status(job_id: int, client: Client, tenant: str) -> Optional[int]:
    """
    Return one of: {-1,0,1,2} representing 'error', 'pending', 'processing', 'processed', respectively or None if record doesn't exist.
    """
    response = (
        client.table("gcs_job_attachment_status")
        .select("status")
        .eq("job_id", job_id)
        .eq("tenant", tenant)
        .execute()
    )
    try:
        if len(response.data) == 0:
            return 0
        if len(response.data) > 1:
            return -1
        else:
            return response.data[0]['status']
    except KeyError:
        return None


def set_job_status_processing(job_id: int, client: Client, time_now: datetime, tenant: str):
    """
    Insert or update job_status row to 'processing'.
    """
    response = (
        client.table("gcs_job_attachment_status")
        .upsert({"job_id": job_id, "status": 1, "last_update": time_now.isoformat(), "tenant": tenant, "error_msg": ""})
        .execute()
    )
    return


def set_job_status_processed(job_id: int, num_images: int, client: Client, time_now: datetime, tenant: str):
    """
    Set job_status to 'processed' with num_images.
    """
    response = (
        client.table("gcs_job_attachment_status")
        .upsert({"job_id": job_id, "status": 2, "last_update": time_now.isoformat(), "tenant": tenant, "error_msg": ""})
        .execute()
    )    
    return


def set_job_status_error(job_id: int, error_message: str, client: Client, time_now: datetime, tenant: str):
    """
    Set job_status to 'error' with error_message.
    """
    response = (
        client.table("gcs_job_attachment_status")
        .upsert({"job_id": job_id, "status": -1, "error_msg": error_message, "last_update": time_now.isoformat(), "tenant": tenant})
        .execute()
    )
    return
