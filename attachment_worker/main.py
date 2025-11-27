from fastapi import FastAPI, HTTPException, Request
from pydantic import BaseModel
from typing import List, Tuple, Dict, Any
import logging
from datetime import datetime
from zoneinfo import ZoneInfo

import modules.fetching as fetch


# from gcs_images import get_signed_image_urls_for_job  # from previous helper module

# # You will replace these with your real implementations:
# from servicetitan_client import fetch_job_image_attachments
from modules.job_status_store import (
    get_job_status,
    set_job_status_processing,
    set_job_status_processed,
    set_job_status_error,
)
from modules.helpers import get_supabase, get_st_client

app = FastAPI()
logger = logging.getLogger("worker")


# -------------------------------------------------------------------
# JSON CONTRACT MODELS
# -------------------------------------------------------------------

class ProcessJobRequest(BaseModel):
    job_id: int
    tenant: str
    force_refresh: bool = False


class ProcessJobResponse(BaseModel):
    status: str
    job_id: int
    num_images: int | None = None
    message: str | None = None


# -------------------------------------------------------------------
# CORE PROCESSING LOGIC
# -------------------------------------------------------------------

def process_job_attachments(job_id: int, force_refresh: bool, tenant: str) -> ProcessJobResponse:
    """
    Core logic:
    - check status
    - fetch attachments from ServiceTitan
    - upload to GCS
    - mark as processed
    """
    sb_client = get_supabase()
    st_client = get_st_client(tenant)

    status_map = {
        -1: 'error',
        0: 'pending',
        1: 'processing',
        2: 'processed',
    }

    current_status = status_map[get_job_status(job_id, sb_client, tenant)]

    if current_status == "processed" and not force_refresh:
        return ProcessJobResponse(
            status="skipped_already_processed",
            job_id=job_id,
            num_images=None,
            message="Job already processed; not refreshing.",
        )

    try:
        set_job_status_processing(job_id, sb_client, datetime.now(ZoneInfo("Australia/Sydney")), tenant)

        # Download attachments to GCS, insert urls and other data to supabase DBs
        count_of_urls = fetch.download_attachments_for_job(job_id, st_client, sb_client)

        # 2. Upload to GCS and generate signed URLs (or you can just ignore URLs here)
        # urls = get_signed_image_urls_for_job(job_id, raw_attachments)

        # 3. Mark as processed in status store
        set_job_status_processed(job_id, count_of_urls, sb_client, datetime.now(ZoneInfo("Australia/Sydney")), tenant)

        return ProcessJobResponse(
            status="processed",
            job_id=job_id,
            num_images=count_of_urls,
        )

    except Exception as e:
        logger.exception(f"Error processing job {job_id}: {e}")
        set_job_status_error(job_id, str(e), sb_client, datetime.now(ZoneInfo("Australia/Sydney")), tenant)

        raise


# -------------------------------------------------------------------
# HTTP ENDPOINT FOR CLOUD TASKS
# -------------------------------------------------------------------

@app.post("/tasks/process-job", response_model=ProcessJobResponse)
async def process_job_endpoint(req: ProcessJobRequest, request: Request):
    """
    Cloud Tasks will POST here with JSON payload matching ProcessJobRequest.
    """
    # (Optional) verify request headers/auth here to ensure it came from Cloud Tasks
    print(f'request caught for {req.job_id}')
    try:
        result = process_job_attachments(req.job_id, req.force_refresh, req.tenant)
        return result
    except Exception as e:
        # Cloud Tasks will treat non-2xx as failure and retry based on config
        raise HTTPException(status_code=500, detail=f"Processing failed: {e}")

if __name__ == '__main__':
    print(process_job_attachments(143554308, True, 'bravogolf'))
