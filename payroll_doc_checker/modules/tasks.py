# create_task.py
from google.cloud import tasks_v2
import datetime
import json

def create_task(url, job_id, tenant, force_refresh=False, project_id='servco1', queue='ST-attachment-download-queue', location='australia-southeast1'):
    client = tasks_v2.CloudTasksClient()
    parent = client.queue_path(project_id, location, queue)

    # Prepare payload
    payload = json.dumps({"job_id": job_id}).encode()
    payload = {
        "job_id": job_id,
        "tenant": tenant,
        "force_refresh": force_refresh
    }
    payload_bytes = json.dumps(payload).encode("utf-8")

    headers = {
        "accept": "application/json",
        "Content-Type": "application/json"
    }

    # Create task
    task = {
        "http_request": {
            "http_method": tasks_v2.HttpMethod.POST,
            "url": url,
            "headers": headers,
            "body": payload_bytes
        }
    }

    # if in_seconds > 0:
    #     d = datetime.datetime.utcnow() + datetime.timedelta(seconds=in_seconds)
    #     timestamp = timestamp_pb2.Timestamp()
    #     timestamp.FromDatetime(d)
    #     task["schedule_time"] = timestamp

    response = client.create_task(request={"parent": parent, "task": task})
    print(f"Created task: {response.name}")
