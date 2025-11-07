import streamlit as st
from concurrent.futures import ThreadPoolExecutor, Future, as_completed
from typing import List, Tuple, Dict, Optional
from servicetitan_api_client import ServiceTitanClient
from google.cloud import secretmanager

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp"}

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

def get_client(state) -> ServiceTitanClient:
    @st.cache_resource
    def _create_client(state) -> ServiceTitanClient:
            state_code = state_codes()[state]
            client = ServiceTitanClient(
                app_key=get_secret("ST_app_key_tester"), 
                tenant=get_secret(f"ST_tenant_id_{state_code}"), 
                client_id=get_secret(f"ST_client_id_{state_code}"), 
                client_secret=get_secret(f"ST_client_secret_{state_code}"), 
                environment="production"
            )
            return client
    return _create_client(state)

@st.cache_data(show_spinner=False)
def fetch_job_attachments(job_id: str, _client: ServiceTitanClient) -> List[Dict[str, str]]:
    # forms/v2/tenant/{tenant}/jobs/{job_id}/attachments
    attachments_url = _client.build_url("forms", f"jobs/{job_id}/attachments", version=2)
    data = _client.get(attachments_url)
    return data.get("data", []) if isinstance(data, dict) else []

def filter_image_attachments(attachments: List[Dict[str, str]]) -> List[Tuple[str, int]]:
    images = []
    for att in attachments:
        name = att.get("fileName")
        att_id = att.get("id")
        if name and att_id:
            ext = name.lower().rsplit(".", 1)[-1] if "." in name else ""
            if f".{ext}" in IMAGE_EXTENSIONS:
                images.append((name, int(att_id)))
    return images

@st.cache_data(show_spinner=False)
def fetch_image_bytes(attachment_id: int, _client: ServiceTitanClient) -> bytes:
    url = _client.build_url("forms", "jobs/attachment", resource_id=attachment_id)
    return _client.get(url)


def main() -> None:
    st.set_page_config(page_title="ServiceTitan Job Images", layout="wide")
    st.title("ServiceTitan Job Image Viewer")
    client = get_client('NSW')

    # # Sidebar inputs
    # with st.sidebar:
    #     st.header("Job Selection")
    #     job_id = st.text_input("Current Job ID", value="143302553")
    #     next_job_id = st.text_input("Next Job ID", value="143812043")
    #     load_button = st.button("Load Job", type="primary")
    
    # Initialize session state flags
    if "prefetch_job_id" not in st.session_state:
        st.session_state.prefetch_job_id: Optional[str] = None
    if "prefetch_ready" not in st.session_state:
        st.session_state.prefetch_ready: bool = False

    if st.session_state.prefetch_ready:
        job_ready = st.session_state.prefetch_job_id
        st.sidebar.success(f"Images for job {job_ready} have been prefetched!")
        st.session_state.prefetch_ready = False
        st.session_state.prefetch_job_id = None
    
    # Sidebar inputs
    with st.sidebar:
        job_id = st.text_input("Current Job ID", value="143302553")
        next_job_id = st.text_input("Next Job ID", value="143812043")
        load_button = st.button("Load Job", type="primary")

    # Session state for prefetching
    if "prefetched" not in st.session_state:
        st.session_state.prefetched: Dict[str, List[Tuple[str, bytes]]] = {}
    if "prefetch_future" not in st.session_state:
        st.session_state.prefetch_future: Optional[Future] = None

    def prefetch(job: str) -> None:
        try:
            attachments = fetch_job_attachments(job, client)
            images_meta = filter_image_attachments(attachments)
            images: List[Tuple[str, bytes]] = []

            from concurrent.futures import ThreadPoolExecutor, as_completed
            with ThreadPoolExecutor(max_workers=min(8, len(images_meta)) or 1) as pool:
                future_to_name = {}
                for filename, att_id in images_meta:
                    fut = pool.submit(fetch_image_bytes, att_id, client)
                    future_to_name[fut] = filename
                for fut in as_completed(future_to_name):
                    filename = future_to_name[fut]
                    try:
                        img_bytes = fut.result()
                    except Exception:
                        continue
                    images.append((filename, img_bytes))
        finally:
            # store result in session state
            st.session_state.prefetched[job] = images
            print(f"prefetched done for {job}")

    if load_button and job_id:
        st.session_state.pop("current_images", None)

        # 1. Start prefetch immediately
        if next_job_id:
            st.session_state.prefetch_job_id = next_job_id
            st.session_state.prefetch_ready = False
            if next_job_id not in st.session_state.prefetched:
                executor = ThreadPoolExecutor(max_workers=1)
                st.session_state.prefetch_future = executor.submit(prefetch, next_job_id)

        # 2. Download current job images concurrently
        with st.spinner("Fetching attachments..."):
            attachments = fetch_job_attachments(job_id, client)
            image_meta = filter_image_attachments(attachments)
            images = []
            max_workers = min(8, len(image_meta)) or 1
            with ThreadPoolExecutor(max_workers=max_workers) as pool:
                future_to_name = {}
                for filename, att_id in image_meta:
                    fut = pool.submit(fetch_image_bytes, att_id, client)
                    future_to_name[fut] = filename
                for fut in as_completed(future_to_name):
                    filename = future_to_name[fut]
                    try:
                        img_bytes = fut.result()
                    except Exception:
                        continue
                    images.append((filename, img_bytes))
        st.session_state.current_images = images
            

    if "current_images" in st.session_state:
        images = st.session_state.current_images
        if images:
            st.subheader(f"Images for Job {job_id}")
            cols = st.columns(3)
            for idx, (filename, img_bytes) in enumerate(images):
                with cols[idx % 3]:
                    st.image(img_bytes, caption=filename, width=150)
        else:
            st.info("No image attachments found for this job.")

    if st.sidebar.button("Next Job") and next_job_id:
        future: Optional[Future] = st.session_state.prefetch_future
        if future is not None:
            try:
                future.result(timeout=0)
            except Exception:
                pass
        images = st.session_state.prefetched.get(next_job_id)
        if images is None:
            # fallback if prefetch not ready
            attachments = fetch_job_attachments(next_job_id, client)
            image_meta = filter_image_attachments(attachments)
            images = [
                (filename, fetch_image_bytes(url, client))
                for filename, url in image_meta
            ]
        st.session_state.current_images = images
        st.rerun()

if __name__ == "__main__":
    main()
