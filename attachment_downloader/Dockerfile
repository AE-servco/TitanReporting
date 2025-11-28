# Use official lightweight Python image
FROM python:3.11-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Set working directory
WORKDIR /attachment_downloader

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app code
COPY . .

# ---- Security: run as non-root user ----
RUN useradd -m workeruser
USER workeruser

# ---- Expose Gunicorn port ----
EXPOSE 8080

# ---- Start FastAPI with Gunicorn+Uvicorn ----
#  -k uvicorn.workers.UvicornWorker → proper ASGI worker
#  --workers 1 keeps concurrency predictable; Cloud Run scales horizontally
#
# If you want >1 worker for CPU-bound stuff, raise it.
CMD ["gunicorn", "main:app", \
     "-k", "uvicorn.workers.UvicornWorker", \
     "--bind", "0.0.0.0:8080", \
     "--workers", "1", \
     "--timeout", "900"]
