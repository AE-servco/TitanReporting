# Use official lightweight Python image
FROM python:3.11-slim

RUN apt-get update && apt-get install -y \
    git \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

# Set environment variables
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Set working directory
WORKDIR /commission_exporter

# Install system dependencies (for some Python libs)
RUN apt-get update && apt-get install -y build-essential

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app code
COPY . .

# Streamlit settings: disable telemetry, set server headless, listen on port 8080
ENV STREAMLIT_SERVER_HEADLESS=true
ENV STREAMLIT_SERVER_PORT=8080
ENV STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

# Expose the port
EXPOSE 8080

# Run Streamlit app
CMD ["streamlit", "run", "Commission Exporter.py"]
