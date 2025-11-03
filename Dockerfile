# Multi-stage Dockerfile for Streamlit app
FROM python:3.11-slim AS base

ENV PIP_NO_CACHE_DIR=1     PYTHONDONTWRITEBYTECODE=1     PYTHONUNBUFFERED=1

WORKDIR /app
COPY requirements.txt ./
RUN pip install --upgrade pip && pip install -r requirements.txt

# Copy app code
COPY . .

# Expose Streamlit default port
EXPOSE 8501

# Streamlit will read .streamlit/config.toml if present
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
