FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    STREAMLIT_SERVER_HEADLESS=true \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app.py cloud_app.py drive_sync.py ./

CMD ["sh", "-c", "streamlit run cloud_app.py --server.address=0.0.0.0 --server.port=${PORT:-8080} --server.headless=true"]
