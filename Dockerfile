FROM python:3.14-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    TZ=Asia/Dhaka \
    REPORTS_DIR=/app/Reports

WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt

RUN useradd --uid 1000 --create-home app \
    && mkdir -p /app/Reports \
    && chown -R app:app /app/Reports
COPY --chown=app:app . .
USER app

EXPOSE 8501
HEALTHCHECK --interval=30s --timeout=10s --start-period=45s --retries=3 \
  CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8501/carton-supply-report/_stcore/health', timeout=5)"

CMD ["python", "-m", "streamlit", "run", "app.py", "--server.address=0.0.0.0", "--server.port=8501", "--server.baseUrlPath=carton-supply-report"]
