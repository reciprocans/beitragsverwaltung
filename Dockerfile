FROM python:3.12-slim
WORKDIR /app
ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1 DATA_DIR=/data
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt && useradd --uid 10001 --create-home appuser && mkdir /data && chown appuser:appuser /data
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-writer libreoffice-calc fonts-dejavu-core && rm -rf /var/lib/apt/lists/*
COPY app.py ./
COPY wsgi.py ./
COPY vorstandsprotokolle ./vorstandsprotokolle
COPY templates ./templates
COPY static ./static
USER appuser
EXPOSE 8000
CMD ["gunicorn", "--workers", "1", "--threads", "4", "--bind", "0.0.0.0:8000", "wsgi:application"]
