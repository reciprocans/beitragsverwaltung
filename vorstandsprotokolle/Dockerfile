FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 PYTHONUNBUFFERED=1

RUN apt-get update \
    && apt-get install -y --no-install-recommends libreoffice-writer libreoffice-calc fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

RUN mkdir -p /data
ENV DATA_DIR=/data
EXPOSE 8787
CMD ["gunicorn", "--bind", "0.0.0.0:8787", "--workers", "2", "--timeout", "120", "app:app"]
