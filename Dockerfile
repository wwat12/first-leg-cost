FROM python:3.12-slim

WORKDIR /app

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=8765

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Writable dirs for session uploads/exports (ephemeral on free tiers)
RUN mkdir -p uploads exports && chmod -R 777 uploads exports

EXPOSE 8765
CMD gunicorn app:app --bind 0.0.0.0:${PORT} --workers 2 --timeout 120
