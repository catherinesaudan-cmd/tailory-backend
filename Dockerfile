# Tailory Backend V2 — image Docker
FROM python:3.11-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-draw \
        libreoffice-core \
        fonts-dejavu-core \
        fontconfig \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

CMD ["sh", "-c", "uvicorn app:app --host 0.0.0.0 --port ${PORT:-8000}"]
