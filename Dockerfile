# Tailory Backend V2.1 — image Docker
# libreoffice-draw   : rasterisation des images EMF/WMF/TIFF des DOCX
# libreoffice-writer : conversion des documents ODT (et doc/rtf) en PDF
FROM python:3.11-slim
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-draw \
        libreoffice-writer \
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
