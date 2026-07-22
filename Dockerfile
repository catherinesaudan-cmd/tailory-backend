# Tailory backend — image avec LibreOffice pour la conversion ODT→PDF
FROM python:3.11-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer libreoffice-draw libreoffice-impress \
        fonts-dejavu fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .

# app.py lit la variable PORT fournie par Render
CMD ["python", "app.py"]
