# Imagen para el Validador de Expediente (Flask + OCR)
# Se usa Docker (en vez de un buildpack normal de Python) porque necesitamos
# instalar tesseract-ocr, el paquete de idioma español, y poppler-utils a
# nivel de sistema operativo — eso no lo instala "pip install".

FROM python:3.11-slim

# Paquetes de sistema: tesseract (OCR) + idioma español + poppler (para
# convertir páginas de PDF a imagen antes de mandarlas al OCR).
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    tesseract-ocr-spa \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Render (y la mayoría de hostings) inyectan la variable PORT en tiempo de
# ejecución; 10000 es solo un valor por default para pruebas locales con Docker.
ENV PORT=10000
EXPOSE 10000

# gunicorn con 2 workers y timeout largo: el OCR de documentos escaneados
# puede tardar varios segundos por archivo, así que le damos margen.
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT} --workers 2 --timeout 120 app:app"]
