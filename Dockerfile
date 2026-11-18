FROM python:3.12-slim
RUN apt-get update && apt-get install -y --no-install-recommends ffmpeg poppler-utils tesseract-ocr tesseract-ocr-eng tesseract-ocr-hin && rm -rf /var/lib/apt/lists/*
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
ENV PYTHONUNBUFFERED=1
CMD gunicorn -w 1 --threads 4 -t 200 -b 0.0.0.0:${PORT:-10000} app:app
