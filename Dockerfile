FROM python:3.11-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY deck_core.py service.py template.pptx ./

ENV PORT=8080
EXPOSE 8080

# gunicorn: 1 worker, timeout ยาวเผื่อดาวน์โหลดรูปเยอะ
CMD ["sh", "-c", "gunicorn --bind 0.0.0.0:${PORT} --workers 1 --timeout 300 service:app"]
