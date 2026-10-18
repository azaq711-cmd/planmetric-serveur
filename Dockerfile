FROM python:3.12-slim

WORKDIR /app
ENV PYTHONUNBUFFERED=1

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py .
COPY static ./static

EXPOSE 10000
# 2 workers, timeout large : l'analyse de dossiers multi-feuilles peut prendre du temps
CMD ["sh", "-c", "gunicorn --workers 2 --timeout 300 --bind 0.0.0.0:${PORT:-10000} app:app"]
