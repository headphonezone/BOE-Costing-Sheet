# Build from the repo root: docker build -f backend/Dockerfile -t boe-backend .
FROM python:3.12-slim

WORKDIR /app

COPY backend/requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY boe_parser.py .
COPY backend/ backend/

ENV PORT=8080
CMD ["sh", "-c", "uvicorn backend.main:app --host 0.0.0.0 --port ${PORT}"]
