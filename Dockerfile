FROM python:3.11-slim

WORKDIR /app

# Install deps first for layer caching
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code only — data files (Excel, startups/, JSON) belong on
# a persistent volume (EFS) or S3, set via DATA_DIR env var at runtime
COPY app.py ai_tasks.py users.py email_utils.py ./
COPY templates/ templates/

EXPOSE 8080
ENV PORT=8080

CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:8080", "--workers", "1", "--timeout", "120"]
