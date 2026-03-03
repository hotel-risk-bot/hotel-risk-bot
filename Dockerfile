FROM python:3.11-slim

# Install poppler-utils for PDF processing (pdf2image)
RUN apt-get update && apt-get install -y --no-install-recommends \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements first for better caching
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY . .

# Railway sets PORT env var; use shell form so $PORT gets expanded
CMD gunicorn web_app:app --bind 0.0.0.0:$PORT --workers 2 --timeout 300 --log-level info
