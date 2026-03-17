FROM python:3.11-slim

# Install system dependencies
RUN apt-get update && apt-get install -y --no-install-recommends \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Cache bust - change this value to force fresh COPY
ARG CACHE_BUST=2026-03-17-v2

# Copy application code (fresh, not cached)
COPY . .

# Make entrypoint executable
RUN chmod +x entrypoint.sh

# Use entrypoint script (shell form ensures $PORT is expanded)
ENTRYPOINT ["./entrypoint.sh"]
