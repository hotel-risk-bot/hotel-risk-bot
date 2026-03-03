#!/bin/bash
set -e
PORT=${PORT:-8080}
echo "Starting gunicorn on port $PORT..."
exec gunicorn web_app:app --bind 0.0.0.0:$PORT --timeout 300 --workers 1 --log-level info
