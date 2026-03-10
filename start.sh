#!/bin/bash
# Start both the Telegram bot and the web app
# The bot runs in the background, gunicorn runs in the foreground

echo "Starting Telegram bot in background..."
python bot.py &
BOT_PID=$!

echo "Starting web app with gunicorn..."
gunicorn web_app:app --bind 0.0.0.0:${PORT:-8080} --workers 1 --timeout 120 --graceful-timeout 30

# If gunicorn exits, also kill the bot
kill $BOT_PID 2>/dev/null
