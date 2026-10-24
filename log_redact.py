"""Keep secrets out of the logs.

python-telegram-bot sends every Telegram call through httpx, and httpx logs each
request URL at INFO ("HTTP Request: POST https://api.telegram.org/bot<TOKEN>/...").
Railway keeps those lines, so the bot token ended up readable in the deploy logs
(found Oct 24 2026 during the Railway credential rotation).

install() does two things:
  * raises the httpx / httpcore loggers to WARNING, so routine request lines are
    no longer written at all;
  * adds a filter to every root handler that masks Telegram bot tokens, OpenAI and
    Anthropic keys, Airtable tokens and Google API keys in any message or traceback
    that still reaches the logs (an exception message can carry a request URL).

Call it once, right after logging.basicConfig().
"""
import logging
import re

_PATTERNS = [
    (re.compile(r"bot\d{6,12}:[A-Za-z0-9_-]{30,}"), "bot<redacted>"),
    (re.compile(r"\b\d{6,12}:AA[A-Za-z0-9_-]{30,}"), "<redacted-telegram-token>"),
    (re.compile(r"sk-ant-[A-Za-z0-9_-]{20,}"), "sk-ant-<redacted>"),
    (re.compile(r"sk-(?:proj-|svcacct-|admin-)?[A-Za-z0-9_-]{20,}"), "sk-<redacted>"),
    (re.compile(r"(pat[A-Za-z0-9]{14})\.[A-Za-z0-9]{20,}"), r"\1.<redacted>"),
    (re.compile(r"AIza[0-9A-Za-z_-]{30,}"), "AIza<redacted>"),
]

_QUIET_LOGGERS = ("httpx", "httpcore")


def redact(text):
    """Return text with every known secret pattern masked."""
    if not text:
        return text
    for pattern, replacement in _PATTERNS:
        text = pattern.sub(replacement, text)
    return text


class SecretRedactingFilter(logging.Filter):
    """Masks secrets in the final message and in any attached traceback."""

    def filter(self, record):
        try:
            message = record.getMessage()
        except Exception:
            return True
        cleaned = redact(message)
        if cleaned != message:
            record.msg = cleaned
            record.args = None
        if record.exc_info and not record.exc_text:
            record.exc_text = logging.Formatter().formatException(record.exc_info)
        if record.exc_text:
            record.exc_text = redact(record.exc_text)
        return True


_FILTER = SecretRedactingFilter()


def install():
    """Quiet the HTTP client loggers and attach the redacting filter to root handlers."""
    for name in _QUIET_LOGGERS:
        logging.getLogger(name).setLevel(logging.WARNING)
    root = logging.getLogger()
    if not root.handlers:
        logging.basicConfig(level=logging.INFO)
    for handler in root.handlers:
        if _FILTER not in handler.filters:
            handler.addFilter(_FILTER)
