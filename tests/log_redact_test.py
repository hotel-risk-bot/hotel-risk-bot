"""Tests for log_redact.py (secrets never reach the logs).

Run with plain Python:   python tests/log_redact_test.py
or with pytest:          pytest tests/log_redact_test.py

All tokens below are made up; they only have the shape of the real ones.
"""
import io
import logging
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(HERE))

import log_redact  # noqa: E402

FAKE_BOT = "1234567890:AAFakeTokenFakeTokenFakeTokenFake1234"
FAKE_OPENAI = "sk-proj-FAKEfakeFAKEfakeFAKEfakeFAKEfake0000"
FAKE_ANTHROPIC = "sk-ant-api03-FAKEfakeFAKEfakeFAKEfake0000"
FAKE_AIRTABLE = "patABCDEFGHIJKLMN." + "0123456789abcdef" * 4
FAKE_GOOGLE = "AIza" + "FakeGoogleKeyFakeGoogleKey0123456"


def _capture():
    """Fresh root logger writing to a buffer, with the filter installed."""
    root = logging.getLogger()
    for h in list(root.handlers):
        root.removeHandler(h)
    buf = io.StringIO()
    handler = logging.StreamHandler(buf)
    handler.setFormatter(logging.Formatter("%(name)s %(levelname)s %(message)s"))
    root.addHandler(handler)
    root.setLevel(logging.INFO)
    logging.getLogger("httpx").setLevel(logging.NOTSET)
    log_redact.install()
    return buf


def test_httpx_request_lines_are_dropped():
    buf = _capture()
    logging.getLogger("httpx").info(
        'HTTP Request: POST https://api.telegram.org/bot%s/getUpdates "HTTP/1.1 200 OK"', FAKE_BOT)
    out = buf.getvalue()
    assert out == "", out


def test_token_masked_in_other_loggers():
    buf = _capture()
    log = logging.getLogger("bot")
    log.info("calling https://api.telegram.org/bot%s/sendMessage", FAKE_BOT)
    log.info("keys %s %s %s %s", FAKE_OPENAI, FAKE_ANTHROPIC, FAKE_AIRTABLE, FAKE_GOOGLE)
    out = buf.getvalue()
    for secret in (FAKE_BOT, FAKE_OPENAI, FAKE_ANTHROPIC, FAKE_AIRTABLE, FAKE_GOOGLE):
        assert secret not in out, out
    assert "bot<redacted>/sendMessage" in out, out
    assert "patABCDEFGHIJKLMN.<redacted>" in out, out


def test_token_masked_in_traceback():
    buf = _capture()
    try:
        raise ConnectionError(
            "HTTPSConnectionPool(host='api.telegram.org'): Max retries exceeded with url: /bot%s/sendMessage"
            % FAKE_BOT)
    except ConnectionError:
        logging.getLogger("loss_run_organizer").exception("Telegram send failed")
    out = buf.getvalue()
    assert FAKE_BOT not in out, out
    assert "Telegram send failed" in out and "bot<redacted>" in out, out


def test_warnings_from_httpx_still_show_redacted():
    buf = _capture()
    logging.getLogger("httpx").warning("retrying https://api.telegram.org/bot%s/getMe", FAKE_BOT)
    out = buf.getvalue()
    assert FAKE_BOT not in out and "bot<redacted>/getMe" in out, out


def test_plain_messages_untouched():
    buf = _capture()
    logging.getLogger("web_app").info("Generated proposal for %s (%d pages)", "Vista Lodging", 42)
    assert "Generated proposal for Vista Lodging (42 pages)" in buf.getvalue()


def test_install_is_idempotent():
    _capture()
    log_redact.install()
    log_redact.install()
    handler = logging.getLogger().handlers[0]
    assert sum(isinstance(f, log_redact.SecretRedactingFilter) for f in handler.filters) == 1


if __name__ == "__main__":
    tests = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    for t in tests:
        t()
        print("ok  ", t.__name__)
    print("%d passed" % len(tests))
