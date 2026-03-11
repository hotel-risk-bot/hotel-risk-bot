"""
Configuration and environment variables for the Hotel Risk Advisor Bot.
Extracted from bot.py for maintainability.
"""

import os
import logging

logger = logging.getLogger(__name__)

# ── Core Credentials ─────────────────────────────────────────────────────
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN") or os.environ.get("TELEGRAM_BOT_TOKEN", "")
AIRTABLE_PAT = os.environ.get("AIRTABLE_PAT", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

# ── Airtable Base IDs ────────────────────────────────────────────────────
SALES_BASE_ID = os.environ.get("SALES_BASE_ID", "appnFKEzmdLbR4CHY")
CONSULTING_BASE_ID = os.environ.get("CONSULTING_BASE_ID", "appOVp1eJUPbNgNXM")

# ── Consulting Table IDs ─────────────────────────────────────────────────
INCIDENTS_TABLE_ID = os.environ.get("INCIDENTS_TABLE_ID", "tblK0V4q84B2hBNra")
ACTIVITY_TABLE_ID = os.environ.get("ACTIVITY_TABLE_ID", "tblESDnmgggtni5a3")
LOCATIONS_TABLE_ID = os.environ.get("LOCATIONS_TABLE_ID", "tbl6f73KwsL4OKzCJ")
CLIENT_TABLE_ID = os.environ.get("CLIENT_TABLE_ID", "tblO0GeWB6DocUA3e")

# ── Sales Table IDs ──────────────────────────────────────────────────────
OPPORTUNITIES_TABLE_ID = os.environ.get("OPPORTUNITIES_TABLE_ID", "tblMKuUsG1cosdQPN")
TASKS_TABLE_ID = os.environ.get("TASKS_TABLE_ID", "tblJVBL95e6qUJud3")
TODO_TABLE_ID = os.environ.get("TODO_TABLE_ID", "tbllOVUzN1oGCrEV7")

# ── API Endpoints ────────────────────────────────────────────────────────
AIRTABLE_API_URL = "https://api.airtable.com/v0"

# ── Optional Module Availability ─────────────────────────────────────────
# These flags are set based on whether optional dependencies import successfully.
# They're used to conditionally enable bot commands.

try:
    from apscheduler.schedulers.asyncio import AsyncIOScheduler  # noqa: F401
    from apscheduler.triggers.cron import CronTrigger  # noqa: F401
    HAS_SCHEDULER = True
except ImportError:
    HAS_SCHEDULER = False

try:
    from sheets_manager import (  # noqa: F401
        get_active_tasks, add_active_task, complete_task,
        get_completed_tasks_today, add_new_business, get_new_business,
        add_lead, get_leads, initialize_sheets,
    )
    HAS_SHEETS = True
except ImportError:
    HAS_SHEETS = False

try:
    from daily_briefing import (  # noqa: F401
        run_morning_briefing, run_afternoon_debrief,
        fetch_upcoming_renewals, classify_renewals,
        send_telegram_message, escape_telegram_markdown,
    )
    HAS_BRIEFING = True
except ImportError:
    HAS_BRIEFING = False

try:
    from marketing_summary import get_marketing_summary  # noqa: F401
    HAS_MARKETING = True
except ImportError:
    HAS_MARKETING = False

try:
    from marketing_update_generator import generate_marketing_update  # noqa: F401
    HAS_MARKETING_UPDATE = True
except ImportError:
    HAS_MARKETING_UPDATE = False

try:
    from proposal_handler import (  # noqa: F401
        get_proposal_conversation_handler, extract_standalone, generate_standalone,
    )
    HAS_PROPOSAL = True
except Exception as _proposal_err:
    HAS_PROPOSAL = False
    logger.warning(f"Proposal module import failed: {_proposal_err}")

try:
    from loss_run_organizer import (  # noqa: F401
        scheduled_organize, organize_loss_runs, send_organize_summary,
        tracker_get_client, tracker_get_all,
    )
    HAS_LOSS_ORGANIZER = True
except ImportError as _lr_err:
    HAS_LOSS_ORGANIZER = False
    logger.warning(f"Loss run organizer import failed: {_lr_err}")
