#!/usr/bin/env python3
"""
HUB International Hotel Risk Advisor Bot (@hotelriskadvisorbot)
Telegram bot for querying Airtable Sales and Consulting databases.
Deployment-ready version using Airtable REST API directly.

Commands:
  /start        - Welcome message
  /help         - Show available commands
  /update       - Get task list
  /status       - View progress
  /add          - Add task (Client | Task | Priority)
  /sales        - Search Sales System
  /consulting   - Search Consulting System (Incidents/Claims)
  /report       - Generate executive PDF report
"""

import os
import json
import logging
import re
import tempfile
import asyncio
from datetime import datetime, date

import requests as http_requests
import unicodedata
from telegram import Update, InputFile
from report_generator import generate_executive_pdf as generate_enhanced_pdf
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

try:
    from apscheduler.schedulers.asyncio import AsyncIOScheduler
    from apscheduler.triggers.cron import CronTrigger
    HAS_SCHEDULER = True
except ImportError:
    HAS_SCHEDULER = False

try:
    from sheets_manager import (
        get_active_tasks, add_active_task, complete_task,
        get_completed_tasks_today, add_new_business, get_new_business,
        add_lead, get_leads, initialize_sheets,
    )
    HAS_SHEETS = True
except ImportError:
    HAS_SHEETS = False

try:
    from daily_briefing import (
        run_morning_briefing, run_afternoon_debrief,
        fetch_upcoming_renewals, classify_renewals,
        send_telegram_message, escape_telegram_markdown,
    )
    HAS_BRIEFING = True
except ImportError:
    HAS_BRIEFING = False

try:
    from marketing_summary import get_marketing_summary
    HAS_MARKETING = True
except ImportError:
    HAS_MARKETING = False

try:
    from proposal_handler import get_proposal_conversation_handler, extract_standalone, generate_standalone
    HAS_PROPOSAL = True
except Exception as _proposal_err:
    HAS_PROPOSAL = False
    logging.getLogger(__name__).warning(f"Proposal module import failed: {_proposal_err}")

# ── Configuration (from environment variables) ────────────────────────────
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")
AIRTABLE_PAT = os.environ.get("AIRTABLE_PAT", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

# Airtable Base IDs
SALES_BASE_ID = "appnFKEzmdLbR4CHY"
CONSULTING_BASE_ID = "appOVp1eJUPbNgNXM"

# Consulting table IDs
INCIDENTS_TABLE_ID = "tblK0V4q84B2hBNra"
ACTIVITY_TABLE_ID = "tblESDnmgggtni5a3"
LOCATIONS_TABLE_ID = "tbl6f73KwsL4OKzCJ"
CLIENT_TABLE_ID = "tblO0GeWB6DocUA3e"

# Sales table IDs
OPPORTUNITIES_TABLE_ID = "tblMKuUsG1cosdQPN"
TASKS_TABLE_ID = "tblJVBL95e6qUJud3"
TODO_TABLE_ID = "tbllOVUzN1oGCrEV7"

AIRTABLE_API_URL = "https://api.airtable.com/v0"


def sanitize_for_pdf(text: str) -> str:
    """Replace non-latin-1 characters with safe ASCII equivalents for PDF output."""
    if not text:
        return text
    replacements = {
        '\u2013': '-',   # en-dash
        '\u2014': '--',  # em-dash
        '\u2018': "'",   # left single quote
        '\u2019': "'",   # right single quote
        '\u201c': '"',   # left double quote
        '\u201d': '"',   # right double quote
        '\u2026': '...', # ellipsis
        '\u2022': '*',   # bullet
        '\u00a0': ' ',   # non-breaking space
        '\u200b': '',    # zero-width space
        '\u2010': '-',   # hyphen
        '\u2011': '-',   # non-breaking hyphen
        '\u2012': '-',   # figure dash
        '\u00b7': '*',   # middle dot
        '\u2032': "'",   # prime
        '\u2033': '"',   # double prime
        '\u00ae': '(R)', # registered
        '\u2122': '(TM)',# trademark
        '\u00a9': '(C)', # copyright
    }
    for char, replacement in replacements.items():
        text = text.replace(char, replacement)
    # Fallback: replace any remaining non-latin-1 chars
    result = []
    for ch in text:
        try:
            ch.encode('latin-1')
            result.append(ch)
        except UnicodeEncodeError:
            # Try to get ASCII equivalent via NFKD decomposition
            decomposed = unicodedata.normalize('NFKD', ch)
            ascii_chars = decomposed.encode('ascii', 'ignore').decode('ascii')
            result.append(ascii_chars if ascii_chars else '?')
    return ''.join(result)

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)


async def safe_reply(update, text, parse_mode=None):
    """Send a reply, splitting into chunks if text exceeds Telegram's 4096 char limit."""
    MAX_LEN = 4000
    if len(text) <= MAX_LEN:
        await update.message.reply_text(text, parse_mode=parse_mode)
        return
    # Split at newlines
    chunks = []
    current = ""
    for line in text.split("\n"):
        if len(current) + len(line) + 1 > MAX_LEN:
            if current:
                chunks.append(current)
            current = line
        else:
            current = current + "\n" + line if current else line
    if current:
        chunks.append(current)
    for chunk in chunks:
        await update.message.reply_text(chunk, parse_mode=parse_mode)



# Ã¢ÂÂÃ¢ÂÂ Telegram Helpers Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

def escape_telegram_dollars(text: str) -> str:
    """Escape dollar signs to prevent Telegram from rendering them as LaTeX.
    Telegram treats $...$ as inline math. We escape $ to prevent this."""
    if not text:
        return text
    # Replace $ with escaped version to prevent LaTeX rendering
    return text.replace("$", "\\$")


async def safe_reply_text(message, text: str, parse_mode: str = None, **kwargs):
    """Send a reply with dollar signs escaped to prevent LaTeX rendering."""
    if parse_mode and parse_mode.lower() == "markdown":
        text = escape_telegram_dollars(text)
    try:
        await message.reply_text(text, parse_mode=parse_mode, **kwargs)
    except Exception as e:
        # If Markdown fails, try without parse_mode
        logger.warning(f"Markdown send failed: {e}, retrying without parse_mode")
        try:
            # Remove markdown formatting and try plain text
            plain = text.replace("*", "").replace("_", "").replace("`", "")
            plain = plain.replace("\\$", "$")  # Restore escaped dollars for plain text
            await message.reply_text(plain)
        except Exception as e2:
            logger.error(f"Plain text send also failed: {e2}")


# Ã¢ÂÂÃ¢ÂÂ Airtable REST API Functions Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

def airtable_headers():
    return {
        "Authorization": f"Bearer {AIRTABLE_PAT}",
        "Content-Type": "application/json",
    }


def airtable_list_records(base_id: str, table_id: str,
                          filter_formula: str = None,
                          max_records: int = 100,
                          sort_field: str = None,
                          sort_direction: str = "desc") -> list:
    """List records from an Airtable table with optional filter."""
    url = f"{AIRTABLE_API_URL}/{base_id}/{table_id}"
    all_records = []
    params = {}

    if filter_formula:
        params["filterByFormula"] = filter_formula
    if max_records:
        params["pageSize"] = min(max_records, 100)
    if sort_field:
        params["sort[0][field]"] = sort_field
        params["sort[0][direction]"] = sort_direction

    offset = None
    while True:
        if offset:
            params["offset"] = offset

        try:
            resp = http_requests.get(url, headers=airtable_headers(), params=params, timeout=30)
            resp.raise_for_status()
            data = resp.json()

            records = data.get("records", [])
            all_records.extend(records)

            offset = data.get("offset")
            if not offset or len(all_records) >= max_records:
                break
        except Exception as e:
            logger.error(f"Airtable API error: {e}")
            break

    return all_records[:max_records]


def airtable_search_records(base_id: str, table_id: str,
                            search_term: str, max_records: int = 20) -> list:
    """Search records using a simple text search across all fields.
    Airtable REST API doesn't have a native search, so we list and filter."""
    # For sales, we use a FIND formula across key fields
    safe_term = search_term.replace('"', '\\"')
    formula = (
        f'OR('
        f'FIND(LOWER("{safe_term}"), LOWER(ARRAYJOIN({{Opportunity Name}}, ","))),'
        f'FIND(LOWER("{safe_term}"), LOWER(ARRAYJOIN({{Opportunity Corporate Name}}, ","))),'
        f'FIND(LOWER("{safe_term}"), LOWER(ARRAYJOIN({{DBA}}, ",")))'
        f')'
    )
    return airtable_list_records(base_id, table_id, filter_formula=formula, max_records=max_records)


def airtable_create_record(base_id: str, table_id: str, fields: dict) -> dict | None:
    """Create a new record in an Airtable table."""
    url = f"{AIRTABLE_API_URL}/{base_id}/{table_id}"
    payload = {"fields": fields}

    try:
        resp = http_requests.post(url, headers=airtable_headers(), json=payload, timeout=30)
        resp.raise_for_status()
        return resp.json()
    except Exception as e:
        logger.error(f"Airtable create error: {e}")
        return None


# Ã¢ÂÂÃ¢ÂÂ Consulting Query Functions Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

def build_filter_formula(client_name: str, status: str = None,
                         min_incurred: float = None, max_incurred: float = None,
                         claim_type: str = None, min_policy_year: int = None) -> str:
    """Build an Airtable filterByFormula string for Incidents Claims."""
    safe_name = client_name.replace('"', '\\"')

    name_conditions = [
        f'FIND(LOWER("{safe_name}"), LOWER(ARRAYJOIN({{Client Name}}, ",")))',
        f'FIND(LOWER("{safe_name}"), LOWER(ARRAYJOIN({{Corporate Name}}, ",")))',
        f'FIND(LOWER("{safe_name}"), LOWER(ARRAYJOIN({{DBA (from Location)}}, ",")))',
        f'FIND(LOWER("{safe_name}"), LOWER(ARRAYJOIN({{Companies}}, ",")))',
    ]
    name_filter = f'OR({", ".join(name_conditions)})'

    conditions = [name_filter]

    if status:
        conditions.append(f'{{Status}} = "{status.title()}"')

    if min_incurred is not None:
        conditions.append(f'{{Incurred}} >= {min_incurred}')

    if max_incurred is not None:
        conditions.append(f'{{Incurred}} <= {max_incurred}')

    if claim_type:
        conditions.append(f'{{Claim Type}} = "{claim_type.title()}"')

    if min_policy_year is not None:
        conditions.append(f'VALUE(ARRAYJOIN({{Policy Year}}, ",")) >= {min_policy_year}')

    if len(conditions) == 1:
        return conditions[0]
    return f'AND({", ".join(conditions)})'


def search_incidents(client_name: str, status: str = None,
                     min_incurred: float = None, max_incurred: float = None,
                     claim_type: str = None, min_policy_year: int = None) -> list:
    """Search Incidents Claims table with full filter support."""
    formula = build_filter_formula(client_name, status, min_incurred, max_incurred,
                                   claim_type, min_policy_year)
    logger.info(f"Airtable filter formula: {formula}")

    records = airtable_list_records(
        CONSULTING_BASE_ID, INCIDENTS_TABLE_ID,
        filter_formula=formula, max_records=100,
    )

    if not records:
        return []

    results = []
    for rec in records:
        fields = rec.get("fields", {})

        incurred = fields.get("Incurred", 0) or 0
        if isinstance(incurred, list):
            incurred = incurred[0] if incurred else 0
        try:
            incurred = float(incurred)
        except (ValueError, TypeError):
            incurred = 0.0

        results.append({
            "record_id": rec.get("id", ""),
            "fields": fields,
            "incurred": incurred,
        })

    results.sort(key=lambda x: x["incurred"], reverse=True)
    return results


# Ã¢ÂÂÃ¢ÂÂ Claims Development Parser Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

def parse_claims_development(raw_data: str) -> list:
    """Parse Activity Rollup Raw Data to extract valuation entries."""
    if not raw_data:
        return []

    entries = raw_data.split("[[Break]]")
    valuations = []

    for entry in entries:
        entry = entry.strip().strip(",").strip()
        if not entry:
            continue

        if "Valuation" not in entry and "Total Incurred:" not in entry:
            continue

        date_match = re.match(r'([\w]+\s+\d{1,2},\s+\d{4})', entry)
        date_str = date_match.group(1) if date_match else "Unknown"

        paid_match = re.search(r'Paid:\s*\$?([\d,.]+)', entry)
        reserved_match = re.search(r'Reserved:\s*\$?([\d,.]+)', entry)
        expenses_match = re.search(r'Expenses:\s*\$?([\d,.]+)', entry)
        incurred_match = re.search(r'Total Incurred:\s*\$?([\d,.]+)', entry)

        def parse_amount(match):
            if match:
                try:
                    return float(match.group(1).replace(",", ""))
                except (ValueError, TypeError):
                    return 0.0
            return 0.0

        total_incurred = parse_amount(incurred_match)

        if total_incurred > 0 or paid_match or reserved_match:
            valuations.append({
                "date": date_str,
                "paid": parse_amount(paid_match),
                "reserved": parse_amount(reserved_match),
                "expenses": parse_amount(expenses_match),
                "total_incurred": total_incurred,
            })

    return valuations


def format_claims_development(valuations: list) -> str:
    """Format claims development valuations into a readable progression."""
    if not valuations:
        return ""

    lines = ["Ã°ÂÂÂ *Claims Development*"]

    for v in valuations:
        detail_parts = []
        if v["paid"] > 0:
            detail_parts.append(f"Paid: ${v['paid']:,.0f}")
        if v["reserved"] > 0:
            detail_parts.append(f"Rsv: ${v['reserved']:,.0f}")
        if v["expenses"] > 0:
            detail_parts.append(f"Exp: ${v['expenses']:,.0f}")

        detail_str = f" ({', '.join(detail_parts)})" if detail_parts else ""
        lines.append(f"Ã¢ÂÂ¢ {v['date']}: *${v['total_incurred']:,.0f}*{detail_str}")

    return "\n".join(lines)


# Ã¢ÂÂÃ¢ÂÂ Report Formatting Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

def get_val(f: dict, field_name: str, default: str = "N/A") -> str:
    """Extract a field value from an Airtable record, handling lists."""
    val = f.get(field_name, default)
    if isinstance(val, list):
        return ", ".join(str(v) for v in val) if val else default
    return str(val) if val else default


def format_claim_report(rec: dict) -> str:
    """Format a single claim record into a Telegram-friendly report."""
    f = rec["fields"]

    claim_num = get_val(f, "Claim #")
    status = get_val(f, "Status")
    claim_type = get_val(f, "Claim Type")
    policy_type = get_val(f, "Policy Type")

    corporate_name = get_val(f, "Corporate Name")
    dba = get_val(f, "DBA (from Location)")
    client_name = get_val(f, "Client Name")

    address = get_val(f, "Address (from Location)")
    city = get_val(f, "City (from Location)")
    state = get_val(f, "State (from Location)")
    zipcode = get_val(f, "ZIP (from Location)")

    incident_date = get_val(f, "Incident Date")
    if incident_date == "N/A":
        incident_date = get_val(f, "DOL")

    policy_year = get_val(f, "Policy Year")

    cause_of_loss = get_val(f, "Cause of Loss Rollup Output")
    if cause_of_loss == "N/A":
        cause_of_loss = get_val(f, "Cause of Loss (from Cause of Loss)")
    risk_hazard = get_val(f, "Risk/Hazard (From Risk/Hazard)")
    brief_desc = get_val(f, "Brief Description")
    summary_of_facts = get_val(f, "Summary of Facts")

    involved_party = get_val(f, "Involved Party (From Involved Party)")
    if involved_party == "N/A":
        involved_party = get_val(f, "Involved Party copy")

    location_of_incident = get_val(f, "Location of Incident")

    incurred = rec.get("incurred", 0)
    paid = get_val(f, "Paid - Rollup")
    reserved_vals = f.get("Reserved Helper", [])
    if isinstance(reserved_vals, list) and reserved_vals:
        reserved = reserved_vals[-1] if reserved_vals else 0
    else:
        reserved = reserved_vals or 0

    expenses = get_val(f, "Expenses Helper")

    carrier = get_val(f, "Carrier")
    if carrier == "N/A":
        carrier = get_val(f, "Carrier (from Policies)")

    policy_num = get_val(f, "Policy # (from Policies)")

    attorney_rep = f.get("Attorney Representation", False)
    attorney_demand = get_val(f, "Attorney Demand")

    status_emoji = "Ã¢ÂÂ" if status == "Open" else "Ã°ÂÂÂ´" if status == "Closed" else "Ã¢ÂÂª"

    # Build the report Ã¢ÂÂ matching exact user-requested format
    lines = []
    lines.append(f"{'Ã¢ÂÂ' * 35}")

    # Ã¢ÂÂÃ¢ÂÂ Date of Loss (first) Ã¢ÂÂÃ¢ÂÂ
    lines.append(f"Ã°ÂÂÂ *Date of Loss:* {incident_date}")
    lines.append("")

    # Ã¢ÂÂÃ¢ÂÂ Claim Details (with location info grouped here) Ã¢ÂÂÃ¢ÂÂ
    lines.append(f"Ã°ÂÂÂ *Claim Details*")
    lines.append(f"Claim #: `{claim_num}`")
    lines.append(f"Status: {status_emoji} {status}")
    lines.append(f"Type: {claim_type}")
    if policy_type != "N/A" and policy_type != claim_type:
        lines.append(f"Policy Type: {policy_type}")
    if policy_year != "N/A":
        lines.append(f"Policy Year: {policy_year}")
    lines.append(f"Property: {dba}")
    lines.append(f"Corporate Name: {corporate_name}")
    lines.append(f"Company: {client_name}")
    if address != "N/A":
        full_addr = f"{address}"
        if city != "N/A":
            full_addr += f", {city}"
        if state != "N/A":
            full_addr += f", {state}"
        if zipcode != "N/A":
            full_addr += f" {zipcode}"
        lines.append(f"Address: {full_addr}")
    lines.append("")

    # Ã¢ÂÂÃ¢ÂÂ Incident Details Ã¢ÂÂÃ¢ÂÂ
    lines.append(f"Ã°ÂÂÂ *Incident Details*")
    lines.append(f"Claimant: {involved_party}")
    lines.append(f"Cause of Loss: {cause_of_loss}")
    if risk_hazard != "N/A":
        lines.append(f"Ã¢ÂÂ Ã¯Â¸Â Hazard: {risk_hazard}")
    if location_of_incident != "N/A":
        lines.append(f"Ã°ÂÂÂ¢ Location of Incident: {location_of_incident}")
    if brief_desc != "N/A":
        lines.append(f"Description: {brief_desc}")
    lines.append("")

    # Ã¢ÂÂÃ¢ÂÂ Financial Summary Ã¢ÂÂÃ¢ÂÂ
    lines.append(f"Ã°ÂÂÂ° *Financial Summary*")
    lines.append(f"Ã¢ÂÂ¢ Total Incurred: ${incurred:,.0f}")
    if paid != "N/A":
        try:
            lines.append(f"Ã¢ÂÂ¢ Paid: ${float(paid):,.0f}")
        except (ValueError, TypeError):
            lines.append(f"Ã¢ÂÂ¢ Paid: {paid}")
    try:
        lines.append(f"Ã¢ÂÂ¢ Reserved: ${float(reserved):,.0f}")
    except (ValueError, TypeError):
        lines.append(f"Ã¢ÂÂ¢ Reserved: {reserved}")
    if expenses != "N/A":
        try:
            exp_vals = f.get("Expenses Helper", [])
            if isinstance(exp_vals, list) and exp_vals:
                lines.append(f"Ã¢ÂÂ¢ Expenses: ${float(exp_vals[-1]):,.0f}")
            else:
                lines.append(f"Ã¢ÂÂ¢ Expenses: ${float(expenses):,.0f}")
        except (ValueError, TypeError):
            pass
    lines.append("")

    # Ã¢ÂÂÃ¢ÂÂ Claims Development Ã¢ÂÂÃ¢ÂÂ
    raw_activity = f.get("Activity Rollup Raw Data", "")
    if raw_activity:
        valuations = parse_claims_development(raw_activity)
        if valuations:
            dev_text = format_claims_development(valuations)
            lines.append(dev_text)
            lines.append("")

    # Ã¢ÂÂÃ¢ÂÂ Summary of Facts Ã¢ÂÂÃ¢ÂÂ
    if summary_of_facts != "N/A" and len(summary_of_facts) > 5:
        sf = summary_of_facts[:500]
        if len(summary_of_facts) > 500:
            sf += "..."
        lines.append(f"Ã°ÂÂÂ *Summary of Facts:*")
        lines.append(sf)
        lines.append("")

    # Ã¢ÂÂÃ¢ÂÂ Attorney Representation Ã¢ÂÂÃ¢ÂÂ
    if attorney_rep:
        lines.append(f"Ã¢ÂÂÃ¯Â¸Â *Attorney Representation:* Yes")
        if attorney_demand != "N/A":
            lines.append(f"Attorney Demand: ${attorney_demand}")
        lines.append("")

    # Ã¢ÂÂÃ¢ÂÂ Carrier / Policy Ã¢ÂÂÃ¢ÂÂ
    if carrier != "N/A":
        lines.append(f"Carrier: {carrier}")
    if policy_num != "N/A":
        lines.append(f"Policy #: {policy_num}")

    return "\n".join(lines)


# Ã¢ÂÂÃ¢ÂÂ Sales System Functions Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

def search_sales(query: str) -> list:
    """Search the Sales System (Opportunities table)."""
    return airtable_search_records(SALES_BASE_ID, OPPORTUNITIES_TABLE_ID, query, max_records=20)


def format_sales_record(rec: dict) -> str:
    """Format a sales opportunity record."""
    f = rec.get("fields", {})

    opp_name = get_val(f, "Opportunity Name")
    if opp_name == "N/A":
        opp_name = get_val(f, "Opportunity Corporate Name")
    dba = get_val(f, "DBA")
    status = get_val(f, "Status")
    market_status = get_val(f, "Market Status")
    eff_date = get_val(f, "Effective Date")
    revenue = get_val(f, "Revenue")
    exp_revenue = get_val(f, "Expiring Revenue")
    nr = get_val(f, "N/R")

    lines = []
    lines.append(f"{'Ã¢ÂÂ' * 35}")
    lines.append(f"Ã°ÂÂÂ¢ *{opp_name}*")
    if dba != "N/A":
        lines.append(f"DBA: {dba}")
    lines.append(f"Status: {status}")
    if market_status != "N/A":
        lines.append(f"Market Status: {market_status}")
    if eff_date != "N/A":
        lines.append(f"Effective Date: {eff_date}")
    if nr != "N/A":
        lines.append(f"New/Renewal: {nr}")
    if revenue != "N/A":
        try:
            lines.append(f"Revenue: ${float(revenue):,.0f}")
        except (ValueError, TypeError):
            lines.append(f"Revenue: {revenue}")
    if exp_revenue != "N/A":
        try:
            lines.append(f"Expiring Revenue: ${float(exp_revenue):,.0f}")
        except (ValueError, TypeError):
            lines.append(f"Expiring Revenue: {exp_revenue}")

    return "\n".join(lines)


# Ã¢ÂÂÃ¢ÂÂ Argument Parser Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

def parse_consulting_args(raw_args: list) -> dict:
    """Parse consulting command arguments with natural language support."""
    full_text = " ".join(raw_args)

    # Remove symbols
    full_text = full_text.replace(">", " ").replace("<", " ").replace("=", " ")

    # Replace natural language phrases with markers
    nl_replacements = [
        (r'\bgreater\s+than\b', '__GT__'),
        (r'\bmore\s+than\b', '__GT__'),
        (r'\babove\b', '__GT__'),
        (r'\bover\b', '__GT__'),
        (r'\bexceeding\b', '__GT__'),
        (r'\bless\s+than\b', '__LT__'),
        (r'\bbelow\b', '__LT__'),
        (r'\bunder\b', '__LT__'),
        (r'\bbetween\b', '__BETWEEN__'),
        (r'\band\b', '__AND__'),
        (r'\blast\s+(\d+)\s+years?\b', r'__LASTYEARS_\1__'),
        (r'\bsince\s+policy\s+year\s+(\d{4})\b', r'__SINCEPY_\1__'),
        (r'\bsince\s+(\d{4})\b', r'__SINCEPY_\1__'),
        (r'\bpolicy\s+year\s+(\d{4})\b', r'__SINCEPY_\1__'),
        (r'\bfrom\s+policy\s+year\s+(\d{4})\b', r'__SINCEPY_\1__'),
        (r'\bfrom\s+(\d{4})\b', r'__SINCEPY_\1__'),
    ]
    for pattern, replacement in nl_replacements:
        full_text = re.sub(pattern, replacement, full_text, flags=re.IGNORECASE)

    tokens = full_text.split()

    client_name = None
    status = None
    min_incurred = None
    max_incurred = None
    claim_type = None
    min_policy_year = None

    name_parts = []
    i = 0
    while i < len(tokens):
        token = tokens[i]
        lower = token.lower()

        if lower in ("open", "closed", "all"):
            status = lower if lower != "all" else None
            i += 1
            continue

        if lower in ("liability", "property"):
            claim_type = lower
            i += 1
            continue

        if lower == "only":
            i += 1
            continue

        lastyears_match = re.match(r'__LASTYEARS_(\d+)__', token)
        if lastyears_match:
            years = int(lastyears_match.group(1))
            min_policy_year = datetime.now().year - years
            i += 1
            continue

        sincepy_match = re.match(r'__SINCEPY_(\d{4})__', token)
        if sincepy_match:
            min_policy_year = int(sincepy_match.group(1))
            i += 1
            continue

        if token == "__GT__":
            i += 1
            while i < len(tokens):
                try:
                    val = float(tokens[i].replace(",", "").replace("$", ""))
                    min_incurred = val
                    i += 1
                    break
                except ValueError:
                    i += 1
            continue

        if token == "__LT__":
            i += 1
            while i < len(tokens):
                try:
                    val = float(tokens[i].replace(",", "").replace("$", ""))
                    max_incurred = val
                    i += 1
                    break
                except ValueError:
                    i += 1
            continue

        if token == "__BETWEEN__":
            i += 1
            nums = []
            while i < len(tokens) and len(nums) < 2:
                t = tokens[i]
                if t == "__AND__":
                    i += 1
                    continue
                try:
                    val = float(t.replace(",", "").replace("$", ""))
                    nums.append(val)
                except ValueError:
                    pass
                i += 1
            if len(nums) >= 2:
                min_incurred = min(nums)
                max_incurred = max(nums)
            elif len(nums) == 1:
                min_incurred = nums[0]
            continue

        if token == "__AND__":
            i += 1
            continue

        try:
            val = float(token.replace(",", "").replace("$", ""))
            if min_incurred is None:
                min_incurred = val
            elif max_incurred is None:
                max_incurred = val
            i += 1
            continue
        except ValueError:
            pass

        name_parts.append(token)
        i += 1

    client_name = " ".join(name_parts).strip()
    if not client_name:
        client_name = " ".join(raw_args)

    return {
        "client_name": client_name,
        "status": status,
        "min_incurred": min_incurred,
        "max_incurred": max_incurred,
        "claim_type": claim_type,
        "min_policy_year": min_policy_year,
    }


# Ã¢ÂÂÃ¢ÂÂ PDF Report Generator Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

def generate_executive_pdf(client_name: str, results: list, query_params: dict) -> str:
    """Generate an executive client report as PDF. Returns file path."""
    from fpdf import FPDF

    class ClaimPDF(FPDF):
        def header(self):
            self.set_font("Helvetica", "B", 11)
            self.set_text_color(0, 51, 102)
            self.cell(0, 8, "HUB International  |  Hotel Franchise Practice", ln=True, align="L")
            self.set_draw_color(0, 102, 204)
            self.set_line_width(0.5)
            self.line(10, self.get_y(), 200, self.get_y())
            self.ln(4)

        def footer(self):
            self.set_y(-15)
            self.set_font("Helvetica", "I", 8)
            self.set_text_color(128, 128, 128)
            self.cell(0, 10, f"Confidential  |  Page {self.page_no()}/{{nb}}  |  Generated {datetime.now().strftime('%m/%d/%Y')}", align="C")

    pdf = ClaimPDF()
    pdf.alias_nb_pages()
    pdf.set_auto_page_break(auto=True, margin=20)
    pdf.add_page()

    # Title
    pdf.set_font("Helvetica", "B", 18)
    pdf.set_text_color(0, 51, 102)
    pdf.cell(0, 12, "Executive Claims Report", ln=True)
    pdf.set_font("Helvetica", "", 14)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(0, 10, f"Client: {sanitize_for_pdf(client_name)}", ln=True)
    pdf.ln(2)

    # Query parameters
    pdf.set_font("Helvetica", "I", 10)
    pdf.set_text_color(80, 80, 80)
    filter_parts = []
    if query_params.get("status"):
        filter_parts.append(f"Status: {query_params['status'].title()}")
    if query_params.get("claim_type"):
        filter_parts.append(f"Type: {query_params['claim_type'].title()}")
    if query_params.get("min_incurred") is not None:
        filter_parts.append(f"Min Incurred: ${query_params['min_incurred']:,.0f}")
    if query_params.get("max_incurred") is not None:
        filter_parts.append(f"Max Incurred: ${query_params['max_incurred']:,.0f}")
    if query_params.get("min_policy_year") is not None:
        filter_parts.append(f"Policy Year >= {query_params['min_policy_year']}")
    if filter_parts:
        pdf.cell(0, 6, "Filters: " + " | ".join(filter_parts), ln=True)
    pdf.cell(0, 6, f"Report Date: {datetime.now().strftime('%B %d, %Y')}", ln=True)
    pdf.ln(4)

    # Executive Summary
    total_incurred = sum(r["incurred"] for r in results)
    total_paid = 0
    total_reserved = 0
    open_count = 0
    closed_count = 0
    attorney_count = 0

    for r in results:
        flds = r["fields"]
        if flds.get("Status") == "Open":
            open_count += 1
        else:
            closed_count += 1
        if flds.get("Attorney Representation"):
            attorney_count += 1
        p = flds.get("Paid - Rollup", 0)
        try:
            total_paid += float(p) if p else 0
        except (ValueError, TypeError):
            pass
        rv = flds.get("Reserved Helper", [])
        if isinstance(rv, list) and rv:
            try:
                total_reserved += float(rv[-1])
            except (ValueError, TypeError):
                pass
        elif rv:
            try:
                total_reserved += float(rv)
            except (ValueError, TypeError):
                pass

    pdf.set_fill_color(240, 245, 250)
    pdf.set_draw_color(0, 102, 204)
    pdf.rect(10, pdf.get_y(), 190, 36, style="DF")

    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(0, 51, 102)
    pdf.cell(0, 8, "  Executive Summary", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(0, 0, 0)

    col_w = 63
    pdf.cell(col_w, 6, f"  Total Claims: {len(results)}", ln=False)
    pdf.cell(col_w, 6, f"Open: {open_count}", ln=False)
    pdf.cell(col_w, 6, f"Closed: {closed_count}", ln=True)

    pdf.cell(col_w, 6, f"  Total Incurred: ${total_incurred:,.0f}", ln=False)
    pdf.cell(col_w, 6, f"Total Paid: ${total_paid:,.0f}", ln=False)
    pdf.cell(col_w, 6, f"Total Reserved: ${total_reserved:,.0f}", ln=True)

    pdf.cell(col_w, 6, f"  Attorney Rep: {attorney_count} claim(s)", ln=True)
    pdf.ln(8)

    # Claims Table Header
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_fill_color(0, 51, 102)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(25, 7, "DOL", fill=True, border=1)
    pdf.cell(30, 7, "Claim #", fill=True, border=1)
    pdf.cell(15, 7, "Status", fill=True, border=1)
    pdf.cell(20, 7, "Type", fill=True, border=1)
    pdf.cell(55, 7, "Property", fill=True, border=1)
    pdf.cell(22, 7, "Incurred", fill=True, border=1, align="R")
    pdf.cell(22, 7, "Paid", fill=True, border=1, align="R")
    pdf.ln()

    # Claims Table Rows
    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(0, 0, 0)
    for idx, r in enumerate(results):
        flds = r["fields"]
        bg = (248, 248, 248) if idx % 2 == 0 else (255, 255, 255)
        pdf.set_fill_color(*bg)

        dol = sanitize_for_pdf(get_val(flds, "Incident Date", ""))
        if dol == "N/A" or not dol:
            dol = sanitize_for_pdf(get_val(flds, "DOL", ""))
        cnum = sanitize_for_pdf(get_val(flds, "Claim #", ""))
        st = sanitize_for_pdf(get_val(flds, "Status", ""))
        ct = sanitize_for_pdf(get_val(flds, "Claim Type", ""))
        prop = sanitize_for_pdf(get_val(flds, "DBA (from Location)", ""))
        inc = r["incurred"]
        p = get_val(flds, "Paid - Rollup", "0")
        try:
            p_val = float(p)
        except (ValueError, TypeError):
            p_val = 0

        if prop and len(prop) > 28:
            prop = prop[:26] + ".."

        pdf.cell(25, 6, dol[:10], fill=True, border=1)
        pdf.cell(30, 6, cnum[:16], fill=True, border=1)
        pdf.cell(15, 6, st[:6], fill=True, border=1)
        pdf.cell(20, 6, ct[:10], fill=True, border=1)
        pdf.cell(55, 6, prop, fill=True, border=1)
        pdf.cell(22, 6, f"${inc:,.0f}", fill=True, border=1, align="R")
        pdf.cell(22, 6, f"${p_val:,.0f}", fill=True, border=1, align="R")
        pdf.ln()

    # Detailed Claims Section
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(0, 51, 102)
    pdf.cell(0, 10, "Detailed Claims Analysis", ln=True)
    pdf.ln(2)

    for idx, r in enumerate(results):
        flds = r["fields"]

        if pdf.get_y() > 240:
            pdf.add_page()

        dol = sanitize_for_pdf(get_val(flds, "Incident Date", "N/A"))
        if dol == "N/A":
            dol = sanitize_for_pdf(get_val(flds, "DOL", "N/A"))
        cnum = sanitize_for_pdf(get_val(flds, "Claim #", "N/A"))
        st = sanitize_for_pdf(get_val(flds, "Status", "N/A"))
        ct = sanitize_for_pdf(get_val(flds, "Claim Type", "N/A"))
        prop = sanitize_for_pdf(get_val(flds, "DBA (from Location)", "N/A"))
        corp = sanitize_for_pdf(get_val(flds, "Corporate Name", "N/A"))
        claimant = sanitize_for_pdf(get_val(flds, "Involved Party (From Involved Party)", "N/A"))
        if claimant == "N/A":
            claimant = sanitize_for_pdf(get_val(flds, "Involved Party copy", "N/A"))
        col_text = sanitize_for_pdf(get_val(flds, "Cause of Loss Rollup Output", "N/A"))
        if col_text == "N/A":
            col_text = sanitize_for_pdf(get_val(flds, "Cause of Loss (from Cause of Loss)", "N/A"))
        hazard = sanitize_for_pdf(get_val(flds, "Risk/Hazard (From Risk/Hazard)", "N/A"))
        loc_inc = sanitize_for_pdf(get_val(flds, "Location of Incident", "N/A"))
        brief = sanitize_for_pdf(get_val(flds, "Brief Description", "N/A"))
        atty = flds.get("Attorney Representation", False)
        carrier = sanitize_for_pdf(get_val(flds, "Carrier", "N/A"))
        if carrier == "N/A":
            carrier = sanitize_for_pdf(get_val(flds, "Carrier (from Policies)", "N/A"))

        # Claim header bar
        pdf.set_fill_color(0, 51, 102)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", "B", 10)
        status_icon = "OPEN" if st == "Open" else "CLOSED"
        pdf.cell(0, 7, f"  {dol}  |  {ct}  |  {cnum}  |  {status_icon}  |  ${r['incurred']:,.0f}", fill=True, ln=True)

        pdf.set_text_color(0, 0, 0)
        pdf.set_font("Helvetica", "", 9)

        pdf.cell(45, 5, f"Property: {prop}", ln=False)
        pdf.cell(0, 5, f"Corporate: {corp}", ln=True)
        pdf.cell(45, 5, f"Claimant: {claimant}", ln=False)
        pdf.cell(0, 5, f"Cause: {col_text}", ln=True)

        detail_line = ""
        if hazard != "N/A":
            detail_line += f"Hazard: {hazard}  |  "
        if loc_inc != "N/A":
            detail_line += f"Location: {loc_inc}  |  "
        if atty:
            detail_line += "Attorney: Yes  |  "
        if carrier != "N/A":
            detail_line += f"Carrier: {carrier}"
        if detail_line:
            pdf.cell(0, 5, detail_line.rstrip("  |  "), ln=True)

        if brief != "N/A":
            pdf.set_font("Helvetica", "I", 8)
            pdf.multi_cell(0, 4, f"Description: {brief[:200]}")

        # Financial progression
        raw_activity = flds.get("Activity Rollup Raw Data", "")
        if raw_activity:
            valuations = parse_claims_development(raw_activity)
            if valuations:
                pdf.set_font("Helvetica", "B", 8)
                pdf.cell(0, 5, "Claims Development:", ln=True)
                pdf.set_font("Helvetica", "", 8)
                for v in valuations:
                    parts = []
                    if v["paid"] > 0:
                        parts.append(f"Paid: ${v['paid']:,.0f}")
                    if v["reserved"] > 0:
                        parts.append(f"Rsv: ${v['reserved']:,.0f}")
                    if v["expenses"] > 0:
                        parts.append(f"Exp: ${v['expenses']:,.0f}")
                    detail = f" ({', '.join(parts)})" if parts else ""
                    pdf.cell(0, 4, f"    {v['date']}: ${v['total_incurred']:,.0f}{detail}", ln=True)

        pdf.ln(3)

    # Save PDF
    filepath = tempfile.mktemp(suffix=".pdf", prefix=f"claims_report_{client_name.replace(' ', '_')}_")
    pdf.output(filepath)
    return filepath


# Ã¢ÂÂÃ¢ÂÂ Command Handlers Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /start command."""
    welcome = (
        "Ã°ÂÂÂ¨ *Hotel Risk Advisor Bot*\n"
        "Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ\n"
        "Welcome! I can help you query the HUB International hotel insurance databases.\n\n"
        "*Claims & Reports:*\n"
        "Ã¢ÂÂ¢ /consulting `query` Ã¢ÂÂ Search Claims\n"
        "Ã¢ÂÂ¢ /report `query` Ã¢ÂÂ Executive PDF Report\n"
        "Ã¢ÂÂ¢ /sales `query` Ã¢ÂÂ Search Sales System\n"
        "Ã¢ÂÂ¢ /marketing `client` Ã¢ÂÂ Marketing Summary\n\n"
        "*Task Management:*\n"
        "Ã¢ÂÂ¢ /task `Client | Task | Priority` Ã¢ÂÂ Add task\n"
        "Ã¢ÂÂ¢ /done `number` Ã¢ÂÂ Complete a task\n"
        "Ã¢ÂÂ¢ /mytasks Ã¢ÂÂ View active tasks\n"
        "Ã¢ÂÂ¢ /update Ã¢ÂÂ Get Airtable task list\n"
        "Ã¢ÂÂ¢ /status Ã¢ÂÂ View progress\n\n"
        "*Business Development:*\n"
        "• /newbiz `Client | Desc | Revenue` — Add opportunity\n"
        "• /lead `Client | Contact | Source | Desc` — Add lead\n"
        "• /renewals — Upcoming renewals (120 days)\n"
        "• /pipeline — View new business pipeline\n\n"
        "*Proposal Generator:*\n"
        "• /proposal `Client Name` — Generate insurance proposal\n"
        "  Upload PDF quotes & Excel SOV, then /extract → /generate\n\n"
        "• /help — Show this message\n\n"
        "*Query Examples:*\n"
        "Ã¢ÂÂ¢ `/consulting Jasmin open liability`\n"
        "Ã¢ÂÂ¢ `/report Ocean Partners last 5 years`\n"
        "Ã¢ÂÂ¢ `/marketing Triton Hospitality`\n"
        "Ã¢ÂÂ¢ `/task Premier | Send loss runs | Urgent`\n"
        "Ã¢ÂÂ¢ `/newbiz Hilton Garden | Property pkg | 15000`\n"
    )
    await safe_reply_text(update.message, welcome, parse_mode="Markdown")


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await start_command(update, context)


async def update_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /update - Get task list from Sales system."""
    await update.message.reply_text("Ã¢ÂÂ³ Fetching task list from Sales System...")

    records = airtable_list_records(
        SALES_BASE_ID, TASKS_TABLE_ID,
        filter_formula='NOT({Task Status} = "Done")',
        max_records=20,
    )

    if not records:
        await update.message.reply_text("No open tasks found.")
        return

    lines = ["Ã°ÂÂÂ *Open Tasks*\n"]
    for rec in records:
        f = rec.get("fields", {})
        name = f.get("Name", "Unnamed")
        task_status = f.get("Task Status", "N/A")
        priority = f.get("Priority", "N/A")
        due = f.get("Due Date", "N/A")
        cam = f.get("CAM", "N/A")

        status_emoji = {"Todo": "Ã°ÂÂÂ´", "In progress": "Ã°ÂÂÂ¡"}.get(task_status, "Ã¢ÂÂª")
        pri_emoji = {"High": "Ã°ÂÂÂ¥", "Medium": "Ã¢ÂÂ¡", "Low": "Ã°ÂÂÂ¤"}.get(priority, "")

        lines.append(f"{status_emoji} {pri_emoji} *{name}*")
        if due != "N/A":
            lines.append(f"   Due: {due}")
        if cam != "N/A":
            lines.append(f"   Assigned: {cam}")
        lines.append("")

    msg = "\n".join(lines)
    if len(msg) > 4000:
        msg = msg[:4000] + "\n\n_...truncated_"
    await safe_reply_text(update.message, msg, parse_mode="Markdown")


async def status_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /status - View progress summary."""
    await update.message.reply_text("Ã¢ÂÂ³ Fetching status from Sales System...")

    records = airtable_list_records(
        SALES_BASE_ID, TASKS_TABLE_ID, max_records=100,
    )

    if not records:
        await update.message.reply_text("No tasks found.")
        return

    total = len(records)
    done = sum(1 for r in records if r.get("fields", {}).get("Task Status") == "Done")
    in_progress = sum(1 for r in records if r.get("fields", {}).get("Task Status") == "In progress")
    todo = sum(1 for r in records if r.get("fields", {}).get("Task Status") == "Todo")

    msg = (
        "Ã°ÂÂÂ *Task Progress*\n"
        "Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ\n"
        f"Total Tasks: {total}\n"
        f"Ã¢ÂÂ Done: {done}\n"
        f"Ã°ÂÂÂ¡ In Progress: {in_progress}\n"
        f"Ã°ÂÂÂ´ Todo: {todo}\n"
    )
    await safe_reply_text(update.message, msg, parse_mode="Markdown")


async def add_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /add Client | Task | Priority."""
    if not context.args:
        await update.message.reply_text(
            "Usage: `/add Client Name | Task Description | Priority`\n"
            "Priority: High, Medium, or Low",
            parse_mode="Markdown",
        )
        return

    text = " ".join(context.args)
    parts = [p.strip() for p in text.split("|")]

    if len(parts) < 2:
        await safe_reply_text(update.message,
            "Please use the format: `/add Client | Task | Priority`",
            parse_mode="Markdown",
            )
        return

    company = parts[0]
    task_desc = parts[1]
    priority = parts[2] if len(parts) > 2 else "Medium"

    valid_priorities = {"high": "High", "medium": "Medium", "low": "Low"}
    priority = valid_priorities.get(priority.lower(), "Medium")

    await update.message.reply_text(f"Ã¢ÂÂ³ Adding task for {company}...")

    result = airtable_create_record(SALES_BASE_ID, TODO_TABLE_ID, {
        "Notes": f"{task_desc}\n\nAdded via Telegram Bot on {datetime.now().strftime('%m/%d/%Y %I:%M %p')}",
        "Status": "Todo",
        "Priority": priority,
    })

    if result:
        await update.message.reply_text(
            f"Ã¢ÂÂ Task added successfully!\n"
            f"Ã°ÂÂÂ {task_desc}\n"
            f"Priority: {priority}",
        )
    else:
        await update.message.reply_text("Ã¢ÂÂ Failed to add task. Please try again.")


# Ã¢ÂÂÃ¢ÂÂ Consulting Query Handler Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

async def run_consulting_query(args: list) -> tuple:
    """Run a consulting query and return (results, params, query_desc)."""
    params = parse_consulting_args(args)

    query_desc = f"Client: *{params['client_name']}*"
    if params["status"]:
        query_desc += f" | Status: *{params['status'].title()}*"
    if params["claim_type"]:
        query_desc += f" | Type: *{params['claim_type'].title()}*"
    if params["min_incurred"] is not None:
        query_desc += f" | Min Incurred: *${params['min_incurred']:,.0f}*"
    if params["max_incurred"] is not None:
        query_desc += f" | Max Incurred: *${params['max_incurred']:,.0f}*"
    if params["min_policy_year"] is not None:
        query_desc += f" | Policy Year Ã¢ÂÂ¥ *{params['min_policy_year']}*"

    results = search_incidents(
        params["client_name"],
        params["status"],
        params["min_incurred"],
        params["max_incurred"],
        params["claim_type"],
        params["min_policy_year"],
    )

    return results, params, query_desc


async def consulting_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /consulting command."""
    if not context.args:
        await update.message.reply_text(
            "Ã°ÂÂÂ *Consulting System Query*\n\n"
            "Usage:\n"
            "Ã¢ÂÂ¢ `/consulting ClientName` Ã¢ÂÂ All claims\n"
            "Ã¢ÂÂ¢ `/consulting ClientName open` Ã¢ÂÂ Open claims\n"
            "Ã¢ÂÂ¢ `/consulting ClientName closed` Ã¢ÂÂ Closed claims\n"
            "Ã¢ÂÂ¢ `/consulting ClientName liability` Ã¢ÂÂ Liability only\n"
            "Ã¢ÂÂ¢ `/consulting ClientName property` Ã¢ÂÂ Property only\n"
            "Ã¢ÂÂ¢ `/consulting ClientName open liability greater than 25000`\n"
            "Ã¢ÂÂ¢ `/consulting ClientName last 5 years`\n"
            "Ã¢ÂÂ¢ `/consulting ClientName closed property last 3 years`\n\n"
            "Searches across Client Name, Corporate Name, DBA, and Company fields.",
            parse_mode="Markdown",
        )
        return

    await safe_reply_text(update.message, "Ã¢ÂÂ³ Searching Consulting System...", parse_mode="Markdown")

    results, params, query_desc = await run_consulting_query(context.args)

    if not results:
        await safe_reply_text(update.message,
            f"No claims found matching your criteria.\n{query_desc}",
            parse_mode="Markdown",
            )
        return

    total_incurred = sum(r["incurred"] for r in results)
    header = (
        f"Ã°ÂÂÂ¨ *Consulting System Results*\n"
        f"Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ\n"
        f"Found *{len(results)}* claim(s)\n"
        f"Total Incurred: *${total_incurred:,.0f}*\n"
        f"{query_desc}\n"
    )
    await safe_reply_text(update.message, header, parse_mode="Markdown")

    for i, rec in enumerate(results[:20]):
        try:
            report = format_claim_report(rec)
            if len(report) > 4000:
                mid = report.rfind("\n", 0, 4000)
                if mid < 0:
                    mid = 4000
                await safe_reply_text(update.message, report[:mid], parse_mode="Markdown")
                await safe_reply_text(update.message, report[mid:], parse_mode="Markdown")
            else:
                await safe_reply_text(update.message, report, parse_mode="Markdown")
        except Exception as e:
            logger.error(f"Error formatting claim: {e}")
            try:
                report = format_claim_report(rec)
                await update.message.reply_text(report)
            except Exception as e2:
                logger.error(f"Error sending claim (plain): {e2}")
                await update.message.reply_text(
                    f"Error displaying claim #{i+1}. Claim #: {rec['fields'].get('Claim #', 'N/A')}"
                )

    if len(results) > 20:
        await safe_reply_text(update.message,
            f"_Showing 20 of {len(results)} results. "
            f"Refine your search to see more specific results._",
            parse_mode="Markdown",
            )


# Ã¢ÂÂÃ¢ÂÂ Report Command Handler Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

async def report_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /report command Ã¢ÂÂ generate executive PDF report."""
    if not context.args:
        await safe_reply_text(update.message,
            "Ã°ÂÂÂ *Executive PDF Report*\n\n"
            "Usage: `/report ClientName [filters]`\n\n"
            "Examples:\n"
            "Ã¢ÂÂ¢ `/report Ocean Partners` Ã¢ÂÂ Full report\n"
            "Ã¢ÂÂ¢ `/report Jasmin open liability` Ã¢ÂÂ Filtered\n"
            "Ã¢ÂÂ¢ `/report Jasmin last 5 years`\n"
            "Ã¢ÂÂ¢ `/report Jasmin closed greater than 25000`\n",
            parse_mode="Markdown",
            )
        return

    await safe_reply_text(update.message, "Ã¢ÂÂ³ Generating executive PDF report...", parse_mode="Markdown")

    results, params, query_desc = await run_consulting_query(context.args)

    if not results:
        await safe_reply_text(update.message,
            f"No claims found matching your criteria.\n{query_desc}",
            parse_mode="Markdown",
            )
        return

    try:
        filepath = generate_enhanced_pdf(params["client_name"], results, params)

        total_incurred = sum(r["incurred"] for r in results)
        caption = (
            f"Ã°ÂÂÂ Executive Claims Report Ã¢ÂÂ {params['client_name']}\n"
            f"{len(results)} claims | Total Incurred: ${total_incurred:,.0f}"
        )

        with open(filepath, "rb") as pdf_file:
            await update.message.reply_document(
                document=InputFile(pdf_file, filename=f"Claims_Report_{params['client_name'].replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.pdf"),
                caption=caption,
            )

        os.unlink(filepath)

    except Exception as e:
        logger.error(f"Error generating PDF: {e}")
        await safe_reply_text(update.message,
            f"Ã¢ÂÂ Ã¯Â¸Â Error generating PDF report: {str(e)}\n\n"
            f"The query found *{len(results)}* claims. Try `/consulting` to view them in chat.",
            parse_mode="Markdown",
            )


# Ã¢ÂÂÃ¢ÂÂ Sales Query Handler Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

async def sales_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /sales command."""
    if not context.args:
        await safe_reply_text(update.message,
            "Ã°ÂÂÂ *Sales System Query*\n\n"
            "Usage: `/sales search term`\n\n"
            "Examples:\n"
            "Ã¢ÂÂ¢ `/sales Marriott`\n"
            "Ã¢ÂÂ¢ `/sales Best Western`\n"
            "Ã¢ÂÂ¢ `/sales Premier Resorts`",
            parse_mode="Markdown",
            )
        return

    query = " ".join(context.args)
    await safe_reply_text(update.message,
        f"Ã¢ÂÂ³ Searching Sales System for: *{query}*...",
        parse_mode="Markdown",
        )

    records = search_sales(query)

    if not records:
        await update.message.reply_text(f"No results found for: {query}")
        return

    header = (
        f"Ã°ÂÂÂ¢ *Sales System Results*\n"
        f"Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ\n"
        f"Found *{len(records)}* result(s) for: *{query}*\n"
    )
    await safe_reply_text(update.message, header, parse_mode="Markdown")

    for i, rec in enumerate(records[:10]):
        try:
            report = format_sales_record(rec)
            await safe_reply_text(update.message, report, parse_mode="Markdown")
        except Exception as e:
            logger.error(f"Error formatting sales record: {e}")
            try:
                report = format_sales_record(rec)
                await update.message.reply_text(report)
            except Exception:
                await update.message.reply_text(f"Error displaying result #{i+1}")

    if len(records) > 10:
        await safe_reply_text(update.message,
            f"_Showing 10 of {len(records)} results._",
            parse_mode="Markdown",
            )


# Ã¢ÂÂÃ¢ÂÂ Google Sheets Task Management Commands Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

async def task_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /task - Add task to Google Sheets Active Tasks."""
    if not HAS_SHEETS:
        await update.message.reply_text("Google Sheets integration not available.")
        return

    if not context.args:
        await update.message.reply_text(
            "Ã°ÂÂÂ *Add Task*\n\n"
            "Usage: `/task Client | Task Description | Priority`\n\n"
            "Priority: Urgent, Today, This Week, Medium, Low\n\n"
            "Examples:\n"
            "Ã¢ÂÂ¢ `/task Premier | Send loss runs to Zurich | Urgent`\n"
            "Ã¢ÂÂ¢ `/task Ocean Partners | Follow up on WC audit | This Week`\n"
            "Ã¢ÂÂ¢ `/task MGM | Review renewal proposal | Today`",
            parse_mode="Markdown",
        )
        return

    text = " ".join(context.args)
    parts = [p.strip() for p in text.split("|")]

    if len(parts) < 2:
        await safe_reply_text(update.message,
            "Please use: `/task Client | Task | Priority`",
            parse_mode="Markdown",
            )
        return

    client = parts[0]
    task_desc = parts[1]
    priority = parts[2] if len(parts) > 2 else "This Week"

    valid_priorities = {
        "urgent": "Urgent", "today": "Today", "asap": "Urgent",
        "this week": "This Week", "high": "This Week",
        "medium": "Medium", "low": "Low",
    }
    priority = valid_priorities.get(priority.lower().strip(), priority.strip())

    await update.message.reply_text(f"Ã¢ÂÂ³ Adding task for {client}...")

    success = add_active_task(client, task_desc, priority)

    if success:
        await safe_reply_text(update.message,
            f"Ã¢ÂÂ Task added to Active Tasks!\n\n"
            f"Ã°ÂÂÂ *{task_desc}*\n"
            f"Client: {client}\n"
            f"Priority: {priority}",
            parse_mode="Markdown",
            )
    else:
        await update.message.reply_text("Ã¢ÂÂ Failed to add task. Check Google Sheets connection.")


async def done_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /done - Mark a task as complete."""
    if not HAS_SHEETS:
        await update.message.reply_text("Google Sheets integration not available.")
        return

    if not context.args:
        # Show task list with numbers
        tasks = get_active_tasks()
        if not tasks:
            await update.message.reply_text("No active tasks to complete!")
            return

        lines = ["Ã°ÂÂÂ *Active Tasks Ã¢ÂÂ Select number to complete:*\n"]
        for i, t in enumerate(tasks, 1):
            lines.append(f"  {i}. [{t['client']}] {t['task']}")
        lines.append("\nUsage: `/done 1` to complete task #1")
        await safe_reply_text(update.message, "\n".join(lines), parse_mode="Markdown")
        return

    try:
        task_num = int(context.args[0])
    except ValueError:
        await safe_reply_text(update.message, "Please provide a task number: `/done 1`", parse_mode="Markdown")
        return

    completed = complete_task(task_num)
    if completed:
        await safe_reply_text(update.message,
            f"Ã¢ÂÂ Task completed!\n\n"
            f"*{completed['task']}*\n"
            f"Client: {completed['client']}\n"
            f"Moved to Completed Tasks tab.",
            parse_mode="Markdown",
            )
    else:
        await safe_reply_text(update.message,
            f"Ã¢ÂÂ Task #{task_num} not found. Use `/done` to see available tasks.",
            parse_mode="Markdown",
            )


async def mytasks_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /mytasks - Show active tasks from Google Sheet."""
    if not HAS_SHEETS:
        await update.message.reply_text("Google Sheets integration not available.")
        return

    await update.message.reply_text("Ã¢ÂÂ³ Fetching active tasks...")
    tasks = get_active_tasks()

    if not tasks:
        await update.message.reply_text("Ã°ÂÂÂ No active tasks. Inbox zero! Ã°ÂÂÂ")
        return

    # Group by priority
    urgent = [t for t in tasks if t.get("priority", "").lower() in ["urgent", "today", "asap"]]
    this_week = [t for t in tasks if t.get("priority", "").lower() in ["this week", "high"]]
    other = [t for t in tasks if t not in urgent and t not in this_week]

    lines = [f"Ã°ÂÂÂ *Active Tasks ({len(tasks)} total)*\n"]
    idx = 1

    if urgent:
        lines.append("Ã°ÂÂÂ¥ *URGENT/TODAY:*")
        for t in urgent:
            due = f" (Due: {t['due_date']})" if t.get('due_date') else ""
            lines.append(f"  {idx}. [{t['client']}] {t['task']}{due}")
            idx += 1
        lines.append("")

    if this_week:
        lines.append("Ã¢ÂÂ¡ *THIS WEEK:*")
        for t in this_week:
            due = f" (Due: {t['due_date']})" if t.get('due_date') else ""
            lines.append(f"  {idx}. [{t['client']}] {t['task']}{due}")
            idx += 1
        lines.append("")

    if other:
        lines.append("Ã°ÂÂÂ *OTHER:*")
        for t in other:
            due = f" (Due: {t['due_date']})" if t.get('due_date') else ""
            lines.append(f"  {idx}. [{t['client']}] {t['task']}{due}")
            idx += 1

    lines.append("\n_Use `/done #` to complete a task_")

    msg = "\n".join(lines)
    if len(msg) > 4000:
        msg = msg[:4000] + "\n\n_...truncated_"
    await safe_reply_text(update.message, msg, parse_mode="Markdown")


async def newbiz_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /newbiz - Add new business opportunity to Google Sheet."""
    if not HAS_SHEETS:
        await update.message.reply_text("Google Sheets integration not available.")
        return

    if not context.args:
        await safe_reply_text(update.message,
            "Ã°ÂÂÂ¼ *Add New Business Opportunity*\n\n"
            "Usage: `/newbiz Client | Description | Est Revenue`\n\n"
            "Examples:\n"
            "Ã¢ÂÂ¢ `/newbiz Hilton Garden Inn | Property & GL pkg | 15000`\n"
            "Ã¢ÂÂ¢ `/newbiz Best Western Plus | Full commercial pkg | 25000`",
            parse_mode="Markdown",
            )
        return

    text = " ".join(context.args)
    parts = [p.strip() for p in text.split("|")]

    if len(parts) < 2:
        await safe_reply_text(update.message,
            "Please use: `/newbiz Client | Description | Est Revenue`",
            parse_mode="Markdown",
            )
        return

    client = parts[0]
    description = parts[1]
    est_revenue = parts[2].strip() if len(parts) > 2 else ""

    # Determine N/R based on context
    nr = "N"  # Default to New

    success = add_new_business(client, description, nr, est_revenue)

    if success:
        rev_display = f" | Est Revenue: ${est_revenue}" if est_revenue else ""
        await safe_reply_text(update.message,
            f"Ã¢ÂÂ New business opportunity added!\n\n"
            f"Ã°ÂÂÂ¼ *{client}*\n"
            f"{description}{rev_display}\n"
            f"Type: New Business",
            parse_mode="Markdown",
            )
    else:
        await update.message.reply_text("Ã¢ÂÂ Failed to add opportunity. Check Google Sheets connection.")


async def lead_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /lead - Add a new lead to Google Sheet Leads tab."""
    if not HAS_SHEETS:
        await update.message.reply_text("Google Sheets integration not available.")
        return

    if not context.args:
        await safe_reply_text(update.message,
            "Ã°ÂÂÂ¯ *Add New Lead*\n\n"
            "Usage: `/lead Client | Contact | Source | Description`\n\n"
            "Examples:\n"
            "Ã¢ÂÂ¢ `/lead Marriott Courtyard | John Smith | Referral | 50 room property in Tampa`\n"
            "Ã¢ÂÂ¢ `/lead Holiday Inn Express | GM Jane | Cold Call | New build opening Q3`",
            parse_mode="Markdown",
            )
        return

    text = " ".join(context.args)
    parts = [p.strip() for p in text.split("|")]

    if len(parts) < 1:
        await safe_reply_text(update.message,
            "Please use: `/lead Client | Contact | Source | Description`",
            parse_mode="Markdown",
            )
        return

    client = parts[0]
    contact = parts[1] if len(parts) > 1 else ""
    source = parts[2] if len(parts) > 2 else ""
    description = parts[3] if len(parts) > 3 else ""

    success = add_lead(client, contact, source, description)

    if success:
        await safe_reply_text(update.message,
            f"Ã¢ÂÂ Lead added!\n\n"
            f"Ã°ÂÂÂ¯ *{client}*\n"
            f"Contact: {contact}\n"
            f"Source: {source}\n"
            f"{description}",
            parse_mode="Markdown",
            )
    else:
        await update.message.reply_text("Ã¢ÂÂ Failed to add lead. Check Google Sheets connection.")


async def pipeline_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /pipeline - View new business pipeline."""
    if not HAS_SHEETS:
        await update.message.reply_text("Google Sheets integration not available.")
        return

    await update.message.reply_text("Ã¢ÂÂ³ Fetching pipeline...")
    items = get_new_business()

    if not items:
        await update.message.reply_text("Ã°ÂÂÂ¼ No new business opportunities in pipeline.")
        return

    lines = [f"Ã°ÂÂÂ¼ *New Business Pipeline ({len(items)} opportunities)*\n"]
    total_revenue = 0
    for nb in items:
        rev = nb.get("est_revenue", "")
        rev_display = ""
        if rev:
            try:
                rev_val = float(str(rev).replace("$", "").replace(",", ""))
                total_revenue += rev_val
                rev_display = f" Ã¢ÂÂ ${rev_val:,.0f}"
            except (ValueError, TypeError):
                rev_display = f" Ã¢ÂÂ {rev}"

        nr_flag = "Ã°ÂÂÂ" if nb.get("nr") == "N" else "Ã°ÂÂÂ"
        lines.append(f"  {nr_flag} *{nb['client']}*{rev_display}")
        if nb.get("description"):
            lines.append(f"    {nb['description']}")
        lines.append(f"    Status: {nb.get('status', 'N/A')} | Added: {nb.get('date_added', '')}")
        lines.append("")

    if total_revenue > 0:
        lines.append(f"\nÃ°ÂÂÂ° *Total Pipeline: ${total_revenue:,.0f}*")

    msg = "\n".join(lines)
    if len(msg) > 4000:
        msg = msg[:4000] + "\n\n_...truncated_"
    await safe_reply_text(update.message, msg, parse_mode="Markdown")


# Ã¢ÂÂÃ¢ÂÂ Renewals Command Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

async def renewals_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /renewals - Show upcoming renewals from Airtable."""
    if not HAS_BRIEFING:
        await update.message.reply_text("Briefing module not available.")
        return

    await update.message.reply_text("Ã¢ÂÂ³ Fetching upcoming renewals from Sales System...")

    records = fetch_upcoming_renewals(120)
    if not records:
        await update.message.reply_text("No upcoming renewals found in the next 120 days.")
        return

    renewals_data = classify_renewals(records)
    all_renewals = renewals_data.get("all_renewals", [])
    submit_alerts = renewals_data.get("submit_alerts", [])
    high_revenue = renewals_data.get("high_revenue", [])

    lines = [f"Ã°ÂÂÂ *Upcoming Renewals Ã¢ÂÂ Next 120 Days ({len(all_renewals)} total)*\n"]

    if submit_alerts:
        lines.append(f"Ã°ÂÂÂ´ *EXPOSED Ã¢ÂÂ NEEDS SUBMISSION ({len(submit_alerts)}):*")
        for r in submit_alerts:
            days = f"{r['days_until']}d" if r.get('days_until') else "TBD"
            rev = f" Ã¢ÂÂ ${r['revenue']:,.0f}" if r.get('revenue') else ""
            lines.append(f"  Ã¢ÂÂ¼Ã¯Â¸Â [{days}] {r['name']}{rev}")
        lines.append("")

    if high_revenue:
        lines.append(f"Ã°ÂÂÂ° *HIGH REVENUE >$5K ({len(high_revenue)}):*")
        for r in high_revenue[:10]:
            days = f"{r['days_until']}d" if r.get('days_until') else "TBD"
            lines.append(f"  $ [{days}] {r['name']} Ã¢ÂÂ ${r['revenue']:,.0f} ({r['status']})")
        lines.append("")

    lines.append("*All Renewals:*")
    sorted_renewals = sorted(all_renewals, key=lambda x: x.get("days_until") or 999)
    for r in sorted_renewals[:25]:
        days = f"{r['days_until']}d" if r.get('days_until') else "TBD"
        rev = f" Ã¢ÂÂ ${r['revenue']:,.0f}" if r.get('revenue') else ""
        flag = " Ã°ÂÂÂ´" if r.get('status', '').lower() in ['submit', 'submitted'] else ""
        lines.append(f"  [{days}] {r['name']}{rev}{flag}")

    if len(all_renewals) > 25:
        lines.append(f"\n_...and {len(all_renewals) - 25} more_")

    msg = "\n".join(lines)
    if len(msg) > 4000:
        msg = msg[:4000] + "\n\n_...truncated_"
    await safe_reply_text(update.message, msg, parse_mode="Markdown")


# Ã¢ÂÂÃ¢ÂÂ Marketing Summary Command Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

async def marketing_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /marketing - Get marketing summary for a client/opportunity."""
    if not HAS_MARKETING:
        await update.message.reply_text("Marketing summary module not available.")
        return

    if not context.args:
        await update.message.reply_text(
            "Ã°ÂÂÂ *Marketing Summary*\n\n"
            "Usage: `/marketing Client Name`\n\n"
            "Shows carrier status for all policy types:\n"
            "Incumbent, Market, Submit, Blocked, Declined, Quoted, Proposed, Bound\n\n"
            "Examples:\n"
            "Ã¢ÂÂ¢ `/marketing Triton Hospitality`\n"
            "Ã¢ÂÂ¢ `/marketing Ocean Partners`\n"
            "Ã¢ÂÂ¢ `/marketing Premier Resorts`",
            parse_mode="Markdown",
        )
        return

    client_name = " ".join(context.args)
    await safe_reply_text(update.message, f"Ã¢ÂÂ³ Generating marketing summary for *{client_name}*...", parse_mode="Markdown")

    try:
        summary = await get_marketing_summary(client_name)
        # Split long messages
        if len(summary) > 4000:
            parts = []
            while summary:
                if len(summary) <= 4000:
                    parts.append(summary)
                    break
                split_at = summary.rfind("\n", 0, 4000)
                if split_at < 0:
                    split_at = 4000
                parts.append(summary[:split_at])
                summary = summary[split_at:]
            for part in parts:
                await safe_reply_text(update.message, part, parse_mode="Markdown")
        else:
            await safe_reply_text(update.message, summary, parse_mode="Markdown")
    except Exception as e:
        logger.error(f"Error generating marketing summary: {e}")
        await update.message.reply_text(
            f"Ã¢ÂÂ Ã¯Â¸Â Error generating marketing summary: {str(e)}",
        )


# Ã¢ÂÂÃ¢ÂÂ Daily Briefing Command (manual trigger) Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

async def marketingtx_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /marketingtx - Get marketing summary with taxes & fees (Premium Tx)."""
    if not HAS_MARKETING:
        await update.message.reply_text("Marketing summary module not available.")
        return

    if not context.args:
        await update.message.reply_text(
            "📊 *Marketing Summary (w/ Taxes & Fees)*\n\n"
            "Usage: `/marketingtx Client Name`\n\n"
            "Same as /marketing but shows *Premium Tx* (total premium\n"
            "including taxes, fees & surcharges) instead of Base Premium.\n\n"
            "Examples:\n"
            "• `/marketingtx Triton Hospitality`\n"
            "• `/marketingtx Ocean Partners`\n"
            "• `/marketingtx Premier Resorts`",
            parse_mode="Markdown",
        )
        return

    client_name = " ".join(context.args)
    await safe_reply_text(update.message, f"⏳ Generating marketing summary (w/ taxes) for *{client_name}*...", parse_mode="Markdown")

    try:
        summary = await get_marketing_summary(client_name, with_taxes=True)
        # Split long messages
        if len(summary) > 4000:
            parts = []
            while summary:
                if len(summary) <= 4000:
                    parts.append(summary)
                    break
                split_at = summary.rfind("\n", 0, 4000)
                if split_at < 0:
                    split_at = 4000
                parts.append(summary[:split_at])
                summary = summary[split_at:]
            for part in parts:
                await safe_reply_text(update.message, part, parse_mode="Markdown")
        else:
            await safe_reply_text(update.message, summary, parse_mode="Markdown")
    except Exception as e:
        logger.error(f"Error generating marketing summary (wtx): {e}")
        await update.message.reply_text(
            f"❌ Error generating marketing summary: {str(e)}",
        )


async def briefing_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /briefing - Manually trigger morning briefing."""
    if not HAS_BRIEFING or not HAS_SHEETS:
        await update.message.reply_text("Briefing modules not available.")
        return

    await update.message.reply_text("Ã¢ÂÂ³ Generating briefing...")

    tasks = get_active_tasks()
    new_business = get_new_business() if HAS_SHEETS else []

    success, body = run_morning_briefing(tasks, new_business)

    if success:
        await update.message.reply_text("Ã¢ÂÂ Morning briefing sent!")
    await safe_reply(update, body)


async def debrief_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /debrief - Manually trigger afternoon debrief."""
    if not HAS_BRIEFING or not HAS_SHEETS:
        await update.message.reply_text("Briefing modules not available.")
        return

    await update.message.reply_text("Ã¢ÂÂ³ Generating debrief...")

    tasks = get_active_tasks()
    completed = get_completed_tasks_today()

    success, body = run_afternoon_debrief(tasks, completed)

    if success:
        await update.message.reply_text("Ã¢ÂÂ Afternoon debrief sent!")
    await safe_reply(update, body)


# Ã¢ÂÂÃ¢ÂÂ Scheduled Jobs Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

def scheduled_morning_briefing():
    """Scheduled job: Run morning briefing at 7 AM EST via Telegram."""
    logger.info("Running scheduled morning briefing...")
    try:
        tasks = get_active_tasks() if HAS_SHEETS else []
        new_business = get_new_business() if HAS_SHEETS else []
        success, body = run_morning_briefing(tasks, new_business)
        logger.info(f"Morning briefing result: {success}")
    except Exception as e:
        logger.error(f"Scheduled morning briefing error: {e}")


def scheduled_afternoon_debrief():
    """Scheduled job: Run afternoon debrief at 4 PM EST via Telegram."""
    logger.info("Running scheduled afternoon debrief...")
    try:
        tasks = get_active_tasks() if HAS_SHEETS else []
        completed = get_completed_tasks_today() if HAS_SHEETS else []
        success, body = run_afternoon_debrief(tasks, completed)
        logger.info(f"Afternoon debrief result: {success}")
    except Exception as e:
        logger.error(f"Scheduled afternoon debrief error: {e}")
# Ã¢ÂÂÃ¢ÂÂ Fallback Message Handler Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle non-command messages, including @command style."""
    text = update.message.text.strip()
    lower = text.lower()

    if lower.startswith("@consulting"):
        remainder = text[len("@consulting"):].strip()
        context.args = remainder.split() if remainder else []
        await consulting_command(update, context)
        return
    elif lower.startswith("@report"):
        remainder = text[len("@report"):].strip()
        context.args = remainder.split() if remainder else []
        await report_command(update, context)
        return
    elif lower.startswith("@sales"):
        remainder = text[len("@sales"):].strip()
        context.args = remainder.split() if remainder else []
        await sales_command(update, context)
        return
    elif lower.startswith("@update"):
        await update_command(update, context)
        return
    elif lower.startswith("@status"):
        await status_command(update, context)
        return
    elif lower.startswith("@marketing"):
        remainder = text[len("@marketing"):].strip()
        context.args = remainder.split() if remainder else []
        await marketing_command(update, context)
        return
    elif lower.startswith("@task"):
        remainder = text[len("@task"):].strip()
        context.args = remainder.split() if remainder else []
        await task_command(update, context)
        return
    elif lower.startswith("@done"):
        remainder = text[len("@done"):].strip()
        context.args = remainder.split() if remainder else []
        await done_command(update, context)
        return
    elif lower.startswith("@mytasks"):
        await mytasks_command(update, context)
        return
    elif lower.startswith("@newbiz"):
        remainder = text[len("@newbiz"):].strip()
        context.args = remainder.split() if remainder else []
        await newbiz_command(update, context)
        return
    elif lower.startswith("@lead"):
        remainder = text[len("@lead"):].strip()
        context.args = remainder.split() if remainder else []
        await lead_command(update, context)
        return
    elif lower.startswith("@renewals"):
        await renewals_command(update, context)
        return
    elif lower.startswith("@pipeline"):
        await pipeline_command(update, context)
        return
    elif lower.startswith("@briefing"):
        await briefing_command(update, context)
        return
    elif lower.startswith("@debrief"):
        await debrief_command(update, context)
        return
    elif lower.startswith("@help") or lower.startswith("@start"):
        await start_command(update, context)
        return
    elif any(word in lower for word in ["help", "commands", "what can you do"]):
        await start_command(update, context)
    else:
        await update.message.reply_text(
            "I didn't understand that. Use /help to see available commands.",
        )


# Ã¢ÂÂÃ¢ÂÂ Error Handler Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

async def error_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Log errors."""
    logger.error(f"Update {update} caused error: {context.error}")
    if update and update.message:
        await update.message.reply_text(
            "Ã¢ÂÂ Ã¯Â¸Â An error occurred. Please try again."
        )


# Ã¢ÂÂÃ¢ÂÂ Main Ã¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂÃ¢ÂÂ

def main():
    """Start the bot."""
    if not TELEGRAM_TOKEN:
        logger.error("TELEGRAM_TOKEN environment variable not set!")
        return
    if not AIRTABLE_PAT:
        logger.error("AIRTABLE_PAT environment variable not set!")
        return

    logger.info("Starting Hotel Risk Advisor Bot...")

    # Initialize Google Sheets tabs
    if HAS_SHEETS:
        try:
            initialize_sheets()
            logger.info("Google Sheets initialized")
        except Exception as e:
            logger.warning(f"Google Sheets init failed (will retry on use): {e}")

    app = Application.builder().token(TELEGRAM_TOKEN).build()

    # Original commands
    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("update", update_command))
    app.add_handler(CommandHandler("status", status_command))
    app.add_handler(CommandHandler("add", add_command))
    app.add_handler(CommandHandler("consulting", consulting_command))
    app.add_handler(CommandHandler("report", report_command))
    app.add_handler(CommandHandler("sales", sales_command))

    # New task management commands
    app.add_handler(CommandHandler("task", task_command))
    app.add_handler(CommandHandler("done", done_command))
    app.add_handler(CommandHandler("mytasks", mytasks_command))

    # New business development commands
    app.add_handler(CommandHandler("newbiz", newbiz_command))
    app.add_handler(CommandHandler("lead", lead_command))
    app.add_handler(CommandHandler("pipeline", pipeline_command))
    app.add_handler(CommandHandler("renewals", renewals_command))

    # Marketing summary
    app.add_handler(CommandHandler("marketing", marketing_command))
    app.add_handler(CommandHandler("marketingtx", marketingtx_command))

    # Manual briefing/debrief triggers
    app.add_handler(CommandHandler("briefing", briefing_command))
    app.add_handler(CommandHandler("debrief", debrief_command))

    # Proposal generator (ConversationHandler - must be before generic MessageHandler)
    if HAS_PROPOSAL:
        app.add_handler(get_proposal_conversation_handler())
        # Standalone fallback handlers for /extract and /generate (catch when ConversationHandler misses)
        app.add_handler(CommandHandler("extract", extract_standalone))
        app.add_handler(CommandHandler("generate", generate_standalone))
        logger.info("Proposal generator module loaded")
    else:
        logger.warning("Proposal generator module not available")

    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    app.add_error_handler(error_handler)

    # Set up scheduled daily briefings via post_init (runs inside the async event loop)
    if HAS_SCHEDULER and HAS_BRIEFING:
        async def post_init(application):
            """Start the scheduler after the event loop is running."""
            try:
                scheduler = AsyncIOScheduler(timezone="US/Eastern")

                # Morning briefing at 7:00 AM EST
                scheduler.add_job(
                    scheduled_morning_briefing,
                    CronTrigger(hour=7, minute=0, timezone="US/Eastern"),
                    id="morning_briefing",
                    name="Morning Briefing",
                    replace_existing=True,
                )

                # Afternoon debrief at 4:00 PM EST
                scheduler.add_job(
                    scheduled_afternoon_debrief,
                    CronTrigger(hour=16, minute=0, timezone="US/Eastern"),
                    id="afternoon_debrief",
                    name="Afternoon Debrief",
                    replace_existing=True,
                )

                scheduler.start()
                logger.info("Scheduler started: Morning briefing at 7AM EST, Debrief at 4PM EST")
            except Exception as e:
                logger.error(f"Scheduler setup failed: {e}")

        app.post_init = post_init
    else:
        if not HAS_SCHEDULER:
            logger.warning("APScheduler not installed - daily briefings disabled")
        if not HAS_BRIEFING:
            logger.warning("Briefing module not available - daily briefings disabled")

    logger.info("Starting polling...")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()

