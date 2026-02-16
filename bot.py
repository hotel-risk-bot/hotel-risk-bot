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
from datetime import datetime

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

# ── Configuration (from environment variables) ────────────────────────────
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")
AIRTABLE_PAT = os.environ.get("AIRTABLE_PAT", "")

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


# ── Airtable REST API Functions ───────────────────────────────────────────

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


# ── Consulting Query Functions ────────────────────────────────────────────

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


# ── Claims Development Parser ──────────────────────────────────────────────

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

    lines = ["📈 *Claims Development*"]

    for v in valuations:
        detail_parts = []
        if v["paid"] > 0:
            detail_parts.append(f"Paid: ${v['paid']:,.0f}")
        if v["reserved"] > 0:
            detail_parts.append(f"Rsv: ${v['reserved']:,.0f}")
        if v["expenses"] > 0:
            detail_parts.append(f"Exp: ${v['expenses']:,.0f}")

        detail_str = f" ({', '.join(detail_parts)})" if detail_parts else ""
        lines.append(f"• {v['date']}: *${v['total_incurred']:,.0f}*{detail_str}")

    return "\n".join(lines)


# ── Report Formatting ──────────────────────────────────────────────────────

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

    status_emoji = "✅" if status == "Open" else "🔴" if status == "Closed" else "⚪"

    # Build the report — matching exact user-requested format
    lines = []
    lines.append(f"{'─' * 35}")

    # ── Date of Loss (first) ──
    lines.append(f"📅 *Date of Loss:* {incident_date}")
    lines.append("")

    # ── Claim Details (with location info grouped here) ──
    lines.append(f"📋 *Claim Details*")
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

    # ── Incident Details ──
    lines.append(f"📋 *Incident Details*")
    lines.append(f"Claimant: {involved_party}")
    lines.append(f"Cause of Loss: {cause_of_loss}")
    if risk_hazard != "N/A":
        lines.append(f"⚠️ Hazard: {risk_hazard}")
    if location_of_incident != "N/A":
        lines.append(f"🏢 Location of Incident: {location_of_incident}")
    if brief_desc != "N/A":
        lines.append(f"Description: {brief_desc}")
    lines.append("")

    # ── Financial Summary ──
    lines.append(f"💰 *Financial Summary*")
    lines.append(f"• Total Incurred: ${incurred:,.0f}")
    if paid != "N/A":
        try:
            lines.append(f"• Paid: ${float(paid):,.0f}")
        except (ValueError, TypeError):
            lines.append(f"• Paid: {paid}")
    try:
        lines.append(f"• Reserved: ${float(reserved):,.0f}")
    except (ValueError, TypeError):
        lines.append(f"• Reserved: {reserved}")
    if expenses != "N/A":
        try:
            exp_vals = f.get("Expenses Helper", [])
            if isinstance(exp_vals, list) and exp_vals:
                lines.append(f"• Expenses: ${float(exp_vals[-1]):,.0f}")
            else:
                lines.append(f"• Expenses: ${float(expenses):,.0f}")
        except (ValueError, TypeError):
            pass
    lines.append("")

    # ── Claims Development ──
    raw_activity = f.get("Activity Rollup Raw Data", "")
    if raw_activity:
        valuations = parse_claims_development(raw_activity)
        if valuations:
            dev_text = format_claims_development(valuations)
            lines.append(dev_text)
            lines.append("")

    # ── Summary of Facts ──
    if summary_of_facts != "N/A" and len(summary_of_facts) > 5:
        sf = summary_of_facts[:500]
        if len(summary_of_facts) > 500:
            sf += "..."
        lines.append(f"📝 *Summary of Facts:*")
        lines.append(sf)
        lines.append("")

    # ── Attorney Representation ──
    if attorney_rep:
        lines.append(f"⚖️ *Attorney Representation:* Yes")
        if attorney_demand != "N/A":
            lines.append(f"Attorney Demand: ${attorney_demand}")
        lines.append("")

    # ── Carrier / Policy ──
    if carrier != "N/A":
        lines.append(f"Carrier: {carrier}")
    if policy_num != "N/A":
        lines.append(f"Policy #: {policy_num}")

    return "\n".join(lines)


# ── Sales System Functions ─────────────────────────────────────────────────

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
    lines.append(f"{'─' * 35}")
    lines.append(f"🏢 *{opp_name}*")
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


# ── Argument Parser ────────────────────────────────────────────────────────

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


# ── PDF Report Generator ──────────────────────────────────────────────────

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


# ── Command Handlers ───────────────────────────────────────────────────────

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /start command."""
    welcome = (
        "🏨 *Hotel Risk Advisor Bot*\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "Welcome! I can help you query the HUB International hotel insurance databases.\n\n"
        "*Available Commands:*\n"
        "• /consulting `query` — Search Claims\n"
        "• /report `query` — Executive PDF Report\n"
        "• /sales `query` — Search Sales System\n"
        "• /update — Get task list\n"
        "• /status — View progress\n"
        "• /add `Client | Task | Priority` — Add task\n"
        "• /help — Show this message\n\n"
        "*Consulting Query Examples:*\n"
        "• `/consulting Jasmin` — All claims\n"
        "• `/consulting Jasmin open` — Open claims\n"
        "• `/consulting Jasmin closed` — Closed claims\n"
        "• `/consulting Jasmin all` — All open + closed\n"
        "• `/consulting Jasmin liability` — Liability only\n"
        "• `/consulting Jasmin property` — Property only\n"
        "• `/consulting Jasmin open liability greater than 25000`\n"
        "• `/consulting Jasmin last 5 years`\n"
        "• `/consulting Jasmin closed property last 3 years`\n\n"
        "*PDF Report:*\n"
        "• `/report Ocean Partners` — Full PDF report\n"
        "• `/report Jasmin open liability last 5 years`\n"
    )
    await update.message.reply_text(welcome, parse_mode="Markdown")


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await start_command(update, context)


async def update_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /update - Get task list from Sales system."""
    await update.message.reply_text("⏳ Fetching task list from Sales System...")

    records = airtable_list_records(
        SALES_BASE_ID, TASKS_TABLE_ID,
        filter_formula='NOT({Task Status} = "Done")',
        max_records=20,
    )

    if not records:
        await update.message.reply_text("No open tasks found.")
        return

    lines = ["📋 *Open Tasks*\n"]
    for rec in records:
        f = rec.get("fields", {})
        name = f.get("Name", "Unnamed")
        task_status = f.get("Task Status", "N/A")
        priority = f.get("Priority", "N/A")
        due = f.get("Due Date", "N/A")
        cam = f.get("CAM", "N/A")

        status_emoji = {"Todo": "🔴", "In progress": "🟡"}.get(task_status, "⚪")
        pri_emoji = {"High": "🔥", "Medium": "⚡", "Low": "💤"}.get(priority, "")

        lines.append(f"{status_emoji} {pri_emoji} *{name}*")
        if due != "N/A":
            lines.append(f"   Due: {due}")
        if cam != "N/A":
            lines.append(f"   Assigned: {cam}")
        lines.append("")

    msg = "\n".join(lines)
    if len(msg) > 4000:
        msg = msg[:4000] + "\n\n_...truncated_"
    await update.message.reply_text(msg, parse_mode="Markdown")


async def status_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /status - View progress summary."""
    await update.message.reply_text("⏳ Fetching status from Sales System...")

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
        "📊 *Task Progress*\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        f"Total Tasks: {total}\n"
        f"✅ Done: {done}\n"
        f"🟡 In Progress: {in_progress}\n"
        f"🔴 Todo: {todo}\n"
    )
    await update.message.reply_text(msg, parse_mode="Markdown")


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
        await update.message.reply_text(
            "Please use the format: `/add Client | Task | Priority`",
            parse_mode="Markdown",
        )
        return

    company = parts[0]
    task_desc = parts[1]
    priority = parts[2] if len(parts) > 2 else "Medium"

    valid_priorities = {"high": "High", "medium": "Medium", "low": "Low"}
    priority = valid_priorities.get(priority.lower(), "Medium")

    await update.message.reply_text(f"⏳ Adding task for {company}...")

    result = airtable_create_record(SALES_BASE_ID, TODO_TABLE_ID, {
        "Notes": f"{task_desc}\n\nAdded via Telegram Bot on {datetime.now().strftime('%m/%d/%Y %I:%M %p')}",
        "Status": "Todo",
        "Priority": priority,
    })

    if result:
        await update.message.reply_text(
            f"✅ Task added successfully!\n"
            f"📌 {task_desc}\n"
            f"Priority: {priority}",
        )
    else:
        await update.message.reply_text("❌ Failed to add task. Please try again.")


# ── Consulting Query Handler ───────────────────────────────────────────────

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
        query_desc += f" | Policy Year ≥ *{params['min_policy_year']}*"

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
            "🔍 *Consulting System Query*\n\n"
            "Usage:\n"
            "• `/consulting ClientName` — All claims\n"
            "• `/consulting ClientName open` — Open claims\n"
            "• `/consulting ClientName closed` — Closed claims\n"
            "• `/consulting ClientName liability` — Liability only\n"
            "• `/consulting ClientName property` — Property only\n"
            "• `/consulting ClientName open liability greater than 25000`\n"
            "• `/consulting ClientName last 5 years`\n"
            "• `/consulting ClientName closed property last 3 years`\n\n"
            "Searches across Client Name, Corporate Name, DBA, and Company fields.",
            parse_mode="Markdown",
        )
        return

    await update.message.reply_text("⏳ Searching Consulting System...", parse_mode="Markdown")

    results, params, query_desc = await run_consulting_query(context.args)

    if not results:
        await update.message.reply_text(
            f"No claims found matching your criteria.\n{query_desc}",
            parse_mode="Markdown",
        )
        return

    total_incurred = sum(r["incurred"] for r in results)
    header = (
        f"🏨 *Consulting System Results*\n"
        f"━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        f"Found *{len(results)}* claim(s)\n"
        f"Total Incurred: *${total_incurred:,.0f}*\n"
        f"{query_desc}\n"
    )
    await update.message.reply_text(header, parse_mode="Markdown")

    for i, rec in enumerate(results[:20]):
        try:
            report = format_claim_report(rec)
            if len(report) > 4000:
                mid = report.rfind("\n", 0, 4000)
                if mid < 0:
                    mid = 4000
                await update.message.reply_text(report[:mid], parse_mode="Markdown")
                await update.message.reply_text(report[mid:], parse_mode="Markdown")
            else:
                await update.message.reply_text(report, parse_mode="Markdown")
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
        await update.message.reply_text(
            f"_Showing 20 of {len(results)} results. "
            f"Refine your search to see more specific results._",
            parse_mode="Markdown",
        )


# ── Report Command Handler ────────────────────────────────────────────────

async def report_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /report command — generate executive PDF report."""
    if not context.args:
        await update.message.reply_text(
            "📄 *Executive PDF Report*\n\n"
            "Usage: `/report ClientName [filters]`\n\n"
            "Examples:\n"
            "• `/report Ocean Partners` — Full report\n"
            "• `/report Jasmin open liability` — Filtered\n"
            "• `/report Jasmin last 5 years`\n"
            "• `/report Jasmin closed greater than 25000`\n",
            parse_mode="Markdown",
        )
        return

    await update.message.reply_text("⏳ Generating executive PDF report...", parse_mode="Markdown")

    results, params, query_desc = await run_consulting_query(context.args)

    if not results:
        await update.message.reply_text(
            f"No claims found matching your criteria.\n{query_desc}",
            parse_mode="Markdown",
        )
        return

    try:
        filepath = generate_enhanced_pdf(params["client_name"], results, params)

        total_incurred = sum(r["incurred"] for r in results)
        caption = (
            f"📄 Executive Claims Report — {params['client_name']}\n"
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
        await update.message.reply_text(
            f"⚠️ Error generating PDF report: {str(e)}\n\n"
            f"The query found *{len(results)}* claims. Try `/consulting` to view them in chat.",
            parse_mode="Markdown",
        )


# ── Sales Query Handler ───────────────────────────────────────────────────

async def sales_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handle /sales command."""
    if not context.args:
        await update.message.reply_text(
            "🔍 *Sales System Query*\n\n"
            "Usage: `/sales search term`\n\n"
            "Examples:\n"
            "• `/sales Marriott`\n"
            "• `/sales Best Western`\n"
            "• `/sales Premier Resorts`",
            parse_mode="Markdown",
        )
        return

    query = " ".join(context.args)
    await update.message.reply_text(
        f"⏳ Searching Sales System for: *{query}*...",
        parse_mode="Markdown",
    )

    records = search_sales(query)

    if not records:
        await update.message.reply_text(f"No results found for: {query}")
        return

    header = (
        f"🏢 *Sales System Results*\n"
        f"━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        f"Found *{len(records)}* result(s) for: *{query}*\n"
    )
    await update.message.reply_text(header, parse_mode="Markdown")

    for i, rec in enumerate(records[:10]):
        try:
            report = format_sales_record(rec)
            await update.message.reply_text(report, parse_mode="Markdown")
        except Exception as e:
            logger.error(f"Error formatting sales record: {e}")
            try:
                report = format_sales_record(rec)
                await update.message.reply_text(report)
            except Exception:
                await update.message.reply_text(f"Error displaying result #{i+1}")

    if len(records) > 10:
        await update.message.reply_text(
            f"_Showing 10 of {len(records)} results._",
            parse_mode="Markdown",
        )


# ── Fallback Message Handler ──────────────────────────────────────────────

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
    elif lower.startswith("@help") or lower.startswith("@start"):
        await start_command(update, context)
        return
    elif any(word in lower for word in ["help", "commands", "what can you do"]):
        await start_command(update, context)
    else:
        await update.message.reply_text(
            "I didn't understand that. Use /help to see available commands.",
        )


# ── Error Handler ─────────────────────────────────────────────────────────

async def error_handler(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Log errors."""
    logger.error(f"Update {update} caused error: {context.error}")
    if update and update.message:
        await update.message.reply_text(
            "⚠️ An error occurred. Please try again."
        )


# ── Main ───────────────────────────────────────────────────────────────────

def main():
    """Start the bot."""
    if not TELEGRAM_TOKEN:
        logger.error("TELEGRAM_TOKEN environment variable not set!")
        return
    if not AIRTABLE_PAT:
        logger.error("AIRTABLE_PAT environment variable not set!")
        return

    logger.info("Starting Hotel Risk Advisor Bot...")

    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("update", update_command))
    app.add_handler(CommandHandler("status", status_command))
    app.add_handler(CommandHandler("add", add_command))
    app.add_handler(CommandHandler("consulting", consulting_command))
    app.add_handler(CommandHandler("report", report_command))
    app.add_handler(CommandHandler("sales", sales_command))

    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    app.add_error_handler(error_handler)

    logger.info("Starting polling...")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()
