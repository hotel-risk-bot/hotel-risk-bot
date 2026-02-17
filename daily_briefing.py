#!/usr/bin/env python3
"""
Daily Briefing and Debrief Email Generator.
Pulls data from Google Sheets (tasks) and Airtable (renewals/opportunities)
to generate morning briefing and afternoon debrief emails.
"""

import os
import json
import logging
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime, date, timedelta

import requests as http_requests

logger = logging.getLogger(__name__)

# Airtable Configuration
AIRTABLE_PAT = os.environ.get("AIRTABLE_PAT", "")
SALES_BASE_ID = "appnFKEzmdLbR4CHY"
OPPORTUNITIES_TABLE_ID = "tblMKuUsG1cosdQPN"
POLICIES_TABLE_ID = "tbl8vZP2oHrinwVfd"

# Email Configuration
GMAIL_ADDRESS = os.environ.get("GMAIL_ADDRESS", "sburkey@riskportfolio.com")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "")
RECIPIENT_EMAIL = os.environ.get("RECIPIENT_EMAIL", "sburkey@riskportfolio.com")

# Telegram Configuration (for sending alerts via bot too)
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")


def airtable_headers():
    return {
        "Authorization": f"Bearer {AIRTABLE_PAT}",
        "Content-Type": "application/json",
    }


def fetch_upcoming_renewals(days_ahead: int = 120) -> list:
    """Fetch opportunities expiring within the next N days from Airtable."""
    today = date.today()
    cutoff = today + timedelta(days=days_ahead)

    today_str = today.strftime("%Y-%m-%d")
    cutoff_str = cutoff.strftime("%Y-%m-%d")

    # Filter: Effective Date is within the next 120 days
    formula = (
        f"AND("
        f"IS_AFTER({{Effective Date}}, '{today_str}'),"
        f"IS_BEFORE({{Effective Date}}, '{cutoff_str}')"
        f")"
    )

    url = f"https://api.airtable.com/v0/{SALES_BASE_ID}/{OPPORTUNITIES_TABLE_ID}"
    params = {
        "filterByFormula": formula,
        "sort[0][field]": "Effective Date",
        "sort[0][direction]": "asc",
        "pageSize": 100,
    }

    all_records = []
    offset = None

    while True:
        if offset:
            params["offset"] = offset
        try:
            resp = http_requests.get(url, headers=airtable_headers(), params=params, timeout=30)
            resp.raise_for_status()
            data = resp.json()
            all_records.extend(data.get("records", []))
            offset = data.get("offset")
            if not offset:
                break
        except Exception as e:
            logger.error(f"Error fetching renewals: {e}")
            break

    return all_records


def fetch_policies_for_opportunity(opportunity_name: str) -> list:
    """Fetch policies linked to a specific opportunity."""
    formula = f"SEARCH(LOWER('{opportunity_name}'), LOWER({{Name}}))"

    url = f"https://api.airtable.com/v0/{SALES_BASE_ID}/{POLICIES_TABLE_ID}"
    params = {
        "filterByFormula": formula,
        "pageSize": 100,
    }

    try:
        resp = http_requests.get(url, headers=airtable_headers(), params=params, timeout=30)
        resp.raise_for_status()
        return resp.json().get("records", [])
    except Exception as e:
        logger.error(f"Error fetching policies: {e}")
        return []


def classify_renewals(records: list) -> dict:
    """Classify renewal opportunities into categories."""
    result = {
        "submit_alerts": [],       # Status = Submit (exposed!)
        "high_revenue": [],        # Expiring revenue > $5,000
        "all_renewals": [],        # All upcoming renewals
    }

    for rec in records:
        flds = rec.get("fields", {})
        opp_name = flds.get("Opportunity Name", flds.get("Name", "Unknown"))
        company = flds.get("Corporate Name", flds.get("DBA", ""))
        if isinstance(company, list):
            company = company[0] if company else ""
        eff_date = flds.get("Effective Date", "")
        status = flds.get("Market Status", flds.get("Status", ""))
        revenue = 0

        # Try to get revenue from multiple possible fields
        for rev_field in ["Expiring Revenue", "Revenue", "Total Premium", "Expiring Premium", "Premium"]:
            rev_val = flds.get(rev_field)
            if rev_val:
                try:
                    revenue = float(str(rev_val).replace("$", "").replace(",", ""))
                    break
                except (ValueError, TypeError):
                    continue

        # Calculate days until renewal
        days_until = None
        if eff_date:
            try:
                eff = datetime.strptime(eff_date[:10], "%Y-%m-%d").date()
                days_until = (eff - date.today()).days
            except (ValueError, TypeError):
                pass

        # Get Account Manager
        am = flds.get("AM", flds.get("Account Manager", ""))
        if isinstance(am, list):
            am = am[0] if am else ""

        renewal_info = {
            "name": opp_name,
            "company": company,
            "effective_date": eff_date,
            "days_until": days_until,
            "status": status,
            "revenue": revenue,
            "record_id": rec.get("id", ""),
            "am": am,
        }

        result["all_renewals"].append(renewal_info)

        # Check for Submit status (exposed)
        if isinstance(status, str) and status.lower() in ["submit", "submitted"]:
            result["submit_alerts"].append(renewal_info)

        # Check for high revenue
        if revenue > 5000:
            result["high_revenue"].append(renewal_info)

    return result


def generate_morning_briefing(tasks: list, renewals_data: dict, new_business: list) -> str:
    """Generate the morning briefing email content."""
    today_str = date.today().strftime("%A, %B %d, %Y")

    lines = []
    lines.append(f"DAILY MORNING BRIEFING - {today_str}")
    lines.append("=" * 60)
    lines.append("")

    # URGENT ALERTS
    submit_alerts = renewals_data.get("submit_alerts", [])
    high_revenue = renewals_data.get("high_revenue", [])

    if submit_alerts or high_revenue:
        lines.append("*** URGENT ALERTS ***")
        lines.append("-" * 40)

        if submit_alerts:
            lines.append("")
            lines.append(f"EXPOSED - NEEDS SUBMISSION ({len(submit_alerts)} opportunities):")
            for r in submit_alerts:
                days = f"{r['days_until']} days" if r['days_until'] else "TBD"
                rev = f"${r['revenue']:,.0f}" if r['revenue'] else "N/A"
                lines.append(f"  ! {r['name']}")
                lines.append(f"    Company: {r['company']} | Effective: {r['effective_date']} | Days: {days}")
                lines.append(f"    Expiring Revenue: {rev}")
                lines.append(f"    ACTION: Send submission to underwriters ASAP")
                lines.append("")

        if high_revenue:
            lines.append(f"HIGH REVENUE RENEWALS >$5K ({len(high_revenue)} opportunities):")
            for r in high_revenue:
                days = f"{r['days_until']} days" if r['days_until'] else "TBD"
                am_str = f" | AM: {r['am']}" if r.get('am') else ""
                lines.append(f"  {r['name']} — ${r['revenue']:,.0f}")
                lines.append(f"    Status: {r['status']} | Days Until: {days}{am_str}")
                lines.append("")

        lines.append("")

    # ACTIVE TASKS
    lines.append("ACTIVE TASKS")
    lines.append("-" * 40)
    if tasks:
        # Group by priority
        urgent = [t for t in tasks if t.get("priority", "").lower() in ["urgent", "today", "asap"]]
        this_week = [t for t in tasks if t.get("priority", "").lower() in ["this week", "high"]]
        other = [t for t in tasks if t not in urgent and t not in this_week]

        if urgent:
            lines.append("")
            lines.append("  URGENT/TODAY:")
            for i, t in enumerate(urgent, 1):
                lines.append(f"    {i}. [{t['client']}] {t['task']}")
                if t.get("due_date"):
                    lines.append(f"       Due: {t['due_date']}")

        if this_week:
            lines.append("")
            lines.append("  THIS WEEK:")
            for i, t in enumerate(this_week, 1):
                lines.append(f"    {i}. [{t['client']}] {t['task']}")
                if t.get("due_date"):
                    lines.append(f"       Due: {t['due_date']}")

        if other:
            lines.append("")
            lines.append("  OTHER:")
            for i, t in enumerate(other, 1):
                lines.append(f"    {i}. [{t['client']}] {t['task']}")
    else:
        lines.append("  No active tasks. Inbox zero!")
    lines.append("")

    # UPCOMING RENEWALS
    all_renewals = renewals_data.get("all_renewals", [])
    lines.append(f"UPCOMING RENEWALS (Next 120 Days) - {len(all_renewals)} total")
    lines.append("-" * 40)
    if all_renewals:
        # Sort by days until
        sorted_renewals = sorted(all_renewals, key=lambda x: x.get("days_until") or 999)
        for r in sorted_renewals[:20]:
            days = f"{r['days_until']}d" if r['days_until'] else "TBD"
            rev = f" ${r['revenue']:,.0f}" if r['revenue'] else ""
            status_str = f" [{r['status']}]" if r.get('status') else ""
            lines.append(f"  [{days:>4s}] {r['name']}{rev}{status_str}")
        if len(all_renewals) > 20:
            lines.append(f"  ... and {len(all_renewals) - 20} more")
    else:
        lines.append("  No upcoming renewals in the next 120 days.")
    lines.append("")

    # NEW BUSINESS PIPELINE
    if new_business:
        lines.append(f"NEW BUSINESS PIPELINE - {len(new_business)} opportunities")
        lines.append("-" * 40)
        for nb in new_business[:10]:
            rev = nb.get("est_revenue", "")
            lines.append(f"  {nb['client']} - {nb['description']} {rev}")
        lines.append("")

    lines.append("")
    lines.append("---")
    lines.append("Hotel Franchise Practice | HUB International")
    lines.append("Generated by RiskAdvisor Bot")

    return "\n".join(lines)


def generate_afternoon_debrief(tasks: list, completed_today: list,
                                renewals_data: dict) -> str:
    """Generate the afternoon debrief email content."""
    today_str = date.today().strftime("%A, %B %d, %Y")

    lines = []
    lines.append(f"DAILY DEBRIEF - {today_str}")
    lines.append("=" * 60)
    lines.append("")

    # COMPLETED TODAY
    lines.append("COMPLETED TODAY")
    lines.append("-" * 40)
    if completed_today:
        for t in completed_today:
            lines.append(f"  [DONE] [{t['client']}] {t['task']}")
    else:
        lines.append("  No tasks marked complete today.")
    lines.append("")

    # STILL OPEN
    lines.append(f"STILL OPEN - {len(tasks)} tasks remaining")
    lines.append("-" * 40)
    if tasks:
        overdue = []
        upcoming = []
        for t in tasks:
            due = t.get("due_date", "")
            if due:
                try:
                    due_date = datetime.strptime(due[:10], "%Y-%m-%d").date()
                    if due_date < date.today():
                        overdue.append(t)
                        continue
                except (ValueError, TypeError):
                    pass
            upcoming.append(t)

        if overdue:
            lines.append("")
            lines.append(f"  OVERDUE ({len(overdue)}):")
            for t in overdue:
                lines.append(f"    !! [{t['client']}] {t['task']} (Due: {t['due_date']})")

        if upcoming:
            lines.append("")
            lines.append(f"  UPCOMING ({len(upcoming)}):")
            for t in upcoming[:15]:
                due = t.get("due_date", "N/A")
                lines.append(f"    [{t['client']}] {t['task']} (Due: {due})")
    else:
        lines.append("  All tasks complete!")
    lines.append("")

    # RENEWAL ALERTS (same as morning for awareness)
    submit_alerts = renewals_data.get("submit_alerts", [])
    high_revenue = renewals_data.get("high_revenue", [])

    if submit_alerts:
        lines.append(f"STILL EXPOSED - NEEDS SUBMISSION ({len(submit_alerts)})")
        lines.append("-" * 40)
        for r in submit_alerts:
            days = f"{r['days_until']}d" if r['days_until'] else "TBD"
            rev_str = f" - ${r['revenue']:,.0f}" if r['revenue'] else ""
            lines.append(f"  ! [{days}] {r['name']}{rev_str}")
        lines.append("")

    if high_revenue:
        lines.append(f"HIGH REVENUE RENEWALS REMINDER ({len(high_revenue)})")
        lines.append("-" * 40)
        for r in high_revenue[:10]:
            lines.append(f"  {r['name']} — ${r['revenue']:,.0f} ({r['status']})")
        lines.append("")

    lines.append("")
    lines.append("---")
    lines.append("Hotel Franchise Practice | HUB International")
    lines.append("Generated by RiskAdvisor Bot")

    return "\n".join(lines)


def send_email(subject: str, body: str, to_email: str = None) -> bool:
    """Send an email via Gmail SMTP."""
    if not GMAIL_APP_PASSWORD:
        logger.error("GMAIL_APP_PASSWORD not set - cannot send email")
        return False

    to_email = to_email or RECIPIENT_EMAIL

    msg = MIMEMultipart()
    msg["From"] = GMAIL_ADDRESS
    msg["To"] = to_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.send_message(msg)
        server.quit()
        logger.info(f"Email sent: {subject}")
        return True
    except Exception as e:
        logger.error(f"Error sending email: {e}")
        return False


def escape_telegram_markdown(text: str) -> str:
    """Escape special characters that Telegram MarkdownV2 interprets as LaTeX.
    Specifically, dollar signs ($) must be escaped to prevent LaTeX rendering."""
    # For regular Markdown mode, we just need to be careful with $ signs
    # Telegram treats $...$ as LaTeX inline math
    # Replace $ with the unicode full-width dollar sign or escape it
    return text.replace("$", "\\$")


async def send_telegram_message(text: str, chat_id: str = None, escape_dollars: bool = True):
    """Send a message via Telegram bot."""
    if not TELEGRAM_TOKEN or not (chat_id or TELEGRAM_CHAT_ID):
        return

    target_chat = chat_id or TELEGRAM_CHAT_ID

    # Escape dollar signs to prevent Telegram LaTeX rendering
    if escape_dollars:
        text = escape_telegram_markdown(text)

    url = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage"
    try:
        http_requests.post(url, json={
            "chat_id": target_chat,
            "text": text,
            "parse_mode": "Markdown",
        }, timeout=15)
    except Exception as e:
        logger.error(f"Error sending Telegram message: {e}")


def run_morning_briefing(tasks: list, new_business: list = None):
    """Execute the full morning briefing workflow."""
    logger.info("Running morning briefing...")

    # Fetch renewals from Airtable
    renewal_records = fetch_upcoming_renewals(120)
    renewals_data = classify_renewals(renewal_records)

    # Generate email
    body = generate_morning_briefing(tasks, renewals_data, new_business or [])

    # Send email
    today_str = date.today().strftime("%m/%d/%Y")
    subject = f"Morning Briefing - {today_str} | {len(tasks)} Tasks | {len(renewals_data['all_renewals'])} Renewals"

    if renewals_data["submit_alerts"]:
        subject = f"[ALERT] {subject} | {len(renewals_data['submit_alerts'])} EXPOSED"

    success = send_email(subject, body)
    logger.info(f"Morning briefing sent: {success}")
    return success, body


def run_afternoon_debrief(tasks: list, completed_today: list = None):
    """Execute the full afternoon debrief workflow."""
    logger.info("Running afternoon debrief...")

    # Fetch renewals from Airtable
    renewal_records = fetch_upcoming_renewals(120)
    renewals_data = classify_renewals(renewal_records)

    # Generate email
    body = generate_afternoon_debrief(tasks, completed_today or [], renewals_data)

    # Send email
    today_str = date.today().strftime("%m/%d/%Y")
    completed_count = len(completed_today) if completed_today else 0
    subject = f"Daily Debrief - {today_str} | {completed_count} Completed | {len(tasks)} Remaining"

    success = send_email(subject, body)
    logger.info(f"Afternoon debrief sent: {success}")
    return success, body
