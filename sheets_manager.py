#!/usr/bin/env python3
"""
Google Sheets Manager for RiskAdvisor Tasks spreadsheet.
Handles reading/writing to Active Tasks, Completed Tasks, Renewals,
New Business, and Leads tabs.

Uses Google Sheets API v4 with a service account.
"""

import os
import json
import logging
from datetime import datetime, date

import requests as http_requests

logger = logging.getLogger(__name__)

# Google Sheets Configuration
SPREADSHEET_ID = "1gwM3AefSzXI_7ECiQW85pGmo772fsxzGCKo-6QyID38"
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")

# Tab names
ACTIVE_TASKS_TAB = "Active Tasks"
COMPLETED_TASKS_TAB = "Completed Tasks"
RENEWALS_TAB = "Renewals"
NEW_BUSINESS_TAB = "New Business"
LEADS_TAB = "Leads"

# Column headers for each tab
ACTIVE_TASKS_HEADERS = ["Priority", "Due Date", "Client", "Task", "Status", "Category", "Notes"]
COMPLETED_TASKS_HEADERS = ["Priority", "Due Date", "Client", "Task", "Status", "Category", "Notes"]
NEW_BUSINESS_HEADERS = ["Date Added", "Client", "Description", "N/R", "Est Revenue", "Status", "AM", "Notes"]
LEADS_HEADERS = ["Date Added", "Client", "Contact", "Source", "Description", "Est Revenue", "Status", "Notes"]


def _get_access_token():
    """Get Google API access token from service account credentials."""
    import jwt
    import time

    if not GOOGLE_SERVICE_ACCOUNT_JSON:
        logger.error("GOOGLE_SERVICE_ACCOUNT_JSON not set")
        return None

    try:
        creds = json.loads(GOOGLE_SERVICE_ACCOUNT_JSON)
    except json.JSONDecodeError:
        logger.error("Invalid GOOGLE_SERVICE_ACCOUNT_JSON format")
        return None

    now = int(time.time())
    payload = {
        "iss": creds["client_email"],
        "scope": "https://www.googleapis.com/auth/spreadsheets",
        "aud": "https://oauth2.googleapis.com/token",
        "iat": now,
        "exp": now + 3600,
    }

    signed_jwt = jwt.encode(payload, creds["private_key"], algorithm="RS256")

    resp = http_requests.post(
        "https://oauth2.googleapis.com/token",
        data={
            "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
            "assertion": signed_jwt,
        },
        timeout=15,
    )
    resp.raise_for_status()
    return resp.json()["access_token"]


def _sheets_headers():
    """Get authorization headers for Google Sheets API."""
    token = _get_access_token()
    if not token:
        return None
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    }


def _sheets_get(range_str: str) -> list:
    """Read values from a Google Sheet range."""
    headers = _sheets_headers()
    if not headers:
        return []

    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}/values/{range_str}"
    try:
        resp = http_requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
        return resp.json().get("values", [])
    except Exception as e:
        logger.error(f"Sheets API read error: {e}")
        return []


def _sheets_append(range_str: str, values: list) -> bool:
    """Append rows to a Google Sheet range."""
    headers = _sheets_headers()
    if not headers:
        return False

    url = (
        f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}/values/{range_str}:append"
        f"?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS"
    )
    body = {"values": values}
    try:
        resp = http_requests.post(url, headers=headers, json=body, timeout=15)
        resp.raise_for_status()
        return True
    except Exception as e:
        logger.error(f"Sheets API append error: {e}")
        return False


def _sheets_update(range_str: str, values: list) -> bool:
    """Update values in a specific range."""
    headers = _sheets_headers()
    if not headers:
        return False

    url = (
        f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}/values/{range_str}"
        f"?valueInputOption=USER_ENTERED"
    )
    body = {"values": values}
    try:
        resp = http_requests.put(url, headers=headers, json=body, timeout=15)
        resp.raise_for_status()
        return True
    except Exception as e:
        logger.error(f"Sheets API update error: {e}")
        return False


def _sheets_delete_row(sheet_id: int, row_index: int) -> bool:
    """Delete a specific row from a sheet by sheet GID."""
    headers = _sheets_headers()
    if not headers:
        return False

    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}:batchUpdate"
    body = {
        "requests": [{
            "deleteDimension": {
                "range": {
                    "sheetId": sheet_id,
                    "dimension": "ROWS",
                    "startIndex": row_index,
                    "endIndex": row_index + 1,
                }
            }
        }]
    }
    try:
        resp = http_requests.post(url, headers=headers, json=body, timeout=15)
        resp.raise_for_status()
        return True
    except Exception as e:
        logger.error(f"Sheets API delete row error: {e}")
        return False


def _ensure_tab_exists(tab_name: str, headers: list) -> bool:
    """Create a tab if it doesn't exist, with headers."""
    hdrs = _sheets_headers()
    if not hdrs:
        return False

    # Check if tab exists
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}"
    try:
        resp = http_requests.get(url, headers=hdrs, timeout=15)
        resp.raise_for_status()
        sheets = resp.json().get("sheets", [])
        for s in sheets:
            if s["properties"]["title"] == tab_name:
                return True  # Already exists
    except Exception as e:
        logger.error(f"Error checking tabs: {e}")
        return False

    # Create the tab
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}:batchUpdate"
    body = {
        "requests": [{
            "addSheet": {
                "properties": {"title": tab_name}
            }
        }]
    }
    try:
        resp = http_requests.post(url, headers=hdrs, json=body, timeout=15)
        resp.raise_for_status()
    except Exception as e:
        logger.error(f"Error creating tab {tab_name}: {e}")
        return False

    # Add headers
    _sheets_update(f"'{tab_name}'!A1", [headers])
    return True


def _get_sheet_id(tab_name: str) -> int | None:
    """Get the numeric sheet ID (GID) for a tab name."""
    hdrs = _sheets_headers()
    if not hdrs:
        return None

    url = f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}"
    try:
        resp = http_requests.get(url, headers=hdrs, timeout=15)
        resp.raise_for_status()
        sheets = resp.json().get("sheets", [])
        for s in sheets:
            if s["properties"]["title"] == tab_name:
                return s["properties"]["sheetId"]
    except Exception as e:
        logger.error(f"Error getting sheet ID: {e}")
    return None


# ── Active Tasks ─────────────────────────────────────────────────────────

def get_active_tasks() -> list:
    """Get all active tasks from the Active Tasks tab.
    Returns list of dicts with keys: row_num, priority, due_date, client, task, status, category, notes
    """
    rows = _sheets_get(f"'{ACTIVE_TASKS_TAB}'!A:G")
    if not rows or len(rows) < 2:
        return []

    tasks = []
    for i, row in enumerate(rows[1:], start=2):  # Skip header, 1-indexed row numbers
        if not row or not any(cell.strip() for cell in row if cell):
            continue
        tasks.append({
            "row_num": i,
            "priority": row[0] if len(row) > 0 else "",
            "due_date": row[1] if len(row) > 1 else "",
            "client": row[2] if len(row) > 2 else "",
            "task": row[3] if len(row) > 3 else "",
            "status": row[4] if len(row) > 4 else "",
            "category": row[5] if len(row) > 5 else "",
            "notes": row[6] if len(row) > 6 else "",
        })
    return tasks


def add_active_task(client: str, task: str, priority: str = "This Week",
                    due_date: str = None, category: str = "Action",
                    notes: str = "") -> bool:
    """Add a new task to the Active Tasks tab."""
    if not due_date:
        due_date = date.today().strftime("%Y-%m-%d")

    if notes:
        notes = f"{notes} | Added via Telegram {datetime.now().strftime('%m/%d/%Y %I:%M %p')}"
    else:
        notes = f"Added via Telegram {datetime.now().strftime('%m/%d/%Y %I:%M %p')}"

    row = [priority, due_date, client, task, "Pending", category, notes]
    return _sheets_append(f"'{ACTIVE_TASKS_TAB}'!A:G", [row])


def complete_task(task_number: int) -> dict | None:
    """Move a task from Active Tasks to Completed Tasks.
    task_number is the display index (1-based) of the task in the active list.
    Returns the completed task dict or None on failure.
    """
    tasks = get_active_tasks()
    if not tasks or task_number < 1 or task_number > len(tasks):
        return None

    task = tasks[task_number - 1]
    actual_row = task["row_num"]

    # Add to Completed Tasks tab
    completed_row = [
        task["priority"],
        task["due_date"],
        task["client"],
        task["task"],
        "Completed",
        task["category"],
        task["notes"],
    ]
    _sheets_append(f"'{COMPLETED_TASKS_TAB}'!A:G", [completed_row])

    # Delete from Active Tasks
    sheet_id = _get_sheet_id(ACTIVE_TASKS_TAB)
    if sheet_id is not None:
        _sheets_delete_row(sheet_id, actual_row - 1)  # 0-indexed for API

    return task


def get_completed_tasks_today() -> list:
    """Get tasks completed today from the Completed Tasks tab."""
    rows = _sheets_get(f"'{COMPLETED_TASKS_TAB}'!A:G")
    if not rows or len(rows) < 2:
        return []

    today_str = date.today().strftime("%Y-%m-%d")
    today_str2 = date.today().strftime("%m/%d/%Y")
    tasks = []
    for row in rows[1:]:
        if not row or not any(cell.strip() for cell in row if cell):
            continue
        notes = row[6] if len(row) > 6 else ""
        # Check if completed today based on notes timestamp
        if today_str in notes or today_str2 in notes or "today" in notes.lower():
            tasks.append({
                "client": row[2] if len(row) > 2 else "",
                "task": row[3] if len(row) > 3 else "",
                "category": row[5] if len(row) > 5 else "",
            })
    return tasks


# ── New Business ─────────────────────────────────────────────────────────

def add_new_business(client: str, description: str, nr: str = "N",
                     est_revenue: str = "", status: str = "Pipeline",
                     am: str = "Stefan", notes: str = "") -> bool:
    """Add a new business opportunity to the New Business tab."""
    _ensure_tab_exists(NEW_BUSINESS_TAB, NEW_BUSINESS_HEADERS)

    date_added = date.today().strftime("%Y-%m-%d")
    row = [date_added, client, description, nr, est_revenue, status, am, notes]
    return _sheets_append(f"'{NEW_BUSINESS_TAB}'!A:H", [row])


def get_new_business() -> list:
    """Get all new business entries."""
    _ensure_tab_exists(NEW_BUSINESS_TAB, NEW_BUSINESS_HEADERS)
    rows = _sheets_get(f"'{NEW_BUSINESS_TAB}'!A:H")
    if not rows or len(rows) < 2:
        return []

    items = []
    for i, row in enumerate(rows[1:], start=1):
        if not row or not any(cell.strip() for cell in row if cell):
            continue
        items.append({
            "num": i,
            "date_added": row[0] if len(row) > 0 else "",
            "client": row[1] if len(row) > 1 else "",
            "description": row[2] if len(row) > 2 else "",
            "nr": row[3] if len(row) > 3 else "",
            "est_revenue": row[4] if len(row) > 4 else "",
            "status": row[5] if len(row) > 5 else "",
            "am": row[6] if len(row) > 6 else "",
            "notes": row[7] if len(row) > 7 else "",
        })
    return items


# ── Leads ────────────────────────────────────────────────────────────────

def add_lead(client: str, contact: str = "", source: str = "",
             description: str = "", est_revenue: str = "",
             notes: str = "") -> bool:
    """Add a new lead to the Leads tab."""
    _ensure_tab_exists(LEADS_TAB, LEADS_HEADERS)

    date_added = date.today().strftime("%Y-%m-%d")
    row = [date_added, client, contact, source, description, est_revenue, "New", notes]
    return _sheets_append(f"'{LEADS_TAB}'!A:H", [row])


def get_leads() -> list:
    """Get all leads."""
    _ensure_tab_exists(LEADS_TAB, LEADS_HEADERS)
    rows = _sheets_get(f"'{LEADS_TAB}'!A:H")
    if not rows or len(rows) < 2:
        return []

    items = []
    for i, row in enumerate(rows[1:], start=1):
        if not row or not any(cell.strip() for cell in row if cell):
            continue
        items.append({
            "num": i,
            "date_added": row[0] if len(row) > 0 else "",
            "client": row[1] if len(row) > 1 else "",
            "contact": row[2] if len(row) > 2 else "",
            "source": row[3] if len(row) > 3 else "",
            "description": row[4] if len(row) > 4 else "",
            "est_revenue": row[5] if len(row) > 5 else "",
            "status": row[6] if len(row) > 6 else "",
            "notes": row[7] if len(row) > 7 else "",
        })
    return items


# ── Renewals Tab (read from sheet, populated by daily sync) ──────────────

def get_renewals_from_sheet() -> list:
    """Get renewals from the Renewals tab."""
    rows = _sheets_get(f"'{RENEWALS_TAB}'!A:G")
    if not rows or len(rows) < 2:
        return []

    items = []
    for row in rows[1:]:
        if not row or not any(cell.strip() for cell in row if cell):
            continue
        items.append({
            "opportunity_name": row[0] if len(row) > 0 else "",
            "dba": row[1] if len(row) > 1 else "",
            "days_until": row[2] if len(row) > 2 else "",
            "effective_date": row[3] if len(row) > 3 else "",
            "market_status": row[4] if len(row) > 4 else "",
            "type": row[5] if len(row) > 5 else "",
            "am": row[6] if len(row) > 6 else "",
        })
    return items


def update_renewals_tab(renewals_data: list) -> bool:
    """Replace the Renewals tab data with fresh data from Airtable.
    renewals_data is a list of dicts with keys matching the column headers.
    """
    # Clear existing data (keep header)
    headers = _sheets_headers()
    if not headers:
        return False

    # Get current data to know how many rows to clear
    current = _sheets_get(f"'{RENEWALS_TAB}'!A:G")
    if current and len(current) > 1:
        clear_url = (
            f"https://sheets.googleapis.com/v4/spreadsheets/{SPREADSHEET_ID}/values/"
            f"'{RENEWALS_TAB}'!A2:G{len(current) + 10}:clear"
        )
        try:
            http_requests.post(clear_url, headers=headers, json={}, timeout=15)
        except Exception as e:
            logger.error(f"Error clearing renewals: {e}")

    # Write new data
    if renewals_data:
        rows = []
        for r in renewals_data:
            rows.append([
                r.get("opportunity_name", ""),
                r.get("dba", ""),
                str(r.get("days_until", "")),
                r.get("effective_date", ""),
                r.get("market_status", ""),
                r.get("type", ""),
                r.get("am", ""),
            ])
        return _sheets_append(f"'{RENEWALS_TAB}'!A:G", rows)
    return True


def initialize_sheets():
    """Ensure all required tabs exist with proper headers."""
    _ensure_tab_exists(NEW_BUSINESS_TAB, NEW_BUSINESS_HEADERS)
    _ensure_tab_exists(LEADS_TAB, LEADS_HEADERS)
    logger.info("Google Sheets tabs initialized")
