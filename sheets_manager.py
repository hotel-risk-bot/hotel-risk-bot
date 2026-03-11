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
import re
import time
from datetime import datetime, date

import requests as http_requests

logger = logging.getLogger(__name__)

# Google Sheets Configuration
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "1gwM3AefSzXI_7ECiQW85pGmo772fsxzGCKo-6QyID38")
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


def _parse_service_account_json(raw):
    """Parse service account JSON, handling Railway env var mangling."""
    if not raw:
        return None
    raw = raw.strip()

    # Attempt 1: Direct parse
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass

    # Attempt 2: Fix real newlines inside JSON string values
    try:
        fixed = re.sub(
            r'"((?:[^"\\]|\\.)*?)"',
            lambda m: '"' + m.group(1).replace('\n', '\\n').replace('\r', '') + '"',
            raw,
            flags=re.DOTALL
        )
        return json.loads(fixed)
    except Exception:
        pass

    # Attempt 3: Line-by-line reconstruction
    try:
        lines = raw.split('\n')
        rebuilt = []
        in_pk = False
        for line in lines:
            s = line.strip()
            if '"private_key"' in s:
                in_pk = True
                rebuilt.append(s)
            elif in_pk:
                if (s.startswith('"') and not s.startswith('"-----')) or s in ('}', '},'):
                    in_pk = False
                    rebuilt.append(s)
                else:
                    if rebuilt:
                        rebuilt[-1] = rebuilt[-1].rstrip() + '\\n' + s
                    else:
                        rebuilt.append(s)
            else:
                rebuilt.append(s)
        rejoined = ' '.join(rebuilt)
        return json.loads(rejoined)
    except Exception as e:
        logger.error(f"All JSON parse attempts failed: {e}")
        return None


def _get_access_token():
    """Get Google API access token from service account credentials."""
    import jwt

    if not GOOGLE_SERVICE_ACCOUNT_JSON:
        logger.error("GOOGLE_SERVICE_ACCOUNT_JSON not set")
        return None

    creds = _parse_service_account_json(GOOGLE_SERVICE_ACCOUNT_JSON)
    if not creds:
        logger.error("Could not parse GOOGLE_SERVICE_ACCOUNT_JSON")
        return None

    # Ensure private_key has proper newline characters
    pk = creds.get("private_key", "")
    if pk and '\n' not in pk and '\\n' in pk:
        creds["private_key"] = pk.replace('\\n', '\n')

    now = int(time.time())
    payload = {
        "iss": creds["client_email"],
        "scope": "https://www.googleapis.com/auth/spreadsheets",
        "aud": "https://oauth2.googleapis.com/token",
        "iat": now,
        "exp": now + 3600,
    }

    try:
        signed_jwt = jwt.encode(payload, creds["private_key"], algorithm="RS256")
    except Exception as e:
        logger.error(f"JWT signing failed: {e}")
        return None

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
                "priority": row[0] if len(row) > 0 else "",
                "due_date": row[1] if len(row) > 1 else "",
                "client": row[2] if len(row) > 2 else "",
                "task": row[3] if len(row) > 3 else "",
                "status": row[4] if len(row) > 4 else "",
                "category": row[5] if len(row) > 5 else "",
                "notes": row[6] if len(row) > 6 else "",
            })
    return tasks


# ── New Business & Leads ──────────────────────────────────────────────────

def add_new_business(client: str, description: str, nr: str = "New",
                      est_revenue: str = "", status: str = "Active",
                      am: str = "", notes: str = "") -> bool:
    """Add a new business opportunity to the New Business tab."""
    date_added = date.today().strftime("%Y-%m-%d")
    row = [date_added, client, description, nr, est_revenue, status, am, notes]
    return _sheets_append(f"'{NEW_BUSINESS_TAB}'!A:H", [row])


def get_new_business() -> list:
    """Get all new business opportunities."""
    rows = _sheets_get(f"'{NEW_BUSINESS_TAB}'!A:H")
    if not rows or len(rows) < 2:
        return []
    return rows[1:]


def add_lead(client: str, contact: str = "", source: str = "",
             description: str = "", est_revenue: str = "",
             status: str = "Active", notes: str = "") -> bool:
    """Add a new lead to the Leads tab."""
    date_added = date.today().strftime("%Y-%m-%d")
    row = [date_added, client, contact, source, description, est_revenue, status, notes]
    return _sheets_append(f"'{LEADS_TAB}'!A:H", [row])


def get_leads() -> list:
    """Get all leads."""
    rows = _sheets_get(f"'{LEADS_TAB}'!A:H")
    if not rows or len(rows) < 2:
        return []
    return rows[1:]


def initialize_sheets() -> bool:
    """Initialize all required tabs in the spreadsheet."""
    success = True
    success &= _ensure_tab_exists(ACTIVE_TASKS_TAB, ACTIVE_TASKS_HEADERS)
    success &= _ensure_tab_exists(COMPLETED_TASKS_TAB, COMPLETED_TASKS_HEADERS)
    success &= _ensure_tab_exists(NEW_BUSINESS_TAB, NEW_BUSINESS_HEADERS)
    success &= _ensure_tab_exists(LEADS_TAB, LEADS_HEADERS)
    return success
