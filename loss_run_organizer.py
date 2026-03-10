#!/usr/bin/env python3
"""
Loss Run Organizer for HUB Hotel Franchise Practice.

Watches a Google Drive "inbox" folder for new loss run PDFs, uses GPT to
extract metadata (client, policy type, carrier, valuation date), then:
  1. Moves the file into the correct client → year → policy-type subfolder
  2. Renames the file with a standard naming convention
  3. Updates a Loss Run Tracker Google Sheet

Designed to run on a schedule via APScheduler inside bot.py on Railway.

Requires:
  - GOOGLE_SERVICE_ACCOUNT_JSON env var (service account with Drive + Sheets scope)
  - OPENAI_API_KEY env var
  - LOSS_RUN_INBOX_FOLDER_ID env var (Google Drive folder ID for the inbox)
  - LOSS_RUN_TRACKER_SHEET_ID env var (Google Sheet ID for the tracker)
  - TELEGRAM_TOKEN + TELEGRAM_CHAT_ID env vars (for notifications)
"""

import os
import io
import json
import logging
import time
import tempfile
import re
from datetime import datetime, date

import jwt
import requests as http_requests

logger = logging.getLogger(__name__)

# ── Configuration ─────────────────────────────────────────────────────────
GOOGLE_SERVICE_ACCOUNT_JSON = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "")
OPENAI_API_KEY = os.environ.get("OPENAI_API_KEY", "")
LOSS_RUN_INBOX_FOLDER_ID = os.environ.get("LOSS_RUN_INBOX_FOLDER_ID", "")
LOSS_RUN_TRACKER_SHEET_ID = os.environ.get("LOSS_RUN_TRACKER_SHEET_ID", "")
TELEGRAM_TOKEN = os.environ.get("TELEGRAM_TOKEN", "")
TELEGRAM_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID", "")

# Google API scopes needed
SCOPES = "https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/spreadsheets"

# Policy type mapping (normalized)
POLICY_TYPES = {
    "property": "Property",
    "general liability": "Liability",
    "liability": "Liability",
    "gl": "Liability",
    "commercial general liability": "Liability",
    "workers compensation": "Workers Comp",
    "workers comp": "Workers Comp",
    "wc": "Workers Comp",
    "work comp": "Workers Comp",
    "umbrella": "Umbrella",
    "excess": "Umbrella",
    "excess liability": "Umbrella",
    "auto": "Auto",
    "commercial auto": "Auto",
    "hired non-owned auto": "Auto",
    "epli": "EPLI",
    "employment practices": "EPLI",
    "crime": "Crime",
    "cyber": "Cyber",
    "inland marine": "Inland Marine",
    "equipment breakdown": "Equipment Breakdown",
    "liquor liability": "Liquor Liability",
}

def _get_tracker_sheet_id():
    """Get tracker sheet ID at runtime."""
    return os.environ.get("LOSS_RUN_TRACKER_SHEET_ID", "").strip()


# Tracker sheet headers
TRACKER_HEADERS = [
    "Client", "Policy Type", "Carrier", "Valuation Date",
    "File Name", "Drive Link", "Date Organized", "Year Folder"
]


# ── Google API Auth ───────────────────────────────────────────────────────

def _parse_service_account_json(raw):
    """Parse service account JSON, handling Railway env var mangling."""
    if not raw:
        return None
    raw = raw.strip()
    # Fix truncated JSON: if it starts with { but doesn't end with }, add it
    if raw.startswith('{') and not raw.endswith('}'):
        raw = raw + '\n}'
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        pass
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
    """Get Google API access token with Drive + Sheets scope."""
    sa_json = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not sa_json:
        logger.error("GOOGLE_SERVICE_ACCOUNT_JSON not set")
        return None

    creds = _parse_service_account_json(sa_json)
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
        "scope": SCOPES,
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


def _auth_headers():
    """Get authorization headers."""
    token = _get_access_token()
    if not token:
        return None
    return {"Authorization": f"Bearer {token}"}


# ── Google Drive API Helpers ──────────────────────────────────────────────

def drive_list_files(folder_id, mime_type=None):
    """List files in a Google Drive folder."""
    headers = _auth_headers()
    if not headers:
        logger.error("DIAGNOSTIC: No auth headers available")
        return []

    # DIAGNOSTIC: Check folder access
    try:
        folder_resp = http_requests.get(
            f"https://www.googleapis.com/drive/v3/files/{folder_id}",
            headers=headers, params={"fields": "id, name, mimeType, owners"}, timeout=15
        )
        if folder_resp.status_code == 200:
            folder_data = folder_resp.json()
            logger.info(f"DIAGNOSTIC: Successfully accessed folder: {folder_data.get('name')} ({folder_id})")
        else:
            logger.error(f"DIAGNOSTIC: Failed to access folder metadata. Status: {folder_resp.status_code}")
    except Exception as e:
        logger.error(f"DIAGNOSTIC: Exception checking folder: {e}")

    # Broad query to find all files in the folder first
    q = f"'{folder_id}' in parents and trashed = false"
    logger.info(f"DIAGNOSTIC: Querying Drive with: {q}")

    files = []
    page_token = None

    while True:
        params = {
            "q": q,
            "fields": "nextPageToken, files(id, name, mimeType, modifiedTime, webViewLink)",
            "pageSize": 100,
        }
        if page_token:
            params["pageToken"] = page_token

        resp = http_requests.get(
            "https://www.googleapis.com/drive/v3/files",
            headers=headers, params=params, timeout=30,
        )
        
        if resp.status_code != 200:
            logger.error(f"DIAGNOSTIC: File list request failed. Status: {resp.status_code}")
            break
            
        data = resp.json()
        batch = data.get("files", [])
        logger.info(f"DIAGNOSTIC: Found {len(batch)} items in this batch")
        for f in batch:
            logger.info(f"DIAGNOSTIC: Item found: {f.get('name')} (MIME: {f.get('mimeType')})")
            
        files.extend(batch)
        page_token = data.get("nextPageToken")
        if not page_token:
            break

    # If mime_type was requested, filter manually (case-insensitive)
    if mime_type:
        files = [f for f in files if f.get("mimeType") == mime_type or f.get("name", "").lower().endswith(".pdf")]

    return files


def drive_download_file(file_id):
    """Download a file's content as bytes."""
    headers = _auth_headers()
    if not headers:
        return None

    resp = http_requests.get(
        f"https://www.googleapis.com/drive/v3/files/{file_id}?alt=media",
        headers=headers, timeout=120,
    )
    resp.raise_for_status()
    return resp.content


def drive_create_folder(name, parent_id):
    """Create a folder in Google Drive. Returns the new folder ID."""
    headers = _auth_headers()
    if not headers:
        return None

    metadata = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id],
    }

    resp = http_requests.post(
        "https://www.googleapis.com/drive/v3/files",
        headers={**headers, "Content-Type": "application/json"},
        json=metadata, timeout=30,
    )
    resp.raise_for_status()
    return resp.json()["id"]


def drive_find_or_create_folder(name, parent_id):
    """Find an existing subfolder by name, or create it."""
    headers = _auth_headers()
    if not headers:
        return None

    q = (
        f"'{parent_id}' in parents and "
        f"name = '{name}' and "
        f"mimeType = 'application/vnd.google-apps.folder' and "
        f"trashed = false"
    )
    resp = http_requests.get(
        "https://www.googleapis.com/drive/v3/files",
        headers=headers,
        params={"q": q, "fields": "files(id, name)", "pageSize": 1},
        timeout=30,
    )
    resp.raise_for_status()
    files = resp.json().get("files", [])

    if files:
        return files[0]["id"]
    return drive_create_folder(name, parent_id)


def drive_move_file(file_id, new_parent_id, new_name=None):
    """Move a file to a new folder and optionally rename it."""
    headers = _auth_headers()
    if not headers:
        return False

    # Get current parents
    resp = http_requests.get(
        f"https://www.googleapis.com/drive/v3/files/{file_id}",
        headers=headers, params={"fields": "parents"}, timeout=15,
    )
    resp.raise_for_status()
    current_parents = ",".join(resp.json().get("parents", []))

    # Build update
    params = {
        "addParents": new_parent_id,
        "removeParents": current_parents,
        "fields": "id, parents, name",
    }
    body = {}
    if new_name:
        body["name"] = new_name

    resp = http_requests.patch(
        f"https://www.googleapis.com/drive/v3/files/{file_id}",
        headers={**headers, "Content-Type": "application/json"},
        params=params, json=body, timeout=30,
    )
    resp.raise_for_status()
    return True


def drive_get_web_link(file_id):
    """Get the web view link for a file."""
    headers = _auth_headers()
    if not headers:
        return None

    resp = http_requests.get(
        f"https://www.googleapis.com/drive/v3/files/{file_id}",
        headers=headers, params={"fields": "webViewLink"}, timeout=15,
    )
    resp.raise_for_status()
    return resp.json().get("webViewLink", "")


# ── GPT Extraction ───────────────────────────────────────────────────────

def extract_loss_run_metadata(pdf_bytes, filename):
    """Use GPT to extract client name, policy type, carrier, and valuation date from a loss run PDF."""
    if not OPENAI_API_KEY:
        logger.error("OPENAI_API_KEY not set")
        return None

    # Extract text from PDF using pdfplumber
    text = ""
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            # Read first 5 pages (loss runs usually have key info on page 1-2)
            for i, page in enumerate(pdf.pages[:5]):
                page_text = page.extract_text() or ""
                text += page_text + "\n"
                if len(text) > 8000:
                    break
    except Exception as e:
        logger.warning(f"pdfplumber extraction failed for {filename}: {e}")

    # If pdfplumber got very little text, try OCR via GPT Vision
    if len(text.strip()) < 200:
        logger.info(f"Low text extraction for {filename}, attempting OCR fallback")
        return _extract_via_ocr(pdf_bytes, filename)

    # Ask GPT to extract metadata
    prompt = f"""You are analyzing an insurance loss run document. Extract the following from the text below:

1. **Client Name** (the insured/policyholder, NOT the insurance company)
2. **Policy Type** (one of: Property, Liability, Workers Comp, Umbrella, Auto, EPLI, Crime, Cyber, Inland Marine, Equipment Breakdown, Liquor Liability)
3. **Carrier** (the insurance company that issued the loss run)
4. **Valuation Date** (also called "report date", "as of date", or "valued as of" — the date through which losses are reported. Format as YYYY-MM-DD)

Original filename: {filename}

Document text:
{text[:6000]}

Return ONLY valid JSON with these exact keys:
{{"client_name": "...", "policy_type": "...", "carrier": "...", "valuation_date": "YYYY-MM-DD"}}

If you cannot determine a field, use "Unknown" for strings or "1900-01-01" for date.
"""

    try:
        resp = http_requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {OPENAI_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": "gpt-4.1-mini",
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.0,
                "max_tokens": 500,
            },
            timeout=60,
        )
        resp.raise_for_status()
        content = resp.json()["choices"][0]["message"]["content"]

        # Parse JSON from response
        json_match = re.search(r'\{[^}]+\}', content, re.DOTALL)
        if json_match:
            metadata = json.loads(json_match.group())
            # Normalize policy type
            metadata["policy_type"] = _normalize_policy_type(metadata.get("policy_type", "Unknown"))
            logger.info(f"Extracted metadata for {filename}: {metadata}")
            return metadata
    except Exception as e:
        logger.error(f"GPT extraction failed for {filename}: {e}")

    return None


def _extract_via_ocr(pdf_bytes, filename):
    """OCR fallback using GPT Vision for scanned PDFs."""
    try:
        from pdf2image import convert_from_bytes
        import base64

        images = convert_from_bytes(pdf_bytes, first_page=1, last_page=2, dpi=150, fmt="jpeg")
        if not images:
            return None

        # Convert first page to base64
        buf = io.BytesIO()
        images[0].save(buf, format="JPEG", quality=70)
        img_b64 = base64.b64encode(buf.getvalue()).decode()

        resp = http_requests.post(
            "https://api.openai.com/v1/chat/completions",
            headers={
                "Authorization": f"Bearer {OPENAI_API_KEY}",
                "Content-Type": "application/json",
            },
            json={
                "model": "gpt-4.1-mini",
                "messages": [{
                    "role": "user",
                    "content": [
                        {"type": "text", "text": (
                            "This is a scanned insurance loss run document. Extract:\n"
                            "1. Client Name (the insured)\n"
                            "2. Policy Type (Property, Liability, Workers Comp, Umbrella, Auto, etc.)\n"
                            "3. Carrier (insurance company)\n"
                            "4. Valuation Date (as YYYY-MM-DD)\n\n"
                            f"Original filename: {filename}\n\n"
                            'Return ONLY valid JSON: {"client_name": "...", "policy_type": "...", "carrier": "...", "valuation_date": "YYYY-MM-DD"}'
                        )},
                        {"type": "image_url", "image_url": {"url": f"data:image/jpeg;base64,{img_b64}", "detail": "low"}},
                    ],
                }],
                "temperature": 0.0,
                "max_tokens": 500,
            },
            timeout=60,
        )
        resp.raise_for_status()
        content = resp.json()["choices"][0]["message"]["content"]
        json_match = re.search(r'\{[^}]+\}', content, re.DOTALL)
        if json_match:
            metadata = json.loads(json_match.group())
            metadata["policy_type"] = _normalize_policy_type(metadata.get("policy_type", "Unknown"))
            return metadata
    except Exception as e:
        logger.error(f"OCR extraction failed for {filename}: {e}")

    return None


def _normalize_policy_type(raw_type):
    """Normalize policy type string to standard categories."""
    if not raw_type:
        return "Other"
    key = raw_type.strip().lower()
    return POLICY_TYPES.get(key, raw_type.title())


# ── Tracker Sheet ─────────────────────────────────────────────────────────

def _sheets_headers():
    """Get auth headers for Sheets API (reuses Drive token with combined scope)."""
    headers = _auth_headers()
    if not headers:
        return None
    return {**headers, "Content-Type": "application/json"}


def tracker_initialize():
    """Ensure the tracker sheet has headers."""
    sheet_id = os.environ.get("LOSS_RUN_TRACKER_SHEET_ID", "").strip()
    if not sheet_id:
        logger.warning("LOSS_RUN_TRACKER_SHEET_ID not set - tracker disabled")
        return False

    headers = _sheets_headers()
    if not headers:
        return False

    # Check if headers exist
    sid = _get_tracker_sheet_id()
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sid}/values/Sheet1!A1:H1"
    try:
        resp = http_requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
        values = resp.json().get("values", [])
        if not values or values[0] != TRACKER_HEADERS:
            # Write headers
            body = {"values": [TRACKER_HEADERS]}
            url = (
                f"https://sheets.googleapis.com/v4/spreadsheets/{sid}"
                f"/values/Sheet1!A1:H1?valueInputOption=USER_ENTERED"
            )
            resp = http_requests.put(url, headers=headers, json=body, timeout=15)
            resp.raise_for_status()
            logger.info("Tracker sheet headers initialized")
        return True
    except Exception as e:
        logger.error(f"Tracker initialization failed: {e}")
        return False


def tracker_add_entry(client, policy_type, carrier, valuation_date, filename, drive_link, year_folder):
    """Add or update an entry in the tracker sheet."""
    if not os.environ.get("LOSS_RUN_TRACKER_SHEET_ID", "").strip():
        return False

    headers = _sheets_headers()
    if not headers:
        return False

    today = date.today().isoformat()
    row = [client, policy_type, carrier, valuation_date, filename, drive_link, today, year_folder]

    # First, check if there's an existing row for this client + policy type to update
    try:
        url = f"https://sheets.googleapis.com/v4/spreadsheets/{_get_tracker_sheet_id()}/values/Sheet1!A:H"
        resp = http_requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
        all_rows = resp.json().get("values", [])

        # Look for existing entry (same client + policy type) — update if this valuation is newer
        for i, existing_row in enumerate(all_rows[1:], start=2):  # skip header
            if len(existing_row) >= 2:
                if (existing_row[0].strip().lower() == client.strip().lower() and
                        existing_row[1].strip().lower() == policy_type.strip().lower()):
                    # Found match — check if new valuation is more recent
                    existing_val = existing_row[3] if len(existing_row) > 3 else ""
                    if valuation_date >= existing_val:
                        # Update this row
                        range_str = f"Sheet1!A{i}:H{i}"
                        url = (
                            f"https://sheets.googleapis.com/v4/spreadsheets/{_get_tracker_sheet_id()}"
                            f"/values/{range_str}?valueInputOption=USER_ENTERED"
                        )
                        resp = http_requests.put(url, headers=headers, json={"values": [row]}, timeout=15)
                        resp.raise_for_status()
                        logger.info(f"Updated tracker: {client} / {policy_type} → {valuation_date}")
                        return True
                    else:
                        logger.info(f"Skipping tracker update — existing valuation {existing_val} is newer")
                        return True

        # No existing entry — append new row
        url = (
            f"https://sheets.googleapis.com/v4/spreadsheets/{_get_tracker_sheet_id()}"
            f"/values/Sheet1!A:H:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS"
        )
        resp = http_requests.post(url, headers=headers, json={"values": [row]}, timeout=15)
        resp.raise_for_status()
        logger.info(f"Added to tracker: {client} / {policy_type} / {carrier} / {valuation_date}")
        return True

    except Exception as e:
        logger.error(f"Tracker update failed: {e}")
        return False


def tracker_get_all():
    """Get all entries from the tracker sheet."""
    if not os.environ.get("LOSS_RUN_TRACKER_SHEET_ID", "").strip():
        return []

    headers = _sheets_headers()
    if not headers:
        return []

    try:
        url = f"https://sheets.googleapis.com/v4/spreadsheets/{_get_tracker_sheet_id()}/values/Sheet1!A:H"
        resp = http_requests.get(url, headers=headers, timeout=15)
        resp.raise_for_status()
        rows = resp.json().get("values", [])
        if len(rows) <= 1:
            return []
        # Return as list of dicts
        header = rows[0]
        return [dict(zip(header, row + [""] * (len(header) - len(row)))) for row in rows[1:]]
    except Exception as e:
        logger.error(f"Tracker read failed: {e}")
        return []


def tracker_get_client(client_name):
    """Get tracker entries for a specific client."""
    all_entries = tracker_get_all()
    return [e for e in all_entries if client_name.lower() in e.get("Client", "").lower()]


# ── Client Folder Resolver ────────────────────────────────────────────────

def find_client_folder(client_name, accounts_folder_id):
    """
    Find the client's folder inside the Accounts folder.
    Uses fuzzy matching — searches for folders whose name contains the client name.
    If no match found, creates a new folder.
    """
    headers = _auth_headers()
    if not headers:
        return None, None

    # Search for folders that might match this client
    # Use a simpler search — get all folders in the accounts parent
    q = (
        f"'{accounts_folder_id}' in parents and "
        f"mimeType = 'application/vnd.google-apps.folder' and "
        f"trashed = false"
    )

    resp = http_requests.get(
        "https://www.googleapis.com/drive/v3/files",
        headers=headers,
        params={"q": q, "fields": "files(id, name)", "pageSize": 500},
        timeout=30,
    )
    resp.raise_for_status()
    folders = resp.json().get("files", [])

    # Try to match — check if client_name appears in folder name or vice versa
    client_lower = client_name.strip().lower()
    best_match = None
    best_score = 0

    for folder in folders:
        folder_lower = folder["name"].strip().lower()
        # Exact containment in either direction
        if client_lower in folder_lower or folder_lower in client_lower:
            score = len(client_lower) / max(len(folder_lower), 1)
            if score > best_score:
                best_score = score
                best_match = folder

    if best_match:
        logger.info(f"Matched client '{client_name}' to folder '{best_match['name']}'")
        return best_match["id"], best_match["name"]

    # No match — create new folder
    logger.info(f"No matching folder for '{client_name}' — creating new one")
    new_id = drive_create_folder(client_name, accounts_folder_id)
    return new_id, client_name


# ── Main Organizer Logic ─────────────────────────────────────────────────

def organize_loss_runs(accounts_folder_id=None):
    """
    Main entry point. Scans the inbox folder, processes each PDF,
    and organizes into the correct location.

    Returns a summary dict with counts and details.
    """
    # Read env vars at runtime (not import time) to ensure Railway vars are available
    inbox_id = os.environ.get("LOSS_RUN_INBOX_FOLDER_ID", "").strip()
    logger.info(f"LOSS_RUN_INBOX_FOLDER_ID = '{inbox_id[:8]}...' (len={len(inbox_id)})")
    logger.info(f"All env keys with LOSS: {[k for k in os.environ if 'LOSS' in k]}")
    if not inbox_id:
        logger.error("LOSS_RUN_INBOX_FOLDER_ID not set")
        return {"error": "Inbox folder not configured", "processed": 0}

    if not accounts_folder_id:
        accounts_folder_id = os.environ.get("ACCOUNTS_FOLDER_ID", "").strip()
    if not accounts_folder_id:
        logger.error("ACCOUNTS_FOLDER_ID not set")
        return {"error": "Accounts folder not configured", "processed": 0}

    # Initialize tracker
    tracker_initialize()

    # Get all files in inbox
    inbox_files = drive_list_files(inbox_id)
    pdf_files = [f for f in inbox_files if f["name"].lower().endswith(".pdf")]

    if not pdf_files:
        logger.info("No PDF files in inbox folder")
        return {"processed": 0, "message": "No loss runs to process"}

    results = {
        "processed": 0,
        "success": [],
        "errors": [],
    }

    for file_info in pdf_files:
        file_id = file_info["id"]
        filename = file_info["name"]
        logger.info(f"Processing: {filename}")

        try:
            # 1. Download the PDF
            pdf_bytes = drive_download_file(file_id)
            if not pdf_bytes:
                results["errors"].append(f"{filename}: Download failed")
                continue

            # 2. Extract metadata via GPT
            metadata = extract_loss_run_metadata(pdf_bytes, filename)
            if not metadata:
                results["errors"].append(f"{filename}: Extraction failed")
                continue

            client_name = metadata.get("client_name", "Unknown")
            policy_type = metadata.get("policy_type", "Other")
            carrier = metadata.get("carrier", "Unknown")
            valuation_date = metadata.get("valuation_date", "1900-01-01")

            # Determine year from valuation date
            try:
                val_year = str(datetime.strptime(valuation_date, "%Y-%m-%d").year)
            except ValueError:
                val_year = str(date.today().year)

            # 3. Find or create client folder
            client_folder_id, client_folder_name = find_client_folder(client_name, accounts_folder_id)
            if not client_folder_id:
                results["errors"].append(f"{filename}: Could not resolve client folder")
                continue

            # 4. Create subfolder path: Client / Loss Runs / {Year} / {Policy Type}
            loss_runs_folder_id = drive_find_or_create_folder("Loss Runs", client_folder_id)
            year_folder_id = drive_find_or_create_folder(val_year, loss_runs_folder_id)
            policy_folder_id = drive_find_or_create_folder(policy_type, year_folder_id)

            # 5. Rename file: YYYY-MM-DD_Carrier_PolicyType_LossRun.pdf
            safe_carrier = re.sub(r'[^\w\s-]', '', carrier).strip().replace(' ', '_')
            safe_policy = policy_type.replace(' ', '_')
            new_name = f"{valuation_date}_{safe_carrier}_{safe_policy}_LossRun.pdf"

            # 6. Move file to destination
            drive_move_file(file_id, policy_folder_id, new_name)

            # 7. Get the web link for the tracker
            web_link = drive_get_web_link(file_id)

            # 8. Update tracker sheet
            tracker_add_entry(
                client=client_folder_name or client_name,
                policy_type=policy_type,
                carrier=carrier,
                valuation_date=valuation_date,
                filename=new_name,
                drive_link=web_link or "",
                year_folder=val_year,
            )

            results["processed"] += 1
            results["success"].append({
                "original": filename,
                "new_name": new_name,
                "client": client_folder_name or client_name,
                "policy_type": policy_type,
                "carrier": carrier,
                "valuation_date": valuation_date,
                "year": val_year,
            })
            logger.info(f"✓ Organized: {filename} → {client_folder_name}/{val_year}/{policy_type}/{new_name}")

        except Exception as e:
            logger.error(f"Error processing {filename}: {e}")
            results["errors"].append(f"{filename}: {str(e)}")

    return results


# ── Telegram Notification ─────────────────────────────────────────────────

def send_organize_summary(results):
    """Send a summary of the organization run to Telegram."""
    if not TELEGRAM_TOKEN or not TELEGRAM_CHAT_ID:
        return

    lines = ["📂 *Loss Run Organizer Summary*\n"]

    if results.get("error"):
        lines.append(f"⚠️ Error: {results['error']}")
    elif results["processed"] == 0:
        lines.append("No new loss runs to process\\.")
    else:
        lines.append(f"✅ *{results['processed']}* loss run\\(s\\) organized:\n")
        for item in results.get("success", []):
            lines.append(
                f"• *{_escape_md(item['client'])}*\n"
                f"  {_escape_md(item['policy_type'])} \\| {_escape_md(item['carrier'])} \\| {_escape_md(item['valuation_date'])}"
            )

    if results.get("errors"):
        lines.append(f"\n⚠️ *{len(results['errors'])}* error\\(s\\):")
        for err in results["errors"][:5]:
            lines.append(f"  • {_escape_md(err)}")

    text = "\n".join(lines)

    try:
        http_requests.post(
            f"https://api.telegram.org/bot{TELEGRAM_TOKEN}/sendMessage",
            json={
                "chat_id": TELEGRAM_CHAT_ID,
                "text": text,
                "parse_mode": "MarkdownV2",
            },
            timeout=15,
        )
    except Exception as e:
        logger.error(f"Telegram notification failed: {e}")


def _escape_md(text):
    """Escape Telegram MarkdownV2 special characters."""
    special = r'_*[]()~`>#+-=|{}.!'
    return ''.join(f'\\{c}' if c in special else c for c in str(text))


# ── Scheduled Run Entry Point ─────────────────────────────────────────────

async def scheduled_organize():
    """Entry point for APScheduler. Runs synchronously in a thread."""
    import asyncio
    loop = asyncio.get_event_loop()
    results = await loop.run_in_executor(None, organize_loss_runs)
    send_organize_summary(results)
    return results


def run_organize_sync():
    """Synchronous entry point for testing or manual runs."""
    results = organize_loss_runs()
    send_organize_summary(results)
    return results


# ── CLI for Testing ───────────────────────────────────────────────────────

if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO)
    results = run_organize_sync()
    print(json.dumps(results, indent=2))
