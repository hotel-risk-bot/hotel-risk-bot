"""
Hotel Insurance Proposal - Web Application
Flask-based web interface for uploading quote PDFs/SOVs, reviewing extracted data,
editing fields, and generating branded DOCX proposals.

Runs alongside the existing Telegram bot on a separate port.
"""

import os
import json
import logging
import tempfile
import uuid
import threading
from pathlib import Path
from datetime import datetime

from flask import Flask, request, jsonify, send_file, render_template

from proposal_extractor import ProposalExtractor, extract_text_from_pdf_smart, extract_text_from_excel
from proposal_generator import generate_proposal
from sov_parser import parse_sov, is_sov_file, aggregate_locations

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# ─── Launch Telegram bot as a subprocess alongside the web app ───
_bot_process = None

def _start_bot_subprocess():
    """Start bot.py as a subprocess if not already running.
    This ensures the Telegram bot runs alongside gunicorn."""
    global _bot_process
    import subprocess
    import sys
    if _bot_process is not None and _bot_process.poll() is None:
        logger.info("Bot subprocess already running (pid=%s)", _bot_process.pid)
        return
    bot_script = os.path.join(os.path.dirname(os.path.abspath(__file__)), "bot.py")
    if not os.path.exists(bot_script):
        logger.warning("bot.py not found at %s - Telegram bot will not start", bot_script)
        return
    logger.info("Starting Telegram bot subprocess: %s", bot_script)
    _bot_process = subprocess.Popen(
        [sys.executable, bot_script],
        stdout=None,
        stderr=None,
        env=os.environ.copy(),
    )
    logger.info("Telegram bot subprocess started (pid=%s)", _bot_process.pid)

# Auto-start the bot when this module is loaded by gunicorn
# Use a file-based lock to ensure only one worker starts the bot
_BOT_LOCK_FILE = os.path.join(tempfile.gettempdir(), "hotel_risk_bot.lock")

# Clean up stale lock file from previous deployments
try:
    os.unlink(_BOT_LOCK_FILE)
except FileNotFoundError:
    pass

def _should_start_bot():
    """Check if this is the first worker to start (use file lock)."""
    try:
        fd = os.open(_BOT_LOCK_FILE, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        os.write(fd, str(os.getpid()).encode())
        os.close(fd)
        return True
    except FileExistsError:
        # Another worker already started the bot
        return False

if os.environ.get("TELEGRAM_TOKEN") or os.environ.get("TELEGRAM_BOT_TOKEN"):
    if _should_start_bot():
        _start_bot_subprocess()
    else:
        logger.info("Bot subprocess already started by another worker - skipping")
else:
    logger.warning("No TELEGRAM_TOKEN or TELEGRAM_BOT_TOKEN found - Telegram bot will not start")

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max upload

# File-backed session storage to survive across workers/restarts
UPLOAD_DIR = os.environ.get("UPLOAD_DIR", tempfile.mkdtemp(prefix="proposal_web_"))
os.makedirs(UPLOAD_DIR, exist_ok=True)
SESSION_FILE = os.path.join(UPLOAD_DIR, "_sessions.json")
logger.info(f"Upload directory: {UPLOAD_DIR}")


def _load_sessions():
    """Load sessions from disk."""
    try:
        if os.path.exists(SESSION_FILE):
            with open(SESSION_FILE, 'r') as f:
                return json.load(f)
    except Exception as e:
        logger.warning(f"Failed to load sessions: {e}")
    return {}


def _save_sessions(sessions):
    """Save sessions to disk."""
    try:
        with open(SESSION_FILE, 'w') as f:
            json.dump(sessions, f)
    except Exception as e:
        logger.warning(f"Failed to save sessions: {e}")


def _get_session(session_id):
    """Get a session by ID."""
    sessions = _load_sessions()
    return sessions.get(session_id)


def _set_session(session_id, data):
    """Set a session by ID."""
    sessions = _load_sessions()
    sessions[session_id] = data
    _save_sessions(sessions)


# ─── Helper functions (ported from proposal_handler.py) ───

def _normalize_coverages(data):
    """Ensure coverages is always a dict with dict values, not a list."""
    if data is None:
        return data
    covs = data.get("coverages", {})
    if isinstance(covs, list):
        normalized = {}
        for item in covs:
            if isinstance(item, dict):
                cov_type = item.get("coverage_type", item.get("type", "unknown"))
                normalized[cov_type] = item
            elif isinstance(item, str):
                normalized[item] = {}
        data["coverages"] = normalized
    elif not isinstance(covs, dict):
        data["coverages"] = {}

    covs = data.get("coverages", {})
    if isinstance(covs, dict):
        for key, val in list(covs.items()):
            if isinstance(val, list):
                if len(val) == 1 and isinstance(val[0], dict):
                    covs[key] = val[0]
                elif len(val) > 1:
                    for item in val:
                        if isinstance(item, dict):
                            covs[key] = item
                            break
                    else:
                        covs[key] = {}
                else:
                    covs[key] = {}
            elif not isinstance(val, dict):
                covs[key] = {}

    return data


def _merge_extraction_results(existing, new_data):
    """Merge extraction results from multiple PDFs into a single data structure."""
    if not existing:
        return _normalize_coverages(new_data)
    if not new_data or "error" in new_data:
        return existing

    _normalize_coverages(existing)
    _normalize_coverages(new_data)

    merged = json.loads(json.dumps(existing))

    # Merge client_info
    existing_ci = merged.get("client_info", {})
    new_ci = new_data.get("client_info", {})
    for key, val in new_ci.items():
        if val and val != "N/A" and (not existing_ci.get(key) or existing_ci.get(key) == "N/A"):
            existing_ci[key] = val
    merged["client_info"] = existing_ci

    # Merge coverages with auto-promotion and standalone prioritization
    existing_covs = merged.get("coverages", {})
    new_covs = new_data.get("coverages", {})

    def _cov_premium(cov):
        """Extract numeric premium from a coverage dict."""
        if not isinstance(cov, dict):
            return 0
        p = cov.get("total_premium") or cov.get("premium") or 0
        try:
            return float(str(p).replace(",", "").replace("$", ""))
        except (ValueError, TypeError):
            return 0

    def _is_standalone(cov):
        """Check if a coverage appears to be a standalone policy (has its own carrier and premium)."""
        if not isinstance(cov, dict):
            return False
        has_carrier = bool(cov.get("carrier", "").strip())
        has_premium = _cov_premium(cov) > 0
        has_limits = bool(cov.get("limits"))
        return has_carrier and (has_premium or has_limits)

    for cov_key, cov_data in new_covs.items():
        if cov_key not in existing_covs:
            existing_covs[cov_key] = cov_data
        elif cov_key == "umbrella":
            if "umbrella_layer_2" not in existing_covs:
                existing_covs["umbrella_layer_2"] = cov_data
            elif "umbrella_layer_3" not in existing_covs:
                existing_covs["umbrella_layer_3"] = cov_data
        elif cov_key == "property":
            if "excess_property" not in existing_covs:
                existing_covs["excess_property"] = cov_data
            elif "excess_property_2" not in existing_covs:
                existing_covs["excess_property_2"] = cov_data
        else:
            # For other coverages (cyber, epli, crime, etc.): prioritize standalone
            # policy over bundled/incidental coverage from another quote
            existing_cov = existing_covs[cov_key]
            new_is_standalone = _is_standalone(cov_data)
            existing_is_standalone = _is_standalone(existing_cov)
            new_premium = _cov_premium(cov_data)
            existing_premium = _cov_premium(existing_cov)
            # Replace if: new is standalone and existing isn't, or new has higher premium
            if (new_is_standalone and not existing_is_standalone) or \
               (new_is_standalone and existing_is_standalone and new_premium > existing_premium):
                logger.info(f"Replacing {cov_key} coverage: {existing_cov.get('carrier', 'unknown')} "
                           f"(${existing_premium:,.0f}) -> {cov_data.get('carrier', 'unknown')} "
                           f"(${new_premium:,.0f})")
                existing_covs[cov_key] = cov_data
    merged["coverages"] = existing_covs

    # Post-merge validation: detect excess LIABILITY misclassified as excess PROPERTY
    # If excess_property has underlying_insurance referencing umbrella/GL, it's actually an umbrella layer
    for ep_key in ["excess_property", "excess_property_2"]:
        ep_data = existing_covs.get(ep_key)
        if not isinstance(ep_data, dict):
            continue
        # Check indicators that this is actually an excess liability, not excess property
        is_liability = False
        # Check 1: underlying_insurance references umbrella or GL
        underlying = ep_data.get("underlying_insurance", [])
        for u in (underlying if isinstance(underlying, list) else []):
            u_cov = str(u.get("coverage", "")).lower() if isinstance(u, dict) else ""
            if any(kw in u_cov for kw in ["umbrella", "general liability", "gl", "excess"]):
                is_liability = True
                break
        # Check 2: tower_structure mentions "xs primary" or "excess of" umbrella
        tower = ep_data.get("tower_structure", [])
        for t in (tower if isinstance(tower, list) else []):
            t_limits = str(t.get("limits", "")).lower() if isinstance(t, dict) else ""
            if "xs" in t_limits and "primary" in t_limits:
                is_liability = True
                break
        # Check 3: carrier name or coverage description contains "excess liability"
        carrier_str = str(ep_data.get("carrier", "")).lower()
        # Check if there are no property-specific fields (no TIV, no building values, no coinsurance)
        has_property_fields = bool(ep_data.get("tiv") or ep_data.get("coinsurance") or ep_data.get("valuation"))
        if not has_property_fields and not is_liability:
            # Additional check: look at the forms/endorsements for liability indicators
            forms = ep_data.get("forms_endorsements", [])
            for f in (forms if isinstance(forms, list) else []):
                f_desc = str(f.get("description", f) if isinstance(f, dict) else f).lower()
                if "liability" in f_desc or "umbrella" in f_desc:
                    is_liability = True
                    break
        if is_liability:
            # Move to the next available umbrella slot
            moved = False
            for umb_slot in ["umbrella", "umbrella_layer_2", "umbrella_layer_3"]:
                if umb_slot not in existing_covs:
                    existing_covs[umb_slot] = existing_covs.pop(ep_key)
                    logger.info(f"Post-merge fix: moved {ep_key} ({ep_data.get('carrier', 'unknown')}) -> {umb_slot} (was misclassified as excess property)")
                    moved = True
                    break
            if not moved:
                logger.warning(f"Post-merge: {ep_key} appears to be excess liability but all umbrella slots are full")

    # Merge locations
    existing_locs = merged.get("locations", [])
    existing_addrs = {loc.get("address", "").upper() for loc in existing_locs}
    for loc in new_data.get("locations", []):
        addr = loc.get("address", "").upper()
        if addr and addr not in existing_addrs:
            existing_locs.append(loc)
            existing_addrs.add(addr)
    merged["locations"] = existing_locs

    # Merge named insureds
    existing_named = merged.get("named_insureds", [])
    if not isinstance(existing_named, list):
        existing_named = [existing_named] if existing_named else []
    def _ni_name(ni):
        if isinstance(ni, dict):
            return ni.get("name", "").strip().upper()
        return str(ni).strip().upper()
    seen_names = {_ni_name(ni) for ni in existing_named}
    for ni in new_data.get("named_insureds", []):
        name_key = _ni_name(ni)
        if name_key and name_key not in seen_names:
            existing_named.append(ni)
            seen_names.add(name_key)
    merged["named_insureds"] = existing_named

    # Merge additional interests
    existing_ai = merged.get("additional_interests", [])
    existing_ai_names = {ai.get("name_address", "") for ai in existing_ai}
    for ai in new_data.get("additional_interests", []):
        if ai.get("name_address", "") not in existing_ai_names:
            existing_ai.append(ai)
    merged["additional_interests"] = existing_ai

    # Merge expiring premiums
    existing_exp = merged.get("expiring_premiums", {})
    new_exp = new_data.get("expiring_premiums", {})
    for key, val in new_exp.items():
        if val and val != 0 and (not existing_exp.get(key) or existing_exp.get(key) == 0):
            existing_exp[key] = val
    merged["expiring_premiums"] = existing_exp

    # Merge payment options
    existing_pay = merged.get("payment_options", [])
    for po in new_data.get("payment_options", []):
        existing_pay.append(po)
    merged["payment_options"] = existing_pay

    return merged


def _enrich_with_sov(extracted_data, sov_data):
    """Enrich extracted data with SOV information (DBA, locations, etc.)."""
    logger.info(f"_enrich_with_sov called. sov_data type={type(sov_data).__name__}, "
                f"has_error={'error' in sov_data if sov_data else 'N/A'}, "
                f"num_locations={len(sov_data.get('locations', [])) if sov_data else 0}")
    if not sov_data or "error" in sov_data:
        logger.warning(f"_enrich_with_sov: skipping - sov_data is {'empty/None' if not sov_data else 'has error: ' + str(sov_data.get('error'))}")
        return

    # Store SOV data
    extracted_data["sov_data"] = sov_data

    # Populate locations from SOV
    sov_locations = []
    for loc in sov_data.get("locations", []):
        loc_entry = {
            "name": loc.get("dba") or loc.get("hotel_flag") or loc.get("corporate_name", ""),
            "address": loc.get("address", ""),
            "city": loc.get("city", ""),
            "state": loc.get("state", ""),
            "zip": loc.get("zip_code", ""),
            "rooms": loc.get("num_rooms", 0),
            "tiv": loc.get("tiv", 0),
            "building_value": loc.get("building_value", 0),
            "contents_value": loc.get("contents_value", 0),
            "bi_value": loc.get("bi_value", 0),
            "construction": loc.get("construction_type", ""),
            "year_built": loc.get("year_built", 0),
            "stories": loc.get("stories", 0),
            "sprinkler": loc.get("sprinkler_pct", ""),
            "roof_type": loc.get("roof_type", ""),
            "roof_year": loc.get("roof_year", 0),
            "flood_zone": loc.get("flood_zone", ""),
            "aop_deductible": loc.get("aop_deductible", 0),
        }
        sov_locations.append(loc_entry)

    # REPLACE GPT-extracted locations with SOV locations (SOV is authoritative for property data)
    logger.info(f"_enrich_with_sov: replacing {len(extracted_data.get('locations', []))} GPT locations "
                f"with {len(sov_locations)} SOV locations")
    for i, sl in enumerate(sov_locations[:3]):
        logger.info(f"  SOV loc {i+1}: name='{sl.get('name', '')}', tiv={sl.get('tiv', 0)}, addr='{sl.get('address', '')}'")
    extracted_data["locations"] = sov_locations
    extracted_data["sov_totals"] = sov_data.get("totals", {})

    # Enrich DBA from SOV named_insured (e.g., "LV Hotels LLC - Latitude Apartment")
    sov_summary = sov_data.get("summary", {})
    sov_named = sov_summary.get("named_insured", "")
    ci = extracted_data.get("client_info", {})
    if sov_named and " - " in sov_named:
        parts = sov_named.split(" - ", 1)
        sov_dba = parts[1].strip()
        if not ci.get("dba"):
            ci["dba"] = sov_dba
        named_insureds = extracted_data.get("named_insureds", [])
        if named_insureds:
            first_ni = named_insureds[0]
            if isinstance(first_ni, dict) and not first_ni.get("dba"):
                first_ni["dba"] = sov_dba

    # Enrich GL schedule_of_classes with SOV address/brand data
    coverages = extracted_data.get("coverages", {})
    gl_cov = coverages.get("general_liability", {})
    sov_locs = sov_data.get("locations", [])
    if gl_cov and sov_locs:
        classes = gl_cov.get("schedule_of_classes", [])
        if classes:
            import re
            sov_lookup = {}
            for loc in sov_locs:
                loc_num = loc.get("location_num", loc.get("building_num", 0))
                if loc_num:
                    sov_lookup[str(loc_num)] = loc

            for cls_entry in classes:
                if not isinstance(cls_entry, dict):
                    continue
                if not cls_entry.get("address") or not cls_entry.get("brand_dba"):
                    loc_str = str(cls_entry.get("location", ""))
                    loc_match = re.search(r"(\d+)", loc_str)
                    if loc_match:
                        loc_num = loc_match.group(1)
                        sov_loc = sov_lookup.get(loc_num)
                        if sov_loc:
                            if not cls_entry.get("address"):
                                addr = sov_loc.get("address", "")
                                city = sov_loc.get("city", "")
                                state = sov_loc.get("state", "")
                                if addr:
                                    cls_entry["address"] = f"{addr}, {city}, {state}" if city else addr
                            if not cls_entry.get("brand_dba"):
                                cls_entry["brand_dba"] = sov_loc.get("dba", "") or sov_loc.get("hotel_flag", "")


# ─── Routes ───

@app.route("/")
def index():
    """Serve the main web interface."""
    return render_template("proposal_web.html")


@app.route("/api/drive-diagnostic")
def drive_diagnostic():
    """Diagnostic endpoint to test Google Drive API access directly."""
    import time as _time
    results = {"timestamp": datetime.now().isoformat(), "steps": []}

    # Step 1: Check env vars
    inbox_id = os.environ.get("LOSS_RUN_INBOX_FOLDER_ID", "").strip()
    sa_json_raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    results["inbox_folder_id"] = inbox_id
    results["sa_json_present"] = bool(sa_json_raw)
    results["sa_json_length"] = len(sa_json_raw)

    if not inbox_id:
        results["error"] = "LOSS_RUN_INBOX_FOLDER_ID not set"
        return jsonify(results), 500
    if not sa_json_raw:
        results["error"] = "GOOGLE_SERVICE_ACCOUNT_JSON not set"
        return jsonify(results), 500

    # Step 2: Parse service account JSON using the same robust parser as loss_run_organizer
    try:
        from loss_run_organizer import _parse_service_account_json
        sa_creds = _parse_service_account_json(sa_json_raw)
        if sa_creds:
            results["steps"].append("JSON parsed via loss_run_organizer parser")
        else:
            # Show first/last 100 chars for debugging
            results["sa_json_first_100"] = sa_json_raw[:100]
            results["sa_json_last_100"] = sa_json_raw[-100:]
            results["error"] = "loss_run_organizer._parse_service_account_json returned None"
            return jsonify(results), 500
    except Exception as e:
        results["sa_json_first_100"] = sa_json_raw[:100]
        results["sa_json_last_100"] = sa_json_raw[-100:]
        results["error"] = f"Parser import/call failed: {e}"
        return jsonify(results), 500

    results["client_email"] = sa_creds.get("client_email", "MISSING")

    # Step 3: Get access token
    try:
        import jwt as _jwt
        pk = sa_creds.get("private_key", "")
        if pk and '\n' not in pk and '\\n' in pk:
            sa_creds["private_key"] = pk.replace('\\n', '\n')
        now = int(_time.time())
        payload = {
            "iss": sa_creds["client_email"],
            "scope": "https://www.googleapis.com/auth/drive https://www.googleapis.com/auth/spreadsheets",
            "aud": "https://oauth2.googleapis.com/token",
            "iat": now, "exp": now + 3600,
        }
        signed_jwt = _jwt.encode(payload, sa_creds["private_key"], algorithm="RS256")
        import requests as _req
        token_resp = _req.post("https://oauth2.googleapis.com/token", data={
            "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
            "assertion": signed_jwt,
        }, timeout=15)
        token_resp.raise_for_status()
        access_token = token_resp.json()["access_token"]
        results["steps"].append(f"Got access token (len={len(access_token)})")
    except Exception as e:
        results["error"] = f"Auth failed: {e}"
        return jsonify(results), 500

    headers = {"Authorization": f"Bearer {access_token}"}

    # Step 4: Check folder metadata
    try:
        import requests as _req
        folder_resp = _req.get(
            f"https://www.googleapis.com/drive/v3/files/{inbox_id}",
            headers=headers, params={"fields": "id,name,mimeType"}, timeout=15
        )
        results["folder_metadata_status"] = folder_resp.status_code
        if folder_resp.status_code == 200:
            results["folder_metadata"] = folder_resp.json()
            results["steps"].append(f"Folder accessible: {folder_resp.json().get('name')}")
        else:
            results["folder_metadata_error"] = folder_resp.text
            results["steps"].append(f"Folder access failed: {folder_resp.status_code}")
    except Exception as e:
        results["steps"].append(f"Folder check exception: {e}")

    # Step 5: List files in folder
    try:
        import requests as _req
        q = f"'{inbox_id}' in parents and trashed = false"
        list_resp = _req.get(
            "https://www.googleapis.com/drive/v3/files",
            headers=headers,
            params={"q": q, "fields": "files(id,name,mimeType)", "pageSize": 100},
            timeout=30
        )
        results["file_list_status"] = list_resp.status_code
        if list_resp.status_code == 200:
            files = list_resp.json().get("files", [])
            results["file_count"] = len(files)
            results["files"] = [{"name": f["name"], "mimeType": f.get("mimeType", "")} for f in files[:30]]
            results["steps"].append(f"Listed {len(files)} files")
        else:
            results["file_list_error"] = list_resp.text
            results["steps"].append(f"File list failed: {list_resp.status_code}")
    except Exception as e:
        results["steps"].append(f"File list exception: {e}")

    # Step 6: Also try with supportsAllDrives for comparison
    try:
        import requests as _req
        q2 = f"'{inbox_id}' in parents and trashed = false"
        list_resp2 = _req.get(
            "https://www.googleapis.com/drive/v3/files",
            headers=headers,
            params={
                "q": q2, "fields": "files(id,name,mimeType)", "pageSize": 100,
                "corpora": "allDrives", "supportsAllDrives": "true",
                "includeItemsFromAllDrives": "true"
            },
            timeout=30
        )
        if list_resp2.status_code == 200:
            files2 = list_resp2.json().get("files", [])
            results["allDrives_file_count"] = len(files2)
            results["steps"].append(f"allDrives query: {len(files2)} files")
        else:
            results["allDrives_error"] = list_resp2.text
            results["steps"].append(f"allDrives query failed: {list_resp2.status_code}")
    except Exception as e:
        results["steps"].append(f"allDrives exception: {e}")

    return jsonify(results)


@app.route("/api/organize")
def organize_endpoint():
    """Trigger loss run organization from the web process (bypasses bot subprocess)."""
    try:
        from loss_run_organizer import organize_loss_runs
        logger.info("[ORGANIZE-WEB] Starting organize_loss_runs from web endpoint")
        results = organize_loss_runs()
        logger.info(f"[ORGANIZE-WEB] Results: {json.dumps(results, default=str)[:2000]}")
        return jsonify(results)
    except Exception as e:
        logger.error(f"[ORGANIZE-WEB] Exception: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500


@app.route("/api/session", methods=["POST"])
def create_session():
    """Create a new proposal session."""
    session_id = str(uuid.uuid4())[:8]
    client_name = request.json.get("client_name", "")
    if not client_name:
        return jsonify({"error": "Client name is required"}), 400

    session_dir = os.path.join(UPLOAD_DIR, session_id)
    os.makedirs(session_dir, exist_ok=True)

    session_data = {
        "client_name": client_name,
        "session_dir": session_dir,
        "files": [],
        "extracted_data": None,
        "sov_data": None,
        "status": "created",
        "created_at": datetime.now().isoformat(),
    }
    _set_session(session_id, session_data)

    return jsonify({"session_id": session_id, "client_name": client_name})


@app.route("/api/upload/<session_id>", methods=["POST"])
def upload_files(session_id):
    """Upload files (PDFs and Excel SOVs) to a session."""
    session = _get_session(session_id)
    if not session:
        return jsonify({"error": "Session not found"}), 404

    if "files" not in request.files:
        return jsonify({"error": "No files uploaded"}), 400

    uploaded = []
    for f in request.files.getlist("files"):
        if not f.filename:
            continue
        ext = Path(f.filename).suffix.lower()
        if ext not in (".pdf", ".xlsx", ".xls", ".csv", ".jpg", ".jpeg", ".png"):
            continue

        safe_name = f"{uuid.uuid4().hex[:8]}_{f.filename}"
        save_path = os.path.join(session["session_dir"], safe_name)
        f.save(save_path)

        if ext == ".pdf":
            file_type = "pdf"
        elif ext in (".jpg", ".jpeg", ".png"):
            file_type = "image"
        else:
            file_type = "excel"
        file_info = {
            "filename": f.filename,
            "path": save_path,
            "type": file_type,
            "size": os.path.getsize(save_path),
        }

        # Check if Excel file is an SOV
        if file_type == "excel" and ext == ".xlsx":
            try:
                if is_sov_file(save_path):
                    file_info["is_sov"] = True
            except Exception:
                file_info["is_sov"] = False

        session["files"].append(file_info)
        uploaded.append({
            "filename": f.filename,
            "type": file_type,
            "is_sov": file_info.get("is_sov", False),
            "size": file_info["size"],
        })

    _set_session(session_id, session)
    return jsonify({"uploaded": uploaded, "total_files": len(session["files"])})


def _run_extraction(session_id):
    """Background worker: extract and structure data from all uploaded files."""
    session = _get_session(session_id)
    if not session:
        return

    extractor = ProposalExtractor()

    try:
        # Process SOV files first
        for file_info in session["files"]:
            if file_info.get("is_sov"):
                logger.info(f"Parsing SOV: {file_info['filename']}")
                sov_data = parse_sov(file_info["path"])
                if "error" not in sov_data:
                    sov_data = aggregate_locations(sov_data)
                    session["sov_data"] = sov_data
                    logger.info(f"SOV parsed: {len(sov_data.get('locations', []))} locations")

        # Step 1: Extract text from ALL files first (fast, no GPT calls)
        all_pdf_texts = []
        all_excel_data = []
        total_files = len([f for f in session["files"] if not f.get("is_sov")])
        processed = 0
        for file_info in session["files"]:
            if file_info["type"] == "pdf":
                logger.info(f"Extracting text from PDF: {file_info['filename']}")
                text = extractor.extract_pdf_text(file_info["path"])
                if text:
                    all_pdf_texts.append({"filename": file_info["filename"], "text": text})
                    logger.info(f"  -> {len(text)} chars extracted")
                    processed += 1
                    session["extract_progress"] = f"Extracted {processed}/{total_files} files"
                    _set_session(session_id, session)
                else:
                    logger.warning(f"No text from {file_info['filename']}")
                    processed += 1
            elif file_info["type"] == "image":
                # Convert image to text via GPT Vision
                logger.info(f"Processing image file: {file_info['filename']}")
                try:
                    import base64
                    with open(file_info["path"], "rb") as img_f:
                        img_b64 = base64.b64encode(img_f.read()).decode("utf-8")
                    ext = Path(file_info["filename"]).suffix.lower()
                    mime = "image/jpeg" if ext in (".jpg", ".jpeg") else "image/png"
                    from openai import OpenAI
                    client = OpenAI()
                    resp = client.chat.completions.create(
                        model="gpt-4.1-mini",
                        messages=[{"role": "user", "content": [
                            {"type": "text", "text": "Extract ALL text from this insurance document image. Preserve structure, numbers, and formatting."},
                            {"type": "image_url", "image_url": {"url": f"data:{mime};base64,{img_b64}", "detail": "high"}}
                        ]}],
                        max_tokens=4000
                    )
                    img_text = resp.choices[0].message.content
                    if img_text:
                        all_pdf_texts.append({"filename": file_info["filename"], "text": img_text})
                        logger.info(f"  -> {len(img_text)} chars extracted from image")
                except Exception as e:
                    logger.warning(f"Image extraction failed for {file_info['filename']}: {e}")
            elif file_info["type"] == "excel" and not file_info.get("is_sov"):
                logger.info(f"Extracting Excel: {file_info['filename']}")
                data = extractor.extract_excel_data(file_info["path"])
                if data:
                    all_excel_data.append({"filename": file_info["filename"], "data": data})

        if not all_pdf_texts and not all_excel_data:
            session["status"] = "error"
            session["extract_error"] = "No data could be extracted from the uploaded files"
            _set_session(session_id, session)
            return

        # Step 2: Single combined GPT call for ALL files at once
        session["extract_progress"] = "Analyzing with AI... this may take 1-3 minutes"
        _set_session(session_id, session)
        logger.info(f"Sending {len(all_pdf_texts)} PDFs + {len(all_excel_data)} Excel files to GPT in single call")
        merged_data = extractor.structure_insurance_data(
            all_pdf_texts, all_excel_data, session["client_name"]
        )

        if not merged_data or "error" in merged_data:
            session["status"] = "error"
            session["extract_error"] = (merged_data.get("error", "Unknown extraction error")
                                        if merged_data else "No data extracted")
            _set_session(session_id, session)
            return

        _normalize_coverages(merged_data)

        # Enrich with SOV data
        has_sov = session.get("sov_data") is not None
        logger.info(f"SOV enrichment check: has_sov={has_sov}, "
                    f"sov_locations={len(session['sov_data'].get('locations', [])) if has_sov else 0}, "
                    f"gpt_locations={len(merged_data.get('locations', []))}")
        if session.get("sov_data"):
            _enrich_with_sov(merged_data, session["sov_data"])
            logger.info(f"After enrichment: {len(merged_data.get('locations', []))} locations, "
                        f"first loc name='{merged_data.get('locations', [{}])[0].get('name', '')}', "
                        f"first loc tiv={merged_data.get('locations', [{}])[0].get('tiv', 0)}")

        session["extracted_data"] = merged_data
        session["status"] = "extracted"
        _set_session(session_id, session)
        logger.info(f"Extraction complete for session {session_id}")

    except Exception as e:
        logger.exception(f"Extraction failed: {e}")
        session["status"] = "error"
        session["extract_error"] = f"Extraction failed: {str(e)}"
        _set_session(session_id, session)


@app.route("/api/extract/<session_id>", methods=["POST"])
def extract_data(session_id):
    """Start extraction in a background thread and return immediately.
    
    The frontend polls GET /api/extract/<session_id>/status for completion.
    This avoids Railway's proxy timeout (~100s) for long-running extractions.
    """
    session = _get_session(session_id)
    if not session:
        return jsonify({"error": "Session not found"}), 404

    if not session["files"]:
        return jsonify({"error": "No files uploaded yet"}), 400

    if session.get("status") == "extracting":
        return jsonify({"status": "extracting", "message": "Extraction already in progress"})

    session["status"] = "extracting"
    session["extract_error"] = None
    _set_session(session_id, session)

    # Start extraction in background thread
    thread = threading.Thread(target=_run_extraction, args=(session_id,), daemon=True)
    thread.start()
    logger.info(f"Started background extraction for session {session_id}")

    return jsonify({"status": "extracting", "message": "Extraction started"})


@app.route("/api/extract/<session_id>/status", methods=["GET"])
def extract_status(session_id):
    """Poll extraction status. Returns data when complete."""
    session = _get_session(session_id)
    if not session:
        return jsonify({"error": "Session not found"}), 404

    status = session.get("status", "unknown")

    if status == "extracting":
        progress = session.get("extract_progress", "Processing files...")
        return jsonify({"status": "extracting", "progress": progress})

    if status == "error":
        error_msg = session.get("extract_error", "Unknown error")
        return jsonify({"status": "error", "error": error_msg}), 400

    if status == "extracted":
        merged_data = session.get("extracted_data", {})
        summary = _build_review_summary(merged_data)
        return jsonify({
            "status": "extracted",
            "summary": summary,
            "data": merged_data,
        })

    return jsonify({"status": status})


@app.route("/api/update/<session_id>", methods=["POST"])
def update_data(session_id):
    """Update the extracted data with user edits."""
    session = _get_session(session_id)
    if not session:
        return jsonify({"error": "Session not found"}), 404

    updated_data = request.json.get("data")
    if not updated_data:
        return jsonify({"error": "No data provided"}), 400

    # Preserve sov_data from the session (it's large and not edited in the UI)
    if session.get("sov_data") and "sov_data" not in updated_data:
        existing_data = session.get("extracted_data", {})
        if existing_data:
            updated_data["sov_data"] = existing_data.get("sov_data")

    session["extracted_data"] = updated_data
    session["status"] = "reviewed"
    _set_session(session_id, session)

    return jsonify({"status": "updated"})


@app.route("/api/generate/<session_id>", methods=["POST"])
def generate_doc(session_id):
    """Generate the DOCX proposal from the (possibly edited) extracted data."""
    session = _get_session(session_id)
    if not session:
        return jsonify({"error": "Session not found"}), 404

    if not session.get("extracted_data"):
        return jsonify({"error": "No extracted data available. Run extraction first."}), 400

    session["status"] = "generating"
    _set_session(session_id, session)

    try:
        client_name = session["client_name"].replace(" ", "_").replace("/", "-")
        timestamp = datetime.now().strftime("%Y%m%d")
        docx_filename = f"Proposal_{client_name}_{timestamp}.docx"
        docx_path = os.path.join(session["session_dir"], docx_filename)

        # Map web UI expiring premiums into the format the generator expects
        gen_data = session["extracted_data"]
        if gen_data.get("expiring_premiums_data"):
            gen_data["expiring_premiums"] = gen_data["expiring_premiums_data"]
            logger.info(f"Mapped expiring premiums: {gen_data['expiring_premiums']}")
        
        # Normalize workers_compensation -> workers_comp for generator compatibility
        covs = gen_data.get("coverages", {})
        if "workers_compensation" in covs and "workers_comp" not in covs:
            covs["workers_comp"] = covs.pop("workers_compensation")
        # Also normalize in expiring_premiums
        ep = gen_data.get("expiring_premiums", {})
        if isinstance(ep, dict) and "workers_compensation" in ep and "workers_comp" not in ep:
            ep["workers_comp"] = ep.pop("workers_compensation")
        
        # Normalize excess_liability / excess variants -> umbrella layer keys
        # The GPT extractor may return excess layers with various key names
        _excess_aliases = ["excess_liability", "excess", "excess_layer_2", "excess_layer_3",
                           "2nd_excess", "second_excess", "3rd_excess", "third_excess"]
        for alias in _excess_aliases:
            if alias in covs:
                # Map to the next available umbrella layer slot
                if "umbrella" not in covs:
                    covs["umbrella"] = covs.pop(alias)
                    logger.info(f"Normalized {alias} -> umbrella")
                elif "umbrella_layer_2" not in covs:
                    covs["umbrella_layer_2"] = covs.pop(alias)
                    logger.info(f"Normalized {alias} -> umbrella_layer_2")
                elif "umbrella_layer_3" not in covs:
                    covs["umbrella_layer_3"] = covs.pop(alias)
                    logger.info(f"Normalized {alias} -> umbrella_layer_3")
                else:
                    logger.warning(f"All umbrella slots full, cannot map {alias}")

        generate_proposal(gen_data, docx_path)

        session["status"] = "complete"
        session["docx_path"] = docx_path
        session["docx_filename"] = docx_filename
        _set_session(session_id, session)

        # Return the DOCX as base64 inline so download works even if session expires
        import base64
        with open(docx_path, "rb") as df:
            docx_b64 = base64.b64encode(df.read()).decode("utf-8")

        return jsonify({
            "status": "complete",
            "filename": docx_filename,
            "download_url": f"/api/download/{session_id}",
            "docx_base64": docx_b64,
        })

    except Exception as e:
        logger.exception(f"Generation failed: {e}")
        session["status"] = "error"
        _set_session(session_id, session)
        return jsonify({"error": f"Document generation failed: {str(e)}"}), 500


@app.route("/api/download/<session_id>", methods=["GET"])
def download_doc(session_id):
    """Download the generated DOCX file."""
    session = _get_session(session_id)
    if not session:
        return jsonify({"error": "Session not found"}), 404

    if not session.get("docx_path") or not os.path.exists(session.get("docx_path", "")):
        return jsonify({"error": "No document available for download"}), 404

    return send_file(
        session["docx_path"],
        as_attachment=True,
        download_name=session.get("docx_filename", "proposal.docx"),
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.route("/api/session/<session_id>", methods=["GET"])
def get_session(session_id):
    """Get session status and data."""
    session = _get_session(session_id)
    if not session:
        return jsonify({"error": "Session not found"}), 404

    return jsonify({
        "session_id": session_id,
        "client_name": session["client_name"],
        "status": session["status"],
        "files": [{"filename": f["filename"], "type": f["type"], "is_sov": f.get("is_sov", False)} for f in session["files"]],
        "has_data": session.get("extracted_data") is not None,
    })


def _build_review_summary(data):
    """Build a human-readable summary of extracted data for the review step."""
    summary = {}

    # Client info
    ci = data.get("client_info", {})
    summary["client_info"] = {
        "named_insured": ci.get("named_insured", ""),
        "dba": ci.get("dba", ""),
        "address": ci.get("address", ""),
        "entity_type": ci.get("entity_type", ""),
    }

    # Coverages found
    coverages = data.get("coverages", {})
    cov_summary = {}
    coverage_display = {
        "property": "Property",
        "excess_property": "Excess Property Layer 1",
        "excess_property_2": "Excess Property Layer 2",
        "general_liability": "General Liability",
        "umbrella": "Umbrella/Excess Layer 1",
        "umbrella_layer_2": "Excess Layer 2",
        "umbrella_layer_3": "Excess Layer 3",
        "workers_compensation": "Workers Compensation",
        "commercial_auto": "Commercial Auto",
        "crime": "Crime/Fidelity",
        "epli": "EPLI",
        "cyber": "Cyber",
        "flood": "Flood",
        "earthquake": "Earthquake",
    }
    for key, cov in coverages.items():
        if isinstance(cov, dict) and cov.get("carrier"):
            cov_summary[key] = {
                "display_name": coverage_display.get(key, key.replace("_", " ").title()),
                "carrier": cov.get("carrier", ""),
                "premium": cov.get("premium", 0),
                "total_premium": cov.get("total_premium", 0),
                "forms_count": len(cov.get("forms_endorsements", []) or []),
                "limits_count": len(cov.get("coverage_limits", []) or []),
            }
    summary["coverages"] = cov_summary

    # Locations count
    summary["locations_count"] = len(data.get("locations", []))
    summary["named_insureds_count"] = len(data.get("named_insureds", []))

    # SOV data
    sov = data.get("sov_data", {})
    if sov:
        summary["sov"] = {
            "locations": len(sov.get("locations", [])),
            "total_tiv": sov.get("totals", {}).get("tiv", 0),
        }

    return summary


# ─── Main ───

if __name__ == "__main__":
    port = int(os.environ.get("WEB_PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
