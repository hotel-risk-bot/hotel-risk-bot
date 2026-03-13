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
from extraction_validator import validate_extraction
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
# Use a STABLE directory path so sessions survive Railway redeploys within the same container
_default_upload_dir = os.path.join(tempfile.gettempdir(), "proposal_web_uploads")
UPLOAD_DIR = os.environ.get("UPLOAD_DIR", _default_upload_dir)
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
                if cov_type not in normalized:
                    normalized[cov_type] = item
                else:
                    # Competing quote — find next available alt slot
                    for alt_n in range(1, 5):
                        alt_key = f"{cov_type}_alt_{alt_n}"
                        if alt_key not in normalized:
                            normalized[alt_key] = item
                            break
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

    # Merge named insureds (with fuzzy deduplication)
    existing_named = merged.get("named_insureds", [])
    if not isinstance(existing_named, list):
        existing_named = [existing_named] if existing_named else []
    def _ni_name(ni):
        if isinstance(ni, dict):
            return ni.get("name", "").strip()
        return str(ni).strip()
    existing_name_strings = [_ni_name(ni) for ni in existing_named]
    for ni in new_data.get("named_insureds", []):
        name_key = _ni_name(ni)
        if name_key and not _is_duplicate_named_insured(name_key, existing_name_strings):
            existing_named.append(ni)
            existing_name_strings.append(name_key)
    merged["named_insureds"] = existing_named

    # Merge additional_named_insureds from carrier quote into named_insureds
    # These come from the "Additional Named Insureds Schedule" on the carrier quote
    # and are the PRIMARY source for named insured data
    additional_named = new_data.get("additional_named_insureds", [])
    if not additional_named:
        additional_named = merged.get("additional_named_insureds", [])
    if additional_named:
        logger.info(f"Merging {len(additional_named)} additional_named_insureds from carrier quote")
        for ani in additional_named:
            ani_name = _ni_name(ani)
            if ani_name and not _is_duplicate_named_insured(ani_name, existing_name_strings):
                # Mark as from carrier quote (primary source)
                if isinstance(ani, dict):
                    ani["relationship"] = ani.get("relationship", "Additional Named Insured")
                existing_named.append(ani)
                existing_name_strings.append(ani_name)
                logger.info(f"  Added additional named insured from carrier quote: {ani_name}")
        merged["named_insureds"] = existing_named
        # Also preserve the raw list for reference
        merged["additional_named_insureds"] = additional_named

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


def _normalize_entity_name(name):
    """Normalize entity name for deduplication.
    Strips punctuation, normalizes whitespace, lowercases.
    'OM Belleville, LLC' and 'OM Belleville LLC' -> 'om belleville llc'
    'Westhampton Hospitality LLC' and 'Westampton Hospitality LLC' -> handled by fuzzy match
    """
    if not name:
        return ""
    import re
    n = name.strip().lower()
    # Remove all punctuation except spaces
    n = re.sub(r'[^a-z0-9\s]', '', n)
    # Normalize whitespace
    n = re.sub(r'\s+', ' ', n).strip()
    return n


def _entity_names_match(name1, name2):
    """Check if two entity names match, allowing for minor spelling differences."""
    n1 = _normalize_entity_name(name1)
    n2 = _normalize_entity_name(name2)
    if not n1 or not n2:
        return False
    # Exact match after normalization
    if n1 == n2:
        return True
    # Check if one contains the other (for partial matches like 'OM Belleville' vs 'OM Belleville LLC')
    if n1 in n2 or n2 in n1:
        return True
    # Fuzzy match: allow 1-2 character differences (handles Westhampton vs Westampton)
    if abs(len(n1) - len(n2)) <= 2:
        # Simple edit distance check - count differing characters
        shorter, longer = (n1, n2) if len(n1) <= len(n2) else (n2, n1)
        if len(longer) > 5:  # Only for names long enough
            # Check character-by-character with shift tolerance
            diffs = 0
            j = 0
            for i in range(len(shorter)):
                if j < len(longer) and shorter[i] == longer[j]:
                    j += 1
                else:
                    diffs += 1
                    # Try skipping a character in the longer string
                    if j + 1 < len(longer) and shorter[i] == longer[j + 1]:
                        j += 2
                    else:
                        j += 1
            diffs += len(longer) - j
            if diffs <= 2:
                return True
    return False


def _is_duplicate_named_insured(new_name, existing_names_list):
    """Check if a named insured already exists in the list (fuzzy match)."""
    for existing in existing_names_list:
        if _entity_names_match(new_name, existing):
            return True
    return False


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
        import re
        sov_lookup = {}
        sov_addr_lookup = {}  # address-based lookup for matching GL locations to SOV
        for loc in sov_locs:
            loc_num = loc.get("location_num", loc.get("building_num", 0))
            if loc_num:
                sov_lookup[str(loc_num)] = loc
            # Build address-based lookup (normalize for fuzzy matching)
            addr = (loc.get("address", "") or "").strip().lower()
            if addr:
                # Normalize common abbreviations
                addr_norm = re.sub(r'\b(road|rd|street|st|avenue|ave|drive|dr|boulevard|blvd|highway|hwy|lane|ln|pike|pk)\b', '', addr)
                addr_norm = re.sub(r'[^a-z0-9]', '', addr_norm)
                sov_addr_lookup[addr_norm] = loc

        classes = gl_cov.get("schedule_of_classes", [])
        if classes:
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

        # PRIORITY: Merge additional_named_insureds from GPT extraction (carrier quote)
        # These come from the "Additional Named Insureds Schedule" on the carrier quote
        # and are the PRIMARY source — they should already be in named_insureds from _merge_extraction_results
        # but handle the case where they weren't merged yet (e.g., single-file extraction)
        additional_named = extracted_data.get("additional_named_insureds", [])
        if additional_named:
            existing_named_list = extracted_data.get("named_insureds", [])
            existing_named_keys = {(ni.get("name", "") if isinstance(ni, dict) else str(ni)).strip().lower()
                                   for ni in existing_named_list}
            added_count = 0
            for ani in additional_named:
                ani_name = (ani.get("name", "") if isinstance(ani, dict) else str(ani)).strip()
                if ani_name and ani_name.lower() not in existing_named_keys:
                    if isinstance(ani, dict):
                        ani["relationship"] = ani.get("relationship", "Additional Named Insured")
                    existing_named_list.append(ani)
                    existing_named_keys.add(ani_name.lower())
                    added_count += 1
            if added_count:
                extracted_data["named_insureds"] = existing_named_list
                logger.info(f"  Added {added_count} additional named insureds from carrier quote (primary source)")

        # SECONDARY: Enrich named_insureds from SOV corporate entities for GL-covered locations
        # ONLY pull named insureds for locations that appear on the liability quote
        gl_location_nums = set()
        gl_addresses_norm = set()
        # Collect GL location identifiers from schedule_of_classes and designated_premises
        for cls_entry in gl_cov.get("schedule_of_classes", []):
            if isinstance(cls_entry, dict):
                loc_str = str(cls_entry.get("location", ""))
                loc_match = re.search(r"(\d+)", loc_str)
                if loc_match:
                    gl_location_nums.add(loc_match.group(1))
                # Also track by address
                cls_addr = (cls_entry.get("address", "") or "").strip().lower()
                if cls_addr:
                    cls_addr_norm = re.sub(r'\b(road|rd|street|st|avenue|ave|drive|dr|boulevard|blvd|highway|hwy|lane|ln|pike|pk)\b', '', cls_addr)
                    cls_addr_norm = re.sub(r'[^a-z0-9]', '', cls_addr_norm)
                    gl_addresses_norm.add(cls_addr_norm)
        for dp in gl_cov.get("designated_premises", []):
            dp_str = str(dp).strip().lower()
            dp_norm = re.sub(r'\b(road|rd|street|st|avenue|ave|drive|dr|boulevard|blvd|highway|hwy|lane|ln|pike|pk)\b', '', dp_str)
            dp_norm = re.sub(r'[^a-z0-9]', '', dp_norm)
            gl_addresses_norm.add(dp_norm)

        # Match GL locations to SOV locations and collect corporate entities
        # Use fuzzy deduplication to avoid duplicates from carrier quote vs SOV
        existing_named_names = [ni.get("name", "").strip() if isinstance(ni, dict) else str(ni).strip()
                               for ni in extracted_data.get("named_insureds", [])]
        new_named_insureds = []
        seen_entities = []  # Use list for fuzzy matching

        for loc in sov_locs:
            loc_num = str(loc.get("location_num", loc.get("building_num", 0)))
            loc_addr = (loc.get("address", "") or "").strip().lower()
            loc_addr_norm = re.sub(r'\b(road|rd|street|st|avenue|ave|drive|dr|boulevard|blvd|highway|hwy|lane|ln|pike|pk)\b', '', loc_addr)
            loc_addr_norm = re.sub(r'[^a-z0-9]', '', loc_addr_norm)

            # Check if this SOV location is on the GL quote
            is_gl_location = (loc_num in gl_location_nums) or (loc_addr_norm in gl_addresses_norm)
            if not is_gl_location:
                continue

            # Get corporate entity name from SOV
            corp_name = (loc.get("corporate_name") or loc.get("client_name") or "").strip()
            dba_name = (loc.get("dba") or loc.get("hotel_flag") or "").strip()

            if corp_name and not _is_duplicate_named_insured(corp_name, existing_named_names) \
               and not _is_duplicate_named_insured(corp_name, seen_entities):
                new_named_insureds.append({
                    "name": corp_name,
                    "dba": dba_name,
                    "relationship": "Named Insured (from SOV)"
                })
                seen_entities.append(corp_name)
                logger.info(f"  SOV named insured enrichment: '{corp_name}' (DBA: '{dba_name}') "
                           f"from GL location {loc_num} / {loc.get('address', '')}")
            elif corp_name:
                logger.info(f"  SOV named insured SKIPPED (duplicate): '{corp_name}' already exists in named insureds")

        if new_named_insureds:
            existing_list = extracted_data.get("named_insureds", [])
            extracted_data["named_insureds"] = existing_list + new_named_insureds
            logger.info(f"  Added {len(new_named_insureds)} named insureds from SOV for GL locations "
                       f"(total now: {len(extracted_data['named_insureds'])})")
        else:
            logger.info("  No new named insureds to add from SOV (none matched GL locations or all already present)")


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
    client_loss_runs_id = os.environ.get("CLIENT_LOSS_RUNS_FOLDER_ID", "").strip()
    accounts_id = os.environ.get("ACCOUNTS_FOLDER_ID", "").strip()
    results["inbox_folder_id"] = inbox_id
    results["client_loss_runs_folder_id"] = client_loss_runs_id or "NOT SET"
    results["accounts_folder_id"] = accounts_id or "NOT SET"
    results["dest_folder_used"] = client_loss_runs_id if client_loss_runs_id else (accounts_id if accounts_id else "NONE")
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


@app.route("/api/cleanup-folders")
def cleanup_folders_endpoint():
    """Merge duplicate client folders and delete empty shells."""
    try:
        from loss_run_organizer import _get_access_token
        import requests as http_req

        CLIENT_LOSS_RUNS_FOLDER_ID = os.environ.get(
            "CLIENT_LOSS_RUNS_FOLDER_ID", "1v2-Y9pKIY4_Jh3X2_ZJOCB7XdNLptS8x"
        )

        # Merge rules: aliases -> canonical
        MERGE_GROUPS = [
            {
                "canonical": "Dalwadi Hospitality Management, LLC",
                "aliases": ["DALWADI HOSPITALITY"],
            },
            {
                "canonical": "Kautilya Management LLC",
                "aliases": [
                    "KAUTILYA MANAGMENT LLC",
                    "Kautilya Columbus Hotel LLC",
                    "Kautilya Hotel Group, LLC",
                    "Kautilya Management LLC; Vinit",
                ],
            },
            {
                "canonical": "Pride Management, Inc.",
                "aliases": ["Pride Management Inc."],
            },
            {
                "canonical": "Star Owner LLC dba Hyatt Place",
                "aliases": ["Star Owner, LLC"],
            },
            {
                "canonical": "Sandalwood Hotel LLC dba America's Best Value Inn",
                "aliases": ["Sandalwood Hotel LLC Company"],
            },
        ]

        token = _get_access_token()
        if not token:
            return jsonify({"error": "Could not get access token"}), 500
        hdrs = {"Authorization": f"Bearer {token}"}

        def _list(folder_id, mime_type=None):
            q = f"'{folder_id}' in parents and trashed = false"
            if mime_type:
                q += f" and mimeType = '{mime_type}'"
            items = []
            pt = None
            while True:
                p = {"q": q, "fields": "nextPageToken,files(id,name,mimeType,parents)", "pageSize": 200}
                if pt:
                    p["pageToken"] = pt
                r = http_req.get("https://www.googleapis.com/drive/v3/files", headers=hdrs, params=p)
                r.raise_for_status()
                d = r.json()
                items.extend(d.get("files", []))
                pt = d.get("nextPageToken")
                if not pt:
                    break
            return items

        def _move(file_id, old_parent, new_parent):
            r = http_req.patch(
                f"https://www.googleapis.com/drive/v3/files/{file_id}",
                headers=hdrs,
                params={"addParents": new_parent, "removeParents": old_parent},
            )
            r.raise_for_status()

        def _find_or_create(name, parent_id):
            q = f"'{parent_id}' in parents and name = '{name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
            r = http_req.get("https://www.googleapis.com/drive/v3/files", headers=hdrs, params={"q": q, "fields": "files(id,name)"})
            r.raise_for_status()
            files = r.json().get("files", [])
            if files:
                return files[0]["id"]
            r = http_req.post("https://www.googleapis.com/drive/v3/files", headers=hdrs,
                            json={"name": name, "mimeType": "application/vnd.google-apps.folder", "parents": [parent_id]})
            r.raise_for_status()
            return r.json()["id"]

        def _move_recursive(src_id, dst_id):
            moved = 0
            children = _list(src_id)
            for ch in children:
                if ch["mimeType"] == "application/vnd.google-apps.folder":
                    sub_dst = _find_or_create(ch["name"], dst_id)
                    moved += _move_recursive(ch["id"], sub_dst)
                else:
                    _move(ch["id"], src_id, dst_id)
                    moved += 1
            return moved

        def _delete_recursive(folder_id):
            children = _list(folder_id)
            for ch in children:
                if ch["mimeType"] == "application/vnd.google-apps.folder":
                    _delete_recursive(ch["id"])
                else:
                    http_req.delete(f"https://www.googleapis.com/drive/v3/files/{ch['id']}", headers=hdrs)
            http_req.delete(f"https://www.googleapis.com/drive/v3/files/{folder_id}", headers=hdrs)

        # Get all folders
        folders = _list(CLIENT_LOSS_RUNS_FOLDER_ID, mime_type="application/vnd.google-apps.folder")
        folder_map = {f["name"]: f["id"] for f in folders}
        results = {"merged": [], "deleted": [], "errors": [], "final_folders": []}

        for group in MERGE_GROUPS:
            canonical_name = group["canonical"]
            canonical_id = folder_map.get(canonical_name)
            if not canonical_id:
                results["errors"].append(f"Canonical folder '{canonical_name}' not found")
                continue
            for alias in group["aliases"]:
                alias_id = folder_map.get(alias)
                if not alias_id:
                    continue
                try:
                    moved = _move_recursive(alias_id, canonical_id)
                    results["merged"].append({"from": alias, "to": canonical_name, "files_moved": moved})
                    _delete_recursive(alias_id)
                    results["deleted"].append(alias)
                except Exception as e:
                    results["errors"].append(f"Error merging '{alias}': {str(e)}")

        # Get final state
        final = _list(CLIENT_LOSS_RUNS_FOLDER_ID, mime_type="application/vnd.google-apps.folder")
        results["final_folders"] = sorted([f["name"] for f in final])
        results["total_folders"] = len(final)

        return jsonify(results)
    except Exception as e:
        logger.error(f"[CLEANUP] Exception: {e}", exc_info=True)
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
        file_info["is_sov"] = False
        if file_type == "excel" and ext == ".xlsx":
            try:
                sov_check = is_sov_file(save_path)
                logger.info(f"SOV check for {f.filename}: {sov_check}")
                file_info["is_sov"] = sov_check
            except Exception as e:
                logger.warning(f"SOV check failed for {f.filename}: {e}")

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
            # If only SOV was uploaded (no quote PDFs), create basic data from SOV
            if session.get("sov_data"):
                logger.info("SOV-only upload: creating basic data from SOV without GPT")
                merged_data = {
                    "client_info": {"named_insured": session["client_name"]},
                    "coverages": {},
                    "locations": [],
                    "named_insureds": [],
                    "client_name": session["client_name"],
                }
            else:
                session["status"] = "error"
                session["extract_error"] = "No data could be extracted from the uploaded files"
                _set_session(session_id, session)
                return
        else:
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
        
        # SAFETY: Ensure sov_data is always available for the generator
        # The web UI doesn't send sov_data back, so it must be preserved from the session
        if not gen_data.get("sov_data") and session.get("sov_data"):
            gen_data["sov_data"] = session["sov_data"]
            logger.info(f"Injected sov_data from session: {len(session['sov_data'].get('locations', []))} locations")
        elif gen_data.get("sov_data"):
            logger.info(f"sov_data already in gen_data: {len(gen_data['sov_data'].get('locations', []))} locations")
        else:
            logger.warning("No sov_data available for proposal generation")
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

        # Run validation checks before generating
        validation = validate_extraction(gen_data, session.get("sov_data"))
        if validation["corrections"]:
            logger.info(f"Auto-corrections applied: {validation['corrections']}")
        if validation["warnings"]:
            logger.warning(f"Validation warnings: {validation['warnings']}")
        if validation["errors"]:
            logger.error(f"Validation errors: {validation['errors']}")
        
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
        "enviro_pack": "Enviro Pack",
        "terrorism": "Terrorism / TRIA",
        "equipment_breakdown": "Equipment Breakdown",
        "liquor_liability": "Liquor Liability",
        "innkeepers_liability": "Innkeepers Liability",
        "environmental": "Environmental / Pollution",
        "workplace_violence": "Workplace Violence",
        "pollution": "Pollution Liability",
        "abuse_molestation": "Sexual Abuse & Molestation",
        "active_assailant": "Active Assailant",
        "inland_marine": "Inland Marine",
        "garage_keepers": "Garage Keepers",
        "wind_deductible_buydown": "Wind Deductible Buy Down",
        "deductible_buydown": "Deductible Buy Down",
        "property_alt_1": "Property (Alt Quote 1)",
        "property_alt_2": "Property (Alt Quote 2)",
        "general_liability_alt_1": "General Liability (Alt Quote 1)",
        "general_liability_alt_2": "General Liability (Alt Quote 2)",
        "umbrella_alt_1": "Umbrella / Excess 2",
        "umbrella_alt_2": "Umbrella / Excess 3",
        "umbrella_alt_3": "Umbrella / Excess 4",
        "workers_compensation_alt_1": "Workers Comp (Alt Quote 1)",
        "cyber_alt_1": "Cyber (Alt Quote 1)",
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
