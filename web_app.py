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
from pathlib import Path
from datetime import datetime

from flask import Flask, request, jsonify, send_file, render_template

from proposal_extractor import ProposalExtractor, extract_text_from_pdf_smart, extract_text_from_excel
from proposal_generator import generate_proposal
from sov_parser import parse_sov, is_sov_file, aggregate_locations

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

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

    # Merge coverages with auto-promotion
    existing_covs = merged.get("coverages", {})
    new_covs = new_data.get("coverages", {})
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
    merged["coverages"] = existing_covs

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
    if not sov_data or "error" in sov_data:
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
        if ext not in (".pdf", ".xlsx", ".xls", ".csv"):
            continue

        safe_name = f"{uuid.uuid4().hex[:8]}_{f.filename}"
        save_path = os.path.join(session["session_dir"], safe_name)
        f.save(save_path)

        file_type = "pdf" if ext == ".pdf" else "excel"
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


@app.route("/api/extract/<session_id>", methods=["POST"])
def extract_data(session_id):
    """Extract and structure data from all uploaded files."""
    session = _get_session(session_id)
    if not session:
        return jsonify({"error": "Session not found"}), 404

    if not session["files"]:
        return jsonify({"error": "No files uploaded yet"}), 400

    session["status"] = "extracting"
    extractor = ProposalExtractor()
    merged_data = None

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

        # Process each PDF individually, then merge
        for file_info in session["files"]:
            if file_info["type"] == "pdf":
                logger.info(f"Extracting PDF: {file_info['filename']}")
                text = extractor.extract_pdf_text(file_info["path"])
                if not text:
                    logger.warning(f"No text from {file_info['filename']}")
                    continue

                pdf_texts = [{"filename": file_info["filename"], "text": text}]
                file_data = extractor.structure_insurance_data(
                    pdf_texts, [], session["client_name"]
                )

                if "error" not in file_data:
                    _normalize_coverages(file_data)
                    merged_data = _merge_extraction_results(merged_data, file_data)
                    logger.info(f"Merged {file_info['filename']}: coverages={list(file_data.get('coverages', {}).keys())}")
                else:
                    logger.error(f"Extraction error for {file_info['filename']}: {file_data['error']}")

            elif file_info["type"] == "excel" and not file_info.get("is_sov"):
                logger.info(f"Extracting Excel: {file_info['filename']}")
                data = extractor.extract_excel_data(file_info["path"])
                excel_data = [{"filename": file_info["filename"], "data": data}]
                file_data = extractor.structure_insurance_data(
                    [], excel_data, session["client_name"]
                )
                if "error" not in file_data:
                    _normalize_coverages(file_data)
                    merged_data = _merge_extraction_results(merged_data, file_data)

        if not merged_data:
            session["status"] = "error"
            return jsonify({"error": "No data could be extracted from the uploaded files"}), 400

        # Enrich with SOV data
        if session.get("sov_data"):
            _enrich_with_sov(merged_data, session["sov_data"])

        session["extracted_data"] = merged_data
        session["status"] = "extracted"
        _set_session(session_id, session)

        # Build a summary for the UI
        summary = _build_review_summary(merged_data)

        return jsonify({
            "status": "extracted",
            "summary": summary,
            "data": merged_data,
        })

    except Exception as e:
        logger.exception(f"Extraction failed: {e}")
        session["status"] = "error"
        _set_session(session_id, session)
        return jsonify({"error": f"Extraction failed: {str(e)}"}), 500


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
        docx_filename = f"HUB_Proposal_{client_name}_{timestamp}.docx"
        docx_path = os.path.join(session["session_dir"], docx_filename)

        # Map web UI expiring premiums into the format the generator expects
        gen_data = session["extracted_data"]
        if gen_data.get("expiring_premiums_data"):
            gen_data["expiring_premiums"] = gen_data["expiring_premiums_data"]
            logger.info(f"Mapped expiring premiums: {gen_data['expiring_premiums']}")

        generate_proposal(gen_data, docx_path)

        session["status"] = "complete"
        session["docx_path"] = docx_path
        session["docx_filename"] = docx_filename
        _set_session(session_id, session)

        return jsonify({
            "status": "complete",
            "filename": docx_filename,
            "download_url": f"/api/download/{session_id}",
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
