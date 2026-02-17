#!/usr/bin/env python3
"""
Marketing Summary Generator.
Pulls policies from the Airtable Sales System (Policies table) for a given
client/opportunity and generates a detailed marketing status summary with
coverage-specific fields (TIV, rates, deductibles, payroll, etc.).

Status Order: Incumbent → Submit → Market → Declined → Quoted → Proposed → Bound → Blocked
Coverage Order: Property → Liability → Umbrella → Workers Comp → Auto → EPLI → Cyber → Flood → Other
"""

import os
import logging
from collections import defaultdict

import requests as http_requests

logger = logging.getLogger(__name__)

AIRTABLE_PAT = os.environ.get("AIRTABLE_PAT", "")
SALES_BASE_ID = "appnFKEzmdLbR4CHY"
POLICIES_TABLE_ID = "tbl8vZP2oHrinwVfd"
OPPORTUNITIES_TABLE_ID = "tblMKuUsG1cosdQPN"

# ── Status order and display ────────────────────────────────────────────

STATUS_ORDER = [
    "Incumbent",
    "Submit",
    "Market",
    "Declined",
    "Quoted",
    "Proposed",
    "Bound",
    "Blocked",
    "Lost",
]

STATUS_EMOJI = {
    "Incumbent": "📋",
    "Submit": "🔴",
    "Market": "🟢",
    "Declined": "❌",
    "Quoted": "💰",
    "Proposed": "📤",
    "Bound": "✅",
    "Blocked": "🚫",
    "Lost": "💔",
}

STATUS_DESCRIPTIONS = {
    "Incumbent": "Current/expiring carrier",
    "Submit": "EXPOSED — Needs submission to underwriter",
    "Market": "Cleared — out to market",
    "Declined": "Underwriter declined",
    "Quoted": "Carrier provided a quote",
    "Proposed": "Proposed to client",
    "Bound": "Client bound with this carrier",
    "Blocked": "Carrier taken by competitor",
    "Lost": "Lost the account",
}

# ── Coverage type canonical ordering ────────────────────────────────────

COVERAGE_ORDER = [
    "Property",
    "Liability",
    "Umbrella",
    "Workers Compensation",
    "Auto",
    "Employment Practices Liability",
    "Cyber",
    "Flood",
    "Equipment Breakdown",
    "Package",
]

COVERAGE_ABBR = {
    "Property": "PROP",
    "Liability": "GL",
    "Umbrella": "UMB",
    "Workers Compensation": "WC",
    "Auto": "AUTO",
    "Employment Practices Liability": "EPLI",
    "Cyber": "CYBER",
    "Flood": "FLOOD",
    "Equipment Breakdown": "EB",
    "Package": "PKG",
    "Automobile": "AUTO",
    "Other": "OTHER",
}


# ── Helpers ─────────────────────────────────────────────────────────────

def airtable_headers():
    return {
        "Authorization": f"Bearer {AIRTABLE_PAT}",
        "Content-Type": "application/json",
    }


def _safe_str(val, default="—"):
    """Convert Airtable value to display string."""
    if val is None:
        return default
    if isinstance(val, list):
        return ", ".join(str(v) for v in val if v is not None) or default
    s = str(val).strip()
    return s if s else default


def _safe_currency(val, default="—"):
    """Format a numeric value as currency."""
    if val is None:
        return default
    try:
        n = float(str(val).replace("$", "").replace(",", ""))
        if n == 0:
            return default
        return f"${n:,.0f}"
    except (ValueError, TypeError):
        return default


def _safe_rate(val, default="—"):
    """Format a rate value."""
    if val is None:
        return default
    try:
        n = float(val)
        if n == 0:
            return default
        return f"${n:.2f}" if n >= 1 else f"${n:.4f}"
    except (ValueError, TypeError):
        return default


def _safe_number(val, default="—"):
    """Format a number."""
    if val is None:
        return default
    try:
        n = float(str(val).replace(",", ""))
        if n == 0:
            return default
        if n == int(n):
            return f"{int(n):,}"
        return f"{n:,.2f}"
    except (ValueError, TypeError):
        return default


def _normalize_coverage_type(policy_type_raw):
    """Normalize policy type to a canonical coverage type string."""
    if isinstance(policy_type_raw, list):
        pt = policy_type_raw[0] if policy_type_raw else "Other"
    else:
        pt = str(policy_type_raw) if policy_type_raw else "Other"

    pt_lower = pt.lower().strip()

    if "property" in pt_lower:
        return "Property"
    elif "liability" in pt_lower and "employment" not in pt_lower:
        return "Liability"
    elif "umbrella" in pt_lower:
        return "Umbrella"
    elif "worker" in pt_lower or "comp" in pt_lower:
        return "Workers Compensation"
    elif "auto" in pt_lower:
        return "Auto"
    elif "employment" in pt_lower or "epli" in pt_lower:
        return "Employment Practices Liability"
    elif "cyber" in pt_lower:
        return "Cyber"
    elif "flood" in pt_lower:
        return "Flood"
    elif "equipment" in pt_lower or "breakdown" in pt_lower:
        return "Equipment Breakdown"
    elif "package" in pt_lower or "pkg" in pt_lower:
        return "Package"
    elif "automobile" in pt_lower:
        return "Auto"
    elif not pt or pt.lower() in ["other", "n/a", "", "none"]:
        return "Other"
    else:
        return pt


def _coverage_sort_key(coverage_type):
    """Return sort key for coverage type ordering."""
    try:
        return COVERAGE_ORDER.index(coverage_type)
    except ValueError:
        return len(COVERAGE_ORDER)


# ── Airtable API calls ─────────────────────────────────────────────────

def _sanitize_for_formula(text: str) -> str:
    """Sanitize text for use in Airtable formula strings."""
    return text.replace('"', '\\"')


def search_opportunity(client_name: str) -> list:
    """Search for opportunities matching a client name (partial match)."""
    safe_name = _sanitize_for_formula(client_name)
    formula = (
        f"OR("
        f"SEARCH(LOWER("{safe_name}"), LOWER({{Opportunity Name}})),"
        f"SEARCH(LOWER("{safe_name}"), LOWER({{Corporate Name}}))"
        f")"
    )

    url = f"https://api.airtable.com/v0/{SALES_BASE_ID}/{OPPORTUNITIES_TABLE_ID}"
    params = {
        "filterByFormula": formula,
        "sort[0][field]": "Effective Date",
        "sort[0][direction]": "desc",
        "pageSize": 20,
    }

    logger.info(f"Searching opportunities for: '{client_name}' with formula: {formula}")

    try:
        resp = http_requests.get(url, headers=airtable_headers(), params=params, timeout=30)
        resp.raise_for_status()
        records = resp.json().get("records", [])
        logger.info(f"Found {len(records)} opportunities for '{client_name}'")
        return records
    except Exception as e:
        logger.error(f"Error searching opportunities: {e}")
        return []


def fetch_policies_for_client(client_name: str) -> list:
    """Fetch all policies related to a client name by searching the Policies table."""
    safe_name = _sanitize_for_formula(client_name)
    formula = (
        f"OR("
        f"SEARCH(LOWER("{safe_name}"), LOWER({{Name}})),"
        f"SEARCH(LOWER("{safe_name}"), LOWER({{Companies}}))"
        f")"
    )
    logger.info(f"Searching policies for: '{client_name}' with formula: {formula}")

    url = f"https://api.airtable.com/v0/{SALES_BASE_ID}/{POLICIES_TABLE_ID}"
    all_records = []
    offset = None

    while True:
        params = {
            "filterByFormula": formula,
            "pageSize": 100,
        }
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
            logger.error(f"Error fetching policies: {e}")
            break

    return all_records


def fetch_policies_by_record_ids(policy_record_ids: list) -> list:
    """Fetch policies by their record IDs (from the Opportunity's Policies linked field)."""
    if not policy_record_ids:
        return []

    all_records = []
    batch_size = 20

    for i in range(0, len(policy_record_ids), batch_size):
        batch = policy_record_ids[i:i + batch_size]
        conditions = [f'RECORD_ID()="{rid}"' for rid in batch]
        formula = f"OR({','.join(conditions)})" if len(conditions) > 1 else conditions[0]

        url = f"https://api.airtable.com/v0/{SALES_BASE_ID}/{POLICIES_TABLE_ID}"
        offset = None

        while True:
            params = {
                "filterByFormula": formula,
                "pageSize": 100,
            }
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
                logger.error(f"Error fetching policies by record IDs: {e}")
                break

    return all_records


# ── Coverage-specific detail formatters ─────────────────────────────────

def _property_details(flds: dict) -> list:
    """Return detail lines for a Property policy."""
    details = []
    tiv = _safe_currency(flds.get("TIV"))
    if tiv != "—":
        details.append(f"    TIV: {tiv}")
    prop_rate = flds.get("Property Rate")
    if prop_rate:
        try:
            r = float(prop_rate)
            if r > 0:
                details.append(f"    Property Rate: ${r:.2f}" if r >= 1 else f"    Property Rate: ${r:.4f}")
        except (ValueError, TypeError):
            pass
    prop_limit = _safe_str(flds.get("Property Limit"))
    if prop_limit != "—":
        details.append(f"    Property Limit: {prop_limit}")
    aop = _safe_str(flds.get("AOP"))
    if aop != "—":
        details.append(f"    AOP Deductible: {aop}")
    wind_type = _safe_str(flds.get("Wind Type"))
    wind = _safe_str(flds.get("Wind"))
    if wind_type != "—" or wind != "—":
        details.append(f"    Wind: {wind} ({wind_type})" if wind_type != "—" else f"    Wind: {wind}")
    aow = _safe_str(flds.get("AOW"))
    if aow != "—":
        details.append(f"    AOW (All Other Wind): {aow}")
    water = _safe_str(flds.get("Water Damage"))
    if water != "—":
        details.append(f"    Water Damage: {water}")
    flood_limit = _safe_str(flds.get("Flood Limit"))
    if flood_limit != "—":
        details.append(f"    Flood Limit: {flood_limit}")
    locs = _safe_number(flds.get("# of Locs"))
    if locs != "—":
        details.append(f"    # of Locations: {locs}")
    return details


def _liability_details(flds: dict) -> list:
    """Return detail lines for a Liability (GL) policy."""
    details = []
    sales = _safe_currency(flds.get("Gross Sales"))
    if sales != "—":
        details.append(f"    Total Sales: {sales}")
    gl_rate = flds.get("GL Rate $")
    if gl_rate:
        try:
            r = float(gl_rate)
            if r > 0:
                details.append(f"    GL Rate: ${r:.2f}")
        except (ValueError, TypeError):
            pass
    gl_rate_u = flds.get("GL Rate (u)")
    if gl_rate_u:
        try:
            r = float(gl_rate_u)
            if r > 0:
                details.append(f"    GL Rate/Unit: ${r:.2f}")
        except (ValueError, TypeError):
            pass
    gl_ded = _safe_str(flds.get("GL Deductible"))
    if gl_ded != "—":
        details.append(f"    GL Deductible: {gl_ded}")
    units = _safe_number(flds.get("Units"))
    if units != "—":
        details.append(f"    # of Units: {units}")
    locs = _safe_number(flds.get("# of Locs"))
    if locs != "—":
        details.append(f"    # of Locations: {locs}")
    return details


def _umbrella_details(flds: dict) -> list:
    """Return detail lines for an Umbrella policy."""
    details = []
    umb_limit = _safe_str(flds.get("UMB Limit"))
    if umb_limit != "—":
        details.append(f"    Umbrella Limit: {umb_limit}")
    sales = _safe_currency(flds.get("Gross Sales"))
    if sales != "—":
        details.append(f"    Total Sales: {sales}")
    units = _safe_number(flds.get("Units"))
    if units != "—":
        details.append(f"    # of Units: {units}")
    locs = _safe_number(flds.get("# of Locs"))
    if locs != "—":
        details.append(f"    # of Locations: {locs}")
    return details


def _wc_details(flds: dict) -> list:
    """Return detail lines for a Workers Compensation policy."""
    details = []
    total_payroll = _safe_currency(flds.get("Total Payroll"))
    if total_payroll != "—":
        details.append(f"    Total Payroll: {total_payroll}")
    else:
        # Try to sum individual payroll fields
        payroll_fields = [
            "Payroll 9052", "Payroll 9053", "Payroll 9054",
            "Payroll 9055", "Payroll 9058", "Payroll 8810",
            "Payroll 8742", "Payroll 1",
        ]
        total = 0
        for pf in payroll_fields:
            val = flds.get(pf)
            if val:
                try:
                    total += float(str(val).replace("$", "").replace(",", ""))
                except (ValueError, TypeError):
                    pass
        if total > 0:
            details.append(f"    Total Payroll: ${total:,.0f}")

    exp_mod = flds.get("Exp Mod")
    if exp_mod is not None:
        try:
            em = float(exp_mod)
            if em > 0:
                details.append(f"    Exp Mod: {em:.2f}")
        except (ValueError, TypeError):
            pass
    return details


def _epli_details(flds: dict) -> list:
    """Return detail lines for an EPLI policy."""
    details = []
    units = _safe_number(flds.get("Units"))
    if units != "—":
        details.append(f"    # of Units: {units}")
    locs = _safe_number(flds.get("# of Locs"))
    if locs != "—":
        details.append(f"    # of Locations: {locs}")
    return details


def _auto_details(flds: dict) -> list:
    """Return detail lines for an Auto policy."""
    details = []
    units = _safe_number(flds.get("Units"))
    if units != "—":
        details.append(f"    # of Vehicles/Units: {units}")
    return details


COVERAGE_DETAIL_FUNCS = {
    "Property": _property_details,
    "Liability": _liability_details,
    "Umbrella": _umbrella_details,
    "Workers Compensation": _wc_details,
    "Employment Practices Liability": _epli_details,
    "Auto": _auto_details,
}


# ── Build the summary ──────────────────────────────────────────────────

def build_marketing_summary(policies: list, client_name: str = "",
                             opportunity_name: str = "") -> str:
    """Build a formatted marketing summary from policy records."""
    if not policies:
        return f"No policies found for {client_name or opportunity_name}."

    # Parse each policy into a structured dict
    parsed = []
    for rec in policies:
        flds = rec.get("fields", {})
        coverage_type = _normalize_coverage_type(flds.get("Policy Type"))

        insurance_co_raw = flds.get("Insurance Company", "")
        if isinstance(insurance_co_raw, list):
            insurance_co = ", ".join(str(ic) for ic in insurance_co_raw if ic) or "N/A"
        else:
            insurance_co = str(insurance_co_raw).strip() if insurance_co_raw else "N/A"

        premium = 0
        try:
            premium = float(str(flds.get("Base Premium", 0) or 0).replace("$", "").replace(",", ""))
        except (ValueError, TypeError):
            pass

        comments = flds.get("Comments", "")
        if isinstance(comments, str) and len(comments) > 150:
            comments = comments[:147] + "..."

        broker = _safe_str(flds.get("Broker ABBR"))
        carrier_abbr = _safe_str(flds.get("Carrier ABBR"))

        parsed.append({
            "name": flds.get("Name", "Unknown"),
            "status": flds.get("Status", "Unknown"),
            "coverage_type": coverage_type,
            "insurance_company": insurance_co,
            "carrier_abbr": carrier_abbr,
            "premium": premium,
            "comments": comments,
            "broker": broker,
            "fields": flds,  # Keep raw fields for coverage-specific details
        })

    # Group by status, then within each status group by coverage type
    by_status = defaultdict(list)
    for p in parsed:
        by_status[p["status"]].append(p)

    # Sort within each status group by coverage type order
    for status in by_status:
        by_status[status].sort(key=lambda p: _coverage_sort_key(p["coverage_type"]))

    # ── Build output ────────────────────────────────────────────────

    title = opportunity_name or client_name or "Client"
    lines = []
    lines.append(f"📊 *Marketing Summary*")
    lines.append(f"*{title}*")
    lines.append(f"━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    lines.append(f"Total Policies: *{len(policies)}*")
    lines.append("")

    # Status overview
    lines.append("*Status Overview:*")
    for status in STATUS_ORDER:
        if status in by_status:
            emoji = STATUS_EMOJI.get(status, "⚪")
            count = len(by_status[status])
            desc = STATUS_DESCRIPTIONS.get(status, "")
            lines.append(f"  {emoji} {status}: *{count}* — _{desc}_")
    for status in by_status:
        if status not in STATUS_ORDER:
            count = len(by_status[status])
            lines.append(f"  ⚪ {status}: *{count}*")
    lines.append("")

    # Detailed breakdown by status → coverage type
    for status in STATUS_ORDER:
        if status not in by_status:
            continue

        emoji = STATUS_EMOJI.get(status, "⚪")
        policies_in_status = by_status[status]

        lines.append(f"{emoji} *{status.upper()}* ({len(policies_in_status)})")
        lines.append(f"{'─' * 35}")

        for p in policies_in_status:
            abbr = COVERAGE_ABBR.get(p["coverage_type"], p["coverage_type"][:4].upper())
            premium_str = f"${p['premium']:,.0f}" if p['premium'] else "N/A"
            carrier_display = p["insurance_company"]

            lines.append(f"  *{abbr} — {carrier_display}*")
            lines.append(f"    Premium: {premium_str}")

            # Coverage-specific details
            detail_func = COVERAGE_DETAIL_FUNCS.get(p["coverage_type"])
            if detail_func:
                detail_lines = detail_func(p["fields"])
                lines.extend(detail_lines)

            # Broker
            broker_val = p["broker"]
            if broker_val != "—":
                lines.append(f"    Broker: {broker_val}")

            # Comments
            if p["comments"]:
                lines.append(f"    💬 _{p['comments']}_")

            lines.append("")

    # Any statuses not in predefined order
    for status in by_status:
        if status in STATUS_ORDER:
            continue
        policies_in_status = by_status[status]
        policies_in_status.sort(key=lambda p: _coverage_sort_key(p["coverage_type"]))
        lines.append(f"⚪ *{status.upper()}* ({len(policies_in_status)})")
        lines.append(f"{'─' * 35}")
        for p in policies_in_status:
            abbr = COVERAGE_ABBR.get(p["coverage_type"], p["coverage_type"][:4].upper())
            premium_str = f"${p['premium']:,.0f}" if p['premium'] else "N/A"
            lines.append(f"  *{abbr} — {p['insurance_company']}*")
            lines.append(f"    Premium: {premium_str}")
            detail_func = COVERAGE_DETAIL_FUNCS.get(p["coverage_type"])
            if detail_func:
                lines.extend(detail_func(p["fields"]))
            broker_val = p["broker"]
            if broker_val != "—":
                lines.append(f"    Broker: {broker_val}")
            if p["comments"]:
                lines.append(f"    💬 _{p['comments']}_")
        lines.append("")

    # ── Action items ────────────────────────────────────────────────

    submit_policies = by_status.get("Submit", [])
    if submit_policies:
        lines.append("⚠️ *ACTION REQUIRED — EXPOSED:*")
        lines.append(f"  {len(submit_policies)} carrier(s) at SUBMIT — submissions needed!")
        for p in submit_policies:
            abbr = COVERAGE_ABBR.get(p["coverage_type"], p["coverage_type"][:4].upper())
            lines.append(f"  🔴 {abbr} — {p['insurance_company']}")
        lines.append("")

    # ── Coverage summary ────────────────────────────────────────────

    coverage_status_map = defaultdict(dict)
    for p in parsed:
        ct = p["coverage_type"]
        st = p["status"]
        if st not in coverage_status_map[ct]:
            coverage_status_map[ct][st] = []
        coverage_status_map[ct][st].append(p)

    lines.append("*Coverage Summary:*")
    sorted_coverages = sorted(coverage_status_map.keys(), key=_coverage_sort_key)
    for ct in sorted_coverages:
        statuses = coverage_status_map[ct]
        abbr = COVERAGE_ABBR.get(ct, ct[:4].upper())
        status_parts = []
        for s in STATUS_ORDER:
            if s in statuses:
                emoji = STATUS_EMOJI.get(s, "⚪")
                count = len(statuses[s])
                status_parts.append(f"{emoji}{s}({count})")
        lines.append(f"  {abbr}: {' | '.join(status_parts)}")

    return "\n".join(lines)


async def get_marketing_summary(client_name: str) -> str:
    """Main entry point: get marketing summary for a client/opportunity.
    
    Supports partial name matching - e.g., 'Pritchard' will find 'Pritchard Hospitality'.
    If no results found with full search term, tries individual words.
    """
    logger.info(f"get_marketing_summary called for: '{client_name}'")
    
    # First, search for the opportunity
    opportunities = search_opportunity(client_name)

    if opportunities:
        # Get the most recent opportunity
        opp = opportunities[0]
        opp_fields = opp.get("fields", {})
        opp_name = opp_fields.get("Opportunity Name", opp_fields.get("Name", client_name))
        logger.info(f"Found opportunity: '{opp_name}'")

        # Get linked Policy record IDs from the Opportunity
        policy_ids = opp_fields.get("Policies", [])
        logger.info(f"Opportunity has {len(policy_ids)} linked policy IDs")

        if policy_ids:
            # Fetch policies by their record IDs
            policies = fetch_policies_by_record_ids(policy_ids)
        else:
            # Fall back to name-based search
            policies = fetch_policies_for_client(client_name)

        if policies:
            return build_marketing_summary(policies, client_name, opp_name)

    # No opportunity found or no policies from opportunity - search policies directly
    logger.info(f"Trying direct policy search for '{client_name}'")
    policies = fetch_policies_for_client(client_name)
    
    if policies:
        return build_marketing_summary(policies, client_name)
    
    # Still nothing - try searching with individual words (for partial matches)
    words = client_name.strip().split()
    if len(words) > 1:
        for word in words:
            if len(word) >= 3:  # Skip very short words
                logger.info(f"Trying individual word search: '{word}'")
                opportunities = search_opportunity(word)
                if opportunities:
                    opp = opportunities[0]
                    opp_fields = opp.get("fields", {})
                    opp_name = opp_fields.get("Opportunity Name", opp_fields.get("Name", client_name))
                    policy_ids = opp_fields.get("Policies", [])
                    if policy_ids:
                        policies = fetch_policies_by_record_ids(policy_ids)
                        if policies:
                            return build_marketing_summary(policies, client_name, opp_name)
                
                policies = fetch_policies_for_client(word)
                if policies:
                    return build_marketing_summary(policies, client_name)
    
    return f"No policies found for {client_name}."
