#!/usr/bin/env python3
"""
Marketing Update DOCX Generator.
Pulls policies from the Airtable Sales System for a given client/opportunity
and generates a professional Marketing Update document in DOCX format.

Two versions:
  - Internal (/marketingsummary): includes commission, revenue, rates, broker info
  - Client-facing (/marketingsummaryclient): hides sensitive financial data

Uses HUB International branding with color-coded carrier comparison tables.
"""

import os
import logging
import tempfile
from collections import defaultdict
from datetime import datetime

import requests as http_requests
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml

logger = logging.getLogger(__name__)

# ── Airtable Config ──
AIRTABLE_PAT = os.environ.get("AIRTABLE_PAT", "")
SALES_BASE_ID = "appnFKEzmdLbR4CHY"
POLICIES_TABLE_ID = "tbl8vZP2oHrinwVfd"
OPPORTUNITIES_TABLE_ID = "tblMKuUsG1cosdQPN"

# ── HUB Colors ──
ELECTRIC_BLUE = RGBColor(0x16, 0x7B, 0xD4)
CLASSIC_BLUE = RGBColor(0x26, 0x38, 0x45)
ARCTIC_GRAY = RGBColor(0xB8, 0xC4, 0xCE)
EGGSHELL = RGBColor(0xF3, 0xF5, 0xF1)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CHARCOAL = RGBColor(0x4A, 0x4A, 0x4A)
DARK_GREEN = RGBColor(0x1B, 0x7A, 0x3D)
DARK_RED = RGBColor(0xC0, 0x39, 0x2B)

ELECTRIC_BLUE_HEX = "167BD4"
CLASSIC_BLUE_HEX = "263845"
EGGSHELL_HEX = "F3F5F1"
LIGHT_GREEN_HEX = "E8F5E9"
LIGHT_RED_HEX = "FDEDEC"
LIGHT_BLUE_HEX = "E3F2FD"
LIGHT_GRAY_HEX = "F0F0F0"
PROPOSED_GREEN_HEX = "E8F5E9"
EXPIRING_GRAY_HEX = "EDEFF1"
PENDING_YELLOW_HEX = "FFF8E1"

LOGO_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "assets", "hub_logo.png")

# ── Coverage ordering ──
COVERAGE_ORDER = [
    "Property", "Liability", "Umbrella", "Workers Compensation",
    "Auto", "Employment Practices Liability", "Cyber", "Flood",
    "Equipment Breakdown", "Terrorism", "Crime", "Inland Marine", "Package",
]

COVERAGE_DISPLAY_NAMES = {
    "Property": "Property",
    "Liability": "General Liability",
    "Umbrella": "Umbrella / Excess Liability",
    "Workers Compensation": "Workers Compensation",
    "Auto": "Commercial Auto",
    "Employment Practices Liability": "Employment Practices Liability (EPLI)",
    "Cyber": "Cyber Liability",
    "Flood": "Flood",
    "Equipment Breakdown": "Equipment Breakdown",
    "Terrorism": "Terrorism / TRIA",
    "Crime": "Crime",
    "Inland Marine": "Inland Marine",
    "Package": "Package",
}

COVERAGE_SHORT = {
    "Property": "Property",
    "Liability": "General Liability",
    "Umbrella": "Umbrella",
    "Workers Compensation": "Workers Comp",
    "Auto": "Auto",
    "Employment Practices Liability": "EPLI",
    "Cyber": "Cyber",
    "Flood": "Flood",
    "Equipment Breakdown": "Equip Breakdown",
    "Terrorism": "Terrorism",
    "Crime": "Crime",
    "Inland Marine": "Inland Marine",
    "Package": "Package",
}

STATUS_PRIORITY = {
    "Incumbent": 0, "Bound": 1, "Proposed": 2, "Quoted": 3,
    "Market": 4, "Submit": 5, "Pending": 6, "Declined": 7,
    "Blocked": 8, "Lost": 9,
}


# ══════════════════════════════════════════════════════════════════════════
# UTILITY FUNCTIONS
# ══════════════════════════════════════════════════════════════════════════

def airtable_headers():
    return {
        "Authorization": f"Bearer {AIRTABLE_PAT}",
        "Content-Type": "application/json",
    }


def _safe_str(val, default="—"):
    if val is None:
        return default
    if isinstance(val, list):
        return ", ".join(str(v) for v in val if v is not None) or default
    s = str(val).strip()
    return s if s else default


def _safe_currency(val, default="—"):
    if val is None:
        return default
    try:
        n = float(str(val).replace("$", "").replace(",", ""))
        if n == 0:
            return default
        return f"${n:,.2f}"
    except (ValueError, TypeError):
        return default


def _safe_currency_int(val, default="—"):
    if val is None:
        return default
    try:
        n = float(str(val).replace("$", "").replace(",", ""))
        if n == 0:
            return default
        return f"${n:,.0f}"
    except (ValueError, TypeError):
        return default


def _safe_number(val, default="—"):
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


def _safe_percent(val, default="—"):
    if val is None:
        return default
    try:
        n = float(val)
        if n == 0:
            return default
        return f"{n:.2%}" if n < 1 else f"{n:.2f}%"
    except (ValueError, TypeError):
        return default


def _get_float(val, default=0):
    if val is None:
        return default
    try:
        return float(str(val).replace("$", "").replace(",", "").replace("%", ""))
    except (ValueError, TypeError):
        return default


def _normalize_coverage_type(policy_type_raw):
    if isinstance(policy_type_raw, list):
        pt = policy_type_raw[0] if policy_type_raw else "Other"
    else:
        pt = str(policy_type_raw) if policy_type_raw else "Other"
    pt_lower = pt.lower().strip()
    if "property" in pt_lower:
        return "Property"
    elif "terror" in pt_lower or "tria" in pt_lower:
        return "Terrorism"
    elif "liability" in pt_lower and "employment" not in pt_lower:
        return "Liability"
    elif "umbrella" in pt_lower or "excess" in pt_lower:
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
    elif "crime" in pt_lower:
        return "Crime"
    elif "inland" in pt_lower or "marine" in pt_lower:
        return "Inland Marine"
    elif "package" in pt_lower or "pkg" in pt_lower:
        return "Package"
    else:
        return pt if pt else "Other"


def _coverage_sort_key(coverage_type):
    try:
        return COVERAGE_ORDER.index(coverage_type)
    except ValueError:
        return len(COVERAGE_ORDER)


def _status_sort_key(status):
    return STATUS_PRIORITY.get(status, 99)


def _map_status_for_display(status):
    """Map Airtable status to display status for the document."""
    s = (status or "").strip()
    if s == "Incumbent":
        return "Expiring"
    elif s in ("Bound", "Proposed"):
        return s
    elif s in ("Quoted",):
        return "Quoted"
    elif s in ("Market", "Submit"):
        return "Pending"
    elif s in ("Declined", "Blocked", "Lost"):
        return s
    else:
        return s or "Unknown"


def _status_color_hex(display_status):
    """Return background color hex for carrier column based on display status."""
    if display_status == "Expiring":
        return EXPIRING_GRAY_HEX
    elif display_status in ("Bound", "Proposed"):
        return PROPOSED_GREEN_HEX
    elif display_status == "Quoted":
        return None  # White / alternating
    elif display_status == "Pending":
        return PENDING_YELLOW_HEX
    else:
        return None


# ══════════════════════════════════════════════════════════════════════════
# AIRTABLE DATA FETCHING (reuse from marketing_summary.py)
# ══════════════════════════════════════════════════════════════════════════

def _sanitize_for_formula(text: str) -> str:
    return text.replace('"', '\\"')


def search_opportunity(client_name: str, upcoming_only: bool = True) -> list:
    safe_name = _sanitize_for_formula(client_name)
    dq = '"'
    search_part = (
        f"OR("
        f"SEARCH(LOWER({dq}{safe_name}{dq}), LOWER({{Opportunity Name}})),"
        f"SEARCH(LOWER({dq}{safe_name}{dq}), LOWER(ARRAYJOIN({{Corporate Name}},{dq}{dq})))"
        f")"
    )
    if upcoming_only:
        formula = f"AND({search_part}, {{Days to Expiration}} >= 0)"
    else:
        formula = search_part

    url = f"https://api.airtable.com/v0/{SALES_BASE_ID}/{OPPORTUNITIES_TABLE_ID}"
    params = {
        "filterByFormula": formula,
        "sort[0][field]": "Effective Date",
        "sort[0][direction]": "asc",
        "pageSize": 20,
    }
    try:
        resp = http_requests.get(url, headers=airtable_headers(), params=params, timeout=30)
        resp.raise_for_status()
        return resp.json().get("records", [])
    except Exception as e:
        logger.error(f"Error searching opportunities: {e}")
        return []


def fetch_policies_by_record_ids(policy_record_ids: list) -> list:
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
            params = {"filterByFormula": formula, "pageSize": 100}
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


def fetch_policies_for_client(client_name: str) -> list:
    safe_name = _sanitize_for_formula(client_name)
    dq = '"'
    formula = (
        f"OR("
        f"SEARCH(LOWER({dq}{safe_name}{dq}), LOWER({{Name}})),"
        f"SEARCH(LOWER({dq}{safe_name}{dq}), LOWER(ARRAYJOIN({{Companies}},{dq}{dq})))"
        f")"
    )
    url = f"https://api.airtable.com/v0/{SALES_BASE_ID}/{POLICIES_TABLE_ID}"
    all_records = []
    offset = None
    while True:
        params = {"filterByFormula": formula, "pageSize": 100}
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


async def resolve_client_data(client_name: str):
    """Resolve client name to opportunity + policies. Returns (opp_fields, policies) or (None, [])."""
    opportunities = search_opportunity(client_name)
    if opportunities:
        opp = opportunities[0]
        opp_fields = opp.get("fields", {})
        policy_ids = opp_fields.get("Policies", [])
        if policy_ids:
            policies = fetch_policies_by_record_ids(policy_ids)
        else:
            policies = fetch_policies_for_client(client_name)
        if policies:
            return opp_fields, policies

    # Direct policy search
    policies = fetch_policies_for_client(client_name)
    if policies:
        return None, policies

    # Try individual words
    words = client_name.strip().split()
    if len(words) > 1:
        for word in words:
            if len(word) >= 3:
                opportunities = search_opportunity(word)
                if opportunities:
                    opp = opportunities[0]
                    opp_fields = opp.get("fields", {})
                    policy_ids = opp_fields.get("Policies", [])
                    if policy_ids:
                        policies = fetch_policies_by_record_ids(policy_ids)
                        if policies:
                            return opp_fields, policies
                policies = fetch_policies_for_client(word)
                if policies:
                    return None, policies

    return None, []


# ══════════════════════════════════════════════════════════════════════════
# POLICY DATA PARSING
# ══════════════════════════════════════════════════════════════════════════

def parse_policies(policies: list):
    """Parse raw Airtable policy records into structured dicts grouped by coverage type."""
    parsed = []
    for rec in policies:
        flds = rec.get("fields", {})
        coverage_type = _normalize_coverage_type(flds.get("Policy Type"))

        insurance_co_raw = flds.get("Insurance Company", "")
        if isinstance(insurance_co_raw, list):
            insurance_co = ", ".join(str(ic) for ic in insurance_co_raw if ic) or "N/A"
        else:
            insurance_co = str(insurance_co_raw).strip() if insurance_co_raw else "N/A"

        status = flds.get("Status", "Unknown")
        display_status = _map_status_for_display(status)

        base_premium = _get_float(flds.get("Base Premium"))
        premium_tx = _get_float(flds.get("Premium Tx"))
        # Use Premium Tx if available, else Base Premium
        premium_with_tax = premium_tx if premium_tx > 0 else base_premium

        commission = _get_float(flds.get("Commission"))
        revenue = _get_float(flds.get("Revenue"))

        parsed.append({
            "coverage_type": coverage_type,
            "carrier": insurance_co,
            "status": status,
            "display_status": display_status,
            "base_premium": base_premium,
            "premium_tx": premium_with_tax,
            "commission": commission,
            "revenue": revenue,
            "broker": _safe_str(flds.get("Broker ABBR")),
            "units": flds.get("Units"),
            "num_locs": flds.get("# of Locs"),
            "tiv": flds.get("TIV"),
            "property_rate": flds.get("Property Rate"),
            "property_limit": _safe_str(flds.get("Property Limit")),
            "aop": _safe_str(flds.get("AOP")),
            "wind_type": _safe_str(flds.get("Wind Type")),
            "wind": _safe_str(flds.get("Wind")),
            "aow": _safe_str(flds.get("AOW")),
            "water_damage": _safe_str(flds.get("Water Damage")),
            "flood_limit": _safe_str(flds.get("Flood Limit")),
            "flood_deductible": _safe_str(flds.get("Flood Deductible")),
            "eq_limit": _safe_str(flds.get("EQ Limit")),
            "eq_deductible": _safe_str(flds.get("Earthquake Deductible")),
            "gross_sales": flds.get("Gross Sales"),
            "gl_rate": flds.get("GL Rate $"),
            "gl_rate_unit": flds.get("GL Rate (u)"),
            "gl_deductible": _safe_str(flds.get("GL Deductible")),
            "umb_limit": _safe_str(flds.get("UMB Limit")),
            "total_payroll": flds.get("Total Payroll"),
            "exp_mod": flds.get("Exp Mod"),
            "safety_credit": flds.get("Safety"),
            "drug_free_credit": flds.get("Drug Free"),
            "comments": flds.get("Comments", ""),
            "fields": flds,
        })

    return parsed


def group_by_coverage(parsed_policies):
    """Group parsed policies by coverage type, sorted by coverage order."""
    by_coverage = defaultdict(list)
    for p in parsed_policies:
        by_coverage[p["coverage_type"]].append(p)

    # Sort carriers within each coverage: Expiring first, then Bound/Proposed, then Quoted, then Pending
    for ct in by_coverage:
        by_coverage[ct].sort(key=lambda p: _status_sort_key(p["status"]))

    return by_coverage


# ══════════════════════════════════════════════════════════════════════════
# DOCX BUILDING UTILITIES
# ══════════════════════════════════════════════════════════════════════════

def set_cell_shading(cell, color_hex):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}" w:val="clear"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_width(cell, width_inches):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{int(width_inches * 1440)}" w:type="dxa"/>')
    existing = tcPr.find(qn('w:tcW'))
    if existing is not None:
        tcPr.remove(existing)
    tcPr.append(tcW)


def set_cell_vertical_alignment(cell, val="center"):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    vAlign = parse_xml(f'<w:vAlign {nsdecls("w")} w:val="{val}"/>')
    existing = tcPr.find(qn('w:vAlign'))
    if existing is not None:
        tcPr.remove(existing)
    tcPr.append(vAlign)


def remove_cell_borders(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'</w:tcBorders>'
    )
    existing = tcPr.find(qn('w:tcBorders'))
    if existing is not None:
        tcPr.remove(existing)
    tcPr.append(borders)


def remove_table_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'<w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'</w:tblBorders>'
    )
    existing = tblPr.find(qn('w:tblBorders'))
    if existing is not None:
        tblPr.remove(existing)
    tblPr.append(borders)


def set_thin_borders(table):
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="D0D0D0"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="D0D0D0"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="D0D0D0"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="D0D0D0"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="D0D0D0"/>'
        f'<w:insideV w:val="single" w:sz="4" w:space="0" w:color="D0D0D0"/>'
        f'</w:tblBorders>'
    )
    existing = tblPr.find(qn('w:tblBorders'))
    if existing is not None:
        tblPr.remove(existing)
    tblPr.append(borders)


def add_page_header(doc):
    section = doc.sections[-1]
    header = section.header
    header.is_linked_to_previous = False
    htable = header.add_table(1, 2, width=Inches(7))
    htable.alignment = WD_TABLE_ALIGNMENT.CENTER
    logo_cell = htable.rows[0].cells[0]
    logo_cell.width = Inches(2.5)
    if os.path.exists(LOGO_PATH):
        p = logo_cell.paragraphs[0]
        run = p.add_run()
        run.add_picture(LOGO_PATH, width=Inches(1.8))
    text_cell = htable.rows[0].cells[1]
    text_cell.width = Inches(4.5)
    p = text_cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run("Franchise Division")
    run.font.size = Pt(12)
    run.font.color.rgb = ELECTRIC_BLUE
    run.font.bold = True
    run.font.name = "Calibri"
    p2 = text_cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p2.paragraph_format.space_before = Pt(0)
    run2 = p2.add_run("Hotel Insurance Programs")
    run2.font.size = Pt(12)
    run2.font.color.rgb = ELECTRIC_BLUE
    run2.font.bold = True
    run2.font.name = "Calibri"
    for row in htable.rows:
        for cell in row.cells:
            remove_cell_borders(cell)


def add_formatted_paragraph(doc, text, size=11, color=CLASSIC_BLUE, bold=False,
                            italic=False, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                            space_before=0, space_after=0):
    p = doc.add_paragraph()
    p.alignment = alignment
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = Pt(size + 3)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = "Calibri"
    return p


def add_section_header(doc, text):
    return add_formatted_paragraph(doc, text, size=22, color=CLASSIC_BLUE, bold=True,
                                   space_before=24, space_after=12)


def add_subsection_header(doc, text):
    return add_formatted_paragraph(doc, text, size=14, color=ELECTRIC_BLUE, bold=True,
                                   space_before=14, space_after=6)


def add_callout_box(doc, text, size=10, shading_hex=EGGSHELL_HEX):
    table = doc.add_table(rows=1, cols=1)
    cell = table.rows[0].cells[0]
    set_cell_shading(cell, shading_hex)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Inches(0.1)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.color.rgb = CLASSIC_BLUE
    run.font.name = "Calibri"
    run.font.italic = True
    return table


def add_rich_callout_box(doc, lines, size=10, shading_hex=EGGSHELL_HEX):
    table = doc.add_table(rows=1, cols=1)
    cell = table.rows[0].cells[0]
    set_cell_shading(cell, shading_hex)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.left_indent = Inches(0.1)
    for i, line_data in enumerate(lines):
        if i > 0:
            p = cell.add_paragraph()
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.left_indent = Inches(0.1)
        text, bold, color = line_data
        run = p.add_run(text)
        run.font.size = Pt(size)
        run.font.color.rgb = color
        run.font.name = "Calibri"
        run.font.bold = bold
    return table


def add_divider(doc):
    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(4)
    p_div.paragraph_format.space_after = Pt(4)
    pPr = p_div._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="12" w:space="1" w:color="{ELECTRIC_BLUE_HEX}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)


def add_page_break(doc):
    doc.add_page_break()


# ══════════════════════════════════════════════════════════════════════════
# TABLE BUILDERS
# ══════════════════════════════════════════════════════════════════════════

def create_premium_summary_table(doc, headers, rows, col_widths=None, highlight_last_row=False):
    """Create the premium summary comparison table."""
    num_cols = len(headers)
    table = doc.add_table(rows=1 + len(rows), cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
    tblLayout = parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>')
    existing = tblPr.find(qn('w:tblLayout'))
    if existing is not None:
        tblPr.remove(existing)
    tblPr.append(tblLayout)
    total_width = sum(col_widths) if col_widths else 7.5
    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="{int(total_width * 1440)}" w:type="dxa"/>')
    existing_tblW = tblPr.find(qn('w:tblW'))
    if existing_tblW is not None:
        tblPr.remove(existing_tblW)
    tblPr.append(tblW)

    if not col_widths:
        col_widths = [total_width / num_cols] * num_cols

    # Header row
    for i, header_text in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER if i >= 2 else WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(header_text)
        run.font.size = Pt(10)
        run.font.color.rgb = WHITE
        run.font.bold = True
        run.font.name = "Calibri"
        set_cell_shading(cell, ELECTRIC_BLUE_HEX)
        set_cell_width(cell, col_widths[i])
        set_cell_vertical_alignment(cell, "center")

    # Data rows
    for row_idx, row_data in enumerate(rows):
        is_total = highlight_last_row and row_idx == len(rows) - 1
        for col_idx, val in enumerate(row_data):
            cell = table.rows[row_idx + 1].cells[col_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            if col_idx < 2:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            run = p.add_run(str(val))
            run.font.size = Pt(9)
            run.font.name = "Calibri"

            if is_total:
                run.font.color.rgb = WHITE
                run.font.bold = True
                set_cell_shading(cell, CLASSIC_BLUE_HEX)
            else:
                run.font.color.rgb = CLASSIC_BLUE
                # Color code $ Change column
                val_str = str(val)
                if "Change" in headers[col_idx] if col_idx < len(headers) else False:
                    pass
                if col_idx >= 2 and val_str.startswith("-$"):
                    run.font.color.rgb = DARK_GREEN
                elif col_idx >= 2 and val_str.startswith("+$"):
                    run.font.color.rgb = DARK_RED
                if row_idx % 2 == 1:
                    set_cell_shading(cell, EGGSHELL_HEX)

            set_cell_width(cell, col_widths[col_idx])
            set_cell_vertical_alignment(cell, "center")

    return table


def create_carrier_comparison_table(doc, coverage_title, metrics, carriers):
    """
    Create a carrier comparison table for a coverage line.
    carriers: list of dicts with keys: name, status, values (dict), notes
    metrics: list of metric labels (rows)
    """
    num_cols = 1 + len(carriers)
    num_rows = 1 + len(metrics)

    has_notes = any(c.get("notes") for c in carriers)
    if has_notes:
        num_rows += 1

    table = doc.add_table(rows=num_rows, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_thin_borders(table)

    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
    tblLayout = parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>')
    existing = tblPr.find(qn('w:tblLayout'))
    if existing is not None:
        tblPr.remove(existing)
    tblPr.append(tblLayout)

    total_width = 7.5
    label_width = 1.6
    carrier_width = (total_width - label_width) / max(len(carriers), 1)

    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="{int(total_width * 1440)}" w:type="dxa"/>')
    existing_tblW = tblPr.find(qn('w:tblW'))
    if existing_tblW is not None:
        tblPr.remove(existing_tblW)
    tblPr.append(tblW)

    # Header row
    header_cell = table.rows[0].cells[0]
    header_cell.text = ""
    p = header_cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(coverage_title)
    run.font.size = Pt(10)
    run.font.color.rgb = WHITE
    run.font.bold = True
    run.font.name = "Calibri"
    set_cell_shading(header_cell, ELECTRIC_BLUE_HEX)
    set_cell_width(header_cell, label_width)
    set_cell_vertical_alignment(header_cell, "center")

    for c_idx, carrier in enumerate(carriers):
        cell = table.rows[0].cells[1 + c_idx]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = Pt(12)

        run = p.add_run(carrier["name"])
        run.font.size = Pt(9)
        run.font.color.rgb = WHITE
        run.font.bold = True
        run.font.name = "Calibri"

        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(3)
        p2.paragraph_format.line_spacing = Pt(10)
        run2 = p2.add_run(f"({carrier['status']})")
        run2.font.size = Pt(8)
        run2.font.color.rgb = RGBColor(0xCC, 0xE5, 0xFF)
        run2.font.italic = True
        run2.font.name = "Calibri"

        set_cell_shading(cell, ELECTRIC_BLUE_HEX)
        set_cell_width(cell, carrier_width)
        set_cell_vertical_alignment(cell, "center")

    # Metric rows
    for m_idx, metric in enumerate(metrics):
        label_cell = table.rows[1 + m_idx].cells[0]
        label_cell.text = ""
        p = label_cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = Pt(12)
        run = p.add_run(metric)
        run.font.size = Pt(9)
        run.font.color.rgb = CLASSIC_BLUE
        run.font.bold = True
        run.font.name = "Calibri"
        set_cell_width(label_cell, label_width)
        set_cell_vertical_alignment(label_cell, "center")

        is_premium_row = (metric == "Premium")

        for c_idx, carrier in enumerate(carriers):
            cell = table.rows[1 + m_idx].cells[1 + c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = Pt(12)

            val = carrier["values"].get(metric, "—")
            run = p.add_run(str(val))
            run.font.size = Pt(9)
            run.font.name = "Calibri"

            if is_premium_row:
                run.font.bold = True
                run.font.size = Pt(10)

            # Color coding by status
            status = carrier["status"]
            bg_hex = _status_color_hex(status)
            if bg_hex:
                set_cell_shading(cell, bg_hex)
            elif m_idx % 2 == 1:
                set_cell_shading(cell, "FAFAFA")

            run.font.color.rgb = CLASSIC_BLUE if status != "Pending" else CHARCOAL

            set_cell_width(cell, carrier_width)
            set_cell_vertical_alignment(cell, "center")

    # Notes row
    if has_notes:
        notes_row_idx = 1 + len(metrics)
        label_cell = table.rows[notes_row_idx].cells[0]
        label_cell.text = ""
        p = label_cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        run = p.add_run("Notes")
        run.font.size = Pt(9)
        run.font.color.rgb = CLASSIC_BLUE
        run.font.bold = True
        run.font.italic = True
        run.font.name = "Calibri"
        set_cell_width(label_cell, label_width)

        for c_idx, carrier in enumerate(carriers):
            cell = table.rows[notes_row_idx].cells[1 + c_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = Pt(11)
            note = carrier.get("notes", "—") or "—"
            run = p.add_run(note)
            run.font.size = Pt(8)
            run.font.color.rgb = CHARCOAL
            run.font.italic = True
            run.font.name = "Calibri"
            set_cell_width(cell, carrier_width)

            bg_hex = _status_color_hex(carrier["status"])
            if bg_hex:
                set_cell_shading(cell, bg_hex)

    return table


# ══════════════════════════════════════════════════════════════════════════
# COVERAGE-SPECIFIC METRIC BUILDERS
# ══════════════════════════════════════════════════════════════════════════

def _build_property_carrier(p, is_internal=True):
    """Build carrier dict for property comparison table."""
    values = {
        "Premium": _safe_currency(p["premium_tx"]) if p["premium_tx"] else "Pending",
        "TIV": _safe_currency_int(p["tiv"]),
        "AOP Deductible": p["aop"],
        "Wind": f"{p['wind']} ({p['wind_type']})" if p["wind"] != "—" and p["wind_type"] != "—" else p["wind"],
        "AOW (All Other Wind)": p["aow"],
        "Water Damage": p["water_damage"],
    }
    if is_internal:
        values["Property Rate"] = _safe_currency(p["property_rate"]) if p["property_rate"] else "—"
    if p["property_limit"] != "—":
        values["Property Limit"] = p["property_limit"]
    if p["flood_limit"] != "—":
        values["Flood Limit"] = p["flood_limit"]
    if p["eq_limit"] != "—":
        values["EQ Limit"] = p["eq_limit"]
    if is_internal:
        if p["commission"]:
            values["Commission"] = _safe_percent(p["commission"])
        if p["revenue"]:
            values["Revenue"] = _safe_currency(p["revenue"])

    # Build notes from comments
    comments = p.get("comments", "")
    if isinstance(comments, str) and len(comments) > 100:
        comments = comments[:97] + "..."

    return {
        "name": p["carrier"],
        "status": p["display_status"],
        "values": values,
        "notes": comments if comments else "",
    }


def _get_property_metrics(carriers_data, is_internal=True):
    """Determine which property metrics to show based on available data."""
    metrics = ["Premium", "TIV"]
    if is_internal:
        metrics.append("Property Rate")
    metrics.extend(["AOP Deductible", "Wind", "AOW (All Other Wind)", "Water Damage"])

    # Only add optional metrics if any carrier has data
    all_values = {}
    for c in carriers_data:
        for k, v in c["values"].items():
            if k not in all_values:
                all_values[k] = []
            all_values[k].append(v)

    optional_metrics = ["Property Limit", "Flood Limit", "EQ Limit"]
    if is_internal:
        optional_metrics.extend(["Commission", "Revenue"])

    for m in optional_metrics:
        if m in all_values and any(v != "—" for v in all_values[m]):
            metrics.append(m)

    return metrics


def _build_gl_carrier(p, is_internal=True):
    values = {
        "Premium": _safe_currency(p["premium_tx"]) if p["premium_tx"] else "Pending",
        "# of Units": _safe_number(p["units"]),
        "Total Sales": _safe_currency_int(p["gross_sales"]),
        "GL Deductible": p["gl_deductible"],
        "# of Locations": _safe_number(p["num_locs"]),
    }
    if is_internal:
        values["GL Rate"] = _safe_currency(p["gl_rate"]) if p["gl_rate"] else "—"
        values["GL Rate/Unit"] = _safe_currency(p["gl_rate_unit"]) if p["gl_rate_unit"] else "—"
    if is_internal:
        if p["commission"]:
            values["Commission"] = _safe_percent(p["commission"])
        if p["revenue"]:
            values["Revenue"] = _safe_currency(p["revenue"])

    comments = p.get("comments", "")
    if isinstance(comments, str) and len(comments) > 100:
        comments = comments[:97] + "..."

    return {
        "name": p["carrier"],
        "status": p["display_status"],
        "values": values,
        "notes": comments if comments else "",
    }


def _get_gl_metrics(carriers_data, is_internal=True):
    metrics = ["Premium", "# of Units", "Total Sales"]
    if is_internal:
        metrics.extend(["GL Rate", "GL Rate/Unit"])
    metrics.extend(["GL Deductible", "# of Locations"])
    if is_internal:
        all_values = {}
        for c in carriers_data:
            for k, v in c["values"].items():
                if k not in all_values:
                    all_values[k] = []
                all_values[k].append(v)
        for m in ["Commission", "Revenue"]:
            if m in all_values and any(v != "—" for v in all_values[m]):
                metrics.append(m)
    return metrics


def _build_umbrella_carrier(p, is_internal=True):
    values = {
        "Premium": _safe_currency(p["premium_tx"]) if p["premium_tx"] else "Pending",
        "# of Units": _safe_number(p["units"]),
        "Umbrella Limit": p["umb_limit"],
        "Total Sales": _safe_currency_int(p["gross_sales"]),
    }
    # Calculate umbrella rate per unit if data available
    if is_internal and p["premium_tx"] and p["units"]:
        try:
            rate_per_unit = p["premium_tx"] / float(p["units"])
            values["Rate/Unit"] = f"${rate_per_unit:,.2f}"
        except (ValueError, TypeError, ZeroDivisionError):
            pass
    if is_internal:
        if p["commission"]:
            values["Commission"] = _safe_percent(p["commission"])
        if p["revenue"]:
            values["Revenue"] = _safe_currency(p["revenue"])

    comments = p.get("comments", "")
    if isinstance(comments, str) and len(comments) > 100:
        comments = comments[:97] + "..."

    return {
        "name": p["carrier"],
        "status": p["display_status"],
        "values": values,
        "notes": comments if comments else "",
    }


def _get_umbrella_metrics(carriers_data, is_internal=True):
    metrics = ["Premium", "# of Units", "Umbrella Limit", "Total Sales"]
    if is_internal:
        all_values = {}
        for c in carriers_data:
            for k, v in c["values"].items():
                if k not in all_values:
                    all_values[k] = []
                all_values[k].append(v)
        if "Rate/Unit" in all_values and any(v != "—" for v in all_values.get("Rate/Unit", [])):
            metrics.append("Rate/Unit")
        for m in ["Commission", "Revenue"]:
            if m in all_values and any(v != "—" for v in all_values[m]):
                metrics.append(m)
    return metrics


def _build_wc_carrier(p, is_internal=True):
    values = {
        "Premium": _safe_currency(p["premium_tx"]) if p["premium_tx"] else "Pending",
        "Total Payroll": _safe_currency_int(p["total_payroll"]),
    }
    if p["exp_mod"]:
        try:
            em = float(p["exp_mod"])
            if em > 0:
                values["Exp Mod"] = f"{em:.2f}"
        except (ValueError, TypeError):
            pass
    if p["safety_credit"]:
        values["Safety Credit"] = "Yes" if p["safety_credit"] else "No"
    if p["drug_free_credit"]:
        values["Drug Free Credit"] = "Yes" if p["drug_free_credit"] else "No"
    if is_internal:
        if p["commission"]:
            values["Commission"] = _safe_percent(p["commission"])
        if p["revenue"]:
            values["Revenue"] = _safe_currency(p["revenue"])

    comments = p.get("comments", "")
    if isinstance(comments, str) and len(comments) > 100:
        comments = comments[:97] + "..."

    return {
        "name": p["carrier"],
        "status": p["display_status"],
        "values": values,
        "notes": comments if comments else "",
    }


def _get_wc_metrics(carriers_data, is_internal=True):
    metrics = ["Premium", "Total Payroll"]
    all_values = {}
    for c in carriers_data:
        for k, v in c["values"].items():
            if k not in all_values:
                all_values[k] = []
            all_values[k].append(v)
    for m in ["Exp Mod", "Safety Credit", "Drug Free Credit"]:
        if m in all_values and any(v not in ("—", "No") for v in all_values[m]):
            metrics.append(m)
    if is_internal:
        for m in ["Commission", "Revenue"]:
            if m in all_values and any(v != "—" for v in all_values[m]):
                metrics.append(m)
    return metrics


def _build_generic_carrier(p, is_internal=True):
    """Generic carrier builder for EPLI, Cyber, Flood, Auto, etc."""
    values = {
        "Premium": _safe_currency(p["premium_tx"]) if p["premium_tx"] else "Pending",
    }
    if p["units"]:
        values["# of Units"] = _safe_number(p["units"])
    if p["num_locs"]:
        values["# of Locations"] = _safe_number(p["num_locs"])

    # Coverage-specific fields
    ct = p["coverage_type"]
    if ct == "Flood":
        if p["flood_limit"] != "—":
            values["Flood Limit"] = p["flood_limit"]
        if p["flood_deductible"] != "—":
            values["Flood Deductible"] = p["flood_deductible"]
    elif ct == "Employment Practices Liability":
        if p["umb_limit"] != "—":
            values["Limit"] = p["umb_limit"]

    if is_internal:
        if p["commission"]:
            values["Commission"] = _safe_percent(p["commission"])
        if p["revenue"]:
            values["Revenue"] = _safe_currency(p["revenue"])
        if p["broker"] != "—":
            values["Broker"] = p["broker"]

    comments = p.get("comments", "")
    if isinstance(comments, str) and len(comments) > 100:
        comments = comments[:97] + "..."

    return {
        "name": p["carrier"],
        "status": p["display_status"],
        "values": values,
        "notes": comments if comments else "",
    }


def _get_generic_metrics(carriers_data, is_internal=True):
    metrics = ["Premium"]
    all_values = {}
    for c in carriers_data:
        for k, v in c["values"].items():
            if k not in all_values:
                all_values[k] = []
            all_values[k].append(v)
    # Add metrics that have data
    optional = ["# of Units", "# of Locations", "Limit", "Flood Limit", "Flood Deductible",
                "Building Limit", "BPP Limit", "Retention"]
    if is_internal:
        optional.extend(["Commission", "Revenue", "Broker"])
    for m in optional:
        if m in all_values and any(v != "—" for v in all_values[m]):
            metrics.append(m)
    return metrics


# Coverage type -> (carrier_builder, metrics_builder) mapping
COVERAGE_BUILDERS = {
    "Property": (_build_property_carrier, _get_property_metrics),
    "Liability": (_build_gl_carrier, _get_gl_metrics),
    "Umbrella": (_build_umbrella_carrier, _get_umbrella_metrics),
    "Workers Compensation": (_build_wc_carrier, _get_wc_metrics),
}


# ══════════════════════════════════════════════════════════════════════════
# PREMIUM COMPARISON DATA BUILDER
# ══════════════════════════════════════════════════════════════════════════

def build_premium_comparison(by_coverage, parsed_policies):
    """Build premium comparison rows: Coverage | Carrier | Expiring | Proposed | $ Change | % Change."""
    rows = []
    total_expiring = 0
    total_proposed = 0
    pending_coverages = []

    sorted_coverages = sorted(by_coverage.keys(), key=_coverage_sort_key)

    for ct in sorted_coverages:
        policies = by_coverage[ct]
        display_name = COVERAGE_SHORT.get(ct, ct)

        # Find expiring (Incumbent) and proposed/bound
        expiring = [p for p in policies if p["status"] == "Incumbent"]
        bound = [p for p in policies if p["status"] in ("Bound", "Proposed")]
        quoted = [p for p in policies if p["status"] == "Quoted"]

        expiring_premium = sum(p["premium_tx"] for p in expiring) if expiring else 0
        proposed_premium = sum(p["premium_tx"] for p in bound) if bound else 0

        # Determine carrier name for the row
        if bound:
            carrier_name = bound[0]["carrier"]
        elif quoted:
            carrier_name = quoted[0]["carrier"]
        elif expiring:
            carrier_name = expiring[0]["carrier"]
        else:
            carrier_name = policies[0]["carrier"] if policies else "—"

        # Check if this is a "bundled" coverage (e.g., EPLI included in GL)
        is_included = False
        for p in policies:
            comments = str(p.get("comments", "")).lower()
            if "included" in comments or "inc " in comments[:10]:
                is_included = True
                break

        if proposed_premium > 0:
            exp_str = _safe_currency(expiring_premium) if expiring_premium else "—"
            prop_str = _safe_currency(proposed_premium)
            if expiring_premium > 0:
                change = proposed_premium - expiring_premium
                pct_change = (change / expiring_premium) * 100
                change_str = f"+${change:,.2f}" if change > 0 else f"-${abs(change):,.2f}"
                pct_str = f"+{pct_change:.1f}%" if change > 0 else f"{pct_change:.1f}%"
                total_expiring += expiring_premium
                total_proposed += proposed_premium
            else:
                change_str = "—"
                pct_str = "—"
                total_proposed += proposed_premium
            rows.append([display_name, carrier_name, exp_str, prop_str, change_str, pct_str])
        elif is_included:
            exp_str = "—" if not expiring_premium else _safe_currency(expiring_premium)
            rows.append([display_name, "Included in GL", exp_str, "Included", "—", "—"])
        else:
            exp_str = _safe_currency(expiring_premium) if expiring_premium else "—"
            if expiring_premium:
                total_expiring += expiring_premium
            pending_coverages.append(display_name)
            rows.append([display_name, carrier_name, exp_str, "Pending", "—", "—"])

    # Total row
    total_change = total_proposed - total_expiring if total_expiring > 0 else 0
    total_pct = (total_change / total_expiring * 100) if total_expiring > 0 else 0
    total_change_str = f"+${total_change:,.2f}" if total_change > 0 else f"-${abs(total_change):,.2f}" if total_change != 0 else "—"
    total_pct_str = f"+{total_pct:.1f}%" if total_change > 0 else f"{total_pct:.1f}%" if total_change != 0 else "—"

    rows.append([
        "TOTAL", "",
        _safe_currency(total_expiring) if total_expiring else "—",
        _safe_currency(total_proposed) if total_proposed else "—",
        total_change_str, total_pct_str,
    ])

    return rows, total_change, total_pct, pending_coverages


def build_premium_comparison_internal(by_coverage, parsed_policies):
    """Build internal premium comparison with Commission and Revenue columns."""
    rows = []
    total_expiring = 0
    total_proposed = 0
    total_commission_revenue = 0
    pending_coverages = []

    sorted_coverages = sorted(by_coverage.keys(), key=_coverage_sort_key)

    for ct in sorted_coverages:
        policies = by_coverage[ct]
        display_name = COVERAGE_SHORT.get(ct, ct)

        expiring = [p for p in policies if p["status"] == "Incumbent"]
        bound = [p for p in policies if p["status"] in ("Bound", "Proposed")]
        quoted = [p for p in policies if p["status"] == "Quoted"]

        expiring_premium = sum(p["premium_tx"] for p in expiring) if expiring else 0
        proposed_premium = sum(p["premium_tx"] for p in bound) if bound else 0

        # Get commission and revenue from bound/proposed carrier
        comm_str = "—"
        rev_str = "—"
        broker_str = "—"
        if bound:
            carrier_name = bound[0]["carrier"]
            if bound[0]["commission"]:
                comm_str = _safe_percent(bound[0]["commission"])
            if bound[0]["revenue"]:
                rev_str = _safe_currency(bound[0]["revenue"])
                total_commission_revenue += bound[0]["revenue"]
            broker_str = bound[0]["broker"]
        elif quoted:
            carrier_name = quoted[0]["carrier"]
            if quoted[0]["commission"]:
                comm_str = _safe_percent(quoted[0]["commission"])
            broker_str = quoted[0]["broker"]
        elif expiring:
            carrier_name = expiring[0]["carrier"]
            broker_str = expiring[0]["broker"]
        else:
            carrier_name = policies[0]["carrier"] if policies else "—"

        is_included = False
        for p in policies:
            comments = str(p.get("comments", "")).lower()
            if "included" in comments or "inc " in comments[:10]:
                is_included = True
                break

        if proposed_premium > 0:
            exp_str = _safe_currency(expiring_premium) if expiring_premium else "—"
            prop_str = _safe_currency(proposed_premium)
            if expiring_premium > 0:
                change = proposed_premium - expiring_premium
                pct_change = (change / expiring_premium) * 100
                change_str = f"+${change:,.2f}" if change > 0 else f"-${abs(change):,.2f}"
                pct_str = f"+{pct_change:.1f}%" if change > 0 else f"{pct_change:.1f}%"
                total_expiring += expiring_premium
                total_proposed += proposed_premium
            else:
                change_str = "—"
                pct_str = "—"
                total_proposed += proposed_premium
            rows.append([display_name, carrier_name, exp_str, prop_str, change_str, pct_str, comm_str, rev_str, broker_str])
        elif is_included:
            exp_str = "—" if not expiring_premium else _safe_currency(expiring_premium)
            rows.append([display_name, "Included in GL", exp_str, "Included", "—", "—", "—", "—", "—"])
        else:
            exp_str = _safe_currency(expiring_premium) if expiring_premium else "—"
            if expiring_premium:
                total_expiring += expiring_premium
            pending_coverages.append(display_name)
            rows.append([display_name, carrier_name, exp_str, "Pending", "—", "—", comm_str, rev_str, broker_str])

    # Total row
    total_change = total_proposed - total_expiring if total_expiring > 0 else 0
    total_pct = (total_change / total_expiring * 100) if total_expiring > 0 else 0
    total_change_str = f"+${total_change:,.2f}" if total_change > 0 else f"-${abs(total_change):,.2f}" if total_change != 0 else "—"
    total_pct_str = f"+{total_pct:.1f}%" if total_change > 0 else f"{total_pct:.1f}%" if total_change != 0 else "—"

    rows.append([
        "TOTAL", "",
        _safe_currency(total_expiring) if total_expiring else "—",
        _safe_currency(total_proposed) if total_proposed else "—",
        total_change_str, total_pct_str, "—",
        _safe_currency(total_commission_revenue) if total_commission_revenue else "—", "—",
    ])

    return rows, total_change, total_pct, pending_coverages


# ══════════════════════════════════════════════════════════════════════════
# MARKET ACTIVITY BUILDER
# ══════════════════════════════════════════════════════════════════════════

def build_market_activity(by_coverage):
    """Build market activity overview rows."""
    rows = []
    sorted_coverages = sorted(by_coverage.keys(), key=_coverage_sort_key)

    for ct in sorted_coverages:
        policies = by_coverage[ct]
        display_name = COVERAGE_SHORT.get(ct, ct)

        total_marketed = len(policies)
        quotes_received = len([p for p in policies if p["status"] in ("Quoted", "Bound", "Proposed")])
        proposed_count = len([p for p in policies if p["status"] in ("Bound", "Proposed")])

        # Determine overall status
        if any(p["status"] == "Bound" for p in policies):
            overall_status = "Bound"
        elif any(p["status"] == "Proposed" for p in policies):
            overall_status = "Proposed"
        elif any(p["status"] == "Quoted" for p in policies):
            overall_status = "Quoted"
        else:
            overall_status = "Pending"

        # Check if included in another coverage
        is_included = False
        for p in policies:
            comments = str(p.get("comments", "")).lower()
            if "included" in comments:
                is_included = True
                break

        if is_included:
            rows.append([display_name, "—", "—", f"Included in GL", overall_status])
        else:
            proposed_str = str(proposed_count) if proposed_count > 0 else "0"
            rows.append([display_name, str(total_marketed), str(quotes_received), proposed_str, overall_status])

    return rows


# ══════════════════════════════════════════════════════════════════════════
# COVERAGE SUMMARY AT A GLANCE
# ══════════════════════════════════════════════════════════════════════════

def build_coverage_summary(by_coverage):
    """Build coverage summary at a glance rows."""
    rows = []
    sorted_coverages = sorted(by_coverage.keys(), key=_coverage_sort_key)

    for ct in sorted_coverages:
        policies = by_coverage[ct]
        display_name = COVERAGE_SHORT.get(ct, ct)

        expiring_carriers = [p["carrier"] for p in policies if p["status"] == "Incumbent"]
        bound_carriers = [p["carrier"] for p in policies if p["status"] in ("Bound", "Proposed")]
        quoted_carriers = [p["carrier"] for p in policies if p["status"] == "Quoted"]
        pending_carriers = [p["carrier"] for p in policies if p["status"] in ("Market", "Submit", "Pending")]

        # Check if included
        is_included = any("included" in str(p.get("comments", "")).lower() for p in policies)

        exp_str = ", ".join(set(expiring_carriers)) if expiring_carriers else "—"
        if is_included:
            bound_str = "Included in GL"
        else:
            bound_str = ", ".join(set(bound_carriers)) if bound_carriers else "—"
        quoted_str = ", ".join(set(quoted_carriers)) if quoted_carriers else "—"
        pending_str = ", ".join(set(pending_carriers)) if pending_carriers else "—"

        rows.append([display_name, exp_str, bound_str, quoted_str, pending_str])

    return rows


# ══════════════════════════════════════════════════════════════════════════
# MAIN DOCUMENT GENERATOR
# ══════════════════════════════════════════════════════════════════════════

def generate_marketing_update_docx(
    opp_fields: dict,
    parsed_policies: list,
    by_coverage: dict,
    client_name: str,
    is_internal: bool = True,
) -> str:
    """
    Generate the Marketing Update DOCX document.

    Args:
        opp_fields: Opportunity record fields (or empty dict)
        parsed_policies: List of parsed policy dicts
        by_coverage: Policies grouped by coverage type
        client_name: Client/opportunity name
        is_internal: If True, include commission, revenue, rates, broker info

    Returns:
        Path to the generated DOCX file.
    """
    doc = Document()

    for section in doc.sections:
        section.left_margin = Inches(0.6)
        section.right_margin = Inches(0.6)
        section.top_margin = Inches(1.2)
        section.bottom_margin = Inches(0.6)

    add_page_header(doc)

    # ── Determine client info ──
    opp_name = ""
    effective_date = ""
    corporate_name = ""
    if opp_fields:
        opp_name = opp_fields.get("Opportunity Name", "")
        effective_date_raw = opp_fields.get("Effective Date", "")
        if effective_date_raw:
            try:
                dt = datetime.strptime(str(effective_date_raw), "%Y-%m-%d")
                effective_date = dt.strftime("%m/%d/%Y")
            except (ValueError, TypeError):
                effective_date = str(effective_date_raw)
        corporate_name = opp_fields.get("Opportunity Corporate Name", "")
        if not corporate_name:
            cn = opp_fields.get("Corporate Name", "")
            if isinstance(cn, list):
                corporate_name = cn[0] if cn else ""
            else:
                corporate_name = str(cn) if cn else ""

    display_name = corporate_name or client_name
    version_label = "Internal" if is_internal else "Client"

    # ════════════════════════════════════════════════════════════════
    # PAGE 1: TITLE & PREMIUM COMPARISON
    # ════════════════════════════════════════════════════════════════

    add_section_header(doc, "Insurance Marketing Update")

    add_formatted_paragraph(doc, f"Prepared for: {display_name}",
                           size=13, color=CLASSIC_BLUE, bold=True, space_before=4, space_after=2)
    if effective_date:
        add_formatted_paragraph(doc, f"Effective Date: {effective_date}  |  Package Program",
                               size=11, color=ELECTRIC_BLUE, bold=False, space_before=0, space_after=2)
    add_formatted_paragraph(doc, f"Report Date: {datetime.now().strftime('%m/%d/%Y')}",
                           size=10, color=CHARCOAL, bold=False, space_before=0, space_after=4)

    add_divider(doc)

    # ── Premium Comparison ──
    add_subsection_header(doc, "Premium Comparison — Expiring vs. Proposed")

    add_callout_box(doc, "Premiums shown include applicable taxes and fees. "
                        "TRIA/Terrorism premiums are not included in totals.")

    add_formatted_paragraph(doc, "", size=4, space_before=0, space_after=0)

    if is_internal:
        premium_headers = ["Coverage", "Carrier", "Expiring", "Proposed", "$ Change", "% Change", "Comm", "Revenue", "Broker"]
        premium_rows, total_change, total_pct, pending_coverages = build_premium_comparison_internal(by_coverage, parsed_policies)
        col_widths = [0.9, 1.1, 0.85, 0.85, 0.8, 0.65, 0.55, 0.7, 0.6]
    else:
        premium_headers = ["Coverage", "Carrier", "Expiring", "Proposed", "$ Change", "% Change"]
        premium_rows, total_change, total_pct, pending_coverages = build_premium_comparison(by_coverage, parsed_policies)
        col_widths = [1.2, 1.5, 1.1, 1.1, 1.1, 1.0]

    create_premium_summary_table(
        doc, premium_headers, premium_rows,
        col_widths=col_widths,
        highlight_last_row=True,
    )

    add_formatted_paragraph(doc, "", size=4, space_before=0, space_after=0)

    # Summary notes
    notes = []
    if total_change != 0:
        change_word = "decrease" if total_change < 0 else "increase"
        notes.append((
            f"Estimated total premium change: {'-' if total_change < 0 else '+'}${abs(total_change):,.2f} "
            f"({total_pct:+.1f}%) based on coverages quoted to date.",
            False, CLASSIC_BLUE
        ))
    if pending_coverages:
        notes.append((
            f"Note: {', '.join(pending_coverages)} premiums are pending and not included in proposed total.",
            False, CHARCOAL
        ))

    if notes:
        add_rich_callout_box(doc, notes, size=10, shading_hex=LIGHT_BLUE_HEX)

    # ════════════════════════════════════════════════════════════════
    # PAGE 2: MARKET ACTIVITY & KEY HIGHLIGHTS
    # ════════════════════════════════════════════════════════════════

    add_page_break(doc)
    add_subsection_header(doc, "Market Activity Overview")

    add_formatted_paragraph(doc,
        "The following summarizes our marketing efforts across all coverage lines. "
        "Our team has engaged multiple carriers to ensure competitive pricing and comprehensive coverage options.",
        size=10, color=CLASSIC_BLUE, space_before=2, space_after=8)

    activity_headers = ["Coverage", "Carriers Marketed", "Quotes Received", "Proposed", "Status"]
    activity_rows = build_market_activity(by_coverage)

    create_premium_summary_table(
        doc, activity_headers, activity_rows,
        col_widths=[1.5, 1.5, 1.3, 1.2, 1.5],
    )

    # Key Highlights section - auto-generated from data
    add_formatted_paragraph(doc, "", size=6, space_before=0, space_after=0)
    add_subsection_header(doc, "Key Highlights & Recommendations")

    highlights = _generate_highlights(by_coverage, parsed_policies)
    for title, description in highlights:
        hl_table = doc.add_table(rows=1, cols=1)
        hl_cell = hl_table.rows[0].cells[0]
        set_cell_shading(hl_cell, EGGSHELL_HEX)
        p_title = hl_cell.paragraphs[0]
        p_title.paragraph_format.space_before = Pt(6)
        p_title.paragraph_format.space_after = Pt(2)
        p_title.paragraph_format.left_indent = Inches(0.1)
        run_t = p_title.add_run(title)
        run_t.font.size = Pt(10)
        run_t.font.color.rgb = ELECTRIC_BLUE
        run_t.font.bold = True
        run_t.font.name = "Calibri"
        p_desc = hl_cell.add_paragraph()
        p_desc.paragraph_format.space_before = Pt(2)
        p_desc.paragraph_format.space_after = Pt(6)
        p_desc.paragraph_format.left_indent = Inches(0.1)
        run_d = p_desc.add_run(description)
        run_d.font.size = Pt(9)
        run_d.font.color.rgb = CLASSIC_BLUE
        run_d.font.name = "Calibri"
        add_formatted_paragraph(doc, "", size=3, space_before=0, space_after=0)

    # ════════════════════════════════════════════════════════════════
    # CARRIER COMPARISON TABLES
    # ════════════════════════════════════════════════════════════════

    add_page_break(doc)
    add_section_header(doc, "Carrier Comparisons")

    color_legend = "Side-by-side carrier comparisons for each coverage line. "
    color_legend += "Expiring carriers shown in gray, proposed/bound in green, quoted in white, and pending in yellow."
    add_callout_box(doc, color_legend)

    add_formatted_paragraph(doc, "", size=6, space_before=0, space_after=0)

    sorted_coverages = sorted(by_coverage.keys(), key=_coverage_sort_key)
    first_on_page = True

    for ct in sorted_coverages:
        policies = by_coverage[ct]
        if not policies:
            continue

        display_title = COVERAGE_DISPLAY_NAMES.get(ct, ct)
        short_title = COVERAGE_SHORT.get(ct, ct)

        # Build carrier data for comparison table
        builder_func, metrics_func = COVERAGE_BUILDERS.get(ct, (_build_generic_carrier, _get_generic_metrics))

        carriers_data = []
        for p in policies:
            carrier_dict = builder_func(p, is_internal=is_internal)
            carriers_data.append(carrier_dict)

        if not carriers_data:
            continue

        metrics = metrics_func(carriers_data, is_internal=is_internal)

        # Page break management - avoid too many tables on one page
        if not first_on_page and len(carriers_data) > 2:
            add_page_break(doc)
            first_on_page = True

        add_subsection_header(doc, display_title)
        create_carrier_comparison_table(doc, short_title, metrics, carriers_data)
        add_formatted_paragraph(doc, "", size=8, space_before=0, space_after=0)
        first_on_page = False

    # ════════════════════════════════════════════════════════════════
    # NEXT STEPS & COVERAGE SUMMARY
    # ════════════════════════════════════════════════════════════════

    add_page_break(doc)
    add_subsection_header(doc, "Next Steps")

    next_steps = _generate_next_steps(by_coverage, parsed_policies)
    for i, step in enumerate(next_steps):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = Pt(14)
        p.paragraph_format.left_indent = Inches(0.3)
        p.paragraph_format.first_line_indent = Inches(-0.3)
        run_num = p.add_run(f"{i+1}.  ")
        run_num.font.size = Pt(10)
        run_num.font.color.rgb = ELECTRIC_BLUE
        run_num.font.bold = True
        run_num.font.name = "Calibri"
        run_text = p.add_run(step)
        run_text.font.size = Pt(10)
        run_text.font.color.rgb = CLASSIC_BLUE
        run_text.font.name = "Calibri"

    # Coverage Summary at a Glance
    add_formatted_paragraph(doc, "", size=8, space_before=0, space_after=0)
    add_subsection_header(doc, "Coverage Summary at a Glance")

    summary_headers = ["Coverage", "Expiring", "Proposed / Bound", "Quoted", "Pending"]
    summary_rows = build_coverage_summary(by_coverage)

    create_premium_summary_table(
        doc, summary_headers, summary_rows,
        col_widths=[1.5, 1.5, 1.5, 1.5, 1.5],
    )

    # ── Contact & Disclaimer ──
    add_formatted_paragraph(doc, "", size=8, space_before=0, space_after=0)
    add_divider(doc)

    add_formatted_paragraph(doc, "Your HUB Hotel Franchise Team",
                           size=12, color=ELECTRIC_BLUE, bold=True, space_before=4, space_after=6)

    contact_table = doc.add_table(rows=3, cols=2)
    contact_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    remove_table_borders(contact_table)

    contacts = [
        ("Stefan Burkey", "Hotel Franchise Practice Leader  |  stefan.burkey@hubinternational.com"),
        ("Maureen Harvey", "Account Executive  |  maureen.harvey@hubinternational.com"),
        ("Sheena Callazo", "Claims Advocate  |  sheena.callazo@hubinternational.com"),
    ]

    for r_idx, (name, role) in enumerate(contacts):
        cell_name = contact_table.rows[r_idx].cells[0]
        cell_name.text = ""
        p_n = cell_name.paragraphs[0]
        p_n.paragraph_format.space_before = Pt(2)
        p_n.paragraph_format.space_after = Pt(2)
        run_n = p_n.add_run(name)
        run_n.font.size = Pt(9)
        run_n.font.color.rgb = CLASSIC_BLUE
        run_n.font.bold = True
        run_n.font.name = "Calibri"
        set_cell_width(cell_name, 1.8)
        remove_cell_borders(cell_name)

        cell_role = contact_table.rows[r_idx].cells[1]
        cell_role.text = ""
        p_r = cell_role.paragraphs[0]
        p_r.paragraph_format.space_before = Pt(2)
        p_r.paragraph_format.space_after = Pt(2)
        run_r = p_r.add_run(role)
        run_r.font.size = Pt(9)
        run_r.font.color.rgb = CHARCOAL
        run_r.font.name = "Calibri"
        set_cell_width(cell_role, 5.7)
        remove_cell_borders(cell_role)

    add_formatted_paragraph(doc, "", size=6, space_before=0, space_after=0)
    add_callout_box(doc, "This marketing update is for informational purposes only and does not constitute a binder of insurance. "
                        "Actual coverage terms, conditions, and exclusions are governed by the policies as issued. "
                        "Please review all policies carefully upon receipt.", size=8)

    # ── Save ──
    version_suffix = "Internal" if is_internal else "Client"
    safe_name = "".join(c for c in display_name if c.isalnum() or c in " _-").strip().replace(" ", "_")
    filename = f"Marketing_Update_{safe_name}_{version_suffix}.docx"
    output_path = os.path.join(tempfile.gettempdir(), filename)
    doc.save(output_path)
    logger.info(f"Marketing Update DOCX saved to {output_path}")
    return output_path


# ══════════════════════════════════════════════════════════════════════════
# AUTO-GENERATED HIGHLIGHTS & NEXT STEPS
# ══════════════════════════════════════════════════════════════════════════

def _generate_highlights(by_coverage, parsed_policies):
    """Auto-generate key highlights from the policy data."""
    highlights = []

    for ct in sorted(by_coverage.keys(), key=_coverage_sort_key):
        policies = by_coverage[ct]
        display_name = COVERAGE_SHORT.get(ct, ct)

        expiring = [p for p in policies if p["status"] == "Incumbent"]
        bound = [p for p in policies if p["status"] in ("Bound", "Proposed")]
        quoted = [p for p in policies if p["status"] == "Quoted"]

        if bound and expiring:
            exp_prem = sum(p["premium_tx"] for p in expiring)
            bound_prem = sum(p["premium_tx"] for p in bound)
            if exp_prem > 0 and bound_prem > 0:
                change = bound_prem - exp_prem
                pct = (change / exp_prem) * 100
                carrier = bound[0]["carrier"]
                if change < 0:
                    highlights.append((
                        f"{display_name} Premium Reduction",
                        f"{carrier} renewal at {_safe_currency(bound_prem)}, "
                        f"a {abs(pct):.1f}% decrease from expiring premium of {_safe_currency(exp_prem)}."
                    ))
                elif change > 0:
                    highlights.append((
                        f"{display_name} Premium Increase",
                        f"{carrier} renewal at {_safe_currency(bound_prem)}, "
                        f"a {pct:.1f}% increase from expiring premium of {_safe_currency(exp_prem)}."
                    ))

        # Note competitive quotes
        if quoted and len(quoted) >= 2:
            carriers = [q["carrier"] for q in quoted]
            highlights.append((
                f"{display_name} — Multiple Quotes",
                f"Received competitive quotes from {', '.join(set(carriers))}."
            ))

    # Check for pending coverages
    pending = []
    for ct in by_coverage:
        policies = by_coverage[ct]
        if not any(p["status"] in ("Bound", "Proposed", "Quoted") for p in policies):
            pending.append(COVERAGE_SHORT.get(ct, ct))

    if pending:
        highlights.append((
            "Pending Coverages",
            f"Awaiting quotes for: {', '.join(pending)}. Will update upon receipt."
        ))

    if not highlights:
        highlights.append((
            "Marketing in Progress",
            "Our team is actively marketing all coverage lines to ensure competitive pricing and comprehensive coverage."
        ))

    return highlights


def _generate_next_steps(by_coverage, parsed_policies):
    """Auto-generate next steps from the policy data."""
    steps = []

    # Check for pending/market status policies
    for ct in sorted(by_coverage.keys(), key=_coverage_sort_key):
        policies = by_coverage[ct]
        display_name = COVERAGE_SHORT.get(ct, ct)

        pending = [p for p in policies if p["status"] in ("Market", "Submit")]
        if pending:
            carriers = list(set(p["carrier"] for p in pending))
            steps.append(
                f"Awaiting {display_name} quotes from {', '.join(carriers)}."
            )

    # Check for quoted but not yet proposed
    for ct in sorted(by_coverage.keys(), key=_coverage_sort_key):
        policies = by_coverage[ct]
        display_name = COVERAGE_SHORT.get(ct, ct)

        quoted = [p for p in policies if p["status"] == "Quoted"]
        bound = [p for p in policies if p["status"] in ("Bound", "Proposed")]
        if quoted and not bound:
            steps.append(
                f"Finalizing {display_name} recommendation based on received quotes."
            )

    steps.append(
        "Full insurance proposal with all coverage sections, forms, and endorsements "
        "will be provided once all outstanding quotes are received."
    )

    return steps


# ══════════════════════════════════════════════════════════════════════════
# MAIN ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════

async def generate_marketing_update(client_name: str, is_internal: bool = True) -> str:
    """
    Main entry point: generate a Marketing Update DOCX for a client.

    Args:
        client_name: Client name to search for in Airtable.
        is_internal: If True, generate internal version with all financial details.

    Returns:
        Path to the generated DOCX file, or error message string.
    """
    logger.info(f"generate_marketing_update called for: '{client_name}' (internal={is_internal})")

    opp_fields, policies = await resolve_client_data(client_name)

    if not policies:
        return f"No policies found for '{client_name}'."

    parsed = parse_policies(policies)
    by_coverage = group_by_coverage(parsed)

    if not opp_fields:
        opp_fields = {}

    try:
        output_path = generate_marketing_update_docx(
            opp_fields=opp_fields,
            parsed_policies=parsed,
            by_coverage=by_coverage,
            client_name=client_name,
            is_internal=is_internal,
        )
        return output_path
    except Exception as e:
        logger.error(f"Error generating marketing update DOCX: {e}", exc_info=True)
        return f"Error generating document: {str(e)}"
