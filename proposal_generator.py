#!/usr/bin/env python3
"""
Hotel Insurance Proposal - DOCX Generator
Generates complete branded DOCX proposals following HUB International design system.
23-section document with full compliance pages.
"""

import os
import logging
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import datetime

logger = logging.getLogger(__name__)

# HUB Design System Colors
ELECTRIC_BLUE = RGBColor(0x16, 0x7B, 0xD4)  # #167BD4
CLASSIC_BLUE = RGBColor(0x26, 0x38, 0x45)   # #263845
ARCTIC_GRAY = RGBColor(0xB8, 0xC4, 0xCE)    # #B8C4CE
EGGSHELL = RGBColor(0xF3, 0xF5, 0xF1)       # #F3F5F1
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
CHARCOAL = RGBColor(0x4A, 0x4A, 0x4A)

ELECTRIC_BLUE_HEX = "167BD4"
CLASSIC_BLUE_HEX = "263845"
ARCTIC_GRAY_HEX = "B8C4CE"
EGGSHELL_HEX = "F3F5F1"

# Logo path
LOGO_PATH = os.path.join(os.path.dirname(__file__), "templates", "hub_logo_horizontal.png")

# Default Service Team
SERVICE_TEAM = [
    {
        "role": "Hotel Franchise Practice Leader",
        "name": "Stefan Burkey",
        "phone": "O: 407-636-8133\nM: 407-782-1900",
        "email": "stefan.burkey@hubinternational.com"
    },
    {
        "role": "Account Executive",
        "name": "Maureen Harvey, CIC, CRM",
        "phone": "O: 407-893-3830\nF: 407-831-3063",
        "email": ""
    },
    {
        "role": "Senior Franchise Claims Advocate",
        "name": "Sheena Callazo, RPLU",
        "phone": "O: 630-468-5674",
        "email": "sheena.callazo@hubinternational.com"
    }
]

OFFICE_LOCATIONS = [
    "HUB International Midwest Limited — 203 N LaSalle, Suite 2000, Chicago, IL 60601",
    "HUB International Midwest Limited — 1411 Opus Place, Suite 450, Downers Grove, IL 60515"
]

# California Licenses
CA_LICENSES = [
    ("Agency Two Insurance Marketing Group, LLC, d/b/a AgencyOne", "0H44808"),
    ("All World Insurance Services, Inc.", "0F69702"),
    ("Avant Brokerage LLC", "0I77138"),
    ("Avant Specialty Claims, LLC", "6003211"),
    ("Avant Underwriters, LLC", "0G67877"),
    ("Brokers' Service Marketing Group II LLC", "0E02001"),
    ("Business Underwriters Associates Agency Inc.", "0C26183"),
    ("Chun-Ha Insurance Services, Inc.", "0F71901"),
    ("Dale Barton Agency", "0137389"),
    ("FNA Insurance Services Inc", "0I72746"),
    ("HUB Heartland, LLC", "0H15020"),
    ("HUB International Insurance Services Inc.", "0757776"),
    ("HUB International Iowa, LLC", "0K02887"),
    ("HUB International Mid-Atlantic Inc.", "0D58520"),
    ("HUB International Midwest Limited (MWW)", "0D08495"),
    ("HUB International Mountain States Limited", "0A96371"),
    ("HUB International New England, LLC", "0F79381"),
    ("HUB International Northeast Limited", "0E16962"),
    ("HUB International Northwest LLC", "0D08450"),
    ("HUB International Texas, Inc.", "0E24644"),
    ("HUB International Transportation Insurance Services Inc", "0D43442"),
    ("ISR Marine Insurance Services LLC", "0I67807"),
    ("Program Brokerage Corporation", "0B27851"),
    ("Sadler & Company, Inc.", "0B57651"),
    ("SBR Services, LLC", "6007384"),
    ("Silverstone Group, LLC", "0D79635"),
    ("Specialty Program Group LLC", "0L09546"),
    ("Squaremouth LLC", "0H58357"),
    ("The HDH Group, Inc.", "0C66771"),
    ("VIU by HUB", "6010202"),
]


# ─── Helper Functions ─────────────────────────────────────────

def set_cell_shading(cell, color_hex):
    """Set cell background color."""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}" w:val="clear"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_border(cell, **kwargs):
    """Set cell borders. kwargs: top, bottom, left, right with values like '1pt'."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}></w:tcBorders>')
    for edge, val in kwargs.items():
        element = parse_xml(
            f'<w:{edge} {nsdecls("w")} w:val="single" w:sz="4" w:space="0" w:color="{ARCTIC_GRAY_HEX}"/>'
        )
        tcBorders.append(element)
    tcPr.append(tcBorders)


def add_formatted_paragraph(doc, text, size=20, color=CLASSIC_BLUE, bold=False,
                            alignment=WD_ALIGN_PARAGRAPH.LEFT, space_before=0, space_after=0):
    """Add a formatted paragraph to the document."""
    p = doc.add_paragraph()
    p.alignment = alignment
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.name = "Calibri"
    return p


def add_section_header(doc, text):
    """Add a 32pt Classic Blue bold section header."""
    return add_formatted_paragraph(doc, text, size=32, color=CLASSIC_BLUE, bold=True,
                                   space_before=30, space_after=20)


def add_subsection_header(doc, text):
    """Add a 24pt Electric Blue bold subsection header."""
    return add_formatted_paragraph(doc, text, size=24, color=ELECTRIC_BLUE, bold=True,
                                   space_before=20, space_after=15)


def create_styled_table(doc, headers, rows, col_widths=None):
    """Create a table with HUB styling: Electric Blue header, alternating rows."""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Style header row
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(header)
        run.font.size = Pt(20)
        run.font.color.rgb = WHITE
        run.font.bold = True
        run.font.name = "Calibri"
        set_cell_shading(cell, ELECTRIC_BLUE_HEX)
    
    # Style data rows
    for row_idx, row_data in enumerate(rows):
        for col_idx, cell_text in enumerate(row_data):
            cell = table.rows[row_idx + 1].cells[col_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            run = p.add_run(str(cell_text))
            run.font.size = Pt(18)
            run.font.color.rgb = CLASSIC_BLUE
            run.font.name = "Calibri"
            # Alternating row colors
            if row_idx % 2 == 1:
                set_cell_shading(cell, EGGSHELL_HEX)
    
    # Set column widths if provided
    if col_widths:
        for row in table.rows:
            for i, width in enumerate(col_widths):
                if i < len(row.cells):
                    row.cells[i].width = Inches(width)
    
    return table


def add_page_header(doc):
    """Add page header with logo left and text right."""
    section = doc.sections[-1]
    header = section.header
    header.is_linked_to_previous = False
    
    # Create a table for the header layout
    htable = header.add_table(1, 2, width=Inches(7))
    htable.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Logo cell (left)
    logo_cell = htable.rows[0].cells[0]
    logo_cell.width = Inches(2.5)
    if os.path.exists(LOGO_PATH):
        p = logo_cell.paragraphs[0]
        run = p.add_run()
        run.add_picture(LOGO_PATH, width=Inches(2.2))
    
    # Text cell (right)
    text_cell = htable.rows[0].cells[1]
    text_cell.width = Inches(4.5)
    p = text_cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = p.add_run("Franchise Division")
    run.font.size = Pt(18)
    run.font.color.rgb = ELECTRIC_BLUE
    run.font.bold = True
    run.font.name = "Calibri"
    p2 = text_cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run2 = p2.add_run("Hotel Insurance Programs")
    run2.font.size = Pt(18)
    run2.font.color.rgb = ELECTRIC_BLUE
    run2.font.bold = True
    run2.font.name = "Calibri"
    
    # Remove borders from header table
    for row in htable.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcBorders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:top w:val="none" w:sz="0" w:space="0"/>'
                f'<w:left w:val="none" w:sz="0" w:space="0"/>'
                f'<w:bottom w:val="none" w:sz="0" w:space="0"/>'
                f'<w:right w:val="none" w:sz="0" w:space="0"/>'
                f'</w:tcBorders>'
            )
            tcPr.append(tcBorders)


def add_callout_box(doc, text, size=18):
    """Add an eggshell background callout/disclaimer box."""
    table = doc.add_table(rows=1, cols=1)
    cell = table.rows[0].cells[0]
    set_cell_shading(cell, EGGSHELL_HEX)
    p = cell.paragraphs[0]
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.color.rgb = CLASSIC_BLUE
    run.font.name = "Calibri"
    return table


def add_page_break(doc):
    """Add a page break."""
    doc.add_page_break()


def fmt_currency(amount):
    """Format a number as currency."""
    if isinstance(amount, (int, float)):
        return f"${amount:,.0f}"
    if isinstance(amount, str):
        if amount.startswith("$"):
            return amount
        try:
            return f"${float(amount.replace(',', '')):,.0f}"
        except (ValueError, AttributeError):
            return amount
    return str(amount)


# ─── Section Generators ───────────────────────────────────────

def generate_cover_page(doc, data):
    """Section 1: Cover Page"""
    ci = data.get("client_info", {})
    client_name = ci.get("named_insured", "Client Name")
    dba = ci.get("dba", "")
    effective_date = ci.get("effective_date", "")
    proposal_date = datetime.date.today().strftime("%B %d, %Y")
    
    # Center everything
    # Logo
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    if os.path.exists(LOGO_PATH):
        run = p.add_run()
        run.add_picture(LOGO_PATH, width=Inches(3))
    
    # Electric Blue line
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(20)
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="48" w:space="1" w:color="{ELECTRIC_BLUE_HEX}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)
    
    # Title
    add_formatted_paragraph(doc, "Commercial Insurance", size=48, color=CLASSIC_BLUE,
                           bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=30)
    add_formatted_paragraph(doc, "Proposal", size=48, color=CLASSIC_BLUE,
                           bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=30)
    
    # Prepared For box (using a table with eggshell background)
    box_table = doc.add_table(rows=1, cols=1)
    box_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    # Set fixed width for centering
    for row in box_table.rows:
        for cell in row.cells:
            cell.width = Inches(5.5)
    # Set table preferred width
    tbl = box_table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="7920" w:type="dxa"/>')  # 5.5 inches = 7920 twips
    # Remove existing tblW if any
    existing_tblW = tblPr.find(qn('w:tblW'))
    if existing_tblW is not None:
        tblPr.remove(existing_tblW)
    tblPr.append(tblW)
    cell = box_table.rows[0].cells[0]
    set_cell_shading(cell, EGGSHELL_HEX)
    
    # Top blue border
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="32" w:space="0" w:color="{ELECTRIC_BLUE_HEX}"/>'
        f'<w:bottom w:val="single" w:sz="32" w:space="0" w:color="{ELECTRIC_BLUE_HEX}"/>'
        f'<w:left w:val="none" w:sz="0" w:space="0"/>'
        f'<w:right w:val="none" w:sz="0" w:space="0"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)
    
    # "Prepared For" text
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(15)
    run = p.add_run("Prepared For")
    run.font.size = Pt(20)
    run.font.color.rgb = CHARCOAL
    run.font.name = "Calibri"
    
    # Client name
    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(10)
    run2 = p2.add_run(client_name)
    run2.font.size = Pt(44)
    run2.font.color.rgb = ELECTRIC_BLUE
    run2.font.bold = True
    run2.font.name = "Calibri"
    
    # DBA if present
    if dba:
        p3 = cell.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run3 = p3.add_run(f"DBA: {dba}")
        run3.font.size = Pt(26)
        run3.font.color.rgb = CLASSIC_BLUE
        run3.font.name = "Calibri"
        p3.paragraph_format.space_after = Pt(15)
    else:
        p2.paragraph_format.space_after = Pt(15)
    
    # Dates - two column table
    date_table = doc.add_table(rows=1, cols=2)
    date_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Proposal Date
    dc1 = date_table.rows[0].cells[0]
    p = dc1.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Proposal Date")
    run.font.size = Pt(16)
    run.font.color.rgb = CLASSIC_BLUE
    run.font.bold = True
    run.font.name = "Calibri"
    p2 = dc1.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run(proposal_date)
    run2.font.size = Pt(16)
    run2.font.color.rgb = CLASSIC_BLUE
    run2.font.name = "Calibri"
    
    # Effective Date
    dc2 = date_table.rows[0].cells[1]
    p = dc2.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Effective Date")
    run.font.size = Pt(16)
    run.font.color.rgb = CLASSIC_BLUE
    run.font.bold = True
    run.font.name = "Calibri"
    p2 = dc2.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run2 = p2.add_run(effective_date)
    run2.font.size = Pt(16)
    run2.font.color.rgb = CLASSIC_BLUE
    run2.font.name = "Calibri"
    
    # Remove borders from date table
    for row in date_table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcBorders = parse_xml(
                f'<w:tcBorders {nsdecls("w")}>'
                f'<w:top w:val="none" w:sz="0" w:space="0"/>'
                f'<w:left w:val="none" w:sz="0" w:space="0"/>'
                f'<w:bottom w:val="none" w:sz="0" w:space="0"/>'
                f'<w:right w:val="none" w:sz="0" w:space="0"/>'
                f'</w:tcBorders>'
            )
            tcPr.append(tcBorders)
    
    # Gray line
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(20)
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="16" w:space="1" w:color="{ARCTIC_GRAY_HEX}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)
    
    # Presented By
    add_formatted_paragraph(doc, "Presented By", size=18, color=CLASSIC_BLUE,
                           alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=15)
    add_formatted_paragraph(doc, "HUB International Midwest Limited", size=18, color=CLASSIC_BLUE,
                           bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_formatted_paragraph(doc, "Franchise Division | Hotel Insurance Programs", size=16,
                           color=ELECTRIC_BLUE, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=20)


def generate_service_team(doc, data):
    """Section 2: Service Team"""
    add_page_break(doc)
    add_page_header(doc)
    add_section_header(doc, "Your Service Team")
    
    headers = ["Role", "Name", "Phone", "Email"]
    rows = []
    for member in SERVICE_TEAM:
        rows.append([member["role"], member["name"], member["phone"], member["email"]])
    
    create_styled_table(doc, headers, rows, col_widths=[2.0, 2.0, 1.8, 2.7])
    
    # Office locations
    add_formatted_paragraph(doc, "", space_before=15)
    for loc in OFFICE_LOCATIONS:
        add_formatted_paragraph(doc, loc, size=16, color=CLASSIC_BLUE)


def generate_premium_summary(doc, data):
    """Section 3: Premium Summary - Expiring vs Proposed"""
    add_page_break(doc)
    add_section_header(doc, "Premium Summary")
    
    coverages = data.get("coverages", {})
    expiring = data.get("expiring_premiums", {})
    
    coverage_names = {
        "property": "Property",
        "general_liability": "General Liability",
        "umbrella": "Umbrella",
        "workers_comp": "Workers Compensation",
        "commercial_auto": "Commercial Auto"
    }
    
    headers = ["Coverage", "Carrier", "Expiring Premium", "Proposed Premium", "$ Change", "% Change"]
    rows = []
    total_expiring = 0
    total_proposed = 0
    
    for key, display_name in coverage_names.items():
        cov = coverages.get(key)
        if not cov:
            continue
        
        carrier = cov.get("carrier", "")
        admitted = cov.get("carrier_admitted", True)
        carrier_display = carrier
        if not admitted:
            carrier_display = f"{carrier}\n(Non-Admitted)"
        
        proposed = cov.get("total_premium", 0)
        exp = expiring.get(key, 0)
        
        dollar_change = proposed - exp if exp else "N/A"
        if exp and exp > 0:
            pct_change = ((proposed - exp) / exp) * 100
            pct_str = f"{pct_change:+.1f}%"
            dollar_str = f"{'+' if dollar_change >= 0 else ''}{fmt_currency(abs(dollar_change))}"
            if dollar_change < 0:
                dollar_str = f"-{fmt_currency(abs(dollar_change))}"
        else:
            pct_str = "New"
            dollar_str = "New"
        
        rows.append([
            display_name,
            carrier_display,
            fmt_currency(exp) if exp else "N/A",
            fmt_currency(proposed),
            dollar_str,
            pct_str
        ])
        
        total_expiring += exp if exp else 0
        total_proposed += proposed
    
    # Total row
    total_dollar = total_proposed - total_expiring
    if total_expiring > 0:
        total_pct = ((total_proposed - total_expiring) / total_expiring) * 100
        total_pct_str = f"{total_pct:+.1f}%"
        total_dollar_str = f"{'+' if total_dollar >= 0 else ''}{fmt_currency(abs(total_dollar))}"
        if total_dollar < 0:
            total_dollar_str = f"-{fmt_currency(abs(total_dollar))}"
    else:
        total_pct_str = "N/A"
        total_dollar_str = "N/A"
    
    rows.append([
        "TOTAL",
        "",
        fmt_currency(total_expiring) if total_expiring else "N/A",
        fmt_currency(total_proposed),
        total_dollar_str,
        total_pct_str
    ])
    
    table = create_styled_table(doc, headers, rows, col_widths=[1.5, 1.5, 1.2, 1.2, 1.0, 0.8])
    
    # Bold the total row
    last_row = table.rows[-1]
    for cell in last_row.cells:
        for p in cell.paragraphs:
            for run in p.runs:
                run.font.bold = True
    
    # Savings/increase callout
    if total_expiring > 0 and total_dollar != 0:
        direction = "savings" if total_dollar < 0 else "increase"
        add_formatted_paragraph(doc, "", space_before=10)
        callout_text = f"Total premium {direction}: {fmt_currency(abs(total_dollar))} ({abs(total_pct):.1f}%)"
        add_callout_box(doc, callout_text)


def generate_payment_options(doc, data):
    """Section 4: Payment Options"""
    add_page_break(doc)
    add_section_header(doc, "Payment Options")
    
    payment_opts = data.get("payment_options", [])
    if payment_opts:
        headers = ["Carrier", "Payment Terms", "Minimum Earned Premium"]
        rows = [[po.get("carrier", ""), po.get("terms", ""), po.get("mep", "")] for po in payment_opts]
        create_styled_table(doc, headers, rows)
    else:
        add_formatted_paragraph(doc, "Payment terms to be confirmed upon binding.")


def generate_subjectivities(doc, data):
    """Section 5: Subjectivities"""
    add_page_break(doc)
    add_section_header(doc, "Subjectivities")
    
    add_formatted_paragraph(doc, "The following items are required prior to or as a condition of binding:", size=20)
    
    coverages = data.get("coverages", {})
    coverage_names = {
        "property": "Property",
        "general_liability": "General Liability",
        "umbrella": "Umbrella",
        "workers_comp": "Workers Compensation",
        "commercial_auto": "Commercial Auto"
    }
    
    has_subjectivities = False
    for key, display_name in coverage_names.items():
        cov = coverages.get(key)
        if cov and cov.get("subjectivities"):
            has_subjectivities = True
            add_subsection_header(doc, display_name)
            for subj in cov["subjectivities"]:
                add_formatted_paragraph(doc, f"☐  {subj}", size=18)
    
    if not has_subjectivities:
        add_formatted_paragraph(doc, "No subjectivities noted. Please confirm with carrier.", size=18)


def generate_named_insureds(doc, data):
    """Section 6: Named Insureds"""
    add_page_break(doc)
    add_section_header(doc, "Named Insureds")
    
    named = data.get("named_insureds", [])
    if named:
        for i, ni in enumerate(named, 1):
            add_formatted_paragraph(doc, f"{i}. {ni}", size=18)
    else:
        add_formatted_paragraph(doc, "1. " + data.get("client_info", {}).get("named_insured", "TBD"), size=18)
    
    # Note box
    add_formatted_paragraph(doc, "", space_before=10)
    add_callout_box(doc, "Note: Additional named insureds may be added as required by franchise agreements or management contracts.")
    
    # Additional Interests
    interests = data.get("additional_interests", [])
    if interests:
        add_subsection_header(doc, "Additional Interests")
        headers = ["Type", "Name & Address", "Description"]
        rows = [[ai.get("type", ""), ai.get("name_address", ""), ai.get("description", "")] for ai in interests]
        create_styled_table(doc, headers, rows)


def generate_information_summary(doc, data):
    """Section 7: Information Summary"""
    add_page_break(doc)
    add_section_header(doc, "Information Summary")
    
    ci = data.get("client_info", {})
    
    headers = ["Item", "Details"]
    rows = [
        ["First Named Insured", ci.get("named_insured", "")],
        ["Mailing Address", ci.get("address", "")],
        ["Entity Type", ci.get("entity_type", "")],
        ["Effective Date", ci.get("effective_date", "")],
        ["Expiration Date", ci.get("expiration_date", "")],
    ]
    if ci.get("sales_exposure_basis"):
        rows.append(["Proposed Sales/Exposure Basis", ci["sales_exposure_basis"]])
    if ci.get("dba"):
        rows.insert(1, ["DBA", ci["dba"]])
    
    create_styled_table(doc, headers, rows, col_widths=[2.5, 5.0])
    
    add_formatted_paragraph(doc, "", space_before=10)
    add_callout_box(doc, "The information contained in this proposal is based on data provided by the insured and/or their representatives. HUB International makes no warranty as to the accuracy of this information.")


def generate_locations(doc, data):
    """Section 8: Locations"""
    add_page_break(doc)
    add_section_header(doc, "Schedule of Locations")
    
    locations = data.get("locations", [])
    if locations:
        has_entity = any(loc.get("corporate_entity") for loc in locations)
        if has_entity:
            headers = ["#", "Corporate Entity", "Address", "City", "State", "ZIP", "Description"]
            rows = [[
                loc.get("number", ""),
                loc.get("corporate_entity", ""),
                loc.get("address", ""),
                loc.get("city", ""),
                loc.get("state", ""),
                loc.get("zip", ""),
                loc.get("description", "")
            ] for loc in locations]
        else:
            headers = ["#", "Address", "City", "State", "ZIP", "Description"]
            rows = [[
                loc.get("number", ""),
                loc.get("address", ""),
                loc.get("city", ""),
                loc.get("state", ""),
                loc.get("zip", ""),
                loc.get("description", "")
            ] for loc in locations]
        create_styled_table(doc, headers, rows)
    else:
        add_formatted_paragraph(doc, "Location schedule to be confirmed.", size=18)


def generate_coverage_section(doc, data, coverage_key, display_name):
    """Generate a coverage section (Property, GL, Umbrella, WC, Auto)."""
    coverages = data.get("coverages", {})
    cov = coverages.get(coverage_key)
    if not cov:
        return
    
    add_page_break(doc)
    add_section_header(doc, display_name)
    
    # Carrier Information
    carrier = cov.get("carrier", "N/A")
    admitted = "Admitted" if cov.get("carrier_admitted", True) else "Non-Admitted"
    am_best = cov.get("am_best_rating", "N/A")
    
    carrier_rows = [
        ["Carrier", carrier],
        ["Admitted Status", admitted],
        ["AM Best Rating", am_best],
    ]
    # Add premium for non-property coverages
    if coverage_key != "property":
        carrier_rows.append(["Total Premium", fmt_currency(cov.get("total_premium", 0))])
    
    create_styled_table(doc, ["Item", "Details"], carrier_rows, col_widths=[2.5, 5.0])
    
    # Limits
    limits = cov.get("limits", [])
    if limits:
        add_subsection_header(doc, "Coverage Limits")
        headers = ["Description", "Limit"]
        rows = [[lim.get("description", ""), lim.get("limit", "")] for lim in limits]
        create_styled_table(doc, headers, rows, col_widths=[4.0, 3.5])
    
    # Deductibles (Property)
    deductibles = cov.get("deductibles", [])
    if deductibles:
        add_subsection_header(doc, "Deductibles")
        headers = ["Description", "Amount"]
        rows = [[ded.get("description", ""), ded.get("amount", "")] for ded in deductibles]
        create_styled_table(doc, headers, rows, col_widths=[4.0, 3.5])
    
    # Schedule of Hazards (GL)
    hazards = cov.get("schedule_of_hazards", [])
    if hazards:
        add_subsection_header(doc, "Schedule of Hazards")
        headers = ["Location", "Classification", "Code", "Basis", "Exposure"]
        rows = [[
            h.get("location", ""),
            h.get("classification", ""),
            h.get("code", ""),
            h.get("basis", ""),
            h.get("exposure", "")
        ] for h in hazards]
        create_styled_table(doc, headers, rows)
    
    # Rating Basis (WC)
    rating = cov.get("rating_basis", [])
    if rating:
        add_subsection_header(doc, "Rating Basis")
        headers = ["State", "Location", "Class Code", "Classification", "Payroll", "Rate"]
        rows = [[
            r.get("state", ""),
            r.get("location", ""),
            r.get("class_code", ""),
            r.get("classification", ""),
            r.get("payroll", ""),
            r.get("rate", "")
        ] for r in rating]
        create_styled_table(doc, headers, rows)
    
    # Vehicle Schedule (Auto)
    vehicles = cov.get("vehicle_schedule", [])
    if vehicles:
        add_subsection_header(doc, "Vehicle Schedule")
        headers = ["Year", "Make", "Model", "VIN", "Garage Location"]
        rows = [[v.get("year", ""), v.get("make", ""), v.get("model", ""),
                 v.get("vin", ""), v.get("garage_location", "")] for v in vehicles]
        create_styled_table(doc, headers, rows)
    
    # Additional Coverages
    addl = cov.get("additional_coverages", [])
    if addl:
        add_subsection_header(doc, "Additional Coverages")
        has_ded = any(ac.get("deductible") for ac in addl)
        if has_ded:
            headers = ["Description", "Limit", "Deductible"]
            rows = [[ac.get("description", ""), ac.get("limit", ""), ac.get("deductible", "")] for ac in addl]
        else:
            headers = ["Description", "Limit"]
            rows = [[ac.get("description", ""), ac.get("limit", "")] for ac in addl]
        create_styled_table(doc, headers, rows)
    
    # Underlying Insurance (Umbrella)
    underlying = cov.get("underlying_insurance", [])
    if underlying:
        add_subsection_header(doc, "Underlying Insurance")
        headers = ["Carrier", "Coverage", "Limits"]
        rows = [[u.get("carrier", ""), u.get("coverage", ""), u.get("limits", "")] for u in underlying]
        create_styled_table(doc, headers, rows)
    
    # Tower Structure (Umbrella)
    tower = cov.get("tower_structure", [])
    if tower:
        add_subsection_header(doc, "Umbrella Tower Structure")
        headers = ["Layer", "Carrier", "Limits", "Premium", "Total Cost (incl. taxes/fees)"]
        rows = [[
            t.get("layer", ""),
            t.get("carrier", ""),
            t.get("limits", ""),
            fmt_currency(t.get("premium", 0)),
            fmt_currency(t.get("total_cost", 0))
        ] for t in tower]
        create_styled_table(doc, headers, rows)
    
    # Forms & Endorsements
    forms = cov.get("forms_endorsements", [])
    if forms:
        add_subsection_header(doc, "Forms & Endorsements")
        headers = ["Form Number", "Description"]
        rows = [[f.get("form_number", ""), f.get("description", "")] for f in forms]
        create_styled_table(doc, headers, rows, col_widths=[2.5, 5.0])


def generate_confirmation_to_bind(doc, data):
    """Section 14: Confirmation to Bind Agreement"""
    add_page_break(doc)
    add_section_header(doc, "Confirmation to Bind Agreement")
    
    add_formatted_paragraph(doc,
        "By signing below, the undersigned authorized representative of the Applicant confirms "
        "the following statements and authorizes HUB International to bind the coverages as outlined "
        "in this proposal, subject to the terms and conditions of the respective policies.",
        size=18, space_after=15)
    
    # Application Statements
    add_subsection_header(doc, "Application Statements")
    
    statements = [
        "The information provided in the applications and supplemental materials is true, accurate, and complete to the best of my knowledge.",
        "I understand that any material misrepresentation or omission may void coverage or result in denial of claims.",
        "I have reviewed the proposed coverages, limits, deductibles, and premiums outlined in this proposal.",
        "I understand that the coverages described herein are subject to the terms, conditions, and exclusions of the actual policies issued.",
        "I authorize HUB International to bind the coverages as outlined in this proposal on behalf of the named insured(s).",
        "I understand that subjectivities, if any, must be satisfied within the timeframes specified or coverage may be subject to cancellation.",
        "I acknowledge that surplus lines placements, if any, are not covered by state guaranty funds.",
        "I have been offered Terrorism Risk Insurance Act (TRIA) coverage and have made my election as indicated in this proposal."
    ]
    
    for i, stmt in enumerate(statements, 1):
        add_formatted_paragraph(doc, f"{i}. {stmt}", size=18, space_after=8)
    
    # Signature block
    add_formatted_paragraph(doc, "", space_before=30)
    sig_table = doc.add_table(rows=5, cols=2)
    sig_table.alignment = WD_TABLE_ALIGNMENT.LEFT
    
    sig_fields = [
        ("Authorized Signature:", ""),
        ("Printed Name:", ""),
        ("Title:", ""),
        ("Date:", ""),
        ("Company:", "")
    ]
    
    for i, (label, val) in enumerate(sig_fields):
        cell_label = sig_table.rows[i].cells[0]
        p = cell_label.paragraphs[0]
        run = p.add_run(label)
        run.font.size = Pt(18)
        run.font.color.rgb = CLASSIC_BLUE
        run.font.bold = True
        run.font.name = "Calibri"
        
        cell_val = sig_table.rows[i].cells[1]
        p = cell_val.paragraphs[0]
        # Add underline for signature line
        pPr = p._p.get_or_add_pPr()
        pBdr = parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'<w:bottom w:val="single" w:sz="4" w:space="1" w:color="{CLASSIC_BLUE_HEX}"/>'
            f'</w:pBdr>'
        )
        pPr.append(pBdr)


def generate_electronic_consent(doc):
    """Section 15: Electronic Documents Consent"""
    add_page_break(doc)
    add_section_header(doc, "Electronic Documents Consent")
    
    add_formatted_paragraph(doc,
        "By signing below, I consent to receive insurance policy documents, endorsements, "
        "certificates of insurance, notices of cancellation, renewal notices, and other "
        "insurance-related documents electronically from HUB International and/or the "
        "insurance carriers providing coverage.",
        size=18, space_after=15)
    
    add_formatted_paragraph(doc,
        "I understand that I may withdraw this consent at any time by providing written "
        "notice to my HUB International representative, and that I may request paper "
        "copies of any documents at no additional charge.",
        size=18, space_after=20)
    
    # Signature line
    sig_table = doc.add_table(rows=3, cols=2)
    for i, label in enumerate(["Authorized Signature:", "Printed Name:", "Date:"]):
        cell = sig_table.rows[i].cells[0]
        p = cell.paragraphs[0]
        run = p.add_run(label)
        run.font.size = Pt(18)
        run.font.color.rgb = CLASSIC_BLUE
        run.font.bold = True
        run.font.name = "Calibri"
        
        cell_val = sig_table.rows[i].cells[1]
        p = cell_val.paragraphs[0]
        pPr = p._p.get_or_add_pPr()
        pBdr = parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'<w:bottom w:val="single" w:sz="4" w:space="1" w:color="{CLASSIC_BLUE_HEX}"/>'
            f'</w:pBdr>'
        )
        pPr.append(pBdr)


def generate_carrier_rating(doc, data):
    """Section 16: Carrier Rating"""
    add_page_break(doc)
    add_section_header(doc, "Carrier Rating")
    
    add_formatted_paragraph(doc,
        "AM Best is a full-service credit rating organization dedicated to serving the insurance "
        "industry. AM Best ratings provide an independent third-party evaluation of an insurer's "
        "financial strength and ability to meet its ongoing insurance policy and contract obligations.",
        size=18, space_after=15)
    
    # Build carrier rating table from all coverages
    coverages = data.get("coverages", {})
    carriers_seen = {}
    coverage_names = {
        "property": "Property",
        "general_liability": "General Liability",
        "umbrella": "Umbrella",
        "workers_comp": "Workers Compensation",
        "commercial_auto": "Commercial Auto"
    }
    
    for key, display_name in coverage_names.items():
        cov = coverages.get(key)
        if cov:
            carrier = cov.get("carrier", "")
            if carrier and carrier not in carriers_seen:
                carriers_seen[carrier] = {
                    "rating": cov.get("am_best_rating", "N/A"),
                    "admitted": "Admitted" if cov.get("carrier_admitted", True) else "Non-Admitted",
                    "coverages": [display_name]
                }
            elif carrier in carriers_seen:
                carriers_seen[carrier]["coverages"].append(display_name)
    
    if carriers_seen:
        headers = ["Carrier", "AM Best Rating", "Admitted Status", "Coverages"]
        rows = []
        for carrier, info in carriers_seen.items():
            rows.append([carrier, info["rating"], info["admitted"], ", ".join(info["coverages"])])
        create_styled_table(doc, headers, rows)


def generate_general_statement(doc):
    """Section 17: General Statement"""
    add_page_break(doc)
    add_section_header(doc, "General Statement")
    
    sections = [
        ("Important Notice", "This proposal of insurance is provided for informational purposes only and does not constitute a binder of insurance. Coverage is not effective until confirmed in writing by the insurance carrier. The actual terms, conditions, and exclusions of coverage will be governed by the policies as issued. Please review your policies carefully upon receipt and report any discrepancies to your HUB International representative immediately."),
        ("Proposal Limitations", "This proposal is based on the information provided to us and the coverages available at the time of preparation. Insurance markets, rates, and terms are subject to change. HUB International makes no guarantee that the proposed coverages, limits, or premiums will remain available at the time of binding."),
        ("Claims Reporting", "All claims and potential claims should be reported immediately to your HUB International representative and directly to the applicable insurance carrier. Failure to report claims promptly may jeopardize coverage."),
        ("Policy Review", "We strongly recommend that you review all insurance policies upon receipt to ensure they accurately reflect the coverages intended. Any errors or omissions should be reported to your HUB International representative within 30 days of policy receipt."),
        ("Surplus Lines Notice", "Certain coverages in this proposal may be placed with surplus lines carriers. Surplus lines carriers are not members of state guaranty funds, and in the event of insolvency, claims may not be covered by state guaranty fund protection. Surplus lines placements are subject to surplus lines taxes and fees as required by applicable state law."),
    ]
    
    for title, text in sections:
        add_subsection_header(doc, title)
        add_formatted_paragraph(doc, text, size=18, space_after=10)


def generate_property_definitions(doc):
    """Section 18: Property Coverage Definitions"""
    add_page_break(doc)
    add_section_header(doc, "Property Coverage Definitions")
    
    definitions = [
        ("Actual Cash Value (ACV)", "The cost to repair or replace damaged property with material of like kind and quality, less depreciation."),
        ("Replacement Cost", "The cost to repair or replace damaged property with material of like kind and quality, without deduction for depreciation."),
        ("Agreed Value", "A predetermined value agreed upon by the insurer and insured, eliminating the coinsurance penalty."),
        ("Coinsurance", "A provision requiring the insured to carry insurance equal to a specified percentage of the property's value. Failure to maintain adequate limits may result in a penalty at the time of loss."),
        ("Business Income", "Coverage for loss of income sustained due to the necessary suspension of operations during the period of restoration following a covered loss."),
        ("Extra Expense", "Coverage for expenses incurred to avoid or minimize the suspension of business operations following a covered loss."),
        ("Ordinance or Law", "Coverage for the increased cost of construction due to enforcement of building codes or ordinances following a covered loss."),
        ("Equipment Breakdown", "Coverage for loss or damage to covered equipment caused by mechanical breakdown, electrical arcing, or other covered causes."),
        ("Flood", "Coverage for direct physical loss or damage caused by flood, as defined in the policy. Flood coverage may be subject to separate limits, deductibles, and waiting periods."),
        ("Earthquake", "Coverage for direct physical loss or damage caused by earthquake or earth movement. Subject to separate limits and deductibles."),
        ("Named Storm", "A storm system that has been designated a tropical storm or hurricane by the National Weather Service. Named storm deductibles typically apply as a percentage of insured values."),
        ("Dampness or dryness of the atmosphere and changes in the temperature", "These perils are typically excluded under standard property policies unless specifically endorsed."),
        ("Artificially generated electrical currents", "Damage caused by artificially generated electrical current, including power surges, is typically excluded unless caused by lightning."),
        ("Explosion of steam boilers", "Damage from steam boiler explosions is typically covered under equipment breakdown coverage rather than standard property coverage."),
        ("Mold", "Mold damage is typically excluded or subject to limited coverage under standard property policies."),
        ("Terrorism", "Coverage for acts of terrorism as defined by the Terrorism Risk Insurance Act (TRIA). See TRIA Disclosure section for details."),
    ]
    
    for term, definition in definitions:
        p = doc.add_paragraph()
        run_term = p.add_run(f"{term}: ")
        run_term.font.size = Pt(18)
        run_term.font.color.rgb = ELECTRIC_BLUE
        run_term.font.bold = True
        run_term.font.name = "Calibri"
        run_def = p.add_run(definition)
        run_def.font.size = Pt(18)
        run_def.font.color.rgb = CLASSIC_BLUE
        run_def.font.name = "Calibri"
        p.paragraph_format.space_after = Pt(8)


def generate_how_we_get_paid(doc):
    """Section 19: How We Get Paid"""
    add_page_break(doc)
    add_section_header(doc, "How We Get Paid")
    
    add_formatted_paragraph(doc,
        "HUB International takes pride in the services our brokerages provide to you, our client, "
        "for insurance and risk management programs. For our efforts we are compensated in a variety "
        "of ways, primarily in the form of commissions and contingency amounts paid by insurance "
        "companies and, in some cases, fees paid by clients or third parties. The means by which we "
        "are compensated are described below.",
        size=18, space_after=15)
    
    add_subsection_header(doc, "Commission income")
    add_formatted_paragraph(doc,
        "Commission, normally calculated as a percentage of the premium paid to the insurer for the "
        "specific policy, is paid to us by the insurer to distribute and service your insurance policy. "
        "Our commission is included in the premium paid by you. The individuals at HUB International "
        "who place and service your insurance may be paid compensation that varies directly with the "
        "commissions we receive.",
        size=18, space_after=10)
    
    add_subsection_header(doc, "Contingency income")
    add_formatted_paragraph(doc,
        "We also receive income through contingency arrangements with most insurers. They are called "
        "\"contingent\" because to qualify for payment we normally need to meet certain criteria, usually "
        "measured on an annual basis. Contingency arrangements vary, but payment under these agreements "
        "is normally the result of growing the business by attracting new customers, helping the insurance "
        "company gather and assess underwriting information and/or working to renew the policies of "
        "existing insureds. There is currently no meaningful method to determine the exact impact that "
        "any particular insurance policy has on contingency arrangements. However, brokers tend to receive "
        "higher contingency payments when they grow their business and retain clients through better service. "
        "In other words, the amount of earned contingency income depends on the overall size and/or "
        "profitability of all of a group of accounts, as opposed to the placement or profitability of any "
        "particular insurance policy. For this reason, the individuals involved in placing or servicing "
        "insurance are rarely, if ever, compensated directly for the contingent income that we receive.",
        size=18, space_after=10)
    
    add_formatted_paragraph(doc,
        "Please also feel free to ask any questions about our compensation generally, or as to your "
        "specific insurance proposal or placement, by contacting your HUB broker or customer service "
        "representative directly, or by calling our client hotline at 1-866-857-4073.",
        size=18, space_after=15)
    
    add_subsection_header(doc, "Privacy Policy")
    add_formatted_paragraph(doc,
        "To view our privacy policy, please visit: www.hubinternational.com/about-us/privacy-policy/",
        size=18)


def generate_hub_advantage(doc):
    """Section 20: The HUB Advantage"""
    add_page_break(doc)
    add_section_header(doc, "Our Commitment — The HUB Advantage")
    
    add_formatted_paragraph(doc,
        "HUB International is dedicated to maintaining and upholding the highest standards of ethical "
        "conduct and integrity in all of our dealings with you, our client. We want to be your trusted "
        "risk advisor, and as such, we need to earn your confidence. So we are making a promise. We call "
        "it The HUB Advantage. Our mission is to make the advantage yours — and this is our commitment.",
        size=18, space_after=15)
    
    commitments = [
        "We strive to secure the most favorable terms from insurers, taking into account all of the circumstances — the risk you need to insure, the cost of insurance, the financial condition of the insurer, the insurer's reputation for service, and any other factors that are specific to your situation.",
        "We are open and honest as to how we are paid for placing your insurance. Our answers to your questions will be forthright and understandable. When we intend to seek a fixed fee for our efforts, we will disclose it to you in writing and obtain your approval prior to coverage being bound.",
        "You make the ultimate decision as to both the terms of insurance and the company providing your coverage. Our objective is to provide you with choices that meet your insurance needs, and to educate you so your decision is fully informed and best suited to your circumstances.",
        "We comply with the laws of every jurisdiction in which we operate, including those that apply to how insurance brokerages and agencies are paid. If the laws change, we will respond in a timely and appropriate manner.",
    ]
    
    for commitment in commitments:
        p = doc.add_paragraph()
        run = p.add_run(f"• {commitment}")
        run.font.size = Pt(18)
        run.font.color.rgb = CLASSIC_BLUE
        run.font.name = "Calibri"
        p.paragraph_format.space_after = Pt(10)
    
    add_formatted_paragraph(doc,
        "We take our responsibility to our customers very seriously. If at any time you feel that we are "
        "not fulfilling your expectations — that we are not meeting our Client Commitment — please contact "
        "your account executive or call our toll free client hotline at 1-866-857-4073, and your concerns "
        "will be addressed as soon as possible.",
        size=18, space_before=15, space_after=15)
    
    add_formatted_paragraph(doc, "The HUB Advantage", size=20, color=ELECTRIC_BLUE, bold=True,
                           alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_formatted_paragraph(doc, "The privilege is ours, but the advantage is yours.", size=20,
                           color=CLASSIC_BLUE, bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER)


def generate_tria_disclosure(doc):
    """Section 21: TRIA Disclosure"""
    add_page_break(doc)
    add_section_header(doc, "Terrorism Risk Insurance Act (TRIA) Disclosure")
    
    add_formatted_paragraph(doc,
        'You are hereby notified that under the Terrorism Risk Insurance Act, as amended in 2015, '
        'the definition of act of terrorism has changed. As defined in Section 102 (1) of the Act: '
        'The term "act of terrorism" means any act or acts that is certified by the Secretary of the '
        'Treasury — in consultation with the Secretary of Homeland Security, and the Attorney General '
        'of the United States — to be an act of terrorism; to be a violent act or an act that is '
        'dangerous to human life, property, or infrastructure; to have resulted in damage within the '
        'United States, or outside the United States in the case of certain air carriers or vessels or '
        'the premises of a United States mission; and to have been committed by an individual or '
        'individuals as a part of an effort to coerce the civilian population of the United States or '
        'to influence the policy or affect the conduct of the United States Government by coercion.',
        size=18, space_after=15)
    
    add_formatted_paragraph(doc,
        'Under the coverage, any losses resulting from certified acts of terrorism may be partially '
        'reimbursed by the United States Government under a formula established by the Terrorism Risk '
        'Insurance Act, as amended. However, your policy may contain other exclusions which might affect '
        'the coverage, such as an exclusion for nuclear events. Under the formal, the United States '
        'Government generally reimburses 80% beginning on January 1, 2020 of covered terrorism losses '
        'exceeding the statutorily established deductible paid by the insurance company providing the '
        'coverage. The Terrorism Risk Insurance Act, as amended, contains a $100 billion cap that limits '
        'United States government reimbursement as well as insurers\' liability for losses resulting from '
        'certified acts of terrorism when the amount of such losses exceed $100 billion in any one calendar '
        'year. If the aggregate insured losses for all insured exceed $100 billion, your coverage may be reduced.',
        size=18)


def generate_california_licenses(doc):
    """Section 22: California Licenses"""
    add_page_break(doc)
    add_section_header(doc, "California Licenses")
    
    add_formatted_paragraph(doc,
        "The following HUB International entities are licensed in the State of California:",
        size=18, space_after=10)
    
    headers = ["Entity Name", "License Number"]
    rows = [[name, lic] for name, lic in CA_LICENSES]
    create_styled_table(doc, headers, rows, col_widths=[5.5, 2.0])


def generate_coverage_recommendations(doc):
    """Section 23: Coverage Recommendations"""
    add_page_break(doc)
    add_section_header(doc, "Coverage Recommendations")
    
    add_formatted_paragraph(doc,
        "HUB International recommends that you consider the following coverages and risk management "
        "strategies to protect your hospitality business. These recommendations are based on our "
        "extensive experience in the hotel and hospitality insurance industry.",
        size=18, space_after=15)
    
    recommendations = [
        ("Umbrella/Excess Liability", "We recommend maintaining umbrella or excess liability coverage with limits adequate to protect your assets. Hotels face unique liability exposures including guest injuries, swimming pool incidents, and food service operations. Higher limits provide additional protection above your primary liability policies."),
        ("Cyber Liability", "Hotels collect and store sensitive guest information including credit card numbers, personal identification, and travel details. Cyber liability coverage protects against data breaches, ransomware attacks, and regulatory fines. We strongly recommend this coverage for all hospitality operations."),
        ("Employment Practices Liability (EPLI)", "Hotels employ large numbers of workers in various roles, creating exposure to employment-related claims including wrongful termination, discrimination, harassment, and wage disputes. EPLI coverage is essential for protecting your business against these claims."),
        ("Crime/Employee Dishonesty", "Hotels handle significant amounts of cash and guest valuables. Crime coverage protects against employee theft, forgery, computer fraud, and funds transfer fraud."),
        ("Flood Insurance", "Standard property policies exclude flood damage. If your properties are located in flood-prone areas, we strongly recommend purchasing separate flood coverage through the National Flood Insurance Program (NFIP) or private flood markets."),
        ("Earthquake Coverage", "Standard property policies exclude earthquake damage. If your properties are located in seismically active areas, we recommend purchasing earthquake coverage with appropriate limits and deductibles."),
        ("Business Income/Extra Expense", "Adequate business income coverage is critical for hotels. A covered loss that forces temporary closure can result in significant lost revenue. We recommend Actual Loss Sustained (ALS) coverage with an extended period of indemnity of at least 12 months."),
        ("Equipment Breakdown", "Hotels rely heavily on mechanical and electrical equipment including HVAC systems, elevators, kitchen equipment, and laundry facilities. Equipment breakdown coverage protects against losses from mechanical failure, electrical arcing, and other equipment-related incidents."),
        ("Liquor Liability", "If your hotel serves alcohol, liquor liability coverage is essential. This coverage protects against claims arising from the sale, service, or furnishing of alcoholic beverages."),
        ("Pollution Liability", "Hotels may face pollution exposures from swimming pool chemicals, cleaning supplies, underground storage tanks, and mold. Pollution liability coverage provides protection against these environmental risks."),
    ]
    
    for title, text in recommendations:
        p = doc.add_paragraph()
        run_title = p.add_run(f"{title}: ")
        run_title.font.size = Pt(18)
        run_title.font.color.rgb = ELECTRIC_BLUE
        run_title.font.bold = True
        run_title.font.name = "Calibri"
        run_text = p.add_run(text)
        run_text.font.size = Pt(18)
        run_text.font.color.rgb = CLASSIC_BLUE
        run_text.font.name = "Calibri"
        p.paragraph_format.space_after = Pt(12)
    
    add_formatted_paragraph(doc, "", space_before=15)
    add_callout_box(doc,
        "Please discuss these recommendations with your HUB International representative to determine "
        "which coverages are appropriate for your specific operations and risk profile. Coverage availability "
        "and pricing may vary based on your individual circumstances.")


# ─── Main Generator ───────────────────────────────────────────

def generate_proposal(data: dict, output_path: str) -> str:
    """
    Generate a complete branded DOCX proposal.
    
    Args:
        data: Structured insurance data from extraction
        output_path: Path to save the DOCX file
        
    Returns:
        Path to the generated DOCX file
    """
    logger.info(f"Generating proposal for: {data.get('client_info', {}).get('named_insured', 'Unknown')}")
    
    doc = Document()
    
    # Set default font
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(18)
    font.color.rgb = CLASSIC_BLUE
    
    # Set margins
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)
    
    # Part 1: Front Matter
    generate_cover_page(doc, data)
    generate_service_team(doc, data)
    generate_premium_summary(doc, data)
    generate_payment_options(doc, data)
    generate_subjectivities(doc, data)
    generate_named_insureds(doc, data)
    generate_information_summary(doc, data)
    generate_locations(doc, data)
    
    # Part 2: Coverage Sections (only if quoted)
    coverages = data.get("coverages", {})
    if "property" in coverages:
        generate_coverage_section(doc, data, "property", "Property Coverage")
    if "general_liability" in coverages:
        generate_coverage_section(doc, data, "general_liability", "General Liability Coverage")
    if "workers_comp" in coverages:
        generate_coverage_section(doc, data, "workers_comp", "Workers Compensation Coverage")
    if "commercial_auto" in coverages:
        generate_coverage_section(doc, data, "commercial_auto", "Commercial Auto Coverage")
    if "umbrella" in coverages:
        generate_coverage_section(doc, data, "umbrella", "Umbrella / Excess Liability Coverage")
    
    # Part 3: Signature Pages
    generate_confirmation_to_bind(doc, data)
    
    # Part 4: Compliance Pages (ALWAYS REQUIRED)
    generate_electronic_consent(doc)
    generate_carrier_rating(doc, data)
    generate_general_statement(doc)
    generate_property_definitions(doc)
    generate_how_we_get_paid(doc)
    generate_hub_advantage(doc)
    generate_tria_disclosure(doc)
    generate_california_licenses(doc)
    generate_coverage_recommendations(doc)
    
    # Save
    doc.save(output_path)
    logger.info(f"Proposal saved to: {output_path}")
    return output_path
