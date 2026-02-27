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

# AM Best Rating Lookup Table - common hospitality insurance carriers
# Updated periodically; used as fallback when quote doesn't include rating
AM_BEST_RATINGS = {
    # Property carriers
    "vantage risk": "A- (VII)",
    "vantage risk specialty insurance": "A- (VII)",
    "tower hill insurance": "A- (VII)",
    "tower hill prime insurance": "A- (VII)",
    "tower hill preferred insurance": "A- (VII)",
    "starr surplus lines": "A (XV)",
    "starr indemnity": "A (XV)",
    "lexington insurance": "A+ (XV)",
    "scottsdale insurance": "A+ (XV)",
    "zurich": "A+ (XV)",
    "zurich american insurance": "A+ (XV)",
    "great lakes insurance": "A+ (XV)",
    "lloyd's of london": "A (XV)",
    "lloyds": "A (XV)",
    "nautilus insurance": "A+ (XV)",
    "empire indemnity": "A (VIII)",
    "colony insurance": "A (VIII)",
    "james river insurance": "A- (VIII)",
    "canopius": "A- (VII)",
    # GL carriers
    "amtrust": "A- (VIII)",
    "amtrust e&s": "A- (VIII)",
    "amtrust financial": "A- (VIII)",
    "associated industries": "A- (VIII)",
    "associated industries insurance": "A- (VIII)",
    "technology insurance company": "A- (VIII)",
    "security national insurance": "A- (VIII)",
    "wesco insurance": "A- (VIII)",
    "southlake specialty insurance": "A- (VIII)",
    "southlake specialty": "A- (VIII)",
    "futuristic underwriters": "A- (VIII)",
    "kinsale insurance": "A (VIII)",
    "kinsale insurance company": "A (VIII)",
    "gotham insurance": "A- (VIII)",
    "gotham insurance company": "A- (VIII)",
    "coaction": "A- (VIII)",
    "coaction specialty insurance": "A- (VIII)",
    "markel insurance": "A (XV)",
    "evanston insurance": "A+ (XV)",
    "general star indemnity": "A++ (XV)",
    "essentia insurance": "A- (VII)",
    "mount vernon fire insurance": "A++ (XV)",
    # Umbrella/Excess carriers
    "palms insurance": "A- (VII)",
    "palms insurance company": "A- (VII)",
    "starstone": "A- (VII)",
    "starstone national insurance": "A- (VII)",
    "ironshore specialty insurance": "A (XV)",
    "westchester surplus lines": "A+ (XV)",
    "great american insurance": "A+ (XV)",
    "argo group": "A- (VIII)",
    "hudson insurance": "A+ (XV)",
    # Crime carriers
    "federal insurance company": "A++ (XV)",
    "federal insurance": "A++ (XV)",
    "chubb": "A++ (XV)",
    "vigilant insurance": "A++ (XV)",
    # WC carriers
    "employers insurance": "A (VIII)",
    "employers compensation insurance": "A (VIII)",
    "zenith insurance": "A (VII)",
    "pinnacol assurance": "A (VII)",
    "texas mutual insurance": "A (VIII)",
    "state compensation insurance fund": "A (VIII)",
    # Auto carriers
    "national interstate insurance": "A (VIII)",
    # Flood carriers
    "selective insurance": "A+ (VIII)",
    "selective": "A+ (VIII)",
    "wright flood": "N/A (NFIP)",
    # Multi-line carriers
    "travelers": "A++ (XV)",
    "hartford": "A+ (XV)",
    "cna": "A (XV)",
    "liberty mutual": "A (XV)",
    "nationwide": "A+ (XV)",
    "berkshire hathaway": "A++ (XV)",
    "aig": "A (XV)",
    "chubb": "A++ (XV)",
}


def _clean_carrier_name(name):
    """Strip (Non-Adm), (Non-Admitted), (Surplus Lines) etc. from carrier names."""
    if not name:
        return name
    import re
    name = re.sub(r'\s*\(Non-Adm(?:itted)?\)', '', name, flags=re.IGNORECASE).strip()
    name = re.sub(r'\s*\(Surplus Lines?\)', '', name, flags=re.IGNORECASE).strip()
    name = re.sub(r'\s*\(E&S\)', '', name, flags=re.IGNORECASE).strip()
    return name


def lookup_am_best(carrier_name):
    """Look up AM Best rating for a carrier. Returns rating or None."""
    if not carrier_name:
        return None
    name_lower = carrier_name.lower().strip()
    # Direct match
    if name_lower in AM_BEST_RATINGS:
        return AM_BEST_RATINGS[name_lower]
    # Partial match - check if any key is contained in the carrier name
    for key, rating in AM_BEST_RATINGS.items():
        if key in name_lower or name_lower in key:
            return rating
    return None

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


def remove_cell_borders(cell):
    """Remove all borders from a cell."""
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


def set_cell_width(cell, inches):
    """Set cell width in inches using XML."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{int(inches * 1440)}" w:type="dxa"/>')
    existing = tcPr.find(qn('w:tcW'))
    if existing is not None:
        tcPr.remove(existing)
    tcPr.append(tcW)


def set_cell_vertical_alignment(cell, align="center"):
    """Set vertical alignment of cell content."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    vAlign = parse_xml(f'<w:vAlign {nsdecls("w")} w:val="{align}"/>')
    existing = tcPr.find(qn('w:vAlign'))
    if existing is not None:
        tcPr.remove(existing)
    tcPr.append(vAlign)


def add_formatted_paragraph(doc, text, size=11, color=CLASSIC_BLUE, bold=False,
                            italic=False, alignment=WD_ALIGN_PARAGRAPH.LEFT,
                            space_before=0, space_after=0):
    """Add a formatted paragraph to the document."""
    p = doc.add_paragraph()
    p.alignment = alignment
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = Pt(14)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.color.rgb = color
    run.font.bold = bold
    run.font.italic = italic
    run.font.name = "Calibri"
    return p


def add_section_header(doc, text):
    """Add a 22pt Classic Blue bold section header with enough space to clear page header."""
    return add_formatted_paragraph(doc, text, size=22, color=CLASSIC_BLUE, bold=True,
                                   space_before=36, space_after=12)


def add_subsection_header(doc, text):
    """Add a 14pt Electric Blue bold subsection header."""
    return add_formatted_paragraph(doc, text, size=14, color=ELECTRIC_BLUE, bold=True,
                                   space_before=12, space_after=8)


def create_styled_table(doc, headers, rows, col_widths=None, header_size=10, body_size=10,
                        total_width=7.5, col_alignments=None, header_alignments=None):
    """Create a table with HUB styling: Electric Blue header, alternating rows.
    
    Args:
        doc: Document object
        headers: List of header strings
        rows: List of row data lists
        col_widths: List of column widths in inches. If None, auto-calculated.
        header_size: Font size for header row (default 10pt)
        body_size: Font size for body rows (default 10pt)
        total_width: Total table width in inches (default 7.5)
        col_alignments: Dict or list of WD_ALIGN_PARAGRAPH values per column for body rows. If None, all left-aligned.
        header_alignments: Dict or list of WD_ALIGN_PARAGRAPH values per column for header row.
                          If None, all center-aligned (default behavior).
    """
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Auto-layout off for fixed widths
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
    tblLayout = parse_xml(f'<w:tblLayout {nsdecls("w")} w:type="fixed"/>')
    existing_layout = tblPr.find(qn('w:tblLayout'))
    if existing_layout is not None:
        tblPr.remove(existing_layout)
    tblPr.append(tblLayout)
    
    # Set total table width
    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="{int(total_width * 1440)}" w:type="dxa"/>')
    existing_tblW = tblPr.find(qn('w:tblW'))
    if existing_tblW is not None:
        tblPr.remove(existing_tblW)
    tblPr.append(tblW)
    
    # Calculate column widths if not provided
    if not col_widths:
        col_widths = [total_width / len(headers)] * len(headers)
    
    # Style header row
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        p = cell.paragraphs[0]
        # Determine header alignment: use header_alignments if specified, else center
        h_align = WD_ALIGN_PARAGRAPH.CENTER
        if header_alignments is not None:
            if isinstance(header_alignments, dict):
                if i in header_alignments and header_alignments[i] is not None:
                    h_align = header_alignments[i]
            elif isinstance(header_alignments, (list, tuple)):
                if i < len(header_alignments) and header_alignments[i] is not None:
                    h_align = header_alignments[i]
        p.alignment = h_align
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = Pt(header_size + 2)
        run = p.add_run(header)
        run.font.size = Pt(header_size)
        run.font.color.rgb = WHITE
        run.font.bold = True
        run.font.name = "Calibri"
        set_cell_shading(cell, ELECTRIC_BLUE_HEX)
        set_cell_width(cell, col_widths[i] if i < len(col_widths) else 1.0)
        set_cell_vertical_alignment(cell, "center")
    
    # Style data rows
    for row_idx, row_data in enumerate(rows):
        for col_idx, cell_text in enumerate(row_data):
            cell = table.rows[row_idx + 1].cells[col_idx]
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(3)
            p.paragraph_format.space_after = Pt(3)
            p.paragraph_format.line_spacing = Pt(body_size + 2)
            run = p.add_run(str(cell_text))
            run.font.size = Pt(body_size)
            run.font.color.rgb = CLASSIC_BLUE
            run.font.name = "Calibri"
            set_cell_width(cell, col_widths[col_idx] if col_idx < len(col_widths) else 1.0)
            set_cell_vertical_alignment(cell, "center")
            # Apply column alignment if specified
            if col_alignments is not None:
                if isinstance(col_alignments, dict):
                    if col_idx in col_alignments and col_alignments[col_idx] is not None:
                        p.alignment = col_alignments[col_idx]
                elif isinstance(col_alignments, (list, tuple)):
                    if col_idx < len(col_alignments) and col_alignments[col_idx] is not None:
                        p.alignment = col_alignments[col_idx]
            # Alternating row colors
            if row_idx % 2 == 1:
                set_cell_shading(cell, EGGSHELL_HEX)
    
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
        run.add_picture(LOGO_PATH, width=Inches(1.8))
    
    # Text cell (right)
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
    
    # Remove borders from header table
    for row in htable.rows:
        for cell in row.cells:
            remove_cell_borders(cell)
    
    # Add page footer with automatic page numbers
    footer = section.footer
    footer.is_linked_to_previous = False
    fp = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    fp.paragraph_format.space_before = Pt(4)
    fp.paragraph_format.space_after = Pt(0)
    # "Page " prefix
    run_prefix = fp.add_run("Page ")
    run_prefix.font.size = Pt(8)
    run_prefix.font.color.rgb = ARCTIC_GRAY
    run_prefix.font.name = "Calibri"
    # Auto page number field
    fldChar1 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
    run_num = fp.add_run()
    run_num._r.append(fldChar1)
    instrText = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>')
    run_num2 = fp.add_run()
    run_num2._r.append(instrText)
    fldChar2 = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
    run_num3 = fp.add_run()
    run_num3._r.append(fldChar2)
    # Style the page number runs
    for r in [run_num, run_num2, run_num3]:
        r.font.size = Pt(8)
        r.font.color.rgb = ARCTIC_GRAY
        r.font.name = "Calibri"


def add_callout_box(doc, text, size=10):
    """Add an eggshell background callout/disclaimer box."""
    table = doc.add_table(rows=1, cols=1)
    cell = table.rows[0].cells[0]
    set_cell_shading(cell, EGGSHELL_HEX)
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.font.color.rgb = CLASSIC_BLUE
    run.font.name = "Calibri"
    run.font.italic = True
    return table


def add_page_break(doc):
    """Add a page break with a spacer paragraph to push content below the header.
    Word suppresses space_before at the top of a new page, so we add an empty
    spacer paragraph with a small font size and fixed spacing to create clearance."""
    doc.add_page_break()
    # Add invisible spacer paragraph - Word won't suppress this
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(0)
    spacer.paragraph_format.space_after = Pt(12)
    spacer.paragraph_format.line_spacing = Pt(2)
    run = spacer.add_run()
    run.font.size = Pt(2)


def fmt_currency(amount):
    """Format a number as currency, preserving cents if present."""
    if isinstance(amount, (int, float)):
        if amount == int(amount):
            return f"${int(amount):,}"
        return f"${amount:,.2f}"
    if isinstance(amount, str):
        if amount.startswith("$"):
            return amount
        try:
            val = float(amount.replace(',', ''))
            if val == int(val):
                return f"${int(val):,}"
            return f"${val:,.2f}"
        except (ValueError, AttributeError):
            return amount
    return str(amount)


# ─── Section Generators ───────────────────────────────────────

def generate_cover_page(doc, data):
    """Section 1: Cover Page - fits on a single page, no page header."""
    ci = data.get("client_info", {})
    client_name = ci.get("named_insured", "Client Name")
    dba = ci.get("dba", "")
    address = ci.get("address", "")
    effective_date = ci.get("effective_date", "")
    proposal_date = datetime.date.today().strftime("%B %d, %Y")
    
    # Ensure cover page section has NO header or footer
    cover_section = doc.sections[0]
    cover_header = cover_section.header
    cover_header.is_linked_to_previous = False
    # Clear any existing header content
    for p in cover_header.paragraphs:
        p.clear()
    # Clear footer on cover page
    cover_footer = cover_section.footer
    cover_footer.is_linked_to_previous = False
    for p in cover_footer.paragraphs:
        p.clear()
    
    # Logo centered
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(40)
    p.paragraph_format.space_after = Pt(6)
    if os.path.exists(LOGO_PATH):
        run = p.add_run()
        run.add_picture(LOGO_PATH, width=Inches(2.5))
    
    # Electric Blue line
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(0)
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>' 
        f'<w:bottom w:val="single" w:sz="36" w:space="1" w:color="{ELECTRIC_BLUE_HEX}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)
    
    # Title - single line
    add_formatted_paragraph(doc, "Commercial Insurance Proposal", size=32, color=CLASSIC_BLUE,
                           bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=16, space_after=36)
    
    # Prepared For box
    box_table = doc.add_table(rows=1, cols=1)
    box_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for row in box_table.rows:
        for cell in row.cells:
            cell.width = Inches(5.5)
    tbl = box_table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else parse_xml(f'<w:tblPr {nsdecls("w")}/>')
    tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:w="7920" w:type="dxa"/>')
    existing_tblW = tblPr.find(qn('w:tblW'))
    if existing_tblW is not None:
        tblPr.remove(existing_tblW)
    tblPr.append(tblW)
    cell = box_table.rows[0].cells[0]
    set_cell_shading(cell, EGGSHELL_HEX)
    
    # Top and bottom blue borders
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="24" w:space="0" w:color="{ELECTRIC_BLUE_HEX}"/>'
        f'<w:bottom w:val="single" w:sz="24" w:space="0" w:color="{ELECTRIC_BLUE_HEX}"/>'
        f'<w:left w:val="none" w:sz="0" w:space="0"/>'
        f'<w:right w:val="none" w:sz="0" w:space="0"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)
    
    # "Prepared For" label
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run("Prepared For")
    run.font.size = Pt(14)
    run.font.color.rgb = CHARCOAL
    run.font.name = "Calibri"
    
    # Client name
    p2 = cell.add_paragraph()
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(4)
    p2.paragraph_format.space_after = Pt(2)
    run2 = p2.add_run(client_name)
    run2.font.size = Pt(28)
    run2.font.color.rgb = ELECTRIC_BLUE
    run2.font.bold = True
    run2.font.name = "Calibri"
    
    # Address if present
    if address:
        p_addr = cell.add_paragraph()
        p_addr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_addr.paragraph_format.space_before = Pt(2)
        p_addr.paragraph_format.space_after = Pt(8)
        run_addr = p_addr.add_run(address)
        run_addr.font.size = Pt(11)
        run_addr.font.color.rgb = CLASSIC_BLUE
        run_addr.font.name = "Calibri"
    
    # DBA if present
    if dba:
        p3 = cell.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p3.paragraph_format.space_after = Pt(8)
        run3 = p3.add_run(f"DBA: {dba}")
        run3.font.size = Pt(14)
        run3.font.color.rgb = CLASSIC_BLUE
        run3.font.name = "Calibri"
    
    # Dates - two column table
    date_table = doc.add_table(rows=2, cols=2)
    date_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Labels row
    for i, label in enumerate(["Proposal Date", "Effective Date"]):
        dc = date_table.rows[0].cells[i]
        p = dc.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(label)
        run.font.size = Pt(10)
        run.font.color.rgb = ARCTIC_GRAY
        run.font.bold = True
        run.font.name = "Calibri"
        remove_cell_borders(dc)
    
    # Values row
    for i, val in enumerate([proposal_date, effective_date]):
        dc = date_table.rows[1].cells[i]
        p = dc.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(val)
        run.font.size = Pt(12)
        run.font.color.rgb = CLASSIC_BLUE
        run.font.bold = True
        run.font.name = "Calibri"
        remove_cell_borders(dc)
    
    # Gray line
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(15)
    p.paragraph_format.space_after = Pt(0)
    pPr = p._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:bottom w:val="single" w:sz="12" w:space="1" w:color="{ARCTIC_GRAY_HEX}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)
    
    # Presented By
    add_formatted_paragraph(doc, "Presented By", size=12, color=CLASSIC_BLUE,
                           alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=10, space_after=2)
    add_formatted_paragraph(doc, "HUB International Midwest Limited", size=13, color=CLASSIC_BLUE,
                           bold=True, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_before=0, space_after=2)
    add_formatted_paragraph(doc, "Franchise Division | Hotel Insurance Programs", size=11,
                           color=ELECTRIC_BLUE, alignment=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)


def generate_service_team(doc, data):
    """Section 2: Service Team - starts a new section with page header."""
    # Add a new section (not just a page break) so the cover page stays header-free
    new_section = doc.add_section()
    new_section.top_margin = Inches(0.7)
    new_section.bottom_margin = Inches(0.6)
    new_section.left_margin = Inches(0.75)
    new_section.right_margin = Inches(0.75)
    add_page_header(doc)
    add_section_header(doc, "Your Service Team")
    
    headers = ["Role", "Name", "Phone", "Email"]
    rows = []
    for member in SERVICE_TEAM:
        rows.append([member["role"], member["name"], member["phone"], member["email"]])
    
    create_styled_table(doc, headers, rows, col_widths=[1.8, 1.6, 1.4, 2.7],
                       header_size=11, body_size=10)
    
    # Office locations
    add_formatted_paragraph(doc, "", space_before=12)
    for loc in OFFICE_LOCATIONS:
        add_formatted_paragraph(doc, loc, size=10, color=CLASSIC_BLUE, space_after=2)


def generate_premium_summary(doc, data):
    """Section 3: Premium Summary - Expiring vs Proposed"""
    add_page_break(doc)
    add_section_header(doc, "Premium Summary")
    
    add_subsection_header(doc, "Coverage Premium Comparison")
    add_formatted_paragraph(doc,
        "Premiums shown include applicable taxes and fees. TRIA/Terrorism premiums are not included.",
        size=9, color=CHARCOAL, space_after=6)
    
    coverages = data.get("coverages", {})
    expiring = data.get("expiring_premiums", {})
    expiring_details = data.get("expiring_details", {})
    
    logger.info(f"Premium Summary - coverages keys: {list(coverages.keys())}")
    logger.info(f"Premium Summary - expiring keys: {list(expiring.keys())} values: {expiring}")
    logger.info(f"Premium Summary - expiring_details keys: {list(expiring_details.keys())}")
    
    # Determine if we have expiring data (enables comparison mode)
    has_expiring = bool(expiring and any(v for v in expiring.values() if v))
    
    coverage_names = {
        "property": "Property",
        "excess_property": "Excess Property (Layer 1)",
        "excess_property_2": "Excess Property (Layer 2)",
        "general_liability": "General Liability",
        "umbrella": "Umbrella / Excess",
        "umbrella_layer_2": "2nd Excess Layer",
        "umbrella_layer_3": "3rd Excess Layer",
        "workers_comp": "Workers Compensation",
        "workers_compensation": "Workers Compensation",
        "commercial_auto": "Commercial Auto",
        "flood": "Flood",
        "epli": "EPLI",
        "cyber": "Cyber",
        "terrorism": "Terrorism / TRIA",
        "crime": "Crime",
        "inland_marine": "Inland Marine",
        "equipment_breakdown": "Equipment Breakdown",
        "liquor_liability": "Liquor Liability",
        "innkeepers_liability": "Innkeepers Liability",
        "environmental": "Environmental / Pollution",
        "workplace_violence": "Workplace Violence",
        "garage_keepers": "Garage Keepers",
        "enviro_pack": "Enviro Pack",
    }
    
    if has_expiring:
        headers = ["Coverage", "Carrier", "Expiring", "Proposed", "$ Change", "% Change"]
    else:
        headers = ["Coverage", "Carrier", "Premium", "Taxes & Fees", "Total"]
    rows = []
    total_expiring = 0
    total_proposed = 0
    total_taxes_fees = 0
    
    # Collect all coverage keys that appear in either proposed or expiring
    all_keys = list(coverage_names.keys())
    
    # Combine proposed umbrella layers into a single total for comparison
    # when expiring has a single combined umbrella premium
    proposed_umb_total = 0
    proposed_umb_carriers = []
    for umb_key in ["umbrella", "umbrella_layer_2", "umbrella_layer_3"]:
        umb_cov = coverages.get(umb_key)
        if umb_cov:
            proposed_umb_total += umb_cov.get("total_premium", 0)
            proposed_umb_carriers.append(_clean_carrier_name(umb_cov.get("carrier", "")))
    
    # Combine expiring umbrella layers into a single total for comparison
    expiring_umb_total = 0
    expiring_umb_carriers = []
    for umb_key in ["umbrella", "umbrella_layer_2", "umbrella_layer_3"]:
        umb_exp = expiring.get(umb_key, 0)
        if umb_exp:
            expiring_umb_total += umb_exp
            umb_detail = expiring_details.get(umb_key, {})
            if umb_detail and umb_detail.get("carrier"):
                expiring_umb_carriers.append(umb_detail.get("carrier"))
    
    for key in all_keys:
        display_name = coverage_names[key]
        cov = coverages.get(key)
        exp = expiring.get(key, 0)
        
        if has_expiring:
            # === COMPARISON MODE (with expiring premiums) ===
            # For umbrella: show combined row instead of individual layers
            if key == "umbrella":
                exp = expiring_umb_total
                if proposed_umb_total > 0:
                    proposed = proposed_umb_total
                    carrier_short = " / ".join(dict.fromkeys(proposed_umb_carriers))
                    if len(carrier_short) > 35:
                        carrier_short = carrier_short[:32] + "..."
                    cov = coverages.get("umbrella")
                elif cov:
                    carrier = _clean_carrier_name(cov.get("carrier", ""))
                    carrier_short = carrier
                    if len(carrier) > 30:
                        carrier_short = carrier.replace("Insurance Company", "Ins Co").replace("Specialty ", "Spec ")
                    proposed = cov.get("total_premium", 0)
                else:
                    exp_detail = expiring_details.get(key, {})
                    carrier_short = " / ".join(expiring_umb_carriers) if expiring_umb_carriers else (_clean_carrier_name(exp_detail.get("carrier", "\u2014")) if exp_detail else "\u2014")
                    proposed = 0
                
                if not exp and not proposed:
                    continue
                
                if exp and exp > 0 and proposed > 0:
                    dollar_change = proposed - exp
                    pct_change = ((proposed - exp) / exp) * 100
                    pct_str = f"{pct_change:+.1f}%"
                    dollar_str = f"+${dollar_change:,.2f}" if dollar_change >= 0 else f"-${abs(dollar_change):,.2f}"
                elif exp and exp > 0 and proposed == 0:
                    dollar_str = "Not Quoted"
                    pct_str = "\u2014"
                elif proposed > 0 and (not exp or exp == 0):
                    dollar_str = "New"
                    pct_str = "New"
                else:
                    dollar_str = "N/A"
                    pct_str = "N/A"
                
                rows.append([
                    display_name,
                    carrier_short,
                    f"${exp:,.2f}" if exp else "N/A",
                    f"${proposed:,.2f}" if proposed else "N/A",
                    dollar_str,
                    pct_str
                ])
                total_expiring += exp if exp else 0
                total_proposed += proposed
                continue
            
            # Skip individual umbrella layers in comparison mode (handled above)
            if key in ("umbrella_layer_2", "umbrella_layer_3"):
                continue
            
            # Skip if neither proposed nor expiring
            if not cov and not exp:
                continue
            
            if cov:
                carrier = _clean_carrier_name(cov.get("carrier", ""))
                carrier_short = carrier
                if len(carrier) > 30:
                    carrier_short = carrier.replace("Insurance Company", "Ins Co").replace("Specialty ", "Spec ")
                proposed = cov.get("total_premium", 0)
            else:
                exp_detail = expiring_details.get(key, {})
                carrier_short = _clean_carrier_name(exp_detail.get("carrier", "\u2014")) if exp_detail else "\u2014"
                proposed = 0
            
            if exp and exp > 0 and proposed > 0:
                dollar_change = proposed - exp
                pct_change = ((proposed - exp) / exp) * 100
                pct_str = f"{pct_change:+.1f}%"
                if dollar_change >= 0:
                    dollar_str = f"+${dollar_change:,.2f}"
                else:
                    dollar_str = f"-${abs(dollar_change):,.2f}"
            elif exp and exp > 0 and proposed == 0:
                dollar_str = "Not Quoted"
                pct_str = "\u2014"
            elif proposed > 0 and (not exp or exp == 0):
                dollar_str = "New"
                pct_str = "New"
            else:
                dollar_str = "N/A"
                pct_str = "N/A"
            
            rows.append([
                display_name,
                carrier_short,
                f"${exp:,.2f}" if exp else "N/A",
                f"${proposed:,.2f}" if proposed else "N/A",
                dollar_str,
                pct_str
            ])
            
            total_expiring += exp if exp else 0
            total_proposed += proposed
        
        else:
            # === SIMPLE MODE (no expiring premiums) ===
            # Show each coverage line individually (no umbrella combining)
            if not cov:
                continue
            
            carrier = _clean_carrier_name(cov.get("carrier", ""))
            carrier_short = carrier
            if len(carrier) > 30:
                carrier_short = carrier.replace("Insurance Company", "Ins Co").replace("Specialty ", "Spec ")
            premium = cov.get("premium", 0) or 0
            total_prem = cov.get("total_premium", 0) or 0
            taxes_fees = total_prem - premium if total_prem > premium else 0
            
            rows.append([
                display_name,
                carrier_short,
                fmt_currency(premium) if premium else "\u2014",
                fmt_currency(taxes_fees) if taxes_fees else "\u2014",
                fmt_currency(total_prem) if total_prem else "\u2014",
            ])
            total_proposed += total_prem
            total_taxes_fees += taxes_fees
    
    # Total row
    if has_expiring:
        total_dollar = total_proposed - total_expiring
        if total_expiring > 0:
            total_pct = ((total_proposed - total_expiring) / total_expiring) * 100
            total_pct_str = f"{total_pct:+.1f}%"
            if total_dollar >= 0:
                total_dollar_str = f"+${total_dollar:,.2f}"
            else:
                total_dollar_str = f"-${abs(total_dollar):,.2f}"
        else:
            total_pct_str = "N/A"
            total_dollar_str = "N/A"
        
        rows.append([
            "TOTAL",
            "",
            f"${total_expiring:,.2f}" if total_expiring else "N/A",
            f"${total_proposed:,.2f}" if total_proposed else "N/A",
            total_dollar_str,
            total_pct_str
        ])
        
        table = create_styled_table(doc, headers, rows,
                                   col_widths=[1.2, 2.0, 1.0, 1.0, 1.0, 0.8],
                                   header_size=10, body_size=10,
                                   col_alignments=[None, None, WD_ALIGN_PARAGRAPH.RIGHT,
                                                   WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT,
                                                   WD_ALIGN_PARAGRAPH.RIGHT])
    else:
        # Simple mode total row
        total_base = total_proposed - total_taxes_fees
        rows.append([
            "TOTAL",
            "",
            fmt_currency(total_base) if total_base else "\u2014",
            fmt_currency(total_taxes_fees) if total_taxes_fees else "\u2014",
            fmt_currency(total_proposed) if total_proposed else "\u2014",
        ])
        
        table = create_styled_table(doc, headers, rows,
                                   col_widths=[1.5, 2.2, 1.2, 1.2, 1.2],
                                   header_size=10, body_size=10,
                                   col_alignments=[None, None, WD_ALIGN_PARAGRAPH.RIGHT,
                                                   WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.RIGHT])
    
    # Bold and shade the total row
    last_row = table.rows[-1]
    for col_idx, cell in enumerate(last_row.cells):
        set_cell_shading(cell, ELECTRIC_BLUE_HEX)
        for p in cell.paragraphs:
            if col_idx >= 2:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            for run in p.runs:
                run.font.bold = True
                run.font.color.rgb = WHITE
    
    # Color-code change columns for comparison mode (green=savings, red=increase)
    if has_expiring:
        GREEN = RGBColor(0x00, 0x80, 0x00)
        RED = RGBColor(0xCC, 0x00, 0x00)
        for row_idx in range(1, len(table.rows) - 1):  # skip header and total
            row = table.rows[row_idx]
            # $ Change column (index 4)
            for p in row.cells[4].paragraphs:
                for run in p.runs:
                    text = run.text.strip()
                    if text.startswith("-"):
                        run.font.color.rgb = GREEN  # savings
                    elif text.startswith("+"):
                        run.font.color.rgb = RED  # increase
            # % Change column (index 5)
            for p in row.cells[5].paragraphs:
                for run in p.runs:
                    text = run.text.strip()
                    if text.startswith("-"):
                        run.font.color.rgb = GREEN
                    elif text.startswith("+"):
                        run.font.color.rgb = RED
    
    # Savings/increase callout (comparison mode only)
    if has_expiring and total_expiring > 0:
        total_dollar = total_proposed - total_expiring
        total_pct = ((total_proposed - total_expiring) / total_expiring) * 100
        if total_dollar != 0:
            direction = "savings" if total_dollar < 0 else "increase"
            add_formatted_paragraph(doc, "", space_before=8)
            callout_text = f"Total premium {direction}: ${abs(total_dollar):,.2f} ({abs(total_pct):.1f}%)"
            add_callout_box(doc, callout_text)
    
    add_formatted_paragraph(doc, "", space_before=6)
    add_callout_box(doc,
        "This comparison is for reference only. Actual coverage terms, conditions, and exclusions "
        "are governed by the policies as issued. Please review all policies carefully upon receipt.")


def generate_payment_options(doc, data):
    """Section 4: Payment Options"""
    add_page_break(doc)
    add_section_header(doc, "Payment Options")
    
    payment_opts = data.get("payment_options", [])
    if payment_opts:
        # Filter out commission-related entries and clean commission text from terms
        import re
        filtered_opts = []
        for po in payment_opts:
            terms = po.get("terms", "")
            carrier = po.get("carrier", "")
            # Skip entries that are purely about commission
            if carrier.lower().strip() in ("commission", "broker fee", "broker"):
                continue
            # Remove commission-related sentences from terms text
            terms = re.sub(r'[^.]*commission[^.]*\.?', '', terms, flags=re.IGNORECASE).strip()
            terms = re.sub(r'[^.]*broker fee[^.]*\.?', '', terms, flags=re.IGNORECASE).strip()
            if terms or po.get("mep"):
                filtered_opts.append({"carrier": carrier, "coverage_type": po.get("coverage_type", ""), "terms": terms, "mep": po.get("mep", "")})
        
        if filtered_opts:
            headers = ["Carrier", "Payment Terms", "Minimum Earned Premium"]
            rows = []
            for po in filtered_opts:
                carrier_name = po.get("carrier", "")
                cov_type = po.get("coverage_type", "")
                # Append coverage type after carrier name (e.g., "Kinsale — Property")
                if cov_type:
                    carrier_display = f"{carrier_name} — {cov_type}"
                else:
                    carrier_display = carrier_name
                rows.append([carrier_display, po.get("terms", ""), po.get("mep", "")])
            create_styled_table(doc, headers, rows, col_widths=[2.5, 3.0, 2.0],
                               header_size=10, body_size=9,
                               col_alignments={2: WD_ALIGN_PARAGRAPH.CENTER})
        else:
            add_formatted_paragraph(doc, "Payment terms to be confirmed upon binding.", size=11)
    else:
        add_formatted_paragraph(doc, "Payment terms to be confirmed upon binding.", size=11)
    
    # Earned premium / cancellation disclaimer - small font, bold, red
    _add_earned_premium_disclaimer(doc)


def generate_subjectivities(doc, data):
    """Section 5: Binding Subjectivities"""
    add_page_break(doc)
    add_section_header(doc, "Binding Subjectivities")
    
    add_formatted_paragraph(doc, "The following items are required prior to or as a condition of binding:",
                           size=11, space_after=8)
    
    coverages = data.get("coverages", {})
    coverage_names = {
        "property": "Property",
        "excess_property": "Excess Property (Layer 1)",
        "excess_property_2": "Excess Property (Layer 2)",
        "general_liability": "General Liability",
        "umbrella": "Umbrella / Excess",
        "umbrella_layer_2": "2nd Excess Layer",
        "umbrella_layer_3": "3rd Excess Layer",
        "workers_comp": "Workers Compensation",
        "workers_compensation": "Workers Compensation",
        "commercial_auto": "Commercial Auto",
        "terrorism": "Terrorism / TRIA",
        "cyber": "Cyber Liability",
        "epli": "Employment Practices Liability",
        "crime": "Crime",
        "flood": "Flood",
        "inland_marine": "Inland Marine",
        "equipment_breakdown": "Equipment Breakdown",
        "liquor_liability": "Liquor Liability",
        "innkeepers_liability": "Innkeepers Liability",
        "environmental": "Environmental / Pollution",
        "workplace_violence": "Workplace Violence",
        "garage_keepers": "Garage Keepers",
        "enviro_pack": "Enviro Pack",
    }
    
    has_subjectivities = False
    for key, display_name in coverage_names.items():
        cov = coverages.get(key)
        if cov and cov.get("subjectivities"):
            has_subjectivities = True
            # Add carrier info with the coverage name
            carrier = _clean_carrier_name(cov.get("carrier", ""))
            header_text = f"{display_name} — {carrier}" if carrier else display_name
            add_subsection_header(doc, header_text)
            for subj in cov["subjectivities"]:
                add_formatted_paragraph(doc, f"☐  {subj}", size=10, space_after=3)
    
    if not has_subjectivities:
        add_formatted_paragraph(doc, "No subjectivities noted. Please confirm with carrier.", size=11)


def _proper_case(name):
    """Convert a name to proper title case, handling special cases.
    ALL CAPS and all lowercase get converted; mixed case is preserved."""
    if not name or not name.strip():
        return name
    s = name.strip()
    # If ALL CAPS or all lowercase, convert to title case
    if s.isupper() or s.islower():
        s = s.title()
    # Fix common abbreviations that should stay uppercase
    for abbr in ["LLC", "LP", "LLP", "INC", "DBA", "II", "III", "IV"]:
        s = s.replace(abbr.title(), abbr)
    return s


def generate_named_insureds(doc, data):
    """Section 6: Named Insureds"""
    add_page_break(doc)
    add_section_header(doc, "Named Insureds")
    
    # Deduplicate named insureds case-insensitively and apply proper case
    raw_named = data.get("named_insureds", [])
    seen = set()
    named = []
    
    # Hotel brand names that should NOT appear in named insured entries
    # (GPT sometimes concatenates brand names from quotes into named insured DBA fields)
    _brand_keywords = {"marriott", "hilton", "ihg", "wyndham", "best western", "choice",
                       "hampton inn", "hampton", "holiday inn", "holiday inn express",
                       "candlewood", "towneplace", "staybridge", "springhill",
                       "comfort inn", "comfort suites", "quality inn", "sleep inn",
                       "la quinta", "days inn", "super 8", "ramada", "baymont",
                       "microtel", "wingate", "hawthorn", "home2", "tru by hilton",
                       "embassy suites", "doubletree", "hyatt", "radisson", "crowne plaza"}
    
    def _sanitize_named_insured(name_str):
        """Remove hotel brand names that GPT may have concatenated into named insured."""
        if not name_str:
            return name_str
        # If the name contains multiple brand keywords, it's likely a GPT hallucination
        name_lower = name_str.lower()
        brand_count = sum(1 for b in _brand_keywords if b in name_lower)
        if brand_count >= 2:
            # Strip everything after the first DBA + entity name
            import re
            # Try to find "LLC DBA <brand>" and keep just the LLC part
            m = re.match(r'^(.+?\b(?:LLC|LP|LLP|Inc|Corp)\b)\s*(?:DBA\s+.+)?$', name_str, re.IGNORECASE)
            if m:
                return m.group(1).strip()
        return name_str
    
    for ni in raw_named:
        if isinstance(ni, dict):
            ni_name = ni.get("name", "")
            ni_dba = ni.get("dba", "")
        else:
            ni_name = str(ni)
            ni_dba = ""
        # Sanitize: remove hallucinated brand concatenations
        ni_name = _sanitize_named_insured(ni_name)
        if ni_dba:
            ni_dba = _sanitize_named_insured(ni_dba)
        key = ni_name.strip().upper()
        if key and key not in seen:
            seen.add(key)
            display = _proper_case(ni_name)
            if ni_dba:
                display += f" DBA {_proper_case(ni_dba)}"
            named.append(display)
    
    # Also check if DBA is available from client_info and append to first named insured
    ci = data.get("client_info", {})
    ci_dba = ci.get("dba", "")
    if named and ci_dba and "DBA" not in named[0].upper():
        named[0] = f"{named[0]} DBA {_proper_case(ci_dba)}"
    
    if named:
        headers = ["#", "Named Insured"]
        rows = [[str(i), ni] for i, ni in enumerate(named, 1)]
        create_styled_table(doc, headers, rows, col_widths=[0.5, 7.0],
                           header_size=10, body_size=10)
    else:
        ni_text = _proper_case(ci.get("named_insured", "TBD"))
        if ci_dba:
            ni_text += f" DBA {_proper_case(ci_dba)}"
        headers = ["#", "Named Insured"]
        rows = [["1", ni_text]]
        create_styled_table(doc, headers, rows, col_widths=[0.5, 7.0],
                           header_size=10, body_size=10)
    
    # Additional Named Insureds
    addl_named = data.get("additional_named_insureds", [])
    if addl_named:
        add_subsection_header(doc, "Additional Named Insureds")
        headers = ["#", "Additional Named Insured"]
        rows = []
        for i, ani in enumerate(addl_named, 1):
            if isinstance(ani, dict):
                name = ani.get("name", "")
                dba = ani.get("dba", "")
                display = f"{name} DBA {dba}" if dba else name
            else:
                display = str(ani)
            rows.append([str(i), display])
        create_styled_table(doc, headers, rows, col_widths=[0.5, 7.0],
                           header_size=10, body_size=10)
    
    # Additional Insureds
    addl_insureds = data.get("additional_insureds", [])
    if addl_insureds:
        add_subsection_header(doc, "Additional Insureds")
        headers = ["#", "Additional Insured", "Relationship"]
        rows = []
        for i, ai in enumerate(addl_insureds, 1):
            if isinstance(ai, dict):
                rows.append([str(i), ai.get("name", ""), ai.get("relationship", "")])
            else:
                rows.append([str(i), str(ai), ""])
        create_styled_table(doc, headers, rows, col_widths=[0.5, 4.5, 2.5],
                           header_size=10, body_size=10)
    
    # Note box
    add_formatted_paragraph(doc, "", space_before=8)
    add_callout_box(doc, "Note: Additional named insureds may be added as required by franchise agreements or management contracts.")
    
    # Additional Interests
    interests = data.get("additional_interests", [])
    if interests:
        add_subsection_header(doc, "Additional Interests")
        headers = ["Type", "Name & Address", "Description"]
        rows = [[ai.get("type", ""), ai.get("name_address", ""), ai.get("description", "")] for ai in interests]
        create_styled_table(doc, headers, rows, col_widths=[1.5, 3.5, 2.5])


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
        rows.insert(1, ["DBA", _proper_case(ci["dba"])])
    
    # Calculate total sales from GL schedule_of_classes exposures
    coverages = data.get("coverages", {})
    gl_cov = coverages.get("general_liability", {})
    if isinstance(gl_cov, dict):
        gl_classes = gl_cov.get("schedule_of_classes", [])
        total_sales = 0
        import re as _re
        for entry in gl_classes:
            if isinstance(entry, dict):
                exposure = entry.get("exposure", "")
                if isinstance(exposure, (int, float)):
                    total_sales += exposure
                elif isinstance(exposure, str):
                    # Parse dollar amounts like "$8,748,612" or "8748612"
                    cleaned = _re.sub(r'[^\d.]', '', exposure.replace(',', ''))
                    if cleaned:
                        try:
                            total_sales += float(cleaned)
                        except ValueError:
                            pass
        if total_sales > 0:
            rows.append(["Total Sales / Exposure", fmt_currency(total_sales)])
        # Also add total_sales from GL coverage if extracted
        elif gl_cov.get("total_sales"):
            rows.append(["Total Sales / Exposure", gl_cov["total_sales"]])
    
    # Add number of locations with property/liability breakdown
    sov_data = data.get("sov_data")
    sov_locs = sov_data.get("locations", []) if sov_data else []
    prop_loc_count = len(sov_locs) if sov_locs else 0
    if not prop_loc_count and coverages.get("property"):
        prop_loc_count = int(coverages["property"].get("num_locations", 0) or 0)
    
    # Count liability locations from schedule_of_classes
    gl_loc_addrs = set()
    if isinstance(gl_cov, dict):
        for entry in gl_cov.get("schedule_of_classes", []):
            if isinstance(entry, dict):
                addr = entry.get("address", "") or entry.get("location", "")
                if addr and addr.strip():
                    gl_loc_addrs.add(_normalize_addr(addr))
    liab_loc_count = len(gl_loc_addrs) if gl_loc_addrs else int(gl_cov.get("num_locations", 0) or 0) if isinstance(gl_cov, dict) else 0
    
    # Calculate UNIQUE location count by merging property and liability addresses
    all_unique_addrs = set()
    # Add property addresses
    if sov_data and sov_data.get("locations"):
        for loc in sov_data["locations"]:
            addr = _normalize_addr(loc.get("address", ""))
            if addr:
                all_unique_addrs.add(addr)
    # Add liability addresses
    for addr in gl_loc_addrs:
        if addr:
            all_unique_addrs.add(addr)
    # Add raw locations (deduped)
    for loc in data.get("locations", []):
        addr = _normalize_addr(loc.get("address", ""))
        if addr:
            all_unique_addrs.add(addr)
    total_loc_count = len(all_unique_addrs) if all_unique_addrs else max(prop_loc_count, liab_loc_count)
    if total_loc_count > 0:
        rows.append(["Total Number of Locations", str(total_loc_count)])
    if prop_loc_count > 0:
        rows.append(["Property Locations", str(prop_loc_count)])
    if liab_loc_count > 0:
        rows.append(["Liability Locations", str(liab_loc_count)])
    
    # Count location types from SOV/property locations and GL schedule_of_classes
    # Use SOV descriptions and GL classifications for PHYSICAL locations only
    # (skip GL exposure-only classes like Hired Auto, Loss Control, Package Stores)
    hotel_count = 0
    office_count = 0
    lro_count = 0
    vacant_count = 0
    other_types = {}
    seen_loc_addrs = set()  # track by normalized address to avoid double-counting
    
    # First: count from SOV locations (most reliable source for property types)
    if sov_data and sov_data.get("locations"):
        for loc in sov_data["locations"]:
            addr = _normalize_addr(loc.get("address", ""))
            if addr in seen_loc_addrs:
                continue
            seen_loc_addrs.add(addr)
            # Determine type from description, hotel_flag, or occupancy
            desc = (loc.get("description", "") or loc.get("hotel_flag", "") or
                    loc.get("occupancy", "") or loc.get("dba", "") or "").lower()
            if any(kw in desc for kw in ["hotel", "motel", "inn", "suite", "lodge", "resort", "hampton", "holiday", "best western", "marriott", "hilton", "ihg", "wyndham", "choice", "comfort", "quality", "candlewood", "towneplace", "springhill"]):
                hotel_count += 1
            elif "office" in desc or "corporate" in desc:
                office_count += 1
            elif "vacant" in desc or "land" in desc:
                vacant_count += 1
            elif "lessor" in desc or "lro" in desc:
                lro_count += 1
            else:
                # Default: if it has a building value, it's likely a hotel
                bldg = loc.get("building_value", 0) or 0
                if bldg > 0:
                    hotel_count += 1
                else:
                    vacant_count += 1
    
    # Second: supplement from GL schedule_of_classes (only PHYSICAL location entries)
    # Skip non-location exposure classes
    _skip_classes = {"hired auto", "non-owned auto", "loss control", "package store",
                     "category vi", "liquor", "sundry", "flat"}
    if isinstance(gl_cov, dict):
        for entry in gl_cov.get("schedule_of_classes", []):
            if isinstance(entry, dict):
                addr = _normalize_addr(entry.get("address", "") or entry.get("location", ""))
                classification = (entry.get("classification", "") or "").lower()
                # Skip non-physical-location exposure classes
                if any(skip in classification for skip in _skip_classes):
                    continue
                if not addr or addr in seen_loc_addrs:
                    continue
                seen_loc_addrs.add(addr)
                if any(kw in classification for kw in ["hotel", "motel", "inn", "suite", "lodge", "resort"]):
                    hotel_count += 1
                elif "office" in classification or "building" in classification:
                    office_count += 1
                elif "lessor" in classification or "lro" in classification:
                    lro_count += 1
                elif "vacant" in classification:
                    vacant_count += 1
                elif classification.strip():
                    type_name = classification.split("-")[0].strip().title()
                    other_types[type_name] = other_types.get(type_name, 0) + 1
    
    type_parts = []
    if hotel_count: type_parts.append(f"{hotel_count} Hotel(s)")
    if office_count: type_parts.append(f"{office_count} Office(s)")
    if lro_count: type_parts.append(f"{lro_count} LRO(s)")
    if vacant_count: type_parts.append(f"{vacant_count} Vacant Land")
    for ot in sorted(other_types):
        type_parts.append(f"{other_types[ot]} {ot}")
    if type_parts:
        rows.append(["Location Types", ", ".join(type_parts)])
    
    # Add TIV from SOV or property quote
    _tiv_added = False
    if sov_data and sov_data.get("totals", {}).get("tiv"):
        rows.append(["Total Insured Value (TIV)", fmt_currency(sov_data["totals"]["tiv"])])
        _tiv_added = True
    elif sov_data and sov_data.get("locations"):
        # Calculate TIV from individual SOV locations
        _total_tiv = sum(loc.get("tiv", 0) or 0 for loc in sov_data["locations"])
        if _total_tiv > 0:
            rows.append(["Total Insured Value (TIV)", fmt_currency(_total_tiv)])
            _tiv_added = True
    if not _tiv_added:
        prop_cov = coverages.get("property", {})
        if isinstance(prop_cov, dict):
            prop_tiv = prop_cov.get("tiv", "")
            if prop_tiv:
                rows.append(["Total Insured Value (TIV)", prop_tiv if isinstance(prop_tiv, str) else fmt_currency(prop_tiv)])
                _tiv_added = True
            elif prop_cov.get("limits"):
                # Try to sum building + BPP + BI from property limits
                _prop_total = 0
                for lim in prop_cov["limits"]:
                    if isinstance(lim, dict):
                        lim_val = lim.get("limit", "")
                        if isinstance(lim_val, str):
                            import re as _re2
                            cleaned = _re2.sub(r'[^\d.]', '', lim_val.replace(',', ''))
                            if cleaned:
                                try:
                                    _prop_total += float(cleaned)
                                except ValueError:
                                    pass
                if _prop_total > 0:
                    rows.append(["Total Insured Value (TIV)", fmt_currency(_prop_total)])
    
    # Add Umbrella / Excess total limit
    _umbrella_total = 0
    for umb_key in ["umbrella", "umbrella_layer_2", "umbrella_layer_3"]:
        umb_cov = coverages.get(umb_key, {})
        if isinstance(umb_cov, dict) and umb_cov.get("carrier"):
            for lim in umb_cov.get("limits", []):
                if isinstance(lim, dict):
                    desc = (lim.get("description", "") or "").lower()
                    if "occurrence" in desc or "each occurrence" in desc:
                        lim_val = lim.get("limit", "")
                        if isinstance(lim_val, str):
                            import re as _re3
                            cleaned = _re3.sub(r'[^\d.]', '', lim_val.replace(',', ''))
                            if cleaned:
                                try:
                                    _umbrella_total += float(cleaned)
                                except ValueError:
                                    pass
                        elif isinstance(lim_val, (int, float)):
                            _umbrella_total += lim_val
                        break  # Only count each occurrence limit once per layer
    if _umbrella_total > 0:
        rows.append(["Total Umbrella / Excess Limit", fmt_currency(_umbrella_total)])
    
    create_styled_table(doc, headers, rows, col_widths=[2.5, 5.0],
                       header_size=10, body_size=10)
    
    add_formatted_paragraph(doc, "", space_before=8)
    add_callout_box(doc, "The information contained in this proposal is based on data provided by the insured and/or their representatives. HUB International makes no warranty as to the accuracy of this information.")


def _normalize_addr(s):
    """Normalize street address for dedup: uppercase, strip, replace common variants.
    Handles U.S. 51 / US 51 / US-51 / Highway 51 / Hwy 51 all mapping to the same form.
    Also strips trailing zip codes."""
    import re as _re_norm
    s = s.strip().upper()
    # Remove periods and commas first
    s = s.replace(".", "").replace(",", "")
    # Normalize route designators: "U.S. 51" / "US 51" / "US-51" / "US HWY 51" -> "HWY 51"
    s = _re_norm.sub(r'\bUS\s*-?\s*(\d)', r'HWY \1', s)
    s = _re_norm.sub(r'\bU\s*S\s*-?\s*(\d)', r'HWY \1', s)
    replacements = {
        " STREET": " ST", " AVENUE": " AVE", " BOULEVARD": " BLVD",
        " DRIVE": " DR", " ROAD": " RD", " LANE": " LN",
        " COURT": " CT", " PLACE": " PL", " CIRCLE": " CIR",
        " HIGHWAY": " HWY", " PARKWAY": " PKWY", " TERRACE": " TER",
        " NORTH": " N", " SOUTH": " S", " EAST": " E", " WEST": " W",
        " NORTHWEST": " NW", " NORTHEAST": " NE", " SOUTHWEST": " SW",
        " SOUTHEAST": " SE",
    }
    for old, new in replacements.items():
        s = s.replace(old, new)
    # Strip trailing zip codes (5-digit or 5+4)
    s = _re_norm.sub(r'\s+\d{5}(-\d{4})?\s*$', '', s)
    s = " ".join(s.split())
    return s


def _normalize_city(s):
    """Normalize city name for dedup: uppercase, remove spaces/punctuation.
    Handles 'La Place' vs 'LaPlace' vs 'LA PLACE' all mapping to 'LAPLACE'."""
    s = s.strip().upper()
    s = s.replace(".", "").replace(",", "").replace("-", "").replace("'", "")
    # Remove ALL spaces so 'LA PLACE' == 'LAPLACE' == 'LA  PLACE'
    s = s.replace(" ", "")
    return s


def _dedup_locations(raw_locations):
    """Deduplicate locations by normalized address."""
    seen_addrs = set()
    locations = []
    for loc in raw_locations:
        addr_key = (_normalize_addr(loc.get("address", "")) + "|" + 
                    _normalize_city(loc.get("city", "")) + "|" +
                    loc.get("state", "").strip().upper())
        if addr_key not in seen_addrs:
            seen_addrs.add(addr_key)
            locations.append(loc)
    return locations


def generate_locations(doc, data):
    """Section 8: Locations — unified schedule with Property/Liability coverage checkmarks."""
    add_page_break(doc)
    add_section_header(doc, "Schedule of Locations")
    
    raw_locations = data.get("locations", [])
    locations = _dedup_locations(raw_locations)
    sov_data = data.get("sov_data")
    coverages = data.get("coverages", {})
    
    # Build a master list of all locations from all sources
    # Each entry: {name, address, city, state, tiv, on_property, on_liability}
    master_locations = []
    seen_addr_keys = set()
    
    # --- Determine which addresses are on the PROPERTY policy ---
    property_addr_keys = set()
    if sov_data and sov_data.get("locations"):
        for loc in sov_data["locations"]:
            addr_key = (_normalize_addr(loc.get("address", "")) + "|" +
                       _normalize_city(loc.get("city", "")) + "|" +
                       loc.get("state", "").strip().upper())
            property_addr_keys.add(addr_key)
    elif "property" in coverages:
        # If no SOV, try to use property quote's schedule_of_values for property addresses
        prop_cov = coverages.get("property", {})
        prop_sov = prop_cov.get("schedule_of_values", []) if isinstance(prop_cov, dict) else []
        if prop_sov:
            for s in prop_sov:
                if isinstance(s, dict):
                    addr = s.get("address", "") or s.get("location", "")
                    if addr:
                        parts = [p.strip() for p in addr.split(",")]
                        street = parts[0] if parts else addr
                        city = parts[1] if len(parts) >= 2 else ""
                        state = ""
                        if len(parts) >= 3:
                            st_m = re.match(r'([A-Z]{2})', parts[2].strip().upper())
                            if st_m: state = st_m.group(1)
                        addr_key = (_normalize_addr(street) + "|" +
                                   _normalize_city(city) + "|" +
                                   state.strip().upper())
                        property_addr_keys.add(addr_key)
        # If still no property addresses, DON'T default to all locations
        # Only locations explicitly on the property quote get the checkmark
        # This prevents GL-only locations from getting property checkmarks
    
    # --- Determine which addresses are on the LIABILITY policy ---
    liability_addr_keys = set()
    gl_cov = coverages.get("general_liability", {})
    gl_classes = gl_cov.get("schedule_of_classes", []) if isinstance(gl_cov, dict) else []
    for entry in gl_classes:
        if isinstance(entry, dict):
            addr = entry.get("address", "")
            if addr:
                # Try to find matching city/state from SOV or locations
                matched = False
                for loc in (sov_data.get("locations", []) if sov_data else []) + locations:
                    if _normalize_addr(addr) in _normalize_addr(loc.get("address", "")) or \
                       _normalize_addr(loc.get("address", "")) in _normalize_addr(addr):
                        addr_key = (_normalize_addr(loc.get("address", "")) + "|" +
                                   _normalize_city(loc.get("city", "")) + "|" +
                                   loc.get("state", "").strip().upper())
                        liability_addr_keys.add(addr_key)
                        matched = True
                        break
                if not matched:
                    liability_addr_keys.add(_normalize_addr(addr) + "||")
    # NOTE: We do NOT default all locations to liability when GL exists.
    # Liability checkmarks are ONLY for locations explicitly listed on the GL carrier quote
    # (from schedule_of_classes addresses and CG2144 designated premises forms).
    # Quotes are primary — SOVs only supplement with additional info.
    
    # --- Pre-scan designated premises to add to liability_addr_keys ---
    import re
    
    # PRIMARY SOURCE: designated_premises array (extracted by GPT from CG2144/NXLL110)
    _cg2144_addrs = []  # Save parsed addresses for the fourth pass later
    designated_premises = gl_cov.get("designated_premises", []) if isinstance(gl_cov, dict) else []
    for raw_addr in designated_premises:
        if not isinstance(raw_addr, str) or not raw_addr.strip() or len(raw_addr.strip()) < 5:
            continue
        raw_addr = raw_addr.strip()
        _cg2144_addrs.append(raw_addr)
        parts = [p.strip() for p in raw_addr.split(",")]
        street = parts[0] if parts else raw_addr
        city = ""
        state = ""
        if len(parts) >= 3:
            city = parts[1]
            st_m = re.match(r'([A-Z]{2})\s*\d*', parts[2].strip().upper())
            if st_m: state = st_m.group(1)
        elif len(parts) == 2:
            st_m = re.match(r'([A-Z]{2})\s*\d*', parts[1].strip().upper())
            if st_m:
                state = st_m.group(1)
            else:
                city = parts[1]
        addr_key = (_normalize_addr(street) + "|" + _normalize_city(city) + "|" + state.strip().upper())
        liability_addr_keys.add(addr_key)
        # Also try matching against SOV/locations for better key resolution
        for loc in (sov_data.get("locations", []) if sov_data else []) + locations:
            if _normalize_addr(street) in _normalize_addr(loc.get("address", "")) or \
               _normalize_addr(loc.get("address", "")) in _normalize_addr(street):
                resolved_key = (_normalize_addr(loc.get("address", "")) + "|" +
                               _normalize_city(loc.get("city", "")) + "|" +
                               loc.get("state", "").strip().upper())
                liability_addr_keys.add(resolved_key)
                break
    
    # FALLBACK: Parse CG2144/NXLL110 form descriptions for addresses
    gl_forms = gl_cov.get("forms_endorsements", []) if isinstance(gl_cov, dict) else []
    for form in gl_forms:
        if not isinstance(form, dict):
            continue
        desc = (form.get("description", "") or "").upper()
        if not any(kw in desc for kw in ["DESIGNATED PREMISES", "CG 21 44", "CG2144", "NXLL110", "NXLL 110", "LIMITATION OF COVERAGE"]):
            continue
        # Parse numbered addresses: "1) 4285 Highway 51, LaPlace, LA 70068"
        addr_pattern = re.findall(r'\d+\)\s*(.+?)(?=\s*\d+\)|$)', desc, re.DOTALL)
        if not addr_pattern:
            addr_pattern = [a.strip() for a in re.split(r'[;\n]', desc) if re.search(r'\d+\s+\w+', a.strip())]
        for raw_addr in addr_pattern:
            raw_addr = raw_addr.strip().rstrip(',')
            if not raw_addr or len(raw_addr) < 5:
                continue
            _cg2144_addrs.append(raw_addr)
            # Parse city/state from address
            parts = [p.strip() for p in raw_addr.split(",")]
            street = parts[0] if parts else raw_addr
            city = ""
            state = ""
            if len(parts) >= 3:
                city = parts[1]
                st_m = re.match(r'([A-Z]{2})\s*\d*', parts[2].strip().upper())
                if st_m: state = st_m.group(1)
            elif len(parts) == 2:
                st_m = re.match(r'([A-Z]{2})\s*\d*', parts[1].strip().upper())
                if st_m:
                    state = st_m.group(1)
                else:
                    city = parts[1]
            addr_key = (_normalize_addr(street) + "|" + _normalize_city(city) + "|" + state.strip().upper())
            liability_addr_keys.add(addr_key)
            # Also try matching against SOV/locations for better key resolution
            for loc in (sov_data.get("locations", []) if sov_data else []) + locations:
                if _normalize_addr(street) in _normalize_addr(loc.get("address", "")) or \
                   _normalize_addr(loc.get("address", "")) in _normalize_addr(street):
                    resolved_key = (_normalize_addr(loc.get("address", "")) + "|" +
                                   _normalize_city(loc.get("city", "")) + "|" +
                                   loc.get("state", "").strip().upper())
                    liability_addr_keys.add(resolved_key)
                    break
    
    # NOTE: Blanket liability fallback removed. Multi-pass extraction (Pass 3) in
    # proposal_extractor.py now handles focused address extraction for GL when
    # designated_premises is empty. Liability checkmarks are only applied to
    # locations explicitly confirmed on the liability quote per Stefan's rule.
    
    # --- Build master location list ---
    # Helper: fuzzy address matching (must be defined before _is_on_liability)
    def _fuzzy_addr_match(addr1, addr2):
        """Check if two normalized addresses refer to the same location.
        Handles cases like '4288 HWY 51' vs '4285 HWY 51' by comparing
        the street name portion after stripping house numbers."""
        if not addr1 or not addr2:
            return False
        if addr1 == addr2:
            return True
        if addr1 in addr2 or addr2 in addr1:
            return True
        # Extract street name without house number for fuzzy match
        num1 = re.match(r'^(\d+)\s+(.+)', addr1)
        num2 = re.match(r'^(\d+)\s+(.+)', addr2)
        if num1 and num2:
            street1 = num1.group(2)
            street2 = num2.group(2)
            house1 = int(num1.group(1))
            house2 = int(num2.group(1))
            # Same street name and house numbers within 20 of each other
            if street1 == street2 and abs(house1 - house2) <= 20:
                return True
        return False
    
    # Helper: check if an addr_key matches any key in liability_addr_keys using fuzzy matching
    def _is_on_liability(addr_key):
        """Check if addr_key is in liability_addr_keys, with fuzzy address matching."""
        if addr_key in liability_addr_keys:
            return True
        # Fuzzy match: compare the street portion of the addr_key against all liability keys
        parts = addr_key.split("|")
        if len(parts) != 3:
            return False
        addr_norm = parts[0]
        state_norm = parts[2]
        for lk in liability_addr_keys:
            lk_parts = lk.split("|")
            if len(lk_parts) != 3:
                continue
            # Must match state (or one is empty)
            if state_norm and lk_parts[2] and state_norm != lk_parts[2]:
                continue
            if _fuzzy_addr_match(addr_norm, lk_parts[0]):
                return True
        return False
    
    # Helper: check if an addr_key matches any key in property_addr_keys
    # Uses STRICT matching (exact normalized address) — no fuzzy tolerance
    # Property checkmarks must be precise: only SOV/property quote locations
    def _is_on_property(addr_key):
        """Check if addr_key is in property_addr_keys. Strict match only."""
        if addr_key in property_addr_keys:
            return True
        # Also try matching just the street portion (ignoring city differences)
        parts = addr_key.split("|")
        if len(parts) != 3:
            return False
        addr_norm = parts[0]
        state_norm = parts[2]
        for pk in property_addr_keys:
            pk_parts = pk.split("|")
            if len(pk_parts) != 3:
                continue
            if state_norm and pk_parts[2] and state_norm != pk_parts[2]:
                continue
            # STRICT: normalized addresses must be identical (no house number tolerance)
            if addr_norm == pk_parts[0]:
                return True
        return False
    
    # First: SOV locations (property locations)
    # Determine fallback corporate name from SOV summary or client_info
    _fallback_corp = ""
    if sov_data and sov_data.get("summary", {}).get("named_insured"):
        _ni = sov_data["summary"]["named_insured"]
        # If named_insured contains " - ", split into corp and DBA
        if " - " in _ni:
            _fallback_corp = _ni.split(" - ")[0].strip()
        else:
            _fallback_corp = _ni.strip()
    if not _fallback_corp:
        _fallback_corp = (ci.get("client_name", "") or "").strip()
    
    # SKIP vacant land — it belongs on property SOV but NOT on the Schedule of Locations
    if sov_data and sov_data.get("locations"):
        for loc in sov_data["locations"]:
            # Filter out vacant land from Schedule of Locations
            desc = (loc.get("description", "") or loc.get("hotel_flag", "") or
                    loc.get("occupancy", "") or loc.get("dba", "") or "").lower()
            if "vacant" in desc or ("land" in desc and "hotel" not in desc and "inn" not in desc):
                # Still track the addr_key so we don't re-add it later
                addr_key = (_normalize_addr(loc.get("address", "")) + "|" +
                           _normalize_city(loc.get("city", "")) + "|" +
                           loc.get("state", "").strip().upper())
                seen_addr_keys.add(addr_key)
                continue
            
            addr_key = (_normalize_addr(loc.get("address", "")) + "|" +
                       _normalize_city(loc.get("city", "")) + "|" +
                       loc.get("state", "").strip().upper())
            # Build "Corporate Name - DBA" format for property name
            corporate_name = (loc.get("corporate_name", "") or "").strip()
            if not corporate_name:
                corporate_name = _fallback_corp
            dba = (loc.get("dba", "") or loc.get("hotel_flag", "") or "").strip()
            if corporate_name and dba:
                name = f"{corporate_name} - {dba}"
            elif dba:
                name = dba
            elif corporate_name:
                name = corporate_name
            else:
                name = "Pending"
            tiv = loc.get("tiv", 0)
            master_locations.append({
                "name": name,
                "address": loc.get("address", ""),
                "city": loc.get("city", ""),
                "state": loc.get("state", ""),
                "tiv": tiv,
                "on_property": _is_on_property(addr_key),
                "on_liability": _is_on_liability(addr_key),
            })
            seen_addr_keys.add(addr_key)
    
    # Second: extracted locations not already in SOV (skip vacant land)
    for loc in locations:
        desc_check = (loc.get("description", "") or loc.get("corporate_entity", "") or "").lower()
        if "vacant" in desc_check or ("land" in desc_check and "hotel" not in desc_check):
            # Track but skip
            addr_key = (_normalize_addr(loc.get("address", "")) + "|" +
                       _normalize_city(loc.get("city", "")) + "|" +
                       loc.get("state", "").strip().upper())
            seen_addr_keys.add(addr_key)
            continue
        addr_key = (_normalize_addr(loc.get("address", "")) + "|" +
                   _normalize_city(loc.get("city", "")) + "|" +
                   loc.get("state", "").strip().upper())
        if addr_key not in seen_addr_keys and loc.get("address"):
            name = loc.get("description", "") or loc.get("corporate_entity", "")
            if not name or not name.strip():
                name = "Pending"
            master_locations.append({
                "name": name,
                "address": loc.get("address", ""),
                "city": loc.get("city", ""),
                "state": loc.get("state", ""),
                "tiv": 0,
                "on_property": _is_on_property(addr_key),
                "on_liability": _is_on_liability(addr_key),
            })
            seen_addr_keys.add(addr_key)
    
    # Third: GL schedule_of_classes locations not already in master list
    # This catches liability-only locations (e.g., LaPlace, vacant land) that aren't on SOV or extracted locations
    import re
    gl_seen_addrs = set()  # Deduplicate GL entries (multiple classes per location)
    
    for entry in gl_classes:
        if not isinstance(entry, dict):
            continue
        addr = entry.get("address", "")
        if not addr:
            continue
        # Skip vacant land entries from GL schedule
        classification = (entry.get("classification", "") or "").lower()
        if "vacant" in classification or ("land" in classification and "hotel" not in classification):
            continue
        # Parse address - may contain "Street, City, ST ZIP" or just street
        addr_norm = _normalize_addr(addr)
        if addr_norm in gl_seen_addrs:
            continue
        gl_seen_addrs.add(addr_norm)
        
        # Check if this address is already in the master list using fuzzy matching
        already_in = False
        for ml in master_locations:
            ml_addr_norm = _normalize_addr(ml.get("address", ""))
            if _fuzzy_addr_match(addr_norm, ml_addr_norm):
                # Mark existing location as on_liability if not already
                ml["on_liability"] = True
                already_in = True
                break
        
        if not already_in:
            # Try to parse city/state from the address string (e.g., "4285 Highway 51, LaPlace, LA 70068")
            parts = [p.strip() for p in addr.split(",")]
            street = parts[0] if parts else addr
            city = ""
            state = ""
            if len(parts) >= 3:
                street = parts[0]
                city = parts[1]
                st_match = re.match(r'([A-Z]{2})\s*\d*', parts[2].strip())
                if st_match:
                    state = st_match.group(1)
            elif len(parts) == 2:
                street = parts[0]
                st_match = re.match(r'([A-Z]{2})\s*\d*', parts[1].strip())
                if st_match:
                    state = st_match.group(1)
                else:
                    city = parts[1]
            
            brand = entry.get("brand_dba", "") or entry.get("classification", "")
            if not brand or not brand.strip():
                brand = "Pending"
            
            # Check fuzzy match against seen_addr_keys too
            addr_key = (_normalize_addr(street) + "|" +
                       _normalize_city(city) + "|" +
                       state.strip().upper())
            key_already_seen = False
            for existing_key in seen_addr_keys:
                existing_parts = existing_key.split("|")
                new_parts = addr_key.split("|")
                if len(existing_parts) == 3 and len(new_parts) == 3:
                    # Relax city matching: if street addresses match, consider same location
                    # (city names can differ: "LA PLACE" vs "LAPLACE")
                    state_match = (not existing_parts[2] or not new_parts[2] or existing_parts[2] == new_parts[2])
                    if _fuzzy_addr_match(existing_parts[0], new_parts[0]) and state_match:
                        key_already_seen = True
                        break
            
            if not key_already_seen:
                # Cross-reference SOV for property name only (NOT TIV — TIV is property-only)
                if sov_data and sov_data.get("locations"):
                    for sov_loc in sov_data["locations"]:
                        if _fuzzy_addr_match(_normalize_addr(street), _normalize_addr(sov_loc.get("address", ""))):
                            # Get name from SOV if brand is generic
                            if brand in ("Pending", "") or brand == entry.get("classification", ""):
                                _cn = (sov_loc.get("corporate_name", "") or "").strip()
                                _db = (sov_loc.get("dba", "") or sov_loc.get("hotel_flag", "") or "").strip()
                                if _cn and _db:
                                    brand = f"{_cn} - {_db}"
                                elif _db:
                                    brand = _db
                                elif _cn:
                                    brand = _cn
                            break
                master_locations.append({
                    "name": brand,
                    "address": street,
                    "city": city,
                    "state": state,
                    "tiv": 0,  # TIV only comes from property SOV, never from GL
                    "on_property": _is_on_property(addr_key),
                    "on_liability": True,
                })
                seen_addr_keys.add(addr_key)
    
    # Fourth: Extract locations from GL forms_endorsements (e.g., CG 21 44 designated premises)
    # These forms often list ALL covered locations with full addresses
    gl_forms = gl_cov.get("forms_endorsements", []) if isinstance(gl_cov, dict) else []
    for form in gl_forms:
        if not isinstance(form, dict):
            continue
        desc = (form.get("description", "") or "").upper()
        # Look for designated premises forms that contain address lists
        if not any(kw in desc for kw in ["DESIGNATED PREMISES", "CG 21 44", "CG2144", "NXLL110", "NXLL 110", "LIMITATION OF COVERAGE"]):
            continue
        # Try to extract addresses from the description text
        # Format: "1) 4285 Highway 51, LaPlace, LA 70068" or similar numbered lists
        addr_pattern = re.findall(r'\d+\)\s*(.+?)(?=\s*\d+\)|$)', desc, re.DOTALL)
        if not addr_pattern:
            # Try semicolon or newline separated
            addr_pattern = [a.strip() for a in re.split(r'[;\n]', desc) if re.search(r'\d+\s+\w+', a.strip())]
        for raw_addr in addr_pattern:
            raw_addr = raw_addr.strip().rstrip(',')
            if not raw_addr or len(raw_addr) < 5:
                continue
            addr_norm = _normalize_addr(raw_addr)
            # Check if already in master list
            already_exists = False
            for ml in master_locations:
                if _fuzzy_addr_match(addr_norm, _normalize_addr(ml.get("address", ""))):
                    ml["on_liability"] = True
                    already_exists = True
                    break
            if not already_exists:
                parts = [p.strip() for p in raw_addr.split(",")]
                street = parts[0] if parts else raw_addr
                city = ""
                state = ""
                if len(parts) >= 3:
                    city = parts[1]
                    st_m = re.match(r'([A-Z]{2})\s*\d*', parts[2].strip().upper())
                    if st_m: state = st_m.group(1)
                elif len(parts) == 2:
                    st_m = re.match(r'([A-Z]{2})\s*\d*', parts[1].strip().upper())
                    if st_m:
                        state = st_m.group(1)
                    else:
                        city = parts[1]
                # Try to find name from SOV data
                loc_name = "Pending"
                if sov_data and sov_data.get("locations"):
                    for sov_loc in sov_data["locations"]:
                        if _fuzzy_addr_match(_normalize_addr(street), _normalize_addr(sov_loc.get("address", ""))):
                            # Use Corporate Name - DBA format
                            _cn = (sov_loc.get("corporate_name", "") or "").strip()
                            _db = (sov_loc.get("dba", "") or sov_loc.get("hotel_flag", "") or "").strip()
                            if _cn and _db:
                                loc_name = f"{_cn} - {_db}"
                            elif _db:
                                loc_name = _db
                            elif _cn:
                                loc_name = _cn
                            if not city and sov_loc.get("city"): city = sov_loc["city"]
                            if not state and sov_loc.get("state"): state = sov_loc["state"]
                            break
                addr_key = (_normalize_addr(street) + "|" + _normalize_city(city) + "|" + state.strip().upper())
                key_already_seen = False
                for existing_key in seen_addr_keys:
                    ep = existing_key.split("|")
                    np = addr_key.split("|")
                    if len(ep) == 3 and len(np) == 3:
                        if _fuzzy_addr_match(ep[0], np[0]) and ep[2] == np[2]:
                            key_already_seen = True
                            break
                if not key_already_seen:
                    master_locations.append({
                        "name": loc_name,
                        "address": street,
                        "city": city,
                        "state": state,
                        "tiv": 0,
                        "on_property": _is_on_property(addr_key),
                        "on_liability": True,
                    })
                    seen_addr_keys.add(addr_key)
    
    if master_locations:
        CHECK = "\u2713"  # Unicode checkmark
        DASH = "\u2014"   # Em-dash for missing coverage (rendered in RED)
        headers = ["#", "Property Name", "Address", "City", "ST", "TIV", "Property", "Liability"]
        rows = []
        total_tiv = 0
        # Track which rows have missing coverages for RED formatting
        missing_property_rows = []  # row indices (0-based in rows list)
        missing_liability_rows = []
        
        for i, loc in enumerate(master_locations, 1):
            tiv_val = loc["tiv"] or 0
            total_tiv += tiv_val
            prop_cell = CHECK if loc["on_property"] else DASH
            liab_cell = CHECK if loc["on_liability"] else DASH
            if not loc["on_property"]:
                missing_property_rows.append(len(rows))  # current row index
            if not loc["on_liability"]:
                missing_liability_rows.append(len(rows))
            rows.append([
                str(i),
                loc["name"],
                loc["address"],
                loc["city"],
                loc["state"],
                fmt_currency(tiv_val) if tiv_val else "",
                prop_cell,
                liab_cell,
            ])
        
        # Add totals row
        rows.append([
            "", "TOTAL", "", "", "",
            fmt_currency(total_tiv) if total_tiv else "",
            "", ""
        ])
        
        L = WD_ALIGN_PARAGRAPH.LEFT
        C = WD_ALIGN_PARAGRAPH.CENTER
        R = WD_ALIGN_PARAGRAPH.RIGHT
        table = create_styled_table(doc, headers, rows,
                          col_widths=[0.3, 1.5, 1.6, 0.9, 0.3, 1.0, 0.7, 0.7],
                          header_size=8, body_size=8,
                          header_alignments={0: L, 1: L, 2: L, 3: L, 4: L, 5: R, 6: C, 7: C},
                          col_alignments={5: R, 6: C, 7: C})
        
        # Apply RED color to missing coverage cells (em-dashes)
        RED = RGBColor(0xCC, 0x00, 0x00)
        for row_idx in missing_property_rows:
            cell = table.rows[row_idx + 1].cells[6]  # +1 for header row
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.color.rgb = RED
                    run.font.bold = True
        for row_idx in missing_liability_rows:
            cell = table.rows[row_idx + 1].cells[7]  # +1 for header row
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.color.rgb = RED
                    run.font.bold = True
        
        # Legend
        add_formatted_paragraph(doc, "", size=4)
        legend_p = doc.add_paragraph()
        legend_p.paragraph_format.space_before = Pt(2)
        legend_p.paragraph_format.space_after = Pt(2)
        run_check = legend_p.add_run("\u2713")
        run_check.font.size = Pt(8)
        run_check.font.color.rgb = RGBColor(0x00, 0x80, 0x00)  # Green
        run_text = legend_p.add_run(" = Covered     ")
        run_text.font.size = Pt(8)
        run_text.font.color.rgb = CHARCOAL
        run_dash = legend_p.add_run(DASH)
        run_dash.font.size = Pt(8)
        run_dash.font.color.rgb = RED
        run_dash.font.bold = True
        run_text2 = legend_p.add_run(" = Not Currently Quoted")
        run_text2.font.size = Pt(8)
        run_text2.font.color.rgb = CHARCOAL
        
        # Add note about SOV
        if sov_data and sov_data.get("locations"):
            add_formatted_paragraph(doc, "", size=6)
            add_formatted_paragraph(doc, "See attached Statement of Values for complete property details.",
                                  size=9, italic=True, color=CHARCOAL)
    else:
        add_formatted_paragraph(doc, "Location schedule to be confirmed.", size=11)


def generate_coverage_section(doc, data, coverage_key, display_name):
    """Generate a coverage section (Property, GL, Umbrella, WC, Auto)."""
    coverages = data.get("coverages", {})
    cov = coverages.get(coverage_key)
    if not cov:
        return
    
    # Skip phantom coverage sections with no meaningful data
    carrier = cov.get("carrier", "") or ""
    premium = cov.get("premium", 0) or 0
    total_premium = cov.get("total_premium", 0) or 0
    limits = cov.get("coverage_limits", []) or []
    if not carrier.strip() and not premium and not total_premium and not limits:
        logger.info(f"Skipping phantom coverage section: {coverage_key} (no carrier, premium, or limits)")
        return
    
    add_page_break(doc)
    add_section_header(doc, display_name)
    
    # Coverage Summary table
    carrier = _clean_carrier_name(cov.get("carrier", "N/A"))
    admitted = "Admitted" if cov.get("carrier_admitted", True) else "Non-Admitted"
    am_best = cov.get("am_best_rating", "N/A")
    # Fallback to lookup table if not provided in quote
    if not am_best or am_best == "N/A":
        looked_up = lookup_am_best(carrier)
        if looked_up:
            am_best = looked_up
    
    add_subsection_header(doc, "Coverage Summary")
    
    carrier_rows = [
        ["Carrier", carrier],
        ["Admitted Status", admitted],
        ["AM Best Rating", am_best],
    ]
    
    # Add wholesaler if present
    if cov.get("wholesaler"):
        carrier_rows.append(["Wholesaler", cov["wholesaler"]])
    
    # Add policy form if present
    if cov.get("policy_form"):
        carrier_rows.append(["Policy Form", cov["policy_form"]])
    
    # Add policy period if present
    if cov.get("policy_period"):
        carrier_rows.append(["Policy Period", cov["policy_period"]])
    
    # Add layer description for excess property
    layer_desc = cov.get("layer_description", "")
    if layer_desc and coverage_key in ("excess_property", "excess_property_2"):
        carrier_rows.append(["Layer", layer_desc])
    
    # Add TIV if present (for property coverages)
    tiv = cov.get("tiv", "")
    if tiv and coverage_key in ("property", "excess_property", "excess_property_2"):
        carrier_rows.append(["Total Insured Value", tiv])
    
    # Add GL deductible if present
    gl_ded = cov.get("gl_deductible", "")
    if gl_ded and gl_ded not in ("$0", "None", "N/A", "", "0"):
        carrier_rows.append(["Deductible", gl_ded])
    
    # Add defense basis if present
    defense = cov.get("defense_basis", "")
    if defense and defense not in ("N/A", ""):
        carrier_rows.append(["Defense Basis", defense])
    
    L = WD_ALIGN_PARAGRAPH.LEFT
    create_styled_table(doc, ["Item", "Details"], carrier_rows, col_widths=[2.5, 5.0],
                       header_size=10, body_size=10,
                       header_alignments={0: L, 1: L})
    
    # Schedule of Values (Property) - prefer SOV data if available
    sov_data = data.get("sov_data")
    sov_from_quote = cov.get("schedule_of_values", [])
    
    if coverage_key == "property" and sov_data and sov_data.get("locations"):
        # Use SOV spreadsheet data for detailed Schedule of Values
        add_subsection_header(doc, "Schedule of Values")
        sov_locs = sov_data["locations"]
        # Check if any location has "other_value" data (Sign, Pool, Other)
        has_other = any(loc.get("other_value", 0) for loc in sov_locs)
        if has_other:
            headers = ["#", "Location", "Building", "Contents", "Other", "BI/Rents", "TIV"]
        else:
            headers = ["#", "Location", "Building", "Contents", "BI/Rents", "TIV"]
        rows = []
        total_other = 0
        for i, loc in enumerate(sov_locs, 1):
            name = loc.get("dba") or loc.get("hotel_flag") or loc.get("corporate_name", "")
            addr = f"{loc.get('address', '')}, {loc.get('city', '')}, {loc.get('state', '')}"
            loc_label = f"{name}\n{addr}" if name else addr
            other_val = loc.get("other_value", 0) or 0
            total_other += other_val
            if has_other:
                rows.append([
                    str(i),
                    loc_label,
                    fmt_currency(loc.get("building_value", 0)),
                    fmt_currency(loc.get("contents_value", 0)),
                    fmt_currency(other_val) if other_val else "\u2014",
                    fmt_currency(loc.get("bi_value", 0)),
                    fmt_currency(loc.get("tiv", 0))
                ])
            else:
                rows.append([
                    str(i),
                    loc_label,
                    fmt_currency(loc.get("building_value", 0)),
                    fmt_currency(loc.get("contents_value", 0)),
                    fmt_currency(loc.get("bi_value", 0)),
                    fmt_currency(loc.get("tiv", 0))
                ])
        # Add totals row
        totals = sov_data.get("totals", {})
        if has_other:
            rows.append([
                "", "TOTAL",
                fmt_currency(totals.get("building_value", 0)),
                fmt_currency(totals.get("contents_value", 0)),
                fmt_currency(total_other) if total_other else "",
                fmt_currency(totals.get("bi_value", 0)),
                fmt_currency(totals.get("tiv", 0))
            ])
            create_styled_table(doc, headers, rows,
                              col_widths=[0.3, 2.0, 1.1, 0.9, 0.8, 0.9, 1.1],
                              header_size=9, body_size=8,
                              col_alignments={2: WD_ALIGN_PARAGRAPH.CENTER, 3: WD_ALIGN_PARAGRAPH.CENTER,
                                             4: WD_ALIGN_PARAGRAPH.CENTER, 5: WD_ALIGN_PARAGRAPH.CENTER,
                                             6: WD_ALIGN_PARAGRAPH.CENTER})
        else:
            rows.append([
                "", "TOTAL",
                fmt_currency(totals.get("building_value", 0)),
                fmt_currency(totals.get("contents_value", 0)),
                fmt_currency(totals.get("bi_value", 0)),
                fmt_currency(totals.get("tiv", 0))
            ])
            create_styled_table(doc, headers, rows,
                              col_widths=[0.3, 2.2, 1.2, 1.0, 1.0, 1.3],
                              header_size=9, body_size=8,
                              col_alignments={2: WD_ALIGN_PARAGRAPH.CENTER, 3: WD_ALIGN_PARAGRAPH.CENTER,
                                             4: WD_ALIGN_PARAGRAPH.CENTER, 5: WD_ALIGN_PARAGRAPH.CENTER})
    elif sov_from_quote:
        add_subsection_header(doc, "Schedule of Values")
        headers = ["Location", "Building", "Contents", "BI/Rents", "TIV"]
        rows = [[
            s.get("location", ""),
            fmt_currency(s.get("building", 0)),
            fmt_currency(s.get("contents", 0)),
            fmt_currency(s.get("business_income", 0)),
            fmt_currency(s.get("tiv", 0))
        ] for s in sov_from_quote]
        create_styled_table(doc, headers, rows,
                          col_widths=[2.0, 1.4, 1.2, 1.2, 1.2],
                          header_size=9, body_size=9)
    
    # Crime Insuring Clauses (3-column: Clause, Limit, Retention)
    insuring_clauses = cov.get("insuring_clauses", [])
    if coverage_key == "crime" and insuring_clauses:
        add_subsection_header(doc, "Insuring Clauses")
        headers = ["Insuring Clause", "Limit", "Retention"]
        rows = []
        for clause in insuring_clauses:
            if isinstance(clause, dict):
                rows.append([
                    clause.get("description", ""),
                    clause.get("limit", ""),
                    clause.get("retention", clause.get("deductible", ""))
                ])
            else:
                rows.append([str(clause), "", ""])
        L = WD_ALIGN_PARAGRAPH.LEFT
        R = WD_ALIGN_PARAGRAPH.RIGHT
        create_styled_table(doc, headers, rows, col_widths=[4.0, 1.5, 1.5],
                           header_size=10, body_size=9,
                           header_alignments={0: L, 1: R, 2: R},
                           col_alignments={1: R, 2: R})
    
    # Limits (non-crime coverages, or crime without insuring_clauses)
    limits = cov.get("limits", [])
    if limits and not (coverage_key == "crime" and insuring_clauses):
        add_subsection_header(doc, "Coverage Limits")
        headers = ["Description", "Limit"]
        rows = [[lim.get("description", ""), lim.get("limit", "")] if isinstance(lim, dict) else [str(lim), ""] for lim in limits]
        # Left-align headers; center Limit values for umbrella/excess
        L = WD_ALIGN_PARAGRAPH.LEFT
        limit_body_align = {}
        if coverage_key in ("umbrella", "cyber", "epli", "flood", "terrorism"):
            limit_body_align = {1: WD_ALIGN_PARAGRAPH.CENTER}
        create_styled_table(doc, headers, rows, col_widths=[4.5, 3.0],
                           header_size=10, body_size=10,
                           header_alignments={0: L, 1: L},
                           col_alignments=limit_body_align)
    
    # Deductibles (Property)
    deductibles = cov.get("deductibles", [])
    if deductibles:
        add_subsection_header(doc, "Deductibles")
        headers = ["Peril", "Deductible"]
        rows = [[ded.get("description", ""), ded.get("amount", "")] if isinstance(ded, dict) else [str(ded), ""] for ded in deductibles]
        L = WD_ALIGN_PARAGRAPH.LEFT
        create_styled_table(doc, headers, rows, col_widths=[4.5, 3.0],
                           header_size=10, body_size=10,
                           header_alignments={0: L, 1: L})
    
    # Coinsurance & Valuation (Property)
    coinsurance = cov.get("coinsurance", [])
    valuation = cov.get("valuation", "")
    if coinsurance or valuation:
        add_subsection_header(doc, "Coinsurance & Valuation")
        headers = ["Coverage", "Coinsurance / Limitation"]
        rows = []
        for ci in coinsurance:
            if isinstance(ci, dict):
                cov_name = ci.get("coverage", "")
                pct = ci.get("percentage", "")
                limitation = ci.get("limitation", "")
                val = limitation if limitation else pct
                if val:
                    rows.append([cov_name, val])
        if valuation:
            rows.append(["Valuation", valuation])
        if rows:
            L = WD_ALIGN_PARAGRAPH.LEFT
            create_styled_table(doc, headers, rows, col_widths=[4.5, 3.0],
                               header_size=10, body_size=10,
                               header_alignments={0: L, 1: L})
    
    # Layer Description (Excess Property)
    layer_desc = cov.get("layer_description", "")
    if layer_desc and coverage_key in ("excess_property", "excess_property_2"):
        add_formatted_paragraph(doc, f"Layer: {layer_desc}", size=11, bold=True, space_after=6)
    
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
        ] if isinstance(h, dict) else [str(h), "", "", "", ""] for h in hazards]
        create_styled_table(doc, headers, rows,
                          col_widths=[1.5, 2.5, 0.8, 1.0, 1.2],
                          header_size=9, body_size=9)
    
    # Schedule of Classes (GL - location exposures)
    classes = cov.get("schedule_of_classes", [])
    if classes:
        add_subsection_header(doc, "Exposures")
        from docx.enum.text import WD_ALIGN_PARAGRAPH as WD_ALIGN
        # Check if we have address/brand data for the enhanced format
        has_address = any(c.get("address") or c.get("brand_dba") for c in classes if isinstance(c, dict))
        # Check if we have class codes and rates (class-code-based format like AmTrust)
        has_class_code = any(c.get("class_code") for c in classes if isinstance(c, dict))
        has_rate = any(c.get("rate") for c in classes if isinstance(c, dict))
        
        if has_address:
            # Enhanced format with address/brand: Address | Brand/DBA | Code | Classification | Rate | Exposure
            if has_class_code or has_rate:
                headers = ["Address", "Brand / DBA", "Code", "Classification", "Rate", "Exposure"]
                rows = []
                for c in classes:
                    if not isinstance(c, dict):
                        continue
                    addr = c.get("address", "")
                    brand = c.get("brand_dba", "")
                    if not addr and c.get("location"):
                        addr = c.get("location", "")
                    classification = c.get("classification", "")
                    class_code = c.get("class_code", "")
                    rate = c.get("rate", "")
                    exposure_basis = c.get("exposure_basis", "")
                    exposure = c.get("exposure", "")
                    # Format exposure with basis
                    if exposure_basis and exposure:
                        exposure_str = f"{exposure} ({exposure_basis})"
                    else:
                        exposure_str = str(exposure) if exposure else ""
                    rows.append([addr, brand, str(class_code), classification, str(rate), exposure_str])
                create_styled_table(doc, headers, rows,
                                  col_widths=[1.8, 1.2, 0.6, 1.5, 0.6, 1.3],
                                  header_size=8, body_size=7,
                                  col_alignments={4: WD_ALIGN.RIGHT, 5: WD_ALIGN.RIGHT})
            else:
                # Address-based without class codes: Address | Brand/DBA | Classification | Exposure | Premium
                headers = ["Address", "Brand / DBA", "Classification", "Exposure", "Premium"]
                rows = []
                for c in classes:
                    if not isinstance(c, dict):
                        continue
                    addr = c.get("address", "")
                    brand = c.get("brand_dba", "")
                    if not addr and c.get("location"):
                        addr = c.get("location", "")
                    classification = c.get("classification", "")
                    exposure_basis = c.get("exposure_basis", "")
                    exposure = c.get("exposure", "")
                    premium = c.get("premium", "")
                    if exposure_basis and exposure:
                        exposure_str = f"{exposure} ({exposure_basis})"
                    else:
                        exposure_str = str(exposure) if exposure else ""
                    rows.append([addr, brand, classification, exposure_str, str(premium) if premium else ""])
                create_styled_table(doc, headers, rows,
                                  col_widths=[2.0, 1.5, 1.5, 1.3, 1.0],
                                  header_size=8, body_size=8,
                                  col_alignments={3: WD_ALIGN.RIGHT, 4: WD_ALIGN.RIGHT})
        else:
            # Class-code-based format (no addresses): Code | Classification | Rate | Exposure | Basis
            headers = ["Code", "Classification", "Rate", "Exposure", "Exposure Basis"]
            rows = []
            for c in classes:
                if not isinstance(c, dict):
                    continue
                class_code = c.get("class_code", c.get("location", ""))
                classification = c.get("classification", "")
                rate = c.get("rate", "")
                exposure = c.get("exposure", "")
                exposure_basis = c.get("exposure_basis", "")
                rows.append([str(class_code), classification, str(rate), str(exposure), exposure_basis])
            create_styled_table(doc, headers, rows,
                              col_widths=[0.8, 2.5, 0.8, 1.5, 1.4],
                              header_size=9, body_size=8,
                              col_alignments={2: WD_ALIGN.RIGHT, 3: WD_ALIGN.RIGHT})
    
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
        create_styled_table(doc, headers, rows,
                          col_widths=[0.6, 1.5, 0.8, 2.0, 1.2, 0.9],
                          header_size=9, body_size=9)
    
    # Vehicle Schedule (Auto)
    vehicles = cov.get("vehicle_schedule", [])
    if vehicles:
        add_subsection_header(doc, "Vehicle Schedule")
        headers = ["Year", "Make", "Model", "VIN", "Garage Location"]
        rows = [[v.get("year", ""), v.get("make", ""), v.get("model", ""),
                 v.get("vin", ""), v.get("garage_location", "")] for v in vehicles]
        create_styled_table(doc, headers, rows,
                          col_widths=[0.6, 1.2, 1.2, 2.5, 2.0],
                          header_size=9, body_size=9)
    
    # Additional Coverages
    addl = cov.get("additional_coverages", [])
    if addl:
        # Use "Sublimits of Liability / Extensions" for property, "Additional Coverages" for others
        addl_title = "Sublimits of Liability / Extensions" if coverage_key == "property" else "Additional Coverages"
        add_subsection_header(doc, addl_title)
        has_ded = any(ac.get("deductible") for ac in addl if isinstance(ac, dict))
        L = WD_ALIGN_PARAGRAPH.LEFT
        if has_ded:
            headers = ["Description", "Limit", "Deductible"]
            rows = [[ac.get("description", ""), ac.get("limit", ""), ac.get("deductible", "")] if isinstance(ac, dict) else [str(ac), "", ""] for ac in addl]
            create_styled_table(doc, headers, rows, col_widths=[3.5, 2.0, 2.0],
                              header_alignments={0: L, 1: L, 2: L})
        else:
            headers = ["Description", "Limit"]
            rows = [[ac.get("description", ""), ac.get("limit", "")] if isinstance(ac, dict) else [str(ac), ""] for ac in addl]
            create_styled_table(doc, headers, rows, col_widths=[4.5, 3.0],
                              header_alignments={0: L, 1: L})
    
    # Underlying Insurance (Umbrella)
    underlying = cov.get("underlying_insurance", [])
    if underlying:
        add_subsection_header(doc, "Underlying Insurance")
        headers = ["Carrier", "Coverage", "Limits"]
        rows = [[u.get("carrier", ""), u.get("coverage", ""), u.get("limits", "")] for u in underlying]
        create_styled_table(doc, headers, rows, col_widths=[2.5, 2.5, 2.5],
                          col_alignments={2: WD_ALIGN_PARAGRAPH.CENTER})
    
    # Tower Structure (Umbrella)
    tower = cov.get("tower_structure", [])
    if tower:
        add_subsection_header(doc, "Umbrella Tower Structure")
        headers = ["Layer", "Carrier", "Limits", "Premium", "Total (incl. taxes/fees)"]
        rows = [[
            t.get("layer", ""),
            t.get("carrier", ""),
            t.get("limits", ""),
            fmt_currency(t.get("premium", 0)),
            fmt_currency(t.get("total_cost", 0))
        ] for t in tower]
        create_styled_table(doc, headers, rows,
                          col_widths=[0.8, 2.0, 1.5, 1.2, 1.5],
                          header_size=9, body_size=9)
    
    # Forms & Endorsements
    forms = cov.get("forms_endorsements", [])
    if forms:
        add_subsection_header(doc, "Forms & Endorsements")
        headers = ["Form Number", "Description"]
        rows = [[f.get("form_number", ""), f.get("description", "")] if isinstance(f, dict) else ["", str(f)] for f in forms]
        L = WD_ALIGN_PARAGRAPH.LEFT
        create_styled_table(doc, headers, rows, col_widths=[2.0, 5.5],
                           header_size=9, body_size=9,
                           header_alignments={0: L, 1: L})
    
    # NOTE: Sales / Exposure section removed — the Exposures table above already
    # renders the full schedule_of_classes data including addresses, classifications,
    # codes, rates, and exposure amounts. Having both was redundant.
    
    # Covered Locations (GL only) - backup list of liability locations from GL quote
    if coverage_key == "general_liability":
        import re as _re_gl
        gl_loc_list = []
        gl_loc_seen = set()
        # Source 1: designated_premises array (PRIMARY - extracted by GPT)
        for dp_addr in cov.get("designated_premises", []):
            if not isinstance(dp_addr, str) or not dp_addr.strip() or len(dp_addr.strip()) < 5:
                continue
            addr_norm = _normalize_addr(dp_addr)
            if addr_norm not in gl_loc_seen:
                gl_loc_seen.add(addr_norm)
                gl_loc_list.append({
                    "address": dp_addr.strip(),
                    "brand": "",
                    "classification": "",
                })
        # Source 2: schedule_of_classes addresses
        for entry in cov.get("schedule_of_classes", []):
            if isinstance(entry, dict):
                addr = entry.get("address", "") or entry.get("location", "")
                if addr and addr.strip():
                    addr_norm = _normalize_addr(addr)
                    if addr_norm not in gl_loc_seen:
                        gl_loc_seen.add(addr_norm)
                        brand = entry.get("brand_dba", "") or ""
                        classification = entry.get("classification", "") or ""
                        gl_loc_list.append({
                            "address": addr.strip(),
                            "brand": brand.strip(),
                            "classification": classification.strip(),
                        })
        # Source 3: CG2144/NXLL110 designated premises forms (fallback)
        for form in cov.get("forms_endorsements", []):
            if not isinstance(form, dict):
                continue
            desc = (form.get("description", "") or "").upper()
            if not any(kw in desc for kw in ["DESIGNATED PREMISES", "CG 21 44", "CG2144", "NXLL110", "NXLL 110", "LIMITATION OF COVERAGE"]):
                continue
            addr_pattern = _re_gl.findall(r'\d+\)\s*(.+?)(?=\s*\d+\)|$)', desc, _re_gl.DOTALL)
            if not addr_pattern:
                addr_pattern = [a.strip() for a in _re_gl.split(r'[;\n]', desc) if _re_gl.search(r'\d+\s+\w+', a.strip())]
            for raw_addr in addr_pattern:
                raw_addr = raw_addr.strip().rstrip(',')
                if not raw_addr or len(raw_addr) < 5:
                    continue
                addr_norm = _normalize_addr(raw_addr)
                if addr_norm not in gl_loc_seen:
                    gl_loc_seen.add(addr_norm)
                    gl_loc_list.append({
                        "address": raw_addr.strip(),
                        "brand": "",
                        "classification": "",
                    })
        
        # Cross-reference GL locations with SOV to pull Corporate Name - DBA
        sov_data = data.get("sov_data")
        if sov_data and sov_data.get("locations"):
            for gl_loc in gl_loc_list:
                if gl_loc["brand"]:  # Already has a brand, skip
                    continue
                gl_addr_norm = _normalize_addr(gl_loc["address"])
                for sov_loc in sov_data["locations"]:
                    sov_addr_norm = _normalize_addr(sov_loc.get("address", ""))
                    # Fuzzy match: check if street addresses match
                    if not gl_addr_norm or not sov_addr_norm:
                        continue
                    # Extract house numbers and street names for comparison
                    gl_num_m = _re_gl.match(r'^(\d+)\s+(.+)', gl_addr_norm)
                    sov_num_m = _re_gl.match(r'^(\d+)\s+(.+)', sov_addr_norm)
                    matched = False
                    if gl_num_m and sov_num_m:
                        if gl_num_m.group(1) == sov_num_m.group(1) and gl_num_m.group(2) == sov_num_m.group(2):
                            matched = True
                    elif gl_addr_norm == sov_addr_norm:
                        matched = True
                    elif gl_addr_norm in sov_addr_norm or sov_addr_norm in gl_addr_norm:
                        matched = True
                    if matched:
                        corp = (sov_loc.get("corporate_name", "") or "").strip()
                        dba = (sov_loc.get("dba", "") or sov_loc.get("hotel_flag", "") or "").strip()
                        if corp and dba:
                            gl_loc["brand"] = f"{corp} - {dba}"
                        elif dba:
                            gl_loc["brand"] = dba
                        elif corp:
                            gl_loc["brand"] = corp
                        break
        
        if gl_loc_list:
            add_subsection_header(doc, "Covered Locations")
            add_formatted_paragraph(doc, 
                "The following locations are covered under this General Liability policy "
                "as identified on the carrier quote:",
                size=9, italic=True, color=CHARCOAL)
            headers = ["#", "Address", "Corporate Name / DBA"]
            rows = []
            for i, loc in enumerate(gl_loc_list, 1):
                rows.append([
                    str(i),
                    loc["address"],
                    loc["brand"],
                ])
            L = WD_ALIGN_PARAGRAPH.LEFT
            create_styled_table(doc, headers, rows,
                              col_widths=[0.4, 3.5, 3.6],
                              header_size=9, body_size=8,
                              header_alignments={0: L, 1: L, 2: L})


def generate_confirmation_to_bind(doc, data):
    """Section 14: Confirmation to Bind Agreement"""
    add_page_break(doc)
    add_section_header(doc, "Confirmation to Bind Agreement")
    
    add_formatted_paragraph(doc,
        "By signing below, the undersigned authorized representative of the Applicant confirms "
        "the following statements and authorizes HUB International to bind the coverages as outlined "
        "in this proposal, subject to the terms and conditions of the respective policies.",
        size=10, space_after=6)
    
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
        "I have been offered Terrorism Risk Insurance Act (TRIA) coverage and have made my election as indicated in this proposal.",
        "I understand that additional policies are available and recommended which include Equipment Breakdown (power surges, electrical arcing, mechanical failure), Employment Practices Liability (excluded by the liability carrier), Pollution (claims such as mold and legionella), Cyber, Flood, Earthquake, Deductible Buy Downs, Sexual Abuse & Molestation. If you would like to get these options quoted please request in writing to the producer or account executive."
    ]
    
    for i, stmt in enumerate(statements, 1):
        add_formatted_paragraph(doc, f"{i}. {stmt}", size=9, space_after=2)
    
    # Earned premium / cancellation disclaimer - small font, bold, red
    _add_earned_premium_disclaimer(doc)
    
    # Signature block
    add_formatted_paragraph(doc, "", space_before=6)
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
        cell_label.width = Inches(2.0)
        p = cell_label.paragraphs[0]
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        run = p.add_run(label)
        run.font.size = Pt(10)
        run.font.color.rgb = CLASSIC_BLUE
        run.font.bold = True
        run.font.name = "Calibri"
        
        cell_val = sig_table.rows[i].cells[1]
        cell_val.width = Inches(5.0)
        p = cell_val.paragraphs[0]
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(4)
        pPr = p._p.get_or_add_pPr()
        pBdr = parse_xml(
            f'<w:pBdr {nsdecls("w")}>'
            f'<w:bottom w:val="single" w:sz="4" w:space="1" w:color="{CLASSIC_BLUE_HEX}"/>'
            f'</w:pBdr>'
        )
        pPr.append(pBdr)


# --- Earned Premium Disclaimer (shared between Payment Options and Confirmation to Bind) ---
_EARNED_PREMIUM_DISCLAIMER = (
    "All insurance policies, including but not limited to property, general liability, "
    "umbrella/excess liability, and ancillary coverages, may be subject to minimum earned premiums "
    "as determined by the issuing carrier. Property policies frequently carry hurricane or named storm "
    "minimum earned premiums, which may require a significant portion of the annual premium to be fully "
    "earned regardless of the policy\u2019s cancellation or replacement date. Liability and umbrella/excess "
    "policies may also carry minimum earned premium provisions that limit or eliminate premium refunds "
    "upon cancellation.\n\n"
    "Additionally, most policies \u2014 both admitted and non-admitted (surplus lines) \u2014 are subject to "
    "short rate cancellation penalties if cancelled mid-term at the insured\u2019s request. Policy fees, "
    "inspection fees, and membership or association fees are typically fully earned at inception and "
    "non-refundable regardless of cancellation.\n\n"
    "These provisions vary by carrier, program, and policy form. Clients should carefully consider the "
    "financial implications of any mid-term policy changes, cancellations, or carrier transitions, as "
    "premium refunds may be limited or unavailable. HUB recommends reviewing all earned premium and "
    "cancellation provisions with your service team prior to binding or making any policy changes."
)

def _add_earned_premium_disclaimer(doc):
    """Add the earned premium disclaimer in small bold red font."""
    RED = RGBColor(0xCC, 0x00, 0x00)
    for para_text in _EARNED_PREMIUM_DISCLAIMER.split("\n\n"):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(4)
        p.paragraph_format.space_after = Pt(2)
        run = p.add_run(para_text.strip())
        run.font.size = Pt(6.5)
        run.font.bold = True
        run.font.color.rgb = RED
        run.font.name = "Calibri"


def generate_electronic_consent(doc):
    """Section 15: Electronic Documents Consent"""
    add_page_break(doc)
    add_section_header(doc, "Electronic Documents Consent")
    
    add_formatted_paragraph(doc,
        "By signing below, I consent to receive insurance policy documents, endorsements, "
        "certificates of insurance, notices of cancellation, renewal notices, and other "
        "insurance-related documents electronically from HUB International and/or the "
        "insurance carriers providing coverage.",
        size=11, space_after=10)
    
    add_formatted_paragraph(doc,
        "I understand that I may withdraw this consent at any time by providing written "
        "notice to my HUB International representative, and that I may request paper "
        "copies of any documents at no additional charge.",
        size=11, space_after=15)
    
    # Signature line
    sig_table = doc.add_table(rows=3, cols=2)
    for i, label in enumerate(["Authorized Signature:", "Printed Name:", "Date:"]):
        cell = sig_table.rows[i].cells[0]
        cell.width = Inches(2.0)
        p = cell.paragraphs[0]
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(8)
        run = p.add_run(label)
        run.font.size = Pt(11)
        run.font.color.rgb = CLASSIC_BLUE
        run.font.bold = True
        run.font.name = "Calibri"
        
        cell_val = sig_table.rows[i].cells[1]
        cell_val.width = Inches(5.0)
        p = cell_val.paragraphs[0]
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(8)
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
    add_section_header(doc, "Carrier Ratings Summary")
    
    add_formatted_paragraph(doc,
        "AM Best is a full-service credit rating organization dedicated to serving the insurance "
        "industry. AM Best ratings provide an independent third-party evaluation of an insurer's "
        "financial strength and ability to meet its ongoing insurance policy and contract obligations.",
        size=11, space_after=10)
    
    # Build carrier rating table from all coverages
    coverages = data.get("coverages", {})
    carriers_seen = {}
    coverage_names = {
        "property": "Property",
        "excess_property": "Excess Property (Layer 1)",
        "excess_property_2": "Excess Property (Layer 2)",
        "general_liability": "General Liability",
        "umbrella": "Umbrella / Excess",
        "umbrella_layer_2": "2nd Excess Layer",
        "umbrella_layer_3": "3rd Excess Layer",
        "workers_comp": "Workers Compensation",
        "workers_compensation": "Workers Compensation",
        "commercial_auto": "Commercial Auto",
        "terrorism": "Terrorism / TRIA",
        "cyber": "Cyber Liability",
        "epli": "Employment Practices Liability",
        "crime": "Crime",
        "flood": "Flood",
        "inland_marine": "Inland Marine",
        "equipment_breakdown": "Equipment Breakdown",
        "liquor_liability": "Liquor Liability",
        "innkeepers_liability": "Innkeepers Liability",
        "environmental": "Environmental / Pollution",
        "workplace_violence": "Workplace Violence",
        "garage_keepers": "Garage Keepers",
        "enviro_pack": "Enviro Pack",
    }
    
    for key, display_name in coverage_names.items():
        cov = coverages.get(key)
        if cov:
            carrier = _clean_carrier_name(cov.get("carrier", ""))
            if carrier and carrier not in carriers_seen:
                rating = cov.get("am_best_rating", "N/A")
                if not rating or rating == "N/A":
                    looked_up = lookup_am_best(carrier)
                    if looked_up:
                        rating = looked_up
                carriers_seen[carrier] = {
                    "rating": rating,
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
        create_styled_table(doc, headers, rows,
                          col_widths=[2.5, 1.2, 1.3, 2.5],
                          header_size=10, body_size=10)


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
        add_formatted_paragraph(doc, text, size=10, space_after=6)


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
        ("Mold", "Mold damage is typically excluded or subject to limited coverage under standard property policies."),
        ("Terrorism", "Coverage for acts of terrorism as defined by the Terrorism Risk Insurance Act (TRIA). See TRIA Disclosure section for details."),
    ]
    
    for term, definition in definitions:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.line_spacing = Pt(13)
        run_term = p.add_run(f"{term}: ")
        run_term.font.size = Pt(10)
        run_term.font.color.rgb = ELECTRIC_BLUE
        run_term.font.bold = True
        run_term.font.name = "Calibri"
        run_def = p.add_run(definition)
        run_def.font.size = Pt(10)
        run_def.font.color.rgb = CLASSIC_BLUE
        run_def.font.name = "Calibri"


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
        size=10, space_after=10)
    
    add_subsection_header(doc, "Commission income")
    add_formatted_paragraph(doc,
        "Commission, normally calculated as a percentage of the premium paid to the insurer for the "
        "specific policy, is paid to us by the insurer to distribute and service your insurance policy. "
        "Our commission is included in the premium paid by you. The individuals at HUB International "
        "who place and service your insurance may be paid compensation that varies directly with the "
        "commissions we receive.",
        size=10, space_after=6)
    
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
        size=10, space_after=6)
    
    add_formatted_paragraph(doc,
        "Please also feel free to ask any questions about our compensation generally, or as to your "
        "specific insurance proposal or placement, by contacting your HUB broker or customer service "
        "representative directly, or by calling our client hotline at 1-866-857-4073.",
        size=10, space_after=10)
    
    add_subsection_header(doc, "Privacy Policy")
    add_formatted_paragraph(doc,
        "To view our privacy policy, please visit: www.hubinternational.com/about-us/privacy-policy/",
        size=10)


def generate_hub_advantage(doc):
    """Section 20: The HUB Advantage"""
    add_page_break(doc)
    add_section_header(doc, "Our Commitment — The HUB Advantage")
    
    add_formatted_paragraph(doc,
        "HUB International is dedicated to maintaining and upholding the highest standards of ethical "
        "conduct and integrity in all of our dealings with you, our client. We want to be your trusted "
        "risk advisor, and as such, we need to earn your confidence. So we are making a promise. We call "
        "it The HUB Advantage. Our mission is to make the advantage yours — and this is our commitment.",
        size=10, space_after=10)
    
    commitments = [
        "We strive to secure the most favorable terms from insurers, taking into account all of the circumstances — the risk you need to insure, the cost of insurance, the financial condition of the insurer, the insurer's reputation for service, and any other factors that are specific to your situation.",
        "We are open and honest as to how we are paid for placing your insurance. Our answers to your questions will be forthright and understandable. When we intend to seek a fixed fee for our efforts, we will disclose it to you in writing and obtain your approval prior to coverage being bound.",
        "We comply with the laws of every jurisdiction in which we operate, including those that apply to how insurance brokerages and agencies are paid. If the laws change, we will respond in a timely and appropriate manner.",
    ]
    
    for commitment in commitments:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = Pt(13)
        run = p.add_run(f"• {commitment}")
        run.font.size = Pt(10)
        run.font.color.rgb = CLASSIC_BLUE
        run.font.name = "Calibri"
    
    add_formatted_paragraph(doc,
        "We take our responsibility to our customers very seriously. If at any time you feel that we are "
        "not fulfilling your expectations — that we are not meeting our Client Commitment — please contact "
        "your account executive or call our toll free client hotline at 1-866-857-4073, and your concerns "
        "will be addressed as soon as possible.",
        size=10, space_before=10, space_after=10)
    
    add_formatted_paragraph(doc, "The HUB Advantage", size=14, color=ELECTRIC_BLUE, bold=True,
                           alignment=WD_ALIGN_PARAGRAPH.CENTER)
    add_formatted_paragraph(doc, "The privilege is ours, but the advantage is yours.", size=12,
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
        size=10, space_after=10)
    
    add_formatted_paragraph(doc,
        'Under the coverage, any losses resulting from certified acts of terrorism may be partially '
        'reimbursed by the United States Government under a formula established by the Terrorism Risk '
        'Insurance Act, as amended. However, your policy may contain other exclusions which might affect '
        'the coverage, such as an exclusion for nuclear events. Under the formula, the United States '
        'Government generally reimburses 80% beginning on January 1, 2020 of covered terrorism losses '
        'exceeding the statutorily established deductible paid by the insurance company providing the '
        'coverage. The Terrorism Risk Insurance Act, as amended, contains a $100 billion cap that limits '
        'United States government reimbursement as well as insurers\' liability for losses resulting from '
        'certified acts of terrorism when the amount of such losses exceed $100 billion in any one calendar '
        'year. If the aggregate insured losses for all insured exceed $100 billion, your coverage may be reduced.',
        size=10)


def generate_california_licenses(doc):
    """Section 22: California Licenses"""
    add_page_break(doc)
    add_section_header(doc, "California Licenses")
    
    add_formatted_paragraph(doc,
        "The following HUB International entities are licensed in the State of California:",
        size=11, space_after=8)
    
    headers = ["Entity Name", "License Number"]
    rows = [[name, lic] for name, lic in CA_LICENSES]
    create_styled_table(doc, headers, rows, col_widths=[5.5, 2.0],
                       header_size=9, body_size=8)


def generate_coverage_recommendations(doc):
    """Section 23: Coverage Recommendations"""
    add_page_break(doc)
    add_section_header(doc, "Coverage Recommendations")
    
    add_formatted_paragraph(doc,
        "HUB International recommends that you consider the following coverages and risk management "
        "strategies to protect your hospitality business. These recommendations are based on our "
        "extensive experience in the hotel and hospitality insurance industry.",
        size=11, space_after=10)
    
    recommendations = [
        ("Umbrella/Excess Liability", "We recommend maintaining umbrella or excess liability coverage with limits adequate to protect your assets. Hotels face unique liability exposures including guest injuries, swimming pool incidents, and food service operations."),
        ("Cyber Liability", "Hotels collect and store sensitive guest information including credit card numbers, personal identification, and travel details. Cyber liability coverage protects against data breaches, ransomware attacks, and regulatory fines."),
        ("Employment Practices Liability (EPLI)", "Hotels employ large numbers of workers in various roles, creating exposure to employment-related claims including wrongful termination, discrimination, harassment, and wage disputes."),
        ("Crime/Employee Dishonesty", "Hotels handle significant amounts of cash and guest valuables. Crime coverage protects against employee theft, forgery, computer fraud, and funds transfer fraud."),
        ("Flood Insurance", "Standard property policies exclude flood damage. If your properties are located in flood-prone areas, we strongly recommend purchasing separate flood coverage."),
        ("Equipment Breakdown", "Hotels rely heavily on mechanical and electrical equipment including HVAC systems, elevators, kitchen equipment, and laundry facilities."),
        ("Liquor Liability", "If your hotel serves alcohol, liquor liability coverage is essential. This coverage protects against claims arising from the sale, service, or furnishing of alcoholic beverages."),
    ]
    
    for title, text in recommendations:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(6)
        p.paragraph_format.line_spacing = Pt(13)
        run_title = p.add_run(f"{title}: ")
        run_title.font.size = Pt(10)
        run_title.font.color.rgb = ELECTRIC_BLUE
        run_title.font.bold = True
        run_title.font.name = "Calibri"
        run_text = p.add_run(text)
        run_text.font.size = Pt(10)
        run_text.font.color.rgb = CLASSIC_BLUE
        run_text.font.name = "Calibri"
    
    add_formatted_paragraph(doc, "", space_before=10)
    add_callout_box(doc,
        "Please discuss these recommendations with your HUB International representative to determine "
        "which coverages are appropriate for your specific operations and risk profile.")


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
    font.size = Pt(11)
    font.color.rgb = CLASSIC_BLUE
    
    # Set margins
    for section in doc.sections:
        section.top_margin = Inches(0.7)
        section.bottom_margin = Inches(0.6)
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
    if "excess_property" in coverages:
        generate_coverage_section(doc, data, "excess_property", "Excess Property Coverage — Layer 1")
    if "excess_property_2" in coverages:
        generate_coverage_section(doc, data, "excess_property_2", "Excess Property Coverage — Layer 2")
    if "general_liability" in coverages:
        generate_coverage_section(doc, data, "general_liability", "General Liability Coverage")
    if "workers_comp" in coverages:
        generate_coverage_section(doc, data, "workers_comp", "Workers Compensation Coverage")
    if "commercial_auto" in coverages:
        generate_coverage_section(doc, data, "commercial_auto", "Commercial Auto Coverage")
    if "umbrella" in coverages:
        generate_coverage_section(doc, data, "umbrella", "Umbrella / Excess Liability Coverage")
    if "umbrella_layer_2" in coverages:
        generate_coverage_section(doc, data, "umbrella_layer_2", "2nd Excess Liability Layer")
    if "umbrella_layer_3" in coverages:
        generate_coverage_section(doc, data, "umbrella_layer_3", "3rd Excess Liability Layer")
    if "cyber" in coverages:
        generate_coverage_section(doc, data, "cyber", "Cyber Liability Coverage")
    if "epli" in coverages:
        generate_coverage_section(doc, data, "epli", "Employment Practices Liability (EPLI) Coverage")
    if "flood" in coverages:
        generate_coverage_section(doc, data, "flood", "Flood Coverage")
    if "terrorism" in coverages:
        generate_coverage_section(doc, data, "terrorism", "Terrorism / TRIA Coverage")
    if "crime" in coverages:
        generate_coverage_section(doc, data, "crime", "Crime Coverage")
    if "inland_marine" in coverages:
        generate_coverage_section(doc, data, "inland_marine", "Inland Marine Coverage")
    
    # Part 3: Coverage Recommendations (before signature pages)
    generate_coverage_recommendations(doc)
    
    # Part 4: Signature Pages
    generate_confirmation_to_bind(doc, data)
    
    # Part 5: Compliance Pages (ALWAYS REQUIRED)
    generate_electronic_consent(doc)
    generate_carrier_rating(doc, data)
    generate_general_statement(doc)
    generate_property_definitions(doc)
    generate_how_we_get_paid(doc)
    generate_hub_advantage(doc)
    generate_tria_disclosure(doc)
    generate_california_licenses(doc)
    
    # Save
    doc.save(output_path)
    logger.info(f"Proposal saved to: {output_path}")
    return output_path
