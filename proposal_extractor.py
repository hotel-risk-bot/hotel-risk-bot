#!/usr/bin/env python3
"""
Hotel Insurance Proposal - Document Extraction & GPT Data Structuring
Extracts text from uploaded PDFs and Excel files, then uses GPT to structure
the data into a standardized format for proposal generation.

Key feature: Smart page-level extraction that identifies quote summary pages
and prioritizes them over forms/endorsements boilerplate.
"""

import os
import json
import logging
import re
import subprocess
import tempfile
from pathlib import Path
from typing import Optional

import openpyxl
from openai import OpenAI

try:
    import pdfplumber
    HAS_PDFPLUMBER = True
except ImportError:
    HAS_PDFPLUMBER = False

logger = logging.getLogger(__name__)

# OpenAI client (lazy initialization)
_client = None
GPT_MODEL = "gpt-4.1-mini"


def _get_openai_client():
    """Lazy-initialize the OpenAI client."""
    global _client
    if _client is None:
        api_key = os.environ.get("OPENAI_API_KEY")
        if not api_key:
            raise RuntimeError(
                "OPENAI_API_KEY environment variable not set. "
                "Please add it to your Railway environment variables."
            )
        _client = OpenAI(api_key=api_key)
    return _client


# Keywords that indicate a page contains important quote data (high priority)
QUOTE_PAGE_KEYWORDS = [
    "QUOTATION", "QUOTE", "PROPOSAL", "INDICATION",
    "COVERAGE –", "COVERAGE -", "COVERAGE:", "COVERAGE SUMMARY",
    "PREMIUM BREAKDOWN", "PREMIUM SUMMARY", "TOTAL PREMIUM",
    "Total Cost of Policy", "Total General Liability Premium",
    "Total Property Premium", "Total Umbrella Premium",
    "LIMITS OF INSURANCE", "LIMITS OF LIABILITY",
    "GENERAL LIABILITY LIMITS", "PROPERTY LIMITS",
    "DEDUCTIBLE", "DEDUCTIBLES",
    "SCHEDULE OF LOCATIONS", "SCHEDULE OF VALUES",
    "INSURED INFORMATION", "NAMED INSURED",
    "EFFECTIVE DATE", "EXPIRATION DATE",
    "AGENCY INFORMATION", "BROKER",
    "PREMIUM:", "RATE:", "EXPOSURE:",
    "DECLARATIONS", "DECLARATION PAGE",
    "COMMERCIAL GENERAL LIABILITY DECLARATIONS",
    "COMMERCIAL PROPERTY DECLARATIONS",
    "WORKERS COMPENSATION DECLARATIONS",
    "COMMERCIAL AUTO DECLARATIONS",
    "UMBRELLA DECLARATIONS",
    "PREMIUM BREAKDOWN",
    "SCHEDULE OF FORMS",
    "SUBJECTIVITIES", "SUBJECTIVES",
    "CONDITIONS & SUBJECTIVES", "CONDITIONS AND SUBJECTIVES",
    "BINDING CONDITIONS", "BINDING REQUIREMENTS",
    "BINDING SUBJECTIVITIES",
    "CLASS CODE", "CLASSIFICATION",
    "OPTIONAL COVERAGES",
    "ADDITIONAL COVERAGES",
    "SUBLIMITS",
    "SUBLIMITS OF LIABILITY",
    "EXTENSIONS OF COVERAGE",
    "COVERAGE EXTENSIONS",
    "FORMS SCHEDULE",
    "ENDORSEMENT SCHEDULE",
    "FORMS AND ENDORSEMENTS",
    "UNDERLYING INSURANCE",
    "TOWER STRUCTURE",
    "RATING BASIS",
    "PAYROLL",
    "INSURING CLAUSE",
    "INSURING AGREEMENT",
    "EMPLOYEE THEFT",
    "FORGERY OR ALTERATION",
    "SOCIAL ENGINEERING",
    "COMPUTER AND FUNDS TRANSFER",
    "FIDELITY",
    "CRIME COVERAGE",
    "FOREFRONT",
]

# Keywords that indicate boilerplate forms/endorsements (low priority)
BOILERPLATE_KEYWORDS = [
    "THIS ENDORSEMENT CHANGES THE POLICY",
    "PLEASE READ IT CAREFULLY",
    "THIS ENDORSEMENT MODIFIES INSURANCE",
    "COMMON POLICY CONDITIONS",
    "COMMERCIAL GENERAL LIABILITY COVERAGE FORM",
    "COMMERCIAL PROPERTY CONDITIONS",
    "COMMERCIAL PROPERTY COVERAGE FORM",
    "CAUSES OF LOSS",
    "TERRORISM RISK INSURANCE ACT",
    "NUCLEAR HAZARD EXCLUSION",
    "EXCLUSION OF TERRORISM",
    "POLICYHOLDER DISCLOSURE",
    "NOTICE OF TERRORISM",
    "Section 102(1) of the Act",
    "means activities against persons",
    "intimidate or coerce a government",
    "© Insurance Services Office",
    "© ISO Properties",
    "Includes copyrighted material",
]


def _score_page(page_text: str) -> float:
    """
    Score a PDF page based on how likely it contains important quote data.
    Higher score = more important.
    """
    text_upper = page_text.upper()
    score = 0.0

    # Positive signals: quote/summary content
    for keyword in QUOTE_PAGE_KEYWORDS:
        if keyword.upper() in text_upper:
            score += 2.0

    # Strong positive: contains dollar amounts with commas (premium figures)
    dollar_amounts = re.findall(r'\$\s*[\d,]+(?:\.\d{2})?', page_text)
    if dollar_amounts:
        score += min(len(dollar_amounts) * 0.5, 5.0)

    # Strong positive: contains percentage rates
    rates = re.findall(r'\d+\.\d{2,4}\s*%?', page_text)
    if rates:
        score += min(len(rates) * 0.3, 3.0)

    # Detect if this is a forms SCHEDULE/LIST page (lists form numbers + descriptions)
    # These pages are HIGH VALUE — they list the forms attached to the policy
    # Do NOT penalize them with boilerplate keywords
    is_forms_schedule = any(kw in text_upper for kw in [
        "FORMS SCHEDULE", "ENDORSEMENT SCHEDULE", "FORMS AND ENDORSEMENTS",
        "SCHEDULE OF FORMS", "FORMS & EXCLUSIONS APPLICABLE",
        "FORMS APPLICABLE", "ENDORSEMENTS APPLICABLE",
        "POLICY FORMS AND ENDORSEMENTS",
    ])
    # Also detect if page has many form numbers (e.g., CP 00 10, CG 00 01, etc.)
    form_numbers = re.findall(r'[A-Z]{2,4}\s*\d{2,4}\s+\d{2,4}', page_text)
    if len(form_numbers) >= 5:
        is_forms_schedule = True
        score += 5.0  # Boost pages with many form numbers

    # Negative signals: boilerplate forms (but NOT forms schedule pages)
    if not is_forms_schedule:
        for keyword in BOILERPLATE_KEYWORDS:
            if keyword.upper() in text_upper:
                score -= 3.0

    # Negative: very long pages with mostly prose (forms text)
    # But NOT forms schedule pages which are tabular
    if not is_forms_schedule and len(page_text) > 3000 and score < 2:
        # Check if it's mostly prose (few numbers, lots of text)
        num_count = len(re.findall(r'\d+', page_text))
        word_count = len(page_text.split())
        if word_count > 0 and num_count / word_count < 0.05:
            score -= 2.0

    return score


def extract_text_from_pdf_smart(pdf_path: str, max_chars: int = 120000) -> str:
    """
    Smart PDF text extraction that prioritizes quote summary pages
    over forms/endorsements boilerplate.

    Extracts text page-by-page, scores each page, and returns the
    highest-scoring pages up to max_chars.
    """
    try:
        # Get total page count
        result = subprocess.run(
            ["pdfinfo", pdf_path],
            capture_output=True, text=True, timeout=30
        )
        pages_match = re.search(r"Pages:\s+(\d+)", result.stdout)
        total_pages = int(pages_match.group(1)) if pages_match else 0

        if total_pages == 0:
            logger.warning(f"pdfinfo returned 0 pages, falling back to pdfplumber")
            return _extract_with_pdfplumber(pdf_path, max_chars)
        
        # Log file size for debugging
        file_size = os.path.getsize(pdf_path)
        logger.info(f"PDF file size: {file_size} bytes")

        logger.info(f"PDF has {total_pages} pages, extracting page-by-page for scoring")

        # Extract text page by page
        page_texts = []
        for page_num in range(1, total_pages + 1):
            result = subprocess.run(
                ["pdftotext", "-layout", "-f", str(page_num), "-l", str(page_num), pdf_path, "-"],
                capture_output=True, text=True, timeout=30
            )
            if result.returncode == 0:
                page_text = result.stdout.strip()
                if page_text:
                    score = _score_page(page_text)
                    page_texts.append({
                        "page": page_num,
                        "text": page_text,
                        "score": score,
                        "chars": len(page_text)
                    })

        if not page_texts:
            logger.warning("No text extracted from any pages via pdftotext, trying pdfplumber fallback")
            return _extract_with_pdfplumber(pdf_path, max_chars)

        # Sort by score (highest first), then by page number for ties
        page_texts.sort(key=lambda x: (-x["score"], x["page"]))

        # Log the top and bottom scored pages
        logger.info(f"Page scoring results ({len(page_texts)} pages with text):")
        for p in page_texts[:10]:
            logger.info(f"  Page {p['page']}: score={p['score']:.1f}, chars={p['chars']}")
        if len(page_texts) > 10:
            logger.info(f"  ... {len(page_texts) - 10} more pages")
            for p in page_texts[-3:]:
                logger.info(f"  Page {p['page']}: score={p['score']:.1f}, chars={p['chars']} (lowest)")

        # Select pages up to max_chars, prioritizing high-score pages
        selected_pages = []
        total_chars = 0
        for p in page_texts:
            if total_chars + p["chars"] > max_chars:
                # If we haven't selected any pages yet, take at least the first one
                if not selected_pages:
                    selected_pages.append(p)
                break
            selected_pages.append(p)
            total_chars += p["chars"]

        # Re-sort selected pages by page number for coherent reading order
        selected_pages.sort(key=lambda x: x["page"])

        logger.info(
            f"Selected {len(selected_pages)} of {len(page_texts)} pages "
            f"({total_chars} chars) for GPT extraction"
        )

        # Combine selected pages with page markers
        parts = []
        for p in selected_pages:
            parts.append(f"\n--- Page {p['page']} ---\n{p['text']}")

        return "\n".join(parts)

    except Exception as e:
        logger.error(f"Smart PDF extraction failed, falling back to pdfplumber: {e}")
        return _extract_with_pdfplumber(pdf_path, max_chars)


def _extract_with_pdfplumber(pdf_path: str, max_chars: int = 120000) -> str:
    """Fallback PDF extraction using pdfplumber (pure Python, no system deps)."""
    if not HAS_PDFPLUMBER:
        logger.warning("pdfplumber not available, falling back to basic pdftotext")
        return extract_text_from_pdf(pdf_path)
    
    try:
        logger.info(f"Using pdfplumber for extraction: {pdf_path}")
        file_size = os.path.getsize(pdf_path)
        logger.info(f"PDF file size: {file_size} bytes")
        
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            logger.info(f"pdfplumber: PDF has {total_pages} pages")
            
            # Extract and score pages
            page_texts = []
            for i, page in enumerate(pdf.pages):
                try:
                    text = page.extract_text() or ""
                    if text.strip():
                        score = _score_page(text)
                        page_texts.append({
                            "page": i + 1,
                            "text": text,
                            "score": score,
                            "chars": len(text)
                        })
                except Exception as e:
                    logger.warning(f"pdfplumber: error on page {i+1}: {e}")
                    continue
            
            if not page_texts:
                logger.warning("pdfplumber: No text extracted from any pages")
                return ""
            
            logger.info(f"pdfplumber: extracted text from {len(page_texts)} of {total_pages} pages")
            
            # Sort by score (highest first)
            page_texts.sort(key=lambda x: (-x["score"], x["page"]))
            
            # Log top scored pages
            for p in page_texts[:10]:
                logger.info(f"  Page {p['page']}: score={p['score']:.1f}, chars={p['chars']}")
            
            # Select pages up to max_chars
            selected_pages = []
            total_chars = 0
            for p in page_texts:
                if total_chars + p["chars"] > max_chars:
                    if not selected_pages:
                        selected_pages.append(p)
                    break
                selected_pages.append(p)
                total_chars += p["chars"]
            
            # Re-sort by page number
            selected_pages.sort(key=lambda x: x["page"])
            
            logger.info(f"pdfplumber: selected {len(selected_pages)} pages ({total_chars} chars)")
            
            parts = []
            for p in selected_pages:
                parts.append(f"\n--- Page {p['page']} ---\n{p['text']}")
            
            return "\n".join(parts)
    
    except Exception as e:
        logger.error(f"pdfplumber extraction failed: {e}")
        return extract_text_from_pdf(pdf_path)


def extract_text_from_pdf(pdf_path: str) -> str:
    """Extract text from a PDF file using pdftotext (basic full extraction)."""
    try:
        result = subprocess.run(
            ["pdftotext", "-layout", pdf_path, "-"],
            capture_output=True, text=True, timeout=60
        )
        if result.returncode == 0 and result.stdout.strip():
            logger.info(f"Extracted {len(result.stdout)} chars from PDF: {pdf_path}")
            return result.stdout
        else:
            # Fallback: try without layout flag
            result = subprocess.run(
                ["pdftotext", pdf_path, "-"],
                capture_output=True, text=True, timeout=60
            )
            if result.returncode == 0:
                logger.info(f"Extracted {len(result.stdout)} chars from PDF (no layout): {pdf_path}")
                return result.stdout
            logger.error(f"pdftotext failed: {result.stderr}")
            return ""
    except Exception as e:
        logger.error(f"PDF extraction error: {e}")
        return ""


def extract_text_from_excel(excel_path: str) -> str:
    """Extract data from an Excel file (SOV or schedule)."""
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
        all_text = []
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            all_text.append(f"\n=== Sheet: {sheet_name} ===\n")
            for row in ws.iter_rows(values_only=True):
                cells = [str(c) if c is not None else "" for c in row]
                if any(c.strip() for c in cells):
                    all_text.append(" | ".join(cells))
        text = "\n".join(all_text)
        logger.info(f"Extracted {len(text)} chars from Excel: {excel_path}")
        return text
    except Exception as e:
        logger.error(f"Excel extraction error: {e}")
        return ""


def extract_document(file_path: str) -> str:
    """Extract text from a document based on its file extension."""
    ext = Path(file_path).suffix.lower()
    if ext == ".pdf":
        return extract_text_from_pdf(file_path)
    elif ext in (".xlsx", ".xls"):
        return extract_text_from_excel(file_path)
    else:
        logger.warning(f"Unsupported file type: {ext}")
        return ""


def extract_document_smart(file_path: str) -> str:
    """Extract text from a document using smart extraction for PDFs."""
    ext = Path(file_path).suffix.lower()
    if ext == ".pdf":
        return extract_text_from_pdf_smart(file_path)
    elif ext in (".xlsx", ".xls"):
        return extract_text_from_excel(file_path)
    else:
        logger.warning(f"Unsupported file type: {ext}")
        return ""


EXTRACTION_SYSTEM_PROMPT = """You are an expert insurance data extraction assistant specializing in hotel and hospitality insurance. You extract structured data from insurance quote documents.

CRITICAL RULES:
1. Extract EVERY form and endorsement number with full description and date
2. NEVER summarize - list everything exactly as shown
3. NEVER write "Additional forms as listed in policy"
4. Extract ALL additional coverages (even if excluded - mark as "Excluded" or "NOT COVERED")
5. For Property: ALWAYS include Flood and Earthquake status (even if excluded)
6. Include ALL taxes, fees, surcharges in premium calculations
7. Exclude TRIA premiums from totals
8. Extract ALL deductibles
9. Extract ALL limits
10. Note carrier name and whether admitted or non-admitted
11. Property additional_coverages (sublimits) is MANDATORY - if you see ANY sublimits, extensions, or coverage limits in the property quote, extract ALL of them. Common ones: Flood, Earthquake, Equipment Breakdown, Ordinance or Law, Spoilage, Business Income Extended Period, Sign Coverage, Accounts Receivable, Valuable Papers, Newly Acquired Property, Transit, Debris Removal, Pollutant Cleanup, Sewer/Drain Backup, Water Damage, Mold/Fungi, Green Building
12. Property forms_endorsements is MANDATORY - extract EVERY form number from the forms schedule page
13. GL forms_endorsements is MANDATORY - extract EVERY form number from the GL forms schedule
14. GL designated_premises is MANDATORY when CG2144/NXLL110 form exists - extract ALL addresses
15. GL schedule_of_classes MUST include ALL class codes with actual dollar exposure amounts
16. Named insureds MUST be exact legal entity names from the quote - do NOT concatenate hotel brand names into entity names
17. For carrier names: Use the ISSUING carrier name (e.g., "Associated Industries Insurance Company" or "Technology Insurance Company" for AmTrust policies, "Palms Insurance Company" for Palms). Do NOT use the wholesale broker name as the carrier.

Return your extraction as a JSON object with the following structure. Only include sections that are present in the documents."""

EXTRACTION_USER_PROMPT = """Extract ALL insurance data from the following quote document(s). Return a JSON object.

The JSON structure should be:
{{
  "client_info": {{
    "named_insured": "Full legal name",
    "dba": "DBA name if any",
    "address": "Full address",
    "entity_type": "LLC/Corp/etc",
    "effective_date": "MM/DD/YYYY",
    "expiration_date": "MM/DD/YYYY",
    "sales_exposure_basis": "Revenue/payroll amount if shown"
  }},
  "coverages": {{
    "property": {{
      "carrier": "Carrier name",
      "carrier_admitted": true or false,
      "am_best_rating": "A+ XV or similar",
      "premium": 0,
      "taxes_fees": 0,
      "total_premium": 0,
      "tria_premium": 0,
      "tiv": "$X Total Insured Value from quote or SOV (e.g., $56,390,000)",
      "limits": [
        {{"description": "Building", "limit": "$X"}},
        {{"description": "Business Personal Property", "limit": "$X"}},
        {{"description": "Business Income", "limit": "ALS or $X"}}
      ],
      "deductibles": [
        {{"description": "All Other Perils", "amount": "$X"}},
        {{"description": "Named Storm", "amount": "X% or $X"}},
        {{"description": "Wind/Hail", "amount": "$X"}}
      ],
      "additional_coverages": [
        {{"description": "Flood", "limit": "$X or NOT COVERED"}},
        {{"description": "Earthquake", "limit": "$X or Excluded"}},
        {{"description": "Equipment Breakdown", "limit": "$X or Included"}},
        {{"description": "Ordinance or Law", "limit": "$X"}},
        {{"description": "Spoilage", "limit": "$X"}},
        {{"description": "Business Income Extended Period", "limit": "X days"}},
        {{"description": "Sign Coverage", "limit": "$X"}}
      ],
      "forms_endorsements": [
        {{"form_number": "CP 00 10 06/07", "description": "Building and Personal Property Coverage Form"}}
      ],
      "subjectivities": ["List of binding requirements"]
    }},
    "general_liability": {{
      "carrier": "Carrier name",
      "carrier_admitted": true or false,
      "am_best_rating": "Rating",
      "premium": 0,
      "taxes_fees": 0,
      "total_premium": 0,
      "tria_premium": 0,
      "gl_deductible": "$X Per Occurrence (or $0 if none)",
      "defense_basis": "In Addition to Limits or Within Limits",
      "limits": [
        {{"description": "Each Occurrence", "limit": "$X"}},
        {{"description": "General Aggregate", "limit": "$X"}},
        {{"description": "Products/Completed Operations", "limit": "$X"}},
        {{"description": "Personal & Advertising Injury", "limit": "$X"}},
        {{"description": "Damage to Rented Premises", "limit": "$X"}},
        {{"description": "Medical Payments", "limit": "$X"}}
      ],
      "aggregate_applies": "Per Location or Per Policy",
      "schedule_of_classes": [
        {{"location": "Loc 1", "address": "Street Address", "brand_dba": "Hotel Brand or DBA Name", "classification": "Hotels/Motels", "class_code": "XXXXX", "rate": "Rate per $100 or flat amount", "exposure_basis": "Sales/Revenue/Area/Units", "exposure": "$X or number", "premium": "$X"}}
      ],
      "additional_coverages": [
        {{"description": "Coverage name", "limit": "$X", "deductible": "$X"}}
      ],
      "designated_premises": [
        "Full address 1 from CG2144/NXLL110 designated premises form",
        "Full address 2"
      ],
      "forms_endorsements": [
        {{"form_number": "CG 00 01 04/13", "description": "Commercial General Liability Coverage Form"}}
      ],
      "subjectivities": []
    }},
    "umbrella": {{
      "carrier": "Primary layer carrier name",
      "carrier_admitted": true or false,
      "am_best_rating": "Rating",
      "premium": 0,
      "taxes_fees": 0,
      "total_premium": 0,
      "tria_premium": 0,
      "limits": [
        {{"description": "Each Occurrence", "limit": "$X"}},
        {{"description": "Aggregate", "limit": "$X"}},
        {{"description": "Self-Insured Retention", "limit": "$X"}}
      ],
      "underlying_insurance": [
        {{"carrier": "Carrier", "coverage": "Auto Liability", "limits": "$X CSL"}},
        {{"carrier": "Carrier", "coverage": "General Liability", "limits": "$X Occ / $X Agg"}}
      ],
      "tower_structure": [
        {{"layer": "Primary", "carrier": "Carrier", "limits": "$5M xs Primary", "premium": 0, "total_cost": 0}}
      ],
      "first_dollar_defense": true,
      "tria_included": true,
      "forms_endorsements": [],
      "subjectivities": []
    }},
    "umbrella_layer_2": {{
      "carrier": "Second layer carrier name (if a second excess layer exists)",
      "carrier_admitted": true or false,
      "am_best_rating": "Rating",
      "premium": 0,
      "taxes_fees": 0,
      "total_premium": 0,
      "limits": [
        {{"description": "Each Occurrence", "limit": "$X"}},
        {{"description": "Aggregate", "limit": "$X"}}
      ],
      "tower_structure": [
        {{"layer": "2nd Excess", "carrier": "Carrier", "limits": "$5M xs $5M", "premium": 0, "total_cost": 0}}
      ],
      "forms_endorsements": [],
      "subjectivities": []
    }},
    "umbrella_layer_3": {{
      "carrier": "Third layer carrier name (if a third excess layer exists)",
      "carrier_admitted": true or false,
      "am_best_rating": "Rating",
      "premium": 0,
      "taxes_fees": 0,
      "total_premium": 0,
      "limits": [
        {{"description": "Each Occurrence", "limit": "$X"}},
        {{"description": "Aggregate", "limit": "$X"}}
      ],
      "tower_structure": [
        {{"layer": "3rd Excess", "carrier": "Carrier", "limits": "$5M xs $10M", "premium": 0, "total_cost": 0}}
      ],
      "forms_endorsements": [],
      "subjectivities": []
    }},
    "workers_comp": {{
      "carrier": "Carrier name",
      "carrier_admitted": true or false,
      "am_best_rating": "Rating",
      "premium": 0,
      "taxes_fees": 0,
      "total_premium": 0,
      "limits": [
        {{"description": "Workers Compensation", "limit": "Statutory"}},
        {{"description": "EL - Each Accident", "limit": "$X"}},
        {{"description": "EL - Disease Policy Limit", "limit": "$X"}},
        {{"description": "EL - Disease Each Employee", "limit": "$X"}}
      ],
      "deductible": {{"amount": "$X", "type": "Per Claim or Per Accident"}},
      "rating_basis": [
        {{"state": "XX", "location": "1", "class_code": "XXXX", "classification": "Hotels", "payroll": "$X", "rate": "X.XX"}}
      ],
      "forms_endorsements": [],
      "subjectivities": []
    }},
    "commercial_auto": {{
      "carrier": "Carrier name",
      "carrier_admitted": true or false,
      "am_best_rating": "Rating",
      "premium": 0,
      "taxes_fees": 0,
      "total_premium": 0,
      "limits": [
        {{"description": "Liability CSL", "limit": "$X"}},
        {{"description": "Uninsured Motorist", "limit": "$X"}},
        {{"description": "Medical Payments", "limit": "$X"}}
      ],
      "vehicle_schedule": [
        {{"year": "XXXX", "make": "Make", "model": "Model", "vin": "VIN"}}
      ],
      "forms_endorsements": [],
      "subjectivities": []
    }},
    "terrorism": {{
      "carrier": "Carrier name",
      "carrier_admitted": true or false,
      "am_best_rating": "Rating",
      "premium": 0,
      "taxes_fees": 0,
      "total_premium": 0,
      "limits": [
        {{"description": "Certified Acts of Terrorism", "limit": "$X or Policy Limit"}},
        {{"description": "Non-Certified Acts / Active Assailant", "limit": "$X or N/A"}}
      ],
      "additional_coverages": [
        {{"description": "Coverage name", "limit": "$X"}}
      ],
      "forms_endorsements": [],
      "subjectivities": []
    }},
    "crime": {{
      "carrier": "Carrier name (e.g., Federal Insurance Company for Chubb)",
      "carrier_admitted": true or false,
      "am_best_rating": "Rating",
      "premium": 0,
      "taxes_fees": 0,
      "total_premium": 0,
      "insuring_clauses": [
        {{"clause": "Employee Theft", "limit": "$X", "retention": "$X"}},
        {{"clause": "Forgery or Alteration", "limit": "$X", "retention": "$X"}},
        {{"clause": "Inside the Premises - Theft of Money & Securities", "limit": "$X", "retention": "$X"}},
        {{"clause": "Inside the Premises - Robbery/Safe Burglary of Other Property", "limit": "$X", "retention": "$X"}},
        {{"clause": "Outside the Premises", "limit": "$X", "retention": "$X"}},
        {{"clause": "Computer and Funds Transfer Fraud", "limit": "$X", "retention": "$X"}},
        {{"clause": "Money Orders and Counterfeit Money", "limit": "$X", "retention": "$X"}},
        {{"clause": "Social Engineering Fraud", "limit": "$X", "retention": "$X"}}
      ],
      "limits": [
        {{"description": "Per Loss Limit", "limit": "$X"}},
        {{"description": "Aggregate Limit", "limit": "$X"}}
      ],
      "forms_endorsements": [],
      "subjectivities": []
    }},
    "cyber": {{
      "carrier": "Carrier name",
      "carrier_admitted": true or false,
      "am_best_rating": "Rating",
      "premium": 0,
      "taxes_fees": 0,
      "total_premium": 0,
      "limits": [
        {{"description": "Aggregate Limit", "limit": "$X"}},
        {{"description": "Retention/Deductible", "limit": "$X"}}
      ],
      "additional_coverages": [
        {{"description": "Coverage name", "limit": "$X"}}
      ],
      "forms_endorsements": [],
      "subjectivities": []
    }}
  }},
  "named_insureds": [
    {{"name": "Full legal entity name", "dba": "DBA/trade name if shown"}}
  ],
  "additional_named_insureds": [
    {{"name": "Entity name", "dba": "DBA if shown"}}
  ],
  "additional_insureds": [
    {{"name": "Entity or person name", "relationship": "Franchisor/Mortgagee/Manager/etc", "description": "Additional details"}}
  ],
  "additional_interests": [
    {{"type": "Mortgagee/Loss Payee/etc", "name_address": "Full name and address", "description": "Description"}}
  ],
  "locations": [
    {{"number": "1", "corporate_entity": "Entity name", "address": "Street", "city": "City", "state": "ST", "zip": "XXXXX", "description": "Hotel/Motel"}}
  ],
  "expiring_premiums": {{
    "property": 0,
    "general_liability": 0,
    "umbrella": 0,
    "workers_comp": 0,
    "commercial_auto": 0,
    "total": 0
  }},
  "payment_options": [
    {{"carrier": "Carrier", "coverage_type": "Property, General Liability, Umbrella / Excess, Workers Compensation, Crime, Terrorism, Equipment Breakdown, EPLI, Cyber, Flood, Auto", "terms": "Payment terms (exclude commission/broker fee info)", "mep": "Minimum earned premium"}}
  ]
}}

IMPORTANT:
- COVERAGE CLASSIFICATION: A standalone terrorism/TRIA policy is NOT general liability. If a document is from Lloyd's of London, AEGIS, or similar and covers ONLY terrorism/TRIA/certified acts of terrorism/active assailant, classify it as "terrorism" NOT "general_liability". General Liability covers bodily injury, property damage, personal & advertising injury with occurrence/aggregate limits. Terrorism covers certified/non-certified acts of terrorism. If a single policy bundles both, put the terrorism portion in "terrorism" and the GL portion in "general_liability".
- Only include coverage sections that appear in the documents
- Extract EVERY form number and endorsement exactly as written
- Include form dates (e.g., "06/07" in "CP 00 10 06/07")
- For total_premium: This MUST be the all-in out-the-door number. Look for "Total Package Cost", "Total Cost of Policy", "Total Policy Cost", "Total Policy Premium", "Total Due", "Grand Total", "Total Amount Due", "Total Estimated Cost", or any final total line. It includes base premium + broker fees + surplus lines tax + stamping fee + fire marshal tax + inspection fees + FSLSO fees + EMPA surcharge + any other taxes/fees/surcharges. If no single total line exists, calculate total_premium = premium + taxes_fees. CRITICAL: total_premium must ALWAYS be >= premium. If the quote shows separate line items for taxes and fees, ADD them ALL to the base premium to get total_premium. For example if GL premium is $163,832 and there are surplus lines taxes of $5,414, stamping fee of $328, broker fee of $3,000, and inspection fee of $252, then taxes_fees = $8,994 and total_premium = $163,832 + $8,994 = $172,826. NEVER use the base premium as total_premium when taxes/fees exist
- For GL policies that include BOTH General Liability AND Liquor Liability in a single package: The "premium" field should be the combined package premium (GL + Liquor), and "total_premium" should be the Total Package Cost (premium + broker fee + surplus lines tax + stamping fee). For example if GL premium is $408,733, Liquor is $10,287, Total Package Premium is $419,020, and Total Package Cost is $442,471, then premium=$419,020 and total_premium=$442,471
- For GL gl_deductible: Extract the per-occurrence deductible if one exists. Look for "Deductible Per Occurrence", "Deductible Liability", "$X,000 Deductible Per Occurrence Including Loss Adjustment Expense", or similar. Include the full description (e.g., "$5,000 Per Occurrence Including Loss Adjustment Expense"). If no GL deductible, set to "$0" or "None".
- For GL defense_basis: Look for "Defense Basis" or whether defense costs are "In Addition to Limits" or "Within Limits of Liability".
- For GL schedule_of_classes: Extract the exposure schedule. This may be location-based OR class-code-based. For class-code-based quotes (like AmTrust), extract each class code entry with: class_code (e.g., "45190"), classification/description, rate (e.g., "9.964" per $100), exposure amount (e.g., "$8,748,612"), and exposure_basis (e.g., "Gross Sales", "Per Acre", "Area", "Liquor Sales", "FLAT"). For location-based quotes, include address, brand_dba, classification, exposure, and premium. Include vacant land, restaurants, liquor, sundry, hired auto, loss control, and all non-hotel entries. Include ALL exposure classes for each location (e.g., Hotel/Motel, Restaurant, Liquor Liability as separate rows). CRITICAL: Always capture the actual dollar amount for exposure (e.g., $8,748,612 not just "Gross Sales"). The exposure_basis describes what the number represents (Gross Sales, Revenue, Area, etc.)
- GL DESIGNATED PREMISES LOCATIONS: If the GL quote includes a form like CG 21 44, CG2144, NXLL110, or any "Limitation of Coverage to Designated Premises" form, extract EVERY location listed in that form. These are ALL the locations covered under the GL policy. Do TWO things: (1) Add each address as a separate schedule_of_classes entry with its full address. (2) Also populate the "designated_premises" array with each full address string exactly as written (e.g., "4285 Highway 51, LaPlace, LA 70068"). The designated_premises array is the AUTHORITATIVE list of GL covered locations. CRITICAL: The CG2144/NXLL110 form typically lists addresses in a numbered format like "1) 4285 Highway 51, LaPlace, LA 70068" followed by "2) 4281 Highway 51..." etc. Extract ALL numbered addresses, not just the first few. There may be 8 or more addresses. Also look for addresses that may appear with labels like "Office:" or "Hotels:" before the numbered list. Extract those too. If the form text is split across multiple pages, combine all addresses from all pages.
- ALWAYS preserve cents in premium amounts (e.g., $60,513.35 not $60,513)
- Mark excluded coverages explicitly
- For Property tiv: Extract the Total Insured Value (TIV) from the property quote or SOV. Look for "Total Insured Value", "TIV", "Total Values", or the sum total on the Schedule of Values. This should be the total of Building + Contents/BPP + Business Income/Rents across all locations. For example if the SOV shows Building Total $42,800,000 + Contents Total $7,700,000 + BI/Rents Total $5,550,000 = TIV $56,390,000. Use the actual SOV/quote total, NOT the per-location coverage limits.
- For Property: ALWAYS include Flood and Earthquake rows even if excluded
- For Property additional_coverages (sublimits/extensions): This section is MANDATORY. Extract ALL sublimits of liability, also called extensions of coverage or additional coverages. Common property sublimits include: Flood, Earthquake, Equipment Breakdown, Ordinance or Law, Spoilage, Business Income Extended Period, Sign Coverage, Accounts Receivable, Valuable Papers, Fine Arts, Newly Acquired Property, Transit, Debris Removal, Pollutant Cleanup, Utility Services, Green Building, Sewer/Drain Backup, Water Damage, Mold/Fungi, and any other sublimit or extension listed in the quote. Include the limit and deductible for each.
- For Property forms_endorsements: This section is MANDATORY. Extract EVERY policy form and endorsement listed in the property quote. Include the exact form number (e.g., CP 00 10 06/07) and description. These are typically listed on a forms schedule or endorsement schedule page. Do NOT skip this section even if the list is long.
- For ALL coverage types subjectivities: This section is CRITICAL. Extract ALL conditions, subjectives, binding requirements, and binding conditions listed in the quote. These are often on a page titled "CONDITIONS & SUBJECTIVES", "BINDING REQUIREMENTS", "BINDING SUBJECTIVITIES", or "BINDING CONDITIONS". Each bullet point or numbered item should be a separate string in the subjectivities array. Include items like: loss control report requirements, certificates of insurance requirements, named insured confirmation, application requirements, ACORD application deadlines, terrorism form requirements, payment of state taxes, inspection/audit contact requirements, and any other conditions the carrier requires before or after binding. Do NOT skip or summarize — extract each condition verbatim as written in the quote.
- For named_insureds: Extract each named insured as an object with "name" and "dba" fields. Do NOT repeat the same entity twice (case-insensitive). If a named insured has a DBA or trade name EXPLICITLY listed in the quote (e.g., "Q Hotels Management LLC DBA Best Western"), split into name="Q Hotels Management LLC" and dba="Best Western". CRITICAL RULES: (1) Only include DBAs that are EXPLICITLY written as "DBA", "d/b/a", or "doing business as" in the documents. (2) Do NOT infer DBAs from hotel brand names, location names, or SOV entries. (3) Do NOT fabricate entity names like "Cajun Lodging LLC" unless that exact name appears in the quote documents. (4) If a named insured appears as "Q HOTEL MANAGEMENT, LLC" in ALL CAPS, extract it exactly as written — the generator will handle proper case formatting. (5) Do NOT create separate named insured entries for each hotel brand — those are locations, not named insureds.
- For additional_named_insureds: Search ALL pages for "Additional Named Insured", "Additional Named Insureds Schedule", "Named Insured Schedule", or similar headings. These are often on a separate page listing multiple entities (e.g., LLCs, management companies, DBAs). Extract every entity listed. Do NOT duplicate entities already in named_insureds.
- For additional_insureds: Search for "Additional Insured", "Additional Insured Schedule", or endorsement pages listing additional insureds (franchisors, mortgagees, managers). Extract all of them.
- CRIME COVERAGE: For crime/fidelity bond policies (e.g., Chubb ForeFront Portfolio, Travelers Crime), extract ALL insuring clauses with their individual limits and retentions. Common insuring clauses include: Employee Theft, Forgery or Alteration, Inside the Premises (Theft of Money & Securities), Inside the Premises (Robbery/Safe Burglary), Outside the Premises, Computer and Funds Transfer Fraud, Money Orders and Counterfeit Money, Social Engineering Fraud. Also extract all endorsements from the forms schedule. If the policy is claims-made, note the retroactive date.
- UMBRELLA/EXCESS LAYERS: When multiple umbrella/excess liability quotes are provided (e.g., separate PDFs for different layers), extract EACH layer as a separate coverage entry. Use "umbrella" for the primary excess layer, "umbrella_layer_2" for the second excess layer ($XM xs $XM), and "umbrella_layer_3" for the third excess layer ($XM xs $XM). Each layer has its own carrier, premium, limits, forms, and subjectivities. The tower_structure field should show that layer's position. Look for "Controlling Underlying" or "Schedule of Underlying" to determine the layer position. If a quote says it sits excess of another carrier's layer, it is a higher layer.

DOCUMENT TEXT:
{document_text}"""


async def extract_and_structure_data(file_paths: list[str]) -> dict:
    """
    Extract text from all uploaded documents and use GPT to structure
    the data into a standardized format for proposal generation.

    Uses smart page-level extraction for PDFs to prioritize quote
    summary pages over forms/endorsements boilerplate.
    """
    # Step 1: Extract text from all documents using smart extraction
    all_text = []
    for fp in file_paths:
        fname = Path(fp).name
        text = extract_document_smart(fp)
        if text:
            all_text.append(f"\n{'='*60}\nFILE: {fname}\n{'='*60}\n{text}")
            logger.info(f"Smart extraction from {fname}: {len(text)} chars")
        else:
            logger.warning(f"No text extracted from: {fname}")

    if not all_text:
        return {"error": "Could not extract text from any uploaded documents."}

    combined_text = "\n".join(all_text)

    # Final safety truncation (should rarely be needed with smart extraction)
    max_chars = 120000
    if len(combined_text) > max_chars:
        logger.warning(f"Combined text truncated from {len(combined_text)} to {max_chars} chars")
        combined_text = combined_text[:max_chars]

    logger.info(f"Sending {len(combined_text)} chars to GPT for extraction")

    # Step 2: Send to GPT for structured extraction
    try:
        response = _get_openai_client().chat.completions.create(
            model=GPT_MODEL,
            messages=[
                {"role": "system", "content": EXTRACTION_SYSTEM_PROMPT},
                {"role": "user", "content": EXTRACTION_USER_PROMPT.format(document_text=combined_text)}
            ],
            response_format={"type": "json_object"},
            temperature=0.1,
            max_tokens=32000
        )

        result_text = response.choices[0].message.content
        finish_reason = response.choices[0].finish_reason
        logger.info(f"GPT response: {len(result_text)} chars, finish_reason={finish_reason}")

        if finish_reason == "length":
            logger.warning("GPT response was truncated (hit max_tokens). Attempting to parse partial JSON.")

        data = json.loads(result_text)
        
        # Normalize coverages: GPT sometimes returns a list instead of dict
        covs = data.get("coverages", {})
        if isinstance(covs, list):
            normalized = {}
            for item in covs:
                if isinstance(item, dict):
                    cov_type = item.get("coverage_type", item.get("type", "unknown"))
                    normalized[cov_type] = item
            data["coverages"] = normalized
            logger.info(f"Normalized coverages from list ({len(covs)} items) to dict ({list(normalized.keys())})")
        elif not isinstance(covs, dict):
            data["coverages"] = {}
        
        # Also fix individual coverage values that are lists instead of dicts
        covs = data.get("coverages", {})
        if isinstance(covs, dict):
            for key, val in list(covs.items()):
                if isinstance(val, list):
                    if len(val) >= 1 and isinstance(val[0], dict):
                        covs[key] = val[0]
                        logger.info(f"Unwrapped list for coverage '{key}'")
                    else:
                        covs[key] = {}
                elif not isinstance(val, dict):
                    covs[key] = {}
        
        logger.info(f"GPT extraction successful. Coverages found: {list(data.get('coverages', {}).keys())}")

        # POST-PROCESSING: Validate and fix common extraction issues
        
        # Fix 1: Clean up named insureds - remove entries with multiple hotel brand names
        _brand_names = {"marriott", "hilton", "ihg", "wyndham", "best western", "choice",
                       "hampton inn", "hampton", "holiday inn", "holiday inn express",
                       "candlewood", "towneplace", "staybridge", "springhill",
                       "comfort inn", "comfort suites", "quality inn", "sleep inn"}
        raw_named = data.get("named_insureds", [])
        cleaned_named = []
        for ni in raw_named:
            ni_name = ni.get("name", "") if isinstance(ni, dict) else str(ni)
            ni_lower = ni_name.lower()
            brand_count = sum(1 for b in _brand_names if b in ni_lower)
            if brand_count >= 3:
                # This is likely a hallucinated concatenation — try to extract just the entity
                import re as _re_fix
                m = _re_fix.match(r'^(.+?\b(?:LLC|LP|LLP|Inc|Corp)\b)', ni_name, _re_fix.IGNORECASE)
                if m:
                    if isinstance(ni, dict):
                        ni["name"] = m.group(1).strip()
                        ni["dba"] = ""  # Clear the hallucinated DBA
                    else:
                        ni = {"name": m.group(1).strip(), "dba": ""}
                    logger.warning(f"Fixed hallucinated named insured: '{ni_name}' -> '{ni.get('name', ni) if isinstance(ni, dict) else ni}'")
            cleaned_named.append(ni)
        data["named_insureds"] = cleaned_named
        
        # Fix 2: Ensure GL carrier name is correct (AmTrust entities often misidentified)
        gl_cov = data.get("coverages", {}).get("general_liability", {})
        if gl_cov:
            carrier = gl_cov.get("carrier", "")
            carrier_lower = carrier.lower()
            # If carrier contains "associated industries" but forms show AmTrust, fix it
            if "associated industries" in carrier_lower:
                gl_cov["carrier"] = "AmTrust E&S (Associated Industries)"
                logger.info(f"Fixed GL carrier: '{carrier}' -> 'AmTrust E&S (Associated Industries)'")
            elif "technology insurance" in carrier_lower:
                gl_cov["carrier"] = "AmTrust E&S (Technology Insurance)"
                logger.info(f"Fixed GL carrier: '{carrier}' -> 'AmTrust E&S (Technology Insurance)'")
        
        # Fix 3: Validate that forms_endorsements is not empty for property and GL
        for cov_key in ["property", "general_liability"]:
            cov = data.get("coverages", {}).get(cov_key, {})
            if cov and not cov.get("forms_endorsements"):
                logger.warning(f"{cov_key} has no forms_endorsements extracted — may need manual review")
        
        # Fix 4: Validate additional_coverages for property
        prop_cov = data.get("coverages", {}).get("property", {})
        if prop_cov and not prop_cov.get("additional_coverages"):
            logger.warning("Property has no additional_coverages (sublimits) extracted — may need manual review")

        # Validate and fix total_premium for each coverage
        for key, cov in data.get("coverages", {}).items():
            premium = cov.get("premium", 0) or 0
            taxes_fees = cov.get("taxes_fees", 0) or 0
            total_premium = cov.get("total_premium", 0) or 0
            
            # Ensure numeric types
            if isinstance(premium, str):
                try: premium = float(str(premium).replace(",", "").replace("$", ""))
                except: premium = 0
            if isinstance(taxes_fees, str):
                try: taxes_fees = float(str(taxes_fees).replace(",", "").replace("$", ""))
                except: taxes_fees = 0
            if isinstance(total_premium, str):
                try: total_premium = float(str(total_premium).replace(",", "").replace("$", ""))
                except: total_premium = 0
            
            # Fallback: if total_premium is less than premium, recalculate
            if total_premium < premium and taxes_fees > 0:
                corrected = premium + taxes_fees
                logger.warning(f"  {key}: total_premium ({total_premium}) < premium ({premium}). "
                             f"Correcting to premium + taxes_fees = {corrected}")
                cov["total_premium"] = corrected
                total_premium = corrected
            elif total_premium == 0 and premium > 0:
                cov["total_premium"] = premium + taxes_fees
                total_premium = cov["total_premium"]
                logger.info(f"  {key}: total_premium was 0, set to premium + taxes_fees = {total_premium}")
            
            logger.info(f"  {key}: carrier={cov.get('carrier', 'N/A')}, premium={premium}, "
                       f"taxes_fees={taxes_fees}, total_premium={total_premium}")

        return data

    except json.JSONDecodeError as e:
        logger.error(f"GPT returned invalid JSON: {e}")
        logger.error(f"Raw response (first 500 chars): {result_text[:500] if 'result_text' in dir() else 'N/A'}")
        return {"error": f"Failed to parse extraction results: {e}"}
    except Exception as e:
        logger.error(f"GPT extraction failed: {e}")
        return {"error": f"AI extraction failed: {e}"}


def format_verification_message(data: dict) -> str:
    """
    Format the extracted data into a verification message for the user
    to review before generating the proposal.
    """
    if "error" in data:
        return f"Extraction Error: {data['error']}"

    lines = []
    lines.append("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    lines.append("📋 EXTRACTED DATA — PLEASE VERIFY")
    lines.append("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

    # Client Info
    ci = data.get("client_info", {})
    if ci:
        lines.append("")
        lines.append("▸ CLIENT INFORMATION")
        lines.append(f"  Named Insured: {ci.get('named_insured', 'N/A')}")
        if ci.get("dba"):
            lines.append(f"  DBA: {ci['dba']}")
        if ci.get("address"):
            lines.append(f"  Address: {ci['address']}")
        if ci.get("effective_date"):
            lines.append(f"  Effective: {ci['effective_date']}")
        if ci.get("expiration_date"):
            lines.append(f"  Expiration: {ci['expiration_date']}")

    # Premium Summary
    coverages = data.get("coverages", {})
    if coverages:
        lines.append("")
        lines.append("▸ PREMIUM SUMMARY")
        lines.append(f"  {'Coverage':<25} {'Carrier':<20} {'Total Premium':>15}")
        lines.append(f"  {'─'*25} {'─'*20} {'─'*15}")

        grand_total = 0
        coverage_names = {
            "property": "Property",
            "general_liability": "General Liability",
            "umbrella": "Umbrella",
            "workers_comp": "Workers Comp",
            "commercial_auto": "Commercial Auto",
            "cyber": "Cyber",
            "epli": "EPLI",
            "flood": "Flood",
            "terrorism": "Terrorism / TRIA",
            "crime": "Crime",
            "inland_marine": "Inland Marine"
        }

        for key, display_name in coverage_names.items():
            cov = coverages.get(key)
            if cov:
                carrier = cov.get("carrier", "N/A")
                total = cov.get("total_premium", 0)
                if not isinstance(total, (int, float)):
                    try:
                        total = float(str(total).replace(",", "").replace("$", ""))
                    except (ValueError, TypeError):
                        total = 0
                admitted = "" if cov.get("carrier_admitted", True) else " (Non-Admitted)"
                lines.append(f"  {display_name:<25} {carrier[:20]:<20} ${total:>12,.0f}")
                if admitted:
                    lines.append(f"  {'':25} {admitted}")
                grand_total += total

        lines.append(f"  {'─'*25} {'─'*20} {'─'*15}")
        lines.append(f"  {'TOTAL':<25} {'':20} ${grand_total:>12,.0f}")

    # Coverage Details
    for key, display_name in [("property", "PROPERTY"), ("general_liability", "GENERAL LIABILITY"),
                               ("umbrella", "UMBRELLA"), ("workers_comp", "WORKERS COMP"),
                               ("commercial_auto", "COMMERCIAL AUTO"), ("cyber", "CYBER"),
                               ("epli", "EPLI"), ("flood", "FLOOD"),
                               ("terrorism", "TERRORISM / TRIA"), ("crime", "CRIME"),
                               ("inland_marine", "INLAND MARINE")]:
        cov = coverages.get(key)
        if not cov:
            continue

        lines.append("")
        lines.append(f"▸ {display_name}")
        lines.append(f"  Carrier: {cov.get('carrier', 'N/A')}")
        lines.append(f"  AM Best: {cov.get('am_best_rating', 'N/A')}")

        # Limits
        limits = cov.get("limits", [])
        if limits and isinstance(limits, list):
            lines.append("  Limits:")
            for lim in limits:
                if isinstance(lim, dict):
                    lines.append(f"    • {lim.get('description', '')}: {lim.get('limit', '')}")
                elif isinstance(lim, str):
                    lines.append(f"    • {lim}")

        # Deductibles
        deductibles = cov.get("deductibles", [])
        if deductibles and isinstance(deductibles, list):
            lines.append("  Deductibles:")
            for ded in deductibles:
                if isinstance(ded, dict):
                    lines.append(f"    • {ded.get('description', '')}: {ded.get('amount', '')}")
                elif isinstance(ded, str):
                    lines.append(f"    • {ded}")

        # Additional Coverages
        addl = cov.get("additional_coverages", [])
        if addl and isinstance(addl, list):
            lines.append("  Additional Coverages:")
            for ac in addl:
                if isinstance(ac, dict):
                    ded_str = f" (Ded: {ac['deductible']})" if ac.get("deductible") else ""
                    lines.append(f"    • {ac.get('description', '')}: {ac.get('limit', '')}{ded_str}")
                elif isinstance(ac, str):
                    lines.append(f"    • {ac}")

        # Forms count
        forms = cov.get("forms_endorsements", [])
        if forms and isinstance(forms, list):
            lines.append(f"  Forms & Endorsements: {len(forms)} extracted")
            for f in forms[:5]:
                if isinstance(f, dict):
                    lines.append(f"    • {f.get('form_number', '')} — {f.get('description', '')}")
                elif isinstance(f, str):
                    lines.append(f"    • {f}")
            if len(forms) > 5:
                lines.append(f"    ... and {len(forms) - 5} more")

        # Subjectivities / Conditions
        subjs = cov.get("subjectivities", [])
        if subjs and isinstance(subjs, list):
            lines.append(f"  Conditions & Subjectivities: {len(subjs)} items")
            for s in subjs[:5]:
                s_text = s if isinstance(s, str) else str(s)
                # Truncate long items for display
                if len(s_text) > 100:
                    s_text = s_text[:97] + "..."
                lines.append(f"    ☐ {s_text}")
            if len(subjs) > 5:
                lines.append(f"    ... and {len(subjs) - 5} more")

    # Locations
    locations = data.get("locations", [])
    if locations:
        lines.append("")
        lines.append(f"▸ LOCATIONS: {len(locations)} found")
        for loc in locations[:5]:
            lines.append(f"  {loc.get('number', '?')}. {loc.get('address', '')} {loc.get('city', '')}, {loc.get('state', '')} {loc.get('zip', '')}")
        if len(locations) > 5:
            lines.append(f"  ... and {len(locations) - 5} more")

    # Named Insureds
    named = data.get("named_insureds", [])
    if named:
        lines.append("")
        lines.append(f"▸ NAMED INSUREDS: {len(named)}")
        for ni in named[:5]:
            lines.append(f"  • {ni}")
        if len(named) > 5:
            lines.append(f"  ... and {len(named) - 5} more")

    lines.append("")
    lines.append("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")
    lines.append("⚠️ PLEASE VERIFY ALL DATA ABOVE")
    lines.append("")
    lines.append("Reply with:")
    lines.append("  ✅ /proposal confirm — to generate the proposal")
    lines.append("  ✏️ Send corrections as a message")
    lines.append("  ❌ /proposal cancel — to cancel")
    lines.append("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━")

    return "\n".join(lines)


async def apply_corrections(data: dict, corrections_text: str) -> dict:
    """Use GPT to apply user corrections to the extracted data."""
    try:
        response = _get_openai_client().chat.completions.create(
            model=GPT_MODEL,
            messages=[
                {"role": "system", "content": "You are an insurance data correction assistant. Apply the user's corrections to the extracted data JSON. Return the corrected JSON object. Only modify the fields mentioned in the corrections, keep everything else the same."},
                {"role": "user", "content": f"Current extracted data:\n{json.dumps(data, indent=2)}\n\nUser corrections:\n{corrections_text}\n\nReturn the corrected JSON:"}
            ],
            response_format={"type": "json_object"},
            temperature=0.1,
            max_tokens=32000
        )
        corrected = json.loads(response.choices[0].message.content)
        logger.info("Corrections applied successfully")
        return corrected
    except Exception as e:
        logger.error(f"Failed to apply corrections: {e}")
        return data


class ProposalExtractor:
    """
    Class wrapper around the extraction functions for use by proposal_handler.
    Provides an OOP interface to the module-level extraction functions.
    """

    def extract_pdf_text(self, pdf_path: str) -> str:
        """Extract text from a PDF file using smart extraction."""
        return extract_text_from_pdf_smart(pdf_path)

    def extract_excel_data(self, excel_path: str) -> str:
        """Extract text from an Excel file."""
        return extract_text_from_excel(excel_path)

    def structure_insurance_data(
        self,
        pdf_texts: list[dict],
        excel_data: list[dict],
        client_name: str,
    ) -> dict:
        """
        Use GPT to structure extracted text into a standardized insurance data format.

        Args:
            pdf_texts: List of dicts with 'filename' and 'text' keys
            excel_data: List of dicts with 'filename' and 'data' keys
            client_name: Name of the client/insured
        """
        # Combine all text sources
        all_text_parts = []
        for item in pdf_texts:
            all_text_parts.append(
                f"\n{'='*60}\nFILE: {item['filename']}\n{'='*60}\n{item['text']}"
            )
        for item in excel_data:
            all_text_parts.append(
                f"\n{'='*60}\nFILE (Excel): {item['filename']}\n{'='*60}\n{item['data']}"
            )

        if not all_text_parts:
            return {"error": "No text extracted from any documents."}

        combined_text = "\n".join(all_text_parts)

        # Safety truncation
        max_chars = 120000
        if len(combined_text) > max_chars:
            logger.warning(
                f"Document text truncated from {len(combined_text)} to {max_chars} chars"
            )
            combined_text = combined_text[:max_chars]

        logger.info(f"Sending {len(combined_text)} chars to GPT for extraction")

        try:
            response = _get_openai_client().chat.completions.create(
                model=GPT_MODEL,
                messages=[
                    {"role": "system", "content": EXTRACTION_SYSTEM_PROMPT},
                    {
                        "role": "user",
                        "content": EXTRACTION_USER_PROMPT.format(
                            document_text=combined_text
                        ),
                    },
                ],
                response_format={"type": "json_object"},
                temperature=0.1,
                max_tokens=32000,
            )

            result_text = response.choices[0].message.content
            finish_reason = response.choices[0].finish_reason
            logger.info(f"GPT response: {len(result_text)} chars, finish_reason={finish_reason}")

            data = json.loads(result_text)
            data["client_name"] = client_name
            
            # Normalize coverages: GPT sometimes returns a list instead of dict
            covs = data.get("coverages", {})
            if isinstance(covs, list):
                # Convert list of coverage dicts to keyed dict
                normalized = {}
                for item in covs:
                    if isinstance(item, dict):
                        cov_type = item.get("coverage_type", item.get("type", "unknown"))
                        normalized[cov_type] = item
                data["coverages"] = normalized
                covs = normalized
            elif not isinstance(covs, dict):
                data["coverages"] = {}
                covs = data["coverages"]
            
            # Also fix individual coverage values that are lists instead of dicts
            if isinstance(covs, dict):
                for key, val in list(covs.items()):
                    if isinstance(val, list):
                        if len(val) >= 1 and isinstance(val[0], dict):
                            covs[key] = val[0]
                            logger.info(f"Unwrapped list for coverage '{key}'")
                        else:
                            covs[key] = {}
                    elif not isinstance(val, dict):
                        covs[key] = {}
            
            logger.info(
                f"GPT extraction successful. Coverages found: "
                f"{list(covs.keys()) if isinstance(covs, dict) else covs}"
            )

            # Log coverage details
            if isinstance(covs, dict):
                for key, cov in covs.items():
                    if isinstance(cov, dict):
                        logger.info(f"  {key}: carrier={cov.get('carrier', 'N/A')}, premium={cov.get('premium', 0)}, total={cov.get('total_premium', 0)}")

            return data

        except json.JSONDecodeError as e:
            logger.error(f"GPT returned invalid JSON: {e}")
            return {"error": f"Failed to parse extraction results: {e}"}
        except Exception as e:
            logger.error(f"GPT extraction failed: {e}")
            return {"error": f"AI extraction failed: {e}"}

    def apply_adjustments(self, data: dict, instructions: str) -> dict:
        """
        Apply user corrections/adjustments to the extracted data using GPT.
        This is a synchronous wrapper; the caller should use asyncio.to_thread.
        """
        import asyncio

        loop = asyncio.new_event_loop()
        try:
            return loop.run_until_complete(apply_corrections(data, instructions))
        finally:
            loop.close()
