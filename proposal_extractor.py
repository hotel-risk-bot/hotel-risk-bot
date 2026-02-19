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
    "BINDING REQUIREMENTS",
    "CLASS CODE", "CLASSIFICATION",
    "OPTIONAL COVERAGES",
    "ADDITIONAL COVERAGES",
    "UNDERLYING INSURANCE",
    "TOWER STRUCTURE",
    "RATING BASIS",
    "PAYROLL",
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

    # Negative signals: boilerplate forms
    for keyword in BOILERPLATE_KEYWORDS:
        if keyword.upper() in text_upper:
            score -= 3.0

    # Negative: very long pages with mostly prose (forms text)
    if len(page_text) > 3000 and score < 2:
        # Check if it's mostly prose (few numbers, lots of text)
        num_count = len(re.findall(r'\d+', page_text))
        word_count = len(page_text.split())
        if word_count > 0 and num_count / word_count < 0.05:
            score -= 2.0

    return score


def extract_text_from_pdf_smart(pdf_path: str, max_chars: int = 80000) -> str:
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


def _extract_with_pdfplumber(pdf_path: str, max_chars: int = 80000) -> str:
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
      "limits": [
        {{"description": "Each Occurrence", "limit": "$X"}},
        {{"description": "General Aggregate", "limit": "$X"}},
        {{"description": "Products/Completed Operations", "limit": "$X"}},
        {{"description": "Personal & Advertising Injury", "limit": "$X"}},
        {{"description": "Damage to Rented Premises", "limit": "$X"}},
        {{"description": "Medical Payments", "limit": "$X"}}
      ],
      "aggregate_applies": "Per Location or Per Policy",
      "additional_coverages": [
        {{"description": "Coverage name", "limit": "$X", "deductible": "$X"}}
      ],
      "forms_endorsements": [
        {{"form_number": "CG 00 01 04/13", "description": "Commercial General Liability Coverage Form"}}
      ],
      "subjectivities": []
    }},
    "umbrella": {{
      "carrier": "Carrier name",
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
        {{"layer": "Primary", "carrier": "Carrier", "limits": "$XM xs $XM", "premium": 0, "total_cost": 0}}
      ],
      "first_dollar_defense": true,
      "tria_included": true,
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
    }}
  }},
  "named_insureds": ["List of all named insureds"],
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
    {{"carrier": "Carrier", "terms": "Payment terms", "mep": "Minimum earned premium"}}
  ]
}}

IMPORTANT:
- Only include coverage sections that appear in the documents
- Extract EVERY form number and endorsement exactly as written
- Include form dates (e.g., "06/07" in "CP 00 10 06/07")
- For total_premium: Use the "Total Cost of Policy" or "Total Policy Premium" figure if shown on the quote. This is the all-in number including base premium + broker fees + surplus lines tax + stamping fee + fire marshal tax + any other taxes/fees. If not shown, calculate total_premium = premium + taxes_fees
- ALWAYS preserve cents in premium amounts (e.g., $60,513.35 not $60,513)
- Mark excluded coverages explicitly
- For Property: ALWAYS include Flood and Earthquake rows even if excluded

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
    max_chars = 80000
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
            max_tokens=16000
        )

        result_text = response.choices[0].message.content
        finish_reason = response.choices[0].finish_reason
        logger.info(f"GPT response: {len(result_text)} chars, finish_reason={finish_reason}")

        if finish_reason == "length":
            logger.warning("GPT response was truncated (hit max_tokens). Attempting to parse partial JSON.")

        data = json.loads(result_text)
        logger.info(f"GPT extraction successful. Coverages found: {list(data.get('coverages', {}).keys())}")

        # Log coverage details
        for key, cov in data.get("coverages", {}).items():
            logger.info(f"  {key}: carrier={cov.get('carrier', 'N/A')}, premium={cov.get('premium', 0)}, total={cov.get('total_premium', 0)}")

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
        if limits:
            lines.append("  Limits:")
            for lim in limits:
                lines.append(f"    • {lim.get('description', '')}: {lim.get('limit', '')}")

        # Deductibles
        deductibles = cov.get("deductibles", [])
        if deductibles:
            lines.append("  Deductibles:")
            for ded in deductibles:
                lines.append(f"    • {ded.get('description', '')}: {ded.get('amount', '')}")

        # Additional Coverages
        addl = cov.get("additional_coverages", [])
        if addl:
            lines.append("  Additional Coverages:")
            for ac in addl:
                ded_str = f" (Ded: {ac['deductible']})" if ac.get("deductible") else ""
                lines.append(f"    • {ac.get('description', '')}: {ac.get('limit', '')}{ded_str}")

        # Forms count
        forms = cov.get("forms_endorsements", [])
        if forms:
            lines.append(f"  Forms & Endorsements: {len(forms)} extracted")
            for f in forms[:5]:
                lines.append(f"    • {f.get('form_number', '')} — {f.get('description', '')}")
            if len(forms) > 5:
                lines.append(f"    ... and {len(forms) - 5} more")

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
            max_tokens=16000
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
        max_chars = 80000
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
                max_tokens=16000,
            )

            result_text = response.choices[0].message.content
            finish_reason = response.choices[0].finish_reason
            logger.info(f"GPT response: {len(result_text)} chars, finish_reason={finish_reason}")

            data = json.loads(result_text)
            data["client_name"] = client_name
            logger.info(
                f"GPT extraction successful. Coverages found: "
                f"{list(data.get('coverages', {}).keys())}"
            )

            # Log coverage details
            for key, cov in data.get("coverages", {}).items():
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
