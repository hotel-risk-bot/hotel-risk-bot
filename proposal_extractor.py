#!/usr/bin/env python3
"""
Hotel Insurance Proposal - Document Extraction & GPT Data Structuring
Extracts text from uploaded PDFs and Excel files, then uses GPT to structure
the data into a standardized format for proposal generation.
"""

import os
import json
import logging
import subprocess
import tempfile
from pathlib import Path
from typing import Optional

import openpyxl
from openai import OpenAI

logger = logging.getLogger(__name__)

# Initialize OpenAI client
client = OpenAI()
GPT_MODEL = "gpt-4.1-mini"


def extract_text_from_pdf(pdf_path: str) -> str:
    """Extract text from a PDF file using pdftotext."""
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
{
  "client_info": {
    "named_insured": "Full legal name",
    "dba": "DBA name if any",
    "address": "Full address",
    "entity_type": "LLC/Corp/etc",
    "effective_date": "MM/DD/YYYY",
    "expiration_date": "MM/DD/YYYY",
    "sales_exposure_basis": "Revenue/payroll amount if shown"
  },
  "coverages": {
    "property": {
      "carrier": "Carrier name",
      "carrier_admitted": true/false,
      "am_best_rating": "A+ XV or similar",
      "premium": 0,
      "taxes_fees": 0,
      "total_premium": 0,
      "tria_premium": 0,
      "limits": [
        {"description": "Building", "limit": "$X"},
        {"description": "Business Personal Property", "limit": "$X"},
        {"description": "Business Income", "limit": "ALS or $X"}
      ],
      "deductibles": [
        {"description": "All Other Perils", "amount": "$X"},
        {"description": "Named Storm", "amount": "X% or $X"},
        {"description": "Wind/Hail", "amount": "$X"}
      ],
      "additional_coverages": [
        {"description": "Flood", "limit": "$X or NOT COVERED"},
        {"description": "Earthquake", "limit": "$X or Excluded"},
        {"description": "Equipment Breakdown", "limit": "$X or Included"},
        {"description": "Ordinance or Law", "limit": "$X"},
        {"description": "Spoilage", "limit": "$X"},
        {"description": "Business Income Extended Period", "limit": "X days"},
        {"description": "Sign Coverage", "limit": "$X"}
      ],
      "forms_endorsements": [
        {"form_number": "CP 00 10 06/07", "description": "Building and Personal Property Coverage Form"}
      ],
      "subjectivities": ["List of binding requirements"]
    },
    "general_liability": {
      "carrier": "Carrier name",
      "carrier_admitted": true/false,
      "am_best_rating": "Rating",
      "premium": 0,
      "taxes_fees": 0,
      "total_premium": 0,
      "tria_premium": 0,
      "limits": [
        {"description": "Each Occurrence", "limit": "$X"},
        {"description": "General Aggregate", "limit": "$X"},
        {"description": "Products/Completed Operations", "limit": "$X"},
        {"description": "Personal & Advertising Injury", "limit": "$X"},
        {"description": "Damage to Rented Premises", "limit": "$X"},
        {"description": "Medical Payments", "limit": "$X"},
        {"description": "Hired & Non-Owned Auto CSL", "limit": "$X"}
      ],
      "aggregate_applies": "Per Location or Per Policy",
      "schedule_of_hazards": [
        {"location": "1", "classification": "Hotels", "code": "XXXXX", "basis": "S", "exposure": "$X"}
      ],
      "additional_coverages": [
        {"description": "Assault & Battery", "limit": "$X", "deductible": "$X"},
        {"description": "Abuse & Molestation", "limit": "$X", "deductible": "$X"},
        {"description": "Employee Benefits Liability", "limit": "$X", "deductible": "$X"}
      ],
      "forms_endorsements": [
        {"form_number": "CG 00 01 04/13", "description": "Commercial General Liability Coverage Form"}
      ],
      "subjectivities": []
    },
    "umbrella": {
      "carrier": "Carrier name",
      "carrier_admitted": true/false,
      "am_best_rating": "Rating",
      "premium": 0,
      "taxes_fees": 0,
      "total_premium": 0,
      "tria_premium": 0,
      "limits": [
        {"description": "Each Occurrence", "limit": "$X"},
        {"description": "Aggregate", "limit": "$X"},
        {"description": "Self-Insured Retention", "limit": "$X"}
      ],
      "underlying_insurance": [
        {"carrier": "Carrier", "coverage": "Auto Liability", "limits": "$X CSL"},
        {"carrier": "Carrier", "coverage": "General Liability", "limits": "$X Occ / $X Agg"}
      ],
      "tower_structure": [
        {"layer": "Primary", "carrier": "Carrier", "limits": "$XM xs $XM", "premium": 0, "total_cost": 0}
      ],
      "first_dollar_defense": true/false,
      "tria_included": true/false,
      "forms_endorsements": [],
      "subjectivities": []
    },
    "workers_comp": {
      "carrier": "Carrier name",
      "carrier_admitted": true/false,
      "am_best_rating": "Rating",
      "premium": 0,
      "taxes_fees": 0,
      "total_premium": 0,
      "limits": [
        {"description": "Workers Compensation", "limit": "Statutory"},
        {"description": "EL - Each Accident", "limit": "$X"},
        {"description": "EL - Disease Policy Limit", "limit": "$X"},
        {"description": "EL - Disease Each Employee", "limit": "$X"}
      ],
      "deductible": {"amount": "$X", "type": "Per Claim or Per Accident"},
      "rating_basis": [
        {"state": "XX", "location": "1", "class_code": "XXXX", "classification": "Hotels", "payroll": "$X", "rate": "X.XX"}
      ],
      "excluded_officers": ["Name - Title"],
      "forms_endorsements": [],
      "subjectivities": []
    },
    "commercial_auto": {
      "carrier": "Carrier name",
      "carrier_admitted": true/false,
      "am_best_rating": "Rating",
      "premium": 0,
      "taxes_fees": 0,
      "total_premium": 0,
      "limits": [
        {"description": "Liability CSL", "limit": "$X"},
        {"description": "Uninsured Motorist", "limit": "$X"},
        {"description": "Medical Payments", "limit": "$X"},
        {"description": "Comprehensive Deductible", "limit": "$X"},
        {"description": "Collision Deductible", "limit": "$X"}
      ],
      "vehicle_schedule": [
        {"year": "XXXX", "make": "Make", "model": "Model", "vin": "VIN", "garage_location": "City, ST"}
      ],
      "driver_schedule": [
        {"name": "Name", "dob": "MM/DD/YYYY", "license_state": "XX", "license_number": "XXXXX"}
      ],
      "forms_endorsements": [],
      "subjectivities": []
    }
  },
  "named_insureds": ["List of all named insureds"],
  "additional_interests": [
    {"type": "Mortgagee/Loss Payee/etc", "name_address": "Full name and address", "description": "Description"}
  ],
  "locations": [
    {"number": "1", "corporate_entity": "Entity name", "address": "Street", "city": "City", "state": "ST", "zip": "XXXXX", "description": "Hotel/Motel"}
  ],
  "expiring_premiums": {
    "property": 0,
    "general_liability": 0,
    "umbrella": 0,
    "workers_comp": 0,
    "commercial_auto": 0,
    "total": 0
  },
  "payment_options": [
    {"carrier": "Carrier", "terms": "Payment terms", "mep": "Minimum earned premium"}
  ]
}

IMPORTANT:
- Only include coverage sections that appear in the documents
- Extract EVERY form number and endorsement exactly as written
- Include form dates (e.g., "06/07" in "CP 00 10 06/07")
- Calculate total_premium = premium + taxes_fees for each coverage
- Mark excluded coverages explicitly
- For Property: ALWAYS include Flood and Earthquake rows even if excluded

DOCUMENT TEXT:
{document_text}"""


async def extract_and_structure_data(file_paths: list[str]) -> dict:
    """
    Extract text from all uploaded documents and use GPT to structure
    the data into a standardized format for proposal generation.
    
    Args:
        file_paths: List of paths to uploaded PDF/Excel files
        
    Returns:
        Structured dict of extracted insurance data
    """
    # Step 1: Extract text from all documents
    all_text = []
    for fp in file_paths:
        fname = Path(fp).name
        text = extract_document(fp)
        if text:
            all_text.append(f"\n{'='*60}\nFILE: {fname}\n{'='*60}\n{text}")
        else:
            logger.warning(f"No text extracted from: {fname}")
    
    if not all_text:
        return {"error": "Could not extract text from any uploaded documents."}
    
    combined_text = "\n".join(all_text)
    
    # Truncate if too long (GPT context limit)
    max_chars = 120000
    if len(combined_text) > max_chars:
        logger.warning(f"Document text truncated from {len(combined_text)} to {max_chars} chars")
        combined_text = combined_text[:max_chars]
    
    logger.info(f"Sending {len(combined_text)} chars to GPT for extraction")
    
    # Step 2: Send to GPT for structured extraction
    try:
        response = client.chat.completions.create(
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
        data = json.loads(result_text)
        logger.info(f"GPT extraction successful. Coverages found: {list(data.get('coverages', {}).keys())}")
        return data
        
    except json.JSONDecodeError as e:
        logger.error(f"GPT returned invalid JSON: {e}")
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
            "commercial_auto": "Commercial Auto"
        }
        
        for key, display_name in coverage_names.items():
            cov = coverages.get(key)
            if cov:
                carrier = cov.get("carrier", "N/A")
                total = cov.get("total_premium", 0)
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
                               ("commercial_auto", "COMMERCIAL AUTO")]:
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
        response = client.chat.completions.create(
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
