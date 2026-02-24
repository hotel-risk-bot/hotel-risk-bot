#!/usr/bin/env python3
"""
Hotel Insurance Proposal - Telegram Conversation Handler
Manages the /proposal command flow: start session, receive uploads, extract data,
verify with user, and generate final DOCX document.
"""

import os
import json
import logging
import tempfile
import asyncio
from pathlib import Path
from datetime import datetime

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ContextTypes, CommandHandler, MessageHandler, CallbackQueryHandler,
    ConversationHandler, filters
)

from proposal_extractor import ProposalExtractor
from proposal_generator import generate_proposal
from sov_parser import parse_sov, is_sov_file, format_sov_summary, aggregate_locations

logger = logging.getLogger(__name__)

# Conversation states
(
    WAITING_FOR_FILES,
    WAITING_FOR_MORE_FILES,
    REVIEWING_EXTRACTION,
    CONFIRMING_GENERATION,
    WAITING_FOR_EXPIRING,
) = range(5)

# Session storage (in-memory, per chat)
proposal_sessions = {}



def _escape_md(text: str) -> str:
    """Escape special Telegram Markdown v1 characters in text (for filenames etc.)."""
    for ch in ('_', '*', '`', '['):
        text = text.replace(ch, f'\\{ch}')
    return text


class ProposalSession:
    """Tracks state for an active proposal generation session."""
    
    def __init__(self, client_name: str, chat_id: int):
        self.client_name = client_name
        self.chat_id = chat_id
        self.uploaded_files = []  # List of (filename, local_path, file_type)
        self.extracted_data = None
        self.processed_files = set()  # Track which files have been extracted
        self.created_at = datetime.now()
        self.work_dir = tempfile.mkdtemp(prefix="proposal_")
    
    def add_file(self, filename: str, local_path: str, file_type: str):
        self.uploaded_files.append({
            "filename": filename,
            "local_path": local_path,
            "file_type": file_type
        })
    
    def get_file_summary(self, escape_md: bool = False) -> str:
        if not self.uploaded_files:
            return "No files uploaded yet."
        lines = []
        for i, f in enumerate(self.uploaded_files, 1):
            fn = _escape_md(f['filename']) if escape_md else f['filename']
            lines.append(f"  {i}. {fn} ({f['file_type']})")
        return "\n".join(lines)
    
    def cleanup(self):
        """Remove temporary files."""
        import shutil
        try:
            shutil.rmtree(self.work_dir, ignore_errors=True)
        except Exception:
            pass


def get_session(chat_id: int) -> ProposalSession:
    """Get the active proposal session for a chat."""
    return proposal_sessions.get(chat_id)


def clear_session(chat_id: int):
    """Clear and cleanup the proposal session."""
    session = proposal_sessions.pop(chat_id, None)
    if session:
        session.cleanup()


async def safe_reply(update: Update, text: str, **kwargs):
    """Send a message, splitting if too long for Telegram's 4096 char limit.
    Falls back to plain text if Markdown parsing fails."""
    MAX_LEN = 4000

    async def _send_chunk(chunk_text, **kw):
        try:
            await update.message.reply_text(chunk_text, **kw)
        except Exception as e:
            if "parse entities" in str(e).lower() or "can't parse" in str(e).lower():
                logger.warning(f"Markdown send failed in safe_reply: {e}, retrying plain")
                plain_kw = {k: v for k, v in kw.items() if k != 'parse_mode'}
                plain = chunk_text.replace('*', '').replace('_', '').replace('`', '')
                await update.message.reply_text(plain, **plain_kw)
            else:
                raise

    if len(text) <= MAX_LEN:
        await _send_chunk(text, **kwargs)
        return
    
    # Split at line breaks
    lines = text.split("\n")
    chunk = ""
    for line in lines:
        if len(chunk) + len(line) + 1 > MAX_LEN:
            if chunk:
                await _send_chunk(chunk, **kwargs)
            chunk = line
        else:
            chunk = chunk + "\n" + line if chunk else line
    if chunk:
        await _send_chunk(chunk, **kwargs)


# ─── Command Handlers ─────────────────────────────────────────

async def proposal_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Start a new proposal session: /proposal [Client Name]"""
    logger.info(f"proposal_start called by user {update.effective_user.id} in chat {update.effective_chat.id}")
    logger.info(f"Raw message text: {update.message.text}")
    args = context.args
    if not args:
        await safe_reply(update,
            "Please provide a client name.\n\n"
            "Usage: `/proposal Client Name`\n"
            "Example: `/proposal RAR Elite Management Inc`",
            parse_mode="Markdown"
        )
        return ConversationHandler.END
    
    client_name = " ".join(args)
    chat_id = update.effective_chat.id
    
    # Clear any existing session
    clear_session(chat_id)
    
    # Create new session
    session = ProposalSession(client_name, chat_id)
    proposal_sessions[chat_id] = session
    
    await safe_reply(update,
        f"📋 **New Proposal Session Started**\n\n"
        f"**Client:** {_escape_md(client_name)}\n\n"
        f"Please upload your insurance quote documents:\n"
        f"• Property quote (PDF)\n"
        f"• General Liability quote (PDF)\n"
        f"• Umbrella/Excess quote (PDF)\n"
        f"• Workers Compensation quote (PDF)\n"
        f"• Commercial Auto quote (PDF)\n"
        f"• Schedule of Values / SOV (Excel)\n\n"
        f"Upload files one at a time. When done, send /extract to process.\n"
        f"Use /expiring to set expiring premiums for comparison.\n"
        f"Send /proposal\\_cancel to cancel.",
        parse_mode="Markdown"
    )
    
    return WAITING_FOR_FILES


async def receive_file(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle uploaded files (PDF or Excel)."""
    logger.info(f"receive_file called by user {update.effective_user.id}")
    chat_id = update.effective_chat.id
    session = get_session(chat_id)
    
    if not session:
        await update.message.reply_text("No active proposal session. Start one with /proposal [Client Name]")
        return ConversationHandler.END
    
    document = update.message.document
    if not document:
        await update.message.reply_text("Please upload a PDF or Excel file.")
        return WAITING_FOR_FILES
    
    filename = document.file_name or "unknown"
    ext = Path(filename).suffix.lower()
    
    if ext not in [".pdf", ".xlsx", ".xls", ".csv"]:
        await update.message.reply_text(
            f"Unsupported file type: {ext}\n"
            f"Please upload PDF or Excel files only."
        )
        return WAITING_FOR_FILES
    
    # Determine file type
    if ext == ".pdf":
        file_type = "pdf"
    elif ext in [".xlsx", ".xls"]:
        file_type = "excel"
    else:
        file_type = "csv"
    
    # Download the file
    try:
        file = await document.get_file()
        local_path = os.path.join(session.work_dir, filename)
        await file.download_to_drive(local_path)
        
        # Log file size for debugging
        actual_size = os.path.getsize(local_path)
        logger.info(f"Downloaded file '{filename}' to '{local_path}', size: {actual_size} bytes")
        
        session.add_file(filename, local_path, file_type)
        
        file_count = len(session.uploaded_files)
        unprocessed = len([f for f in session.uploaded_files if f['filename'] not in session.processed_files])
        extract_hint = f"\n\n\U0001f4cc **{unprocessed} new file(s)** ready for extraction. Send /extract to process." if unprocessed > 0 and session.extracted_data else ""
        safe_fn = _escape_md(filename)
        await safe_reply(update,
            f"\u2705 Received: **{safe_fn}** ({file_type.upper()})\n\n"
            f"**Files uploaded ({file_count}):**\n{session.get_file_summary(escape_md=True)}"
            f"{extract_hint}\n\n"
            f"Upload more files or send /extract when ready.",
            parse_mode="Markdown"
        )
        
    except Exception as e:
        logger.error(f"Error downloading file: {e}")
        await update.message.reply_text(f"Error downloading file: {e}")
    
    return WAITING_FOR_FILES


def _normalize_coverages(data):
    """Ensure coverages is always a dict, not a list."""
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
    return data


def _merge_extraction_results(existing: dict, new_data: dict) -> dict:
    """Merge extraction results from multiple PDFs into a single data structure.
    
    Each PDF is extracted individually. This function merges the coverages,
    locations, named insureds, etc. from each extraction into one unified result.
    New coverages are added; existing coverages are NOT overwritten.
    Client info is merged (fill in blanks from new data).
    """
    if not existing:
        return _normalize_coverages(new_data)
    if not new_data or "error" in new_data:
        return existing
    
    # Normalize coverages in both inputs
    _normalize_coverages(existing)
    _normalize_coverages(new_data)
    
    merged = json.loads(json.dumps(existing))  # Deep copy
    
    # Merge client_info - fill in blanks
    existing_ci = merged.get("client_info", {})
    new_ci = new_data.get("client_info", {})
    for key, val in new_ci.items():
        if val and val != "N/A" and (not existing_ci.get(key) or existing_ci.get(key) == "N/A"):
            existing_ci[key] = val
    merged["client_info"] = existing_ci
    
    # Merge coverages - add new coverage types, don't overwrite existing
    existing_covs = merged.get("coverages", {})
    new_covs = new_data.get("coverages", {})
    for cov_key, cov_data in new_covs.items():
        if cov_key not in existing_covs:
            existing_covs[cov_key] = cov_data
            logger.info(f"Merged new coverage: {cov_key} from {cov_data.get('carrier', 'unknown')}")
        else:
            logger.info(f"Coverage {cov_key} already exists, keeping existing")
    merged["coverages"] = existing_covs
    
    # Merge locations - add new ones (avoid duplicates by address)
    existing_locs = merged.get("locations", [])
    existing_addrs = {loc.get("address", "").upper() for loc in existing_locs}
    for loc in new_data.get("locations", []):
        addr = loc.get("address", "").upper()
        if addr and addr not in existing_addrs:
            existing_locs.append(loc)
            existing_addrs.add(addr)
    merged["locations"] = existing_locs
    
    # Merge named insureds - add unique ones
    existing_named = set(merged.get("named_insureds", []))
    for ni in new_data.get("named_insureds", []):
        existing_named.add(ni)
    merged["named_insureds"] = list(existing_named)
    
    # Merge additional interests
    existing_ai = merged.get("additional_interests", [])
    existing_ai_names = {ai.get("name_address", "") for ai in existing_ai}
    for ai in new_data.get("additional_interests", []):
        if ai.get("name_address", "") not in existing_ai_names:
            existing_ai.append(ai)
    merged["additional_interests"] = existing_ai
    
    # Merge expiring premiums - fill in zeros
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


async def extract_data(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Process uploaded files and extract insurance data.
    
    Each PDF is processed individually with its own GPT call, then results
    are merged. This prevents large PDFs from overwhelming smaller ones.
    Only processes files that haven't been extracted yet.
    """
    logger.info(f"extract_data called by user {update.effective_user.id}")
    chat_id = update.effective_chat.id
    session = get_session(chat_id)
    
    if not session:
        await update.message.reply_text("No active proposal session. Start one with /proposal [Client Name]")
        return ConversationHandler.END
    
    if not session.uploaded_files:
        await update.message.reply_text("No files uploaded yet. Please upload at least one quote document.")
        return WAITING_FOR_FILES
    
    # Find files that haven't been processed yet
    new_files = [f for f in session.uploaded_files if f["filename"] not in session.processed_files]
    
    if not new_files and session.extracted_data:
        await update.message.reply_text(
            "All files have already been extracted.\n"
            "Upload more files or send /generate to create the proposal."
        )
        return REVIEWING_EXTRACTION
    
    # If no new files but also no extracted data, process all files
    if not new_files:
        new_files = session.uploaded_files
    
    await safe_reply(update,
        f"⏳ **Processing {len(new_files)} file(s) individually...**\n\n"
        f"Each document is extracted separately for accuracy.\n"
        f"This may take 1-2 minutes per file.",
        parse_mode="Markdown"
    )
    
    try:
        extractor = ProposalExtractor()
        
        for file_info in new_files:
            filename = file_info["filename"]
            logger.info(f"Processing file individually: {filename}")
            
            if file_info["file_type"] == "pdf":
                text = extractor.extract_pdf_text(file_info["local_path"])
                logger.info(f"PDF '{filename}': extracted {len(text)} chars")
                
                if not text:
                    logger.warning(f"No text extracted from PDF: {filename}")
                    await update.message.reply_text(f"⚠️ Could not extract text from: {filename}")
                    session.processed_files.add(filename)
                    continue
                
                # Extract this single PDF with GPT
                pdf_texts = [{"filename": filename, "text": text}]
                file_data = await asyncio.to_thread(
                    extractor.structure_insurance_data,
                    pdf_texts,
                    [],
                    session.client_name
                )
                
                if "error" in file_data:
                    logger.error(f"Extraction error for {filename}: {file_data['error']}")
                    await update.message.reply_text(f"⚠️ Error extracting {filename}: {file_data['error']}")
                else:
                    # Normalize coverages (GPT may return list instead of dict)
                    _normalize_coverages(file_data)
                    
                    # Log what was found in this file
                    covs_found = list(file_data.get('coverages', {}).keys())
                    logger.info(f"File '{filename}' coverages: {covs_found}")
                    for key in covs_found:
                        cov = file_data['coverages'][key]
                        logger.info(f"  {key}: carrier={cov.get('carrier', 'N/A')}, "
                                   f"total_premium={cov.get('total_premium', 0)}")
                    
                    # Merge with existing data
                    session.extracted_data = _merge_extraction_results(
                        session.extracted_data, file_data
                    )
                    
                    covs_display = [c.replace('_', ' ') for c in covs_found]
                    await safe_reply(update,
                        f"\u2705 **{_escape_md(filename)}** \u2014 found: {', '.join(covs_display) if covs_display else 'no coverages'}",
                        parse_mode="Markdown"
                    )
                
            elif file_info["file_type"] in ["excel", "csv"]:
                local_path = file_info["local_path"]
                
                # Check if this is an SOV spreadsheet
                if file_info["file_type"] == "excel" and is_sov_file(local_path):
                    logger.info(f"Detected SOV spreadsheet: {filename}")
                    sov_data = await asyncio.to_thread(parse_sov, local_path)
                    
                    # Aggregate building-level rows into location-level summaries
                    if "error" not in sov_data:
                        sov_data = aggregate_locations(sov_data)
                    
                    if "error" in sov_data:
                        await update.message.reply_text(
                            f"\u26a0\ufe0f SOV parse error for {filename}: {sov_data['error']}"
                        )
                    else:
                        # Store SOV data in session
                        if session.extracted_data is None:
                            session.extracted_data = {}
                        session.extracted_data["sov_data"] = sov_data
                        
                        # Also populate locations from SOV if not already set
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
                        
                        session.extracted_data["locations"] = sov_locations
                        session.extracted_data["sov_totals"] = sov_data.get("totals", {})
                        
                        sov_summary = format_sov_summary(sov_data)
                        await safe_reply(update, f"\u2705 **{_escape_md(filename)}** \u2014 SOV parsed:\n\n{sov_summary}", parse_mode="Markdown")
                else:
                    # Generic Excel processing via GPT
                    data = extractor.extract_excel_data(local_path)
                    excel_data = [{"filename": filename, "data": data}]
                    file_data = await asyncio.to_thread(
                        extractor.structure_insurance_data,
                        [],
                        excel_data,
                        session.client_name
                    )
                    if "error" not in file_data:
                        _normalize_coverages(file_data)
                        session.extracted_data = _merge_extraction_results(
                            session.extracted_data, file_data
                        )
            
            session.processed_files.add(filename)
        
        # Final check
        if not session.extracted_data:
            await update.message.reply_text(
                "❌ Could not extract data from any files.\n"
                "Please check your documents and try again."
            )
            return WAITING_FOR_FILES
        
        # Normalize final merged data
        _normalize_coverages(session.extracted_data)
        
        # Log final merged results
        coverages_found = list(session.extracted_data.get('coverages', {}).keys())
        logger.info(f"Final merged extraction. Coverages: {coverages_found}")
        for key in coverages_found:
            cov = session.extracted_data['coverages'][key]
            logger.info(f"  {key}: carrier={cov.get('carrier', 'N/A')}, "
                       f"total_premium={cov.get('total_premium', 0)}")
        
        # Build verification summary
        summary = build_verification_summary(session.extracted_data)
        
        await safe_reply(update,
            f"📊 **Extraction Complete — Verification Checkpoint**\n\n"
            f"{summary}\n\n"
            f"**Commands:**\n"
            f"• /expiring — Set expiring premiums for comparison\n"
            f"• /override — Manually override a premium (e.g. /override UMB 15000)\n"
            f"• /generate — Accept and generate proposal\n"
            f"• /adjust [instructions] — Request changes\n"
            f"• /proposal\\_cancel — Cancel session",
            parse_mode="Markdown"
        )
        
        return REVIEWING_EXTRACTION
        
    except Exception as e:
        logger.error(f"Error extracting data: {e}", exc_info=True)
        await update.message.reply_text(
            f"❌ Error during extraction: {str(e)}\n\n"
            f"Please check your uploaded files and try again with /extract, "
            f"or start over with /proposal [Client Name]."
        )
        return WAITING_FOR_FILES


def build_verification_summary(data: dict) -> str:
    """Build a human-readable verification summary of extracted data."""
    _normalize_coverages(data)
    lines = []
    
    # Client Info
    ci = data.get("client_info", {})
    lines.append("**CLIENT INFORMATION**")
    lines.append(f"  Named Insured: {ci.get('named_insured', 'N/A')}")
    if ci.get("dba"):
        lines.append(f"  DBA: {ci['dba']}")
    lines.append(f"  Effective Date: {ci.get('effective_date', 'N/A')}")
    lines.append(f"  Address: {ci.get('address', 'N/A')}")
    lines.append("")
    
    # Locations
    locations = data.get("locations", [])
    lines.append(f"**LOCATIONS** ({len(locations)} found)")
    for loc in locations[:5]:  # Show first 5
        desc = loc.get("description", loc.get("address", ""))
        lines.append(f"  • {desc}")
    if len(locations) > 5:
        lines.append(f"  ... and {len(locations) - 5} more")
    lines.append("")
    
    # Coverage Summary
    coverages = data.get("coverages", {})
    coverage_names = {
        "property": "PROPERTY",
        "general_liability": "GENERAL LIABILITY",
        "umbrella": "UMBRELLA/EXCESS",
        "workers_comp": "WORKERS COMPENSATION",
        "commercial_auto": "COMMERCIAL AUTO",
        "cyber": "CYBER",
        "epli": "EPLI",
        "flood": "FLOOD",
        "terrorism": "TERRORISM / TRIA",
        "crime": "CRIME",
        "employee_benefits": "EMPLOYEE BENEFITS",
        "equipment_breakdown": "EQUIPMENT BREAKDOWN",
        "inland_marine": "INLAND MARINE",
        "umbrella_layer_2": "2ND EXCESS LAYER",
        "umbrella_layer_3": "3RD EXCESS LAYER"
    }
    
    total_premium = 0
    for key, display in coverage_names.items():
        cov = coverages.get(key)
        if cov:
            carrier = cov.get("carrier", "N/A")
            premium = cov.get("total_premium", 0)
            if premium is None:
                premium = 0
            elif not isinstance(premium, (int, float)):
                try:
                    premium = float(str(premium).replace(",", "").replace("$", ""))
                except (ValueError, TypeError):
                    premium = 0
            total_premium += premium
            admitted = "Admitted" if cov.get("carrier_admitted", True) else "Non-Admitted"
            
            lines.append(f"**{display}**")
            lines.append(f"  Carrier: {carrier} ({admitted})")
            lines.append(f"  Premium: ${premium:,.2f}" if isinstance(premium, (int, float)) else f"  Premium: {premium}")
            
            # Key limits
            limits = cov.get("limits", [])
            if limits and isinstance(limits, list):
                for lim in limits[:3]:
                    if isinstance(lim, dict):
                        lines.append(f"  {lim.get('description', '')}: {lim.get('limit', '')}")
                    elif isinstance(lim, str):
                        lines.append(f"  {lim}")
            
            # Deductibles
            deds = cov.get("deductibles", [])
            if deds and isinstance(deds, list):
                for ded in deds[:2]:
                    if isinstance(ded, dict):
                        lines.append(f"  Deductible: {ded.get('description', '')} — {ded.get('amount', '')}")
                    elif isinstance(ded, str):
                        lines.append(f"  Deductible: {ded}")
            
            # Forms count
            forms = cov.get("forms_endorsements", [])
            if forms:
                lines.append(f"  Forms/Endorsements: {len(forms)} listed")
            
            lines.append("")
    
    lines.append(f"**TOTAL PROPOSED PREMIUM: ${total_premium:,.2f}**")
    
    return "\n".join(lines)


async def adjust_data(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Handle adjustment requests to extracted data."""
    chat_id = update.effective_chat.id
    session = get_session(chat_id)
    
    if not session or not session.extracted_data:
        await update.message.reply_text("No extracted data to adjust. Run /extract first.")
        return REVIEWING_EXTRACTION
    
    instructions = " ".join(context.args) if context.args else update.message.text.replace("/adjust", "").strip()
    
    if not instructions:
        await update.message.reply_text(
            "Please provide adjustment instructions.\n"
            "Example: /adjust Change the property deductible to $25,000"
        )
        return REVIEWING_EXTRACTION
    
    await update.message.reply_text("⏳ Applying adjustments...")
    
    try:
        extractor = ProposalExtractor()
        updated_data = await asyncio.to_thread(
            extractor.apply_adjustments,
            session.extracted_data,
            instructions
        )
        
        session.extracted_data = updated_data
        summary = build_verification_summary(updated_data)
        
        await safe_reply(update,
            f"✅ **Adjustments Applied**\n\n"
            f"{summary}\n\n"
            f"• /generate — Accept and generate proposal\n"
            f"• /adjust [instructions] — More changes\n"
            f"• /proposal\\_cancel — Cancel",
            parse_mode="Markdown"
        )
        
    except Exception as e:
        logger.error(f"Error applying adjustments: {e}", exc_info=True)
        await update.message.reply_text(f"Error applying adjustments: {e}")
    
    return REVIEWING_EXTRACTION


async def generate_doc(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Generate the final DOCX proposal."""
    logger.info(f"generate_doc called by user {update.effective_user.id}")
    chat_id = update.effective_chat.id
    session = get_session(chat_id)
    
    if not session:
        await update.message.reply_text("No active proposal session. Start one with /proposal [Client Name]")
        return ConversationHandler.END
    
    # Auto-extract if not done yet
    if not session.extracted_data:
        if not session.uploaded_files:
            await update.message.reply_text("No files uploaded yet. Please upload documents first, then /extract or /generate.")
            return WAITING_FOR_FILES
        
        logger.info("Auto-extracting data before generating document")
        # Run extraction first
        result = await extract_data(update, context)
        session = get_session(chat_id)  # Re-fetch in case extraction updated it
        if not session or not session.extracted_data:
            return result if result is not None else WAITING_FOR_FILES
    
    # Check if extraction returned an error
    if isinstance(session.extracted_data, dict) and "error" in session.extracted_data:
        await update.message.reply_text(
            f"\u274c Extraction had an error: {session.extracted_data['error']}\n\n"
            f"Please try /extract again or upload different files."
        )
        return WAITING_FOR_FILES
    
    # Normalize coverages before generating
    _normalize_coverages(session.extracted_data)
    
    # Log what we have
    coverages = session.extracted_data.get('coverages', {})
    logger.info(f"Generating document with coverages: {list(coverages.keys())}")
    for key, cov in coverages.items():
        logger.info(f"  {key}: carrier={cov.get('carrier', 'N/A')}, premium={cov.get('total_premium', 0)}")
    
    # Enrich GL schedule_of_classes with SOV address/brand data
    sov_data = session.extracted_data.get('sov_data', {})
    sov_locations = sov_data.get('locations', []) if sov_data else []
    gl_cov = coverages.get('general_liability', {})
    if gl_cov and sov_locations:
        classes = gl_cov.get('schedule_of_classes', [])
        if classes:
            # Build lookup from location number to SOV data
            sov_lookup = {}
            for loc in sov_locations:
                loc_num = loc.get('location_num', loc.get('building_num', 0))
                if loc_num:
                    sov_lookup[str(loc_num)] = loc
            
            for cls_entry in classes:
                if not isinstance(cls_entry, dict):
                    continue
                # If no address/brand, try to enrich from SOV
                if not cls_entry.get('address') or not cls_entry.get('brand_dba'):
                    loc_str = str(cls_entry.get('location', ''))
                    # Extract location number from strings like "Loc 1", "Location 1", "1"
                    import re
                    loc_match = re.search(r'(\d+)', loc_str)
                    if loc_match:
                        loc_num = loc_match.group(1)
                        sov_loc = sov_lookup.get(loc_num)
                        if sov_loc:
                            if not cls_entry.get('address'):
                                addr = sov_loc.get('address', '')
                                city = sov_loc.get('city', '')
                                state = sov_loc.get('state', '')
                                if addr:
                                    cls_entry['address'] = f"{addr}, {city}, {state}" if city else addr
                            if not cls_entry.get('brand_dba'):
                                cls_entry['brand_dba'] = sov_loc.get('dba', '') or sov_loc.get('hotel_flag', '')
            logger.info(f"Enriched {len(classes)} GL schedule_of_classes entries with SOV data")
    
    # Log expiring data
    exp_premiums = session.extracted_data.get('expiring_premiums', {})
    exp_details = session.extracted_data.get('expiring_details', {})
    logger.info(f"Expiring premiums in extracted_data: {exp_premiums}")
    logger.info(f"Expiring details keys in extracted_data: {list(exp_details.keys())}")
    logger.info(f"Full extracted_data top-level keys: {list(session.extracted_data.keys())}")
    logger.info(f"Full extracted_data: {session.extracted_data}"[:2000])
    
    await safe_reply(update,
        "📝 **Generating proposal document...**\n\n"
        "Creating branded DOCX with all coverage sections, compliance pages, and signature blocks.\n"
        "This may take a moment.",
        parse_mode="Markdown"
    )
    
    try:
        # Generate filename
        client_name = session.client_name.replace(" ", "_").replace("/", "-")
        timestamp = datetime.now().strftime("%Y%m%d")
        docx_filename = f"HUB_Proposal_{client_name}_{timestamp}.docx"
        docx_path = os.path.join(session.work_dir, docx_filename)
        
        # Generate the DOCX
        await asyncio.to_thread(
            generate_proposal,
            session.extracted_data,
            docx_path
        )
        
        # Send the file
        # Escape underscores in dynamic strings for Telegram Markdown
        safe_filename = docx_filename.replace("_", "\\_")
        safe_client = session.client_name.replace("_", "\\_")
        with open(docx_path, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=docx_filename,
                caption=(
                    f"\u2705 **Proposal Generated**\n\n"
                    f"**Client:** {safe_client}\n"
                    f"**File:** {safe_filename}\n\n"
                    f"Review the document and make any final edits as needed.\n"
                    f"Send /proposal to start a new proposal."
                ),
                parse_mode="Markdown"
            )
        
        # Cleanup session
        clear_session(chat_id)
        
        return ConversationHandler.END
        
    except Exception as e:
        logger.error(f"Error generating proposal: {e}", exc_info=True)
        safe_err = str(e).replace("_", " ")
        await update.message.reply_text(
            f"\u274c Error generating proposal: {safe_err}\n\n"
            f"Try /generate again or /proposal\_cancel to start over."
        )
        return REVIEWING_EXTRACTION


def _parse_dollar(s: str) -> float:
    """Parse a dollar string like '$61,487' or '61487' into a float."""
    return float(s.replace("$", "").replace(",", "").replace("\\", "").strip())


def _parse_expiring_block(raw_text: str) -> tuple:
    """Parse multi-line expiring coverage block.
    
    Accepts format like:
        PROP — Tower Hill Insurance
            Premium: $61,487
            TIV: $15,042,080
            ...
            💬 AA including TRIA
        
        GL — Southlake Specialty Insurance Company
            Premium: $49,483
            ...
    
    Returns (expiring_premiums dict, expiring_details dict, parsed_summary list)
    """
    import re
    
    # Coverage abbreviation mapping
    coverage_map = {
        "prop": "property", "property": "property",
        "gl": "general_liability", "liability": "general_liability",
        "general_liability": "general_liability",
        "umb": "umbrella", "umbrella": "umbrella", "excess": "umbrella",
        "wc": "workers_comp", "workers": "workers_comp",
        "workers_comp": "workers_comp", "comp": "workers_comp",
        "auto": "commercial_auto", "commercial_auto": "commercial_auto",
        "flood": "flood", "epli": "epli", "cyber": "cyber",
        "crime": "crime", "crim": "crime",
        "eb": "employee_benefits", "employee_benefits": "employee_benefits",
        "equipment_breakdown": "equipment_breakdown",
        "equipment": "equipment_breakdown",
        "inland_marine": "inland_marine",
        "im": "inland_marine", "bop": "bop",
    }
    
    display_names = {
        "property": "Property", "general_liability": "General Liability",
        "umbrella": "Umbrella", "workers_comp": "Workers Comp",
        "commercial_auto": "Commercial Auto", "flood": "Flood",
        "epli": "EPLI", "cyber": "Cyber", "crime": "Crime",
        "employee_benefits": "Employee Benefits",
        "equipment_breakdown": "Equipment Breakdown",
        "inland_marine": "Inland Marine", "bop": "BOP",
    }
    
    expiring_premiums = {}  # key -> premium amount
    expiring_details = {}   # key -> {carrier, premium, details: {}, notes}
    parsed_summary = []
    
    # Strip Telegram's escaped dollar signs: \$ -> $
    raw_text = raw_text.replace("\\$", "$")
    
    lines = raw_text.strip().split("\n")
    current_key = None
    current_entry = None
    
    # Pattern for coverage header: "PROP — Carrier Name" or "GL - Carrier Name"
    header_pattern = re.compile(
        r'^\s*([A-Za-z_]+)\s*[\u2014\-\u2013]+\s*(.+)$'
    )
    # Pattern for detail line: "  Key: Value"
    detail_pattern = re.compile(r'^\s+([^:]+):\s*(.+)$')
    # Pattern for note line: "  💬 some note"
    note_pattern = re.compile(r'^\s*💬\s*(.+)$')
    
    for line in lines:
        line_stripped = line.strip()
        if not line_stripped:
            continue
        
        # Check for coverage header
        header_match = header_pattern.match(line)
        if header_match:
            # Save previous entry
            if current_key and current_entry:
                cov_key = coverage_map.get(current_key.lower())
                if cov_key:
                    expiring_details[cov_key] = current_entry
                    if current_entry.get("premium"):
                        expiring_premiums[cov_key] = current_entry["premium"]
                        parsed_summary.append(
                            f"\u2022 **{display_names.get(cov_key, cov_key)}** — "
                            f"{current_entry.get('carrier', 'N/A')}: "
                            f"${current_entry['premium']:,.0f}"
                        )
            
            abbrev = header_match.group(1).strip()
            carrier = header_match.group(2).strip()
            current_key = abbrev.lower()
            current_entry = {
                "carrier": carrier,
                "premium": 0,
                "details": {},
                "notes": ""
            }
            continue
        
        # Check for note line
        note_match = note_pattern.match(line)
        if note_match and current_entry:
            current_entry["notes"] = note_match.group(1).strip()
            continue
        
        # Check for detail line
        detail_match = detail_pattern.match(line)
        if detail_match and current_entry:
            field_name = detail_match.group(1).strip()
            field_value = detail_match.group(2).strip()
            
            # Store the raw detail
            current_entry["details"][field_name] = field_value
            
            # If this is the Premium field, extract the number
            if field_name.lower() == "premium":
                try:
                    current_entry["premium"] = _parse_dollar(field_value)
                except (ValueError, TypeError):
                    pass
            continue
    
    # Save the last entry
    if current_key and current_entry:
        cov_key = coverage_map.get(current_key.lower())
        if cov_key:
            expiring_details[cov_key] = current_entry
            if current_entry.get("premium"):
                expiring_premiums[cov_key] = current_entry["premium"]
                parsed_summary.append(
                    f"\u2022 **{display_names.get(cov_key, cov_key)}** — "
                    f"{current_entry.get('carrier', 'N/A')}: "
                    f"${current_entry['premium']:,.0f}"
                )
    
    return expiring_premiums, expiring_details, parsed_summary


async def set_expiring(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Set expiring premiums and details for the proposal.
    
    Accepts two formats:
    
    1. Simple inline:
        /expiring property 60000 gl 50000 umbrella 5000
    
    2. Rich multi-line:
        /expiring
        PROP — Tower Hill Insurance
            Premium: $61,487
            TIV: $15,042,080
            AOP Deductible: $5,000
            💬 AA including TRIA
        
        GL — Southlake Specialty
            Premium: $49,483
            Total Sales: $4,000,000
            💬 inc $1M EPLI
    """
    chat_id = update.effective_chat.id
    session = get_session(chat_id)
    
    if not session:
        await update.message.reply_text("No active proposal session. Start one with /proposal [Client Name]")
        return ConversationHandler.END
    
    raw_text = update.message.text.replace("/expiring", "").strip()
    # Strip Telegram's escaped dollar signs: \$ -> $
    raw_text = raw_text.replace("\\$", "$")
    
    if not raw_text:
        await update.message.reply_text(
            "\u2139\ufe0f **Set Expiring Premiums**\n\n"
            "Paste your expiring program details below.\n"
            "I'm waiting for your next message.\n\n"
            "Example format:\n"
            "`PROP \u2014 Tower Hill Insurance`\n"
            "`    Premium: $61,487`\n"
            "`    TIV: $15,042,080`\n"
            "`    AOP Deductible: $5,000`\n\n"
            "`GL \u2014 Southlake Specialty`\n"
            "`    Premium: $49,483`\n"
            "`    Total Sales: $4,000,000`\n\n"
            "**Coverage abbreviations:** PROP, GL, UMB, WC, AUTO, FLOOD, EPLI, CYBER\n\n"
            "Or use simple format: `property 60000 gl 50000`",
            parse_mode="Markdown"
        )
        return WAITING_FOR_EXPIRING
    
    # Detect format: multi-line (has \u2014 or - with carrier) vs simple (key value pairs)
    import re
    has_header = bool(re.search(r'[A-Za-z_]+\s*[\u2014\-\u2013]+\s*.+', raw_text))
    
    if has_header:
        # Rich multi-line format
        expiring_premiums, expiring_details, parsed_summary = _parse_expiring_block(raw_text)
    else:
        # Simple inline format: property 60000 gl 50000
        simple_map = {
            "property": "property", "prop": "property",
            "gl": "general_liability", "liability": "general_liability",
            "umbrella": "umbrella", "umb": "umbrella", "excess": "umbrella",
            "wc": "workers_comp", "workers": "workers_comp", "comp": "workers_comp",
            "auto": "commercial_auto",
            "flood": "flood", "epli": "epli", "cyber": "cyber",
        }
        simple_display = {
            "property": "Property", "general_liability": "General Liability",
            "umbrella": "Umbrella", "workers_comp": "Workers Comp",
            "commercial_auto": "Commercial Auto", "flood": "Flood",
            "epli": "EPLI", "cyber": "Cyber",
        }
        tokens = raw_text.replace(",", "").replace("$", "").split()
        expiring_premiums = {}
        expiring_details = {}
        parsed_summary = []
        i = 0
        while i < len(tokens):
            tok = tokens[i].lower()
            if tok in simple_map:
                cov_key = simple_map[tok]
                if i + 1 < len(tokens):
                    try:
                        amount = float(tokens[i + 1])
                        expiring_premiums[cov_key] = amount
                        parsed_summary.append(
                            f"\u2022 **{simple_display.get(cov_key, cov_key)}**: ${amount:,.0f}"
                        )
                        i += 2
                        continue
                    except ValueError:
                        pass
            i += 1
    
    if not expiring_premiums:
        await safe_reply(update,
            "\u26a0\ufe0f Could not parse any expiring premiums.\n\n"
            "Make sure each coverage section has a `Premium: $XX,XXX` line.\n"
            "Or use simple format: `/expiring property 60000 gl 50000`",
            parse_mode="Markdown"
        )
        if session.extracted_data:
            return REVIEWING_EXTRACTION
        return WAITING_FOR_FILES
    
    # Store in session data
    if not session.extracted_data:
        session.extracted_data = {
            "expiring_premiums": expiring_premiums,
            "expiring_details": expiring_details
        }
    else:
        session.extracted_data["expiring_premiums"] = expiring_premiums
        session.extracted_data["expiring_details"] = expiring_details
    
    # Build response
    response = "\u2705 **Expiring Program Set**\n\n"
    response += "\n".join(parsed_summary) + "\n"
    
    # Show details for rich format
    if expiring_details:
        for cov_key, entry in expiring_details.items():
            if entry.get("notes"):
                display = {
                    "property": "Property", "general_liability": "GL",
                    "umbrella": "Umbrella", "workers_comp": "WC",
                    "commercial_auto": "Auto", "flood": "Flood",
                }.get(cov_key, cov_key)
                response += f"  💬 {display}: {entry['notes']}\n"
    
    total_exp = sum(v for v in expiring_premiums.values() if isinstance(v, (int, float)))
    response += f"\n**Total Expiring: ${total_exp:,.0f}**\n\n"
    
    if session.extracted_data and session.extracted_data.get("coverages"):
        response += (
            "Send /generate to create the proposal, "
            "or /extract to re-extract data."
        )
    else:
        response += "Upload quote documents and send /extract to continue."
    
    await safe_reply(update, response, parse_mode="Markdown")
    
    if session.extracted_data and session.extracted_data.get("coverages"):
        return REVIEWING_EXTRACTION
    return WAITING_FOR_FILES


async def receive_expiring_text(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Receive pasted expiring program text after /expiring was called with no args."""
    chat_id = update.effective_chat.id
    session = get_session(chat_id)
    
    if not session:
        await update.message.reply_text("No active proposal session. Start one with /proposal [Client Name]")
        return ConversationHandler.END
    
    raw_text = update.message.text.strip()
    
    if not raw_text:
        await update.message.reply_text("Please paste your expiring program details, or send /proposal_cancel to cancel.")
        return WAITING_FOR_EXPIRING
    
    # Strip Telegram's escaped dollar signs: \$ -> $
    raw_text = raw_text.replace("\\$", "$")
    
    # Reuse the same parsing logic from set_expiring
    import re
    has_header = bool(re.search(r'[A-Za-z_]+\s*[\u2014\-\u2013]+\s*.+', raw_text))
    
    if has_header:
        expiring_premiums, expiring_details, parsed_summary = _parse_expiring_block(raw_text)
    else:
        # Simple inline format: property 60000 gl 50000
        simple_map = {
            "property": "property", "prop": "property",
            "gl": "general_liability", "liability": "general_liability",
            "umbrella": "umbrella", "umb": "umbrella", "excess": "umbrella",
            "wc": "workers_comp", "workers": "workers_comp", "comp": "workers_comp",
            "auto": "commercial_auto",
            "flood": "flood", "epli": "epli", "cyber": "cyber",
        }
        simple_display = {
            "property": "Property", "general_liability": "General Liability",
            "umbrella": "Umbrella", "workers_comp": "Workers Comp",
            "commercial_auto": "Commercial Auto", "flood": "Flood",
            "epli": "EPLI", "cyber": "Cyber",
        }
        tokens = raw_text.replace(",", "").replace("$", "").split()
        expiring_premiums = {}
        expiring_details = {}
        parsed_summary = []
        i = 0
        while i < len(tokens):
            tok = tokens[i].lower()
            if tok in simple_map:
                cov_key = simple_map[tok]
                if i + 1 < len(tokens):
                    try:
                        amount = float(tokens[i + 1])
                        expiring_premiums[cov_key] = amount
                        parsed_summary.append(
                            f"\u2022 **{simple_display.get(cov_key, cov_key)}**: ${amount:,.0f}"
                        )
                        i += 2
                        continue
                    except ValueError:
                        pass
            i += 1
    
    if not expiring_premiums:
        await safe_reply(update,
            "\u26a0\ufe0f Could not parse any expiring premiums.\n\n"
            "Make sure each coverage section has a `Premium: $XX,XXX` line.\n"
            "Or use simple format: `property 60000 gl 50000`\n\n"
            "Try pasting again, or send /proposal_cancel to cancel.",
            parse_mode="Markdown"
        )
        return WAITING_FOR_EXPIRING
    
    # Store in session data
    if not session.extracted_data:
        session.extracted_data = {
            "expiring_premiums": expiring_premiums,
            "expiring_details": expiring_details
        }
    else:
        session.extracted_data["expiring_premiums"] = expiring_premiums
        session.extracted_data["expiring_details"] = expiring_details
    
    # Build response
    response = "\u2705 **Expiring Program Set**\n\n"
    response += "\n".join(parsed_summary) + "\n"
    
    if expiring_details:
        for cov_key, entry in expiring_details.items():
            if entry.get("notes"):
                display = {
                    "property": "Property", "general_liability": "GL",
                    "umbrella": "Umbrella", "workers_comp": "WC",
                    "commercial_auto": "Auto", "flood": "Flood",
                }.get(cov_key, cov_key)
                response += f"  Notes {display}: {entry['notes']}\n"
    
    total_exp = sum(v for v in expiring_premiums.values() if isinstance(v, (int, float)))
    response += f"\n**Total Expiring: ${total_exp:,.0f}**\n\n"
    
    if session.extracted_data and session.extracted_data.get("coverages"):
        response += (
            "Send /generate to create the proposal, "
            "or /extract to re-extract data."
        )
    else:
        response += "Upload quote documents and send /extract to continue."
    
    await safe_reply(update, response, parse_mode="Markdown")
    
    if session.extracted_data and session.extracted_data.get("coverages"):
        return REVIEWING_EXTRACTION
    return WAITING_FOR_FILES


async def proposal_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Cancel the current proposal session."""
    chat_id = update.effective_chat.id
    clear_session(chat_id)
    await update.message.reply_text("❌ Proposal session cancelled.")
    return ConversationHandler.END


async def proposal_status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Show current proposal session status."""
    chat_id = update.effective_chat.id
    session = get_session(chat_id)
    
    if not session:
        await update.message.reply_text(
            "No active proposal session.\n"
            "Start one with: /proposal [Client Name]"
        )
        return
    
    status_lines = [
        f"📋 **Active Proposal Session**\n",
        f"**Client:** {session.client_name}",
        f"**Started:** {session.created_at.strftime('%I:%M %p')}",
        f"**Files uploaded:** {len(session.uploaded_files)}",
        session.get_file_summary(),
        f"**Data extracted:** {'Yes' if session.extracted_data else 'No'}",
    ]
    
    await safe_reply(update, "\n".join(status_lines), parse_mode="Markdown")


async def override_premium(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Manually override the total premium for one or more coverages.
    
    Usage:
        /override GL 44650.25
        /override GL 44650.25 UMB 15000 CYBER 5000
        /override GL 44650.25, UMB 15000, CYBER 5000
    
    This sets the total_premium in the extracted data for the specified coverage(s).
    Useful when agency commissions, fees, or taxes need to be manually added.
    """
    chat_id = update.effective_chat.id
    session = get_session(chat_id)
    
    if not session:
        await update.message.reply_text("No active proposal session. Start one with /proposal [Client Name]")
        return ConversationHandler.END
    
    if not session.extracted_data or not session.extracted_data.get("coverages"):
        await update.message.reply_text("No extracted data yet. Upload files and run /extract first.")
        return REVIEWING_EXTRACTION
    
    # Normalize coverages in case GPT returned a list
    _normalize_coverages(session.extracted_data)
    
    raw_text = update.message.text.replace("/override", "").strip()
    raw_text = raw_text.replace("\\$", "$")  # Strip Telegram escaped dollars
    
    if not raw_text:
        # Show current premiums and usage
        coverages = session.extracted_data.get("coverages", {})
        lines = ["\u2139\ufe0f **Manual Premium Override**\n"]
        lines.append("Current premiums (total w/ taxes & fees):")
        for key, cov in coverages.items():
            display = key.replace('_', ' ').title()
            tp = cov.get('total_premium', 0)
            lines.append(f"  {display}: ${tp:,.2f}" if isinstance(tp, (int, float)) else f"  {display}: {tp}")
        lines.append("")
        lines.append("**Usage:** `/override COVERAGE AMOUNT`")
        lines.append("**Examples:**")
        lines.append("`/override GL 44650.25`")
        lines.append("`/override GL 44500 UMB 18500 CYBER 5000`")
        lines.append("`/override GL 44500, UMB 18500, CYBER 5000`")
        lines.append("")
        lines.append("**Coverage abbreviations:** PROP, GL, UMB, WC, AUTO, FLOOD, EPLI, CYBER, TERR")
        await safe_reply(update, "\n".join(lines), parse_mode="Markdown")
        return REVIEWING_EXTRACTION
    
    # Parse: COVERAGE AMOUNT
    abbrev_map = {
        "prop": "property", "property": "property",
        "gl": "general_liability", "liability": "general_liability",
        "umb": "umbrella", "umbrella": "umbrella", "excess": "umbrella",
        "wc": "workers_comp", "workers": "workers_comp", "comp": "workers_comp",
        "auto": "commercial_auto",
        "flood": "flood", "epli": "epli", "cyber": "cyber",
        "terr": "terrorism", "terrorism": "terrorism", "tria": "terrorism",
        "crime": "crime", "crim": "crime",
        "eb": "employee_benefits", "employee_benefits": "employee_benefits",
        "equipment_breakdown": "equipment_breakdown",
        "equipment": "equipment_breakdown",
        "inland": "inland_marine",
    }
    display_names = {
        "property": "Property", "general_liability": "General Liability",
        "umbrella": "Umbrella", "workers_comp": "Workers Comp",
        "commercial_auto": "Commercial Auto", "flood": "Flood",
        "epli": "EPLI", "cyber": "Cyber", "terrorism": "Terrorism/TRIA",
        "crime": "Crime", "employee_benefits": "Employee Benefits",
        "equipment_breakdown": "Equipment Breakdown",
        "inland_marine": "Inland Marine",
    }
    
    # Strip commas used as separators between pairs, and dollar signs
    raw_text = raw_text.replace(",", " ").replace("$", "").strip()
    tokens = raw_text.split()
    
    if len(tokens) < 2:
        await update.message.reply_text(
            "Please provide both coverage and amount.\n"
            "Example: /override GL 44650.25\n"
            "Multiple: /override GL 44500 UMB 18500 CYBER 5000"
        )
        return REVIEWING_EXTRACTION
    
    coverages = session.extracted_data.get("coverages", {})
    
    # Parse pairs: iterate tokens looking for COVERAGE AMOUNT pairs
    results = []  # List of (display_name, old_premium, new_premium)
    errors = []
    i = 0
    while i < len(tokens):
        token = tokens[i]
        cov_abbrev = token.lower()
        cov_key = abbrev_map.get(cov_abbrev)
        
        if not cov_key:
            # Try to see if it's a number (stray amount without coverage)
            try:
                float(token)
                errors.append(f"Amount '{token}' has no coverage before it")
            except ValueError:
                errors.append(f"Unknown coverage: {token}")
            i += 1
            continue
        
        # Next token should be the amount
        if i + 1 >= len(tokens):
            errors.append(f"No amount provided for {token.upper()}")
            i += 1
            continue
        
        amount_str = tokens[i + 1]
        try:
            amount = float(amount_str)
        except ValueError:
            errors.append(f"Invalid amount for {token.upper()}: {amount_str}")
            i += 2
            continue
        
        if cov_key not in coverages:
            # Create a new coverage entry with just the premium
            coverages[cov_key] = {
                "carrier": "TBD",
                "carrier_admitted": False,
                "premium": amount,
                "taxes_fees": 0,
                "total_premium": amount,
                "limits": [],
                "forms_endorsements": [],
            }
            display = display_names.get(cov_key, cov_key)
            results.append((display + " (NEW)", 0, amount))
            i += 2
            continue
        
        old_premium = coverages[cov_key].get("total_premium", 0)
        coverages[cov_key]["total_premium"] = amount
        display = display_names.get(cov_key, cov_key)
        results.append((display, old_premium, amount))
        i += 2
    
    if not results and errors:
        await update.message.reply_text(
            "Could not process overrides:\n" + "\n".join(f"  • {e}" for e in errors) +
            "\n\nExample: /override GL 44500 UMB 18500"
        )
        return REVIEWING_EXTRACTION
    
    # Build confirmation message
    lines = ["\u2705 **Premium Override(s) Applied**\n"]
    for display, old_val, new_val in results:
        lines.append(f"**{display}:**")
        lines.append(f"  Previous: ${old_val:,.2f}")
        lines.append(f"  Updated: ${new_val:,.2f}")
        lines.append("")
    
    if errors:
        lines.append("\u26a0\ufe0f **Warnings:**")
        for e in errors:
            lines.append(f"  • {e}")
        lines.append("")
    
    lines.append("Send /generate to create the proposal with updated premiums, ")
    lines.append("or /override again for more changes.")
    
    await safe_reply(update, "\n".join(lines), parse_mode="Markdown")
    return REVIEWING_EXTRACTION


def get_proposal_conversation_handler() -> ConversationHandler:
    """Create and return the ConversationHandler for /proposal."""
    return ConversationHandler(
        entry_points=[CommandHandler("proposal", proposal_start)],
        states={
            WAITING_FOR_FILES: [
                MessageHandler(filters.Document.ALL, receive_file),
                CommandHandler("extract", extract_data),
                CommandHandler("expiring", set_expiring),
                CommandHandler("override", override_premium),
                CommandHandler("generate", generate_doc),  # Auto-extract if needed
                CommandHandler("proposal_cancel", proposal_cancel),
            ],
            REVIEWING_EXTRACTION: [
                MessageHandler(filters.Document.ALL, receive_file),  # Accept more files after extraction
                CommandHandler("generate", generate_doc),
                CommandHandler("adjust", adjust_data),
                CommandHandler("expiring", set_expiring),
                CommandHandler("override", override_premium),
                CommandHandler("extract", extract_data),  # Re-extract
                CommandHandler("proposal_cancel", proposal_cancel),
            ],
            WAITING_FOR_EXPIRING: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_expiring_text),
                MessageHandler(filters.Document.ALL, receive_file),
                CommandHandler("expiring", set_expiring),
                CommandHandler("override", override_premium),
                CommandHandler("extract", extract_data),
                CommandHandler("generate", generate_doc),
                CommandHandler("proposal_cancel", proposal_cancel),
            ],
        },
        fallbacks=[
            CommandHandler("extract", extract_data),
            CommandHandler("expiring", set_expiring),
            CommandHandler("override", override_premium),
            CommandHandler("generate", generate_doc),
            CommandHandler("proposal_cancel", proposal_cancel),
            CommandHandler("proposal", proposal_start),  # Restart
        ],
        per_chat=True,
        per_user=True,
    )


async def extract_standalone(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Standalone /extract handler for when no conversation is active."""
    chat_id = update.effective_chat.id
    session = get_session(chat_id)
    if session:
        # There IS a session but the ConversationHandler didn't catch it
        # This can happen after a redeployment. Try to process anyway.
        logger.info(f"extract_standalone: found orphaned session for chat {chat_id}, processing")
        await extract_data(update, context)
    else:
        await update.message.reply_text(
            "No active proposal session.\n"
            "Start one with: /proposal [Client Name]\n"
            "Then upload your quote documents and send /extract."
        )


async def generate_standalone(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Standalone /generate handler for when no conversation is active."""
    chat_id = update.effective_chat.id
    session = get_session(chat_id)
    if session:
        logger.info(f"generate_standalone: found orphaned session for chat {chat_id}, processing")
        await generate_doc(update, context)
    else:
        await update.message.reply_text(
            "No active proposal session.\n"
            "Start one with: /proposal [Client Name]\n"
            "Then upload documents, /extract, and /generate."
        )
