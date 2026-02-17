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

logger = logging.getLogger(__name__)

# Conversation states
(
    WAITING_FOR_FILES,
    WAITING_FOR_MORE_FILES,
    REVIEWING_EXTRACTION,
    CONFIRMING_GENERATION,
) = range(4)

# Session storage (in-memory, per chat)
proposal_sessions = {}


class ProposalSession:
    """Tracks state for an active proposal generation session."""
    
    def __init__(self, client_name: str, chat_id: int):
        self.client_name = client_name
        self.chat_id = chat_id
        self.uploaded_files = []  # List of (filename, local_path, file_type)
        self.extracted_data = None
        self.created_at = datetime.now()
        self.work_dir = tempfile.mkdtemp(prefix="proposal_")
    
    def add_file(self, filename: str, local_path: str, file_type: str):
        self.uploaded_files.append({
            "filename": filename,
            "local_path": local_path,
            "file_type": file_type
        })
    
    def get_file_summary(self) -> str:
        if not self.uploaded_files:
            return "No files uploaded yet."
        lines = []
        for i, f in enumerate(self.uploaded_files, 1):
            lines.append(f"  {i}. {f['filename']} ({f['file_type']})")
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
    """Send a message, splitting if too long for Telegram's 4096 char limit."""
    MAX_LEN = 4000
    if len(text) <= MAX_LEN:
        await update.message.reply_text(text, **kwargs)
        return
    
    # Split at line breaks
    lines = text.split("\n")
    chunk = ""
    for line in lines:
        if len(chunk) + len(line) + 1 > MAX_LEN:
            if chunk:
                await update.message.reply_text(chunk, **kwargs)
            chunk = line
        else:
            chunk = chunk + "\n" + line if chunk else line
    if chunk:
        await update.message.reply_text(chunk, **kwargs)


# ─── Command Handlers ─────────────────────────────────────────

async def proposal_start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Start a new proposal session: /proposal [Client Name]"""
    logger.info(f"proposal_start called by user {update.effective_user.id} in chat {update.effective_chat.id}")
    logger.info(f"Raw message text: {update.message.text}")
    args = context.args
    if not args:
        await update.message.reply_text(
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
    
    await update.message.reply_text(
        f"📋 **New Proposal Session Started**\n\n"
        f"**Client:** {client_name}\n\n"
        f"Please upload your insurance quote documents:\n"
        f"• Property quote (PDF)\n"
        f"• General Liability quote (PDF)\n"
        f"• Umbrella/Excess quote (PDF)\n"
        f"• Workers Compensation quote (PDF)\n"
        f"• Commercial Auto quote (PDF)\n"
        f"• Schedule of Values / SOV (Excel)\n\n"
        f"Upload files one at a time. When done, send /extract to process.\n"
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
        
        session.add_file(filename, local_path, file_type)
        
        file_count = len(session.uploaded_files)
        await update.message.reply_text(
            f"✅ Received: **{filename}** ({file_type.upper()})\n\n"
            f"**Files uploaded ({file_count}):**\n{session.get_file_summary()}\n\n"
            f"Upload more files or send /extract when ready.",
            parse_mode="Markdown"
        )
        
    except Exception as e:
        logger.error(f"Error downloading file: {e}")
        await update.message.reply_text(f"Error downloading file: {e}")
    
    return WAITING_FOR_FILES


async def extract_data(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Process uploaded files and extract insurance data."""
    logger.info(f"extract_data called by user {update.effective_user.id}")
    chat_id = update.effective_chat.id
    session = get_session(chat_id)
    
    if not session:
        await update.message.reply_text("No active proposal session. Start one with /proposal [Client Name]")
        return ConversationHandler.END
    
    if not session.uploaded_files:
        await update.message.reply_text("No files uploaded yet. Please upload at least one quote document.")
        return WAITING_FOR_FILES
    
    await update.message.reply_text(
        f"⏳ **Processing {len(session.uploaded_files)} file(s)...**\n\n"
        f"Extracting text, analyzing coverages, and structuring data.\n"
        f"This may take 1-2 minutes.",
        parse_mode="Markdown"
    )
    
    try:
        extractor = ProposalExtractor()
        
        # Extract text from all files
        all_texts = []
        excel_data = []
        
        for file_info in session.uploaded_files:
            if file_info["file_type"] == "pdf":
                text = extractor.extract_pdf_text(file_info["local_path"])
                all_texts.append({
                    "filename": file_info["filename"],
                    "text": text
                })
            elif file_info["file_type"] in ["excel", "csv"]:
                data = extractor.extract_excel_data(file_info["local_path"])
                excel_data.append({
                    "filename": file_info["filename"],
                    "data": data
                })
        
        # Use GPT to structure the data
        structured_data = await asyncio.to_thread(
            extractor.structure_insurance_data,
            all_texts,
            excel_data,
            session.client_name
        )
        
        session.extracted_data = structured_data
        
        # Build verification summary
        summary = build_verification_summary(structured_data)
        
        await safe_reply(update,
            f"📊 **Extraction Complete — Verification Checkpoint**\n\n"
            f"{summary}\n\n"
            f"**Commands:**\n"
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
        "commercial_auto": "COMMERCIAL AUTO"
    }
    
    total_premium = 0
    for key, display in coverage_names.items():
        cov = coverages.get(key)
        if cov:
            carrier = cov.get("carrier", "N/A")
            premium = cov.get("total_premium", 0)
            total_premium += premium
            admitted = "Admitted" if cov.get("carrier_admitted", True) else "Non-Admitted"
            
            lines.append(f"**{display}**")
            lines.append(f"  Carrier: {carrier} ({admitted})")
            lines.append(f"  Premium: ${premium:,.0f}" if isinstance(premium, (int, float)) else f"  Premium: {premium}")
            
            # Key limits
            limits = cov.get("limits", [])
            if limits:
                for lim in limits[:3]:
                    lines.append(f"  {lim.get('description', '')}: {lim.get('limit', '')}")
            
            # Deductibles
            deds = cov.get("deductibles", [])
            if deds:
                for ded in deds[:2]:
                    lines.append(f"  Deductible: {ded.get('description', '')} — {ded.get('amount', '')}")
            
            # Forms count
            forms = cov.get("forms_endorsements", [])
            if forms:
                lines.append(f"  Forms/Endorsements: {len(forms)} listed")
            
            lines.append("")
    
    lines.append(f"**TOTAL PROPOSED PREMIUM: ${total_premium:,.0f}**")
    
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
    
    await update.message.reply_text(
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
        with open(docx_path, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=docx_filename,
                caption=(
                    f"✅ **Proposal Generated**\n\n"
                    f"**Client:** {session.client_name}\n"
                    f"**File:** {docx_filename}\n\n"
                    f"Review the document and make any final edits as needed.\n"
                    f"Send /proposal [Client Name] to start a new proposal."
                ),
                parse_mode="Markdown"
            )
        
        # Cleanup session
        clear_session(chat_id)
        
        return ConversationHandler.END
        
    except Exception as e:
        logger.error(f"Error generating proposal: {e}", exc_info=True)
        await update.message.reply_text(
            f"❌ Error generating proposal: {str(e)}\n\n"
            f"Try /generate again or /proposal\\_cancel to start over."
        )
        return REVIEWING_EXTRACTION


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
    
    await update.message.reply_text("\n".join(status_lines), parse_mode="Markdown")


def get_proposal_conversation_handler() -> ConversationHandler:
    """Create and return the ConversationHandler for /proposal."""
    return ConversationHandler(
        entry_points=[CommandHandler("proposal", proposal_start)],
        states={
            WAITING_FOR_FILES: [
                MessageHandler(filters.Document.ALL, receive_file),
                CommandHandler("extract", extract_data),
                CommandHandler("generate", generate_doc),  # Auto-extract if needed
                CommandHandler("proposal_cancel", proposal_cancel),
            ],
            REVIEWING_EXTRACTION: [
                CommandHandler("generate", generate_doc),
                CommandHandler("adjust", adjust_data),
                CommandHandler("extract", extract_data),  # Re-extract
                CommandHandler("proposal_cancel", proposal_cancel),
            ],
        },
        fallbacks=[
            CommandHandler("extract", extract_data),
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
