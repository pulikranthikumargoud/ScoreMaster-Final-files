import asyncio
import json
import base64
import time
import html                                     # 🆕 for HTML parse_mode escaping
from datetime import datetime
import logging
import os
import re
import shutil
import sqlite3
import pypdf
try:
    from master_pdf_processor import process_master_pdf   # used by handle_master_pdf; /PT also imports it lazily
except Exception as _mpp_err:                             # a missing/broken processor must not stop the bot starting
    process_master_pdf = None
    print(f"⚠️ master_pdf_processor not loaded at startup: {_mpp_err}")
from urllib.parse import urljoin

import aiosqlite
from bs4 import BeautifulSoup
import httpx
import pandas as pd
from playwright.async_api import async_playwright
from telegram import Update
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    ContextTypes,
    MessageHandler,
    filters,
)
from telegram.request import HTTPXRequest
from telegram.error import RetryAfter, NetworkError, TimedOut, Forbidden, BadRequest

# ==============================================================================
# 1. CONFIGURATION & DIRECTORIES
# ==============================================================================
BOT_TOKEN = os.getenv("BOT_TOKEN", "8859800373:AAF9UDn1hlK8ypRPBwlE3JiNW0ba1V-wQxY")  #ScoreMaster
# 🌟 ADD YOUR IDS HERE (Separated by commas inside the brackets)
ADMIN_IDS = [
    1001653060,  # Kranthikumar
    #23456789,   # Add friend/admin ID here
    #987654321    # Add another admin ID here
]  # 🌟 YOUR TELEGRAM ID (Only you can use /broadcast)

# 🌟 NEW: JSON Backup Setup
USERS_JSON_FILE = "users_backup.json"
user_json_lock = asyncio.Lock()

DB_FILE = os.getenv("DB_FILE", "exam_evaluator.db")
STORAGE_DIR = os.getenv("STORAGE_DIR", "downloaded_sheets")
PDF_STORAGE_DIR = os.getenv("PDF_STORAGE_DIR", "downloaded_pdfs")
SCORECARD_IMG_DIR = os.getenv("SCORECARD_IMG_DIR", "scorecard_images")

# ==============================================================================
# 🌟 MASTER FILES — organized by exam/subject, auto-created, self-healing
# ==============================================================================
MASTER_ROOT_DIR    = os.getenv("MASTER_ROOT_DIR", "master_files")
MASTER_PDF_DIR     = os.path.join(MASTER_ROOT_DIR, "master_pdfs")
MASTER_TESTS_DIR   = os.path.join(MASTER_ROOT_DIR, "master_tests")
MASTER_SHEETS_DIR  = os.path.join(MASTER_ROOT_DIR, "master_sheets")
MASTER_JSON_DIR    = os.path.join(MASTER_ROOT_DIR, "master_json")

TEST_STORAGE_DIR = os.getenv("TEST_STORAGE_DIR", "generated_tests")
os.makedirs(TEST_STORAGE_DIR, exist_ok=True)

# ==============================================================================
# TEST GENERATOR CONFIGURATION
# ==============================================================================
TEST_PASSWORD = "KKGEEE"        # Set a string like "1234" to require a password. Set to None for no password.
                                # eg: TEST_PASSWORD = "My#Super$Secret@Pass_2026!" 
                                # eg: TEST_PASSWORD = "KranthikumarGoudEEE" 
                                # 🌟 MUST HAVE QUOTES .

SEND_MOCK_TEST_TO_USER = True   # True = Sends Mock Test to user. False = Saves locally only.
SAVE_INDIVIDUAL_PDF = False     # True = Saves individual candidate PDFs. False = Saves disk space.
TEST_EXPIRY_DAYS = None         # Set an integer like 7 for 7 days expiry. Set to None for no expiry.

# ==============================================================================
# 🌟 MASTER SAVE FLAGS (Laptop Storage & Smart Cache)
# ==============================================================================
# These flags control whether the Master files are generated and saved to your computer.
# If True, the files are saved to the `master_files/` folder.

# SAVE_MASTER_PDF
# True  = Generates 1 Master PDF per shift. (Skips generation if it already exists).
# False = Does not generate Master PDFs.
SAVE_MASTER_PDF             = True   

# SAVE_MASTER_TEST_HTML (SMART CACHE)
# True  = Saves 1 Master Test HTML per shift. 
#         If a 2nd user submits the same shift, the bot skips downloading 500+ 
#         images and instantly sends the cached file. Saves massive data/time.
# False = NO CACHE. Downloads images and generates a fresh HTML for EVERY user.
SAVE_MASTER_TEST_HTML       = True   

# SAVE_MASTER_RAW_SHEET
# True  = Saves 1 raw response sheet HTML per shift to your hard drive.
# False = Does not save the raw HTML file.
SAVE_MASTER_RAW_SHEET       = True   

# ==============================================================================
# 🌟 SEPARATE TELEGRAM UPLOAD FLAGS (For Master Files)
# ==============================================================================
# These flags control whether the Master files are uploaded to your Telegram Admin Channel.
# Note: The files will still be saved to your laptop if the SAVE_MASTER flags above are True.

# True  = Uploads to your Telegram Admin Channel.
# False = Saves to your computer folder ONLY, skips Telegram upload (saves your internet data).

SEND_MASTER_PDF_TO_ADMIN    = False   # Controls the Master Answer Key PDF
SEND_MASTER_TEST_TO_ADMIN   = True   # Controls the Master Mock Test HTML
SEND_MASTER_SHEET_TO_ADMIN  = False   # Controls the Master Raw Sheet HTML

# ==============================================================================
# 🌟 JSON MASTER FILES — Two independent versions (small + full)
# ==============================================================================

# --- SMALL JSON (~200 KB) — meta + info + test_data (editable via /regen_from_json)
SEND_MASTER_JSON_TO_ADMIN   = True    # 🆕 send raw JSON to admin channel
SAVE_MASTER_JSON            = True    # 🆕 save JSON to laptop
# --- FULL JSON (~3-4 MB) — everything above + raw_source_html (full RRB HTML)
SAVE_MASTER_JSON_FULL          = True    # save full JSON to laptop
SEND_MASTER_JSON_FULL_TO_ADMIN = True    # send full JSON to admin channel


# copying, highlighting, and right-clicking.
COPY_PROTECTION = True  # 🌟 True = Block copy for ALL exams. False = Allow copy for ALL exams.

# 🌟 NEW: Master PDF Watermark Configuration
ADD_WATERMARK_TO_MASTER = True              # Set to False if you don't want a watermark on Master PDFs
WATERMARK_TEXT = "@KranthikumargoudEEE"     # The text you want to appear (e.g., "DIGIALM", "SSC", or your bot name)

# 🌟 ADMIN AUTO-COPY (silent — user is never notified)
SEND_COPY_TO_ADMIN = True                    # master switch: scorecard + mock test copies
SEND_COPY_TO_ADMIN_DMS = False                # 👈 NEW: False = skip ADMIN_IDS DMs (channels still receive)
SEND_PDF_COPY_TO_ADMIN = False               # 👈 NEW: PDF copies only if True (saves data when off)
ADMIN_COPY_CHAT_IDS = [
    "-1004468515115",                        # channel 1 (yours now)
    # "-1001234567890",                     # channel 2 (add later)
    # "-1001234567890:3383",                # channel 3 with topic (add later)
]
LOG_FILE = "bot.log"

EXCEL_FILE = os.getenv("EXCEL_FILE", "exam_results.xlsx")
EXCEL_USER_HISTORY = os.getenv("EXCEL_USER_HISTORY", "user_submissions_history.xlsx")
UNRECOGNIZED_LOG = "unrecognized_submissions.txt"

CHANNEL_LINK = os.getenv("CHANNEL_LINK", "https://t.me/kranthikumargoudEEE")
CHANNEL_LINK1 = os.getenv("CHANNEL_LINK1", "https://whatsapp.com/channel/0029Vam4uhw4CrfcNOn3eV0K")

# Marking schemes: (Positive Marks, Negative Deduction)
# ==============================================================================
# 🌟 HOW TO ADD NEW EXAMS / MARKING SCHEMES:
# ==============================================================================
# 1. Identify the exam name from the "Subject" field in the HTML 
#    (e.g., "BEL Probationary Engineer").
# 2. Add the keyword (e.g., "PROBATIONARY") to EXAM_MARKING_SCHEMES 
#    with its (positive, negative) marks.
# 3. Restart the bot.
# 
# ⚠️ NO URL HARDCODING REQUIRED. The bot dynamically scans the HTML text 
# and Subject field. It will safely say "Competitive Exam" for unknown exams.
# ==============================================================================

EXAM_MARKING_SCHEMES = {
    "SSC JE PAPER 2": (3.0, 1.0),
    "SSC JE": (1.0, 0.25),
    "SSC CGL TIER 2": (3.0, 1.0),
    "SSC CGL": (2.0, 0.50),
    "SSC CHSL TIER 2": (3.0, 1.0),
    "SSC CHSL": (2.0, 0.50),
    "SSC MTS": (3.0, 1.0),
    "SSC CPO": (1.0, 0.25),
    "SSC GD": (2.0, 0.25),
    "SSC STENOGRAPHER": (1.0, 0.25),
    "SSC JHT": (1.0, 0.25),
    "SSC SELECTION POST": (2.0, 0.50),
    "SSC": (2.0, 0.50),
    "COAL INDIA": (1.0, 0.0),
    "VIDYUT": (1.0, 0.0),
    "RRB": (1.0, 0.333333),
    "BARC": (3.0, 1.0),
    
    # 🌟 ADDED BEL HERE (Matches "Probationary Engineer Electrical")
    "PROBATIONARY ENGINEER": (1.0, 0.25), 
    
    "COMPETITIVE EXAM": (1.0, 0.0), 
    "DEFAULT": (1.0, 0.0), 
}

for d in [STORAGE_DIR, PDF_STORAGE_DIR, SCORECARD_IMG_DIR,
          MASTER_ROOT_DIR, MASTER_PDF_DIR, MASTER_TESTS_DIR, MASTER_SHEETS_DIR, MASTER_JSON_DIR]:
    os.makedirs(d, exist_ok=True)

# NEW: Prevents race conditions if 2 users submit the exact same shift at the same second
generating_masters = set()

excel_lock = asyncio.Lock()
history_lock = asyncio.Lock()
db_lock = asyncio.Lock()
# 🌟 Fast scorecards get their own small pool → never wait for slow PDFs
scorecard_semaphore = asyncio.Semaphore(3)   # ~3s per render; safe on 8GB laptop
pdf_semaphore        = asyncio.Semaphore(2)   # heavy renders (10-20s each)
# (old playwright_semaphore kept as alias for backward-compat)
playwright_semaphore = scorecard_semaphore

from logging.handlers import RotatingFileHandler

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
    handlers=[
        logging.StreamHandler(),
        RotatingFileHandler(LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=3, encoding="utf-8"),
    ],
)
logger = logging.getLogger(__name__)
logging.getLogger("httpx").setLevel(logging.WARNING)



# ==============================================================================
# 🆕 ATOMIC FILE WRITE HELPERS (crash-safe)
# ==============================================================================
def atomic_write_bytes(file_path: str, data: bytes):
    """Write bytes atomically — crash-safe, never leaves corrupt target."""
    dir_name = os.path.dirname(file_path) or "."
    os.makedirs(dir_name, exist_ok=True)
    tmp_path = file_path + ".tmp"
    with open(tmp_path, "wb") as f:
        f.write(data)
        f.flush()
        try:
            os.fsync(f.fileno())
        except Exception:
            pass
    os.replace(tmp_path, file_path)


def atomic_write_text(file_path: str, text: str):
    atomic_write_bytes(file_path, text.encode("utf-8"))


# ==============================================================================
# 🆕 IMAGE CACHE — reuse successful downloads across users / shifts
# ==============================================================================
IMAGE_CACHE_DIR = os.path.join("image_cache")
os.makedirs(IMAGE_CACHE_DIR, exist_ok=True)


def _img_cache_key(url: str) -> str:
    """SHA256 of URL → short hex filename."""
    import hashlib
    return hashlib.sha256(url.encode("utf-8")).hexdigest()[:24] + ".txt"


def _img_cache_load(url: str):
    """Return cached data-URI if present, else None."""
    path = os.path.join(IMAGE_CACHE_DIR, _img_cache_key(url))
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = f.read().strip()
                if data.startswith("data:image"):
                    return data
        except Exception:
            pass
    return None


def _img_cache_save(url: str, data_uri: str):
    """Write data-URI to disk cache (atomic)."""
    try:
        path = os.path.join(IMAGE_CACHE_DIR, _img_cache_key(url))
        atomic_write_text(path, data_uri)
    except Exception as e:
        logger.warning("Image cache write failed: %s", e)    




# ==============================================================================
# 🆕 HELPER: NORMALIZE UID (fuzzy matching for GATE keys)
# ==============================================================================
def _normalize_uid(uid: str) -> str:
    """Lowercase + strip every non-alphanumeric char.
    Makes 'EE_ELEC_02-02-2025_2-30-PM' == 'EE_ELEC_02_02_2025_2_30_PM' == 'eeelec02022025230pm'."""
    return re.sub(r"[^a-z0-9]", "", (uid or "").lower())



# ==============================================================================
# 🆕 HELPER: MASTER PATHS
# ==============================================================================
def get_master_paths(info: dict) -> dict:
    def _s(v, fallback):
        s = re.sub(r"[^a-zA-Z0-9_-]", "_", str(v or "")).strip("_")
        return s or fallback

    safe_exam   = _s(info.get("exam_name"), "Unknown_Exam")[:60]
    safe_branch = _s(info.get("branch"), "General")[:60]
    safe_date   = _s(info.get("test_date"), "NA")
    safe_time   = _s(info.get("test_time"), "NA")

    if (safe_branch.lower() == safe_exam.lower() or safe_branch.lower() in ("general", "n_a", "na", "")):
        subject_dir = "_All"
    else:
        subject_dir = safe_branch

    base_name = f"{safe_date}_{safe_time}"
    return {
        "key":            f"{safe_exam}_{safe_date}_{safe_time}_{safe_branch}",
        "pdf_path":       os.path.join(MASTER_PDF_DIR,    safe_exam, subject_dir, base_name + ".pdf"),
        "test_path":      os.path.join(MASTER_TESTS_DIR,  safe_exam, subject_dir, base_name + ".html"),
        "sheet_path":     os.path.join(MASTER_SHEETS_DIR, safe_exam, subject_dir, base_name + ".html"),
        "json_path":      os.path.join(MASTER_JSON_DIR,   safe_exam, subject_dir, base_name + ".json"),
        "json_full_path": os.path.join(MASTER_JSON_DIR,   safe_exam, subject_dir, base_name + "_full.json"),
    }
# ==============================================================================
# 🆕 HELPER: BLOCKED USERS (Auto-Delete)
# ==============================================================================
async def remove_blocked_user(user_id: int) -> str:
    """Deletes blocked users and returns a formatted string for logging."""
    async with aiosqlite.connect(DB_FILE) as db:
        async with db.execute("SELECT full_name, username FROM users WHERE user_id = ?", (user_id,)) as c:
            row = await c.fetchone()
        
        name = row[0] if row else "Unknown"
        uname = row[1] if row else "Unknown"
        
        await db.execute("DELETE FROM users WHERE user_id = ?", (user_id,))
        await db.commit()
    
    logger.info("🗑️ Auto-removed blocked user %s (%s | %s) from database.", user_id, name, uname)
    return f"• ID: <code>{user_id}</code> | Name: {html.escape(name)} | User: {html.escape(uname)}"

# ==============================================================================
# 🆕 HELPER: GENUINE ERROR NOTIFICATION
# ==============================================================================
async def notify_genuine_error(context, failure_type, user_id, user_full_name, username_str, info, source_url, error, user_message=""):
    admin_text = (
        f"🚨 <b>GENUINE ERROR — {html.escape(failure_type)}</b>\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"👤 <b>User:</b> {html.escape(user_full_name)} ({html.escape(username_str)}) | ID: <code>{user_id}</code>\n"
        f"🏛 <b>Exam:</b> {html.escape(str(info.get('exam_name','N/A')))}\n"
        f"🧑 <b>Candidate:</b> {html.escape(str(info.get('name','N/A')))}\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"❌ <b>Failed:</b> {html.escape(failure_type)}\n"
        f"💥 <b>Error:</b> {html.escape(type(error).__name__)}: {html.escape(str(error))[:300]}\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"🔗 <b>Sheet URL:</b>\n<code>{html.escape(source_url)}</code>"
    )
    for target_chat, thread_id in _get_all_copy_targets(exclude_user_id=user_id):
        try:
            kwargs = dict(chat_id=target_chat, text=admin_text, parse_mode="HTML")
            if thread_id is not None:
                kwargs["message_thread_id"] = thread_id
            await context.bot.send_message(**kwargs)
        except Exception:
            # 🆕 Channel is down → queue this alert for the retry loop
            await queue_failed_delivery(
                context, target_chat, "send_message", "(alert)",
                admin_text, "HTML", thread_id,
                user_id, user_full_name, source_url,
            )

    if not user_message:
        user_message = (
            f"⚠️ <b>{html.escape(failure_type)} failed</b>\n\n"
            f"Your sheet was evaluated, but delivery failed due to a network error.\n\n"
            f"🔄 <b>How to retry:</b>\nResend this link after <b>10 minutes</b>:\n"
            f"<code>{html.escape(source_url)}</code>"
        )
    try:
        await context.bot.send_message(chat_id=user_id, text=user_message, parse_mode="HTML", disable_web_page_preview=True)
    except Exception: pass


async def queue_failed_delivery(context, chat_id, action, file_path, caption, parse_mode="HTML", thread_id=None, user_id="N/A", user_name="N/A", source_url="N/A"):
    """Queue a failed delivery with dedup + exponential backoff scheduling."""

    # 1. Dedup: same (chat, file, action) queued within last hour → skip
    async with aiosqlite.connect(DB_FILE) as db:
        async with db.execute(
            "SELECT id FROM pending_deliveries "
            "WHERE chat_id=? AND file_path=? AND action=? "
            "AND created_at > datetime('now','-1 hour') LIMIT 1",
            (str(chat_id), file_path, action),
        ) as c:
            existing = await c.fetchone()
        if existing:
            logger.info("⏭️ Duplicate queue entry skipped: %s", os.path.basename(file_path))
            return

        # 2. Insert with first retry scheduled 2 minutes out
        await db.execute(
            """INSERT INTO pending_deliveries
               (chat_id, action, file_path, caption, parse_mode, thread_id,
                user_id, user_name, source_url, retry_count, next_retry_at)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 0, datetime('now','+2 minutes'))""",
            (str(chat_id), action, file_path, caption, parse_mode, thread_id,
             str(user_id), user_name, source_url),
        )
        # 3. Audit log — permanent record
        await db.execute(
            "INSERT INTO failure_reports (event_type, user_id, user_name, source_url, file_path, error) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            ("queued", str(user_id), user_name, source_url, os.path.basename(file_path), None),
        )
        await db.commit()

    logger.info("📥 Queued failed delivery for %s: %s %s", chat_id, action, file_path)

    # 4. Notify log channels once
    admin_msg = (
        f"⚠️ <b>DELIVERY FAILED & QUEUED</b>\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"👤 <b>User:</b> {html.escape(str(user_name))} (ID: <code>{user_id}</code>)\n"
        f"🔗 <b>Link:</b> <code>{html.escape(str(source_url))}</code>\n"
        f"📂 <b>Action:</b> {action}\n"
        f"📄 <b>File:</b> {os.path.basename(file_path)}\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"<i>First retry in 2 min · backs off to 6h cap.</i>"
    )
    for target_chat, topic_id in _get_all_copy_targets():
        try:
            kwargs = dict(chat_id=target_chat, text=admin_msg, parse_mode="HTML")
            if topic_id is not None:
                kwargs["message_thread_id"] = topic_id
            await context.bot.send_message(**kwargs)
            await asyncio.sleep(0.25)
        except Exception:
            pass  # bot.log still records it


async def _report_recovery(app, user_id, user_name, source_url, file_path):
    """Deliver a 'recovered' notice to log channels + audit table."""
    async with aiosqlite.connect(DB_FILE) as db:
        await db.execute(
            "INSERT INTO failure_reports (event_type, user_id, user_name, source_url, file_path, error) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            ("recovered", str(user_id), user_name, source_url, os.path.basename(file_path), None),
        )
        await db.commit()

    text = (
        f"✅ <b>RECOVERY SUCCESS</b>\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"👤 <b>User:</b> {html.escape(str(user_name))} (ID: <code>{user_id}</code>)\n"
        f"🔗 <b>Link:</b> <code>{html.escape(str(source_url))}</code>\n"
        f"📄 <b>File:</b> {os.path.basename(file_path)}\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"<i>Delivered successfully after network restored.</i>"
    )
    for target_chat, topic_id in _get_all_copy_targets():
        try:
            k = dict(chat_id=target_chat, text=text, parse_mode="HTML")
            if topic_id is not None:
                k["message_thread_id"] = topic_id
            await app.bot.send_message(**k)
            await asyncio.sleep(0.25)
        except Exception:
            pass


async def _report_permanent_failure(app, user_id, user_name, source_url, file_path, error=None):
    """Deliver a 'permanent failure' notice to log channels + audit table."""
    async with aiosqlite.connect(DB_FILE) as db:
        await db.execute(
            "INSERT INTO failure_reports (event_type, user_id, user_name, source_url, file_path, error) "
            "VALUES (?, ?, ?, ?, ?, ?)",
            ("dropped", str(user_id), user_name, source_url, os.path.basename(file_path),
             str(error) if error else "file missing on disk"),
        )
        await db.commit()

    text = (
        f"❌ <b>DELIVERY PERMANENTLY FAILED</b>\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"👤 <b>User:</b> {html.escape(str(user_name))} (ID: <code>{user_id}</code>)\n"
        f"🔗 <b>Link:</b> <code>{html.escape(str(source_url))}</code>\n"
        f"📄 <b>File:</b> {os.path.basename(file_path)}\n"
        f"💥 <b>Reason:</b> {html.escape(str(error or 'file missing on disk'))[:200]}\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"<i>Row removed from queue. Resubmit the link if needed.</i>"
    )
    for target_chat, topic_id in _get_all_copy_targets():
        try:
            k = dict(chat_id=target_chat, text=text, parse_mode="HTML")
            if topic_id is not None:
                k["message_thread_id"] = topic_id
            await app.bot.send_message(**k)
            await asyncio.sleep(0.25)
        except Exception:
            pass


async def retry_failed_deliveries(app):
    """Retry queue with per-row exponential backoff + circuit breaker.

    Tick every 60s. Only pick rows whose next_retry_at has arrived.
    Backoff: 2m → 8m → 32m → 2h → 6h (cap).
    Circuit breaker: after 5 consecutive network failures, pause 5 min.
    """
    _consecutive_failures = 0
    _circuit_open_until = 0

    while True:
        await asyncio.sleep(60)

        # Circuit breaker: if open, skip this tick
        if time.time() < _circuit_open_until:
            continue

        try:
            async with aiosqlite.connect(DB_FILE) as db:
                async with db.execute(
                    """SELECT id, chat_id, action, file_path, caption, parse_mode,
                              thread_id, user_id, user_name, source_url, retry_count
                       FROM pending_deliveries
                       WHERE next_retry_at IS NULL OR next_retry_at <= datetime('now')
                       ORDER BY next_retry_at ASC LIMIT 20"""
                ) as c:
                    rows = await c.fetchall()

            if not rows:
                _consecutive_failures = 0
                continue

            any_success = False

            for row in rows:
                (rid, chat_id_str, action, file_path, caption, parse_mode,
                 thread_id, user_id, user_name, source_url, retry_count) = row

                try:
                    chat_id = int(chat_id_str)
                except (ValueError, TypeError):
                    chat_id = chat_id_str

                # File gone → permanent failure (only for file-based actions)
                # send_message has no file — its "file_path" is a placeholder
                if action in ("send_photo", "send_document") and not os.path.exists(file_path):
                    logger.warning("File missing for retry — dropping: %s", file_path)
                    await _report_permanent_failure(
                        app, user_id, user_name, source_url, file_path,
                        "file missing on disk"
                    )
                    async with aiosqlite.connect(DB_FILE) as db:
                        await db.execute("DELETE FROM pending_deliveries WHERE id=?", (rid,))
                        await db.commit()
                    continue

                try:
                    kwargs = dict(chat_id=chat_id, caption=caption, parse_mode=parse_mode)
                    if thread_id is not None:
                        kwargs["message_thread_id"] = thread_id

                    if action == "send_photo":
                        with open(file_path, "rb") as f:
                            kwargs["photo"] = f
                            await app.bot.send_photo(**kwargs)
                    elif action == "send_document":
                        with open(file_path, "rb") as f:
                            kwargs["document"] = f
                            await app.bot.send_document(**kwargs)
                    elif action == "send_message":
                        kwargs.pop("caption", None)
                        await app.bot.send_message(text=caption, **kwargs)
                    else:
                        raise ValueError(f"Unknown action: {action}")

                    # Success → delete + report
                    async with aiosqlite.connect(DB_FILE) as db:
                        await db.execute("DELETE FROM pending_deliveries WHERE id=?", (rid,))
                        await db.commit()

                    any_success = True
                    _consecutive_failures = 0
                    await _report_recovery(app, user_id, user_name, source_url, file_path)

                except Exception as e:
                    # Failure → exponential backoff
                    _consecutive_failures += 1
                    backoff_seconds = min(3600 * 6, 120 * (4 ** min(retry_count, 6)))
                    # 2m → 8m → 32m → 2h 8m → 6h (capped)

                    async with aiosqlite.connect(DB_FILE) as db:
                        await db.execute(
                            """UPDATE pending_deliveries
                               SET retry_count = retry_count + 1,
                                   last_attempt_at = datetime('now'),
                                   next_retry_at = datetime('now', ?),
                                   last_error = ?
                               WHERE id = ?""",
                            (f"+{backoff_seconds} seconds", str(e)[:200], rid),
                        )
                        await db.commit()

                    logger.warning(
                        "Retry failed (row #%s, attempt %s, next in %ss): %s",
                        rid, retry_count + 1, backoff_seconds, e
                    )

            # Circuit breaker: if everything failed, pause
            if not any_success and _consecutive_failures >= 5:
                _circuit_open_until = time.time() + 300
                logger.warning(
                    "🔴 Circuit OPEN — pausing all retries for 5 min "
                    "(consecutive failures: %s)", _consecutive_failures
                )
                try:
                    for target_chat, topic_id in _get_all_copy_targets():
                        kwargs = dict(
                            chat_id=target_chat,
                            text=(
                                f"🔴 <b>RETRY CIRCUIT OPEN</b>\n"
                                f"━━━━━━━━━━━━━━━━━━━━\n"
                                f"{_consecutive_failures} consecutive network failures.\n"
                                f"All retries paused for 5 minutes.\n\n"
                                f"<i>If DNS is down on the bot host, fix that first.</i>"
                            ),
                            parse_mode="HTML",
                        )
                        if topic_id is not None:
                            kwargs["message_thread_id"] = topic_id
                        await app.bot.send_message(**kwargs)
                except Exception:
                    pass

        except Exception as e:
            logger.error(f"Retry loop error: {e}")


# ==============================================================================
# 🆕 CHANGE 5 — Hourly health digest
# ==============================================================================
async def hourly_health_digest(app):
    """Send one health summary per hour: queue depth, DNS status, 1h stats."""
    while True:
        await asyncio.sleep(3600)
        try:
            # Direct DNS test — cheapest health check available
            dns_ok = True
            try:
                loop = asyncio.get_event_loop()
                await asyncio.wait_for(
                    loop.getaddrinfo("api.telegram.org", 443),
                    timeout=5.0,
                )
            except Exception:
                dns_ok = False

            async with aiosqlite.connect(DB_FILE) as db:
                async with db.execute("SELECT COUNT(*) FROM pending_deliveries") as c:
                    pending = (await c.fetchone())[0]
                async with db.execute(
                    "SELECT event_type, COUNT(*) FROM failure_reports "
                    "WHERE created_at > datetime('now','-1 hour') GROUP BY event_type"
                ) as c:
                    stats = dict(await c.fetchall())

            badge = "🟢 DNS OK" if dns_ok else "🔴 <b>DNS DOWN</b>"
            text = (
                f"📊 <b>Hourly Health — {badge}</b>\n"
                f"━━━━━━━━━━━━━━━━━━━━\n"
                f"⏳ Pending retries: <b>{pending}</b>\n"
                f"✅ Recovered (1h): {stats.get('recovered', 0)}\n"
                f"📥 Queued (1h): {stats.get('queued', 0)}\n"
                f"❌ Dropped (1h): {stats.get('dropped', 0)}\n"
                f"━━━━━━━━━━━━━━━━━━━━\n"
                f"🕐 {datetime.now().strftime('%d/%m/%Y %H:%M')}"
            )

            for target_chat, topic_id in _get_all_copy_targets():
                try:
                    kwargs = dict(chat_id=target_chat, text=text, parse_mode="HTML")
                    if topic_id is not None:
                        kwargs["message_thread_id"] = topic_id
                    await app.bot.send_message(**kwargs)
                except Exception:
                    pass
        except Exception as e:
            logger.error("Health digest error: %s", e)


# ==============================================================================
# 🆕 CHANGE 6 — /queue admin command
# ==============================================================================
async def queue_status_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return await update.message.reply_text("❌ Not authorized.")

    async with aiosqlite.connect(DB_FILE) as db:
        async with db.execute(
            "SELECT id, chat_id, file_path, retry_count, next_retry_at, last_error "
            "FROM pending_deliveries ORDER BY next_retry_at ASC LIMIT 20"
        ) as c:
            rows = await c.fetchall()
        async with db.execute("SELECT COUNT(*) FROM pending_deliveries") as c:
            total = (await c.fetchone())[0]

    if not rows:
        return await update.message.reply_text("✅ Queue is empty.")

    text = f"📥 <b>Pending Deliveries ({total} total, showing first {len(rows)})</b>\n"
    text += "━━━━━━━━━━━━━━━━━━━━\n"
    for rid, chat, path, rc, nxt, err in rows:
        text += (
            f"• <b>#{rid}</b> → <code>{chat}</code>\n"
            f"  📄 <code>{os.path.basename(str(path))[:50]}</code>\n"
            f"  🔁 Retries: {rc} | ⏰ Next: {nxt or 'now'}\n"
            f"  💥 {html.escape(str(err or '-'))[:80]}\n"
        )
    await update.message.reply_text(text, parse_mode="HTML")



def cleanup_orphan_files():
    """Delete .tmp / .tmp_* files from crashed writes. Never touches _FAILED_MANUAL_UPLOADS/."""
    cleaned = 0
    for folder in [MASTER_PDF_DIR, MASTER_TESTS_DIR, MASTER_SHEETS_DIR,
                   PDF_STORAGE_DIR, TEST_STORAGE_DIR, SCORECARD_IMG_DIR,
                   STORAGE_DIR, MASTER_JSON_DIR]:
        if not os.path.isdir(folder):
            continue
        for root, _, files in os.walk(folder):
            # 🛡️ Skip the failed-manual-uploads folder entirely
            if "_FAILED_MANUAL_UPLOADS" in root:
                continue
            for fname in files:
                # Remove both `foo.tmp` (suffix) and `.tmp_foo` (prefix)
                if fname.endswith(".tmp") or fname.startswith(".tmp_"):
                    try:
                        os.remove(os.path.join(root, fname))
                        cleaned += 1
                    except Exception:
                        pass
    if cleaned:
        logger.info("🧹 Cleaned %s orphaned .tmp files", cleaned)


# ==============================================================================
# 🆕 RATE LIMITING
# ==============================================================================
_rate_limit_lock = asyncio.Lock()
_user_submission_times: dict = {}
MAX_SUBMISSIONS_PER_HOUR = 5


async def check_rate_limit(user_id: int) -> bool:
    """True = allowed, False = rate-limited. Admins bypass."""
    if user_id in ADMIN_IDS:
        return True
    now = time.time()
    cutoff = now - 3600
    async with _rate_limit_lock:
        times = [t for t in _user_submission_times.get(user_id, []) if t > cutoff]
        if len(times) >= MAX_SUBMISSIONS_PER_HOUR:
            _user_submission_times[user_id] = times
            return False
        times.append(now)
        _user_submission_times[user_id] = times
        return True


# ==============================================================================
# 🆕 DUPLICATE DETECTION
# ==============================================================================
async def check_recent_submission(user_id: int, hall_ticket: str, source_url: str):
    """Return dict if same user + same hall_ticket + same URL within last 10 min."""
    if user_id in ADMIN_IDS:       # 👈 ADD THIS LINE
        return None                # 👈 ADD THIS LINE
        
    async with aiosqlite.connect(DB_FILE) as db:
        cursor = await db.execute(
            "SELECT total_score, attempted, correct, incorrect, timestamp "
            "FROM candidates WHERE user_id=? AND hall_ticket=? AND response_url=? "
            "ORDER BY timestamp DESC LIMIT 1",
            (user_id, hall_ticket, source_url),
        )
        row = await cursor.fetchone()
    if not row:
        return None
    try:
        submitted_at = datetime.strptime(row[4], "%Y-%m-%d %H:%M:%S")
        age_seconds = (datetime.now() - submitted_at).total_seconds()
    except Exception:
        return None
    if age_seconds > 600:   # 🆕 10 min — covers mock test processing window
        return None
    return {
        "total_score": row[0],
        "attempted": row[1],
        "correct": row[2],
        "incorrect": row[3],
        "age_seconds": age_seconds,
    }


# ==============================================================================
# 🆕 STARTUP NOTIFICATION
# ==============================================================================
async def send_startup_notification(app):
    """Send 'Bot Online' + DB health to all copy targets."""
    try:
        async with aiosqlite.connect(DB_FILE) as db:
            async with db.execute("SELECT COUNT(*) FROM users") as c:
                total_users = (await c.fetchone())[0]
            async with db.execute("SELECT COUNT(*) FROM candidates") as c:
                total_sheets = (await c.fetchone())[0]
            async with db.execute(
                "SELECT COUNT(*) FROM submission_logs WHERE submitted_at >= datetime('now','-1 day')"
            ) as c:
                today = (await c.fetchone())[0]

        text = (
            f"🟢 *Bot Online*\n"
            f"⏰ {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"👥 Total users: {total_users}\n"
            f"📄 Total sheets: {total_sheets}\n"
            f"📅 Submissions (last 24h): {today}\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"✅ Peer cache primed\n"
            f"✅ Copy channels: {len(ADMIN_COPY_CHAT_IDS)}\n"
            f"✅ Admin DMs: {len(ADMIN_IDS)}"
        )

        for chat, topic in _get_all_copy_targets():
            try:
                kwargs = dict(chat_id=chat, text=text, parse_mode="Markdown")
                if topic is not None:
                    kwargs["message_thread_id"] = topic
                await app.bot.send_message(**kwargs)
            except Exception as e:
                logger.warning("Startup notify failed → %s: %s", chat, e)
    except Exception as e:
        logger.warning("Startup notification error: %s", e)


# ==============================================================================
# 🆕 HOURLY BACKUP
# ==============================================================================
async def periodic_backup():
    """Hourly DB backup + rate-limit cleanup."""
    while True:
        try:
            # 👇 NEW: Auto-remove expired admins 👇
            global ADMIN_IDS
            async with aiosqlite.connect(DB_FILE) as db:
                await db.execute("DELETE FROM admins WHERE expires_at IS NOT NULL AND expires_at <= datetime('now')")
                async with db.execute("SELECT user_id FROM admins WHERE expires_at IS NULL OR expires_at > datetime('now')") as c:
                    rows = await c.fetchall()
                    ADMIN_IDS.clear()
                    ADMIN_IDS.extend([row[0] for row in rows])
                await db.commit()
            # 👆 END NEW 👆

            # 🆕 Purge failure_reports older than 30 days (prevents unbounded growth)
            async with aiosqlite.connect(DB_FILE) as db:
                cursor = await db.execute(
                    "DELETE FROM failure_reports WHERE created_at < datetime('now','-30 days')"
                )
                if cursor.rowcount > 0:
                    logger.info("🧹 Purged %s old failure_reports rows", cursor.rowcount)
                await db.commit()

            if os.path.exists(DB_FILE):
                ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                dst = f"backup_hourly_{ts}.db"
                shutil.copyfile(DB_FILE, dst)
                backups = sorted(f for f in os.listdir(".") if f.startswith("backup_hourly_"))
                for old in backups[:-24]:
                    try:
                        os.remove(old)
                    except Exception:
                        pass
                logger.info("✅ Hourly DB backup: %s", dst)

            # Clean rate-limit cache (drop users idle > 2h)
            cutoff = time.time() - 7200
            async with _rate_limit_lock:
                stale = [uid for uid, times in _user_submission_times.items()
                         if not times or times[-1] < cutoff]
                for uid in stale:
                    _user_submission_times.pop(uid, None)
            if stale:
                logger.info("🧹 Rate-limit cache cleaned: %s users", len(stale))
        except Exception as e:
            logger.warning("Hourly backup failed: %s", e)
        await asyncio.sleep(3600)



# ==============================================================================
# 2. DATABASE INITIALIZATION
# ==============================================================================
async def init_db():
    if os.path.exists(DB_FILE):
        backup_db_name = f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
        shutil.copyfile(DB_FILE, backup_db_name)

    async with aiosqlite.connect(DB_FILE) as db:
        await db.execute("PRAGMA journal_mode=WAL")       # 🆕 crash-safe
        await db.execute("PRAGMA synchronous=NORMAL")     # 🆕 safe + fast
        await db.execute("""
            CREATE TABLE IF NOT EXISTS candidates (
                user_id INTEGER,
                exam_name TEXT,
                hall_ticket TEXT PRIMARY KEY,
                candidate_name TEXT,
                branch TEXT,
                total_score REAL,
                total_questions INTEGER,
                attempted INTEGER,
                correct INTEGER,
                incorrect INTEGER,
                test_date TEXT,
                test_time TEXT,
                test_center TEXT,
                response_url TEXT,
                timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
        await db.execute("""
            CREATE TABLE IF NOT EXISTS submission_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                user_full_name TEXT,
                username TEXT,
                exam_name TEXT,
                hall_ticket TEXT,
                candidate_name TEXT,
                response_url TEXT,
                submitted_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)
        
        # 🌟 NEW: Users Table for Broadcast Tracking
        await db.execute("""
            CREATE TABLE IF NOT EXISTS users (
                user_id INTEGER PRIMARY KEY,
                username TEXT,
                full_name TEXT,
                first_seen DATETIME DEFAULT CURRENT_TIMESTAMP,
                last_seen DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # 👇 NEW: Admins Table with Expiry 👇
        await db.execute("""
            CREATE TABLE IF NOT EXISTS admins (
                user_id INTEGER PRIMARY KEY,
                added_by INTEGER,
                added_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                expires_at DATETIME
            )
        """)

        # 👇 NEW: Retry Queue for Network Failures 👇
        await db.execute("""
            CREATE TABLE IF NOT EXISTS pending_deliveries (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                chat_id TEXT,
                action TEXT,
                file_path TEXT,
                caption TEXT,
                parse_mode TEXT,
                thread_id INTEGER,
                user_id TEXT,
                user_name TEXT,
                source_url TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # 🆕 Backoff columns (safe upgrade — only adds if missing)
        cursor = await db.execute("PRAGMA table_info(pending_deliveries)")
        cols = [row[1] for row in await cursor.fetchall()]
        for col, ddl in [
            ("retry_count",     "ALTER TABLE pending_deliveries ADD COLUMN retry_count INTEGER DEFAULT 0"),
            ("last_attempt_at", "ALTER TABLE pending_deliveries ADD COLUMN last_attempt_at DATETIME"),
            ("next_retry_at",   "ALTER TABLE pending_deliveries ADD COLUMN next_retry_at DATETIME"),
            ("last_error",      "ALTER TABLE pending_deliveries ADD COLUMN last_error TEXT"),
        ]:
            if col not in cols:
                await db.execute(ddl)
                logger.info("✅ Added column pending_deliveries.%s", col)

        # 🆕 Audit table — permanent record of every queue event
        await db.execute("""
            CREATE TABLE IF NOT EXISTS failure_reports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                event_type TEXT,
                user_id TEXT,
                user_name TEXT,
                source_url TEXT,
                file_path TEXT,
                error TEXT,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # 👇 SAFE UPDATE: Drop ONLY the gate_keys table (user data is untouched) 👇
        # await db.execute("DROP TABLE IF EXISTS gate_keys") 
        
        await db.execute("""
            CREATE TABLE IF NOT EXISTS gate_keys (
                gate_uid TEXT PRIMARY KEY,
                keys_json TEXT
            )
        """)
        
        # Upgrade table if it already exists without expires_at
        cursor = await db.execute("PRAGMA table_info(admins)")
        columns = [row[1] for row in await cursor.fetchall()]
        if "expires_at" not in columns:
            await db.execute("ALTER TABLE admins ADD COLUMN expires_at DATETIME")

        # Seed the DB with the hardcoded ADMIN_IDS if the table is empty
        async with db.execute("SELECT COUNT(*) FROM admins") as c:
            count = (await c.fetchone())[0]
            if count == 0:
                for aid in ADMIN_IDS:
                    await db.execute("INSERT OR IGNORE INTO admins (user_id, added_by, expires_at) VALUES (?, ?, NULL)", (aid, 0))
        # 👆 NEW 👆

        cursor = await db.execute("PRAGMA table_info(candidates)")
        columns = [row[1] for row in await cursor.fetchall()]
        for col in ["test_date", "test_time", "test_center", "response_url"]:
            if col not in columns:
                await db.execute(f"ALTER TABLE candidates ADD COLUMN {col} TEXT")
        await db.commit()
    logger.info("Database initialized successfully.")


# ==============================================================================
# 3. HTML IMAGE INLINER
# ==============================================================================
async def embed_all_images(html: str, base_url: str = "") -> str:
    soup = BeautifulSoup(html, "html.parser")
    async with httpx.AsyncClient(
        timeout=15.0,
        verify=False,
        limits=httpx.Limits(max_connections=5, max_keepalive_connections=5)
    ) as client:
        for img in soup.find_all("img"):
            src = img.get("src")
            if not src:
                continue
            if base_url and not src.startswith(("http://", "https://", "data:")):
                src = urljoin(base_url, src)
            if src.startswith("data:"):
                continue

            # 🌟 FIX: Normalize (collapse /// in path, keep https:// intact)
            norm_src = re.sub(r'(?<!:)/{2,}', '/', src)

            # 🌟 CACHE: Try disk cache first
            cached = _img_cache_load(norm_src)
            if cached:
                img["src"] = cached
                continue

            success = False
            try:
                resp = await client.get(norm_src, headers={"User-Agent": "Mozilla/5.0", "Referer": base_url})
                if resp.status_code == 200:
                    content_type = resp.headers.get("content-type", "image/png")
                    b64 = base64.b64encode(resp.content).decode("utf-8")
                    data_uri = f"data:{content_type};base64,{b64}"
                    img["src"] = data_uri
                    _img_cache_save(norm_src, data_uri)
                    success = True
            except Exception:
                pass

            # 🌟 UNIVERSAL CURL FALLBACK (for PDFs + scorecard)
            if not success:
                try:
                    import subprocess, tempfile
                    fd, temp_path = tempfile.mkstemp(suffix=".png")
                    os.close(fd)
                    cmd = [
                        "curl", "-L", "-s", "-o", temp_path,
                        "-H", "User-Agent: Mozilla/5.0",
                        "-H", f"Referer: {base_url or 'https://rrb.digialm.com/'}",
                        norm_src
                    ]
                    result = subprocess.run(cmd, capture_output=True, timeout=15)
                    if result.returncode == 0 and os.path.exists(temp_path):
                        with open(temp_path, "rb") as f:
                            content = f.read()
                        is_image = (
                            len(content) > 100 and (
                                content.startswith(b'\x89PNG') or
                                content.startswith(b'\xff\xd8\xff') or
                                content.startswith(b'GIF8')
                            )
                        )
                        if is_image:
                            b64 = base64.b64encode(content).decode("utf-8")
                            data_uri = f"data:image/png;base64,{b64}"
                            img["src"] = data_uri
                            _img_cache_save(norm_src, data_uri)
                            logger.info("✅ curl fallback (PDF) worked: %s", norm_src[-40:])
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                except Exception as e:
                    logger.warning("curl fallback (PDF) failed: %s", e)
    return str(soup)


# ==============================================================================
# 4. ROBUST PDF GENERATOR (PASSAGE EXTRACTION & PAGE BLOAT FIX)
# ==============================================================================
PDF_PRINT_STYLES = """
<style>
    @page { size: A4 portrait; margin: 8mm 10mm 8mm 10mm; }
    *, *::before, *::after { box-sizing: border-box !important; }

    html, body {
        width: 100% !important; margin: 0 !important; padding: 0 !important;
        background: #ffffff !important; color: #111827 !important;
        font-family: Arial, sans-serif !important; font-size: 11px !important;
        line-height: 1.35 !important;
        -webkit-print-color-adjust: exact !important; 
        print-color-adjust: exact !important;
    }

    .print-pnl, .grn-btn, .watermarks, #watermark,
    .watermark-container, #lblWatermark, #qDivE, #qDivH, #lblmenu,
    .language-switch, a[href*="Challenge_Item"], a.challenge-btn,
    input, button, script, style { display: none !important; }

    img { max-width: 100% !important; height: auto !important; }

    .section-title-bar {
        page-break-before: always !important;
        break-before: page !important;
        page-break-after: avoid !important;
        break-after: avoid !important;
        display: block !important;
    }
    
    .main-info-pnl + .section-title-bar {
        page-break-before: auto !important;
        break-before: auto !important;
    }
    
    .question-pnl {
        page-break-inside: avoid !important;
        break-inside: avoid !important;
    }
    .main-info-pnl {
        page-break-inside: avoid !important;
        break-inside: avoid !important;
    }

    .wrapper { width: 100% !important; }
    .main-info-pnl {
        background-color: #fffbd8 !important; border: 1px solid #ffbd52 !important;
        margin: 6px 0 10px 0 !important; padding: 8px !important;
    }
    .main-info-pnl table { border-collapse: collapse !important; width: auto !important; max-width: 100% !important; }
    .main-info-pnl table td { padding: 3px 8px !important; border: 1px solid #bbb !important; font-size: 11px !important; vertical-align: top !important; }
    
    .section-lbl {
        display: block !important; background-color: #99999b !important; color: #ffffff !important;
        padding: 6px 8px !important; font-weight: bold !important; font-size: 12px !important;
        margin: 10px 0 6px 0 !important; page-break-after: avoid !important; break-after: avoid !important;
    }
    .section-lbl span, .section-lbl .bold { color: #ffffff !important; }
    
    .question-pnl {
        background-color: #ffffff !important; border: 1px solid #cccccc !important;
        margin-bottom: 8px !important; padding: 4px 6px !important; overflow: hidden !important;
    }
    
    .questionRowTbl { width: 70% !important; float: left !important; border-collapse: collapse !important; }
    .questionRowTbl td { padding: 3px 2px !important; font-size: 11px !important; vertical-align: top !important; word-break: break-word !important; }
    
    .menu-tbl {
        width: 28% !important; float: right !important; background-color: #eef2f5 !important;
        border: 1px solid #cccccc !important; border-radius: 4px !important;
        padding: 4px 6px !important; font-size: 10px !important;
    }
    .menu-tbl td { padding: 2px 4px !important; font-size: 10px !important; min-width: 0 !important; }
    
    .rw { border-bottom: 1px solid #eeeeee !important; }
    .rightAns, .rightAns * { color: #40c64b !important; -webkit-print-color-adjust: exact !important; }
    .wrngAns, .wrngAns * { color: #f61818 !important; -webkit-print-color-adjust: exact !important; }
    
    .questionRowTbl img.tick, .questionRowTbl img.cross { 
        display: inline-block !important; width: 14px !important; height: 14px !important; 
        vertical-align: middle !important; margin-right: 3px !important; 
    }
    
    .extracted-passage-box {
        display: block !important; width: 100% !important; background: #fafafa !important;
        border: 1px solid #d4d4d4 !important; border-left: 4px solid #1a73e8 !important;
        padding: 8px 12px !important; margin: 8px 0 10px 0 !important; font-size: 11px !important;
    }
    
    /* 🌟 WATERMARK FOR MASTER PDF */
    .master-watermark {
        position: fixed;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%) rotate(-45deg);
        font-size: 80px;
        font-family: Arial, sans-serif;
        font-weight: bold;
        color: rgba(0, 0, 0, 0.08) !important; 
        z-index: 9999;
        pointer-events: none;
        white-space: nowrap;
        -webkit-print-color-adjust: exact;
        print-color-adjust: exact;
    }
</style>
"""

async def save_sheet_as_pdf(html_content: str, pdf_path: str, base_url: str = "", is_master: bool = False, info: dict = None) -> bool:
    """Returns True only if the PDF was actually written to disk."""
    pdf_dir = os.path.dirname(pdf_path)
    try:
        if pdf_dir:
            os.makedirs(pdf_dir, exist_ok=True)

        html_ready = await embed_all_images(html_content, base_url)
        soup = BeautifulSoup(html_ready, "html.parser")

        # 1. Base Cleanup
        for tag in soup(["script", "noscript", "iframe"]): tag.decompose()
        for elem in soup.find_all(id=["lblWatermark", "qDivE", "qDivH", "lblmenu", "watermark"]): elem.decompose()
        for link in soup.find_all("a", href=re.compile(r"Challenge_Item", re.I)): link.decompose()
        for btn in soup.find_all(["input", "button"]): btn.decompose()
        # 🌟 FIX: Do NOT delete print-pnl entirely, or we lose the SSC header.
        for a in soup.find_all("a", string=re.compile(r"print", re.I)): a.decompose()
        for a in soup.find_all("a", onclick=re.compile(r"print", re.I)): a.decompose()

        is_ssc = "SSC ONLINE EXAMINATION" in html_ready or "cbexams" in base_url

        # 2. SSC vs DigiALM layout prep
        if is_ssc:
            sections = soup.find_all(class_="section-cntnr")
            for i, sec in enumerate(sections):
                if i > 0:
                    existing_style = sec.get("style", "")
                    sec["style"] = f"{existing_style} page-break-before: always !important; break-before: page !important;"

            for info_pnl in soup.find_all(class_="main-info-pnl"):
                existing_style = info_pnl.get("style", "")
                info_pnl["style"] = f"{existing_style} page-break-inside: avoid !important; break-inside: avoid !important;"
        else:
            for tag in soup(["style"]): tag.decompose()

            for td in soup.find_all("td"):
                txt = td.get_text()
                if "Comprehension:" in txt or "Read the given passage" in txt:
                    parent_table = td.find_parent("table")
                    if parent_table:
                        passage_div = soup.new_tag("div")
                        passage_div["class"] = "extracted-passage-box"
                        passage_div.extend(td.contents)
                        parent_table.insert_before(passage_div)
                        td.decompose()

            for elem in soup.find_all(["table", "tr", "td", "div"]):
                if elem.has_attr("width"): del elem["width"]
                if elem.has_attr("style"):
                    elem["style"] = re.sub(r'width\s*:\s*[^;]+;?', '', elem["style"], flags=re.I)

        # ======================================================================
        # MASTER PDF LOGIC (Pure Answer Key)
        # ======================================================================
        if is_master:
            if is_ssc:
                # --- SSC MASTER LOGIC ---
                # 1. Remove Legend
                for tbl in soup.find_all("table"):
                    txt = tbl.get_text()
                    if "Correct Option selected" in txt and "Wrong Option selected" in txt:
                        tbl.decompose()
       
                # 2. 🌟 Remove the WHOLE broken candidate table, then rebuild cleanly
                for tbl in soup.find_all("table"):
                    txt = tbl.get_text(" ", strip=True)
                    if ("Roll No" in txt or "Candidate Name" in txt) and \
                       ("Exam Level" in txt or "Test Date" in txt or "Test Time" in txt):
                        tbl.decompose()

                if info:
                    detail_txt = (
                        f'<div class="master-info-block" style="'
                        f'background:#fffbd8; border:1px solid #ffbd52; padding:12px 20px; '
                        f'margin:10px 0 15px 0; border-radius:4px; font-size:12px;">'
                        f'<table style="border-collapse:collapse;">'
                        f'<tr><td style="padding:4px 12px;font-weight:bold;">Exam Level</td><td style="padding:4px 12px;">:</td><td style="padding:4px 12px;">{info.get("branch","N/A")}</td></tr>'
                        f'<tr><td style="padding:4px 12px;font-weight:bold;">Test Date</td><td style="padding:4px 12px;">:</td><td style="padding:4px 12px;">{info.get("test_date","N/A")}</td></tr>'
                        f'<tr><td style="padding:4px 12px;font-weight:bold;">Test Time / Shift</td><td style="padding:4px 12px;">:</td><td style="padding:4px 12px;">{info.get("test_time","N/A")}</td></tr>'
                        f'</table></div>'
                    )
                    details_soup = BeautifulSoup(detail_txt, "html.parser")
                    wrapper = soup.find("div", class_="wrapper") or soup.body
                    if wrapper:
                        first_table = wrapper.find("table")
                        if first_table:
                            first_table.insert_after(details_soup)
                        else:
                            wrapper.insert(0, details_soup)

                # 3. Remove Undertaking and Notes
                for tr in soup.find_all("tr"):
                    txt = tr.get_text()
                    if "Undertaking by the candidate" in txt or "The candidates may note" in txt or "Click Here for PART-" in txt:
                        tr.decompose()

                # 4. Process Question Tables
                for tbl in soup.find_all("table"):
                    qno_count = len(tbl.find_all(string=re.compile(r"Q\.No\s*:\s*\d+")))
                    if qno_count != 1:
                        continue

                    for tr in tbl.find_all("tr"):
                        txt = tr.get_text()
                        if "Hindi version" in txt or "Click Here to Challenge" in txt or "Not Answered" in txt:
                            tr.decompose()
                            continue

                        tds = tr.find_all("td")
                        if len(tds) >= 2:
                            if "Q.No" in tds[0].get_text():
                                for td in tds:
                                    if "bgcolor" in td.attrs: del td["bgcolor"]
                                    td["style"] = "background-color: transparent !important; border: none !important;"
                                continue

                            bg = str(tds[0].get("bgcolor", "")).lower()
                            if "green" in bg or "yellow" in bg:
                                tr["style"] = "background-color: #2e7d32 !important; color: #ffffff !important; font-weight: bold !important; padding: 6px 10px !important; border-radius: 4px !important;"
                                tds[0].decompose()
                                for td in tr.find_all("td"):
                                    if td.has_attr("class"):
                                        td["class"] = [c for c in td["class"] if c.lower() not in ["rightans", "right", "correct"]]
                                    td["style"] = "background-color: transparent !important; border: none !important; padding: 4px !important; color: #ffffff !important;"
                            else:
                                for td in tds:
                                    if "bgcolor" in td.attrs: del td["bgcolor"]
                                    td["style"] = "background-color: transparent !important; border: none !important; padding: 4px !important;"

                    tbl["style"] = "width: 100% !important; background-color: #ffffff !important; border: 1px solid #ddd !important; border-radius: 4px !important; margin-bottom: 15px !important; border-collapse: collapse !important;"

                for q_tbl in soup.find_all("table", class_="questionRowTbl"):
                    q_tbl["style"] = "width: 100% !important; float: none !important; display: block !important; margin: 0 !important;"
                    if q_tbl.has_attr("width"):
                        del q_tbl["width"]

                for rw_td in soup.find_all("td", class_="rw"):
                    rw_td["style"] = "display: block !important; width: 100% !important; padding-bottom: 15px !important;"

            else:
                # --- DIGIALM MASTER LOGIC ---
                for menu in soup.find_all(class_="menu-tbl"):
                    menu.decompose()

                for q_tbl in soup.find_all(class_="questionRowTbl"):
                    q_tbl["style"] = (q_tbl.get("style", "") + " width: 100% !important; float: none !important; display: block !important;").strip()
                    if q_tbl.has_attr("width"):
                        del q_tbl["width"]

                for q_pnl in soup.find_all(class_="question-pnl"):
                    q_pnl["style"] = (q_pnl.get("style", "") + " overflow: visible !important;").strip()

                for tr in soup.find_all("tr"):
                    txt = tr.get_text()
                    if re.search(r'Status\s*:', txt, re.I) or re.search(r'Chosen Option\s*:', txt, re.I):
                        tr.decompose()

                for tbl in soup.find_all("table"):
                    txt = tbl.get_text()
                    if "Correct Option selected" in txt and "Wrong Option selected" in txt:
                        tbl.decompose()
                for div in soup.find_all("div"):
                    txt = div.get_text()
                    if "Correct Option selected" in txt and "Wrong Option selected" in txt:
                        div.decompose()

                for img in soup.find_all("img"):
                    src = img.get("src", "").lower()
                    if any(k in src for k in ["tick", "cross", "correct", "wrong", "check"]):
                        img.decompose()

                for td in soup.find_all("td"):
                    bg = str(td.get("bgcolor", "")).lower()
                    st = str(td.get("style", "")).lower()
                    cls = str(td.get("class", [])).lower()

                    is_correct = ("green" in bg or "green" in st or "yellow" in bg or "yellow" in st or "rightans" in cls)

                    if is_correct:
                        if td.has_attr("class"):
                            td["class"] = [c for c in td["class"] if c.lower() not in ["rightans", "right", "correct"]]
                        if "bgcolor" in td.attrs: td["bgcolor"] = "#2e7d32"
                        td["style"] = "background-color: #2e7d32 !important; color: #ffffff !important; font-weight: bold !important; padding: 6px 10px !important; border-radius: 4px !important;"
                    else:
                        if "bgcolor" in td.attrs: del td["bgcolor"]
                        if "style" in td.attrs:
                            td["style"] = re.sub(r'background-color\s*:\s*[^;]+;?', '', td["style"], flags=re.I)
                            td["style"] = re.sub(r'color\s*:\s*(red|green|yellow|gray|#f61818|#40c64b)[^;]*;?', '', td["style"], flags=re.I)
                        td["style"] = (td.get("style", "") + " color: #111827 !important; background-color: transparent !important;").strip()

                    for child in td.find_all(True):
                        if child.has_attr("style"):
                            child["style"] = re.sub(r'background-color\s*:\s*[^;]+;?', '', child["style"], flags=re.I)
                            child["style"] = re.sub(r'color\s*:\s*(red|green|yellow|gray|#f61818|#40c64b)[^;]*;?', '', child["style"], flags=re.I)
                        if child.has_attr("class"):
                            child["class"] = [c for c in child["class"] if c not in ["rightAns", "wrngAns", "right", "wrong"]]
        # ======================================================================

        # 🌟 Inject Watermark if it's a Master PDF
        if is_master and ADD_WATERMARK_TO_MASTER:
            watermark_div = soup.new_tag("div")
            watermark_div["class"] = "master-watermark"
            watermark_div.string = WATERMARK_TEXT
            if soup.body:
                soup.body.insert(0, watermark_div)

        if soup.head:
            soup.head.append(BeautifulSoup(PDF_PRINT_STYLES, "html.parser"))
        else:
            soup.insert(0, BeautifulSoup(PDF_PRINT_STYLES, "html.parser"))

        final_html = str(soup)

        async with pdf_semaphore:
            async with async_playwright() as p:
                browser = await p.chromium.launch(headless=True, args=["--disable-gpu", "--disable-dev-shm-usage", "--no-sandbox"])
                page = await browser.new_page()
                await page.set_content(final_html, wait_until="load", timeout=60000)
                await page.emulate_media(media="print")

                pdf_bytes = await page.pdf(
                    format="A4", print_background=True, prefer_css_page_size=False,
                    margin={"top": "8mm", "bottom": "8mm", "left": "10mm", "right": "10mm"},
                )
                await browser.close()

        # ---- write with directory guarantee + retry on FileNotFound/Permission ----
        saved = False
        for attempt in range(3):
            try:
                if pdf_dir:
                    os.makedirs(pdf_dir, exist_ok=True)
                atomic_write_bytes(pdf_path, pdf_bytes)
                saved = True
                logger.info("✅ PDF saved: %s", pdf_path)
                break
            except (PermissionError, FileNotFoundError) as e:
                logger.warning("PDF write attempt %s failed (%s): %s", attempt + 1, type(e).__name__, e)
                await asyncio.sleep(0.5)

        if not saved:
            if pdf_dir:
                os.makedirs(pdf_dir, exist_ok=True)
            dir_name, file_name = os.path.split(pdf_path)
            name_base, ext = os.path.splitext(file_name)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            snapshot_path = os.path.join(dir_name, f"{name_base}_snapshot_{ts}{ext}")
            try:
                atomic_write_bytes(snapshot_path, pdf_bytes)
                logger.warning("⚠️ File locked. Saved PDF snapshot to: '%s'", snapshot_path)
                return True
            except Exception as e:
                logger.error("❌ Snapshot fallback also failed: %s", e)
                return False

        return True

    except Exception as e:
        logger.error("❌ PDF generation failed for %s: %s", pdf_path, e, exc_info=True)
        return False


# ==============================================================================
# 5. UNIVERSAL EXAM HEADER & NAME DETECTOR
# ==============================================================================
def extract_header_and_exam_name(soup: BeautifulSoup, source_url: str = ""):
    main_info = soup.find("div", class_=re.compile(r"main-info-pnl", re.I)) or soup.find("table")
    header_html = None
    exam_name_text = None

    PHOTO_WORDS  = ("photograph", "photo", "signature", "thumb", "applicant")
    # NOTE: "logo" removed — BEL uses iONNewLogo.jpg as its banner
    # Small icons are still caught by "icon", "tick", "cross", "spacer", etc.
    BAD_SRC_KEYS = ("tick", "cross", "wrong", "correct", "blank", "dot",
                    "icon", "close_icon", "spacer", "pixel", "1x1")

    if main_info:
        valid_imgs = []
        filtered_count = 0

        for img in main_info.find_all("img"):
            # 1. Normalize src
            src = (img.get("src") or "").strip().strip("'\"").strip()
            src_lower = (src + " " + (img.get("alt") or "")).lower()

            # 2. Base64 → ALWAYS a photo, never a banner
            if src_lower.startswith("data:") or "base64," in src_lower:
                filtered_count += 1
                continue

            # 3. Row text check (catches "Exam Day Photograph", "Application Photograph")
            row = img.find_parent("tr")
            if row and any(w in row.get_text(" ", strip=True).lower() for w in PHOTO_WORDS):
                filtered_count += 1
                continue

            # 4. Ancestor class/id check
            bad_anc = False
            for anc in img.parents:
                if not hasattr(anc, "get"):
                    continue
                anc_str = " ".join([
                    " ".join(anc.get("class") or []),
                    anc.get("id") or "",
                    anc.get("name") or ""
                ]).lower()
                if any(w in anc_str for w in PHOTO_WORDS):
                    bad_anc = True
                    break
            if bad_anc:
                filtered_count += 1
                continue

            # 5. Inside ANY table → info-cell content, not a header banner
            if img.find_parent("table") is not None:
                filtered_count += 1
                continue

            # 6. Keyword blacklist
            if any(k in src_lower for k in BAD_SRC_KEYS):
                filtered_count += 1
                continue

            # 7. Width sanity (photos are 200px, banners are 1000+)
            w_attr = (img.get("width") or "").strip()
            if w_attr.isdigit() and int(w_attr) < 400:
                filtered_count += 1
                continue

            valid_imgs.append(img)

        logger.info("🖼️ Header filter: kept %d banners, filtered %d non-banner images",
                    len(valid_imgs), filtered_count)

        if valid_imgs:
            header_html = ('<div style="display:flex;justify-content:center;'
                           'align-items:center;flex-wrap:wrap;gap:10px;'
                           'width:100%;margin-bottom:10px;">')
            for img in valid_imgs:
                src = (img.get("src") or "").strip()
                header_html += (
                    f'<img src="{src}" style="max-width:100%;max-height:90px;'
                    f'height:auto;object-fit:contain;border-radius:4px;" />'
                )
            header_html += '</div>'

    # 2. EXTRACT "SUBJECT" FROM TABLE (Works perfectly for BEL, RRB, Vidyut)
    extracted_subject = None
    for tbl in soup.find_all("table"):
        for row in tbl.find_all("tr"):
            cells = [td.get_text(strip=True) for td in row.find_all(["td", "th"])]
            for i in range(0, len(cells) - 1, 2):
                key = cells[i].lower()
                if "subject" in key or "discipline" in key:
                    val = cells[i+1].strip()
                    if val and len(val) > 3:
                        extracted_subject = val
                        break
            if extracted_subject: break
        if extracted_subject: break

    # 3. SCAN TEXT DYNAMICALLY (Uses your dictionary as the source of truth)
    known_keywords = [k for k in EXAM_MARKING_SCHEMES.keys() if k not in ["DEFAULT", "COMPETITIVE EXAM"]]

    if main_info:
        raw_banner_text = main_info.get_text("\n", strip=True)
        lines = [re.sub(r"\s+", " ", l).strip() for l in raw_banner_text.splitlines() if len(l.strip()) > 3]

        for line in lines:
            if any(skip in line.lower() for skip in ["participant", "roll", "hall ticket", "print", "test center", "correct answer", "incorrect answer", "note", "options shown", "chosen option"]):
                continue
            
            # Look for ANY keyword you put in EXAM_MARKING_SCHEMES
            if any(k in line.upper() for k in known_keywords):
                exam_name_text = line
                break

    # 4. Prioritize: Known Text > Subject > Title > "Competitive Exam"
    if not exam_name_text:
        if extracted_subject:
            exam_name_text = extracted_subject
        else:
            title = soup.find("title")
            if title and len(title.get_text(strip=True)) > 4:
                exam_name_text = re.sub(r"^(?:DigiALM\s*[-:]|TCS\s*iON\s*[-:])\s*", "", title.get_text(strip=True), flags=re.I)[:60]
            else:
                exam_name_text = "Competitive Exam"

    if not header_html:
        header_html = f'<div class="text-exam-title">{exam_name_text}</div>'

    return header_html, exam_name_text

# ==============================================================================
# 6. SCORECARD IMAGE RENDERER (CLEAN OFFICIAL DESIGN)
# ==============================================================================
async def generate_scorecard_image(
    candidate_info: dict,
    section_stats: dict,
    total_score: float,
    max_marks: float,
    total_questions: int,
    attempted: int,
    correct: int,
    wrong: int,
    header_html: str,
    base_url: str = "",
) -> str:
    soup_header = BeautifulSoup(header_html, "html.parser")

    # 🚫 HARD BLOCK: no base64 / photo-looking image can enter the banner
    for bad in soup_header.find_all("img"):
        s = (bad.get("src") or "").strip().strip("'\"").lower()
        if s.startswith("data:") or "base64," in s or any(
            k in s for k in ("photo", "signature", "thumb", "tick", "cross")
        ):
            bad.decompose()

    async with httpx.AsyncClient(timeout=12.0) as client:
        for img in soup_header.find_all("img"):
            src = img.get("src")
            if not src:
                continue
            if base_url and not src.startswith(("http://", "https://", "data:")):
                src = urljoin(base_url, src)
            if src.startswith("data:"):
                continue
            try:
                resp = await client.get(src, headers={"User-Agent": "Mozilla/5.0"})
                if resp.status_code == 200:
                    content_type = resp.headers.get("content-type", "image/png")
                    b64 = base64.b64encode(resp.content).decode("utf-8")
                    img["src"] = f"data:{content_type};base64,{b64}"
            except Exception:
                pass
    clean_header_html = str(soup_header)

    sec_rows = "".join(
        f"""
        <tr>
            <td class="sec-name">{name}</td>
            <td>{s['total']}</td>
            <td>{s['total'] - s['attempted']}</td>
            <td>{s['correct']}</td>
            <td>{s['wrong']}</td>
            <td class="sec-marks">{s['score']:.2f}</td>
        </tr>
        """
        for name, s in section_stats.items()
    )

    html_template = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="UTF-8">
    <style>
        * {{ box-sizing: border-box; }}
        body {{
            margin: 0;
            padding: 24px;
            background: #eef2f7;
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
            display: flex;
            justify-content: center;
            -webkit-font-smoothing: antialiased;
        }}
        .scorecard-box {{
            width: 780px;
            background: #ffffff;
            border-radius: 16px;
            box-shadow: 0 4px 20px rgba(0,0,0,0.06);
            border: 1px solid #dbe3eb;
            overflow: hidden;
        }}
        .banner-container {{
            width: 100%;
            background: #ffffff;
            padding: 18px 20px 14px 20px;
            text-align: center;
            border-bottom: 2px solid #eef2f7;
            display: flex;
            justify-content: center;
            align-items: center;
            flex-direction: column;
        }}
        .banner-container img {{
            max-width: 95%;
            height: auto;
            image-rendering: -webkit-optimize-contrast;
            margin: 0 auto;
            display: inline-block;
        }}
        .text-exam-title {{
            font-size: 22px;
            font-weight: 700;
            color: #1a202c;
            padding: 6px 0;
            text-align: center;
            letter-spacing: 0.3px;
        }}
        .profile-table {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13.5px;
            background: #fafbfc;
            border-bottom: 2px solid #eef2f7;
        }}
        .profile-table td {{
            padding: 9px 20px;
            border-bottom: 1px solid #edf2f7;
        }}
        .profile-table .lbl {{
            font-weight: 600;
            color: #4a5568;
            width: 25%;
        }}
        .profile-table .val {{
            font-weight: 500;
            color: #1a202c;
            width: 25%;
        }}
        .score-pill {{
            margin: 16px 20px;
            background: #f0f7ff;
            border: 1px solid #cce3ff;
            border-radius: 12px;
            padding: 14px 20px;
            display: flex;
            justify-content: space-around;
            align-items: center;
            text-align: center;
        }}
        .score-pill .item .big {{
            font-size: 26px;
            font-weight: 800;
            color: #0066cc;
        }}
        .score-pill .item .small {{
            font-size: 12px;
            font-weight: 600;
            color: #64748b;
            text-transform: uppercase;
            margin-top: 2px;
        }}
        .table-wrap {{
            padding: 0 20px 20px 20px;
        }}
        table.section-data {{
            width: 100%;
            border-collapse: collapse;
            font-size: 13.5px;
            text-align: center;
        }}
        table.section-data th {{
            background: #2b6cb0;
            color: #ffffff;
            padding: 10px 8px;
            font-weight: 600;
            font-size: 13px;
        }}
        table.section-data td {{
            padding: 9px 10px;
            border-bottom: 1px solid #e2e8f0;
            color: #2d3748;
        }}
        table.section-data td.sec-name {{
            text-align: left;
            font-weight: 500;
        }}
        table.section-data td.sec-marks {{
            font-weight: 700;
            color: #2b6cb0;
        }}
        tr.total-row {{
            background: #fefcbf;
            font-weight: 700;
        }}
        tr.total-row td {{
            border-top: 2px solid #ecc94b;
            color: #744210;
        }}
        .footer {{
            text-align: center;
            font-size: 11px;
            color: #a0aec0;
            padding: 10px;
            background: #f7fafc;
            border-top: 1px solid #edf2f7;
        }}
    </style>
    </head>
    <body>
    <div class="scorecard-box" id="capture">
        <div class="banner-container">
            {clean_header_html}
        </div>
        <table class="profile-table">
            <tr>
                <td class="lbl">Participant Name</td><td class="val">{candidate_info['name']}</td>
                <td class="lbl">Hall Ticket Number</td><td class="val">{candidate_info['hall_ticket']}</td>
            </tr>
            <tr>
                <td class="lbl">Subject / Discipline</td><td class="val">{candidate_info['branch']}</td>
                <td class="lbl">Test Center</td><td class="val">{candidate_info.get('test_center', 'N/A')}</td>
            </tr>
            <tr>
                <td class="lbl">Test Date</td><td class="val">{candidate_info.get('test_date', 'N/A')}</td>
                <td class="lbl">Test Time</td><td class="val">{candidate_info.get('test_time', 'N/A')}</td>
            </tr>
        </table>

        <div class="score-pill">
            <div class="item">
                <div class="big">{total_score:.2f} / {max_marks:.2f}</div>
                <div class="small">Total Marks</div>
            </div>
            <div class="item">
                <div class="big">{(correct / attempted * 100) if attempted else 0:.1f}%</div>
                <div class="small">Accuracy</div>
            </div>
            <div class="item">
                <div class="big">{attempted} / {total_questions}</div>
                <div class="small">Attempted</div>
            </div>
        </div>

        <div class="table-wrap">
            <table class="section-data">
                <thead>
                    <tr>
                        <th style="text-align: left; width: 44%;">Section</th>
                        <th>Total</th>
                        <th>NA</th>
                        <th>Right</th>
                        <th>Wrong</th>
                        <th>Marks</th>
                    </tr>
                </thead>
                <tbody>
                    {sec_rows}
                    <tr class="total-row">
                        <td style="text-align: left;">Total</td>
                        <td>{total_questions}</td>
                        <td>{total_questions - attempted}</td>
                        <td>{correct}</td>
                        <td>{wrong}</td>
                        <td>{total_score:.2f}</td>
                    </tr>
                </tbody>
            </table>
        </div>
        <div class="footer">
            Generated by @KranthikumargoudEEE &bull; {datetime.now().strftime('%d/%m/%Y %H:%M')}
        </div>
    </div>
    </body>
    </html>
    """

    img_path = os.path.join(SCORECARD_IMG_DIR, f"{candidate_info['hall_ticket']}.png")
    async with scorecard_semaphore:
        try:
            async with async_playwright() as p:
                browser = await p.chromium.launch(headless=True, args=["--no-sandbox"])
                page = await browser.new_page(viewport={"width": 840, "height": 1100}, device_scale_factor=2)
                await page.set_content(html_template, wait_until="networkidle")
                elem = await page.query_selector("#capture")
                if elem:
                    await elem.screenshot(path=img_path)
                await browser.close()
                return img_path
        except Exception as e:
            logger.error(f"Scorecard generation error: {e}")
            return ""


# ==============================================================================
# 7. RESPONSE DATA EXTRACTION
# ==============================================================================
def parse_digialm_html(html_content: str, source_url: str = ""):
    soup = BeautifulSoup(html_content, "html.parser")
    letter_to_num = {"A": "1", "B": "2", "C": "3", "D": "4"}

    header_html, exam_name = extract_header_and_exam_name(soup, source_url)

    candidate_info = {
        "exam_name": exam_name,
        "header_html": header_html,
        "name": "Candidate",
        "hall_ticket": "N/A",
        "branch": "General",
        "test_date": "N/A",
        "test_time": "N/A",
        "test_center": "N/A",
    }

    for tbl in soup.find_all("table"):
        for row in tbl.find_all("tr"):
            cells = [td.get_text(strip=True) for td in row.find_all(["td", "th"])]
            for i in range(0, len(cells) - 1, 2):
                key = cells[i].lower()
                val = cells[i + 1].strip()
                if not val or len(val) > 75 or val.startswith("Q."):
                    continue
                if any(x in key for x in ["participant name", "candidate name"]):
                    candidate_info["name"] = val
                elif any(x in key for x in ["hall ticket", "roll number", "participant id", "candidate id"]):
                    candidate_info["hall_ticket"] = val
                elif any(x in key for x in ["subject", "discipline", "branch"]):
                    candidate_info["branch"] = val
                elif "test date" in key or "exam date" in key:
                    candidate_info["test_date"] = val
                elif "test time" in key or "exam time" in key:
                    candidate_info["test_time"] = val
                elif any(x in key for x in ["test center", "test centre"]):
                    candidate_info["test_center"] = val

    questions = []
    seen_qids = set()
    sections = soup.find_all(class_=re.compile(r"section-lbl|section-name|section_title", re.I))

    if sections:
        for sec in sections:
            sec_raw = sec.get_text(strip=True)
            sec_title = re.sub(r"^(?:Section\s*[:\-]?|Part\s*[:\-]?)", "", sec_raw, flags=re.I).strip()
            parent = sec.find_parent(class_=re.compile(r"grp-cntnr|section|main-table", re.I)) or soup
            panels = parent.find_all(class_=re.compile(r"question-pnl|menu-tbl|questionRow|rw", re.I))

            for pnl in panels:
                txt = pnl.get_text(" ", strip=True)
                qid_m = re.search(r"Question\s*ID\s*[:\-]?\s*(\d+)", txt, re.I)
                if not qid_m or qid_m.group(1).strip() in seen_qids:
                    continue
                qid = qid_m.group(1).strip()
                seen_qids.add(qid)

                chosen = "--"
                
                # 🌟 ULTIMATE UN-SHUFFLE: Works for GATE, BEL, and RRB (Text + Image options)
                option_map = {}
                if "Question Type : NAT" not in txt and "Question Type :NAT" not in txt:
                    for tr in pnl.find_all("tr"):
                        tds = tr.find_all("td")
                        if len(tds) >= 2:
                            display_letter = None
                            # Check first column (BEL style) or start of second column (GATE style)
                            m0 = re.match(r"^([A-D])\.?$", tds[0].get_text(strip=True))
                            m1 = re.match(r"^([A-D])\.", tds[1].get_text(strip=True))
                            
                            if m0:
                                display_letter = m0.group(1)
                            elif m1:
                                display_letter = m1.group(1)
                            
                            if display_letter:
                                img = tds[1].find("img")
                                true_letter = display_letter
                                if img:
                                    img_name = img.get("name", "") or img.get("src", "")
                                    suffix_match = re.search(r'([a-d])(?:_\w+)?\.(?:jpe?g|png|gif|bmp)$', img_name, re.I)
                                    if suffix_match:
                                        true_letter = suffix_match.group(1).upper()
                                option_map[display_letter] = letter_to_num.get(true_letter, true_letter)

                # 🌟 THE ULTIMATE REGEX: Captures A-D (MCQ), A,D (MSQ), and 1, 2, 142.85 (RRB/NAT)
                chosen_m = re.search(
                    r"(?:Chosen\s*Option|Given\s*Option|Option\s*Marked|Given\s*Answer)\s*[:\-]?\s*"
                    r"([A-D](?:\s*[,;]\s*[A-D])*|[\d\.\-]+|--)",
                    txt, re.I
                )
                
                if chosen_m:
                    v = chosen_m.group(1).strip().upper()
                    if "," in v or ";" in v:
                        parts = re.split(r"[,;]", v)
                        chosen = ",".join(sorted(
                            option_map.get(p.strip(), letter_to_num.get(p.strip(), p.strip()))
                            for p in parts if p.strip()
                        ))
                    else:
                        chosen = option_map.get(v, letter_to_num.get(v, v))

                correct_opt = None
                right_elem = pnl.find(class_=re.compile(r"rightAns|correct|right_ans", re.I))
                if right_elem:
                    m = re.search(r"^([A-D1-4])\.", right_elem.get_text(strip=True), re.I)
                    if m:
                        correct_opt = letter_to_num.get(m.group(1).upper(), m.group(1).upper())

                if not correct_opt:
                    for idx, row in enumerate(pnl.find_all("tr"), start=1):
                        img = row.find("img")
                        if img and any(k in (img.get("src", "") + img.get("alt", "")).lower() for k in ["tick", "correct", "right", "check"]):
                            m = re.search(r"^([A-D1-4])\.", row.get_text(strip=True), re.I)
                            correct_opt = letter_to_num.get(m.group(1).upper(), str(idx)) if m else str(idx)
                            break

                # 🌟 ULTIMATE FIX: Extract true Master Q.No from the question image filename
                # Prefer images with a 'name' attribute; skip data: URIs (base64 formulas
                # can contain random 'q<digit>' substrings that cause wrong matches).
                master_qno = None
                for img in pnl.find_all("img"):
                    name_attr = (img.get("name") or "").strip()
                    src_attr  = (img.get("src")  or "").strip()

                    # Skip base64 data URIs entirely
                    if src_attr.startswith("data:"):
                        continue

                    candidate = name_attr or src_attr
                    if not candidate:
                        continue

                    # Look for 'q' followed by digits — catches ga4q7.jpg (7), eeq15.jpg (15)
                    m_qno = re.search(r'q(\d+)', candidate, re.I)
                    if m_qno:
                        master_qno = m_qno.group(1)
                        break
                
                if not master_qno:
                    m_qno_text = re.search(r"\b(?:Question\s*No\.?|Q\.?)\s*:?\s*(\d+)\b", txt, re.I)
                    if m_qno_text:
                        master_qno = m_qno_text.group(1)
                
                # Fallback if no image is found (e.g., text-only questions)
                if not master_qno:
                    m_qno_text = re.search(r"\b(?:Question\s*No\.?|Q\.?)\s*:?\s*(\d+)\b", txt, re.I)
                    if m_qno_text:
                        master_qno = m_qno_text.group(1)
                # logger.info(f"🔍 MAPPING: HTML Q#{len(questions)+1} -> PDF Q#{master_qno} (QID: {qid})")

                q_type = "MCQ"
                if "Question Type : NAT" in txt or "Question Type :NAT" in txt:
                    q_type = "NAT"
                elif "Question Type : MSQ" in txt or "Question Type :MSQ" in txt:
                    q_type = "MSQ"

                has_full_marks_note = "Full Marks is being awarded" in txt or "Full Marks is being awarded to all candidates" in txt
                ambiguity_opts = []
                amb_match = re.search(r"options are chosen\s*-\s*([\d,]+)", txt, re.I)
                if amb_match:
                    ambiguity_opts = [opt.strip() for opt in amb_match.group(1).split(",")]

                questions.append({
                    "qid": qid,
                    "master_qno": master_qno,
                    "q_type": q_type,
                    "chosen": chosen,
                    "correct": correct_opt,
                    "section": sec_title or "General",
                    "panel_text": txt,
                    "full_marks_note": has_full_marks_note,
                    "ambiguity_opts": ambiguity_opts
                })

    return candidate_info, questions

# ==============================================================================
# 7.5 NEW SSC CBEXAMS PARSER
# ==============================================================================
def parse_ssc_cbexams(html_pages: list):
    all_questions = []
    candidate_info = {
        "exam_name": "Combined Graduate Level Examination",
        "header_html": "",
        "name": "Candidate",
        "hall_ticket": "N/A",
        "branch": "General",
        "test_date": "N/A",
        "test_time": "N/A",
        "test_center": "N/A",
    }
    seen_qids = set()

    for html in html_pages:
        soup = BeautifulSoup(html, "html.parser")

        # 1. Candidate Info Extraction
        if candidate_info["hall_ticket"] == "N/A":
            ddl = soup.find("select", id="ddltest")
            if ddl:
                opt = ddl.find("option", selected=True)
                if opt and opt.get_text(strip=True):
                    candidate_info["exam_name"] = opt.get_text(strip=True)

            for tr in soup.find_all("tr"):
                tds = tr.find_all(["td", "th"])
                if len(tds) >= 2:
                    k = re.sub(r"\s+", " ", tds[0].get_text(strip=True)).lower()
                    
                    # 🌟 FIX: Handle 3-column tables (Label | : | Value)
                    if len(tds) >= 3 and tds[1].get_text(strip=True) in [":", ""]:
                        v = re.sub(r"\s+", " ", tds[2].get_text(strip=True)).lstrip(":").strip()
                    else:
                        v = re.sub(r"\s+", " ", tds[1].get_text(strip=True)).lstrip(":").strip()

                    if "candidate name" in k:
                        candidate_info["name"] = v
                    elif "roll no" in k:
                        candidate_info["hall_ticket"] = v
                    elif "test date" in k:
                        candidate_info["test_date"] = v
                    elif "test time" in k:
                        candidate_info["test_time"] = v
                    elif "centre" in k or "center" in k:
                        candidate_info["test_center"] = v
                    # 🌟 Extract Exam Level as the Subject
                    elif "exam level" in k or "subject" in k:
                        if v:
                            candidate_info["branch"] = v

        # 2. Section Extraction
        current_section = "General"
        lbl_subj = soup.find("span", id="lblsubject")
        if lbl_subj and lbl_subj.get_text(strip=True):
            current_section = lbl_subj.get_text(strip=True)

        # 3. Question Extraction
        for tbl in soup.find_all("table"):
            # Ensure this is a SINGLE question table, not the main container table with the legend
            qno_cells = tbl.find_all(string=re.compile(r"Q\.No\s*:\s*\d+", re.I))
            if not qno_cells or len(qno_cells) > 1:
                continue
            
            q_match = re.search(r"Q\.No\s*:\s*(\d+)", qno_cells[0], re.I)
            if not q_match:
                continue
            qid = q_match.group(1).strip()
            if qid in seen_qids:
                continue
            seen_qids.add(qid)

            is_unanswered = bool(tbl.find(string=re.compile(r"Not Answered", re.I)))
            chosen_opt = "--"
            correct_opt = None
            opt_counter = 0
            has_green_or_red = False

            for tr in tbl.find_all("tr"):
                tds = tr.find_all("td")
                if len(tds) < 2:
                    continue
                
                # CRITICAL FIX: Skip the row that contains "Q.No" and the question image!
                if "Q.No" in tds[0].get_text():
                    continue

                if tds[1].find("img"):
                    opt_counter += 1
                    opt_str = str(opt_counter)

                    status_td = tds[0]
                    color = (status_td.get("bgcolor") or status_td.get("style") or "").lower()

                    if "green" in color:
                        chosen_opt = opt_str
                        correct_opt = opt_str
                        has_green_or_red = True
                    elif "red" in color:
                        chosen_opt = opt_str
                        has_green_or_red = True
                    elif "yellow" in color:
                        correct_opt = opt_str

            if is_unanswered and not has_green_or_red:
                chosen_opt = "--"

            all_questions.append({
                "qid": qid,
                "chosen": chosen_opt,
                "correct": correct_opt,
                "section": current_section,
                "panel_text": tbl.get_text(" ", strip=True),
            })

    # 🌟 FORCE SUBJECT TO BE EXAM NAME IF IT'S STILL "GENERAL"
    if candidate_info["branch"] == "General" or not candidate_info["branch"]:
        candidate_info["branch"] = candidate_info["exam_name"]

    exam_title = candidate_info["exam_name"].upper()
    
    # (Delete the extra duplicate line here)

    candidate_info["header_html"] = (
        f'<div style="font-size: 20px; font-weight: 700; text-align: center; '
        f'padding: 10px; color: #1e293b; text-transform: uppercase;">'
        f'{exam_title}'
        f'</div>'
    )
    return candidate_info, all_questions


# ==============================================================================
# 🌟 MASTER PDF HANDLER (NEW FEATURE)
# ==============================================================================
async def safe_download_document(doc, save_path, max_retries=3):
    """Downloads a Telegram document safely with SSL decryption error handling and retries."""
    for attempt in range(max_retries):
        try:
            tg_file = await doc.get_file()
            await tg_file.download_to_drive(save_path)
            if os.path.exists(save_path) and os.path.getsize(save_path) > 0:
                return True
        except Exception as e:
            logger.warning(f"Download attempt {attempt + 1} failed due to SSL/Network error: {e}")
            if attempt == max_retries - 1:
                raise e
            await asyncio.sleep(2 * (attempt + 1))
    return False


async def handle_master_pdf(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handles Master PDF uploads and generates a locked HTML test."""
    if update.effective_user.id not in ADMIN_IDS:
        return await update.message.reply_text("❌ Admin only.")
        
    doc = update.message.document
    if not doc.file_name.lower().endswith(".pdf"):
        return await update.message.reply_text("❌ Please send a valid PDF file.")

    msg = await update.message.reply_text("📥 Downloading Master PDF…")
    
    # Download PDF safely with SSLError retry protection
    tmp_pdf = f"tmp_master_{doc.file_id}.pdf"
    out_html = f"master_test_{doc.file_id}.html"
    
    try:
        await safe_download_document(doc, tmp_pdf)
    except Exception as e:
        logger.error(f"Master PDF download failed: {e}")
        return await msg.edit_text(f"❌ Failed to download file due to SSL/Network error. Please try sending it again.")

    await msg.edit_text("🔍 Extracting Questions, Tables & Images… (This may take a minute)")

    try:
        # Run the flawless processor in a separate thread so the bot doesn't freeze
        loop = asyncio.get_running_loop()
        
        # This calls your exact master_pdf_processor.py script.
        questions, html_code = await loop.run_in_executor(
            None, process_master_pdf, tmp_pdf, None
        )
        
        answered = sum(1 for q in questions if q["answer"])
        
        # 🌟 Inject the Password Lock directly into the HTML string!
        password_overlay = f"""
        <div id="lockScreen" style="position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(15,23,42,0.95);backdrop-filter:blur(10px);display:flex;justify-content:center;align-items:center;z-index:9999;">
          <div style="background:#fff;padding:40px;border-radius:16px;text-align:center;max-width:400px;width:90%;">
            <h2>🔒 Test Locked</h2>
            <p>Enter the password to unlock this mock test.</p>
            <input type="password" id="pwdInput" style="width:100%;padding:12px;border:2px solid #e2e8f0;border-radius:8px;margin-bottom:15px;text-align:center;font-size:16px;">
            <button onclick="document.getElementById('pwdInput').value==='{TEST_PASSWORD}'?document.getElementById('lockScreen').style.display='none':alert('Incorrect Password')" style="width:100%;padding:12px;background:#3b82f6;color:#fff;border:none;border-radius:8px;font-weight:600;cursor:pointer;">Unlock Test</button>
          </div>
        </div>
        """
        
        # Insert the lock screen right after the <body> tag
        locked_html = html_code.replace("<body>", f"<body>{password_overlay}")

        # Save the locked HTML to a file
        with open(out_html, "w", encoding="utf-8") as f:
            f.write(locked_html)
        
        await msg.edit_text(
            f"✅ Extracted {len(questions)} questions\n"
            f"✅ Detected {answered} answers\n"
            f"📤 Uploading HTML test…"
        )
        
        # Send the file to the user
        with open(out_html, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=f"MockTest_{len(questions)}Q_{answered}Ans.html",
                caption=f"📝 {len(questions)} Questions | 🔑 {answered} Answers\n🔒 Password: `{TEST_PASSWORD}`\n\nUse this file in any browser to attempt!"
            )
            
    except Exception as e:
        logger.error(f"Master PDF processing error: {e}")
        await msg.edit_text(f"❌ Error processing PDF: {e}")
    finally:
        # Cleanup temp files
        if os.path.exists(tmp_pdf): os.remove(tmp_pdf)
        if os.path.exists(out_html): os.remove(out_html)



# ==============================================================================
# 8. EXCEL SYNC (HANDLES EXCEL OPEN-FILE LOCKS)
# ==============================================================================
async def update_excel_sheet():
    async with excel_lock:
        try:
            loop = asyncio.get_running_loop()

            def _sync():
                conn = sqlite3.connect(DB_FILE)
                query = """
                    SELECT 
                        COALESCE(exam_name, 'General Exam') AS "Exam Name",
                        candidate_name AS "Candidate Name",
                        hall_ticket AS "Roll No / ID",
                        branch AS "Branch / Subject",
                        total_score AS "Marks Obtained",
                        total_questions AS "Total Questions",
                        attempted AS "Total Attempted",
                        correct AS "Correct Answers",
                        incorrect AS "Incorrect Answers",
                        test_date AS "Test Date",
                        test_time AS "Test Time",
                        test_center AS "Test Center Name",
                        response_url AS "Response URL",
                        timestamp AS "Submitted Time"
                    FROM candidates
                    ORDER BY "Exam Name", branch, total_score DESC
                """
                df = pd.read_sql_query(query, conn)
                conn.close()

                if df.empty:
                    return

                target = EXCEL_FILE
                try:
                    with pd.ExcelWriter(target, engine="openpyxl") as writer:
                        df.to_excel(writer, index=False, sheet_name="Master Rankings")
                        for (exam, branch), grp in df.groupby(["Exam Name", "Branch / Subject"]):
                            clean_name = re.sub(r"[\:\\\/\?\*\[\]]", "_", f"{exam} - {branch}")[:30]
                            grp.to_excel(writer, index=False, sheet_name=clean_name)
                except PermissionError:
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    target = f"exam_results_snapshot_{ts}.xlsx"
                    df.to_excel(target, index=False, sheet_name="Master Rankings")
                    logger.warning(f"'{EXCEL_FILE}' was locked. Saved snapshot: '{target}'")

            await loop.run_in_executor(None, _sync)
        except Exception as e:
            logger.error(f"Excel error: {e}")


async def update_user_history_excel():
    async with history_lock:
        try:
            loop = asyncio.get_running_loop()

            def _sync_history():
                conn = sqlite3.connect(DB_FILE)
                query = """
                    SELECT 
                        user_id AS "Telegram User ID",
                        user_full_name AS "User Full Name",
                        username AS "Username",
                        exam_name AS "Exam Name",
                        candidate_name AS "Candidate Name",
                        hall_ticket AS "Hall Ticket / Roll No",
                        response_url AS "DigiALM Response Link",
                        submitted_at AS "Submitted At"
                    FROM submission_logs
                    ORDER BY id DESC
                """
                df = pd.read_sql_query(query, conn)
                conn.close()

                if df.empty:
                    return

                target = EXCEL_USER_HISTORY
                try:
                    with pd.ExcelWriter(target, engine="openpyxl") as writer:
                        df.to_excel(writer, index=False, sheet_name="All Link Submissions")
                        for uid, grp in df.groupby("Telegram User ID"):
                            grp.to_excel(writer, index=False, sheet_name=f"User_{uid}"[:30])
                except PermissionError:
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    target = f"user_submissions_snapshot_{ts}.xlsx"
                    df.to_excel(target, index=False, sheet_name="All Link Submissions")
                    logger.warning(f"'{EXCEL_USER_HISTORY}' was locked. Saved snapshot: '{target}'")

            await loop.run_in_executor(None, _sync_history)
        except Exception as e:
            logger.error(f"User History Excel error: {e}")



# ==============================================================================
# 8.5 HTML TEST GENERATOR
# ==============================================================================
import json
import random
from datetime import timedelta

async def embed_test_images_async(questions_data, base_url, referer_url=""):
    """Multi-strategy download with cache + context-aware fallback."""
    all_imgs = {}

    # 1. Collect all image URLs (question text + options)
    for q in questions_data:
        for k, v in q.items():
            if k == "options" and isinstance(v, list):
                for opt in v:
                    if isinstance(opt, str) and "<img" in opt:
                        for src in re.findall(r'src=["\']([^"\']+)["\']', opt):
                            if src.startswith("data:"): continue
                            full_url = src if src.startswith("http") else urljoin(base_url, src)
                            all_imgs[src] = full_url
            elif isinstance(v, str) and "<img" in v:
                for src in re.findall(r'src=["\']([^"\']+)["\']', v):
                    if src.startswith("data:"): continue
                    full_url = src if src.startswith("http") else urljoin(base_url, src)
                    all_imgs[src] = full_url

    total_imgs = len(all_imgs)
    if total_imgs == 0:
        return

    logger.info("Embedding %s unique images (cache + retry)...", total_imgs)
    src_map = {}
    cache_hits = 0
    failed_urls = []

    # Small placeholder SVG for failed question images
    _PLACEHOLDER_B64 = base64.b64encode(
        b'<svg xmlns="http://www.w3.org/2000/svg" width="140" height="38">'
        b'<rect width="140" height="38" fill="#f5f5f5" stroke="#d0d0d0" '
        b'stroke-dasharray="4,3" rx="5"/>'
        b'<text x="70" y="24" text-anchor="middle" fill="#999" '
        b'font-family="Arial,sans-serif" font-size="11">image unavailable</text>'
        b'</svg>'
    ).decode("utf-8")
    PLACEHOLDER_SVG = "data:image/svg+xml;base64," + _PLACEHOLDER_B64

    # 2. Download with cache + multi-strategy fallback
    async with httpx.AsyncClient(
        timeout=30.0,
        verify=False,
        follow_redirects=True,
        limits=httpx.Limits(max_connections=5, max_keepalive_connections=5)
    ) as client:

        count = 0
        for src, full_url in all_imgs.items():
            count += 1
            if count % 100 == 0 or count == total_imgs:
                logger.info("Downloading image %s / %s... (cache hits: %s)", count, total_imgs, cache_hits)

            # 🌟 FIX: Collapse /// to / but protect https://
            norm_url = re.sub(r'(?<!:)/{2,}', '/', full_url)

            # 🌟 CACHE: Try disk cache first
            cached = _img_cache_load(norm_url)
            if cached:
                src_map[src] = cached
                cache_hits += 1
                continue

            success = False
            
            # 🌟 DEFINE STRATEGIES (The part I accidentally deleted previously!)
            strategies = [
                (
                    "simple-pdf-mimic",
                    {
                        "User-Agent": "Mozilla/5.0",
                        "Referer": referer_url or "https://rrb.digialm.com/",
                    },
                    norm_url,
                ),
                (
                    "ua-only",
                    {"User-Agent": "Mozilla/5.0"},
                    norm_url,
                ),
                (
                    "cache-bust",
                    {
                        "User-Agent": "Mozilla/5.0",
                        "Referer": referer_url or "https://rrb.digialm.com/",
                    },
                    norm_url + ("&" if "?" in norm_url else "?") + f"_cb={int(time.time()*1000)}",
                ),
            ]

            # 1. Try httpx first (Fast)
            for label, headers, req_url in strategies:
                if success:
                    break
                try:
                    resp = await client.get(req_url, headers=headers)
                    ct = resp.headers.get("content-type", "").lower()
                    if resp.status_code == 200 and "image" in ct and len(resp.content) > 500:
                        b64 = base64.b64encode(resp.content).decode("utf-8")
                        data_uri = f"data:{ct};base64,{b64}"
                        src_map[src] = data_uri
                        _img_cache_save(norm_url, data_uri)
                        success = True
                        # logger.info("✅ %s via [%s]", os.path.basename(norm_url)[:50], label)
                        break
                except Exception:
                    pass
                await asyncio.sleep(0.3)

            # 2. 🌟 UNIVERSAL FALLBACK: IF HTTPX FAILS FOR ANY REASON, USE CURL
            if not success:
                try:
                    logger.info("⚠️ httpx failed, using curl fallback for: %s", norm_url[-40:])
                    import subprocess
                    import tempfile
                    
                    # Create a temporary file to hold the downloaded image
                    fd, temp_path = tempfile.mkstemp(suffix=".png")
                    os.close(fd)
                    
                    # Execute curl exactly like it works on your PC
                    cmd = [
                        "curl", "-L", "-s", "-o", temp_path,
                        "-H", "User-Agent: Mozilla/5.0",
                        "-H", f"Referer: {referer_url or 'https://rrb.digialm.com/'}",
                        norm_url
                    ]
                    
                    # Run curl silently and wait for it to finish
                    result = subprocess.run(cmd, capture_output=True, timeout=15)
                    
                    if result.returncode == 0 and os.path.exists(temp_path):
                        with open(temp_path, "rb") as f:
                            content = f.read()
                        
                        # Ensure the file is actually an image (> 100 bytes) and not an HTML error page
                        is_image = False
                        if len(content) > 100:
                            if content.startswith(b'\x89PNG') or content.startswith(b'\xff\xd8\xff') or content.startswith(b'GIF8'):
                                is_image = True
                        
                        if is_image:
                            b64 = base64.b64encode(content).decode("utf-8")
                            data_uri = f"data:image/png;base64,{b64}"
                            src_map[src] = data_uri
                            _img_cache_save(norm_url, data_uri)
                            success = True
                            # logger.info("✅ curl successfully bypassed firewall: %s", norm_url[-40:])
                            
                    # Clean up the temp file
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                        
                except Exception as e:
                    logger.error("❌ curl fallback failed: %s", e)

            if not success:
                logger.error("❌ All strategies failed: %s", norm_url)
                failed_urls.append(norm_url)

            await asyncio.sleep(0.02)

    # 3. Apply results — context-aware fallback for failures
    def _patch_imgs(html_str, is_option=False):
        soup = BeautifulSoup(html_str, "html.parser")
        changed = False
        for img in soup.find_all("img"):
            src = img.get("src")
            if not src:
                continue

            # Direct match
            if src in src_map:
                img["src"] = src_map[src]
                changed = True
                continue

            if src.startswith("data:"):
                continue  # Already embedded (MathML inline, etc.)

            # Normalized match (collapse /// to /)
            norm_src = re.sub(r'(?<!:)/{2,}', '/', src)
            if norm_src in src_map:
                img["src"] = src_map[norm_src]
                changed = True
                continue

            # Failed download → context-aware fallback
            failed_urls.append(norm_src)
            if is_option:
                img.decompose()
                # Strip invisible spaces (\xa0) to check if the option is truly empty
                remaining_text = soup.get_text(strip=True).replace('\xa0', '').strip()
                if not remaining_text and not soup.find("img"):
                    ph = soup.new_tag("span")
                    ph["style"] = "color:#999; font-style:italic; font-size:0.9em;"
                    ph.string = "[Image unavailable]"
                    soup.append(ph)
            else:
                # Question: small placeholder chip
                img["src"] = PLACEHOLDER_SVG
                if not img.get("alt"):
                    img["alt"] = "Image unavailable"
                img["style"] = "max-height:38px;display:inline-block;vertical-align:middle;border:0;"
            changed = True
        return str(soup) if changed else html_str

    for q in questions_data:
        for k, v in q.items():
            if k == "options" and isinstance(v, list):
                for idx, opt in enumerate(v):
                    if isinstance(opt, str) and "<img" in opt:
                        v[idx] = _patch_imgs(opt, is_option=True)
            elif isinstance(v, str) and "<img" in v:
                q[k] = _patch_imgs(v, is_option=False)

    # Log summary
    if cache_hits:
        logger.info("⚡ Cache saved %s network downloads", cache_hits)
    if failed_urls:
        try:
            stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            with open("failed_images.log", "a", encoding="utf-8") as f:
                for u in failed_urls:
                    f.write(f"[{stamp}] {u}\n")
            logger.warning("📋 %s failed URLs logged to failed_images.log", len(failed_urls))
        except Exception:
            pass

    logger.info("Image embedding complete: %s downloaded, %s cached, %s failed",
                len(src_map) - cache_hits, cache_hits, len(failed_urls))

def extract_test_data(html_input, is_ssc, questions_meta, exam_name=""):
    """Accepts a single HTML string OR a list of page strings."""
    pages = [html_input] if isinstance(html_input, str) else html_input
    test_data = []
    is_rrb = "RRB" in exam_name.upper()

    def safe_html(html_str):
        if not html_str: return ""
        return str(html_str).replace("`", "\\`").replace("${", "\\${")

    for page_html in pages:
        soup = BeautifulSoup(page_html, "html.parser")
        if is_ssc:
            # =================================================================
            # SSC LOGIC (Unchanged)
            # =================================================================
            for tbl in soup.find_all("table"):
                qno_count = len(tbl.find_all(string=re.compile(r"Q\.No\s*:\s*\d+", re.I)))
                if qno_count != 1:
                    continue
                qid, q_html, opts, correct_idx = "", "", [], 0
                for tr in tbl.find_all("tr"):
                    txt = tr.get_text()
                    if "Q.No" in txt:
                        qno_match = re.search(r"Q\.No\s*:\s*(\d+)", txt, re.I)
                        if qno_match:
                            qid = qno_match.group(1).strip()
                            q_html = safe_html(tr.find_all("td")[1])
                    elif "Hindi version" in txt or "Challenge" in txt:
                        continue
                    else:
                        tds = tr.find_all("td")
                        if len(tds) >= 2:
                            for i in range(0, len(tds), 2):
                                if i + 1 < len(tds):
                                    status_td = tds[i]
                                    opt_td = tds[i + 1]
                                    bg = str(status_td.get("bgcolor", "")).lower()
                                    if "green" in bg or "yellow" in bg:
                                        correct_idx = len(opts) + 1
                                    
                                    opt_html = str(opt_td)
                                    opt_soup = BeautifulSoup(opt_html, "html.parser")
                                    for img in opt_soup.find_all("img"):
                                        alt_text = img.get("alt", "").strip()
                                        src = img.get("src", "").lower()
                                        if any(k in src for k in ["spacer", "blank", "tick", "cross", "dummy", "1x1"]) or src.startswith("data:image/gif;base64,R0lGODlhAQAB"):
                                            if alt_text:
                                                img.replace_with(alt_text)
                                            else:
                                                img.decompose()
                                    opts.append(safe_html(opt_soup))
                if qid and len(opts) >= 2:
                    q_meta = next((q for q in questions_meta if q["qid"] == qid), {})
                    test_data.append({
                        "section": q_meta.get("section") or "General",
                        "original_qno": q_meta.get("master_qno"),
                        "passage": "",
                        "question": q_html,
                        "options": opts,
                        "q_type": "MCQ",
                        "answer": str(correct_idx),
                        "solution": "",           # 🆕 editable — add solutions here
                        "solution_images": [],    # 🆕 list of base64 data URIs (method 3)
                        "difficulty": "",         # 🆕 optional admin tag
                        "topic": "",              # 🆕 optional admin tag
                        "admin_notes": "",        # 🆕 internal — never shown to users
                        "positive_marks": q_meta.get("pos_marks", 2.0),
                        "negative_marks": q_meta.get("neg_marks", 0.50),
                    })
        else:
            # =================================================================
            # DIGIALM (RRB / AP Vidyut / GATE) LOGIC (MERGED & FIXED)
            # =================================================================
            for pnl in soup.find_all("div", class_="question-pnl"):
                q_html, opts, correct_idx = "", [], 0
                passage_html = ""
                
                # Detect GATE Question Types
                pnl_text = pnl.get_text(" ", strip=True)
                is_nat = "Question Type : NAT" in pnl_text or "Question Type :NAT" in pnl_text
                is_msq = "Question Type : MSQ" in pnl_text or "Question Type :MSQ" in pnl_text

                # 1. Extract Passage (Grabs BOTH "Comprehension:" label and the text)
                in_passage = False
                for tr in pnl.find_all("tr"):
                    txt = tr.get_text(" ", strip=True)
                    
                    if "Comprehension:" in txt or "Read the given passage" in txt:
                        in_passage = True
                        
                    if in_passage:
                        if re.match(r"^Q\.\s*\d+", txt) or "SubQuestion No" in txt:
                            break
                            
                        tds = tr.find_all("td")
                        if len(tds) >= 2:
                            target_td = tds[1]
                            nested = target_td.find("table")
                            if nested:
                                first_td = nested.find("td")
                                if first_td:
                                    target_td = first_td
                            
                            clean_text = target_td.get_text("\n", strip=True)
                            clean_text = clean_text.replace("\n", "<br/>")
                            if clean_text:
                                passage_html += clean_text + "<br/>"

                # 2. Extract Question Text (FIXED: Checks for img if text is empty for GATE)
                q_td = None
                for td in pnl.find_all("td"):
                    style = str(td.get("style", "")).lower()
                    if "text-align: left" in style and (td.get_text(strip=True) or td.find("img")) and not td.get("colspan"):
                        q_td = td
                        break
                
                if q_td:
                    q_html = safe_html(q_td.decode_contents())
                else:
                    for tr in pnl.find_all("tr"):
                        tds = tr.find_all("td")
                        if len(tds) >= 2 and re.match(r"^Q\.\s*\d+", tds[0].get_text(strip=True)):
                            for td in tds[1:]:
                                if td.get_text(strip=True) or td.find("img"):
                                    q_html = safe_html(td.decode_contents())
                                    break
                            if q_html:
                                break

                # 🌟 FIXED: Extract qid from the whole panel text (works for GATE/RRB)
                qid = ""
                m = re.search(r"Question\s*ID\s*[:\-]?\s*(\d+)", pnl.get_text(" ", strip=True), re.I)
                if m:
                    qid = m.group(1)

                # 3. Extract Options (Handles full-mark, NAT, and MSQ)
                q_meta = next((q for q in questions_meta if q["qid"] == qid), {})
                
                q_type = "MCQ"
                if re.search(r"Question\s*Type\s*:\s*NAT", pnl_text, re.I):
                    q_type = "NAT"
                elif re.search(r"Question\s*Type\s*:\s*MSQ", pnl_text, re.I):
                    q_type = "MSQ"
                elif q_meta.get("q_type"):
                    q_type = q_meta.get("q_type")

                if q_type == "NAT":
                    opts = []  # NAT has no options — just a text box
                else:
                    seen_opts = set()

                    # 🌟 FIX: recursive scan of ALL <tr> in the panel
                    #    (works for RRB/BEL empty <td> rows AND GATE "A./B." rows)
                    #    `recursive=False` on questionRowTbl missed <tbody>-wrapped rows.
                    for tr in pnl.find_all("tr"):
                        if len(opts) >= 4:
                            break

                        tds = tr.find_all("td")
                        if len(tds) < 2:
                            continue

                        first_text = tds[0].get_text(strip=True).replace(" ", "")
                        row_text  = tr.get_text(" ", strip=True)

                        # --- Skip question / metadata / passage rows ---
                        if re.match(r"^Q\.?\s*\d+$", first_text, re.I):
                            continue
                        if any(k in row_text for k in (
                            "QuestionID", "Question ID", "Comprehension:",
                            "Read the given passage", "SubQuestionNo", "Note:"
                        )):
                            continue

                        # --- Skip summary rows (multiple rightAns/wrngAns in same <tr>) ---
                        opt_tds = [td for td in tds if "rightAns" in str(td.get("class", []))
                                                       or "wrngAns" in str(td.get("class", []))]
                        if len(opt_tds) > 1:
                            continue

                        opt_td    = tds[1]
                        opt_class = str(opt_td.get("class", []))
                        opt_soup  = BeautifulSoup(str(opt_td), "html.parser")

                        # strip tick/cross icons (check filename in src ONLY, ignore alt text!)
                        for img in opt_soup.find_all("img"):
                            src_url = (img.get("src", "") or "").lower()
                            if any(k in src_url for k in ("tick.", "cross.", "check.", "wrong.", "correct.")):
                                img.decompose()

                        txt = opt_soup.get_text(strip=True)

                        # --- Is this an option row? ---
                        is_opt_row = False

                        # 🅰️ GATE-style: first cell is "A.", "B.", "C.", "D."
                        if re.match(r"^[A-D]\.?$", first_text, re.I):
                            is_opt_row = True

                        # 🅱️ RRB/BEL-style: empty first cell OR class="rightAns"/"wrngAns"
                        elif "rightAns" in opt_class or "wrngAns" in opt_class:
                            is_opt_row = True
                        elif first_text == "" or first_text == "Ans" or "Options" in first_text:
                            if opt_td.find("img") or txt:
                                is_opt_row = True

                        if not is_opt_row:
                            continue

                        # skip truly empty cells
                        if not txt and not opt_soup.find("img"):
                            continue

                        opt_html = safe_html(str(opt_soup))
                        if opt_html in seen_opts:
                            continue
                        seen_opts.add(opt_html)
                        opts.append(opt_html)

                # 4. Append Question (Accept NAT with 0 options)
                if q_html:
                    # 🌟 OVERRIDE WITH CORRECT ANSWER FROM KEY
                    correct_ans_from_key = str(q_meta.get("correct", "")).strip()
                    
                    is_two_mark = "2 MARK" in pnl_text.upper() or "MARKS : 2" in pnl_text.upper()
                    pos_fallback = 2.0 if is_two_mark else 1.0
                    neg_fallback = 0.0 if q_type in ["NAT", "MSQ"] else (pos_fallback / 3.0)
                    
                    test_data.append({
                        "section": q_meta.get("section") or "General",
                        "original_qno": q_meta.get("master_qno"),
                        "passage": passage_html,
                        "question": q_html,
                        "options": opts,
                        "q_type": q_type,
                        "answer": correct_ans_from_key if correct_ans_from_key else str(correct_idx),
                        "solution": "",           # 🆕 editable
                        "solution_images": [],    # 🆕 list of base64 data URIs
                        "difficulty": "",         # 🆕 optional
                        "topic": "",              # 🆕 optional
                        "admin_notes": "",        # 🆕 internal only
                        "positive_marks": q_meta.get("pos_marks", pos_fallback),
                        "negative_marks": q_meta.get("neg_marks", neg_fallback),
                        "full_marks_note": q_meta.get("full_marks_note", False),
                        "ambiguity_opts": q_meta.get("ambiguity_opts", []),
                    })

    return test_data                



def generate_html_test(info: dict, test_data: list) -> str:
    import json
    import re
    import os
    import base64
    import hashlib
    from datetime import datetime, timedelta
    
    sections_data = {}
    for q in test_data:
        sec_name = q.get("section") or "General"
        # Only add if the question actually exists
        if q.get("question"):
            sections_data.setdefault(sec_name, []).append({
            "question_heading": q.get("passage", ""),
            "question": q.get("question", ""),
            "answer": q.get("answer", ""),
            "positive_marks": q.get("positive_marks", 1.0),
            "negative_marks": q.get("negative_marks", 0.0),
            "full_marks_note": q.get("full_marks_note", False),      
            "ambiguity_opts": q.get("ambiguity_opts", []),           
            "q_type": q.get("q_type", "MCQ"),  # 🌟 FIX: THIS WAS MISSING!
            "solution": q.get("solution", ""),
            "solution_images": q.get("solution_images", []),
            **{f"option_{i+1}": opt for i, opt in enumerate(q.get("options", []))}
        })

    sections_list = list(sections_data.keys())
    grouped_questions = list(sections_data.values())
    total_q = len(test_data)
    
    safe_main = info.get("exam_name", "Competitive Exam")
    safe_sub = info.get("name", "Candidate")
    duration_minutes = 60

    # ========================================================================
    # 🛡️ Anti-tamper: XOR + split obfuscation (no per-user tracking)
    # ========================================================================
    _xor_key = 0x5A  # 90

    def _xor_encode(s):
        return ",".join(str(ord(c) ^ _xor_key) for c in s)

    def _split3(s):
        n = len(s)
        third = max(1, n // 3)
        return s[:third], s[third:2 * third], s[2 * third:]

    _b_enc = _xor_encode("@KranthikumargoudEEE")
    _t_enc = _xor_encode("t.me/kranthikumargoudEEE")
    _w_enc = _xor_encode("whatsapp.com/channel/0029Vam4uhw4CrfcNOn3eV0K")

    _bp1, _bp2, _bp3 = _split3(_b_enc)
    _tp1, _tp2, _tp3 = _split3(_t_enc)
    _wp1, _wp2, _wp3 = _split3(_w_enc)
    
    # 🌟 NEW: Convert the global switch to a string for JavaScript
    prevent_copy = "true" if COPY_PROTECTION else "false"
    
    # 🌟 Bulletproof Password & Expiry (Handles both int and string)
    password = str(TEST_PASSWORD) if TEST_PASSWORD else ""
    
    try:
        expiry_days = int(TEST_EXPIRY_DAYS) if TEST_EXPIRY_DAYS else 0
    except ValueError:
        expiry_days = 0
        
    expiry_iso = ""
    if expiry_days > 0:
        expiry_date = datetime.now() + timedelta(days=expiry_days)
        expiry_iso = expiry_date.strftime("%Y-%m-%d")

    encrypted_payload = None
    json_string = json.dumps({"sections": sections_list, "questions": grouped_questions}).replace("</", "<\\/").replace("\u2028", "\\u2028").replace("\u2029", "\\u2029")
    
    if password:
        try:
            from cryptography.hazmat.primitives.ciphers.aead import AESGCM
            salt = os.urandom(16)
            iv = os.urandom(12)
            key = hashlib.pbkdf2_hmac('sha256', password.encode('utf-8'), salt, 200000, 32)
            plaintext = json_string.encode('utf-8')
            aesgcm = AESGCM(key)
            ct_with_tag = aesgcm.encrypt(iv, plaintext, None)
            encrypted_payload = {
                'salt': base64.b64encode(salt).decode('ascii'),
                'iv': base64.b64encode(iv).decode('ascii'),
                'ct': base64.b64encode(ct_with_tag).decode('ascii'),
                'iterations': 200000,
            }
        except ImportError:
            pass

    json_string_to_embed = "null" if encrypted_payload else json_string
    password_js = "null" if encrypted_payload else (json.dumps(password) if password else '""')
    expiry_js = json.dumps(expiry_iso) if expiry_iso else '""'
    encrypted_json_js = json.dumps(encrypted_payload) if encrypted_payload else "null"
    
    total_marks = sum(float(q.get("positive_marks", 1.00)) for q in test_data)

    section_timer_inputs_html = ""
    for sec in sections_list:
        section_timer_inputs_html += f"""
        <div class="d-flex justify-content-between align-items-center mb-2">
            <span class="fw-bold">{sec}</span>
            <input type="number" class="form-control form-control-sm sec-timer-input" data-sec="{sec}" value="15" min="1" style="width: 80px;">
        </div>
        """

    html_template = r"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Mock Test - __SAFE_MAIN__</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <script src="https://code.jquery.com/jquery-1.8.0.min.js"></script>
    <script src="https://code.jquery.com/ui/1.11.4/jquery-ui.min.js"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/jqueryui-touch-punch/0.2.3/jquery.ui.touch-punch.min.js"></script>

    <script>
        const PREVENT_COPY = __PREVENT_COPY__;
        
        if (PREVENT_COPY) {
            document.addEventListener('contextmenu', event => event.preventDefault());
            document.onkeydown = function(e) {
                if(e.keyCode == 123) { return false; }
                if(e.ctrlKey && e.shiftKey && e.keyCode == 'I'.charCodeAt(0)) { return false; }
                if(e.ctrlKey && e.shiftKey && e.keyCode == 'C'.charCodeAt(0)) { return false; }
                if(e.ctrlKey && e.shiftKey && e.keyCode == 'J'.charCodeAt(0)) { return false; }
                if(e.ctrlKey && e.keyCode == 'U'.charCodeAt(0)) { return false; }
            };
            document.addEventListener('selectstart', function(e) { e.preventDefault(); });
            document.addEventListener('copy', function(e) { e.preventDefault(); });
            document.addEventListener('cut', function(e) { e.preventDefault(); });
        }
    </script>

    <style>
        :root { --primary-color: #1f4e79; --primary-gradient: linear-gradient(135deg, #1f4e79, #163a5a); --bg-color: #f4f6f9; --option-border: #e0e0e0; }
        body { background-color: var(--bg-color); font-family: 'Segoe UI', sans-serif; padding-bottom: 80px; }
        
        .top-navbar { background: var(--primary-color); color: white; padding: 10px 20px 0 20px; font-weight: bold; display: flex; flex-direction: column; align-items: stretch; position: sticky; top: 0; z-index: 100; }
        .brand-text { font-size: 0.9rem; font-weight: normal; opacity: 0.8; margin-left: 10px; }
        
        #section-tabs-bar { display: flex; gap: 4px; margin-top: 12px; padding-bottom: 0; overflow-x: auto; white-space: nowrap; scrollbar-width: thin; }
        .section-tab { border: 1px solid #ccc; border-bottom: none; border-radius: 8px 8px 0 0; padding: 8px 20px; font-size: 0.9rem; flex-shrink: 0; background: #e9ecef; color: #333; cursor: pointer; transition: all 0.2s; }
        .section-tab:hover:not(.locked-tab) { background: #d1d5db; }
        .section-tab.active-tab { background: #ffffff; color: #1f4e79; font-weight: 800; border-top: 3px solid #ffc107; }
        .section-tab.locked-tab { opacity: 0.6; cursor: not-allowed; background: #cbd5e1; }

        #access-gate { position: fixed; inset: 0; background: var(--primary-gradient); z-index: 99999; display: flex; align-items: center; justify-content: center; padding: 20px; }
        .gate-card { background: white; border-radius: 22px; padding: 40px 30px; max-width: 460px; width: 100%; text-align: center; box-shadow: 0 20px 60px rgba(0,0,0,0.35); }
        .gate-title { font-size: 1.7rem; font-weight: 800; margin-bottom: 10px; color: #333; }
        .gate-input { width: 100%; padding: 13px 16px; font-size: 1.05rem; border: 2px solid #ddd; border-radius: 10px; margin-bottom: 14px; }
        .gate-btn { background: var(--primary-gradient); color: white; border: none; padding: 13px 40px; border-radius: 28px; font-weight: bold; font-size: 1rem; width: 100%; cursor: pointer; }
        .gate-error { color: #dc3545; font-weight: bold; min-height: 22px; margin-bottom: 12px; font-size: 0.92rem; }

        #start-screen { display: flex; flex-direction: column; align-items: center; padding: 40px 20px; }
        .app-logo { width: 100px; height: 100px; background: var(--primary-color); border-radius: 50%; display: flex; align-items: center; justify-content: center; color: white; font-size: 50px; margin-bottom: 20px; box-shadow: 0 4px 15px rgba(31, 78, 121, 0.4); }
        .app-title { font-size: 2.5rem; font-weight: 800; color: #333; margin-bottom: 10px; }
        
        .option-label { display: flex; align-items: center; padding: 12px; border: 1px solid #ddd; border-radius: 8px; margin-bottom: 10px; cursor: pointer; background: #fff; transition: all 0.2s; }
        .option-label:hover { background: #f8f9fa; }
        .option-label.selected { border-color: #6a5acd; background: #f0ebff; }
        .option-label.correct-option { border-color: #28a745; background: #d4edda; }
        .option-label.wrong-option { border-color: #dc3545; background: #f8d7da; }
        .option-number { width: 30px; height: 30px; border-radius: 50%; background: #e9ecef; display: flex; align-items: center; justify-content: center; font-weight: bold; margin-right: 15px; flex-shrink: 0; }
        .option-label.selected .option-number { background: #6a5acd; color: white; }
        .option-label.correct-option .option-number { background: #28a745; color: white; }
        .option-label.wrong-option .option-number { background: #dc3545; color: white; }
        
        .question-text { font-size: 1.1rem; margin-bottom: 20px; }
        .passage-box { background: #eef3ff; border: 1px solid #cdd7f0; padding: 15px; border-radius: 8px; margin-bottom: 15px; max-height: 300px; overflow-y: auto; }
        .bottom-nav { position: fixed; bottom: 0; left: 0; right: 0; background: #fff; padding: 15px; display: flex; justify-content: space-between; z-index: 1000; border-top: 1px solid #ddd; }
        .modal-header { background: linear-gradient(135deg, #7b2cbf, #5a189a); color: white; }
        /* 🌟 PREMIUM DASHBOARD & NAVIGATION CSS */
        .legend-box { width: 16px; height: 16px; border-radius: 4px; display: inline-block; vertical-align: middle; margin-right: 5px; border: 1px solid #ccc; }
        .grid-container { display: grid; grid-template-columns: repeat(auto-fill, minmax(45px, 1fr)); gap: 10px; }
        .grid-btn { width: 45px; height: 45px; border-radius: 8px; border: 1px solid #ccc; background: #fff; font-weight: bold; display: flex; align-items: center; justify-content: center; cursor: pointer; transition: 0.2s; position: relative; }
        .grid-btn:hover { background: #eee; }
        .grid-btn.answered { background: #28a745; color: white; border-color: #28a745; }
        .grid-btn.marked { background: #6a5acd; color: white; border-color: #6a5acd; }
        .grid-btn.marked::after { content: ''; position: absolute; top: -4px; right: -4px; width: 12px; height: 12px; background: #f0ad4e; border-radius: 50%; border: 2px solid #fff; }
        .grid-btn.visited { background: #e9ecef; }
        .grid-btn.current { border: 3px solid #222; }

        .dashboard-card { border-radius: 15px; box-shadow: 0 4px 20px rgba(0,0,0,0.08); background: #f8f9fa; }
        .dashboard-header { background: linear-gradient(135deg, #1f4e79, #163a5a); color: white; padding: 20px; border-radius: 15px 15px 0 0; }
        .score-circle { position: relative; width: 120px; height: 120px; border-radius: 50%; background: conic-gradient(var(--primary-color) 0%, #e9ecef 0%); display: flex; align-items: center; justify-content: center; margin: 0 auto; }
        .score-inner { width: 100px; height: 100px; border-radius: 50%; background: white; display: flex; flex-direction: column; align-items: center; justify-content: center; }
        .score-text { font-size: 1.5rem; font-weight: bold; color: var(--primary-color); }
        .stat-box-colored { border-radius: 10px; padding: 15px; color: white; text-align: center; font-weight: bold; }
        .bg-purple { background: #6a5acd; }
        .bg-teal { background: #20c997; }
        .bg-pink { background: #e83e8c; }
        .bg-orange { background: #fd7e14; }
        .detail-box { background: #fff; border-radius: 10px; padding: 15px; text-align: center; border: 1px solid #eee; }
        
        .q-timer { font-size: 0.85rem; background: #ffeeba; color: #856404; padding: 4px 10px; border-radius: 15px; font-weight: bold; border: 1px solid #ffe8a1; }

        #full-paper-screen { display: none; padding-bottom: 50px; }
        .full-q-card { margin-bottom: 25px; border-bottom: 2px solid #ddd; padding-bottom: 20px; page-break-inside: avoid; }
        .full-opt { padding: 8px 12px; margin-bottom: 5px; border: 1px solid #eee; border-radius: 6px; font-size: 0.95rem; }
        .full-opt.correct { background: #d4edda; border-color: #28a745; font-weight: bold; color: #155724; }
        .full-opt.wrong { background: #f8d7da; border-color: #dc3545; color: #721c24; }
        .full-opt.user-selected { border-left: 5px solid #6a5acd; }

        /* 🌟 TCS SCIENTIFIC CALCULATOR CSS — FULL WORKING VERSION */
        .calc-btn-float { position: fixed; bottom: 90px; right: 20px; z-index: 1050; background: #4a90e2; color: white; border: none; border-radius: 50%; width: 60px; height: 60px; font-size: 24px; box-shadow: 0 4px 15px rgba(0,0,0,0.2); display: none; justify-content: center; align-items: center; cursor: pointer; }
        .calc-btn-float:hover { transform: scale(1.05); }
        #calc-wrapper { position: fixed; top: 100px; left: max(0px, calc(100vw - 483px)); z-index: 1049; display: none; }
        .calc_close { background: #e74c3c; width: 16px; height: 16px; color: #fff; cursor: pointer; font-size: 12px; font-weight: 900; line-height: 16px; position: absolute; right: 12px; text-align: center; top: 12px; border-radius: 5px; }
        #keyPad_Help { color: #000; background: #ddd; width: 54px; height: 25px; padding: 1px; cursor: pointer; font-size: 12px; font-weight: bold; line-height: 20px; position: absolute; right: 79px; text-align: center; top: 7px; border-radius: 4px; }
        #keyPad_Helpback { color: #000; background: #ddd; width: 54px; height: 25px; padding: 1px; cursor: pointer; font-size: 12px; font-weight: bold; line-height: 20px; position: absolute; right: 79px; text-align: center; top: 7px; border-radius: 4px; }
        .calc_min { background: #555; width: 16px; height: 3px; color: #fff; cursor: pointer; font-size: 12px; font-weight: bold; line-height: 25px; position: absolute; right: 42px; text-align: center; top: 25px; border-radius: 2px; }
        .calc_max { background: #555; width: 16px; height: 16px; color: #fff; cursor: pointer; font-size: 12px; font-weight: bold; line-height: 25px; position: absolute; right: 42px; text-align: center; top: 12px; border-radius: 2px; }
        .hide { display: none; }
        .reduceWidth { width: 200px !important; }
        .reduceHeader { display: block; overflow: hidden; text-overflow: ellipsis; white-space: pre; width: 74px !important; }
        #mainContentArea { margin-top: 0px; padding: 10px; }
        .calc_container { background: #dadada; border: 1px solid #aaa; border-radius: 5px; box-shadow: 0 0 15px 0 rgba(0,0,0,0.1) inset; padding: 0px; width: 463px; border-bottom: 7px solid #999; position: relative; }
        .left_sec { float: left; }
        .right_sec { float: left; height: 180px; }
        .calc_row { display: block; overflow: hidden; margin-bottom: 5px; }
        #keyPad a { background: #f1f1f1; border-color: #aaa; border-image: none; border-radius: 4px; border-style: solid; border-width: 1px 1px 3px; color: #444; cursor: pointer; float: left; font-size: 12px; font-weight: bold; height: 25px; line-height: 20px; margin-right: 5px; text-align: center; text-decoration: none; vertical-align: middle; width: 35px; }
        #keyPad a.min { font-size: 10.5px; }
        #keyPad a.last { margin-right: 0px; }
        #keyPad a:active { position: relative; height: 25px; border-bottom: 1px solid #aaa; }
        #keyPad_btnEnter:active { position: absolute !important; bottom: 17px !important; }
        #keyPad a:hover { background: #eaeaea; }
        .text_container { display: inline; }
        .memoryhide { visibility: hidden; position: absolute; top: 52px; right: 429px; bottom: 0; opacity: 0; transition: opacity 1s, visibility, 1s; }
        .memoryshow { visibility: visible; opacity: 1; }
        .keyPad_TextBox, .keyPad_TextBox1 { background: #fff; border-color: #aaa; border-image: none; border-radius: 3px; border-style: solid; border-width: 1px 1px 3px; font-size: 20px; height: 30px; line-height: 30px; text-align: right; text-indent: 10px; vertical-align: middle; width: 434px; float: left; margin-bottom: 7px; padding-right: 5px; }
        .degree_radian { float: left; height: 25px; width: 80px; display: inline-block; font-size: 9px; margin-top: 4px; vertical-align: middle; }
        #keyPad_btnDot { width: 35px !important; }
        #keyPad_btn0 { width: 76px !important; }
        .calc_arrows { font-size: 24px !important; }
        #keyPad #keyPad_btnEnter { background: #2ecc71; border-color: #27ae60; bottom: 17px; color: #fff; font-size: 20px; line-height: 47px; right: 10px; width: 35px; height: 53px; position: absolute; }
        #keyPad #keyPad_btnClr, #keyPad #keyPad_btnInverseSign, #keyPad #keyPad_btnAllClr { background: #e74c3c; color: #fff; border-color: #c0392b; width: 35px !important; }
        #keyPad #keyPad_btnBack { background: #e74c3c; color: #fff; border-color: #c0392b; width: 76px !important; }
        #keyPad #keyPad_btnSquare, #keyPad #keyPad_btnCube, #keyPad #keyPad_btnExp, #keyPad #keyPad_btnYpowX { font-size: 12px; }
        .clear { clear: both; }
        .superscript, .subscript { font-size: 0.7em; position: relative; top: -0.5em; }
        .subscript { top: 0.5em; }
        #helpContent { height: 263px; margin-right: 6px; overflow: auto; padding: 10px; }
        #helptopDiv { margin-bottom: 5px; padding: 10px; background-color: #4286f3; color: #ffffff; cursor: move; touch-action: none; user-select: none; -webkit-user-select: none; -webkit-tap-highlight-color: transparent; }
        @media (max-width: 560px) {
            #calc-wrapper { top: 55px; left: 5px; width: calc(100vw - 10px); max-height: calc(100vh - 65px); overflow-y: auto; overflow-x: hidden; -webkit-overflow-scrolling: touch; }
            .calc_container { width: 100% !important; max-width: 100%; box-sizing: border-box; }
            #mainContentArea { padding: 6px; }
            .keyPad_TextBox, .keyPad_TextBox1 { width: calc(100% - 20px); font-size: 16px; height: 28px; line-height: 26px; box-sizing: border-box; padding-right: 2px; }
            .left_sec { width: 100%; box-sizing: border-box; }
            #keyPad a { width: calc((100% - 55px) / 11); height: 28px; line-height: 22px; font-size: 10px; padding: 0; box-sizing: border-box; }
            #keyPad #keyPad_btn0, #keyPad #keyPad_btnBack { width: calc((100% - 55px) / 11 * 2 + 5px) !important; }
            #keyPad #keyPad_btnEnter { right: 6px; bottom: 6px; width: calc((100% - 55px) / 11); height: 58px; line-height: 50px; font-size: 16px; }
            .degree_radian { width: 50px; font-size: 8px; height: 22px; margin-top: 3px; }
            .degree_radian input { margin: 0; }
        }
        
        /* 🌟 WATERMARK FOR PDF */
        .watermark { display: none; }
        
        @media print {
            .top-navbar, .bottom-nav, .calc-btn-float, #calc-wrapper, .modal, #quiz-screen, #print-controls { display: none !important; }
            #full-paper-screen { display: block !important; padding: 0 !important; }
            body { background: white; padding: 0; margin: 0; position: relative; }
            
            .watermark {
                display: block !important;
                position: fixed !important;
                top: 50% !important;
                left: 50% !important;
                transform: translate(-50%, -50%) rotate(-45deg) !important;
                font-size: 80px !important;
                font-weight: bold !important;
                color: rgba(0, 0, 0, 0.08) !important;
                z-index: 9999 !important;
                pointer-events: none !important;
                white-space: nowrap !important;
            }
        }


        /* =========================================================
           🌟 RESPONSIVE FIXES — laptop + mobile comfort
           ========================================================= */

        /* ---------- 1. Images NEVER overflow (critical) ---------- */
        #question-container img,
        #full-paper-screen img,
        .passage-box img,
        .solution-box img,
        .question-text img,
        .full-opt img,
        .option-label img,
        .modal-body img,
        .sol-card img {
            max-width: 100% !important;
            height: auto !important;
            display: inline-block;
            vertical-align: middle;
        }

        /* ---------- 2. Tables stay inside the card (critical) ---------- */
        #question-container table,
        #full-paper-screen table,
        .passage-box table,
        .option-label table,
        .full-opt table,
        .modal-body table,
        .sol-card table {
            max-width: 100% !important;
            table-layout: auto;
            word-break: break-word;
            border-collapse: collapse;
        }
        #question-container td,
        #question-container th,
        #full-paper-screen td,
        #full-paper-screen th,
        .passage-box td,
        .passage-box th,
        .sol-card td,
        .sol-card th {
            overflow-wrap: anywhere;
            word-break: break-word;
        }

        /* ---------- 2b. Bilingual (English + Telugu) text and merged tables from /PT ---------- */
        .telugu-text { display: block; color: #475569; font-size: 0.95em; font-weight: 600; margin-top: 6px; }
        .extracted-table { width: 100%; border-collapse: collapse; margin: 12px 0; font-size: 0.92em; background: #fff; border: 1px solid #cbd5e1; }
        .extracted-table th, .extracted-table td { border: 1px solid #cbd5e1; padding: 8px 12px; text-align: left; vertical-align: top; color: #334155; }
        .extracted-table th { background: #f1f5f9; font-weight: 600; }
        .extracted-table tr:nth-child(even) td { background: #f8fafc; }
        .extracted-table .telugu-text { margin-top: 3px; }

        /* ---------- 3. Question header wraps, never overflows ---------- */
        #question-container .card > .d-flex.justify-content-between {
            flex-wrap: wrap;
            row-gap: 8px;
        }

        /* ---------- 4. Section timer rows wrap nicely ---------- */
        #section-time-setup .d-flex.justify-content-between {
            flex-wrap: wrap;
            row-gap: 6px;
        }

        /* ---------- 5. Long option text / questions break correctly ---------- */
        .question-text,
        .option-content,
        .full-opt,
        .sol-body,
        .explanation-box {
            overflow-wrap: anywhere;
            word-break: break-word;
        }

        /* ---------- 6. Mobile comfort (≤640px) ---------- */
        @media (max-width: 640px) {

            /* Top navbar: title + subtitle on one line, buttons wrap below */
            .top-navbar > .d-flex {
                flex-wrap: wrap;
                row-gap: 8px;
            }
            .top-navbar > .d-flex > div:first-child {
                flex: 1 1 100%;
                min-width: 0;
            }
            .top-navbar .brand-text {
                display: block;
                max-width: 100%;
                white-space: normal;      /* 🌟 Allows the text to wrap */
                overflow-wrap: anywhere;  /* Breaks long strings if needed */
                line-height: 1.3;
                margin-top: 4px !important;
            }

            /* Question card breathing room */
            #question-container .card { padding: 1rem !important; }
            .question-text { font-size: 1rem; line-height: 1.45; }

            /* Compact option rows */
            .option-label { padding: 10px !important; }
            .option-number {
                width: 26px !important;
                height: 26px !important;
                font-size: 0.82rem !important;
                margin-right: 10px !important;
            }

            /* Compact per-question timer */
            .q-timer { font-size: 0.75rem; padding: 3px 8px; }

            /* Section tabs scroll nicely, tighter padding */
            .section-tab { padding: 6px 12px; font-size: 0.8rem; }

            /* Results modal — smaller headings */
            #final-score { font-size: 2.5rem !important; }
            .modal-title { font-size: 1rem !important; }

            /* Full paper cards — better spacing on phone */
            .full-q-card { margin-bottom: 18px; padding-bottom: 14px; }
            .full-opt { padding: 6px 10px; font-size: 0.9rem; }

            /* Grid buttons — compact */
            .grid-btn {
                width: 42px;
                height: 42px;
                font-size: 0.85rem;
            }
        }

        /* ---------- 7. Extra-tight screens (≤360px, iPhone SE) ---------- */
        @media (max-width: 360px) {
            .top-navbar .brand-text { font-size: 0.65rem !important; }
            #question-container .card { padding: 0.75rem !important; }
            .option-number {
                width: 24px !important;
                height: 24px !important;
                font-size: 0.78rem !important;
                margin-right: 8px !important;
            }
            .grid-btn { width: 38px; height: 38px; }
        }

        /* ---------- 8. Print / PDF safety ---------- */
        @media print {
            #question-container img,
            #full-paper-screen img,
            .passage-box img,
            .sol-card img {
                max-width: 100% !important;
            }
            .sol-card,
            .full-q-card,
            .question-text,
            .explanation-box {
                page-break-inside: avoid;
            }
        }

    </style>
</head>
<body>

<div class="watermark"><span class="brand"></span></div>

<div id="access-gate">
    <div id="gate-card" class="gate-card">
        <div style="color:#999; font-size:0.9rem;">Checking access...</div>
    </div>
</div>

<div class="top-navbar" id="top-nav" style="display:none;">
    <div class="d-flex justify-content-between align-items-center w-100">
        <!-- 🌟 NEW: Adding Subject, Date, and Time to the top navbar -->
        <div>
           <div>🎓 <span class="brand"></span></div>
            <div class="brand-text text-white" style="font-size:0.75rem; margin-top:2px;">
                __SAFE_MAIN__ | Sub: __SAFE_SUBJECT__ | Date: __SAFE_DATE__ | Time: __SAFE_TIME__
            </div>
        </div>
        <div class="d-flex gap-3">
            <div id="timer-display" class="badge bg-light text-dark fs-6"><i class="far fa-clock"></i> --:--</div>
            <button class="btn btn-sm btn-info fw-bold text-white" id="btn-show-results" style="display:none;" onclick="resultsModal.show()">Results</button>
            <button class="btn btn-sm btn-warning fw-bold" onclick="openGrid()">Question Paper</button>
        </div>
    </div>
    <div id="section-tabs-bar"></div>
</div>

<div id="setup-screen" class="container mt-5">
    <div class="card shadow-lg p-4 mx-auto" style="max-width: 600px;">
        <div style="text-align:center; margin-bottom: 15px; width: 100%;">
            __SAFE_HEADER_HTML__
        </div>
        <h4 class="text-center mb-3 text-primary fw-bold">__SAFE_MAIN__</h4>
        
        <!-- 🌟 NEW: Subject, Date, and Time Info -->
        <div class="text-center text-muted mb-4" style="font-size: 0.95rem;">
            <div class="mb-1"><strong>Subject:</strong> __SAFE_SUBJECT__</div>
            <div><strong>Date:</strong> __SAFE_DATE__ &nbsp;|&nbsp; <strong>Time:</strong> __SAFE_TIME__</div>
        </div>
        <hr class="mb-4">

        <p class="text-center text-muted mb-1">Total Questions: __QCOUNT__ | Total Marks: __TOTAL_MARKS_2F__</p>
        
        <div class="mb-3">
            <label class="fw-bold mb-2">Timer Mode:</label>
            <select id="timer-mode" class="form-select" onchange="toggleTimerUI()">
                <option value="full">Full Test (Single Timer)</option>
                <option value="sectional">Sectional (Timer per section)</option>
            </select>
        </div>
        <div id="full-time-setup" class="mb-3">
            <label class="fw-bold mb-2">Total Time (Minutes):</label>
            <input type="number" id="full-time-input" class="form-control" value="__DURATION__" min="1">
        </div>
        <div id="section-time-setup" class="mb-3" style="display: none;">
            <label class="fw-bold mb-2 text-primary">Customize Section Timers (Minutes):</label>
            <div class="p-3 bg-light rounded border">__SECTION_INPUTS__</div>
        </div>
        <div class="mb-3 form-check">
            <input type="checkbox" class="form-check-input" id="calc-toggle" checked>
            <label class="form-check-label" for="calc-toggle">Enable Scientific Calculator</label>
        </div>
        <button class="btn btn-primary w-100 fw-bold fs-5 mt-3" onclick="startTest()">Start Test</button>
        <div class="mt-4 text-center">
            <p class="fw-bold text-secondary mb-2">📢 Join our channel for latest updates:</p>
            <a data-ch-tg href="#" target="_blank" class="btn btn-sm btn-outline-primary me-2 mb-2">👉 Telegram Channel</a>
            <a data-ch-wa href="#" target="_blank" class="btn btn-sm btn-outline-success mb-2">👉 WhatsApp Channel</a>
        </div>
    </div>
</div>

<div id="quiz-screen" class="container mt-4" style="display: none;">
    <div id="question-container"></div>
</div>

<div id="full-paper-screen" class="container mt-4"></div>

<div class="bottom-nav" id="bottom-nav" style="display: none; max-width: 900px; margin: 0 auto; border-radius: 12px 12px 0 0;">
    <button class="btn btn-outline-secondary fw-bold" onclick="prevQuestion()" id="btn-prev">Prev</button>
    <button class="btn btn-warning fw-bold text-white" onclick="toggleReview()" id="btn-review">Review</button>
    <button class="btn btn-primary fw-bold" onclick="nextQuestion()" id="btn-next">Next</button>
</div>

<button class="calc-btn-float" id="calc-float-btn" onclick="toggleCalc()">🧮</button>

<div id="calc-wrapper">
    <div id="keyPad" class="ui-widget-content calc_container">
        <div id="helptopDiv">
            <span>Scientific Calculator</span>
            <div href="#nogo" id="keyPad_Help" class="help_back">Help</div>
            <div style="display:none;" href="#nogo" id="keyPad_Helpback" class="help_back">Back</div>
        </div>
        <div class="calc_min" id="calc_min"></div>
        <div class="calc_close" id="closeButton">✖</div>

        <div id="mainContentArea">
            <input type="text" id="keyPad_UserInput1" class="keyPad_TextBox1" readonly />
            <div class="text_container">
                <input type="text" id="keyPad_UserInput" class="keyPad_TextBox" maxlength="30" readonly />
                <span id="memory" class="memoryhide"><font size="2">M</font></span>
            </div>
            <div class="clear"></div>
            <div class="left_sec">
                <div class="calc_row clear">
                    <a href="javascript:void(0)" id="keyPad_btnMod" class="keyPad_btnBinaryOp">mod</a>
                    <div class="degree_radian">
                        <input type="radio" name="degree_or_radian" value="deg" checked="checked">Deg
                        <input type="radio" name="degree_or_radian" value="rad">Rad
                    </div>
                    <a href="javascript:void(0)" id="keyPad_MC" class="keyPad_btnMemoryOp">MC</a>
                    <a href="javascript:void(0)" id="keyPad_MR" class="keyPad_btnMemoryOp">MR</a>
                    <a href="javascript:void(0)" id="keyPad_MS" class="keyPad_btnMemoryOp">MS</a>
                    <a href="javascript:void(0)" id="keyPad_M+" class="keyPad_btnMemoryOp">M+</a>
                    <a href="javascript:void(0)" id="keyPad_M-" class="keyPad_btnMemoryOp">M-</a>
                </div>
                <div class="calc_row clear">
                    <a href="javascript:void(0)" id="keyPad_btnSinH" class="keyPad_btnUnaryOp min">sinh</a>
                    <a href="javascript:void(0)" id="keyPad_btnCosinH" class="keyPad_btnUnaryOp min">cosh</a>
                    <a href="javascript:void(0)" id="keyPad_btnTgH" class="keyPad_btnUnaryOp min">tanh</a>
                    <a href="javascript:void(0)" id="keyPad_EXP" class="keyPad_btnBinaryOp">Exp</a>
                    <a href="javascript:void(0)" id="keyPad_btnOpen" class="keyPad_btnBinaryOp">(</a>
                    <a href="javascript:void(0)" id="keyPad_btnClose" class="keyPad_btnBinaryOp">)</a>
                    <a href="javascript:void(0)" id="keyPad_btnBack" class="keyPad_btnCommand calc_arrows"><div style="position: relative; top: -3px">&#8592;</div></a>
                    <a href="javascript:void(0)" id="keyPad_btnAllClr" class="keyPad_btnCommand">C</a>
                    <a href="javascript:void(0)" id="keyPad_btnInverseSign" class="keyPad_btnUnaryOp">+/-</a>
                    <a href="javascript:void(0)" id="keyPad_btnSquareRoot" class="keyPad_btnUnaryOp"><div style="position: relative; top: 1px">&#8730;</div></a>
                </div>
                <div class="calc_row clear" style="margin-top: 5px;">
                    <a href="javascript:void(0)" id="keyPad_btnAsinH" class="keyPad_btnUnaryOp min"><span class='baseele'>sinh</span><span class='superscript'>-1</span></a>
                    <a href="javascript:void(0)" id="keyPad_btnAcosH" class="keyPad_btnUnaryOp min"><span class='baseele'>cosh</span><span class='superscript'>-1</span></a>
                    <a href="javascript:void(0)" id="keyPad_btnAtanH" class="keyPad_btnUnaryOp min"><span class='baseele'>tanh</span><span class='superscript'>-1</span></a>
                    <a href="javascript:void(0)" id="keyPad_btnLogBase2" class="keyPad_btnUnaryOp"><span class='baseele'>log</span><span class='subscript'>2</span><span class='baseele'>x</span></a>
                    <a href="javascript:void(0)" id="keyPad_btnLn" class="keyPad_btnUnaryOp">ln</a>
                    <a href="javascript:void(0)" id="keyPad_btnLg" class="keyPad_btnUnaryOp">log</a>
                    <a href="javascript:void(0)" id="keyPad_btn7" class="keyPad_btnNumeric">7</a>
                    <a href="javascript:void(0)" id="keyPad_btn8" class="keyPad_btnNumeric">8</a>
                    <a href="javascript:void(0)" id="keyPad_btn9" class="keyPad_btnNumeric">9</a>
                    <a href="javascript:void(0)" id="keyPad_btnDiv" class="keyPad_btnBinaryOp">/</a>
                    <a href="javascript:void(0)" id="keyPad_%" class="keyPad_btnBinaryOp">%</a>
                </div>
                <div class="calc_row clear">
                    <a href="javascript:void(0)" id="keyPad_btnPi" class="keyPad_btnConst">&#960;</a>
                    <a href="javascript:void(0)" id="keyPad_btnE" class="keyPad_btnConst">e</a>
                    <a href="javascript:void(0)" id="keyPad_btnFact" class="keyPad_btnUnaryOp">n!</a>
                    <a href="javascript:void(0)" id="keyPad_btnYlogX" class="keyPad_btnBinaryOp"><span class='baseele'>log</span><span class='subscript'>y</span><span class='baseele'>x</span></a>
                    <a href="javascript:void(0)" id="keyPad_btnExp" class="keyPad_btnUnaryOp"><span class='baseele'>e</span><span class='superscript'>x</span></a>
                    <a href="javascript:void(0)" id="keyPad_btn10X" class="keyPad_btnUnaryOp"><span class='baseele'>10</span><span class='superscript'>x</span></a>
                    <a href="javascript:void(0)" id="keyPad_btn4" class="keyPad_btnNumeric">4</a>
                    <a href="javascript:void(0)" id="keyPad_btn5" class="keyPad_btnNumeric">5</a>
                    <a href="javascript:void(0)" id="keyPad_btn6" class="keyPad_btnNumeric">6</a>
                    <a href="javascript:void(0)" id="keyPad_btnMult" class="keyPad_btnBinaryOp"><div style="position: relative; top: 3px; font-size: 20px">*</div></a>
                    <a href="javascript:void(0)" id="keyPad_btnInverse" class="keyPad_btnUnaryOp"><span class='baseele'>1/x</span></a>
                </div>
                <div class="calc_row clear">
                    <a href="javascript:void(0)" id="keyPad_btnSin" class="keyPad_btnUnaryOp min">sin</a>
                    <a href="javascript:void(0)" id="keyPad_btnCosin" class="keyPad_btnUnaryOp min">cos</a>
                    <a href="javascript:void(0)" id="keyPad_btnTg" class="keyPad_btnUnaryOp min">tan</a>
                    <a href="javascript:void(0)" id="keyPad_btnYpowX" class="keyPad_btnBinaryOp"><span class='baseele'>x</span><span class='superscript'>y</span></a>
                    <a href="javascript:void(0)" id="keyPad_btnCube" class="keyPad_btnUnaryOp"><span class='baseele'>x</span><span class='superscript'>3</span></a>
                    <a href="javascript:void(0)" id="keyPad_btnSquare" class="keyPad_btnUnaryOp"><span class='baseele'>x</span><span class='superscript'>2</span></a>
                    <a href="javascript:void(0)" id="keyPad_btn1" class="keyPad_btnNumeric">1</a>
                    <a href="javascript:void(0)" id="keyPad_btn2" class="keyPad_btnNumeric">2</a>
                    <a href="javascript:void(0)" id="keyPad_btn3" class="keyPad_btnNumeric">3</a>
                    <a href="javascript:void(0)" id="keyPad_btnMinus" class="keyPad_btnBinaryOp"><div style="position: relative; top: -1px; font-size: 20px">-</div></a>
                </div>
                <div class="calc_row clear">
                    <a href="javascript:void(0)" id="keyPad_btnAsin" class="keyPad_btnUnaryOp min"><span class='baseele'>sin</span><span class='superscript'>-1</span></a>
                    <a href="javascript:void(0)" id="keyPad_btnAcos" class="keyPad_btnUnaryOp min"><span class='baseele'>cos</span><span class='superscript'>-1</span></a>
                    <a href="javascript:void(0)" id="keyPad_btnAtan" class="keyPad_btnUnaryOp min"><span class='baseele'>tan</span><span class='superscript'>-1</span></a>
                    <a href="javascript:void(0)" id="keyPad_btnYrootX" class="keyPad_btnBinaryOp"><span class='superscript' style='top: -8px;'>y</span><span class='baseele' style='font-size: 1.2em; margin: -6px 0 0 -9px;'>&#8730;x</span></a>
                    <a href="javascript:void(0)" id="keyPad_btnCubeRoot" class="keyPad_btnUnaryOp"><font size="3">&#8731;</font></a>
                    <a href="javascript:void(0)" id="keyPad_btnAbs" class="keyPad_btnUnaryOp"><span class='baseele'>|x|</span></a>
                    <a href="javascript:void(0)" id="keyPad_btn0" class="keyPad_btnNumeric">0</a>
                    <a href="javascript:void(0)" id="keyPad_btnDot" class="keyPad_btnNumeric">.</a>
                    <a href="javascript:void(0)" id="keyPad_btnPlus" class="keyPad_btnBinaryOp">+</a>
                    <a href="javascript:void(0)" id="keyPad_btnEnter" class="keyPad_btnCommand"><div style="margin-bottom: 2px;">=</div></a>
                </div>
            </div>
            <div class="clear"></div>

            <div id="helpContent" onmousedown="return false" style="display:none;">
                <h3 style="text-align:center;"><strong>Calculator Instructions</strong></h3>
                Allows you to perform basic and complex mathematical operations such as modulus, square root, cube root, trigonometric, exponential, logarithmic, hyperbolic functions, etc.<br><br>
                <h3 style="text-decoration: underline; color: green;">Do's:</h3>
                <ul>
                    <li>Be sure to press [C] when beginning a new calculation.</li>
                    <li>Use predefined operations such as p (Pi), log, Exp to save time.</li>
                    <li>Use memory functions for cumulative totals.</li>
                    <strong>[M+]: Add to memory. [MR]: Recall. [M-]: Subtract.</strong>
                    <li>Select angle unit (Deg or Rad) before calculating. Default is Degree.</li>
                </ul>
                <h3><span style="text-decoration: underline; color: red;">Don'ts:</span></h3>
                <ul>
                    <li>Perform multiple operations together.</li>
                    <li>Leave parenthesis unbalanced.</li>
                    <li>Change angle unit (Deg/Rad) mid-calculation.</li>
                </ul>
                <h3><span style="text-decoration: underline;">Limitations:</span></h3>
                <ul>
                    <li>Keyboard operation is disabled.</li>
                    <li>Factorial precise up to 14 digits.</li>
                    <li>Logarithmic/Hyperbolic precise up to 5 digits.</li>
                    <li>Mod on 15+ digit decimals is imprecise.</li>
                    <li>Supported range: 10(-323) to 10(308).</li>
                </ul>
            </div>
        </div>
    </div>
</div>

<div class="modal fade" id="gridModal" tabindex="-1">
  <div class="modal-dialog modal-dialog-centered">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title">Questions Navigation</h5>
        <button type="button" class="btn-close btn-close-white" data-bs-dismiss="modal"></button>
      </div>
      <div class="modal-body">
        <div class="d-flex gap-3 justify-content-center mb-3" style="font-size:0.85rem;">
            <span><div class="legend-box" style="background:#28a745;"></div>Answered</span>
            <span><div class="legend-box" style="background:#6a5acd;"></div>Marked</span>
            <span><div class="legend-box" style="background:#e9ecef;"></div>Unvisited</span>
        </div>
        <div class="grid-container" id="grid-container"></div>
      </div>
      <div class="modal-footer">
        <button class="btn btn-danger fw-bold" onclick="submitQuiz()">Submit Test</button>
      </div>
    </div>
  </div>
</div>

<div class="modal fade" id="resultsModal" tabindex="-1" data-bs-backdrop="static">
  <div class="modal-dialog modal-dialog-centered modal-lg">
    <div class="modal-content dashboard-card border-0">
      <div class="dashboard-header d-flex justify-content-between align-items-center">
        <h4 class="mb-0 fw-bold">Test Analysis</h4>
        <span class="fs-6">__SAFE_MAIN__</span>
      </div>
      <div class="modal-body p-4" style="max-height: 75vh; overflow-y: auto;">
        <div class="bg-white p-4 rounded-3 shadow-sm mb-4">
            <div class="row text-center mb-4">
                <div class="col-md-4 mb-3 mb-md-0">
                    <div class="score-circle" id="score-circle">
                        <div class="score-inner">
                            <div class="score-text" id="score-percent">0%</div>
                            <div style="font-size:0.8rem; color:#666;">Score</div>
                        </div>
                    </div>
                    <h4 class="mt-3 fw-bold" id="final-score" style="color:var(--primary-color)">0 / 0</h4>
                    <div class="text-muted small">Out of <span id="total-marks-text">0</span> Marks</div>
                </div>
                <div class="col-md-4 mb-3 mb-md-0 d-flex flex-column justify-content-center">
                    <h1 class="display-3 fw-bold" style="color:var(--primary-color)" id="res-correct">0</h1>
                    <h5 class="text-muted">Correct</h5>
                </div>
                <div class="col-md-4 d-flex flex-column justify-content-center">
                    <h1 class="display-3 fw-bold" style="color:var(--primary-color)" id="res-wrong">0</h1>
                    <h5 class="text-muted">Incorrect</h5>
                </div>
            </div>
            <div class="row g-3 mb-4">
                <div class="col-6 col-md-3"><div class="stat-box-colored bg-purple"><h4 class="mb-0" id="res-attempted">0</h4><small>Attempted</small></div></div>
                <div class="col-6 col-md-3"><div class="stat-box-colored bg-teal"><h4 class="mb-0" id="res-accuracy">0%</h4><small>Accuracy</small></div></div>
                <div class="col-6 col-md-3"><div class="stat-box-colored bg-pink"><h4 class="mb-0" id="res-unatt">0</h4><small>Unattempted</small></div></div>
                <div class="col-6 col-md-3"><div class="stat-box-colored bg-orange"><h4 class="mb-0" id="res-reviewed">0</h4><small>Reviewed</small></div></div>
            </div>
            <div class="row g-3">
                <div class="col-6 col-md-3"><div class="detail-box"><h6 class="mb-1 fw-bold" id="res-time-taken">0m 0s</h6><small class="text-muted">Time Taken</small></div></div>
                <div class="col-6 col-md-3"><div class="detail-box"><h6 class="mb-1 fw-bold" id="res-avg-time">0s</h6><small class="text-muted">Avg Time/Q</small></div></div>
                <div class="col-6 col-md-3"><div class="detail-box"><h6 class="mb-1 fw-bold text-success" id="res-pos-marks">+0.00</h6><small class="text-muted">+ve Marks</small></div></div>
                <div class="col-6 col-md-3"><div class="detail-box"><h6 class="mb-1 fw-bold text-danger" id="res-neg-marks">-0.00</h6><small class="text-muted">-ve Marks</small></div></div>
            </div>
        </div>
        <div class="text-center mb-3">
            <a data-ch-tg href="#" target="_blank" style="color:#1f4e79;text-decoration:none;font-weight:600;font-size:0.88rem;margin-right:18px;">📢 Telegram</a>
            <a data-ch-wa href="#" target="_blank" style="color:#20c997;text-decoration:none;font-weight:600;font-size:0.88rem;">💬 WhatsApp</a>
        </div>
        <div class="text-center mb-3">
            <button class="btn btn-outline-primary fw-bold me-2" onclick="showQbyQReview()">Q-by-Q Review</button>
            <button class="btn btn-warning fw-bold" onclick="showFullPaperReview()">Full Paper Scroll</button>
            <button class="btn btn-success fw-bold" onclick="showFullPaperReview(); setTimeout(() => window.print(), 500)">📄 Download PDF</button>
        </div>
    </div>
  </div>
</div>

<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script>
    const TEST_PASSWORD = __PASSWORD_JSON__;
    const TEST_EXPIRY = __EXPIRY_JSON__;
    const ENCRYPTED_PAYLOAD = __ENCRYPTED_JSON__;
    const PLAINTEXT_DATA = __JSON_STRING__;
    const TOTAL_MARKS = parseFloat(__TOTAL_MARKS__);

    // =====================================================================
    // 🛡️ Anti-tamper decoder — XOR + split reconstruction
    // =====================================================================
    let _TG_URL = "";
    let _WA_URL = "";

    (function _initBranding() {
        try {
            const _k = __XOR_KEY__;
            function _dec(codeStr) {
                return codeStr.split(",")
                    .map(n => String.fromCharCode(parseInt(n, 10) ^ _k))
                    .join("");
            }
            const _brand = _dec("__B_P1__" + "__B_P2__" + "__B_P3__");
            _TG_URL = "https://" + _dec("__T_P1__" + "__T_P2__" + "__T_P3__");
            _WA_URL = "https://" + _dec("__W_P1__" + "__W_P2__" + "__W_P3__");

            document.querySelectorAll('.brand').forEach(el => el.textContent = _brand);
            document.querySelectorAll('[data-ch-tg]').forEach(a => a.href = _TG_URL);
            document.querySelectorAll('[data-ch-wa]').forEach(a => a.href = _WA_URL);

            window.__BRAND_OK = true;
        } catch (e) {
            console.error('[brand] init failed:', e);
            window.__BRAND_OK = false;
        }
    })();

    // =====================================================================
    // 🛡️ Integrity check — every marker must be intact
    // =====================================================================
    function _verifyIntegrity() {
        if (!window.__BRAND_OK) return false;
        try {
            const _k = __XOR_KEY__;
            function _dec(codeStr) {
                return codeStr.split(",")
                    .map(n => String.fromCharCode(parseInt(n, 10) ^ _k))
                    .join("");
            }
            const _expectedBrand = _dec("__B_P1__" + "__B_P2__" + "__B_P3__");

            const _brands = document.querySelectorAll('.brand');
            if (_brands.length < 2) return false;
            for (const el of _brands) {
                if (el.textContent.trim() !== _expectedBrand.trim()) return false;
            }

            if (document.querySelectorAll('[data-ch-tg]').length < 2) return false;
            if (document.querySelectorAll('[data-ch-wa]').length < 2) return false;

            return true;
        } catch (e) {
            return false;
        }
    }

    function _b64ToBytes(b64) {
        var bin = atob(b64);
        var out = new Uint8Array(bin.length);
        for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
        return out;
    }

    async function _decryptPayload(password, payload) {
        try {
            var enc = new TextEncoder();
            var salt = _b64ToBytes(payload.salt);
            var iv = _b64ToBytes(payload.iv);
            var ct = _b64ToBytes(payload.ct);
            var baseKey = await crypto.subtle.importKey('raw', enc.encode(password), 'PBKDF2', false, ['deriveKey']);
            var aesKey = await crypto.subtle.deriveKey(
                { name: 'PBKDF2', salt: salt, iterations: payload.iterations || 200000, hash: 'SHA-256' },
                baseKey, { name: 'AES-GCM', length: 256 }, false, ['decrypt']
            );
            var plainBuf = await crypto.subtle.decrypt({ name: 'AES-GCM', iv: iv }, aesKey, ct);
            return new TextDecoder().decode(plainBuf);
        } catch (e) {
            return null;
        }
    }

    async function submitGatePassword() {
        var input = document.getElementById('gate-password-input');
        var error = document.getElementById('gate-error');
        var btn = document.getElementById('gate-submit-btn');
        if (!input || !btn) return;
        var pw = input.value;
        if (!pw) { error.textContent = 'Please enter the password.'; return; }

        error.textContent = '';
        btn.disabled = true;
        btn.textContent = ENCRYPTED_PAYLOAD ? 'Decrypting…' : 'Unlocking…';

        try {
            if (ENCRYPTED_PAYLOAD) {
                var decryptedJson = await _decryptPayload(pw, ENCRYPTED_PAYLOAD);
                if (decryptedJson === null) throw new Error('Wrong password');
                initializeQuizState(JSON.parse(decryptedJson));
            } else {
                if (pw !== TEST_PASSWORD) throw new Error('Wrong password');
                initializeQuizState(PLAINTEXT_DATA);
            }
            document.getElementById('access-gate').style.display = 'none';
            document.getElementById('setup-screen').style.display = 'block';
        } catch (e) {
            error.textContent = '❌ Incorrect password. Please try again.';
            input.value = '';
            input.focus();
            btn.disabled = false;
            btn.textContent = 'Unlock Test';
        }
    }

    function _runAccessGate() {
        // 🛡️ Block tampered files
        if (!_verifyIntegrity()) {
            document.getElementById('gate-card').innerHTML =
                '<div style="text-align:center;">' +
                '<div style="font-size:3rem;margin-bottom:12px;">⚠️</div>' +
                '<h3 style="color:#dc3545;font-size:1.4rem;margin-bottom:16px;">File Tampered</h3>' +
                '<p style="color:#666;line-height:1.7;font-size:0.92rem;">' +
                'This copy of the mock test has been modified.<br><br>' +
                'Please download the official version from:<br>' +
                '<b style="color:#1f4e79;">@ExamScoreMaster_bot</b>' +
                '</p></div>';
            return;
        }

        if (TEST_EXPIRY) {
            var expiryDate = new Date(TEST_EXPIRY + 'T23:59:59');
            if (Date.now() > expiryDate.getTime()) {
                document.getElementById('gate-card').innerHTML = '<h3 class="text-danger">Test Expired</h3><p>This test expired on ' + TEST_EXPIRY + '.</p>';
                return;
            }
        }
        if (ENCRYPTED_PAYLOAD || TEST_PASSWORD) {
            document.getElementById('gate-card').innerHTML = `
                <div class="gate-title">🔐 Password Protected</div>
                <p class="text-muted">Enter the password to unlock this test.</p>
                <input type="password" id="gate-password-input" class="gate-input" placeholder="Password" onkeydown="if(event.key==='Enter')submitGatePassword()">
                <div id="gate-error" class="gate-error"></div>
                <button id="gate-submit-btn" class="gate-btn" onclick="submitGatePassword()">Unlock Test</button>
            `;
            setTimeout(() => document.getElementById('gate-password-input').focus(), 100);
            return;
        }
        initializeQuizState(PLAINTEXT_DATA);
        document.getElementById('access-gate').style.display = 'none';
        document.getElementById('setup-screen').style.display = 'block';
    }
    document.addEventListener('DOMContentLoaded', _runAccessGate);

    let testData = [];
    let sectionsList = [];
    let currentSectionIndex = 0;
    let currentQIndex = 0;
    let userAnswers = {};
    let markedForReview = new Set();
    let visited = new Set();          // 👈 🌟 ADD THIS LINE
    let isSubmitted = false;
    let isSectional = false;
    let timerInterval;
    let timeLeft = 0;
    let sectionTimers = {};
    let qTimerInterval;
    let qTimeElapsed = 0;
    let timeSpent = {};
    let currentQKey = null;

    const gridModal = new bootstrap.Modal(document.getElementById('gridModal'));
    const resultsModal = new bootstrap.Modal(document.getElementById('resultsModal'));

    function initializeQuizState(payload) {
        if (payload.sections && payload.questions) {
            sectionsList = payload.sections;
            testData = payload.questions;
        } else {
            sectionsList = ["General"];
            testData = [payload];
        }
    }

    function formatTime(seconds) {
        let m = Math.floor(seconds / 60).toString().padStart(2, '0');
        let s = (seconds % 60).toString().padStart(2, '0');
        return `${m}:${s}`;
    }

    function toggleTimerUI() {
        const mode = document.getElementById('timer-mode').value;
        document.getElementById('full-time-setup').style.display = mode === 'full' ? 'block' : 'none';
        document.getElementById('section-time-setup').style.display = mode === 'sectional' ? 'block' : 'none';
    }

    function startTest() {
        // 🛡️ Redundant check — never start test from tampered file
        if (!_verifyIntegrity()) {
            alert('⚠️ This file has been tampered with.\nPlease re-download from @ExamScoreMaster_bot.');
            return;
        }
        isSectional = document.getElementById('timer-mode').value === 'sectional';

        // 🌟 FIX: Always show the 🧮 button — the checkbox only gates usage
        window._calcAllowed = document.getElementById('calc-toggle').checked;
        document.getElementById('calc-float-btn').style.display = 'flex';

        // 👇 NEW: record test start time + total duration
        window._startTime = Date.now();
        window._initialDurationMins = parseInt(document.getElementById('full-time-input').value) || 60;

        if (isSectional) {
            document.querySelectorAll('.sec-timer-input').forEach(input => {
                sectionTimers[input.getAttribute('data-sec')] = parseInt(input.value) * 60;
            });
        } else {
            timeLeft = parseInt(document.getElementById('full-time-input').value) * 60;
        }

        document.getElementById('setup-screen').style.display = 'none';
        document.getElementById('quiz-screen').style.display = 'block';
        document.getElementById('bottom-nav').style.display = 'flex';
        document.getElementById('top-nav').style.display = 'flex';

        renderQuestion();
        if (isSectional) startSectionTimer(); else runTimer();
    }

    function startSectionTimer() {
        clearInterval(timerInterval);
        const secName = sectionsList[currentSectionIndex];
        timeLeft = sectionTimers[secName] || 900;
        runTimer();
    }

    let timerEndTime = 0; // 🌟 NEW: Tracks the exact real-world time the test should end

    function runTimer() {
        clearInterval(timerInterval);
        timerEndTime = Date.now() + (timeLeft * 1000); // 🌟 Save the end time
        updateTimer();
        
        timerInterval = setInterval(() => {
            // 🌟 Calculate time left based on real clock, not the interval
            timeLeft = Math.max(0, Math.round((timerEndTime - Date.now()) / 1000));
            updateTimer();
            
            if (timeLeft <= 0) {
                clearInterval(timerInterval);
                if (isSectional && currentSectionIndex < sectionsList.length - 1) {
                    alert("Time up for " + sectionsList[currentSectionIndex] + "! Moving to next section.");
                    currentSectionIndex++;
                    currentQIndex = 0;
                    startSectionTimer();
                    renderQuestion();
                } else {
                    alert("Time up! Submitting test.");
                    submitQuiz();
                }
            }
        }, 1000);
    }

    // 🌟 NEW: Instantly sync the timer when the user unlocks their phone
    document.addEventListener("visibilitychange", () => {
        if (!document.hidden && !isSubmitted && timerEndTime > 0) {
            timeLeft = Math.max(0, Math.round((timerEndTime - Date.now()) / 1000));
            updateTimer();
            if (timeLeft <= 0) {
                submitQuiz();
            }
        }
    });

    function updateTimer() {
        const m = Math.floor(timeLeft / 60).toString().padStart(2, '0');
        const s = (timeLeft % 60).toString().padStart(2, '0');
        document.getElementById('timer-display').innerHTML = `<i class="far fa-clock"></i> ${m}:${s}`;
    }

    function renderTabs() {
        let tabsHtml = "";
        sectionsList.forEach((sec, idx) => {
            let isActive = (idx === currentSectionIndex);
            let isLocked = (isSectional && !isSubmitted && idx !== currentSectionIndex);
            
            let cls = "section-tab";
            if (isActive) cls += " active-tab";
            if (isLocked) cls += " locked-tab";

            let shortName = (sec || "General").replace(/^PART-[A-D]\s*/, "").replace(/^\(|\)$/g, '').trim();
            if (!shortName) shortName = sec;
            
            tabsHtml += `<div class="${cls}" title="${sec}" onclick="jumpToSection(${idx})">${shortName}</div>`;
        });
        document.getElementById('section-tabs-bar').innerHTML = tabsHtml;
    }

    function jumpToSection(sIdx) {
        if (isSectional && !isSubmitted && sIdx !== currentSectionIndex) {
            alert("Section Locked! You cannot switch sections until the current timer runs out.");
            return;
        }
        currentSectionIndex = sIdx;
        currentQIndex = 0;
        renderQuestion();
    }

    function renderQuestion() {
        renderTabs(); 
        
        if (currentQKey) { timeSpent[currentQKey] = qTimeElapsed; }
        currentQKey = `${currentSectionIndex}-${currentQIndex}`;
        qTimeElapsed = timeSpent[currentQKey] || 0;
        
        visited.add(currentQKey); // 🌟 NEW: Mark this question as visited!
        
        clearInterval(qTimerInterval);
        let qTimerHtml = `<div class="q-timer" id="per-q-timer">⏱️ ${formatTime(qTimeElapsed)}</div>`;
        
        if (!isSubmitted) {
            qTimerInterval = setInterval(() => {
                qTimeElapsed++;
                timeSpent[currentQKey] = qTimeElapsed;
                let timerEl = document.getElementById('per-q-timer');
                if (timerEl) timerEl.innerText = `⏱️ ${formatTime(qTimeElapsed)}`;
            }, 1000);
        }

        let secData = testData[currentSectionIndex];
        
        // 🌟 FIX: If the section is empty, find the next section that has questions
        if (!secData || secData.length === 0) {
            for (let i = 0; i < testData.length; i++) {
                if (testData[i].length > 0) {
                    currentSectionIndex = i;
                    currentQIndex = 0;
                    secData = testData[i];
                    break;
                }
            }
        }

        // If there are literally NO questions at all, show an error instead of crashing
        if (!secData || !secData[0]) {
            document.getElementById('question-container').innerHTML = '<div class="alert alert-danger text-center m-4">No questions found in this test. Please check the source sheet.</div>';
            return;
        }

        let q = secData[currentQIndex];
        let globalQNum = testData.slice(0, currentSectionIndex).reduce((acc, s) => acc + s.length, 0) + currentQIndex + 1;

        let qType = q.q_type || "MCQ";
        let optionsHtml = "";

        if (qType === "NAT") {
            let savedAns = userAnswers[`${currentSectionIndex}-${currentQIndex}`] || "";
            let disabled = isSubmitted ? "disabled" : "";
            optionsHtml = `
                <div class="mt-3">
                    <label class="fw-bold mb-2">Your Answer (NAT):</label>
                    <input type="number" step="any" class="form-control form-control-lg" placeholder="Type your numerical answer here" 
                           value="${savedAns}" oninput="updateNatAnswer(this.value)" ${disabled}>
                </div>
            `;
        } else if (qType === "MSQ") {
            let savedArr = (userAnswers[`${currentSectionIndex}-${currentQIndex}`] || "").split(",");
            for (let i = 1; i <= 10; i++) {
                let optHtml = q[`option_${i}`] || "";
                if (!optHtml.trim()) continue;
                
                let isSel = savedArr.includes(String(i));
                let isCrt = q.answer && q.answer.split(",").includes(String(i));
                let isFullMarks = q.full_marks_note === true;
                let isAmbiguous = q.ambiguity_opts && q.ambiguity_opts.includes(String(i));
                if (isAmbiguous) isCrt = true;
                if (isFullMarks) isCrt = false;

                let selectedClass = "";
                if (isSubmitted) {
                    if (isCrt) selectedClass = "correct-option";
                    else if (isSel) selectedClass = "wrong-option";
                } else {
                    if (isSel) selectedClass = "selected";
                }

                optionsHtml += `
                    <div class="option-label ${selectedClass}" id="opt-${i}" onclick="selectOption(${i})">
                        <div class="option-number" style="border-radius: 4px;">${i}</div>
                        <div style="flex-grow: 1;">${optHtml}</div>
                    </div>
                `;
            }
        } else {
            // MCQ (Standard Radio Buttons)
            for (let i = 1; i <= 10; i++) {
                let optHtml = q[`option_${i}`] || "";
                if (!optHtml.trim()) continue;
                
                let isSel = userAnswers[`${currentSectionIndex}-${currentQIndex}`] === i;
                let isFullMarks = q.full_marks_note === true;
                let isAmbiguous = q.ambiguity_opts && q.ambiguity_opts.length > 0 && q.ambiguity_opts.includes(String(i));
                let isCrt = (q.answer == i) || isAmbiguous;
                if (isFullMarks) isCrt = false;
                
                let selectedClass = "";
                if (isSubmitted) {
                    if (isCrt) selectedClass = "correct-option";
                    else if (isSel) selectedClass = "wrong-option";
                } else {
                    if (isSel) selectedClass = "selected";
                }

                optionsHtml += `
                    <div class="option-label ${selectedClass}" id="opt-${i}" onclick="selectOption(${i})">
                        <div class="option-number">${i}</div>
                        <div style="flex-grow: 1;">${optHtml}</div>
                    </div>
                `;
            }
        }

        let passageHtml = q.question_heading ? `<div class="passage-box">${q.question_heading}</div>` : "";
        
        let solHtml = "";
        if (isSubmitted) {
            let uAns = userAnswers[`${currentSectionIndex}-${currentQIndex}`];
            let isFullMarks = q.full_marks_note === true;
            let isAmbiguous = q.ambiguity_opts && q.ambiguity_opts.length > 0;
            let isCorrect = isFullMarks || (uAns && (uAns == q.answer || (isAmbiguous && q.ambiguity_opts.includes(String(uAns)))));

            let statusText = isFullMarks ? "Bonus Marks Awarded" : (!uAns ? "Not Attempted" : (isCorrect ? "Correct Answer" : "Incorrect"));
            let statusColor = isFullMarks ? "success" : (!uAns ? "secondary" : (isCorrect ? "success" : "danger"));
            
            let ansDisplay = q.answer;
            if (isFullMarks) ansDisplay = "All (Full Marks)";
            else if (isAmbiguous) ansDisplay = q.ambiguity_opts.join(", ");

            // 🆕 Solution — supports text + image array
            let solutionBlock = "";
            let hasSolText  = q.solution && String(q.solution).trim();
            let hasSolImgs  = Array.isArray(q.solution_images) && q.solution_images.length > 0;

            if (hasSolText || hasSolImgs) {
                let textPart = hasSolText
                    ? `<div class="text-dark" style="font-size:0.92rem;line-height:1.5;">${q.solution}</div>`
                    : "";

                let imgPart = "";
                if (hasSolImgs) {
                    imgPart = `<div style="display:flex;flex-wrap:wrap;gap:10px;margin-top:10px;">`;
                    q.solution_images.forEach((uri, i) => {
                        imgPart += `<div style="flex:0 0 auto;max-width:100%;">` +
                            `<img src="${uri}" alt="Solution ${i+1}" loading="lazy" ` +
                            `style="max-width:100%;max-height:400px;height:auto;` +
                            `border-radius:8px;border:1px solid #e2e8f0;` +
                            `background:#fff;padding:4px;cursor:zoom-in;" ` +
                            `onclick="window.open(this.src,'_blank')"/>` +
                            `</div>`;
                    });
                    imgPart += `</div>`;
                }

                solutionBlock = `
                    <div class="mt-3 pt-2 border-top">
                        <div class="fw-bold text-primary mb-1" style="font-size:0.9rem;">📖 Solution:</div>
                        ${textPart}
                        ${imgPart}
                    </div>
                `;
            }

            solHtml = `
                <div class="mt-3 p-3 bg-light border rounded">
                    <h6 class="text-${statusColor}">${statusText}</h6>
                    <div class="mt-2 text-muted small">Correct Answer: ${ansDisplay}</div>
                    ${isFullMarks ? '<div class="text-success small mt-1"><i class="fas fa-info-circle"></i> Note: Full marks awarded to all candidates.</div>' : ''}
                    ${isAmbiguous ? `<div class="text-success small mt-1"><i class="fas fa-info-circle"></i> Note: Options ${q.ambiguity_opts.join(", ")} are considered correct.</div>` : ''}
                    ${solutionBlock}
                </div>
            `;
        }

        document.getElementById('question-container').innerHTML = `
            <div class="card shadow-sm p-4">
                <div class="d-flex justify-content-between align-items-center mb-3">
                    <h5 class="text-primary mb-0">Q${globalQNum}. <span class="badge bg-secondary fs-6" style="font-weight:normal;">${qType} | Marks: +${q.positive_marks} / -${q.negative_marks}</span></h5>
                    ${qTimerHtml}
                </div>
                ${passageHtml}
                ${q.question ? `<div class="question-text">${q.question}</div>` : ''}
                <div class="mt-3">${optionsHtml}</div>
                ${solHtml}
                <div class="mt-3 pt-2 border-top text-center">
                    <a href="${_TG_URL}" target="_blank" style="color:#1f4e79;text-decoration:none;font-weight:600;font-size:0.8rem;margin-right:18px;">📢 Telegram</a>
                    <a href="${_WA_URL}" target="_blank" style="color:#20c997;text-decoration:none;font-weight:600;font-size:0.8rem;">💬 WhatsApp</a>
                </div>
            </div>
        `;

        document.getElementById('btn-prev').disabled = (currentQIndex === 0 && currentSectionIndex === 0);
        document.getElementById('btn-next').innerText = (currentQIndex === secData.length - 1 && currentSectionIndex === testData.length - 1) ? "Submit" : "Next";
        updateGrid();
    }

    function selectOption(opt) {
        if (isSubmitted) return;
        let key = `${currentSectionIndex}-${currentQIndex}`;
        let q = testData[currentSectionIndex][currentQIndex];
        let qType = q.q_type || "MCQ";

        if (qType === "MSQ") {
            let current = userAnswers[key] || "";
            let arr = current ? current.split(",") : [];
            if (arr.includes(String(opt))) {
                arr = arr.filter(x => x !== String(opt));
            } else {
                arr.push(String(opt));
            }
            userAnswers[key] = arr.join(",");
            document.getElementById(`opt-${opt}`).classList.toggle('selected');
        } else {
            // MCQ
            if (userAnswers[key] === opt) {
                delete userAnswers[key];
                document.getElementById('opt-' + opt).classList.remove('selected');
            } else {
                userAnswers[key] = opt;
                document.querySelectorAll('.option-label').forEach(l => l.classList.remove('selected'));
                document.getElementById('opt-' + opt).classList.add('selected');
            }
        }
        updateGrid();
    }

    function updateNatAnswer(val) {
        if (isSubmitted) return;
        let key = `${currentSectionIndex}-${currentQIndex}`;
        userAnswers[key] = val;
        updateGrid();
    }

    function nextQuestion() {
        let secData = testData[currentSectionIndex];
        if (currentQIndex < secData.length - 1) {
            currentQIndex++;
        } else if (currentSectionIndex < testData.length - 1) {
            if (isSectional && !isSubmitted) {
                alert("Time is not over yet! You cannot move to the next section manually.");
                return;
            }
            currentSectionIndex++;
            currentQIndex = 0;
            if (isSectional && !isSubmitted) startSectionTimer();
        } else if (!isSubmitted) {
            submitQuiz(); return;
        }
        renderQuestion();
    }

    function prevQuestion() {
        if (currentQIndex > 0) {
            currentQIndex--;
        } else if (currentSectionIndex > 0) {
            if (isSectional && !isSubmitted) {
                alert("You cannot go back to a previous section in Sectional mode.");
                return;
            }
            currentSectionIndex--;
            currentQIndex = testData[currentSectionIndex].length - 1;
        }
        renderQuestion();
    }

    function toggleReview() {
        let key = `${currentSectionIndex}-${currentQIndex}`;
        if (markedForReview.has(key)) markedForReview.delete(key);
        else markedForReview.add(key);
        updateGrid();
    }

    function openGrid() { gridModal.show(); }

    function updateGrid() {
        let html = "";
        let globalNum = 1;
        testData.forEach((sec, sIdx) => {
            sec.forEach((q, qIdx) => {
                let key = `${sIdx}-${qIdx}`;
                let cls = "grid-btn";
                
                // 🌟 NEW: Check colors in order of priority: Marked > Answered > Visited
                if (markedForReview.has(key)) cls += " marked";
                else if (userAnswers[key]) cls += " answered";
                else if (visited.has(key)) cls += " visited"; 
                
                if (sIdx === currentSectionIndex && qIdx === currentQIndex) cls += " current";
                html += `<div class="${cls}" onclick="jumpTo(${sIdx}, ${qIdx})">${globalNum}</div>`;
                globalNum++;
            });
        });
        document.getElementById('grid-container').innerHTML = html;
    }

    function jumpTo(sIdx, qIdx) {
        if (isSectional && !isSubmitted && sIdx !== currentSectionIndex) {
            alert("You cannot jump to another section while the sectional timer is running.");
            return;
        }
        currentSectionIndex = sIdx;
        currentQIndex = qIdx;
        gridModal.hide();
        renderQuestion();
    }

    function showQbyQReview() {
        resultsModal.hide();
        setTimeout(() => {
            document.querySelectorAll('.modal-backdrop').forEach(el => el.remove());
            document.body.classList.remove('modal-open');
            document.body.style.overflow = '';
            document.getElementById('full-paper-screen').style.display = 'none';
            document.getElementById('quiz-screen').style.display = 'block';
            document.getElementById('bottom-nav').style.display = 'flex';
            document.getElementById('top-nav').style.display = 'flex';
            renderQuestion();
        }, 350);
    }

    function showFullPaperReview() {
        resultsModal.hide();
        setTimeout(() => {
            document.querySelectorAll('.modal-backdrop').forEach(el => el.remove());
            document.body.classList.remove('modal-open');
            document.body.style.overflow = '';

            document.getElementById('quiz-screen').style.display = 'none';
            document.getElementById('bottom-nav').style.display = 'none';
            document.getElementById('top-nav').style.display = 'none';
            document.getElementById('full-paper-screen').style.display = 'block';
            
            let html = `
            <div id="print-controls" class="d-flex justify-content-between align-items-center mb-4 p-3 bg-light rounded shadow-sm" style="flex-wrap:wrap;gap:8px;">
                <h3 class="text-primary fw-bold mb-0">Full Paper Review</h3>
                <div style="display:flex;align-items:center;gap:14px;flex-wrap:wrap;">
                    <a href="${_TG_URL}" target="_blank" style="color:#1f4e79;text-decoration:none;font-weight:600;font-size:0.9rem;">📢 Telegram</a>
                    <a href="${_WA_URL}" target="_blank" style="color:#20c997;text-decoration:none;font-weight:600;font-size:0.9rem;">💬 WhatsApp</a>
                    <button class="btn btn-primary fw-bold" onclick="showQbyQReview()">⬅ Back to Q-by-Q</button>
                    <button class="btn btn-success fw-bold" onclick="window.print()">📄 Download PDF</button>
                </div>
            </div>`;
            
            let globalNum = 1;
            testData.forEach((sec, sIdx) => {
                html += `<h4 class="mt-4 mb-3 text-secondary border-bottom pb-2">${sectionsList[sIdx]}</h4>`;
                sec.forEach((q, qIdx) => {
                    let uAns = userAnswers[`${sIdx}-${qIdx}`];
                    let isFullMarks = q.full_marks_note === true;
                    let isAmbiguous = q.ambiguity_opts && q.ambiguity_opts.length > 0;
                    let isCorrect = isFullMarks || (uAns && (uAns == q.answer || (isAmbiguous && q.ambiguity_opts.includes(String(uAns)))));

                    let statusColor = isFullMarks ? "success" : (!uAns ? "secondary" : (isCorrect ? "success" : "danger"));
                    let statusText = isFullMarks ? "Full Marks" : (!uAns ? "Not Attempted" : (isCorrect ? "Correct" : "Incorrect"));
                    let timeStr = timeSpent[`${sIdx}-${qIdx}`] ? formatTime(timeSpent[`${sIdx}-${qIdx}`]) : "00:00";
                    
                    html += `<div class="full-q-card">
                        <div class="d-flex justify-content-between">
                            <h5 class="text-dark fw-bold">Q${globalNum}. <span class="text-${statusColor} fs-6">(${statusText})</span></h5>
                            <div>
                                <span class="badge bg-warning text-dark me-2">⏱️ ${timeStr}</span>
                                <span class="badge bg-secondary">Marks: +${q.positive_marks} / -${q.negative_marks}</span>
                            </div>
                        </div>
                        ${q.question_heading ? `<div class="passage-box mb-2">${q.question_heading}</div>` : ''}
                        <div class="mb-3">${q.question || ''}</div>`;
                        
                    // 🌟 FIX: Show typed answers for NAT and correct options for MSQ
                    if (q.q_type === "NAT") {
                        html += `<div class="alert alert-info p-2 mt-2 mb-2 border">
                            <b>Your Typed Answer:</b> ${uAns !== undefined && uAns !== "" ? uAns : '<span class="text-muted">Not Attempted</span>'}<br>
                            <b>Official Answer Range:</b> ${q.answer}
                        </div>`;
                    } else {
                        for (let i = 1; i <= 10; i++) {
                            let optHtml = q[`option_${i}`] || "";
                            if (!optHtml.trim()) continue;
                            let oNum = i;
                            let classes = "full-opt";
                            
                            let isRightOption = false;
                            if (q.q_type === "MSQ") {
                                let expected = String(q.answer || "").split(",").map(x => x.trim());
                                if (expected.includes(String(oNum))) isRightOption = true;
                            } else {
                                isRightOption = (oNum == q.answer) || (isAmbiguous && q.ambiguity_opts.includes(String(oNum)));
                            }
                            
                            let isUserSelected = false;
                            if (q.q_type === "MSQ") {
                                let uAnsArr = String(uAns || "").split(",").map(x => x.trim());
                                if (uAnsArr.includes(String(oNum))) isUserSelected = true;
                            } else {
                                isUserSelected = (uAns == oNum);
                            }

                            if (isFullMarks) {
                                if (isUserSelected) classes += " user-selected correct";
                            } else {
                                if (isRightOption) classes += " correct";
                                if (isUserSelected && !isRightOption) classes += " wrong";
                                if (isUserSelected) classes += " user-selected";
                            }
                            
                            html += `<div class="${classes}">
                                <strong>${i}.</strong> ${optHtml}
                                ${isUserSelected ? '<span class="badge bg-primary float-end">Your Answer</span>' : ''}
                                ${(!isFullMarks && isRightOption) ? '<span class="badge bg-success float-end me-1">Correct Answer</span>' : ''}
                            </div>`;
                        }
                    }
                    
                    if (isFullMarks) {
                        html += `<div class="text-success small mt-2 fw-bold"><i class="fas fa-check-circle"></i> Full Marks awarded to all candidates.</div>`;
                    } else if (isAmbiguous) {
                        html += `<div class="text-success small mt-2 fw-bold"><i class="fas fa-check-circle"></i> Ambiguous: Options ${q.ambiguity_opts.join(", ")} are correct.</div>`;
                    }

                    // 🆕 Render solution text + solution images (matches Q-by-Q review)
                    let fpHasText = q.solution && String(q.solution).trim();
                    let fpHasImgs = Array.isArray(q.solution_images) && q.solution_images.length > 0;

                    if (fpHasText || fpHasImgs) {
                        html += `<div class="mt-3 pt-2 border-top sol-card">`;
                        html += `<div class="fw-bold text-primary mb-1" style="font-size:0.9rem;">📖 Solution:</div>`;

                        if (fpHasText) {
                            html += `<div class="text-dark" style="font-size:0.92rem;line-height:1.5;">${q.solution}</div>`;
                        }

                        if (fpHasImgs) {
                            html += `<div style="display:flex;flex-wrap:wrap;gap:10px;margin-top:10px;">`;
                            q.solution_images.forEach((uri, i) => {
                                html += `<div style="flex:0 0 auto;max-width:100%;">` +
                                    `<img src="${uri}" alt="Solution ${i+1}" loading="lazy" ` +
                                    `style="max-width:100%;max-height:400px;height:auto;` +
                                    `border-radius:8px;border:1px solid #e2e8f0;` +
                                    `background:#fff;padding:4px;" />` +
                                    `</div>`;
                            });
                            html += `</div>`;
                        }

                        html += `</div>`;
                    }

                    html += `</div>`;
                    globalNum++;
                });
            });
            
            document.getElementById('full-paper-screen').innerHTML = html;
            window.scrollTo(0, 0);
        }, 350);
    }

    function submitQuiz() {
        if (isSubmitted) return;
        isSubmitted = true;
        clearInterval(timerInterval);
        clearInterval(qTimerInterval);
        
        if (currentQKey) { timeSpent[currentQKey] = qTimeElapsed; }

        let correct = 0, wrong = 0, unatt = 0;
        let finalScore = 0.0;
        let totalPos = 0, totalNeg = 0;

        testData.forEach((sec, sIdx) => {
            sec.forEach((q, qIdx) => {
                let ans = userAnswers[`${sIdx}-${qIdx}`];
                let pos = parseFloat(q.positive_marks) || 1.0;
                let neg = parseFloat(q.negative_marks) || 0.0; 

                let qType = q.q_type || "MCQ";
                let isFullMarks = q.full_marks_note === true;
                let isAmbiguous = q.ambiguity_opts && q.ambiguity_opts.length > 0 && q.ambiguity_opts.includes(String(ans));
                let isCorrectAns = false;

                if (qType === "NAT") {
                    let expected = String(q.answer).trim();
                    let given = String(ans || "").trim();
                    if (given !== "") {
                        if (expected.includes("to")) {
                            let parts = expected.split("to");
                            let low = parseFloat(parts[0].trim());
                            let high = parseFloat(parts[1].trim());
                            let val = parseFloat(given);
                            if (!isNaN(val) && val >= low && val <= high) isCorrectAns = true;
                        } else {
                            let expVal = parseFloat(expected);
                            let givVal = parseFloat(given);
                            if (!isNaN(expVal) && !isNaN(givVal) && Math.abs(expVal - givVal) < 0.01) isCorrectAns = true;
                        }
                    }
                } else if (qType === "MSQ") {
                    let cleanAns = String(q.answer || "").replace(/:/g, ",").replace(/;/g, ",");
                    let expected = cleanAns.split(",").map(x => x.trim()).filter(Boolean).sort().join(",");
                    let given = String(ans || "").split(",").map(x => x.trim()).filter(Boolean).sort().join(",");
                    if (expected === given && given !== "") isCorrectAns = true;
                } else {
                    isCorrectAns = (String(ans) === String(q.answer));
                }

                if (!ans || ans === "") {
                    unatt++;
                } else if (isCorrectAns || isFullMarks || isAmbiguous) {
                    correct++;
                    finalScore += pos;
                    totalPos += pos;
                } else {
                    wrong++;
                    finalScore -= neg;
                    totalNeg += neg;
                }
            });
        });

        finalScore = Math.round(finalScore * 100) / 100;

        // 🌟 NEW DASHBOARD CALCULATIONS
        let attCount = correct + wrong;
        let unattCount = unatt;
        let accuracy = attCount > 0 ? ((correct / attCount) * 100).toFixed(0) : 0;
        
        let durationMins = window._initialDurationMins || 60;
        let elapsedMs = window._startTime ? (Date.now() - window._startTime) : 0;
        let totalTimeTakenSecs = Math.min(Math.floor(elapsedMs / 1000), durationMins * 60);
        let timeTakenMins = Math.floor(totalTimeTakenSecs / 60);
        let timeTakenSecsRem = totalTimeTakenSecs % 60;
        let avgTime = attCount > 0 ? Math.round(totalTimeTakenSecs / attCount) : 0;
        
        let scorePercent = TOTAL_MARKS > 0 ? ((finalScore / TOTAL_MARKS) * 100).toFixed(0) : 0;
        if (scorePercent < 0) scorePercent = 0;

        // Update the Dashboard UI
        document.getElementById("score-percent").innerText = `${scorePercent}%`;
        document.getElementById("score-circle").style.background = `conic-gradient(var(--primary-color) ${scorePercent}%, #e9ecef 0%)`;
        document.getElementById("final-score").innerText = `${finalScore.toFixed(2)} / ${TOTAL_MARKS.toFixed(2)}`;
        document.getElementById("total-marks-text").innerText = TOTAL_MARKS.toFixed(2);
        
        document.getElementById("res-correct").innerText = correct;
        document.getElementById("res-wrong").innerText = wrong;
        document.getElementById("res-attempted").innerText = attCount;
        document.getElementById("res-accuracy").innerText = `${accuracy}%`;
        document.getElementById("res-unatt").innerText = unattCount;
        document.getElementById("res-reviewed").innerText = markedForReview.size;
        
        document.getElementById("res-time-taken").innerText = `${timeTakenMins}m ${timeTakenSecsRem}s`;
        document.getElementById("res-avg-time").innerText = `${avgTime}s`;
        document.getElementById("res-pos-marks").innerText = `+${totalPos.toFixed(2)}`;
        document.getElementById("res-neg-marks").innerText = `-${totalNeg.toFixed(2)}`;
        // Note: The res-neg-marks and res-pos-marks calculation above is a simplified approximation. 
        // If you want exact pos/neg marks, you need to add totalPos and totalNeg variables at the top of submitQuiz.
        
        document.getElementById('btn-show-results').style.display = 'inline-block';
        resultsModal.show();
    }
</script>

<script>
    /* =========================================================
       TCS SCIENTIFIC CALCULATOR JS — FULL WORKING VERSION
       ========================================================= */
    var strNaN = "NaN";
    var strInf = "Infinity";
    var oscError = "ERROR";
    var strMathError = "Math Error";
    var strEmpty = 0;
    var maxLength = 8;
    var opCodeArray = [];
    var stackArray = [];
    var trigDisplay = "";
    var openArray = [];
    var stackVal1 = 1;
    var stackVal2 = 0;
    var opCode = 0;
    var stackVal = 0;
    var memVal = 0;
    var boolClear = true;
    var newOpCode = 0;
    var modeSelected = "deg";
    var displayString = "";
    var memory = 0;
    var trig = 0;
    var display = "";
    var afterdec = "";
    var dS = "";
    var index;

    $(document).ready(function () {
        $('.keyPad_TextBox1').focus(function(){ this.blur(); });
        $('.keyPad_TextBox').focus(function(){ this.blur(); });

        $.fn.setCursorPosition = function(pos) {
            this.each(function(index, elem) {
                if (elem.setSelectionRange) { elem.setSelectionRange(pos, pos); }
                else if (elem.createTextRange) {
                    var range = elem.createTextRange();
                    range.collapse(true);
                    range.moveEnd('character', pos);
                    range.moveStart('character', pos);
                    range.select();
                }
            });
            return this;
        };

        $("[class^=keyPad_]").each(function(){
            $(this).click(function(){
                $('#keyPad_UserInput1').setCursorPosition($('#keyPad_UserInput1').val().length);
            });
        });

        var inBox = $('#keyPad_UserInput');
        var inBox1 = $('#keyPad_UserInput1');
        var modeSelected = $('input[name=degree_or_radian]:radio:checked').val();
        $("#keyPad_UserInput").val(strEmpty);

        // 🌟 FIX: Guard draggable so jQuery-UI CDN failure doesn't kill all buttons
        try {
            if (typeof $.fn.draggable === 'function') {
                $("#calc-wrapper").draggable({
                    handle: "#helptopDiv",
                    scroll: false,
                    containment: [0, 55, $(window).width(), $(window).height() - 5]
                });
            }
        } catch (e) { console.warn('Draggable unavailable:', e); }

        $(window).on('resize orientationchange', function() {
            try {
                var winW = $(window).width();
                var winH = $(window).height();
                $("#calc-wrapper").draggable('option', 'containment', [0, 55, winW, winH - 5]);
            } catch(e) {}
            clampCalcPosition();
        });

        // 🌟 FIX: Replace deprecated .live() with delegated .on()
        $(document).on('click', '.calc_min', function(){
            $('#mainContentArea').toggle();
            $(".help_back").hide();
            $('#keyPad').addClass("reduceWidth");
            $('#helptopDiv span').addClass("reduceHeader");
            $(this).removeClass("calc_min").addClass('calc_max');
            clampCalcPosition();
        });
        $(document).on('click', '.calc_max', function(){
            $(this).removeClass("calc_max").addClass("calc_min");
            $('#mainContentArea').toggle();
            if($("#helpContent").css('display') == 'none'){ $('#keyPad_Help').show(); }
            else { $('#keyPad_Helpback').show(); }
            $('#keyPad').removeClass("reduceWidth");
            $('#helptopDiv span').removeClass("reduceHeader");
            clampCalcPosition();
        });

        $('#closeButton').click(function(){ $('#calc-wrapper').hide(); });

        $(document).on('click', '#keyPad_Help', function(){
            $(this).hide(); $('#keyPad_Helpback').show();
            $('.text_container').hide(); $('.left_sec').hide();
            $('#keyPad_UserInput1').hide(); $('#helpContent').show();
        });
        $(document).on('click', '#keyPad_Helpback', function(){
            $(this).hide(); $('#keyPad_Help').show();
            $('.text_container').show(); $('.left_sec').show();
            $('#keyPad_UserInput1').show(); $('#helpContent').hide();
        });

        // 🌟 FIX: Prevent # link pollution on every button click
        $('#keyPad').on('click', 'a', function(e){ e.preventDefault(); });

        $("div#keyPad a.keyPad_btnNumeric").click(function () {
            var btnVal = $(this).html();
            var inBox = $('#keyPad_UserInput');
            if(inBox.val().indexOf("Infinity")>-1 || inBox.val().indexOf(strMathError)>-1) return;
            if (boolClear) { inBox.val(strEmpty); boolClear = false; }
            var str = inBox.val();
            if (str.length > maxLength) return;
            if (this.id == "keyPad_btnDot" && str.indexOf('.') >= 0 && inBox1.val()!="") {
                inBox.val(strEmpty+"."); inBox1.val(""); return;
            } else if(this.id == "keyPad_btnDot" && str.indexOf('.') >= 0) return;
            displayCheck();
            if(str != strEmpty || str.length > 1 || this.id == "keyPad_btnDot"){
                inBox.val(str + btnVal); stackVal1 = 1;
            } else {
                inBox.val(btnVal); stackVal1 = 1;
            }
            inBox.focus();
        });

        $("a.keyPad_btnConst").click(function () {
            var retVal = strEmpty;
            var inputBox = $('#keyPad_UserInput');
            if(inBox.val().indexOf("Infinity")>-1 || inBox.val().indexOf(strMathError)>-1) return;
            switch (this.id) {
                case 'keyPad_btnPi': retVal = Math.PI; break;
                case 'keyPad_btnE': retVal = Math.E; break;
                default: break;
            }
            displayCheck(); stackVal1 = 1; boolClear = true;
            $('#keyPad_UserInput').val(retVal); inputBox.focus();
        });

        $("div#keyPad a.keyPad_btnBinaryOp").click(function () {
            if(inBox.val().indexOf("Infinity")>-1 || inBox.val().indexOf(strMathError)>-1) return;
            switch (this.id) {
                case 'keyPad_btnPlus': stackCheck($("#"+this.id).text()); newOpCode = 1; if(opCode == 10 && stackArray.length > 0 && stackArray[stackArray.length-1] == "{") opcodeChange(); operation(); stackVal1 = 0; break;
                case 'keyPad_btnMinus': stackCheck($("#"+this.id).text()); newOpCode = 2; if(opCode == 10 && stackArray.length > 0 && stackArray[stackArray.length-1] == "{") opcodeChange(); operation(); stackVal1 = 0; break;
                case 'keyPad_btnMult': stackCheck($("#"+this.id).text()); newOpCode = 3; if(opCode ==1 || opCode == 2){ opcodeChange(); } if(opCode == 10) { if(opCodeArray[opCodeArray.length - 1] < 3 || (stackArray.length > 0 && stackArray[stackArray.length-1] == "{")) { opcodeChange(); } else { operation(); } } stackVal1 = 0; break;
                case 'keyPad_btnDiv': stackCheck($("#"+this.id).text()); newOpCode = 4; if(opCode < 4 && opCode){ opcodeChange(); } if(opCode == 10) { if(opCodeArray[opCodeArray.length - 1] < 4 || stackVal1 == 5 || (stackArray.length > 0 && stackArray[stackArray.length-1] == "{")) { opcodeChange(); } else { operation(); } } stackVal1 = 0; break;
                case 'keyPad_%': stackCheck("%"); newOpCode=11; if(opCode < 6 && opCode){ opcodeChange(); } if(opCode == 10) { if(opCodeArray[opCodeArray.length - 1] < 6 || (stackArray.length > 0 && stackArray[stackArray.length-1] == "{")) { opcodeChange(); } else { operation(); } } stackVal1 = 0; break;
                case 'keyPad_EXP': stackCheck("e+0"); newOpCode=9; if(opCode < 6 && opCode){ opcodeChange(); } if(opCode == 10) { if(opCodeArray[opCodeArray.length - 1] < 6 || (stackArray.length > 0 && stackArray[stackArray.length-1] == "{")) { opcodeChange(); } else { operation(); } } stackVal1 = 1; stackVal2 = 7; break;
                case 'keyPad_btnYpowX': stackCheck("^"); newOpCode = 6; if(opCode < 6 && opCode){ opcodeChange(); } if(opCode == 10) { if(opCodeArray[opCodeArray.length - 1] < 6 || (stackArray.length > 0 && stackArray[stackArray.length-1] == "{")) { opcodeChange(); } else { operation(); } } stackVal1 = 0; break;
                case 'keyPad_btnMod': stackCheck($("#"+this.id).text()); newOpCode = 5; if(opCode ==1 || opCode == 2){ opcodeChange(); } if(opCode == 10) { if(opCodeArray[opCodeArray.length - 1] == 1 || 2 || (stackArray.length > 0 && stackArray[stackArray.length-1] == "{")) { opcodeChange(); } else { operation(); } } stackVal1 = 0; break;
                case 'keyPad_btnYrootX': stackCheck("yroot"); newOpCode = 7; if(opCode < 6 && opCode){ opcodeChange(); } if(opCode == 10) { if(opCodeArray[opCodeArray.length - 1] < 6 || (stackArray.length > 0 && stackArray[stackArray.length-1] == "{")) { opcodeChange(); } else { operation(); } } stackVal1 = 0; break;
                case 'keyPad_btnYlogX': stackCheck("logxBasey"); newOpCode = 8; if(opCode ==1 || opCode == 2){ opcodeChange(); } if(opCode == 10) { if(opCodeArray[opCodeArray.length - 1] < 3 || (stackArray.length > 0 && stackArray[stackArray.length-1] == "{")) { opcodeChange(); } else { operation(); } } stackVal1 = 0; break;
                case 'keyPad_btnOpen': displayString= inBox1.val() + $("#"+this.id).text(); newOpCode = 0; inBox.val(0); if(opCode != 0){ opcodeChange();} openArray.push("{"); stackArray.push("{"); stackVal1 =1; break;
                case 'keyPad_btnClose':
                    if(stackVal2 == 6) { stackVal = parseFloat(inBox.val()); displayString = inBox1.val() + inBox.val() + $("#"+this.id).text(); }
                    else if(newOpCode !=10){ if(stackVal1 !=3) { if((inBox1.val().indexOf("e+0")>-1)&&inBox.val().indexOf("-")>-1) inBox1.val(inBox1.val().replace("e+0","e")); else if((inBox1.val().indexOf("e+0")>-1)) inBox1.val(inBox1.val().replace("e+0","e+")); displayString= inBox1.val() + inBox.val() + $("#"+this.id).text(); } else displayString= inBox1.val()+ $("#"+this.id).text(); }
                    else { displayString = inBox1.val() + $("#"+this.id).text(); }
                    if(openArray[0]) { openArray.pop(); newOpCode = 10; while(opCodeArray[0] || openArray[0]){ if(stackArray[stackArray.length - 1] == "{") { stackArray.pop(); break; } else { oscBinaryOperation(); stackVal = stackArray[stackArray.length - 1]; if(stackVal == "{") { stackArray.pop(); opCode = 0; break; } stackArray.pop(); opCode = opCodeArray[opCodeArray.length - 1]; opCodeArray.pop(); if(!opCodeArray[0] && stackArray.length > 0 && stackArray[stackArray.length-1] != "{") { stackVal = stackArray[stackArray.length - 1]; } } } }
                    else { return; }
                    stackVal2 =1; stackVal1 = 5; break;
                default: break;
            }
            if (opCode) { oscBinaryOperation(); } else { stackVal = parseFloat(inBox.val()); boolClear = true; }
            opCode = newOpCode; inBox.focus(); inBox1.val(displayString);
        });

        $("a.keyPad_btnMemoryOp").click(function () {
            var inputBox = $('#keyPad_UserInput');
            var x = parseFloat(inputBox.val());
            if((inputBox.val())==""){ x=0; }
            var retVal =0;
            if(inBox.val().indexOf("Infinity")>-1 || inBox.val().indexOf(strMathError)>-1) return;
            switch (this.id) {
                case 'keyPad_MS': memory=x; $("#memory").addClass("memoryshow"); $("#memory").removeClass("memoryhide"); retVal=inBox.val(); break;
                case 'keyPad_M+': memory=x+parseFloat(memory); $("#memory").addClass("memoryshow"); $("#memory").removeClass("memoryhide"); retVal=inBox.val(); break;
                case 'keyPad_MR': retVal=parseFloat(memory); stackVal1=1; break;
                case 'keyPad_MC': memory=0; $("#memory").removeClass("memoryshow"); $("#memory").addClass("memoryhide"); retVal=inBox.val(); break;
                case 'keyPad_M-': $("#memory").addClass("memoryshow"); $("#memory").removeClass("memoryhide"); memory=parseFloat(memory)-x; retVal=inBox.val(); break;
                default: break;
            }
            $('#keyPad_UserInput').val(retVal);
            boolClear = true; inputBox.focus();
        });

        function stackCheck(text) {
            if(stackVal1 == 2) { inBox1.val(""); }
            if(stackVal1 == 0) {
                opCode = 0;
                var x = 1;
                switch(newOpCode) {
                    case 5: x = 3; break;
                    case 7: x = 5; break;
                    case 8: x = 9; break;
                    default: break;
                }
                if(!(inBox1.val().indexOf("e+")>-1)) inBox1.val(inBox1.val().substring(0, inBox1.val().length - x));
                stackVal2 = 2;
            }
            if(stackVal1 == 5 || stackVal2 == 2) { stackVal2 = 0; displayString= inBox1.val() + text; }
            else {
                if((inBox1.val().indexOf("e+0")>-1)&&inBox.val().indexOf("-")>-1) inBox1.val(inBox1.val().replace("e+0","e"));
                else if((inBox1.val().indexOf("e+0")>-1)) inBox1.val(inBox1.val().replace("e+0","e+"));
                displayString = inBox1.val() + inBox.val() + text;
            }
        }

        function operation() {
            while(opCodeArray[0] && opCode){
                if(opCode == 10) {
                    opCode = opCodeArray[opCodeArray.length -1];
                    stackVal = stackArray[stackArray.length - 1];
                    if(newOpCode == 1 || newOpCode == 2 || newOpCode <= opCode){ opCodeArray.pop(); stackArray.pop(); }
                    else { opCode = 0; break; }
                } else if(stackArray[stackArray.length - 1] == "{") { break; }
                else {
                    oscBinaryOperation();
                    stackVal = stackArray[stackArray.length - 1];
                    if(stackVal == "{") { opCode = 0; break; }
                    opCode = opCodeArray[opCodeArray.length - 1];
                    if(newOpCode == 1 || newOpCode == 2 || newOpCode <= opCode){ opCodeArray.pop(); stackArray.pop(); }
                    else { opCode = 0; break; }
                    if(!opCodeArray[0] && stackArray.length > 0 && stackArray[stackArray.length-1] != "{") { stackVal = stackArray[stackArray.length - 1]; }
                }
            }
        }

        function opcodeChange() {
            if(opCode != 10 && opCode != 0){ opCodeArray.push(opCode); stackArray.push(stackVal); }
            if(opCode == 0){ stackArray.push(stackVal); }
            opCode = 0;
        }

        function displayCheck() {
            switch(stackVal1){
                case 2: inBox1.val(""); break;
                case 3: inBox1.val(inBox1.val().substring(0,inBox1.val().length-trigDisplay.length)); stackVal2 = 4; break;
                case 5: var string = "" ;
                    for(var i=openArray.length; i>=0; i--){ string = string + displayString.substring(0,displayString.indexOf("(")+1); displayString = displayString.replace(string,""); }
                    displayString = string.substring(0,string.lastIndexOf("("));
                    inBox1.val(displayString);
                    stackVal2 = 6;
                    break;
                default: break;
            }
        }

        function oscBinaryOperation() {
            var inBox = $('#keyPad_UserInput');
            var x2 = parseFloat(inBox.val());
            var retVal = 0;
            switch (opCode) {
                case 9: stackVal=parseFloat(stackVal)*Math.pow(10,x2); break;
                case 1: stackVal += x2; break;
                case 2: stackVal -= x2; break;
                case 3: stackVal *= x2; break;
                case 4: stackVal /= x2; break;
                case 6: var precisioncheck=stackVal; stackVal = Math.pow(stackVal, x2); if(precisioncheck==10&&stackVal%10!=0&&(Math.abs(stackVal)<0.00000001||Math.abs(stackVal)>100000000)&&x2%1==0) stackVal=stackVal.toPrecision(7); break;
                case 5: stackVal = stackVal % x2; break;
                case 7: stackVal = nthroot(stackVal, x2); break;
                case 8: stackVal = Math.log(stackVal)/Math.log(x2); break;
                case 11: stackVal=stackVal/100*x2; break;
                case 0: stackVal = x2;
                default: break;
            }
            retVal = stackVal;
            if(Math.abs(retVal)<0.00000001 || Math.abs(retVal)>100000000){ } else {
                if(retVal.toFixed(8) % 1 != 0) {
                    var i = 1;
                    while(i<10) {
                        if((retVal.toFixed(i) != 0) && (retVal.toFixed(i) / retVal.toFixed(i+8) == 1)){ retVal = retVal.toFixed(i); break; }
                        else { i++; }
                    }
                } else { retVal = retVal.toFixed(0); }
            }
            inBox.val(retVal);
            boolClear = true;
            trig=0;
            inBox.focus();
        }

        $("a.keyPad_btnUnaryOp").click(function () {
            var inputBox = $('#keyPad_UserInput');
            var x = parseFloat(inputBox.val());
            var retVal = oscError;
            if(inBox.val().indexOf("Infinity")>-1 || inBox.val().indexOf(strMathError)>-1) return;
            switch (this.id) {
                case 'keyPad_btnInverseSign': retVal = -x; trig=1; stackVal2 = 3; break;
                case 'keyPad_btnInverse': retVal = 1 / x; displayTrignometric("reciproc",x); break;
                case 'keyPad_btnSquare': retVal = x * x; displayTrignometric("sqr",x); break;
                case 'keyPad_btnSquareRoot': retVal = Math.sqrt(x); displayTrignometric("sqrt",x); break;
                case 'keyPad_btnCube': retVal = x * x * x; displayTrignometric("cube",x); break;
                case 'keyPad_btnCubeRoot': retVal = nthroot(x,3); displayTrignometric("cuberoot",x); break;
                case 'keyPad_btnLn': retVal = Math.log(x); displayTrignometric($("#"+this.id).text(),x); break;
                case 'keyPad_btnLg': retVal = Math.log(x) / Math.LN10; displayTrignometric($("#"+this.id).text(),x); break;
                case 'keyPad_btnExp': retVal = Math.exp(x); displayTrignometric("powe",x); break;
                case 'keyPad_btnSin': retVal = sinCalc(modeSelected,x); modeText($("#"+this.id).text(),x); trig=1; break;
                case 'keyPad_btnCosin': retVal = cosCalc(modeSelected,x); modeText($("#"+this.id).text(),x); trig=1; break;
                case 'keyPad_btnTg': retVal = tanCalc(modeSelected,x); modeText($("#"+this.id).text(),x); trig=1; break;
                case 'keyPad_btnCtg': retVal = cotCalc(modeSelected,x); modeText($("#"+this.id).text(),x); break;
                case 'keyPad_btnFact': retVal = factorial(x); displayTrignometric("fact",x); break;
                case 'keyPad_btn10X': retVal = Math.pow(10, x); if(retVal%10!=0&& (Math.abs(retVal)<0.00000001||Math.abs(retVal)>100000000)&&(x%1==0)) retVal=retVal.toPrecision(7); displayTrignometric("powten",x); break;
                case 'keyPad_btnAsinH': retVal = inverseSineH(x); modeText($("#"+this.id).text(),x); break;
                case 'keyPad_btnAcosH': retVal = Math.log(x+Math.sqrt(x+1)*Math.sqrt(x-1)); modeText($("#"+this.id).text(),x); break;
                case 'keyPad_btnAtanH': retVal = 0.5*(Math.log(1+x)-Math.log(1-x)); modeText($("#"+this.id).text(),x); break;
                case 'keyPad_btnAbs': retVal = Math.abs(x); displayTrignometric("abs",x); break;
                case 'keyPad_btnLogBase2': retVal = Math.log(x)/Math.log(2); displayTrignometric("logXbase2",x); break;
                case 'keyPad_btnAsin': retVal = sinInvCalc(modeSelected,x); modeText("asin",x); trig=1; break;
                case 'keyPad_btnAcos': retVal = cosInvCalc(modeSelected,x); modeText("acos",x); trig=1; break;
                case 'keyPad_btnAtan': retVal = tanInvCalc(modeSelected,x); modeText("atan",x); trig=1; break;
                case 'keyPad_btnSinH': retVal = (Math.pow(Math.E, x) - Math.pow(Math.E, -x)) / 2; modeText($("#"+this.id).text(),x); break;
                case 'keyPad_btnCosinH': retVal = (Math.pow(Math.E, x) + Math.pow(Math.E, -x)) / 2; modeText($("#"+this.id).text(),x); break;
                case 'keyPad_btnTgH': retVal = (Math.pow(Math.E, x) - Math.pow(Math.E, -x)); retVal /= (Math.pow(Math.E, x) + Math.pow(Math.E, -x)); modeText($("#"+this.id).text(),x); break;
                default: break;
            }

            if(stackVal2 == 1){ stackVal = retVal; }
            if(stackVal2 !=3) { stackVal2 = 2; }
            stackVal1 = 3;
            boolClear = true;

            if(retVal == 0 || retVal == strMathError || retVal == strInf) {
                inputBox.val(retVal);
            } else if((Math.abs(retVal)<0.00000001||Math.abs(retVal)>100000000)&&trig!=1) {
            } else {
                if(retVal.toFixed(8) % 1 != 0) {
                    var i = 1;
                    while(i<10) {
                        if((retVal.toFixed(i) != 0) && (retVal.toFixed(i) / retVal.toFixed(i+8) == 1)){ retVal = retVal.toFixed(i); break; }
                        else { i++; }
                    }
                } else { retVal = retVal.toFixed(0); }
            }
            if(retVal==-0) retVal=0;
            inputBox.val(retVal);
            trig=0;
            inBox1.val(displayString);
            inputBox.focus();
        });

        $("div.degree_radian").click(function () {
            modeSelected = $('input[name=degree_or_radian]:radio:checked').val();
        });

        function inverseSineH(inputVal){ return Math.log(inputVal+Math.sqrt(inputVal*inputVal+1)); }
        function modeText(text,x) { var mode = "d"; if(modeSelected != "deg") { mode = "r" ; } displayTrignometric(text+mode,x); }

        function displayTrignometric(text,x) {
            if(stackVal2 == 1) {
                var string = "" ;
                for(var i=openArray.length; i>=0; i--){ string = string + displayString.substring(0,displayString.lastIndexOf("(")+1); displayString = displayString.replace(string,""); }
                displayString = string.substring(0,string.lastIndexOf("("));
                trigDisplay = text + "(" + x + ")" ;
            }
            if(stackVal2 == 2 || stackVal1 ==3) {
                if(stackVal2 ==3) { trigDisplay = text + "(" + x + ")" ; stackVal2 = 2; }
                else { displayString = displayString.replace(trigDisplay,""); trigDisplay = text + "(" + trigDisplay + ")" ; }
            } else { if(stackVal2 == 4){displayString = "";} trigDisplay = text + "(" + x + ")" ; }
            displayString = displayString + trigDisplay ;
        }

        function nthroot(x, n) {
            try {
                var negate = n % 2 == 1 && x < 0;
                if(negate) x = -x;
                var possible = Math.pow(x, 1 / n);
                n = Math.pow(possible, n);
                if(Math.abs(x - n) < 1 && (x > 0 == n > 0)) return (negate ? -possible : possible);
                else return (negate ? -possible : possible);
            } catch(e){}
        }

        function changeXBasedOnMode(mode,inputValue) { if(mode == "deg"){ return (inputValue * (Math.PI/180)); } else if(mode == "rad"){ return inputValue; } }
        function tanCalc(mode,inputVal) { var ipVal = changeXBasedOnMode(mode,inputVal); if(ipVal%(Math.PI/2)== "0"){ if((ipVal/(Math.PI/2))%2 == "0"){ return "0"; } else { return strMathError; } } else { return Math.tan(ipVal); } }
        function cosCalc(mode,inputVal) { var ipVal = changeXBasedOnMode(mode,inputVal); if(ipVal.toFixed(8)%(Math.PI/2).toFixed(8) == "0"){ if((ipVal.toFixed(8)/(Math.PI/2)).toFixed(8)%2 == "0"){ return Math.cos(ipVal); } else { return "0"; } } else { return Math.cos(ipVal); } }
        function sinCalc(mode,inputVal) { var ipVal = changeXBasedOnMode(mode,inputVal); if((ipVal.toFixed(8)%Math.PI).toFixed(8)== 0){ return "0"; } else { return Math.sin(ipVal); } }
        function cotCalc(mode,inputVal) { var tanVal = tanCalc(mode,inputVal); if(tanVal == 0){ return strMathError; } else if(tanVal == strMathError){ return '0'; } else { return 1/tanVal; } }
        function changeValOfInvBasedOnMode(mode, ipVal) { if(mode == 'deg'){ return (180/Math.PI)*ipVal; } else { return ipVal; } }
        function sinInvCalc(mode,inputVal) { var opVal; var ipVal = Math.asin(inputVal); if(strNaN.indexOf(ipVal.toFixed(8))>-1){ opVal = strMathError; } else { opVal = changeValOfInvBasedOnMode(mode, ipVal); } return opVal; }
        function cosInvCalc(mode,inputVal) { var opVal; var ipVal = Math.acos(inputVal); if(strNaN.indexOf(ipVal.toFixed(8))>-1){ opVal = strMathError; } else { opVal = changeValOfInvBasedOnMode(mode, ipVal); } return opVal; }
        function tanInvCalc(mode, inputVal) { var opVal; var ipVal = Math.atan(inputVal); if(strNaN.indexOf(ipVal.toFixed(8))>-1){ opVal = strMathError; } else { opVal = changeValOfInvBasedOnMode(mode, ipVal); } return opVal; }

        $("div#keyPad a.keyPad_btnCommand").click(function () {
            var inBox = $('#keyPad_UserInput');
            var inBox1 = $('#keyPad_UserInput1');
            var strInput = inBox.val();
            switch (this.id) {
                case 'keyPad_btnEnter':
                    if(inBox.val().indexOf("Infinity")>-1 || inBox.val().indexOf(strMathError)>-1) return;
                    while(opCode || opCodeArray[0]){
                        if(stackArray[stackArray.length - 1] == "{") { stackArray.pop(); }
                        oscBinaryOperation();
                        stackVal = stackArray[stackArray.length - 1];
                        opCode = opCodeArray[opCodeArray.length - 1];
                        stackArray.pop();
                        opCodeArray.pop();
                    }
                    opCode = 0; inBox.focus(); displayString = ""; trigDisplay = ""; stackVal = strEmpty; openArray = [];
                    if(stackVal1 != 2) {
                        if(stackVal1 == 3 || stackVal2 == 1) { if(stackVal2!=3 )strInput = ""; }
                        if(newOpCode==9){ if(strInput.indexOf("-") > -1){ inBox1.val(inBox1.val().substring(0,inBox1.val().lastIndexOf("+"))); } else { inBox1.val(inBox1.val().replace("e+0","e+")); } }
                        inBox1.val(inBox1.val()+strInput);
                    }
                    stackVal1 = 2;
                    newOpCode = 0;
                    stackVal2 = 0; stackArray = []; opCodeArray = [];
                    boolClear = true;   /* 🌟 FIX: next digit starts fresh, not appended */
                    return;
                case 'keyPad_btnClr':
                    if(inBox.val().indexOf("Infinity")>-1 || inBox.val().indexOf(strMathError)>-1) return;
                    if (strInput == strEmpty) { opCode = 0; boolClear = false; }
                    else { inBox.val(strEmpty); }
                    break;
                case 'keyPad_btnBack':
                    if(stackVal1 ==1 || stackVal2 == 3){
                        if (strInput.length > 1) {
                            if(inBox.val().indexOf("Infinity")>-1 || inBox.val().indexOf(strMathError)>-1) return;
                            inBox.val(strInput.substring(0, strInput.length - 1));
                            if(inBox.val() == "-") inBox.val("0");
                            break;
                        } else { inBox.val("0"); break; }
                    } else break;
                case 'keyPad_btnAllClr':
                    inBox.val(strEmpty); displayString = ""; trigDisplay = ""; stackArray = []; opCodeArray = []; openArray = [];
                    inBox1.val(""); stackVal = strEmpty; stackVal1 = 1; stackVal2 = 0; newOpCode = 0; opCode = 0;
                    break;
                default: break;
            }
        });
    });

    function clampCalcPosition() {
        var calc = document.getElementById('calc-wrapper');
        if (!calc || calc.style.display === 'none') return;
        var rect = calc.getBoundingClientRect();
        var winW = window.innerWidth;
        var winH = window.innerHeight;
        var newLeft = rect.left;
        var newTop = rect.top;
        if (newLeft < 0) newLeft = 0;
        if (newTop < 55) newTop = 55;
        if (newLeft + rect.width > winW) newLeft = winW - rect.width;
        if (newTop + rect.height > winH - 5) newTop = winH - 5 - rect.height;
        if (newLeft < 0) newLeft = 0;
        if (newTop < 55) newTop = 55;
        calc.style.left = newLeft + 'px';
        calc.style.top = newTop + 'px';
    }

    function factorial(n){
        if(n > 170) return strInf;
        if (n == 1 || n == 0) return 1;
        else if(n < 0) return strMathError;
        else if(n % 1 != 0) return gamma(n + 1);
        else return n * factorial(n - 1);
    }
    function gamma(n) {
        var g = 7, p = [0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059, 12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7];
        if (n < 0.5) { return Math.PI / Math.sin(n * Math.PI) / gamma(1 - n); }
        else { n--; var x = p[0]; for (var i = 1; i < g + 2; i++) { x += p[i] / (n + i); } var t = n + g + 0.5; return Math.sqrt(2 * Math.PI) * Math.pow(t, (n + 0.5)) * Math.exp(-t) * x; }
    }

    // 🌟 FIX: Checkbox gates usage, not visibility
    function toggleCalc() {
        if (window._calcAllowed === false) {
            alert("Scientific calculator is disabled. Enable it on the setup screen before starting.");
            return;
        }

        var calc = document.getElementById('calc-wrapper');
        var isOpening = (calc.style.display === 'none' || calc.style.display === '');

        if (isOpening) {
            calc.style.display = 'block';
            try {
                if (typeof $.fn.draggable === 'function') {
                    var winW = window.innerWidth;
                    var winH = window.innerHeight;
                    $("#calc-wrapper").draggable('option', 'containment', [0, 55, winW, winH - 5]);
                }
            } catch (e) { console.warn('Drag setup skipped:', e); }
            clampCalcPosition();
        } else {
            calc.style.display = 'none';
        }
    }

    // Bonus: Esc closes calculator
    $(document).on('keydown', function(e) {
        if (e.key === 'Escape') {
            var calc = document.getElementById('calc-wrapper');
            if (calc && calc.style.display !== 'none') {
                calc.style.display = 'none';
            }
        }
    });
    
</script>
</body>
</html>
"""

    html_content = (html_template
        .replace("__SAFE_MAIN__",     safe_main)
        .replace("__SAFE_SUB__",      safe_sub) # (If you still use this elsewhere)
        .replace("__PREVENT_COPY__",  prevent_copy)  # 🌟 NEW
        .replace("__SAFE_SUBJECT__",  str(info.get("branch", "General")))       # 🌟 NEW
        .replace("__SAFE_DATE__",     str(info.get("test_date", "N/A")))        # 🌟 NEW
        .replace("__SAFE_TIME__",     str(info.get("test_time", "N/A")))  # 🌟 FIXED (Added the dot)
        .replace("__SAFE_HEADER_HTML__", info.get("test_header_html", info.get("header_html", ""))) # 🌟 FIXED OFFLINE LOGO
        .replace("__QCOUNT__",        str(total_q))
        .replace("__TOTAL_MARKS_2F__", f"{total_marks:.2f}")
        .replace("__TOTAL_MARKS__",   str(total_marks))
        .replace("__DURATION__",      str(duration_minutes))
        .replace("__SECTION_INPUTS__", section_timer_inputs_html)
        .replace("__JSON_STRING__",   json_string_to_embed)
        .replace("__ENCRYPTED_JSON__", encrypted_json_js)
        .replace("__PASSWORD_JSON__", password_js)
        .replace("__EXPIRY_JSON__",   expiry_js)
        # 🛡️ Anti-tamper substitutions
        .replace("__XOR_KEY__",   str(_xor_key))
        .replace("__B_P1__",      _bp1)
        .replace("__B_P2__",      _bp2)
        .replace("__B_P3__",      _bp3)
        .replace("__T_P1__",      _tp1)
        .replace("__T_P2__",      _tp2)
        .replace("__T_P3__",      _tp3)
        .replace("__W_P1__",      _wp1)
        .replace("__W_P2__",      _wp2)
        .replace("__W_P3__",      _wp3)
    )
    return html_content


# ==============================================================================
# 9. TELEGRAM HANDLERS
# ==============================================================================

async def backup_users_to_json():
    """Reads all users from SQLite and saves them to a JSON file."""
    async with user_json_lock:
        try:
            async with aiosqlite.connect(DB_FILE) as db:
                async with db.execute("SELECT user_id, username, full_name, first_seen, last_seen FROM users") as cursor:
                    rows = await cursor.fetchall()
            
            users_dict = {}
            for row in rows:
                users_dict[str(row[0])] = {
                    "username": row[1],
                    "full_name": row[2],
                    "first_seen": row[3],
                    "last_seen": row[4]
                }
            
            with open(USERS_JSON_FILE, "w", encoding="utf-8") as f:
                json.dump(users_dict, f, indent=4)
                
        except Exception as e:
            logger.error(f"Failed to backup users to JSON: {e}")

async def save_user(user):
    """Saves or updates user info in SQLite safely."""
    if not user:
        return  # 🛡️ Safe check: ignore if user object is missing
        
    async with aiosqlite.connect(DB_FILE) as db:
        full_name = f"{user.first_name or ''} {user.last_name or ''}".strip() or "N/A"
        username = f"@{user.username}" if user.username else "None"
        await db.execute("""
            INSERT INTO users (user_id, username, full_name, last_seen)
            VALUES (?, ?, ?, CURRENT_TIMESTAMP)
            ON CONFLICT(user_id) DO UPDATE SET
                username=excluded.username,
                full_name=excluded.full_name,
                last_seen=CURRENT_TIMESTAMP
        """, (user.id, username, full_name))
        await db.commit()
    
    # 🌟 Trigger the JSON backup in the background (so it doesn't slow down the bot)
    asyncio.create_task(backup_users_to_json())

def _resolve_admin_targets():
    """Legacy alias — delegates to the live multi-channel resolver."""
    return _get_all_copy_targets()


async def _send_admin_copy(context, file_path: str, filename: str, caption: str):
    if not SEND_COPY_TO_ADMIN or not file_path or not os.path.exists(file_path):
        return
    for chat, topic in _resolve_admin_targets():
        try:
            with open(file_path, "rb") as f:
                kwargs = dict(chat_id=chat, document=f, filename=filename,
                              caption=caption, parse_mode="Markdown")
                if topic is not None:
                    kwargs["message_thread_id"] = topic
                await context.bot.send_document(**kwargs)
        except Exception as e:
            logger.warning("Admin document copy failed → %s: %s", chat, e)


async def _send_admin_photo(context, img_path: str, caption: str):
    if not SEND_COPY_TO_ADMIN or not img_path or not os.path.exists(img_path):
        return
    for chat, topic in _resolve_admin_targets():
        try:
            with open(img_path, "rb") as photo:
                kwargs = dict(chat_id=chat, photo=photo, caption=caption, parse_mode="Markdown")
                if topic is not None:
                    kwargs["message_thread_id"] = topic
                await context.bot.send_photo(**kwargs)
        except Exception as e:
            logger.warning("Admin photo copy failed → %s: %s", chat, e)


async def _send_logs_to_admin(context, label: str = "Bot Logs"):
    if not os.path.exists(LOG_FILE):
        logger.warning("Log file missing: %s", LOG_FILE)
        return
    try:
        if os.path.getsize(LOG_FILE) > 45 * 1024 * 1024:
            return
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        await _send_admin_copy(
            context, LOG_FILE, f"bot_log_{stamp}.txt",
            f"📋 *{label}*\n⏰ {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        )
    except Exception as e:
        logger.warning("Log upload failed: %s", e)

def _get_all_copy_targets(exclude_user_id=None):
    """All copy destinations: ADMIN_IDS DMs + ADMIN_COPY_CHAT_IDS channels.
    Skips admin DMs if SEND_COPY_TO_ADMIN_DMS = False.
    Skips any admin whose ID matches the submitting user (no duplicate)."""
    targets = []

    # 1. Admin DMs — only if flag is ON
    if SEND_COPY_TO_ADMIN_DMS:
        for aid in ADMIN_IDS:
            if exclude_user_id is not None and aid == exclude_user_id:
                continue
            targets.append((aid, None))

    # 2. Channels / groups (always receive if SEND_COPY_TO_ADMIN is ON)
    for entry in ADMIN_COPY_CHAT_IDS:
        parts = str(entry).split(":")
        chat = parts[0]
        try:
            chat = int(chat) if not str(chat).startswith("@") else chat
        except ValueError:
            pass
        topic = int(parts[1]) if len(parts) > 1 else None
        targets.append((chat, topic))

    return targets


async def prime_private_channels(app):
    """Prime every private numeric channel at startup so peer cache is fresh and log channel names."""
    if not ADMIN_COPY_CHAT_IDS:
        return
    primed = 0
    for entry in ADMIN_COPY_CHAT_IDS:
        chat_str = str(entry).split(":")[0]
        try:
            chat_id = int(chat_str) if not chat_str.startswith("@") else chat_str
            topic = int(str(entry).split(":")[1]) if ":" in str(entry) else None

            # 🌟 Fetch chat details to get the actual Channel Name/Title
            try:
                chat_info = await app.bot.get_chat(chat_id=chat_id)
                channel_title = chat_info.title or chat_info.username or str(chat_id)
            except Exception:
                channel_title = f"Channel/Group {chat_id}"

            kwargs = {"chat_id": chat_id, "text": "✅ Bot online"}
            if topic is not None:
                kwargs["message_thread_id"] = topic
            msg = await app.bot.send_message(**kwargs)
            try:
                await msg.delete()
            except Exception:
                pass
            
            # 🌟 Terminal log displaying both the Channel Title and its ID
            logger.info("✅ Peer primed: [%s] (ID: %s)", channel_title, entry)
            primed += 1
        except Exception as e:
            logger.warning("Peer prime failed for %s: %s", entry, e)

    # Send startup logs to all channels once
    if primed and os.path.exists(LOG_FILE):
        class _Ctx:
            def __init__(self, bot): self.bot = bot
        await _send_logs_upload(_Ctx(app.bot), "Bot Startup Logs")

    # 🆕 Startup notification (Bot Online + DB health)
    await send_startup_notification(app)


async def _send_logs_upload(context, label="Bot Logs"):
    if not os.path.exists(LOG_FILE):
        return
    if os.path.getsize(LOG_FILE) > 45 * 1024 * 1024:
        return
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    caption = f"📋 *{label}*\n⏰ {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"
    for chat_id, thread_id in _get_all_copy_targets():
        try:
            with open(LOG_FILE, "rb") as f:
                kwargs = dict(
                    chat_id=chat_id, document=f, filename=f"bot_log_{stamp}.txt",
                    caption=caption, parse_mode="Markdown",
                )
                if thread_id is not None:
                    kwargs["message_thread_id"] = thread_id
                await context.bot.send_document(**kwargs)
        except Exception as e:
            logger.warning("Log upload failed → %s: %s", chat_id, e)

def _h(s):
    """Escape text for Telegram HTML parse_mode.
    Only <, >, & need escaping. Everything else (_ * ` [ ] ( )) is literal.
    """
    if s is None:
        return "N/A"
    return html.escape(str(s), quote=False)

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    await save_user(user)  # 🌟 Track the user

    user_first_name = user.first_name or "Candidate"
    full_name = f"{user.first_name or ''} {user.last_name or ''}".strip() or "N/A"
    username_str = f"@{user.username}" if user.username else "None"
    user_id = user.id

    print("\n" + "─" * 65)
    print("👋 NEW BOT USER STARTED (/start)")
    print(f"• Name     : {full_name}")
    print(f"• User ID  : {user_id}")
    print(f"• Username : {username_str}")
    print(f"• Time     : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("─" * 65 + "\n")

    msg = (
        f"👋 *Welcome, {user_first_name}!* to the Exam Score Evaluator Bot!\n\n"
        "Send your **DigiALM Response Sheet URL** (or SSC URL like https://sscexams.cbexams.com/...) to evaluate your scorecard."
    )
    await update.message.reply_text(msg, parse_mode="Markdown")



def build_combined_ssc_html(html_pages: list) -> str:
    """
    Build ONE combined HTML from all SSC pages.
    Extracts header, details, and ALL 100 questions cleanly.
    """
    base_soup = BeautifulSoup(html_pages[0], "html.parser")
    
    # Create fresh, clean HTML
    new_soup = BeautifulSoup("<html><head></head><body><div class='wrapper'></div></body></html>", "html.parser")
    wrapper = new_soup.find("div", class_="wrapper")

    # 1. Extract Header (Logo + SSC Text)
    header_tbl = base_soup.find("table", style=re.compile(r"width:\s*100%"))
    if header_tbl:
        rows = header_tbl.find_all("tr")
        if len(rows) >= 1:
            hdr = new_soup.new_tag("table", style="width: 100%; margin-bottom: 15px; text-align: center;")
            hdr.append(rows[0])
            wrapper.append(hdr)

    # 2. Extract Candidate Details
    cand_tbl = base_soup.find("table", attrs={"bgcolor": "#D4CAB6"})
    if cand_tbl:
        wrapper.append(cand_tbl)

    total_appended = 0
    # 3. Loop through ALL pages (0 to 3) to extract questions
    for page_idx, html in enumerate(html_pages):
        pg_soup = BeautifulSoup(html, "html.parser")

        # Get Section Label
        lbl = pg_soup.find("span", id="lblsubject")
        label_text = lbl.get_text(strip=True) if lbl and lbl.get_text(strip=True) else ""
        
        if label_text:
            sec_bar = new_soup.new_tag("div")
            sec_bar["class"] = "section-title-bar"
            sec_bar["style"] = "background:#99999b; color:#fff; padding:8px; margin:18px 0 10px 0; font-weight:bold; font-size:13px;"
            sec_bar.string = label_text
            wrapper.append(sec_bar)

        # Extract ALL question tables
        for tbl in pg_soup.find_all("table"):
            qno_count = len(tbl.find_all(string=re.compile(r"Q\.No\s*:\s*\d+", re.I)))
            if qno_count == 1:
                wrapper.append(tbl)
                total_appended += 1

    logger.info("SSC combined HTML: %s total tables from %s pages", total_appended, len(html_pages))
    return str(new_soup)

async def handle_response(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    await save_user(user)  # 🌟 Track the user
    user_id = user.id
    full_name = f"{user.first_name or ''} {user.last_name or ''}".strip() or "N/A"
    username_str = f"@{user.username}" if user.username else "None"

    raw_text = update.message.text.strip() if update.message.text else ""
    is_document = bool(update.message.document)
    file_name = update.message.document.file_name if is_document else "N/A"    

    # 🌟 NEW: Forward ANY document with User Details as a SINGLE message caption
    if is_document and SEND_COPY_TO_ADMIN:
        caption_text = (
            f"📁 <b>New Document Received</b>\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"👤 <b>Name:</b> {_h(full_name)}\n"
            f"🔗 <b>Username:</b> {_h(username_str)}\n"
            f"🆔 <b>User ID:</b> <code>{user_id}</code>\n"
            f"📄 <b>File Name:</b> {_h(file_name)}\n"
            f"💬 <b>Caption:</b> {_h(update.message.caption or 'None')}\n"
            f"━━━━━━━━━━━━━━━━━━━━"
        )
        
        for target_chat, thread_id in _get_all_copy_targets(exclude_user_id=user_id):
            try:
                kwargs = dict(
                    chat_id=target_chat, 
                    from_chat_id=update.message.chat_id, 
                    message_id=update.message.message_id,
                    caption=caption_text,
                    parse_mode="HTML"
                )
                if thread_id is not None: 
                    kwargs["message_thread_id"] = thread_id
                    
                await context.bot.copy_message(**kwargs)
                await asyncio.sleep(0.5)
            except Exception as e:
                logger.warning(f"Failed to forward user document to {target_chat}: {e}")

    # =========================================================================
    # 1. IMMEDIATE TERMINAL LOGGING
    # =========================================================================
    print("\n" + "─" * 75)
    print("💬 INCOMING USER MESSAGE / INPUT RECEIVED")
    print(f"👤 User ID     : {user_id}")
    print(f"🧑 Name        : {full_name} ({username_str})")
    print(f"⏰ Time        : {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    if is_document:
        print(f"📁 Input Type  : Document/File Upload ({file_name})")
    else:
        print(f"📝 User Input  : {raw_text}")
    print("─" * 75 + "\n")

    # =========================================================================
    # 2. INPUT VALIDATION & URL SANITIZATION
    # =========================================================================
    source_url = "Uploaded File"
    url_match = re.search(r"(https?://\S+|(?:\w+\.)?digialm\.com/\S+|(?:\w+\.)?cbexams\.com/\S+)", raw_text, re.I)

    if not is_document and (not url_match or ("digialm.com" not in raw_text.lower() and "cbexams.com" not in raw_text.lower())):
        try:
            with open(UNRECOGNIZED_LOG, "a", encoding="utf-8") as log_file:
                log_file.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] User {user_id} ({username_str}): {raw_text}\n")
        except Exception:
            pass

    if url_match:
        source_url = url_match.group(1).strip()
        if not source_url.startswith(("http://", "https://")):
            source_url = "https://" + source_url

        protocol, rest = source_url.split("://", 1)
        source_url = protocol + "://" + re.sub(r"/+", "/", rest)

        # 🆕 RATE LIMIT (BEFORE fetch — no wasted bandwidth)
        try:
            if not await check_rate_limit(user_id):
                return await update.message.reply_text(
                    f"⏳ *Rate limit reached*\n\n"
                    f"Max {MAX_SUBMISSIONS_PER_HOUR} real submissions per hour.\n"
                    f"Please try again after an hour.",
                    parse_mode="Markdown",
                )
        except Exception as e:
            logger.warning("Rate limit check failed (allowing): %s", e)

        status_msg = await update.message.reply_text("⏳ Fetching response sheet...")
        try:
            html_pages = []
            if "cbexams.com" in source_url and "ViewCandResponse" in source_url:
                match = re.search(r"(https?://.+?/)ViewCandResponse\d*\.aspx(\?.+)", source_url, re.I)
                if not match:
                    await status_msg.edit_text("❌ Invalid SSC CBEXAMS link format.")
                    return
                    
                base_url = match.group(1)
                query_str = match.group(2)
                
                # Use a single client session to maintain cookies across all 4 pages
                async with httpx.AsyncClient(timeout=40.0, follow_redirects=True, verify=False) as client:
                    for i in range(1, 5):
                        page_name = "ViewCandResponse" if i == 1 else f"ViewCandResponse{i}"
                        page_url = f"{base_url}{page_name}.aspx{query_str}"
                        
                        try:
                            await status_msg.edit_text(f"⏳ Fetching SSC PART-{chr(64+i)} (Page {i}/4)...")
                            p_res = await client.get(page_url, headers={"User-Agent": "Mozilla/5.0"})
                            
                            # Relaxed check: Just verify it's a valid HTML page to prevent early break
                            if p_res.status_code == 200:
                                html_pages.append(p_res.text)
                                logger.info("SSC page %s fetched OK | len=%s", i, len(p_res.text))
                            else:
                                logger.warning("SSC page %s failed HTTP %s. Stopping.", i, p_res.status_code)
                                break
                            await asyncio.sleep(1.5)  # prevent SSC rate-limit / early break
                        except Exception as e:
                            logger.error("Error fetching SSC page %s: %s", i, e)
                            break
                
                # 🌟 FIX: Merge pages properly (Grab ALL question tables directly)
                # ✅ Combine all pages in one clean pass (uses helper function)
                if len(html_pages) > 1:
                    all_html = build_combined_ssc_html(html_pages)
                else:
                    all_html = html_pages[0] if html_pages else ""
            else:
                async with httpx.AsyncClient(timeout=40.0, follow_redirects=True, verify=False) as client:
                    res = await client.get(
                        source_url,
                        headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
                    )
                    if res.status_code != 200:
                        await status_msg.edit_text(f"❌ Failed to download sheet (HTTP {res.status_code}).")
                        return
                    all_html = res.text
                    html_pages = [all_html]

        except Exception as e:
            await status_msg.edit_text(f"❌ Network Error: {e}")
            return

    elif is_document:
        # 🆕 RATE LIMIT for file uploads too
        try:
            if not await check_rate_limit(user_id):
                return await update.message.reply_text(
                    f"⏳ *Rate limit reached*\n\n"
                    f"Max {MAX_SUBMISSIONS_PER_HOUR} real submissions per hour.\n"
                    f"Please try again after an hour.",
                    parse_mode="Markdown",
                )
        except Exception as e:
            logger.warning("Rate limit check failed (allowing): %s", e)

        status_msg = await update.message.reply_text("⏳ Reading file...")
        try:
            doc = await context.bot.get_file(update.message.document.file_id)
            raw_b = await doc.download_as_bytearray()
            all_html = raw_b.decode("utf-8", errors="ignore")
            html_pages = [all_html]
        except Exception as e:
            await status_msg.edit_text(f"❌ Error: {e}")
            return
    else:
        await update.message.reply_text("❌ Please send a valid DigiALM/SSC response sheet URL or upload the HTML file.")
        return


    # 🌟 FIX: Convert all images to Base64 ONCE so PDF and HTML Test work offline
    # NOTE: image embedding is intentionally deferred.
    # PDFs call embed_all_images() internally; the Mock Test uses
    # embed_test_images_async() only on the extracted question data.
    logger.info("Parsing HTML (pre-embed) | len=%s", len(all_html))

    
    # =========================================================================
    # 3. SHEET PARSING & DYNAMIC SCORING
    # =========================================================================
    if "cbexams.com" in source_url or "SSC ONLINE EXAMINATION" in all_html:
        info, questions = parse_ssc_cbexams(html_pages)
    else:
        info, questions = parse_digialm_html(all_html, source_url)

    # 🌟 FINAL FIX: If Exam Name is exactly the Candidate Name, use the Subject instead
    if info.get("exam_name", "").strip().upper() == info.get("name", "").strip().upper():
        if info.get("branch"):
            info["exam_name"] = info["branch"]

    if not questions:
        await status_msg.edit_text("❌ Could not parse questions from this sheet. Start again /start and send Response URL.")
        return

    # 🆕 DUPLICATE CHECK (fail-open, 2-min window)
    # 🆕 DUPLICATE CHECK (10-min window — covers mock test processing time)
    try:
        recent = await check_recent_submission(user_id, info["hall_ticket"], source_url)
        if recent:
            secs = int(recent["age_seconds"])
            wait_left = max(0, 600 - secs)

            if secs < 60:
                age_str = f"{secs} sec"
            else:
                m = secs // 60
                s = secs % 60
                age_str = f"{m} min {s} sec" if s else f"{m} min"

            wait_min = wait_left // 60
            wait_sec = wait_left % 60
            if wait_min and wait_sec:
                wait_str = f"{wait_min} min {wait_sec} sec"
            elif wait_min:
                wait_str = f"{wait_min} min"
            else:
                wait_str = f"{wait_sec} sec"

            await status_msg.edit_text(
                f"ℹ️ *Already submitted {age_str} ago*\n\n"
                f"📊 Score: {recent['total_score']:.2f}\n"
                f"✅ Correct: {recent['correct']}  |  ❌ Wrong: {recent['incorrect']}\n"
                f"⏭️ Attempted: {recent['attempted']}\n\n"
                f"⏳ The bot is still processing this submission.\n"
                f"If your mock test hasn't arrived yet, please wait — "
                f"it will be delivered automatically.\n\n"
                f"🔄 To force regeneration, resend the link after **{wait_str}**.",
                parse_mode="Markdown",
            )
            return
    except Exception as e:
        logger.warning("Duplicate check failed (continuing): %s", e)
        # do nothing — continue normally
    # Strip URL parameters (?EncKey=...) to prevent random tokens triggering incorrect exams
    clean_url = source_url.split("?")[0].lower()
    exam_detection_text = info.get("exam_name", "").upper()

    # 🌟 STRICT URL CHECKS — order matters (RRB/SSC claimed first, GATE last)
    is_rrb_exam = (
        "rrb.digialm.com" in clean_url
        or "rrb" in clean_url
        or "railway" in clean_url
        or "RRB" in exam_detection_text
        or "RAILWAY" in exam_detection_text
    )
    is_ssc_exam = (
        "cbexams.com" in clean_url
        or "ssc" in clean_url
        or "SSC" in exam_detection_text
    )

    # 🌟 FIXED: GATE detection — STRICT ONLY
    #   - Must NOT be RRB
    #   - Must NOT be SSC
    #   - Must have GATE as a standalone word in exam name
    #     OR GATE<year digits> in URL (gate24, gate25, gate26)
    #   - REMOVED the dangerous "GATE in all_html" check that matched
    #     "navigate", "aggregate", "gateway", "investigate", etc.
    is_gate_exam = (
        not is_rrb_exam
        and not is_ssc_exam
        and (
            bool(re.search(r"\bGATE\b", exam_detection_text)) 
            or bool(re.search(r"gate\d{2}", clean_url))        
            or "GATE 20" in all_html.upper()   # 🌟 FIX: Matches the 2026 HTML banner
            or "GATE202" in all_html.upper()
        )
    )

    # 🌟 NEW: GATE ANSWER KEY CHECK (First time per branch)
    # 🌟 NEW: GATE UNIQUE SHIFT CHECK (Subject + Date + Time)
    if is_gate_exam:
        context.user_data['gate_source_url'] = source_url  # 🌟 ALWAYS update the link
        gate_branch = info.get("branch", "General")
        gate_date = info.get("test_date", "NA")
        gate_time = info.get("test_time", "NA")
        
        # Create a unique ID for this specific shift (e.g., "EE_ELECTRICAL_02-02-2025_2-30-PM")
        safe_branch = re.sub(r"[^a-zA-Z0-9]", "_", gate_branch)[:30]
        safe_date = re.sub(r"[^a-zA-Z0-9]", "-", gate_date)
        safe_time = re.sub(r"[^a-zA-Z0-9]", "-", gate_time)
        gate_uid = f"{safe_branch}_{safe_date}_{safe_time}"

        # 1. Exact match first (fast path)
        async with aiosqlite.connect(DB_FILE) as db:
            async with db.execute("SELECT keys_json FROM gate_keys WHERE gate_uid = ?", (gate_uid,)) as c:
                row = await c.fetchone()

            # 2. 🆕 Fuzzy match fallback — ignores _ vs - vs typos in separators
            if not row:
                async with db.execute("SELECT gate_uid, keys_json FROM gate_keys") as c:
                    all_keys = await c.fetchall()

                target_norm = _normalize_uid(gate_uid)
                for stored_uid, keys_json in all_keys:
                    if _normalize_uid(stored_uid) == target_norm:
                        row = (keys_json,)
                        logger.info(
                            "🔑 Fuzzy-matched GATE key: '%s' → '%s'",
                            gate_uid, stored_uid
                        )
                        break
        
        if not row:
            # First time for this specific shift! Ask for PDFs.
            pending_file = f"pending_gate_{user_id}.html"
            atomic_write_text(pending_file, all_html)
            
            context.user_data['gate_upload_step'] = 'key'
            context.user_data['gate_uid'] = gate_uid
            context.user_data['gate_branch'] = gate_branch
            context.user_data['gate_source_url'] = source_url  # 🌟 Save link for PDF forwarding
            
            await status_msg.edit_text(
                f"🔐 *GATE Answer Key Required*\n\n"
                f"To grade your score, I need the official **Answer Key PDF** for this shift.\n\n"
                f"📚 *{gate_branch}*\n"
                f"📅 *{gate_date} | {gate_time}*\n\n"
                f"Please upload the **Answer Key PDF** now.",
                parse_mode="Markdown"
            )
            return
            
        else:
            # 🌟 KEY FOUND! APPLY IT TO THE QUESTIONS USING MASTER Q.NO
            try:
                gate_data = json.loads(row[0])
                gate_keys = gate_data.get("keys", {})
                
                matched = 0
                for q in questions:
                    master_qno = str(q.get("master_qno", "")).strip()
                    
                    if master_qno and master_qno in gate_keys:
                        key_data = gate_keys[master_qno]
                        
                        # Handle both old string format and new dict format
                        if isinstance(key_data, dict):
                            key_val = str(key_data.get("key", "")).strip()
                            q["key_marks"] = int(key_data.get("marks", 1))
                        else:
                            key_val = str(key_data).strip()
                            q["key_marks"] = 1   # 🌟 legacy string keys default to 1 mark
                        
                        # 🌟 FIX: Handle "MTA" or "MTA*" (Marks To All) for GATE 2026
                        if key_val.upper() in ["MTA", "MTA*"]:
                            q["full_marks_note"] = True
                            q["correct"] = "MTA"
                        # Handle NAT Ranges (e.g., "7 to 71", "-2.00 to -1.94 OR 1.94 to 2.00")
                        elif "to" in key_val:
                            # Just store the range string. The scoring loop will evaluate it.
                            q["correct"] = key_val
                        else:
                            # 🌟 FIX: Convert key letters (A/B/C/D) → numbers (1/2/3/4)
                            # because HTML parser stores chosen answers as numbers.
                            letter_to_num = {"A": "1", "B": "2", "C": "3", "D": "4"}
                            
                            # MSQ: "A;C" or "A,C" → "1,3"
                            if ";" in key_val or "," in key_val or ":" in key_val:
                                letters = re.split(r"[;,]", key_val)
                                numbers = [letter_to_num.get(l.strip().upper(), l.strip()) for l in letters if l.strip()]
                                q["correct"] = ",".join(sorted(numbers))
                            else:
                                # MCQ or single NAT: "A" → "1", "7" → "7"
                                single = key_val.strip().upper()
                                q["correct"] = letter_to_num.get(single, single)
                            
                        matched += 1
                
                if matched == 0:
                    logger.warning("GATE Key found, but 0 questions matched. Check Master Q.No parsing.")
                else:
                    logger.info(f"GATE Key applied: {matched}/{len(questions)} questions matched.")
                    
            except Exception as e:
                logger.error(f"Failed to apply GATE keys: {e}")
    
    # 🌟 SPECIFIC SSC EXAMS (Only checked if it's SSC)
    is_je = is_ssc_exam and bool(re.search(r"\bJE\b|JUNIOR ENGINEER", exam_detection_text))
    is_cgl = is_ssc_exam and bool(re.search(r"\bCGL\b|COMBINED GRADUATE", exam_detection_text))
    is_chsl = is_ssc_exam and bool(re.search(r"\bCHSL\b|COMBINED HIGHER SECONDARY|10\+2", exam_detection_text))
    is_mts = is_ssc_exam and bool(re.search(r"\bMTS\b|MULTI[\s\-]*TASKING", exam_detection_text))
    is_cpo = is_ssc_exam and bool(re.search(r"\bCPO\b|CENTRAL POLICE|SI IN DELHI", exam_detection_text))
    is_gd = is_ssc_exam and bool(re.search(r"\bGD\b|CONSTABLE[\s\-]*GD|GENERAL DUTY", exam_detection_text))
    is_steno = is_ssc_exam and "STENOGRAPHER" in exam_detection_text
    is_jht = is_ssc_exam and bool(re.search(r"\bJHT\b|HINDI TRANSLATOR", exam_detection_text))
    is_sel_post = is_ssc_exam and "SELECTION POST" in exam_detection_text

    section_stats = {}
    total_correct = total_incorrect = total_attempted = 0
    total_score = 0.0
    total_max_marks = 0.0

    for q in questions:
        sec = q["section"]
        sec_upper = sec.upper()
        pnl_txt = q.get("panel_text", "").upper()

        q_pos, q_neg = 1.0, 0.0 # Default fallback

        if is_gate_exam:
            # 🌟 FIX: Use exact marks from the PDF Key if available!
            if "key_marks" in q:
                q_pos = float(q["key_marks"])
            else:
                # Fallback to guessing (only if key_marks is missing)
                master_qno = q.get("master_qno", "")
                if master_qno and str(master_qno).isdigit():
                    q_num = int(master_qno)
                    is_ga = "GENERAL APTITUDE" in sec_upper or "GA" in sec_upper
                    if is_ga:
                        q_pos = 2.0 if q_num > 5 else 1.0
                    else:
                        q_pos = 2.0 if q_num > 35 else 1.0
                else:
                    is_two_mark = ("2 MARK" in sec_upper or "2M" in sec_upper or "MARKS : 2" in pnl_txt or "MARKS:2" in pnl_txt)
                    q_pos = 2.0 if is_two_mark else 1.0
            
            is_nat_or_msq = ("NAT" in sec_upper or "MSQ" in sec_upper or "QUESTION TYPE : NAT" in pnl_txt or "QUESTION TYPE : MSQ" in pnl_txt)
            q_neg = 0.0 if is_nat_or_msq else (q_pos / 3.0)

        elif is_cgl:
            if "TIER 2" in exam_detection_text or "TIER-2" in exam_detection_text or "TIER 2" in sec_upper or "TIER-2" in sec_upper:
                if "STATISTICS" in sec_upper or "JSO" in sec_upper:
                    q_pos, q_neg = 2.0, 0.50
                elif "COMPUTER" in sec_upper or "DEST" in sec_upper:
                    q_pos, q_neg = 1.0, 0.0  # Qualifying
                else:
                    q_pos, q_neg = 3.0, 1.0
            else:
                q_pos, q_neg = 2.0, 0.50

        elif is_chsl:
            if "TIER 2" in exam_detection_text or "TIER-2" in exam_detection_text or "TIER 2" in sec_upper or "SESSION-1" in sec_upper:
                q_pos, q_neg = 3.0, 1.0
            else:
                q_pos, q_neg = 2.0, 0.50

        elif is_rrb_exam:
            q_pos, q_neg = 1.0, 0.333333  # 🌟 RRB JE and all RRB exams get -1/3 (0.33)

        elif is_je:
            if "PAPER 2" in exam_detection_text or "PAPER-2" in exam_detection_text or "PAPER 2" in sec_upper or "PAPER-2" in sec_upper:
                q_pos, q_neg = 3.0, 1.0
            else:
                q_pos, q_neg = 1.0, 0.25  # 🌟 SSC JE stays -0.25

        elif is_mts:
            is_session_2 = (
                "SESSION-II" in sec_upper or "SESSION 2" in sec_upper or "SESSION-2" in sec_upper or
                "GENERAL AWARENESS" in sec_upper or "ENGLISH" in sec_upper or
                "GENERAL AWARENESS" in pnl_txt or "ENGLISH" in pnl_txt
            )
            q_pos, q_neg = (3.0, 1.0) if is_session_2 else (3.0, 0.0)

        elif is_cpo:
            q_pos, q_neg = 1.0, 0.25
        elif is_gd:
            q_pos, q_neg = 2.0, 0.25
        elif is_steno or is_jht:
            q_pos, q_neg = 1.0, 0.25
        elif is_sel_post:
            q_pos, q_neg = 2.0, 0.50
        else:
            for key, scheme in EXAM_MARKING_SCHEMES.items():
                if key in exam_detection_text or key in sec_upper:
                    q_pos, q_neg = scheme
                    break

        total_max_marks += q_pos

        # 🌟 FIX: Save the exact calculated marks to the question object
        q["pos_marks"] = q_pos
        q["neg_marks"] = q_neg

        if sec not in section_stats:
            section_stats[sec] = {"total": 0, "attempted": 0, "correct": 0, "wrong": 0, "score": 0.0}
        section_stats[sec]["total"] += 1

        # --- 2. Score Calculation ---
        if q.get("full_marks_note", False):
            total_correct += 1
            total_score += q_pos
            section_stats[sec]["correct"] += 1
            section_stats[sec]["score"] += q_pos
            if q["chosen"] not in ["--", ""]:
                total_attempted += 1
                section_stats[sec]["attempted"] += 1

        elif q.get("ambiguity_opts", []) and str(q["chosen"]) in q["ambiguity_opts"]:
            total_attempted += 1
            section_stats[sec]["attempted"] += 1
            total_correct += 1
            total_score += q_pos
            section_stats[sec]["correct"] += 1
            section_stats[sec]["score"] += q_pos

        elif q["chosen"] not in ["--", ""]:
            total_attempted += 1
            section_stats[sec]["attempted"] += 1
            
            is_correct_q = False
            
            # 🌟 FIX: Proper numeric evaluation for GATE NAT questions
            if is_gate_exam and q.get("q_type") == "NAT":
                try:
                    expected = str(q["correct"]).strip()
                    given = str(q["chosen"]).strip()
                    if given != "" and given != "--":
                        if "to" in expected:
                            # Handle ranges (e.g., "7 to 7.1" or "-2.00 to -1.94 OR 1.94 to 2.00")
                            ranges = expected.split(" OR ")
                            for r in ranges:
                                if "to" in r:
                                    low, high = [float(x.strip()) for x in r.split("to")]
                                    if low <= float(given) <= high:
                                        is_correct_q = True
                                        break
                        else:
                            # Handle exact numbers with 0.01 tolerance
                            exp_val = float(expected)
                            giv_val = float(given)
                            if abs(exp_val - giv_val) < 0.01:
                                is_correct_q = True
                except ValueError:
                    # Fallback to string compare if float fails
                    if str(q["chosen"]) == str(q["correct"]):
                        is_correct_q = True
            else:
                # Standard string compare for MCQ/MSQ and other exams
                if q["correct"] and str(q["chosen"]) == str(q["correct"]):
                    is_correct_q = True
                    
            # logger.info(f"📝 SCORING Q#{q.get('master_qno')}: HTML Chosen='{q['chosen']}', PDF Key='{q['correct']}', Result={'CORRECT' if is_correct_q else 'WRONG'}")
            
            if is_correct_q:
                total_correct += 1
                total_score += q_pos
                section_stats[sec]["correct"] += 1
                section_stats[sec]["score"] += q_pos
            else:
                total_incorrect += 1
                total_score -= q_neg
                section_stats[sec]["wrong"] += 1
                section_stats[sec]["score"] -= q_neg

    # =========================================================================
    # 4. DATABASE SAVE (fast, necessary, keep synchronous)
    # =========================================================================
    async with db_lock:
        async with aiosqlite.connect(DB_FILE) as db:
            await db.execute("""
                INSERT OR REPLACE INTO candidates 
                (user_id, exam_name, hall_ticket, candidate_name, branch, total_score, total_questions, attempted, correct, incorrect, test_date, test_time, test_center, response_url, timestamp)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                user_id, info["exam_name"], info["hall_ticket"], info["name"], info["branch"],
                round(total_score, 2), len(questions), total_attempted, total_correct, total_incorrect,
                info["test_date"], info["test_time"], info["test_center"], source_url,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ))
            await db.execute("""
                INSERT INTO submission_logs
                (user_id, user_full_name, username, exam_name, hall_ticket, candidate_name, response_url, submitted_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                user_id, full_name, username_str, info["exam_name"], info["hall_ticket"], info["name"],
                source_url, datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ))
            await db.commit()

    # =========================================================================
    # 5. SAVE RAW HTML + COMPUTE ALL PATHS (fast, local disk)
    # =========================================================================
    safe_exam = re.sub(r"[^a-zA-Z0-9_-]", "_", info["exam_name"])
    safe_ht   = re.sub(r"[^a-zA-Z0-9_-]", "_", info["hall_ticket"])
    safe_name = re.sub(r"[^a-zA-Z0-9_-]", "_", info["name"])

    exam_html_dir = os.path.join(STORAGE_DIR, safe_exam)
    exam_pdf_dir  = os.path.join(PDF_STORAGE_DIR, safe_exam)
    os.makedirs(exam_html_dir, exist_ok=True)
    os.makedirs(exam_pdf_dir, exist_ok=True)

    html_file = os.path.join(exam_html_dir, f"{safe_ht}_{safe_name}.html")
    pdf_file  = os.path.join(exam_pdf_dir,  f"{safe_ht}_{safe_name}.pdf")

    atomic_write_text(html_file, all_html)

    base_url = source_url.rsplit("/", 1)[0] + "/" if "/" in source_url.split("?")[0] else ""

    if "cbexams.com" in source_url and len(html_pages) > 1:
        all_html_for_pdf = all_html
        logger.info("SSC PDF source = combined HTML (len=%s)", len(all_html_for_pdf))
    else:
        all_html_for_pdf = all_html

    # Pre-compute master key/path (used by background worker)
    # 🌟 Master paths
    master = get_master_paths(info)
    master_key = master["key"]
    master_pdf_path = master["pdf_path"]
    master_test_path = master["test_path"]
    master_sheet_path = master["sheet_path"]

    # =========================================================================
    # 6. ⚡ GENERATE SCORECARD FIRST (top priority — no PDF competition)
    # =========================================================================
    t_scorecard_start = datetime.now()
    img_path = await generate_scorecard_image(
        candidate_info=info,
        section_stats=section_stats,
        total_score=total_score,
        max_marks=total_max_marks,
        total_questions=len(questions),
        attempted=total_attempted,
        correct=total_correct,
        wrong=total_incorrect,
        header_html=info["header_html"],
        base_url=base_url,
    )
    scorecard_secs = (datetime.now() - t_scorecard_start).total_seconds()

    # Terminal evaluation summary (log-only, no blocking)
    accuracy = (total_correct / total_attempted * 100) if total_attempted > 0 else 0.0
    print("\n" + "═" * 75)
    print("📥 CANDIDATE EVALUATION COMPLETE")
    print("═" * 75)
    print(f"👤 User ID     : {user_id} ({full_name} | {username_str})")
    print(f"🔗 Input Link  : {source_url}")
    print(f"🏛 Exam Name   : {info['exam_name']}")
    print(f"🧑 Candidate   : {info['name']}")
    print(f"🏷 Roll / ID   : {info['hall_ticket']}")
    print(f"⚙️ Branch/Sub  : {info['branch']}")
    print(f"📅 Test Date   : {info.get('test_date', 'N/A')}  |  ⏰ Test Time: {info.get('test_time', 'N/A')}")
    print(f"📊 Total Marks : {total_score:.2f} / {total_max_marks:.2f} (Attempted: {total_attempted}/{len(questions)})")
    print(f"🎯 Accuracy    : {accuracy:.1f}%")
    print(f"🖼️ Scorecard   : {img_path}  (rendered in {scorecard_secs:.1f}s)")
    print("─" * 75)
    print("📋 SECTION BREAKDOWN:")
    for s_name, s in section_stats.items():
        print(f"  • {s_name:<36} : {s['score']:>6.2f} Marks | R: {s['correct']:>2} | W: {s['wrong']:>2} | NA: {s['total'] - s['attempted']:>2}")
    print("═" * 75 + "\n")

    # =========================================================================
    # 7. ⚡ SEND SCORECARD IMMEDIATELY (before queuing anything else)
    # =========================================================================
    photo_caption = (
        f"📸 *Scorecard Snapshot* @KranthikumargoudEEE\n\n"
        f"📢 *Join our channel for latest updates:*\n"
        f"👉 [Click Here to Join Telegram Channel]({CHANNEL_LINK})\n"
        f"👉 [Click Here to Join Whatsapp Channel]({CHANNEL_LINK1})\n"
        f"{CHANNEL_LINK}"
    )

    if not (img_path and os.path.exists(img_path)):
        await status_msg.edit_text("❌ Failed to generate scorecard image.")
        return

    sent = False
    for attempt in range(3):
        try:
            with open(img_path, "rb") as photo:
                await update.message.reply_photo(photo=photo, caption=photo_caption, parse_mode="Markdown", write_timeout=120.0)
            sent = True
            break
        except RetryAfter as e:
            await asyncio.sleep(e.retry_after + 1)
        except (TimedOut, NetworkError) as e:
            logger.warning("Scorecard send attempt %s failed: %s", attempt + 1, e)
            await asyncio.sleep(2)

    if not sent:
        await notify_genuine_error(
            context, "Scorecard image delivery", user_id, full_name, username_str, info, source_url,
            RuntimeError("Scorecard upload failed after 3 attempts")
        )
        await queue_failed_delivery(context, user_id, "send_photo", img_path, photo_caption, "Markdown", None, user_id, full_name, source_url)
        try:
            await status_msg.edit_text(f"⚠️ Scorecard upload failed (network).\n\nResend after 10 mins:\n<code>{html.escape(source_url)}</code>", parse_mode="HTML")
        except Exception: pass
        return
    logger.info("⚡ Scorecard delivered to user in %.1fs", scorecard_secs)

    # 🌟 ADMIN COPY — Scorecard (ADMIN_IDS DMs + all ADMIN_COPY_CHAT_IDS channels)

    if SEND_COPY_TO_ADMIN and img_path and os.path.exists(img_path):
        admin_caption = (
            f"📸 <b>Admin Copy — Scorecard</b>\n"
            f"👤 <b>From:</b> {_h(full_name)} (ID: <code>{user_id}</code>)\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"🏛 <b>Exam:</b> {_h(info['exam_name'])}\n"
            f"📅 <b>Date:</b> {_h(info.get('test_date','N/A'))}\n"
            f"⏰ <b>Shift / Time:</b> {_h(info.get('test_time','N/A'))}\n"
            f"📚 <b>Subject:</b> {_h(info.get('branch','General'))}\n"
            f"🧑 <b>Candidate:</b> {_h(info['name'])}\n"
            f"🏷 <b>Roll No:</b> {_h(info['hall_ticket'])}\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"📊 <b>Score:</b> {total_score:.2f} / {total_max_marks:.2f}\n"
            f"✅ <b>Correct:</b> {total_correct}  |  ❌ <b>Wrong:</b> {total_incorrect}\n"
            f"⏭️ <b>Attempted:</b> {total_attempted} / {len(questions)}\n"
            f"🎯 <b>Accuracy:</b> {(total_correct / total_attempted * 100) if total_attempted else 0:.1f}%\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"🔗 <b>Sheet URL:</b>\n<code>{html.escape(source_url)}</code>"
        )
        for target_chat, thread_id in _get_all_copy_targets(exclude_user_id=user_id):
            try:
                with open(img_path, "rb") as photo:
                    kwargs = dict(
                        chat_id=target_chat,
                        photo=photo,
                        caption=admin_caption,
                        parse_mode="HTML",
                    )
                    if thread_id is not None:
                        kwargs["message_thread_id"] = thread_id
                    await context.bot.send_photo(**kwargs)
            except Exception as e:
                logger.warning("Admin scorecard copy failed for %s: %s", target_chat, e)
                await queue_failed_delivery(context, target_chat, "send_photo", img_path, admin_caption, "HTML", thread_id, user_id, full_name, source_url)

    # =========================================================================
    # 8. 🧵 QUEUE ALL SLOW WORK AS BACKGROUND TASKS (do not block handler)
    # =========================================================================

    # --- 8a. PDF worker: user PDF, then master PDF (if not cached) ---
    async def _pdf_worker():
        # 🌟 INDIVIDUAL PDF (+ optional admin copy)
        if SAVE_INDIVIDUAL_PDF:
            try:
                ok = await save_sheet_as_pdf(all_html_for_pdf, pdf_file, base_url, is_master=False)

                # 🌟 ADMIN COPY — Individual PDF (only if SEND_PDF_COPY_TO_ADMIN is ON)
                if ok and os.path.exists(pdf_file) and SEND_COPY_TO_ADMIN and SEND_PDF_COPY_TO_ADMIN:
                    admin_pdf_caption = (
                        f"📄 <b>Admin Copy — Individual PDF</b>\n"
                        f"👤 User: {_h(full_name)} (ID: <code>{user_id}</code>)\n"
                        f"🏛 Exam: {_h(info['exam_name'])}\n"
                        f"🧑 Candidate: {_h(info['name'])}\n"
                        f"🏷 Roll: {_h(info['hall_ticket'])}"
                    )
                    for target_chat, thread_id in _get_all_copy_targets(exclude_user_id=user_id):
                        try:
                            kwargs = dict(
                                chat_id=target_chat,
                                filename=os.path.basename(pdf_file),
                                caption=admin_pdf_caption,
                                parse_mode="HTML",
                            )
                            if thread_id is not None:
                                kwargs["message_thread_id"] = thread_id
                            with open(pdf_file, "rb") as f:
                                kwargs["document"] = f
                                await context.bot.send_document(**kwargs)
                        except Exception as e:
                            logger.warning("Admin PDF copy failed for %s: %s", target_chat, e)
                            await queue_failed_delivery(context, target_chat, "send_document", pdf_file, admin_pdf_caption, "HTML", thread_id, user_id, full_name, source_url)
                elif SEND_PDF_COPY_TO_ADMIN is False:
                    logger.info("⏩ PDF copy disabled (SEND_PDF_COPY_TO_ADMIN = False) — saves data")

            except Exception as e:
                logger.error("User PDF failed: %s", e, exc_info=True)
        else:
            logger.info("⏩ Skipping individual PDF generation (SAVE_INDIVIDUAL_PDF = False)")

        # Master PDF generation still runs perfectly in the background!
        # =====================================================================
        # 🌟 MASTER PDF
        # =====================================================================
        if SAVE_MASTER_PDF and not os.path.exists(master_pdf_path) and master_key not in generating_masters:
            generating_masters.add(master_key)
            try:
                logger.info("🌟 Generating NEW Master PDF: %s", master_pdf_path)
                ok = await save_sheet_as_pdf(all_html_for_pdf, master_pdf_path, base_url, is_master=True, info=info)
                if ok and os.path.exists(master_pdf_path):
                    logger.info("✅ Master PDF saved: %s", master_pdf_path)
                    # 🌟 SEPARATE FLAG: SEND_MASTER_PDF_TO_ADMIN
                    if SEND_MASTER_PDF_TO_ADMIN and SEND_COPY_TO_ADMIN:
                        admin_caption = f"📚 <b>Master PDF</b>\n🏛 {_h(info.get('exam_name','N/A'))}\n📅 {_h(info.get('test_date','N/A'))} | ⏰ {_h(info.get('test_time','N/A'))}\n🔗 <code>{html.escape(source_url)}</code>"
                        for target_chat, thread_id in _get_all_copy_targets(exclude_user_id=user_id):
                            try:
                                with open(master_pdf_path, "rb") as f:
                                    kwargs = dict(chat_id=target_chat, document=f, filename=os.path.basename(master_pdf_path), caption=admin_caption, parse_mode="HTML")
                                    if thread_id is not None: kwargs["message_thread_id"] = thread_id
                                    await context.bot.send_document(**kwargs)
                            except Exception as e: 
                                logger.warning("Admin Master PDF copy failed: %s", e)
                                await queue_failed_delivery(context, target_chat, "send_document", master_pdf_path, admin_caption, "HTML", thread_id, user_id, full_name, source_url)
            except Exception as e:
                logger.error("Master PDF failed: %s", e)
            finally:
                generating_masters.discard(master_key)

        # =====================================================================
        # 🌟 MASTER RAW SHEET
        # =====================================================================
        if SAVE_MASTER_RAW_SHEET and not os.path.exists(master_sheet_path):
            try:
                atomic_write_text(master_sheet_path, all_html_for_pdf)
                # 🌟 SEPARATE FLAG: SEND_MASTER_SHEET_TO_ADMIN
                if SEND_MASTER_SHEET_TO_ADMIN and SEND_COPY_TO_ADMIN:
                    admin_caption = f"🗂️ <b>Master Raw Sheet</b>\n🏛 {_h(info.get('exam_name','N/A'))}\n📅 {_h(info.get('test_date','N/A'))} | ⏰ {_h(info.get('test_time','N/A'))}\n🔗 <code>{html.escape(source_url)}</code>"
                    for target_chat, thread_id in _get_all_copy_targets(exclude_user_id=user_id):
                        try:
                            with open(master_sheet_path, "rb") as f:
                                kwargs = dict(chat_id=target_chat, document=f, filename=os.path.basename(master_sheet_path), caption=admin_caption, parse_mode="HTML")
                                if thread_id is not None: kwargs["message_thread_id"] = thread_id
                                await context.bot.send_document(**kwargs)
                        except Exception as e: 
                            logger.warning("Admin Master Raw Sheet copy failed: %s", e)
                            await queue_failed_delivery(context, target_chat, "send_document", master_sheet_path, admin_caption, "HTML", thread_id, user_id, full_name, source_url)
            except Exception as e: logger.error("Master raw sheet save failed: %s", e)

    # --- 8b. Excel sync worker (independent of scorecard) ---
    async def _excel_worker():
        try:
            await update_excel_sheet()
        except Exception as e:
            logger.error("update_excel_sheet failed: %s", e)
        try:
            await update_user_history_excel()
        except Exception as e:
            logger.error("update_user_history_excel failed: %s", e)

    # --- 8c. Mock Test worker (generates + sends HTML to user) ---
    # --- 8c. Mock Test worker (generates + sends HTML to user) ---
    async def _mock_test_worker():
        # 🌟 FIX: Define all safe names ONCE at the very top.
        # This prevents the "UnboundLocalError" and avoids duplicates entirely.
        safe_exam_t = re.sub(r"[^a-zA-Z0-9_-]", "_", info.get('exam_name', 'Exam'))[:40]
        safe_date_t = re.sub(r"[^a-zA-Z0-9_-]", "_", str(info.get("test_date", "NA")))
        safe_time_t = re.sub(r"[^a-zA-Z0-9_-]", "_", str(info.get("test_time", "NA")))
        safe_subj_t = re.sub(r"[^a-zA-Z0-9_-]", "_", str(info.get("branch", "General")))[:25]
        safe_ht = re.sub(r"[^a-zA-Z0-9_-]", "_", info.get("hall_ticket", "NA"))

        try:
            # =================================================================
            # 🌟 SMART CACHE CHECK
            # =================================================================
            if SAVE_MASTER_TEST_HTML and os.path.exists(master_test_path):
                logger.info("⚡ Cache hit! Reusing master for user %s", user_id)
                if SEND_MOCK_TEST_TO_USER:
                    await status_msg.edit_text("⚡ Sending your interactive Mock Test (cached)...")
                    def _md(s): return str(s).replace("*", "").replace("_", " ").replace("`", "").replace("[", "(").replace("]", ")")
                    caption = (
                        f"📝 *Interactive Mock Test Generated!*\n"
                        f"━━━━━━━━━━━━━━━━━━━━\n"
                        f"🏛 *Exam:* {_md(info.get('exam_name', 'Mock Test'))}\n"
                        f"📅 *Date:* {_md(info.get('test_date', 'N/A'))}\n"
                        f"⏰ *Shift / Time:* {_md(info.get('test_time', 'N/A'))}\n"
                        f"📚 *Subject:* {_md(info.get('branch', 'General'))}\n"
                        f"🧑 *Candidate:* {_md(info.get('name', 'Candidate'))}\n"
                        f"━━━━━━━━━━━━━━━━━━━━\n"
                        f"✅ Full Test or Sectional Timers (Customizable)\n"
                        f"✅ Fully Offline (Open in Chrome/Safari)\n\n"
                        f"Powered by @KranthikumargoudEEE"
                    )
                    try:
                        with open(master_test_path, "rb") as test_file:
                            await update.message.reply_document(
                                document=test_file, 
                                filename=f"MockTest_{safe_date_t}_{safe_time_t}_{safe_ht}.html",
                                caption=caption, 
                                parse_mode="Markdown"
                            )
                    except Exception as e:
                        await queue_failed_delivery(context, user_id, "send_document", master_test_path, caption, "Markdown", None, user_id, full_name, source_url)
                        raise e
                    try: await status_msg.delete()
                    except Exception: pass
                return # EXIT EARLY

            if SEND_MOCK_TEST_TO_USER:
                await status_msg.edit_text("⏳ Generating your interactive Mock Test...")

            is_ssc_test = "cbexams.com" in source_url or "SSC ONLINE EXAMINATION" in all_html
            test_data = (
                extract_test_data(html_pages, True, questions, info.get("exam_name", ""))
                if is_ssc_test else
                extract_test_data(all_html, False, questions, info.get("exam_name", ""))
            )
            logger.info("Test data extracted: %s questions (ssc=%s)", len(test_data), is_ssc_test)

            if not test_data:
                if SEND_MOCK_TEST_TO_USER:
                    await status_msg.delete()
                logger.warning("No questions parsed for HTML test.")
                return

            # ⚡ Step 1: Embed images so everything is base64 (self-contained)
            await embed_test_images_async(test_data, base_url, source_url)
            logger.info("Test images embedded")

            # ⚡ Step 2: Embed header banner too (for offline use)
            test_header_html = await embed_all_images(info.get("header_html", ""), base_url)

            _hdr_soup = BeautifulSoup(test_header_html, "html.parser")
            for bad in _hdr_soup.find_all("img"):
                s = (bad.get("src") or "").strip().strip("'\"").lower()
                if any(k in s for k in ("photo", "signature", "thumb")):
                    bad.decompose()
            test_header_html = str(_hdr_soup)

            info["test_header_html"] = test_header_html

            # ============================================================
            # 📦 Step 3a: Save SMALL JSON (meta + info + test_data)
            # ============================================================
            if SAVE_MASTER_JSON:
                try:
                    small_payload = {
                        "schema_version": "2.0",
                        "json_variant":   "small",
                        "meta": {
                            "generated_at":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "user_id":         user_id,
                            "user_full_name":  full_name,
                            "user_username":   username_str,
                            "source_url":      source_url,
                            "exam_name":       info.get("exam_name"),
                            "branch":          info.get("branch"),
                            "test_date":       info.get("test_date"),
                            "test_time":       info.get("test_time"),
                            "test_center":     info.get("test_center"),
                            "candidate_name":  info.get("name"),
                            "hall_ticket":     info.get("hall_ticket"),
                            "total_questions": len(test_data),
                            "master_key":      master_key,
                        },
                        "info": {
                            "exam_name":        info.get("exam_name"),
                            "branch":           info.get("branch"),
                            "test_date":        info.get("test_date"),
                            "test_time":        info.get("test_time"),
                            "test_center":      info.get("test_center"),
                            "name":             info.get("name"),
                            "hall_ticket":      info.get("hall_ticket"),
                            "header_html":      info.get("header_html", ""),
                            "test_header_html": info.get("test_header_html", ""),
                        },
                        "test_data": test_data,
                    }

                    json_path = master["json_path"]
                    atomic_write_text(
                        json_path,
                        json.dumps(small_payload, ensure_ascii=False, indent=2)
                    )
                    file_size_mb = os.path.getsize(json_path) / (1024 * 1024)
                    logger.info("📦 Master JSON (small) saved: %s (%.2f MB)", json_path, file_size_mb)

                    if SEND_MASTER_JSON_TO_ADMIN and SEND_COPY_TO_ADMIN:
                        json_caption = (
                            f"📦 <b>Master JSON — SMALL (Editable)</b>\n"
                            f"━━━━━━━━━━━━━━━━━━━━\n"
                            f"🏛 <b>Exam:</b> {_h(info.get('exam_name','N/A'))}\n"
                            f"📚 <b>Subject:</b> {_h(info.get('branch','General'))}\n"
                            f"📅 <b>Date:</b> {_h(info.get('test_date','N/A'))}\n"
                            f"⏰ <b>Shift:</b> {_h(info.get('test_time','N/A'))}\n"
                            f"🧑 <b>Candidate:</b> {_h(info.get('name','N/A'))}\n"
                            f"🏷 <b>Roll No:</b> {_h(info.get('hall_ticket','N/A'))}\n"
                            f"━━━━━━━━━━━━━━━━━━━━\n"
                            f"📊 <b>Questions:</b> {len(test_data)}\n"
                            f"💾 <b>Size:</b> {file_size_mb:.2f} MB\n"
                            f"✏️ <b>Edit</b> <code>solution</code> → <code>/regen_from_json</code>\n"
                            f"━━━━━━━━━━━━━━━━━━━━\n"
                            f"👤 {_h(full_name)} (ID: <code>{user_id}</code>)\n"
                            f"🔗 <b>Source:</b>\n<code>{html.escape(source_url)}</code>"
                        )
                        for target_chat, thread_id in _get_all_copy_targets(exclude_user_id=user_id):
                            try:
                                with open(json_path, "rb") as jf:
                                    kwargs = dict(
                                        chat_id=target_chat,
                                        document=jf,
                                        filename=os.path.basename(json_path),
                                        caption=json_caption,
                                        parse_mode="HTML",
                                    )
                                    if thread_id is not None:
                                        kwargs["message_thread_id"] = thread_id
                                    await context.bot.send_document(**kwargs)
                                await asyncio.sleep(0.3)
                            except Exception as e:
                                logger.warning("Admin small JSON send failed: %s", e)
                                await queue_failed_delivery(
                                    context, target_chat, "send_document", json_path,
                                    json_caption, "HTML", thread_id, user_id, full_name, source_url
                                )
                except Exception as e:
                    logger.error("Small JSON save/send failed: %s", e)

            # ============================================================
            # 📦 Step 3b: Save FULL JSON (everything + raw_source_html)
            #            — keeps file on failure, pings admin with path
            # ============================================================
            if SAVE_MASTER_JSON_FULL or SEND_MASTER_JSON_FULL_TO_ADMIN:
                try:
                    full_payload = {
                        "schema_version": "2.0",
                        "json_variant":   "full",
                        "meta": {
                            "generated_at":    datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            "user_id":         user_id,
                            "user_full_name":  full_name,
                            "user_username":   username_str,
                            "source_url":      source_url,
                            "exam_name":       info.get("exam_name"),
                            "branch":          info.get("branch"),
                            "test_date":       info.get("test_date"),
                            "test_time":       info.get("test_time"),
                            "test_center":     info.get("test_center"),
                            "candidate_name":  info.get("name"),
                            "hall_ticket":     info.get("hall_ticket"),
                            "total_questions": len(test_data),
                            "master_key":      master_key,
                            "raw_source_html_length": len(all_html_for_pdf),
                        },
                        "info": {
                            "exam_name":        info.get("exam_name"),
                            "branch":           info.get("branch"),
                            "test_date":        info.get("test_date"),
                            "test_time":        info.get("test_time"),
                            "test_center":      info.get("test_center"),
                            "name":             info.get("name"),
                            "hall_ticket":      info.get("hall_ticket"),
                            "header_html":      info.get("header_html", ""),
                            "test_header_html": info.get("test_header_html", ""),
                        },
                        "test_data":       test_data,
                        "raw_source_html": all_html_for_pdf,
                    }

                    full_json_str = json.dumps(full_payload, ensure_ascii=False, indent=2)

                    # Path decision: permanent (SAVE=True) vs temp (SAVE=False)
                    is_temp = not SAVE_MASTER_JSON_FULL
                    if SAVE_MASTER_JSON_FULL:
                        json_full_path = master["json_full_path"]
                    else:
                        json_full_path = os.path.join(
                            MASTER_JSON_DIR, f".tmp_{master_key}_full.json"
                        )

                    atomic_write_text(json_full_path, full_json_str)
                    file_size_mb = os.path.getsize(json_full_path) / (1024 * 1024)
                    logger.info(
                        "📦 Master JSON (FULL) %s: %s (%.2f MB)",
                        "temp" if is_temp else "saved", json_full_path, file_size_mb
                    )

                    # Upload to admin
                    failed_targets = []
                    if SEND_MASTER_JSON_FULL_TO_ADMIN and SEND_COPY_TO_ADMIN:
                        json_caption = (
                            f"📦 <b>Master JSON — FULL (Archive)</b>\n"
                            f"━━━━━━━━━━━━━━━━━━━━\n"
                            f"🏛 <b>Exam:</b> {_h(info.get('exam_name','N/A'))}\n"
                            f"📚 <b>Subject:</b> {_h(info.get('branch','General'))}\n"
                            f"📅 <b>Date:</b> {_h(info.get('test_date','N/A'))}\n"
                            f"⏰ <b>Shift:</b> {_h(info.get('test_time','N/A'))}\n"
                            f"🧑 <b>Candidate:</b> {_h(info.get('name','N/A'))}\n"
                            f"🏷 <b>Roll No:</b> {_h(info.get('hall_ticket','N/A'))}\n"
                            f"━━━━━━━━━━━━━━━━━━━━\n"
                            f"📊 <b>Questions:</b> {len(test_data)}\n"
                            f"🗂️ <b>Raw HTML:</b> {len(all_html_for_pdf):,} chars\n"
                            f"💾 <b>Size:</b> {file_size_mb:.2f} MB\n"
                            f"♻️ <b>Full recovery</b> — reparse anytime\n"
                            f"━━━━━━━━━━━━━━━━━━━━\n"
                            f"👤 {_h(full_name)} (ID: <code>{user_id}</code>)\n"
                            f"🔗 <b>Source:</b>\n<code>{html.escape(source_url)}</code>"
                        )
                        display_name = os.path.basename(master["json_full_path"])
                        for target_chat, thread_id in _get_all_copy_targets(exclude_user_id=user_id):
                            try:
                                with open(json_full_path, "rb") as jf:
                                    kwargs = dict(
                                        chat_id=target_chat,
                                        document=jf,
                                        filename=display_name,
                                        caption=json_caption,
                                        parse_mode="HTML",
                                    )
                                    if thread_id is not None:
                                        kwargs["message_thread_id"] = thread_id
                                    await context.bot.send_document(**kwargs)
                                logger.info("✅ Full JSON sent → %s", target_chat)
                                await asyncio.sleep(0.3)
                            except Exception as e:
                                logger.warning("❌ Full JSON send failed → %s: %s", target_chat, e)
                                failed_targets.append((target_chat, thread_id, str(e)[:200]))

                    # ─── Handle result: on fail → promote + ping + keep ───
                    final_path = json_full_path
                    if failed_targets:
                        if is_temp:
                            failed_dir = os.path.join(MASTER_JSON_DIR, "_FAILED_MANUAL_UPLOADS")
                            os.makedirs(failed_dir, exist_ok=True)
                            manual_path = os.path.join(failed_dir, f"{master_key}_full.json")
                            try:
                                os.replace(json_full_path, manual_path)
                                logger.warning("⚠️ Promoted failed temp → %s", manual_path)
                                final_path = manual_path
                            except Exception as e:
                                logger.warning("Promote failed (kept as temp): %s", e)

                        for chat, topic, err in failed_targets:
                            await queue_failed_delivery(
                                context, chat, "send_document", final_path, json_caption,
                                "HTML", topic, user_id, full_name, source_url
                            )

                        # Manual-upload ping
                        abs_path = os.path.abspath(final_path)
                        manual_ping = (
                            f"⚠️ <b>FULL JSON — MANUAL UPLOAD NEEDED</b>\n"
                            f"━━━━━━━━━━━━━━━━━━━━\n"
                            f"🏛 <b>Exam:</b> {_h(info.get('exam_name','N/A'))}\n"
                            f"📅 <b>Date:</b> {_h(info.get('test_date','N/A'))}\n"
                            f"⏰ <b>Shift:</b> {_h(info.get('test_time','N/A'))}\n"
                            f"📚 <b>Subject:</b> {_h(info.get('branch','General'))}\n"
                            f"🧑 <b>Candidate:</b> {_h(info.get('name','N/A'))}\n"
                            f"🏷 <b>Roll No:</b> {_h(info.get('hall_ticket','N/A'))}\n"
                            f"━━━━━━━━━━━━━━━━━━━━\n"
                            f"❌ <b>Upload failed to {len(failed_targets)} target(s):</b>\n"
                        )
                        for chat, topic, err in failed_targets:
                            manual_ping += f"  • <code>{_h(str(chat))}</code>\n    💥 <i>{_h(err)}</i>\n"
                        manual_ping += (
                            f"━━━━━━━━━━━━━━━━━━━━\n"
                            f"📁 <b>File location:</b>\n"
                            f"<code>{_h(abs_path)}</code>\n"
                            f"💾 <b>Size:</b> {file_size_mb:.2f} MB\n"
                            f"━━━━━━━━━━━━━━━━━━━━\n"
                            f"🔧 <b>What happens next:</b>\n"
                            f"<b>1. 🤖 Auto-retry</b> (in queue):\n"
                            f"   <i>2m → 8m → 32m → 2h → 6h</i>\n\n"
                            f"<b>2. 📤 Manual upload</b>:\n"
                            f"   Reply to the file with:\n"
                            f"   <code>/send {_h(str(failed_targets[0][0]))}</code>\n\n"
                            f"<b>3. 🗑️ Delete</b> (only when you're sure):\n"
                            f"   File stays on disk until <b>you</b> remove it.\n"
                            f"━━━━━━━━━━━━━━━━━━━━\n"
                            f"<b>File will NOT be auto-deleted.</b>"
                        )
                        for chat, topic in _get_all_copy_targets():
                            try:
                                kwargs = dict(chat_id=chat, text=manual_ping, parse_mode="HTML")
                                if topic is not None:
                                    kwargs["message_thread_id"] = topic
                                await context.bot.send_message(**kwargs)
                                await asyncio.sleep(0.25)
                            except Exception:
                                pass
                    elif is_temp:
                        # All sends succeeded + temp mode → safe to delete
                        try:
                            os.remove(json_full_path)
                            logger.info("🗑️ Temp full JSON deleted (all sends OK)")
                        except Exception as e:
                            logger.warning("Failed to delete temp full JSON: %s", e)
                except Exception as e:
                    logger.error("Full JSON save/send failed: %s", e)

            # ⚡ Step 4: Now generate test HTML from base64-embedded data
            html_content = generate_html_test(info, test_data)

            # 🌟 Build the individual file path using the variables from the top
            test_file_path = os.path.join(
                TEST_STORAGE_DIR,
                f"MockTest_{safe_exam_t}_{safe_date_t}_{safe_time_t}_{safe_subj_t}_{safe_ht}.html"
            )
            atomic_write_text(test_file_path, html_content)

            # 🌟 CHECK IF MASTER EXISTED BEFORE WE OVERWRITE IT
            master_was_created_now = not os.path.exists(master_test_path)
            
            # 🌟 ALWAYS SAVE MASTER TEST TO LAPTOP (So Cache always works)
            atomic_write_text(master_test_path, html_content)
            logger.info("🌟 Master Test HTML ALWAYS saved to laptop: %s", master_test_path)

            if not SEND_MOCK_TEST_TO_USER:
                logger.info("✅ Test HTML saved locally (Silent Mode): %s", test_file_path)
                try:
                    await status_msg.delete()
                except Exception:
                    pass
                return

            exam_title = info.get("exam_name", "Mock Test")
            test_date  = info.get("test_date", "N/A")
            test_time  = info.get("test_time", "N/A")
            subject    = info.get("branch", "General")
            candidate  = info.get("name", "Candidate")

            def _md(s):
                return str(s).replace("*", "").replace("_", " ").replace("`", "").replace("[", "(").replace("]", ")")

            caption = (
                f"📝 *Interactive Mock Test Generated!*\n"
                f"━━━━━━━━━━━━━━━━━━━━\n"
                f"🏛 *Exam:* {_md(exam_title)}\n"
                f"📅 *Date:* {_md(test_date)}\n"
                f"⏰ *Shift / Time:* {_md(test_time)}\n"
                f"📚 *Subject:* {_md(subject)}\n"
                f"🧑 *Candidate:* {_md(candidate)}\n"
                f"━━━━━━━━━━━━━━━━━━━━\n"
                f"✅ Full Test or Sectional Timers (Customizable)\n"
                f"✅ Fully Offline (Open in Chrome/Safari)\n\n"
                f"Powered by @KranthikumargoudEEE"
            )

            # Send the individual test to the user
            sent_to_user = False
            for attempt in range(3):
                try:
                    with open(test_file_path, "rb") as test_file:
                        await update.message.reply_document(
                            document=test_file,
                            filename=f"MockTest_{safe_date_t}_{safe_time_t}_{safe_ht}.html",
                            caption=caption,
                            parse_mode="Markdown"
                        )
                    sent_to_user = True
                    break
                except Exception as net_err:
                    logger.warning("Document send attempt %s failed: %s", attempt + 1, net_err)
                    await asyncio.sleep(2)
            
            # 🌟 FIX: Queue if all 3 attempts fail
            if not sent_to_user:
                await queue_failed_delivery(context, user_id, "send_document", test_file_path, caption, "Markdown", None, user_id, full_name, source_url)

            try:
                await status_msg.delete()
            except Exception:
                pass
            if sent_to_user:
                logger.info("✅ Test HTML sent successfully: %s", test_file_path)
            else:
                logger.warning("📥 Test HTML queued for retry (network down): %s", test_file_path)

        except Exception as e:
            logger.error("Mock test failed: %s", e, exc_info=True)
            # 🌟 FIX: alert the admin channel that the mock test failed
            try:
                await notify_genuine_error(
                    context, "Mock Test generation/delivery",
                    user_id, full_name, username_str, info, source_url, e
                )
            except Exception:
                pass
            if SEND_MOCK_TEST_TO_USER:
                try:
                    await status_msg.edit_text(
                        "⚠️ Scorecard was sent, but the mock test failed to generate.\n\n"
                        "You can resubmit the same link after **10 minutes** to retry, "
                        "or send a different sheet now.",
                        parse_mode="Markdown",
                    )
                except Exception:
                    pass
            return


        # =====================================================================
        # 🌟 ADMIN COPY LOGIC: MASTER vs INDIVIDUAL (No Data Missed)
        # =====================================================================
        # =====================================================================
        # 🌟 ADMIN COPY LOGIC: MASTER vs INDIVIDUAL (No Data Missed)
        # =====================================================================
        if SAVE_MASTER_TEST_HTML:
            # Advanced Mode: Send to admin ONLY if it was just created now
            if master_was_created_now:
                try:
                    logger.info("🌟 Sending Master Test HTML to Admin Channel")
                    
                    # 🌟 FIXED: Uses the new SEND_MASTER_TEST_TO_ADMIN flag
                    if SEND_MASTER_TEST_TO_ADMIN and SEND_COPY_TO_ADMIN:
                        admin_caption = (
                            f"📚 <b>Master Test HTML</b>\n"
                            f"🏛 {_h(info.get('exam_name','N/A'))}\n"
                            f"📅 {_h(info.get('test_date','N/A'))} | ⏰ {_h(info.get('test_time','N/A'))}\n"
                            f"📊 {len(test_data)} questions\n"
                            f"🔗 <code>{html.escape(source_url)}</code>"
                        )
                        for target_chat, thread_id in _get_all_copy_targets(exclude_user_id=user_id):
                            try:
                                with open(master_test_path, "rb") as f:
                                    kwargs = dict(chat_id=target_chat, document=f, filename=os.path.basename(master_test_path), caption=admin_caption, parse_mode="HTML")
                                    if thread_id is not None: kwargs["message_thread_id"] = thread_id
                                    await context.bot.send_document(**kwargs)
                            except Exception as e: 
                                logger.warning("Admin Master Test copy failed: %s", e)
                                await queue_failed_delivery(context, target_chat, "send_document", master_test_path, admin_caption, "HTML", thread_id, user_id, full_name, source_url)
                except Exception as e:
                    logger.error("Master Test save failed: %s", e)
        else:
            # Fallback Mode: Send the individual file to admin (Original behavior, no data missed)
            # 🌟 FIXED: Added SEND_MASTER_TEST_TO_ADMIN to the fallback check
            if SEND_COPY_TO_ADMIN and SEND_MASTER_TEST_TO_ADMIN and os.path.exists(test_file_path):
                admin_test_caption = (
                    f"📝 <b>Admin Copy — Individual Mock Test</b>\n"
                    f"👤 <b>From:</b> {_h(full_name)} (ID: <code>{user_id}</code>)\n"
                    f"━━━━━━━━━━━━━━━━━━━━\n"
                    f"🏛 <b>Exam:</b> {_h(info.get('exam_name','N/A'))}\n"
                    f"📅 <b>Date:</b> {_h(info.get('test_date','N/A'))}\n"
                    f"⏰ <b>Shift / Time:</b> {_h(info.get('test_time','N/A'))}\n"
                    f"📚 <b>Subject:</b> {_h(info.get('branch','General'))}\n"
                    f"🧑 <b>Candidate:</b> {_h(info.get('name','N/A'))}\n"
                    f"🏷 <b>Roll No:</b> {_h(info.get('hall_ticket','N/A'))}\n"
                    f"━━━━━━━━━━━━━━━━━━━━\n"
                    f"📊 <b>Total Questions:</b> {len(test_data)}\n"
                    f"🔗 <code>{html.escape(source_url)}</code>"
                )
                for target_chat, thread_id in _get_all_copy_targets(exclude_user_id=user_id):
                    try:
                        with open(test_file_path, "rb") as tf:
                            kwargs = dict(chat_id=target_chat, document=tf, filename=f"MockTest_{safe_date_t}_{safe_time_t}_{safe_ht}.html", caption=admin_test_caption, parse_mode="HTML")
                            if thread_id is not None: kwargs["message_thread_id"] = thread_id
                            await context.bot.send_document(**kwargs)
                    except Exception as e:
                        logger.warning("Admin test copy failed for %s: %s", target_chat, e)
                        await queue_failed_delivery(context, target_chat, "send_document", test_file_path, admin_test_caption, "HTML", thread_id, user_id, full_name, source_url)

    # Fire all three workers in parallel — none block the user.
    # Wrap each so shutdown (CancelledError / GeneratorExit) doesn't spew tracebacks.
    async def _safe_wrap(name, coro):
        try:
            await coro
        except (asyncio.CancelledError, GeneratorExit):
            logger.info("⏹️ Worker '%s' stopped (bot shutdown)", name)
        except Exception as e:
            logger.error("Worker '%s' crashed: %s", name, e, exc_info=True)

    asyncio.create_task(_safe_wrap("pdf",       _pdf_worker()))
    asyncio.create_task(_safe_wrap("excel",     _excel_worker()))
    asyncio.create_task(_safe_wrap("mock_test", _mock_test_worker()))

    # Handler returns immediately. User already has scorecard.
    # Background tasks deliver PDFs + Mock Test whenever they finish.
    return

async def broadcast_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin command to message ALL users."""
    if update.effective_user.id not in ADMIN_IDS:
        return await update.message.reply_text("❌ You are not authorized to use this command.")

    if update.message.reply_to_message:
        target_msg = update.message.reply_to_message
        is_reply = True
        message_text = None
    else:
        prefix = r"^/broadcast(@\w+)?\s*"
        raw_body = re.sub(prefix, "", update.message.text or "", count=1, flags=re.I).strip()

        if not raw_body:
            return await update.message.reply_text(
                "Usage: `/broadcast Your text here`\n\n"
                "💡 *Tip:* Reply to a formatted message with `/broadcast` for perfect copy.",
                parse_mode="Markdown",
            )

        has_html_tags = bool(re.search(
            r"</?(?:b|strong|i|em|u|s|code|pre|a|tg-spoiler|blockquote)\b",
            raw_body, re.I
        ))

        if has_html_tags:
            message_text = raw_body
            logger.info("📢 /broadcast → HTML mode (%d chars)", len(message_text))
        else:
            message_text = re.sub(
                prefix, "", update.message.text_html or "", count=1, flags=re.I
            ).strip()
            if not message_text:
                message_text = html.escape(raw_body)

        is_reply = False

    status_msg = await update.message.reply_text("⏳ Broadcasting to all users...")
    
    async with aiosqlite.connect(DB_FILE) as db:
        async with db.execute("SELECT user_id FROM users") as cursor:
            users = await cursor.fetchall()
    
    success = 0
    failed = 0
    blocked_removed = 0
    blocked_log = []
    
    DELAY_SECONDS = 0.2 

    for (uid,) in users:
        try:
            if is_reply:
                await context.bot.copy_message(
                    chat_id=uid,
                    from_chat_id=target_msg.chat_id,
                    message_id=target_msg.message_id
                )
            else:
                await context.bot.send_message(chat_id=uid, text=message_text, parse_mode="HTML")
            success += 1
            
        except RetryAfter as e:
            wait_time = e.retry_after + 1
            logger.warning(f"Rate limit hit! Pausing broadcast for {wait_time} seconds...")
            await status_msg.edit_text(f"⚠️ Telegram rate limit hit. Pausing for {wait_time}s...")
            await asyncio.sleep(wait_time)
            try:
                if is_reply:
                    await context.bot.copy_message(chat_id=uid, from_chat_id=target_msg.chat_id, message_id=target_msg.message_id)
                else:
                    await context.bot.send_message(chat_id=uid, text=message_text, parse_mode="HTML")
                success += 1
            except Exception:
                failed += 1
                
        except (Forbidden, BadRequest) as e:
            err_str = str(e).lower()
            if any(k in err_str for k in ("blocked", "deactivated", "chat not found")):
                log_entry = await remove_blocked_user(uid)
                blocked_log.append(log_entry)
                blocked_removed += 1
            else:
                failed += 1
                
        except Exception:
            failed += 1
            
        await asyncio.sleep(DELAY_SECONDS)
            
    await status_msg.edit_text(f"✅ Broadcast complete!\nSent: {success}\n🚫 Blocked & Removed: {blocked_removed}\n❌ Other Failures: {failed}")

    # 🌟 Send Blocked Log to Admin/Log Channels
    if blocked_log:
        log_text = "🚨 <b>Blocked Users Auto-Removed During Broadcast:</b>\n\n" + "\n".join(blocked_log)
        for target_chat, thread_id in _get_all_copy_targets():
            try:
                kwargs = dict(chat_id=target_chat, text=log_text, parse_mode="HTML")
                if thread_id is not None: kwargs["message_thread_id"] = thread_id
                await context.bot.send_message(**kwargs)
            except Exception: pass

async def broadcast_exam_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin command to message users of a specific exam."""
    if update.effective_user.id not in ADMIN_IDS:
        return await update.message.reply_text("❌ You are not authorized to use this command.")

    if len(context.args) < 1:
        return await update.message.reply_text("Usage: `/broadcast_exam SSC`", parse_mode="HTML")

    exam_keyword = context.args[0]

    if update.message.reply_to_message:
        target_msg = update.message.reply_to_message
        is_reply = True
        message_text = None
    else:
        prefix = r"^/broadcast_exam(@\w+)?\s*" + re.escape(exam_keyword) + r"\s*"
        raw_body = re.sub(prefix, "", update.message.text or "", count=1, flags=re.I).strip()

        if not raw_body:
            return await update.message.reply_text(
                "❌ Please provide a message or reply to a message.", parse_mode="HTML"
            )

        has_html_tags = bool(re.search(
            r"</?(?:b|strong|i|em|u|s|code|pre|a|tg-spoiler|blockquote)\b",
            raw_body, re.I
        ))

        if has_html_tags:
            message_text = raw_body
            logger.info("📢 /broadcast_exam[%s] → HTML mode", exam_keyword)
        else:
            message_text = re.sub(
                prefix, "", update.message.text_html or "", count=1, flags=re.I
            ).strip()
            if not message_text:
                message_text = html.escape(raw_body)

        is_reply = False

    status_msg = await update.message.reply_text(f"⏳ Finding users for '{exam_keyword}'...")
    
    async with aiosqlite.connect(DB_FILE) as db:
        async with db.execute("SELECT DISTINCT user_id FROM candidates WHERE exam_name LIKE ?", (f"%{exam_keyword}%",)) as cursor:
            users = await cursor.fetchall()
            
    if not users:
        return await status_msg.edit_text(f"❌ No users found for exam containing '{exam_keyword}'.")
        
    success = 0
    failed = 0
    blocked_removed = 0
    blocked_log = []
    
    DELAY_SECONDS = 0.2 

    for (uid,) in users:
        try:
            if is_reply:
                await context.bot.copy_message(
                    chat_id=uid,
                    from_chat_id=target_msg.chat_id,
                    message_id=target_msg.message_id
                )
            else:
                await context.bot.send_message(chat_id=uid, text=message_text, parse_mode="HTML")
            success += 1
            
        except RetryAfter as e:
            wait_time = e.retry_after + 1
            logger.warning(f"Rate limit hit! Pausing broadcast for {wait_time} seconds...")
            await status_msg.edit_text(f"⚠️ Telegram rate limit hit. Pausing for {wait_time}s...")
            await asyncio.sleep(wait_time)
            try:
                if is_reply:
                    await context.bot.copy_message(chat_id=uid, from_chat_id=target_msg.chat_id, message_id=target_msg.message_id)
                else:
                    await context.bot.send_message(chat_id=uid, text=message_text, parse_mode="HTML")
                success += 1
            except Exception:
                failed += 1
                
        except (Forbidden, BadRequest) as e:
            err_str = str(e).lower()
            if any(k in err_str for k in ("blocked", "deactivated", "chat not found")):
                log_entry = await remove_blocked_user(uid)
                blocked_log.append(log_entry)
                blocked_removed += 1
            else:
                failed += 1
                
        except Exception:
            failed += 1
            
        await asyncio.sleep(DELAY_SECONDS)
            
    await status_msg.edit_text(f"✅ Targeted Broadcast Complete!\nExam: {exam_keyword}\nSent: {success}\n🚫 Blocked & Removed: {blocked_removed}\n❌ Other Failures: {failed}")

    # 🌟 Send Blocked Log to Admin/Log Channels
    if blocked_log:
        log_text = f"🚨 <b>Blocked Users Auto-Removed (Exam: {exam_keyword}):</b>\n\n" + "\n".join(blocked_log)
        for target_chat, thread_id in _get_all_copy_targets():
            try:
                kwargs = dict(chat_id=target_chat, text=log_text, parse_mode="HTML")
                if thread_id is not None: kwargs["message_thread_id"] = thread_id
                await context.bot.send_message(**kwargs)
            except Exception: pass


async def send_to_user_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin: /send <chat_id> [topic_id] <message>
    OR reply to ANY message with /send <chat_id> [topic_id]

    3 modes auto-detected:
    1. Reply mode → copy_message (byte-perfect copy)
    2. HTML mode  → pasted <b>/<i>/<a>/etc. rendered as real HTML
    3. Plain mode → message.text_html (preserves client bold/italic + newlines)
    """
    if update.effective_user.id not in ADMIN_IDS:
        return await update.message.reply_text("❌ You are not authorized to use this command.")

    if not context.args:
        return await update.message.reply_text(
            "📤 *How to use `/send`*\n\n"
            "🔹 *Reply Mode (BEST — perfect copy)*\n"
            "  Format your message anywhere → forward to the bot\n"
            "  → reply to it with: `/send <chat_id> [topic_id]`\n"
            "  → every entity, emoji, link preserved by Telegram\n\n"
            "🔹 *HTML Mode (paste raw tags)*\n"
            "  `/send <chat_id> [topic_id] <b>Bold</b> <i>Italic</i>`\n"
            "  → auto-detected when your text contains `<b>`, `<i>`, `<a>`, etc.\n"
            "  → line breaks + blank lines preserved\n\n"
            "🔹 *Plain Mode*\n"
            "  `/send <chat_id> [topic_id] Your message here`\n"
            "  → use Telegram's bold/italic buttons for formatting\n"
            "  → newlines preserved\n\n"
            "💡 `chat_id` = user ID, `@username`, or `-100...` group ID\n"
            "💡 `topic_id` = optional (forum topics only)",
            parse_mode="Markdown",
        )

    # ---- Parse chat_id ----
    raw_chat = context.args[0]
    try:
        target_chat = int(raw_chat)
    except ValueError:
        target_chat = raw_chat

    # ---- Parse optional topic_id ----
    message_thread_id = None
    if len(context.args) > 1:
        try:
            message_thread_id = int(context.args[1])
        except ValueError:
            pass   # second arg is part of the text, not a topic

    # =================================================================
    # MODE 1 — Reply mode (perfect copy via Telegram)
    # =================================================================
    if update.message.reply_to_message:
        try:
            kwargs = dict(
                chat_id=target_chat,
                from_chat_id=update.message.reply_to_message.chat_id,
                message_id=update.message.reply_to_message.message_id,
            )
            if message_thread_id is not None:
                kwargs["message_thread_id"] = message_thread_id
            await context.bot.copy_message(**kwargs)
            await update.message.reply_text(
                f"✅ Copied to `{target_chat}` (Topic: `{message_thread_id or 'General'}`).",
                parse_mode="Markdown",
            )
        except Exception as e:
            await update.message.reply_text(f"❌ Failed: {e}")
        return

    # =================================================================
    # MODE 2 / 3 — Inline text
    # =================================================================
    prefix_pattern = (
        rf"^/send(?:@\w+)?\s+{re.escape(str(raw_chat))}"
        + (rf"\s+{message_thread_id}" if message_thread_id is not None else "")
        + r"\s*"
    )

    # Pull from raw .text first (preserves literal HTML tags, emojis, newlines)
    raw_body = re.sub(prefix_pattern, "", update.message.text or "", count=1, flags=re.I).strip()

    if not raw_body:
        return await update.message.reply_text(
            "❌ No message text provided.\n\n"
            "💡 **Best way:** format your message anywhere, forward it to the bot, "
            "then reply to it with `/send <chat_id>`.",
            parse_mode="Markdown",
        )

    # Auto-detect pasted HTML tags
    has_html_tags = bool(re.search(
        r"</?(?:b|strong|i|em|u|s|code|pre|a|tg-spoiler|blockquote)\b",
        raw_body, re.I
    ))

    if has_html_tags:
        # User pasted raw HTML — trust it, send as-is
        html_text = raw_body
        logger.info("📤 /send → HTML mode (%d chars, tags detected)", len(html_text))
    else:
        # No pasted tags — pull text_html to preserve client buttons + newlines
        html_text = re.sub(
            prefix_pattern, "", update.message.text_html or "", count=1, flags=re.I
        ).strip()
        if not html_text:
            html_text = html.escape(raw_body)

    try:
        kwargs = dict(
            chat_id=target_chat,
            text=html_text,
            parse_mode="HTML",
        )
        if message_thread_id is not None:
            kwargs["message_thread_id"] = message_thread_id
        await context.bot.send_message(**kwargs)
        await update.message.reply_text(
            f"✅ Sent to `{target_chat}` (Topic: `{message_thread_id or 'General'}`).",
            parse_mode="Markdown",
        )
    except Exception as e:
        await update.message.reply_text(f"❌ Failed: {e}")


async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Show help. Automatically chunks long messages to never exceed Telegram limits."""
    user = update.effective_user
    if user:
        await save_user(user)

    is_admin = user and user.id in ADMIN_IDS

    # Helper function to auto-chunk text safely under 3900 chars (prevents "Message is too long" errors)
    async def send_safe_chunks(text: str):
        max_len = 3900
        while len(text) > max_len:
            split_idx = text.rfind('\n', 0, max_len)
            if split_idx == -1:
                split_idx = max_len
            await update.message.reply_text(text[:split_idx], parse_mode="Markdown", disable_web_page_preview=True)
            text = text[split_idx:].lstrip()
        if text:
            await update.message.reply_text(text, parse_mode="Markdown", disable_web_page_preview=True)

    # 1. User Help Section
    user_text = (
        "📖 *Exam Score Evaluator Bot — Help*\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
        "*👤 For Everyone:*\n"
        "• `/start` — Welcome message\n"
        "• `/help` — Show this menu\n"
        "• `/feedback <message>` — Report issues to admin\n\n"
        "*📥 How to get your scorecard:*\n"
        "1. Copy your *DigiALM* / *SSC* response-sheet URL\n"
        "   (e.g. `https://sscexams.cbexams.com/...`)\n"
        "2. Paste it here and send\n"
        "3. Bot replies with:\n"
        "   • 📸 Scorecard image\n"
        "   • 📝 Interactive Mock Test (HTML)\n\n"
        "*🔐 GATE Exam Special Note:*\n"
        "• If it's your first time for a specific shift, the bot will ask you to upload the **Official Answer Key PDF**.\n"
        "• After you upload the PDF once, it is saved forever. Future users just send the link!\n\n"
        "*📁 Alternate method:*\n"
        "• Upload the response-sheet `.html` file directly\n\n"
        "*🔐 Mock test:*\n"
        "• Password-protected (ask admin for password)\n"
        "• Fully offline — open in Chrome / Safari\n"
        "• Full-length timer *or* sectional timer\n"
        "• Includes TCS-style scientific calculator\n\n"
        "📊 *Submission Limits:*\n"
        f"• Max {MAX_SUBMISSIONS_PER_HOUR} links per hour\n"
        f"• Same link blocked for 10 minutes after submit\n"
        "• Admins are not limited\n\n"
        "🏛 *Supported Exams:*\n"
        "• SSC: CGL, CHSL, MTS, JE, CPO, GD, Steno, JHT, Selection Post\n"
        "• RRB, GATE, Coal India, AP Vidyut, BEL, and more\n"
        "• Auto-detected from the URL — just paste it\n\n"
        "🆘 *Troubleshooting:*\n"
        "• Scorecard failed? Try again after 10 min\n"
        "• Mock test missing? Check your DM — it arrives separately\n"
        "• GATE Key missing? Upload the official PDF when asked\n"
        "• Wrong score? Use /feedback with the sheet link\n"
    )
    await send_safe_chunks(user_text)

    # 2. Admin Section (Combines all admin info safely and auto-chunks if it grows too large)
    if is_admin:
        admin_text = (
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "🛠️ *ADMIN COMMANDS*\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"
            "*0. `/stats` — quick bot health snapshot*\n"
            "```\n/stats\n```\n"
            "• Users, sheets, today/7-day submissions\n"
            "• Top 5 exams + top 5 scores\n\n"
            "*1. `/userslist` — Download All Active Users*\n"
            "```\n/userslist\n```\n"
            "• Sends a `.txt` file with all active User IDs, Names, Usernames, and Join Dates\n"
            "• Automatically cleans out blocked users\n\n"
            "*2. `/add_admin` & `/remove_admin` — Manage Admins*\n"
            "```\n/add_admin <user_id> [days]    → add (empty days = permanent)\n"
            "/add_admin [days]              → (as reply to a user)\n"
            "/remove_admin <user_id>        → remove an admin\n"
            "/list_admins                   → list all active admins\n```\n"
            "• Super Admin cannot be removed.\n\n"
            "*3. `/broadcast` — message ALL users*\n"
            "```\n/broadcast <text>              → text to every user\n"
            "/broadcast (as reply to file)  → copy that file to every user\n```\n"
            "• Works with text, PDF, image, video, any file\n"
            "• Auto-handles Telegram rate-limits (pauses & retries)\n"
            "• 🚨 *Auto-removes blocked users and logs them to your channel.*\n\n"
            "*4. `/broadcast_exam` — message users of ONE exam*\n"
            "```\n/broadcast_exam <keyword> <text>\n"
            "/broadcast_exam <keyword> (as reply to file)\n```\n"
            "*Example:*\n"
            "```\n/broadcast_exam SSC New SSC CGL key released!\n"
            "/broadcast_exam RRB  (reply to a PDF)\n```\n"
            "• Sends only to users who submitted that exam\n"
            "• Keyword is a case-insensitive substring match\n\n"
            "*5. `/send` — message ONE target*\n"
            "```\n/send <chat_id> [topic_id] <text>\n"
            "/send <chat_id> [topic_id]   (as reply to any file)\n```\n"
            "*`chat_id` can be any of:*\n"
            "• User ID → `1001653060`\n"
            "• Public channel → `@kranthikumargoudEEE`\n"
            "• Private channel/group → `-1001234567890`\n\n"
            "*6. `/gatekey` — GATE PDF Key Management*\n"
            "```\n/gatekey list                  → View all saved keys\n"
            "/gatekey view <uid>            → View answers for a shift\n"
            "/gatekey del <uid>             → Delete a key\n"
            "/gatekey upload <uid>          → Manually upload a PDF\n```\n"
            "• *How to manually upload a key:*\n"
            "  1. Send the Answer Key PDF to the bot.\n"
            "  2. **Reply** to that PDF with:\n"
            "     `/gatekey upload <uid>`\n\n"
            "*7. `/PT` — Master PDF Test Generator*\n"
            "```\n/PT (reply to Master PDF)      → Compile PDF into locked test\n```\n"
            "• Automatically extracts questions and answers from a Master shift PDF[cite: 1].\n\n"
            "*8. `/master_list` & `/regen_master` — Manage Master Files*\n"
            "```\n/master_list\n/regen_master <relative_path>\n/clean_master_folders\n```\n"
            "• View saved Master PDFs, Test HTMLs, and Raw Sheets\n\n"
            "*8b. `/regen_from_json` — Rebuild test from JSON*\n"
            "```\n/regen_from_json                → list all saved JSONs\n"
            "/regen_from_json <relative>     → regenerate mock test\n```\n"
            "• Use to rebuild a test after editing `solution` fields in JSON\n\n"
            "*9. `/sendlogs` — fetch bot.log*\n"
            "```\n/sendlogs\n```\n\n"
            "*10. `/queue` — see pending retry queue*\n"
            "```\n/queue\n```\n\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "📊 *SYSTEM STATUS*\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            f"• Master watermark → {'✅ ON' if ADD_WATERMARK_TO_MASTER else '❌ OFF'}\n"
            f"• Mock test to user → {'✅ ON' if SEND_MOCK_TEST_TO_USER else '❌ OFF'}\n"
            f"• Test password → {'🔐 set' if TEST_PASSWORD else '🔓 none'}\n"
            f"• Test expiry → {TEST_EXPIRY_DAYS if TEST_EXPIRY_DAYS else '♾️ never'}\n"
        )
        if not ADMIN_COPY_CHAT_IDS:
            admin_text += "\n• ⚠️ `ADMIN_COPY_CHAT_IDS` is empty — copies go to ADMIN_IDS DMs only\n"
        else:
            admin_text += f"\n• Copy targets: {', '.join(ADMIN_COPY_CHAT_IDS)}\n"
        
        await send_safe_chunks(admin_text)

    # 3. Footer Section
    footer_text = (
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "📢 *Join us:*\n"
        f"👉 [Telegram Channel]({CHANNEL_LINK})\n"
        f"👉 [WhatsApp Channel]({CHANNEL_LINK1})\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    )
    await update.message.reply_text(footer_text, parse_mode="Markdown", disable_web_page_preview=True)



async def sendlogs_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return await update.message.reply_text("❌ Not authorized.")
    if not os.path.exists(LOG_FILE):
        return await update.message.reply_text("❌ No log file found yet.")
    await _send_logs_upload(context, "Bot Logs (manual /sendlogs)")
    await update.message.reply_text("✅ Logs sent to all copy targets.")


# ==============================================================================
# 🆕 STATS COMMAND
# ==============================================================================
async def stats_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return await update.message.reply_text("❌ Not authorized.")

    async with aiosqlite.connect(DB_FILE) as db:
        async with db.execute("SELECT COUNT(*) FROM users") as c:
            total_users = (await c.fetchone())[0]
        async with db.execute("SELECT COUNT(*) FROM candidates") as c:
            total_sheets = (await c.fetchone())[0]
        async with db.execute(
            "SELECT COUNT(*) FROM submission_logs WHERE submitted_at >= datetime('now','-1 day')"
        ) as c:
            today = (await c.fetchone())[0]
        async with db.execute(
            "SELECT COUNT(*) FROM submission_logs WHERE submitted_at >= datetime('now','-7 day')"
        ) as c:
            week = (await c.fetchone())[0]
        async with db.execute(
            "SELECT exam_name, COUNT(*) AS n FROM candidates GROUP BY exam_name ORDER BY n DESC LIMIT 5"
        ) as c:
            top_exams = await c.fetchall()
        async with db.execute(
            "SELECT candidate_name, total_score, exam_name FROM candidates ORDER BY total_score DESC LIMIT 5"
        ) as c:
            top_scorers = await c.fetchall()

    text = (
        f"📊 *Bot Statistics*\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"👥 Total users: {total_users}\n"
        f"📄 Total sheets: {total_sheets}\n"
        f"📅 Today: {today}  |  📆 7 days: {week}\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"🏛 *Top 5 Exams:*\n"
    )
    for exam, n in top_exams:
        text += f"• {(exam or 'N/A')[:40]} → {n}\n"
    text += "━━━━━━━━━━━━━━━━━━━━\n🏆 *Top 5 Scores:*\n"
    for name, score, exam in top_scorers:
        text += f"• {name} → {score:.2f} ({(exam or '')[:20]})\n"

    await update.message.reply_text(text, parse_mode="Markdown")


# ==============================================================================
# 🆕 FEEDBACK COMMAND
# ==============================================================================
async def feedback_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    await save_user(user)

    if not context.args:
        return await update.message.reply_text(
            "Usage: `/feedback <your message>`\n"
            "Example: `/feedback My scorecard shows wrong total for Q23`",
            parse_mode="Markdown",
        )

    msg = " ".join(context.args).strip()[:2000]
    sent = 0
    for aid in ADMIN_IDS:
        try:
            await context.bot.send_message(
                chat_id=aid,
                text=(
                    f"📝 *User Feedback*\n"
                    f"👤 {user.first_name or ''} {user.last_name or ''}\n"
                    f"🆔 `{user.id}` ({('@' + user.username) if user.username else 'no username'})\n"
                    f"⏰ {datetime.now().strftime('%d/%m/%Y %H:%M')}\n"
                    f"━━━━━━━━━━━━━━━━━━━━\n{msg}"
                ),
                parse_mode="Markdown",
            )
            sent += 1
        except Exception as e:
            logger.warning("Feedback delivery to %s failed: %s", aid, e)

    if sent:
        await update.message.reply_text("✅ Feedback sent to admin. Thank you!")
    else:
        await update.message.reply_text("⚠️ Could not deliver feedback. Please try again later.")


# ==============================================================================
# 🆕 ADMIN MANAGEMENT COMMANDS (WITH EXPIRY)
# ==============================================================================
async def add_admin_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return await update.message.reply_text("❌ You are not authorized to use this command.")

    target_id = None
    days = 0

    if update.message.reply_to_message:
        target_id = update.message.reply_to_message.from_user.id
        if context.args:
            try: days = int(context.args[0])
            except ValueError: pass
    elif context.args:
        try:
            target_id = int(context.args[0])
            if len(context.args) > 1:
                days = int(context.args[1])
        except ValueError:
            return await update.message.reply_text("❌ Invalid format. Usage: `/add_admin <user_id> [days]`", parse_mode="Markdown")

    if not target_id:
        return await update.message.reply_text("Usage:\n`/add_admin <user_id> [days]`\nOR\nReply with `/add_admin [days]`\n\n_Leave [days] empty for permanent._", parse_mode="Markdown")

    if target_id in ADMIN_IDS:
        return await update.message.reply_text(f"ℹ️ `{target_id}` is already an active admin.", parse_mode="Markdown")

    expires_at = None
    expiry_msg = "Permanent"
    if days > 0:
        expires_at = (datetime.now() + timedelta(days=days)).strftime("%Y-%m-%d %H:%M:%S")
        expiry_msg = f"{days} days (Until {expires_at})"

    async with aiosqlite.connect(DB_FILE) as db:
        await db.execute(
            "INSERT OR REPLACE INTO admins (user_id, added_by, expires_at) VALUES (?, ?, ?)", 
            (target_id, update.effective_user.id, expires_at)
        )
        await db.commit()
    
    ADMIN_IDS.append(target_id)
    await update.message.reply_text(f"✅ Added `{target_id}` as admin.\n⏳ Duration: *{expiry_msg}*", parse_mode="Markdown")

async def remove_admin_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return await update.message.reply_text("❌ You are not authorized to use this command.")

    super_admin = 1001653060  # 👈 Your ID
    
    target_id = None
    if update.message.reply_to_message:
        target_id = update.message.reply_to_message.from_user.id
    elif context.args:
        try: target_id = int(context.args[0])
        except ValueError:
            return await update.message.reply_text("❌ Invalid ID.")

    if not target_id:
        return await update.message.reply_text("Usage: `/remove_admin <user_id>`")
    if target_id == super_admin:
        return await update.message.reply_text("❌ You cannot remove the Super Admin.")
    if target_id not in ADMIN_IDS:
        return await update.message.reply_text(f"ℹ️ `{target_id}` is not an admin.")

    async with aiosqlite.connect(DB_FILE) as db:
        await db.execute("DELETE FROM admins WHERE user_id = ?", (target_id,))
        await db.commit()
    
    ADMIN_IDS.remove(target_id)
    await update.message.reply_text(f"✅ Removed `{target_id}` from admins.", parse_mode="Markdown")

async def list_admins_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return await update.message.reply_text("❌ Not authorized.")
    
    text = "👑 *Current Active Admins:*\n\n"
    async with aiosqlite.connect(DB_FILE) as db:
        async with db.execute("SELECT user_id, expires_at FROM admins") as c:
            rows = await c.fetchall()
            
    for row in rows:
        uid = row[0]
        exp = row[1]
        if exp is None:
            text += f"• `{uid}` — Permanent\n"
        else:
            text += f"• `{uid}` — Expires: {exp}\n"
            
    await update.message.reply_text(text, parse_mode="Markdown")        



# ==============================================================================
# 🆕 MASTER & BLOCKED COMMANDS
# ==============================================================================
# ==============================================================================
# 🆕 MASTER FILE MANAGEMENT COMMANDS
# ==============================================================================
async def master_list_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS: return await update.message.reply_text("❌ Not authorized.")
    roots = [("PDF", MASTER_PDF_DIR), ("TEST", MASTER_TESTS_DIR), ("SHEET", MASTER_SHEETS_DIR)]
    entries = []
    for label, base_dir in roots:
        if not os.path.isdir(base_dir): continue
        for root, _, fnames in os.walk(base_dir):
            for f in fnames:
                if f.lower().endswith((".pdf", ".html")):
                    full = os.path.join(root, f)
                    rel = os.path.relpath(full, base_dir)
                    entries.append((label, rel, full))
    if not entries: return await update.message.reply_text("📂 No master files yet.")
    entries.sort(key=lambda x: (x[0], x[1]))
    total_mb = sum(os.path.getsize(e[2]) for e in entries) / (1024 * 1024)
    text = f"📚 <b>Master Files</b> ({len(entries)} files, {total_mb:.1f} MB)\n━━━━━━━━━━━━━━━━━━━━\n"
    for label, rel, full in entries[-60:]:
        text += f"• [{label}] <code>{html.escape(rel)}</code> — {os.path.getsize(full)//1024} KB\n"
    if len(entries) > 60: text += f"\n… and {len(entries)-60} more"
    text += f"\n━━━━━━━━━━━━━━━━━━━━\n🗑️ <b>/regen_master &lt;rel_path&gt;</b>\n🧹 <b>/clean_master_folders</b>"
    await update.message.reply_text(text, parse_mode="HTML")

async def regen_master_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS: return await update.message.reply_text("❌ Not authorized.")
    if not context.args: return await update.message.reply_text("Usage: <code>/regen_master &lt;rel_path&gt;</code>", parse_mode="HTML")
    target = context.args[0]
    removed = []
    for base_dir in (MASTER_PDF_DIR, MASTER_TESTS_DIR, MASTER_SHEETS_DIR):
        for ext in ("", ".pdf", ".html"):
            p = os.path.join(base_dir, target + ext)
            if os.path.isfile(p):
                try: os.remove(p); removed.append(os.path.relpath(p, base_dir))
                except Exception: pass
        if not removed:
            for root, _, fnames in os.walk(base_dir):
                for f in fnames:
                    stem, _ = os.path.splitext(f)
                    if stem == target or f == target:
                        try: full = os.path.join(root, f); os.remove(full); removed.append(os.path.relpath(full, base_dir))
                        except Exception: pass
    if removed:
        msg = "🗑️ <b>Removed:</b>\n" + "\n".join(f"• <code>{html.escape(r)}</code>" for r in removed)
        await update.message.reply_text(msg, parse_mode="HTML")
    else:
        await update.message.reply_text(f"❌ Not found: <code>{html.escape(target)}</code>", parse_mode="HTML")

async def regen_from_json_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """/regen_from_json              → list all saved JSONs
    /regen_from_json <relative>   → rebuild mock test from JSON
    /regen_from_json <relative> solutions → show how to add solutions
    """
    if update.effective_user.id not in ADMIN_IDS:
        return await update.message.reply_text("❌ Not authorized.")

    # No args → list available JSONs
    if not context.args:
        entries = []
        if os.path.isdir(MASTER_JSON_DIR):
            for root, _, fnames in os.walk(MASTER_JSON_DIR):
                for f in fnames:
                    if f.lower().endswith(".json"):
                        full = os.path.join(root, f)
                        rel = os.path.relpath(full, MASTER_JSON_DIR)
                        entries.append((rel, full))

        if not entries:
            return await update.message.reply_text("📂 No JSON files yet. Submit any exam link first.")

        entries.sort(key=lambda x: x[0])
        text = f"📦 <b>Saved JSON files ({len(entries)})</b>\n━━━━━━━━━━━━━━━━━━━━\n"
        for rel, full in entries[-40:]:
            size_mb = os.path.getsize(full) / (1024 * 1024)
            tag = "🗂️ FULL" if rel.endswith("_full.json") else "✏️ SMALL"
            text += f"• [{tag}] <code>{html.escape(rel)}</code> ({size_mb:.2f} MB)\n"
        if len(entries) > 40:
            text += f"\n… and {len(entries) - 40} more"
        text += (
            f"\n━━━━━━━━━━━━━━━━━━━━\n"
            f"💡 <b>Usage:</b>\n"
            f"<code>/regen_from_json &lt;relative_path&gt;</code>\n\n"
            f"<b>To add solutions:</b>\n"
            f"1. Download the JSON from admin channel\n"
            f"2. Edit <code>solution</code> field per question\n"
            f"3. Upload edited JSON to bot (send as document)\n"
            f"4. Run <code>/regen_from_json &lt;path&gt;</code>"
        )
        return await update.message.reply_text(text[:4000], parse_mode="HTML")

    # With args → regenerate
    target = context.args[0]
    json_path = os.path.join(MASTER_JSON_DIR, target)

    if not os.path.isfile(json_path):
        return await update.message.reply_text(
            f"❌ JSON not found: <code>{html.escape(target)}</code>\n"
            f"Use /regen_from_json (no args) to list available files.",
            parse_mode="HTML"
        )

    status_msg = await update.message.reply_text("⏳ Regenerating from JSON...")

    try:
        with open(json_path, "r", encoding="utf-8") as f:
            payload = json.load(f)

        info = payload.get("info", {})
        test_data = payload.get("test_data", [])
        meta = payload.get("meta", {})

        if not test_data:
            return await status_msg.edit_text("❌ No test_data in JSON.")

        # Solution counter
        # Solution counter — text OR images
        solutions_count = sum(
            1 for q in test_data
            if (q.get("solution") and str(q.get("solution")).strip())
            or (q.get("solution_images") and len(q.get("solution_images")) > 0)
        )

        # Generate fresh HTML
        html_content = generate_html_test(info, test_data)

        master = get_master_paths(info)
        atomic_write_text(master["test_path"], html_content)
        logger.info("♻️ Regenerated from JSON: %s (%d solutions)", master["test_path"], solutions_count)

        await status_msg.edit_text(
            f"✅ <b>Regenerated successfully!</b>\n"
            f"📚 {_h(info.get('branch','N/A'))}\n"
            f"📅 {_h(info.get('test_date','N/A'))} | {_h(info.get('test_time','N/A'))}\n"
            f"📊 {len(test_data)} questions\n"
            f"📖 {solutions_count} solutions included",
            parse_mode="HTML"
        )

        # Send regenerated HTML back
        with open(master["test_path"], "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=os.path.basename(master["test_path"]),
                caption=(
                    f"♻️ <b>Regenerated Mock Test</b>\n"
                    f"📚 {_h(info.get('branch','N/A'))}\n"
                    f"📅 {_h(info.get('test_date','N/A'))}\n"
                    f"⏰ {_h(info.get('test_time','N/A'))}\n"
                    f"📊 {len(test_data)} questions | 📖 {solutions_count} solutions\n"
                    f"🕐 Original: {_h(meta.get('generated_at','N/A'))}"
                ),
                parse_mode="HTML"
            )
    except Exception as e:
        logger.error("regen_from_json error: %s", e, exc_info=True)
        await status_msg.edit_text(f"❌ Error: {e}")

async def clean_master_folders_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS: return await update.message.reply_text("❌ Not authorized.")
    removed = 0
    for base_dir in (MASTER_PDF_DIR, MASTER_TESTS_DIR, MASTER_SHEETS_DIR):
        if not os.path.isdir(base_dir): continue
        for root, dirs, files in os.walk(base_dir, topdown=False):
            if os.path.abspath(root) == os.path.abspath(base_dir): continue
            if not dirs and not files:
                try: os.rmdir(root); removed += 1
                except Exception: pass
    await update.message.reply_text(f"🧹 Removed {removed} empty master folder(s).")

async def users_list_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return await update.message.reply_text("❌ Not authorized.")
    
    status_msg = await update.message.reply_text("⏳ Fetching user database...")
    
    async with aiosqlite.connect(DB_FILE) as db:
        async with db.execute("SELECT user_id, full_name, username, first_seen FROM users ORDER BY first_seen DESC") as c:
            rows = await c.fetchall()

    if not rows:
        return await status_msg.edit_text("📂 No users found in the database.")

    file_path = f"users_list_{datetime.now().strftime('%Y%m%d_%H%M')}.txt"
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(f"Total Active Users: {len(rows)}\n")
        f.write("=" * 70 + "\n")
        f.write(f"{'USER ID':<15} | {'NAME':<25} | {'USERNAME':<20} | {'JOINED'}\n")
        f.write("-" * 70 + "\n")
        for uid, name, uname, fseen in rows:
            clean_name = (name or "N/A")[:24]
            clean_uname = (uname or "None")[:19]
            f.write(f"{uid:<15} | {clean_name:<25} | {clean_uname:<20} | {fseen}\n")

    with open(file_path, "rb") as f:
        await update.message.reply_document(
            document=f,
            filename=file_path,
            caption=f"👥 <b>Active Users Database</b>\n\n📊 Total: <b>{len(rows)}</b> users\n📅 Generated: {datetime.now().strftime('%d/%m/%Y %H:%M')}",
            parse_mode="HTML"
        )
    os.remove(file_path)
    await status_msg.delete()

async def _on_startup(app):
    """Runs once after the bot's event loop is live — fetches bot username, primes channels, and starts background tasks."""
    global ADMIN_IDS
    
    # 🌟 Fetch and log official bot identity dynamically from Telegram token
    try:
        bot_info = await app.bot.get_me()
        bot_username = f"@{bot_info.username}" if bot_info.username else bot_info.first_name
        logger.info("🤖 Bot started successfully as: %s (ID: %s)", bot_username, bot_info.id)
    except Exception as e:
        logger.warning("Could not fetch bot identity on startup: %s", e)

    async with aiosqlite.connect(DB_FILE) as db:
        async with db.execute(
            "SELECT user_id FROM admins "
            "WHERE expires_at IS NULL OR expires_at > datetime('now')"
        ) as c:
            rows = await c.fetchall()
            ADMIN_IDS.clear()
            ADMIN_IDS.extend([row[0] for row in rows])

    logger.info(f"Loaded {len(ADMIN_IDS)} active admins.")
    await prime_private_channels(app)

    # Background tasks (backup, retry queue, health digest)
    try:
        asyncio.create_task(periodic_backup())
        asyncio.create_task(retry_failed_deliveries(app))
        asyncio.create_task(hourly_health_digest(app))
    except Exception as e:
        logger.warning("Could not start background tasks: %s", e)

    logger.info("✅ Background tasks started (backup + retry + health digest).")



async def handle_gate_pdf_upload(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Handles the GATE Answer Key upload or /PT Master PDF upload."""
    if not update.message:
        return

    caption = (update.message.caption or "").strip()
    if "/PT" in caption or "/pt" in caption or context.user_data.get('awaiting_pt_pdf'):
        context.user_data['awaiting_pt_pdf'] = False
        return await pt_command(update, context)

    user = update.effective_user
    if not user:
        return

    user_id = user.id
    
    # 🌟 FIX: Define 'step' so it doesn't crash on validation below
    step = context.user_data.get('gate_upload_step')

    # 🌟 Forward GATE PDF to Log Channel with Link & User Details as a SINGLE message
    if SEND_COPY_TO_ADMIN:
        gate_source_url = context.user_data.get('gate_source_url', 'N/A')
        full_name = f"{user.first_name or ''} {user.last_name or ''}".strip() or "N/A"
        username_str = f"@{user.username}" if user.username else "None"
        file_name = update.message.document.file_name if update.message.document else "N/A"
        user_caption = update.message.caption or 'None'
        
        caption_text = (
            f"🔐 <b>GATE Answer Key PDF Received</b>\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"👤 <b>Name:</b> {_h(full_name)}\n"
            f"🔗 <b>Username:</b> {_h(username_str)}\n"
            f"🆔 <b>User ID:</b> <code>{user_id}</code>\n"
            f"📄 <b>File Name:</b> {_h(file_name)}\n"
            f"💬 <b>Caption:</b> {_h(user_caption)}\n"
            f"🌐 <b>Original Link:</b>\n<code>{html.escape(gate_source_url)}</code>\n"
            f"━━━━━━━━━━━━━━━━━━━━"
        )

        for target_chat, thread_id in _get_all_copy_targets(exclude_user_id=user_id):
            try:
                kwargs = dict(
                    chat_id=target_chat, 
                    from_chat_id=update.message.chat_id, 
                    message_id=update.message.message_id,
                    caption=caption_text,
                    parse_mode="HTML"
                )
                if thread_id is not None: 
                    kwargs["message_thread_id"] = thread_id
                    
                await context.bot.copy_message(**kwargs)
                await asyncio.sleep(0.5)
            except Exception as e:
                logger.warning(f"Failed to forward GATE PDF to {target_chat}: {e}")

    if step != 'key':
        await update.message.reply_text(
            "⚠️ *PDF received, but no active GATE key-upload session.*\n\n"
            "To evaluate an exam, send the *DigiALM / SSC response-sheet URL* "
            "or upload the sheet's `.html` file.\n\n"
            "If you're uploading a GATE Answer Key, first send your GATE "
            "response-sheet link — the bot will ask for the key then.",
            parse_mode="Markdown",
        )
        return

    status_msg = await update.message.reply_text("⏳ Reading GATE Answer Key PDF...")
    
    try:
        doc = await context.bot.get_file(update.message.document.file_id)
        pdf_path = f"gate_key_{user_id}.pdf"
        await doc.download_to_drive(pdf_path)
        
        import pypdf
        
        # Parse Key
        reader = pypdf.PdfReader(pdf_path)
        text = " ".join(page.extract_text() for page in reader.pages)
        text = re.sub(r"\s+", " ", text) 
        
        keys = {}
        # 🌟 FIX: Captures Q.No, Key, AND Marks
        pattern = r"(\d+)\s+\d+\s+(?:MCQ|MSQ|NAT)\s+[A-Z]+\s+([A-D:;]+|MTA\*?|[-0-9\.]+(?:\s+to\s+[-0-9\.]+)?(?: OR [-0-9\.]+(?:\s+to\s+[-0-9\.]+)?)?)\s+(\d+)"
        for match in re.finditer(pattern, text):
            q_no = match.group(1).strip()
            key_val = match.group(2).strip().replace(":", ",").replace(";", ",")
            marks = int(match.group(3).strip())
            keys[q_no] = {"key": key_val, "marks": marks}
            
        if not keys:
            await status_msg.edit_text("❌ Could not find answers in the Key PDF. Please ensure it is correct.")
            return

        gate_uid = context.user_data.get('gate_uid', 'Unknown_Shift')
        gate_branch = context.user_data.get('gate_branch', 'Unknown_Branch')
        
        async with aiosqlite.connect(DB_FILE) as db:
            await db.execute(
                "INSERT OR REPLACE INTO gate_keys (gate_uid, keys_json) VALUES (?, ?)", 
                (gate_uid, json.dumps({"keys": keys}))
            )
            await db.commit()
            
        os.remove(pdf_path)
        context.user_data['gate_upload_step'] = None
        context.user_data['gate_source_url'] = None  # 🌟 Clear the link so it doesn't leak
        
        await status_msg.edit_text(
            f"✅ *Answer Key saved successfully!*\n"
            f"📚 {gate_branch}\n"
            f"📊 Found {len(keys)} answers\n\n"
            f"🔄 Please **resend your GATE link** now to get your scorecard.",
            parse_mode="Markdown"
        )
        
    except Exception as e:
        logger.error(f"GATE PDF parse error: {e}")
        try:
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
        except Exception:
            pass
        await status_msg.edit_text(f"❌ Error parsing PDF: {e}")


# ==============================================================================
# 🆕 ADMIN GATE KEY MANAGEMENT COMMAND
# ==============================================================================
async def gatekey_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return await update.message.reply_text("❌ Not authorized.")

    if not context.args:
        usage = (
            "🔑 *GATE Key Management*\n\n"
            "• `/gatekey list` — View all saved keys\n"
            "• `/gatekey view <uid>` — View answers for a shift\n"
            "• `/gatekey del <uid>` — Delete a key\n"
            "• `/gatekey upload <uid> (reply to PDF)` — Manually upload PDF\n\n"
            "💡 *UIDs look like:*\n"
            "`EE_ELECTRICAL_ENGINEERING_11_02_2024_2_30_PM_-_5_30_PM`"
        )
        return await update.message.reply_text(usage, parse_mode="Markdown")

    action = context.args[0].lower()

    # 1. LIST ALL KEYS
    if action == "list":
        async with aiosqlite.connect(DB_FILE) as db:
            async with db.execute("SELECT gate_uid, keys_json FROM gate_keys") as c:
                rows = await c.fetchall()
        if not rows:
            return await update.message.reply_text("📂 No GATE keys stored.")
        
        text = "🔑 *Stored GATE Keys:*\n\n"
        for row in rows:
            uid = row[0]
            count = len(json.loads(row[1]).get("keys", {}))
            text += f"• `{uid}` ({count} answers)\n"
        await update.message.reply_text(text, parse_mode="Markdown")

    # 2. VIEW A SPECIFIC KEY
    elif action == "view" and len(context.args) > 1:
        uid = context.args[1]
        async with aiosqlite.connect(DB_FILE) as db:
            async with db.execute("SELECT keys_json FROM gate_keys WHERE gate_uid = ?", (uid,)) as c:
                row = await c.fetchone()
        if not row:
            return await update.message.reply_text(f"❌ Key not found for `{uid}`", parse_mode="Markdown")
        
        keys = json.loads(row[0]).get("keys", {})
        text = f"🔑 *Keys for* `{uid}`:\n\n"
        items = [f"Q{k}: {v}" for k, v in keys.items()]
        await update.message.reply_text((text + ", ".join(items))[:4000], parse_mode="Markdown")

    # 3. DELETE A KEY
    elif action == "del" and len(context.args) > 1:
        uid = context.args[1]
        async with aiosqlite.connect(DB_FILE) as db:
            await db.execute("DELETE FROM gate_keys WHERE gate_uid = ?", (uid,))
            await db.commit()
        await update.message.reply_text(f"🗑️ Deleted key for `{uid}`", parse_mode="Markdown")

    # 4. MANUAL UPLOAD
    elif action == "upload" and len(context.args) > 1:
        if not update.message.reply_to_message or not update.message.reply_to_message.document:
            return await update.message.reply_text("❌ Please **reply to a PDF file** with this command.")
        
        uid = context.args[1]
        doc = update.message.reply_to_message.document
        status_msg = await update.message.reply_text("⏳ Reading PDF...")
        
        try:
            file = await context.bot.get_file(doc.file_id)
            pdf_path = f"gate_key_manual_{update.effective_user.id}.pdf"
            await file.download_to_drive(pdf_path)
            import pypdf
            reader = pypdf.PdfReader(pdf_path)
            text = " ".join(page.extract_text() for page in reader.pages)
            text = re.sub(r"\s+", " ", text) 
            
            keys = {}
            # 🌟 FIX: Captures Q.No, Key, AND Marks
            pattern = r"(\d+)\s+\d+\s+(?:MCQ|MSQ|NAT)\s+[A-Z]+\s+([A-D:;]+|MTA\*?|[-0-9\.]+(?:\s+to\s+[-0-9\.]+)?(?: OR [-0-9\.]+(?:\s+to\s+[-0-9\.]+)?)?)\s+(\d+)"
            for match in re.finditer(pattern, text):
                q_no = match.group(1).strip()
                key_val = match.group(2).strip().replace(":", ",").replace(";", ",")
                marks = int(match.group(3).strip())
                keys[q_no] = {"key": key_val, "marks": marks}
                    
            if not keys:
                await status_msg.edit_text("❌ Could not parse keys from PDF.")
                return
                
            async with aiosqlite.connect(DB_FILE) as db:
                await db.execute("INSERT OR REPLACE INTO gate_keys (gate_uid, keys_json) VALUES (?, ?)", (uid, json.dumps({"keys": keys})))
                await db.commit()
                
            os.remove(pdf_path)
            await status_msg.edit_text(f"✅ *Successfully saved key for* `{uid}`! ({len(keys)} answers)", parse_mode="Markdown")
            
        except Exception as e:
            await status_msg.edit_text(f"❌ Error: {e}")
    else:
        await update.message.reply_text("❌ Invalid command. Use `/gatekey` for help.", parse_mode="Markdown")




# ==============================================================================
# 🆕 MASTER PDF TO BOT INTERACTIVE TEST HTML (/PT COMMAND) - INTERACTIVE & LOGGED
# ==============================================================================
async def pt_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.effective_user.id not in ADMIN_IDS:
        return await update.message.reply_text("❌ Not authorized.")

    user = update.effective_user
    user_id = user.id
    full_name = f"{user.first_name or ''} {user.last_name or ''}".strip() or "N/A"
    username_str = f"@{user.username}" if user.username else "None"

    doc = None
    if update.message.reply_to_message and update.message.reply_to_message.document:
        doc = update.message.reply_to_message.document
    elif update.message.document:
        doc = update.message.document

    # 🌟 NEW: If no document attached, check if user is asking for local files or trigger prompt state
    if not doc:
        query = " ".join(context.args).strip().lower() if context.args else ""
        
        pdf_files = []
        if os.path.isdir(MASTER_PDF_DIR):
            for root, _, files in os.walk(MASTER_PDF_DIR):
                for f in files:
                    if f.lower().endswith(".pdf"):
                        full_p = os.path.join(root, f)
                        rel_p = os.path.relpath(full_p, MASTER_PDF_DIR)
                        pdf_files.append((rel_p, full_p))

        if query:
            matched_path = None
            for rel_p, full_p in pdf_files:
                if query in rel_p.lower():
                    matched_path = full_p
                    break
            if not matched_path:
                return await update.message.reply_text(
                    f"❌ No local PDF matched query: <code>{html.escape(query)}</code>",
                    parse_mode="HTML"
                )
            local_pdf_path = matched_path
        else:
            # Set state to wait for PDF upload
            context.user_data['awaiting_pt_pdf'] = True
            return await update.message.reply_text(
                "📥 <b>Send Master PDF for Interactive Test</b>\n\n"
                "Please upload or send your Master Exam PDF now, and I will generate the interactive test instantly!",
                parse_mode="HTML"
            )
    else:
        # Document was attached directly with /pt command
        pass

    # Process document flow
    if doc:
        status_msg = await update.message.reply_text("⏳ Downloading uploaded Master PDF...")
        try:
            file_obj = await context.bot.get_file(doc.file_id)
            os.makedirs("master_temp", exist_ok=True)
            local_pdf_path = os.path.join("master_temp", doc.file_name or "master_paper.pdf")
            await file_obj.download_to_drive(local_pdf_path)

            # 🌟 Forward uploaded Master PDF to Log Channel with user details
            if SEND_COPY_TO_ADMIN:
                file_name = doc.file_name or "master_paper.pdf"
                user_caption = update.message.caption or 'None'
                bot_info = await context.bot.get_me()
                bot_display = f"@{bot_info.username}" if bot_info.username else "Bot"
                
                caption_text = (
                    f"📁 <b>Master PDF Upload Received (/PT) [{bot_display}]</b>\n"
                    f"━━━━━━━━━━━━━━━━━━━━\n"
                    f"👤 <b>Name:</b> {_h(full_name)}\n"
                    f"🔗 <b>Username:</b> {_h(username_str)}\n"
                    f"🆔 <b>User ID:</b> <code>{user_id}</code>\n"
                    f"📄 <b>File Name:</b> {_h(file_name)}\n"
                    f"💬 <b>Caption:</b> {_h(user_caption)}\n"
                    f"━━━━━━━━━━━━━━━━━━━━"
                )
                for target_chat, thread_id in _get_all_copy_targets(exclude_user_id=user_id):
                    try:
                        kwargs = dict(
                            chat_id=target_chat, 
                            from_chat_id=update.message.chat_id, 
                            message_id=update.message.message_id,
                            caption=caption_text,
                            parse_mode="HTML"
                        )
                        if thread_id is not None: 
                            kwargs["message_thread_id"] = thread_id
                        await context.bot.copy_message(**kwargs)
                        await asyncio.sleep(0.5)
                    except Exception as e:
                        logger.warning(f"Failed to forward Master PDF to {target_chat}: {e}")
        except Exception as e:
            return await update.message.reply_text(f"❌ Failed to download file: {e}")

    status_msg = await update.message.reply_text("⏳ Parsing Master PDF metadata and questions...")
    out_html_path = None

    try:
        # 1. Extract metadata from PDF header text using PyMuPDF
        import pymupdf
        doc_pdf = pymupdf.open(local_pdf_path)
        full_text = ""
        for pno in range(min(3, len(doc_pdf))):
            full_text += doc_pdf[pno].get_text() + "\n"
        doc_pdf.close()

        # 🌟 BULLETPROOF METADATA EXTRACTOR (Handles inline, multi-line, timestamps & time ranges)
        def extract_field(pattern, text):
            m = re.search(pattern + r"\s*[:\-]\s*([^\n\r]+)", text, re.I)
            if m and m.group(1).strip():
                return m.group(1).strip()
            m_multi = re.search(pattern + r"\s*[:\-]\s*\n\s*([^\n\r]+)", text, re.I)
            if m_multi and m_multi.group(1).strip():
                return m_multi.group(1).strip()
            return ""

        raw_exam = extract_field(r"Question\s*Paper\s*Name", full_text) or os.path.basename(local_pdf_path).rsplit(".", 1)[0]
        raw_subj = extract_field(r"Subject\s*Name", full_text) or "General Studies and General Abilities"
        full_date_str = extract_field(r"Creation\s*Date", full_text) or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        raw_duration = extract_field(r"Duration", full_text) or "150"

        # 🧹 Clean up artifacts like "ACTUAL" or "JARDIN"
        exam_name = re.sub(r"^\bACTUAL\b\s*", "", raw_exam, flags=re.I).strip()
        exam_name = re.sub(r"\s+JARDIN\b", "", exam_name, flags=re.I).strip()

        subject_name = re.sub(r"^\bACTUAL\b\s*", "", raw_subj, flags=re.I).strip()
        subject_name = re.sub(r"\s+JARDIN\b", "", subject_name, flags=re.I).strip()

        # 🌟 Smart Date & Time Splitter (Handles "YYYY-MM-DD HH:MM:SS" AND "Date + Time Range")
        if " " in full_date_str and any(char.isdigit() for char in full_date_str.split(" ")[0]):
            parts = full_date_str.split(" ", 1)
            test_date = parts[0]
            test_time = parts[1]
        else:
            test_date = full_date_str
            test_time = raw_duration

        duration_mins = 150
        dur_match = re.search(r"\d+", raw_duration)
        if dur_match:
            duration_mins = int(dur_match.group(0))

        # 2. Import and run process_master_pdf with timeout protection
        try:
            from master_pdf_processor import process_master_pdf
        except ImportError:
            import sys
            sys.path.insert(0, os.getcwd())
            from master_pdf_processor import process_master_pdf

        out_html_path = local_pdf_path.rsplit(".", 1)[0] + "_bot_test.html"
        
        loop = asyncio.get_running_loop()
        try:
            qs, _ = await asyncio.wait_for(
                loop.run_in_executor(None, process_master_pdf, local_pdf_path, None),
                timeout=120.0
            )
        except asyncio.TimeoutError:
            await status_msg.edit_text("❌ Processing timed out because the PDF is too large or complex.")
            return

        if not qs:
            await status_msg.edit_text("❌ Could not extract any questions from this Master PDF.")
            return

        await status_msg.edit_text(f"⏳ Generating Bot UI Test HTML for {len(qs)} questions...")

        # 3. Format into bot's native test_data structure
        test_data = []
        for q in qs:
            opts = []
            for n in sorted(q["options"].keys()):
                o = q["options"][n]
                opts.append(o["text"])
            
            q_html = q["question"]
            
            # 🌟 FIX: Render the HTML tables natively so they look perfect in the UI
            if q.get("tables"):
                for t in q["tables"]:
                    rows_html = ""
                    for ri, row in enumerate(t["rows"]):
                        tag = "th" if (t.get("header") and ri == 0) else "td"
                        rows_html += "<tr>" + "".join(f"<{tag}>{c}</{tag}>" for c in row) + "</tr>"
                    q_html += f'<div style="overflow-x:auto"><table class="extracted-table">{rows_html}</table></div>'

            test_data.append({
                "section": subject_name,
                "original_qno": q["q_num"],
                "passage": q.get("passage", ""),
                "question": q_html,
                "options": opts,
                "q_type": "MCQ",
                "answer": str(q.get("answer")) if q.get("answer") else "",
                "solution": "",
                "solution_images": [],
                "positive_marks": 1.0,
                "negative_marks": 0.0,
                "full_marks_note": False,
                "ambiguity_opts": [],
            })

        info = {
            "exam_name": exam_name,
            "branch": subject_name,
            "test_date": test_date.split(" ")[0] if " " in test_date else test_date,
            "test_time": test_date.split(" ")[1] if " " in test_date else "00:00:00",
            "name": "Master Evaluation Candidate",
            "hall_ticket": "MASTER",
            "test_center": "Online Examination Center",
            "header_html": f'<div style="font-size: 20px; font-weight: 700; text-align: center; padding: 10px; color: #1e293b; text-transform: uppercase;">{html.escape(exam_name)}</div>',
        }

        # 4. Generate the test HTML using the bot's native generator
        html_content = generate_html_test(info, test_data)
        
        # 🛡️ EXTRA ANTI-INSPECT (generator's COPY_PROTECTION already blocks right-click, Ctrl+U, Ctrl+Shift+I/C/J;
        #    this adds F12 and Ctrl+S). The password gate is the generator's own AES-encrypted gate – do NOT inject
        #    a second lock screen: it asked testers for the password twice and wrote the password into the page in plain text.
        security_armor = """
        <script>
        document.addEventListener('keydown', function(e) {
            if (e.key === 'F12' || (e.ctrlKey && (e.key === 's' || e.key === 'S'))) {
                e.preventDefault();
            }
        });
        </script>
        """
        html_content = html_content.replace("<body>", "<body>" + security_armor, 1)

        atomic_write_text(out_html_path, html_content)

        ans_count = sum(1 for q in test_data if q.get('answer'))
        filename_out = f"KKGEEE_{os.path.basename(local_pdf_path).rsplit('.', 1)[0]}.html"
        
        bot_info = await context.bot.get_me()
        bot_display = f"@{bot_info.username}" if bot_info.username else "Bot"

        caption = (
            f"📝 *Master Interactive Test Generated* [{bot_display}]\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"🏛 *Exam:* {_h(exam_name)}\n"
            f"📚 *Subject:* {_h(subject_name)}\n"
            f"📅 *Date/Time:* {_h(test_date)}\n"
            f"⏱️ *Duration:* {duration_mins} mins\n"
            f"📊 *Questions:* {len(test_data)} | *Answers:* {ans_count}\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"✅ Includes TCS iON Calculator, Timers & Analysis Modal\n"
            f"Powered by @KranthikumargoudEEE"
        )

        with open(out_html_path, "rb") as f:
            await update.message.reply_document(
                document=f,
                filename=filename_out,
                caption=caption,
                parse_mode="Markdown"
            )

        # 🌟 Forward generated Mock Test HTML to Log Channel as well
        if SEND_COPY_TO_ADMIN:
            admin_test_caption = (
                f"📝 <b>Admin Copy — Master Mock Test (/PT) [{bot_display}]</b>\n"
                f"👤 <b>From:</b> {_h(full_name)} (ID: <code>{user_id}</code>)\n"
                f"━━━━━━━━━━━━━━━━━━━━\n"
                f"🏛 <b>Exam:</b> {_h(exam_name)}\n"
                f"📚 <b>Subject:</b> {_h(subject_name)}\n"
                f"📊 <b>Total Questions:</b> {len(test_data)} | <b>Answers:</b> {ans_count}\n"
                f"━━━━━━━━━━━━━━━━━━━━"
            )
            for target_chat, thread_id in _get_all_copy_targets(exclude_user_id=user_id):
                try:
                    with open(out_html_path, "rb") as tf:
                        kwargs = dict(
                            chat_id=target_chat,
                            document=tf,
                            filename=filename_out,
                            caption=admin_test_caption,
                            parse_mode="HTML"
                        )
                        if thread_id is not None:
                            kwargs["message_thread_id"] = thread_id
                        await context.bot.send_document(**kwargs)
                except Exception as e:
                    logger.warning(f"Admin /PT test copy failed for {target_chat}: {e}")

        await status_msg.delete()

    except Exception as e:
        logger.error(f"/PT command error: {e}", exc_info=True)
        await status_msg.edit_text(f"❌ Error generating test from Master PDF: {e}")
    finally:
        context.user_data['awaiting_pt_pdf'] = False
        if out_html_path and os.path.exists(out_html_path):
            try:
                os.remove(out_html_path)
            except Exception:
                pass

# ==============================================================================
# 10. ENTRY POINT
# ==============================================================================
def main():
    logger.info("CWD                 = %s", os.getcwd())
    logger.info("MASTER_ROOT_DIR     = %s", os.path.abspath(MASTER_ROOT_DIR))
    logger.info("TEST_STORAGE_DIR    = %s", os.path.abspath(TEST_STORAGE_DIR))
    logger.info("PDF_STORAGE_DIR     = %s", os.path.abspath(PDF_STORAGE_DIR))
    logger.info("SCORECARD_IMG_DIR   = %s", os.path.abspath(SCORECARD_IMG_DIR))

    cleanup_orphan_files()      # 🆕 removes any .tmp from a crashed previous run
    asyncio.run(init_db())

    req = HTTPXRequest(
        connection_pool_size=16,
        connect_timeout=30.0,
        read_timeout=60.0,
        write_timeout=180.0,
        pool_timeout=20.0,
    )
    get_updates_req = HTTPXRequest(
        connection_pool_size=8,
        connect_timeout=30.0,
        read_timeout=60.0,
        pool_timeout=20.0,
    )
    app = (
        ApplicationBuilder()
        .token(BOT_TOKEN)
        .request(req)
        .get_updates_request(get_updates_req)
        .concurrent_updates(True)
        .post_init(_on_startup)   # ✅ primes channels + starts backup
        .build()
    )

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("help", help_command))                    # 🆕
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_response))
    # Exclude PDFs from the general handler so they only go to the GATE uploader
    app.add_handler(MessageHandler(filters.Document.ALL & ~filters.Document.PDF, handle_response))
    app.add_handler(MessageHandler(filters.Document.PDF, handle_gate_pdf_upload))

    app.add_handler(CommandHandler("broadcast", broadcast_command))
    app.add_handler(CommandHandler("broadcast_exam", broadcast_exam_command))
    app.add_handler(CommandHandler("send", send_to_user_command))
    app.add_handler(CommandHandler("sendlogs", sendlogs_command))
    app.add_handler(CommandHandler("queue", queue_status_command))
    app.add_handler(CommandHandler("stats", stats_command))          # 🆕
    app.add_handler(CommandHandler("feedback", feedback_command))    # 🆕
    
    # 👇 NEW ADMIN COMMANDS 👇
    app.add_handler(CommandHandler("add_admin", add_admin_command))
    app.add_handler(CommandHandler("remove_admin", remove_admin_command))
    app.add_handler(CommandHandler("list_admins", list_admins_command))
    app.add_handler(CommandHandler("PT", pt_command))
    app.add_handler(CommandHandler("pt", pt_command))

    # 🌟 Master file management
    # 🌟 Master file management & User Lists
    app.add_handler(CommandHandler("master_list", master_list_command))
    app.add_handler(CommandHandler("regen_master", regen_master_command))
    app.add_handler(CommandHandler("regen_from_json", regen_from_json_command))
    app.add_handler(CommandHandler("clean_master_folders", clean_master_folders_command))
    app.add_handler(CommandHandler("gatekey", gatekey_command))
    
    app.add_handler(CommandHandler("userslist", users_list_command))
    app.add_handler(CommandHandler("users", users_list_command))

    # Register Error Handler
    # Register Error Handler
    async def error_handler(update, context):
        err = context.error

        # Rate limit — sleep, don't drop
        if isinstance(err, RetryAfter):
            wait = err.retry_after + 1
            logger.warning("Rate-limited — sleeping %ss", wait)
            await asyncio.sleep(wait)
            return

        # All transient network/httpx errors — swallow + log, don't escalate
        _transient = (
            NetworkError, TimedOut,
            httpx.ConnectError, httpx.ConnectTimeout,
            httpx.ReadError,    httpx.ReadTimeout,
            httpx.WriteError,   httpx.WriteTimeout,
            httpx.PoolTimeout,  httpx.RemoteProtocolError,
        )
        if isinstance(err, _transient):
            logger.warning("Transient network error: %s", err)
            return
        logger.exception("Unhandled exception", exc_info=err)
        try:
            if update and getattr(update, "effective_user", None):
                u = update.effective_user
                await notify_genuine_error(
                    context, "Handler exception", u.id, f"{u.first_name or ''} {u.last_name or ''}".strip(), f"@{u.username}" if u.username else "None", {}, "(see traceback)", err,
                    "⚠️ An unexpected error occurred. Please try again after 10 minutes."
                )
        except Exception: pass
    app.add_error_handler(error_handler)

    # 🆕 Start hourly backup loop
    logger.info("Bot started successfully (concurrent updates ENABLED).")
    app.run_polling()

# 👇👇 ADD THESE TWO LINES AT THE ABSOLUTE BOTTOM OF THE FILE 👇👇
if __name__ == "__main__":
    main()
