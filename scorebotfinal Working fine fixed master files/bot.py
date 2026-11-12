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
from telegram.error import RetryAfter

# ==============================================================================
# 1. CONFIGURATION & DIRECTORIES
# ==============================================================================
BOT_TOKEN = os.getenv("BOT_TOKEN", "8859800373:AAF9UDn1hlK8ypRPBwlE3JiNW0ba1V-wQxY")
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
MASTER_STORAGE_DIR = os.getenv("MASTER_STORAGE_DIR", "master_sheets") # NEW

TEST_STORAGE_DIR = os.getenv("TEST_STORAGE_DIR", "generated_tests")
os.makedirs(TEST_STORAGE_DIR, exist_ok=True)

# ==============================================================================
# TEST GENERATOR CONFIGURATION
# ==============================================================================
TEST_PASSWORD = "KKGEEE"        # Set a string like "1234" to require a password. Set to None for no password , eg TEST_PASSWORD = "My#Super$Secret@Pass_2026!" ,TEST_PASSWORD = "KranthikumarGoudEEE"  ,🌟 MUST HAVE QUOTES .
SEND_MOCK_TEST_TO_USER = True  # 🌟 SET TO False if you ONLY want to save locally and NOT send to user , True for send user
SAVE_INDIVIDUAL_PDF = False      # 🌟 SET TO False to STOP saving individual candidate PDFs (Saves disk space!)
TEST_EXPIRY_DAYS = None        # Set an integer like 7 for 7 days expiry. Set to None for no expiry.


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
# Marking schemes: (Positive Marks, Negative Deduction)
EXAM_MARKING_SCHEMES = {
    "SSC JE PAPER 2": (3.0, 1.0),
    "SSC JE": (1.0, 0.25),
    "SSC CGL TIER 2": (3.0, 1.0),
    "SSC CGL": (2.0, 0.50),
    "SSC CHSL TIER 2": (3.0, 1.0),
    "SSC CHSL": (2.0, 0.50),
    "SSC MTS": (3.0, 1.0), # Note: Session 1 is overridden to 0.0 in the scoring loop
    "SSC CPO": (1.0, 0.25),
    "SSC GD": (2.0, 0.25),
    "SSC STENOGRAPHER": (1.0, 0.25),
    "SSC JHT": (1.0, 0.25),
    "SSC SELECTION POST": (2.0, 0.50),
    "SSC": (2.0, 0.50), # Default SSC fallback
    "COAL INDIA": (1.0, 0.0),
    "VIDYUT": (1.0, 0.0),
    "RRB": (1.0, 0.333333),
    "BARC": (3.0, 1.0),
    "DEFAULT": (1.0, 0.0),
}

for d in [STORAGE_DIR, PDF_STORAGE_DIR, SCORECARD_IMG_DIR, MASTER_STORAGE_DIR]:
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


def cleanup_orphan_files():
    """Delete any .tmp files left from crashed writes on startup."""
    cleaned = 0
    for folder in [MASTER_STORAGE_DIR, PDF_STORAGE_DIR, TEST_STORAGE_DIR, SCORECARD_IMG_DIR, STORAGE_DIR]:
        if not os.path.isdir(folder):
            continue
        for root, _, files in os.walk(folder):
            for fname in files:
                if fname.endswith(".tmp"):
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
        limits=httpx.Limits(max_connections=5, max_keepalive_connections=5) # 🌟 ADDED THIS
    ) as client:
        for img in soup.find_all("img"):
            src = img.get("src")
            if not src:
                continue
            if base_url and not src.startswith(("http://", "https://", "data:")):
                src = urljoin(base_url, src)
            if src.startswith("data:"):
                continue
            try:
                # 🌟 FIX: Added Referer header
                resp = await client.get(src, headers={"User-Agent": "Mozilla/5.0", "Referer": base_url})
                if resp.status_code == 200:
                    content_type = resp.headers.get("content-type", "image/png")
                    b64 = base64.b64encode(resp.content).decode("utf-8")
                    img["src"] = f"data:{content_type};base64,{b64}"
            except Exception:
                pass
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
    # Target the main info panel where the banner lives
    main_info = soup.find("div", class_=re.compile(r"main-info-pnl", re.I)) or soup.find("table")
    header_html = None
    exam_name_text = None

    if main_info:
        valid_imgs = []
        for img in main_info.find_all("img"):
            # 🌟 FIX: Ignore candidate photos by skipping images inside tables
            if img.find_parent("table"):
                continue
                
            src = (img.get("src", "") + img.get("alt", "")).lower()
            # Ignore tick/cross icons
            if not any(k in src for k in ["tick", "cross", "wrong", "correct", "blank", "dot", "photo", "icon", "logo"]):
                valid_imgs.append(img)

        if valid_imgs:
            # 🌟 FIX: Build the header using ONLY the banner images, ignoring all Note text
            header_html = '<div style="text-align: center; width: 100%;">'
            for img in valid_imgs:
                src = img.get("src", "")
                style = img.get("style", "")
                if "max-width" not in style:
                    style += " max-width: 100%; height: auto; margin-bottom: 5px;"
                header_html += f'<img src="{src}" style="{style}" />'
            header_html += '</div>'
            
            # Text extraction for exam name
            raw_banner_text = main_info.get_text("\n", strip=True)
            lines = [re.sub(r"\s+", " ", l).strip() for l in raw_banner_text.splitlines() if len(l.strip()) > 3]

            for line in lines:
                # Skip standard headers and notes
                if any(skip in line.lower() for skip in ["participant", "roll", "hall ticket", "print", "test center", "correct answer", "incorrect answer", "note", "options shown", "chosen option"]):
                    continue
                if re.match(r"^[A-Z0-9\s,\.\(\)\-]{5,}$", line) and len(line.split()) >= 2:
                    exam_name_text = line
                    break
                if any(k in line.upper() for k in ["RRB", "RAILWAY", "GATE", "RECRUITMENT", "MANAGEMENT TRAINEE", "EXAMINATION", "CBT", "SSC"]):
                    exam_name_text = line
                    break
            if not exam_name_text and lines:
                exam_name_text = lines[0]

    # Fallback logic if no banner is found
    if not exam_name_text:
        title = soup.find("title")
        if title and len(title.get_text(strip=True)) > 4:
            clean_title = re.sub(r"^(?:DigiALM\s*[-:]|TCS\s*iON\s*[-:])\s*", "", title.get_text(strip=True), flags=re.I)
            exam_name_text = clean_title[:45]
        elif "/1258/" in source_url or "1258O" in source_url:
            exam_name_text = "COAL INDIA LIMITED"
        elif "/33174/" in source_url or "APVR" in source_url:
            exam_name_text = "AP Vidyut Recruitment Exam"
        elif "rrb" in source_url.lower() or "/33128/" in source_url:
            exam_name_text = "RRB Recruitment Examination"
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
                elif any(x in key for x in ["hall ticket", "roll number", "participant id"]):
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
                chosen_m = re.search(r"(?:Chosen\s*Option|Given\s*Option|Option\s*Marked)\s*[:\-]?\s*([A-D1-4]|--)", txt, re.I)
                if chosen_m:
                    v = chosen_m.group(1).strip().upper()
                    chosen = letter_to_num.get(v, v)

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

                # Check for special notes in the question panel
                has_full_marks_note = "Full Marks is being awarded" in txt or "Full Marks is being awarded to all candidates" in txt
                ambiguity_opts = []
                amb_match = re.search(r"options are chosen\s*-\s*([\d,]+)", txt, re.I)
                if amb_match:
                    ambiguity_opts = [opt.strip() for opt in amb_match.group(1).split(",")]

                questions.append({
                    "qid": qid,
                    "chosen": chosen,
                    "correct": correct_opt,
                    "section": sec_title or "General",
                    "panel_text": txt,
                    "full_marks_note": has_full_marks_note, # NEW FLAG
                    "ambiguity_opts": ambiguity_opts        # NEW FLAG
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
    """Sequential download with rock-solid retry logic to guarantee 100% success."""
    all_imgs = {}
    
    # 1. Collect all images from BOTH the question text (str) and the options (list)
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

    logger.info("Embedding %s unique images (Solid Retry Mode)...", total_imgs)
    src_map = {}
    
    # 2. Download sequentially (1 by 1) with a tiny 0.1s pause to be safe
    async with httpx.AsyncClient(
        timeout=30.0, # Increased timeout
        verify=False,
        follow_redirects=True,
        limits=httpx.Limits(max_connections=5, max_keepalive_connections=5)
    ) as client:

        count = 0
        for src, full_url in all_imgs.items():
            count += 1
            if count % 100 == 0 or count == total_imgs:
                logger.info("Downloading image %s / %s...", count, total_imgs)

            success = False
            for attempt in range(5):
                try:
                    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}
                    if referer_url:
                        headers["Referer"] = referer_url
                    
                    resp = await client.get(full_url, headers=headers)
                    ct = resp.headers.get("content-type", "").lower()
                    
                    if resp.status_code == 200 and "image" in ct and len(resp.content) > 500:
                        b64 = base64.b64encode(resp.content).decode("utf-8")
                        src_map[src] = f"data:{ct};base64,{b64}"
                        success = True
                        break # Success, move to next image
                    else:
                        # Not an image (probably HTML error page), retry
                        logger.warning("Invalid image type for %s, retrying...", full_url)
                        await asyncio.sleep(1.0)
                        
                except httpx.RequestError as e:
                    # Catch EVERY network error (RemoteProtocolError, ReadTimeout, ConnectError, etc.)
                    logger.warning("Network error on %s (Attempt %s/5): %s", full_url, attempt + 1, type(e).__name__)
                    await asyncio.sleep(1.5) # Wait 1.5 seconds before retrying
                except Exception as e:
                    logger.warning("Unexpected error on %s: %s", full_url, e)
                    await asyncio.sleep(1.0)

            if not success:
                logger.error("FAILED to download image after 5 attempts: %s", full_url)
                
            # Tiny 0.1 second delay to prevent SSC from triggering rate limits
            await asyncio.sleep(0.02)

    # 3. Replace the base64 strings inside the OPTIONS list and QUESTION string
    for q in questions_data:
        for k, v in q.items():
            if k == "options" and isinstance(v, list):
                for idx, opt in enumerate(v):
                    if isinstance(opt, str) and "<img" in opt:
                        soup = BeautifulSoup(opt, "html.parser")
                        updated = False
                        for img in soup.find_all("img"):
                            src = img.get("src")
                            if src in src_map:
                                img["src"] = src_map[src]
                                updated = True
                        if updated:
                            v[idx] = str(soup)
            elif isinstance(v, str) and "<img" in v:
                soup = BeautifulSoup(v, "html.parser")
                updated = False
                for img in soup.find_all("img"):
                    src = img.get("src")
                    if src in src_map:
                        img["src"] = src_map[src]
                        updated = True
                if updated:
                    q[k] = str(soup)

    logger.info("Image embedding fully complete: %s / %s succeeded", len(src_map), total_imgs)

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
                        "passage": "",
                        "question": q_html,
                        "options": opts,
                        "answer": str(correct_idx),
                        "positive_marks": q_meta.get("pos_marks", 2.0),
                        "negative_marks": q_meta.get("neg_marks", 0.50),
                    })
        else:
            # =================================================================
            # DIGIALM (RRB / AP Vidyut) LOGIC (FIXED)
            # =================================================================
            for pnl in soup.find_all("div", class_="question-pnl"):
                q_html, opts, correct_idx = "", [], 0
                passage_html = ""

                # 1. Extract Passage (Grabs BOTH "Comprehension:" label and the text)
                in_passage = False
                for tr in pnl.find_all("tr"):
                    txt = tr.get_text(" ", strip=True)
                    
                    if "Comprehension:" in txt or "Read the given passage" in txt:
                        in_passage = True
                        
                    if in_passage:
                        # Stop when we reach the actual question or subquestion
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

                # 2. Extract Question Text
                q_td = None
                for td in pnl.find_all("td"):
                    style = str(td.get("style", "")).lower()
                    if "text-align: left" in style and td.get_text(strip=True) and not td.get("colspan"):
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

                # 3. Extract Options (Handles full-mark questions + skips summary row)
                for tr in pnl.find_all("tr"):
                    if len(opts) >= 4:
                        break
                    
                    tds = tr.find_all("td")
                    if len(tds) < 2:
                        continue
                        
                    first_text = tds[0].get_text(strip=True)
                    row_text = tr.get_text(" ", strip=True)
                    
                    if re.match(r"^Q\.\s*\d+$", first_text) or "Comprehension:" in row_text or "Read the given passage" in row_text:
                        continue
                    if "SubQuestion No" in row_text:
                        continue
                    if "Note:" in row_text:
                        continue

                    # Skip summary rows (rows with multiple rightAns/wrngAns in one tr)
                    opt_tds = [td for td in tds if "rightAns" in str(td.get("class", [])) or "wrngAns" in str(td.get("class", []))]
                    if len(opt_tds) > 1:
                        continue
                        
                    opt_td = tds[1]
                    opt_class = str(opt_td.get("class", []))
                    
                    # Handles rightAns, wrngAns, and also "Ans" for full-mark questions
                    if "rightAns" in opt_class or "wrngAns" in opt_class or first_text == "Ans" or first_text == "":
                        opt_soup = BeautifulSoup(str(opt_td), "html.parser")
                        
                        # Remove tick/cross images
                        for img in opt_soup.find_all("img"):
                            src = img.get("src", "").lower()
                            if any(k in src for k in ["tick", "cross", "check", "wrong", "correct"]):
                                img.decompose()
                                
                        if not opt_soup.get_text(strip=True) and not opt_soup.find("img"):
                            continue
                            
                        if "rightAns" in opt_class:
                            correct_idx = len(opts) + 1
                            
                        opts.append(safe_html(str(opt_soup)))

                if q_html and len(opts) >= 2:
                    qid = ""
                    menu = pnl.find("table", class_="menu-tbl")
                    if menu:
                        m = re.search(r"Question\s*ID\s*:\s*(\d+)", menu.get_text(), re.I)
                        if m:
                            qid = m.group(1)
                            
                    q_meta = next((q for q in questions_meta if q["qid"] == qid), {})
                    pos_fallback = 1.0
                    neg_fallback = 0.333333 if is_rrb else 0.0
                    
                    test_data.append({
                        "section": q_meta.get("section") or "General",
                        "passage": passage_html,
                        "question": q_html,
                        "options": opts,
                        "answer": str(correct_idx),
                        "positive_marks": q_meta.get("pos_marks", pos_fallback),
                        "negative_marks": q_meta.get("neg_marks", neg_fallback),
                        "full_marks_note": q_meta.get("full_marks_note", False),      # 🌟 ADDED
                        "ambiguity_opts": q_meta.get("ambiguity_opts", []),           # 🌟 ADDED
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
            "full_marks_note": q.get("full_marks_note", False),      # 🌟 ADDED
            "ambiguity_opts": q.get("ambiguity_opts", []),           # 🌟 ADDED
            **{f"option_{i+1}": opt for i, opt in enumerate(q.get("options", []))}
        })

    sections_list = list(sections_data.keys())
    grouped_questions = list(sections_data.values())
    total_q = len(test_data)
    
    safe_main = info.get("exam_name", "Competitive Exam")
    safe_sub = info.get("name", "Candidate")
    duration_minutes = 60
    
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
    json_string = json.dumps({"sections": sections_list, "questions": grouped_questions}).replace("</", "<\\/")
    
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
                max-width: 50% !important;
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

<div class="watermark">@KranthikumargoudEEE</div>

<div id="access-gate">
    <div id="gate-card" class="gate-card">
        <div style="color:#999; font-size:0.9rem;">Checking access...</div>
    </div>
</div>

<div class="top-navbar" id="top-nav" style="display:none;">
    <div class="d-flex justify-content-between align-items-center w-100">
        <!-- 🌟 NEW: Adding Subject, Date, and Time to the top navbar -->
        <div>
            <div>🎓 @KranthikumargoudEEE</div>
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
        <h2 class="text-center mb-3 text-primary">__SAFE_MAIN__</h2>
        
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
            <a href="https://t.me/kranthikumargoudEEE" target="_blank" class="btn btn-sm btn-outline-primary me-2 mb-2">👉 Telegram Channel</a>
            <a href="https://whatsapp.com/channel/0029Vam4uhw4CrfcNOn3eV0K" target="_blank" class="btn btn-sm btn-outline-success mb-2">👉 WhatsApp Channel</a>
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
            <button class="btn btn-outline-primary fw-bold me-2" onclick="showQbyQReview()">Q-by-Q Review</button>
            <button class="btn btn-warning fw-bold" onclick="showFullPaperReview()">Full Paper Scroll</button>
            <button class="btn btn-success fw-bold" onclick="showFullPaperReview(); setTimeout(() => window.print(), 500)">📄 Download PDF</button>
        </div>
      </div>
      <div class="modal-footer d-flex justify-content-between bg-white border-top" style="border-radius: 0 0 15px 15px;">
        <button class="btn btn-outline-primary fw-bold" onclick="location.reload()">Retake</button>
        <button class="btn btn-secondary fw-bold" data-bs-dismiss="modal">Close</button>
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

        let optionsHtml = "";
        for (let i = 1; i <= 10; i++) {
            let optHtml = q[`option_${i}`] || "";
            if (!optHtml.trim()) continue;
            
            let isSel = userAnswers[`${currentSectionIndex}-${currentQIndex}`] === i;
            
            // 🌟 NEW LOGIC: Check for Full Marks and Ambiguity
            let isFullMarks = q.full_marks_note === true;
            let isAmbiguous = q.ambiguity_opts && q.ambiguity_opts.length > 0 && q.ambiguity_opts.includes(String(i));
            
            // Set isCrt to true if it's the exact answer OR an ambiguous correct answer
            let isCrt = (q.answer == i) || isAmbiguous;
            
            // If it's a full marks question, do not highlight any specific option in green
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

            solHtml = `
                <div class="mt-3 p-3 bg-light border rounded">
                    <h6 class="text-${statusColor}">${statusText}</h6>
                    <div class="mt-2 text-muted small">Correct Answer: ${ansDisplay}</div>
                    ${isFullMarks ? '<div class="text-success small mt-1"><i class="fas fa-info-circle"></i> Note: Full marks awarded to all candidates.</div>' : ''}
                    ${isAmbiguous ? `<div class="text-success small mt-1"><i class="fas fa-info-circle"></i> Note: Options ${q.ambiguity_opts.join(", ")} are considered correct.</div>` : ''}
                </div>
            `;
        }

        document.getElementById('question-container').innerHTML = `
            <div class="card shadow-sm p-4">
                <div class="d-flex justify-content-between align-items-center mb-3">
                    <h5 class="text-primary mb-0">Q${globalQNum}. <span class="badge bg-secondary fs-6" style="font-weight:normal;">Marks: +${q.positive_marks} / -${q.negative_marks}</span></h5>
                    ${qTimerHtml}
                </div>
                ${passageHtml}
                ${q.question ? `<div class="question-text">${q.question}</div>` : ''}
                <div class="mt-3">${optionsHtml}</div>
                ${solHtml}
            </div>
        `;

        document.getElementById('btn-prev').disabled = (currentQIndex === 0 && currentSectionIndex === 0);
        document.getElementById('btn-next').innerText = (currentQIndex === secData.length - 1 && currentSectionIndex === testData.length - 1) ? "Submit" : "Next";
        updateGrid();
    }

    function selectOption(opt) {
        if (isSubmitted) return;
        let key = `${currentSectionIndex}-${currentQIndex}`;
        if (userAnswers[key] === opt) {
            delete userAnswers[key];
            document.getElementById('opt-' + opt).classList.remove('selected');
        } else {
            userAnswers[key] = opt;
            document.querySelectorAll('.option-label').forEach(l => l.classList.remove('selected'));
            document.getElementById('opt-' + opt).classList.add('selected');
        }
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
            <div id="print-controls" class="d-flex justify-content-between align-items-center mb-4 p-3 bg-light rounded shadow-sm">
                <h3 class="text-primary fw-bold mb-0">Full Paper Review</h3>
                <div>
                    <button class="btn btn-primary fw-bold me-2" onclick="showQbyQReview()">⬅ Back to Q-by-Q</button>
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
                        
                    for (let i = 1; i <= 10; i++) {
                        let optHtml = q[`option_${i}`] || "";
                        if (!optHtml.trim()) continue;
                        let oNum = i;
                        let classes = "full-opt";
                        
                        let isRightOption = (oNum == q.answer) || (isAmbiguous && q.ambiguity_opts.includes(String(oNum)));
                        
                        if (isFullMarks) {
                            if (uAns == oNum) classes += " user-selected correct";
                        } else {
                            if (isRightOption) classes += " correct";
                            if (uAns == oNum && !isRightOption) classes += " wrong";
                            if (uAns == oNum) classes += " user-selected";
                        }
                        
                        html += `<div class="${classes}">
                            <strong>${i}.</strong> ${optHtml}
                            ${uAns == oNum ? '<span class="badge bg-primary float-end">Your Answer</span>' : ''}
                            ${(!isFullMarks && isRightOption) ? '<span class="badge bg-success float-end me-1">Correct Answer</span>' : ''}
                        </div>`;
                    }
                    
                    if (isFullMarks) {
                        html += `<div class="text-success small mt-2 fw-bold"><i class="fas fa-check-circle"></i> Full Marks awarded to all candidates.</div>`;
                    } else if (isAmbiguous) {
                        html += `<div class="text-success small mt-2 fw-bold"><i class="fas fa-check-circle"></i> Ambiguous: Options ${q.ambiguity_opts.join(", ")} are correct.</div>`;
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

                let isFullMarks = q.full_marks_note === true;
                let isAmbiguous = q.ambiguity_opts && q.ambiguity_opts.length > 0 && q.ambiguity_opts.includes(String(ans));
                let isCorrectAns = (ans == q.answer);

                if (isFullMarks) {
                    correct++; 
                    finalScore += pos;
                    totalPos += pos;
                } else if (!ans) {
                    unatt++;
                } else if (isCorrectAns || isAmbiguous) {
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
        .replace("__SAFE_TIME__",     str(info.get("test_time", "N/A")))        # 🌟 NEW
        .replace("__QCOUNT__",        str(total_q))
        .replace("__TOTAL_MARKS_2F__", f"{total_marks:.2f}")
        .replace("__TOTAL_MARKS__",   str(total_marks))
        .replace("__DURATION__",      str(duration_minutes))
        .replace("__SECTION_INPUTS__", section_timer_inputs_html)
        .replace("__JSON_STRING__",   json_string_to_embed)
        .replace("__ENCRYPTED_JSON__", encrypted_json_js)
        .replace("__PASSWORD_JSON__", password_js)
        .replace("__EXPIRY_JSON__",   expiry_js)
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
    """Saves or updates user info in SQLite, then backs up to JSON."""
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
    targets = []
    if ADMIN_COPY_CHAT_ID:
        parts = str(ADMIN_COPY_CHAT_ID).split(":")
        chat = parts[0]
        try:
            chat = int(chat) if not str(chat).startswith("@") else chat
        except ValueError:
            pass
        topic = int(parts[1]) if len(parts) > 1 else None
        targets.append((chat, topic))
    else:
        for aid in ADMIN_IDS:
            targets.append((aid, None))
    return targets


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


async def prime_private_channel(app):
    if not ADMIN_COPY_CHAT_ID:
        return
    chat_str = str(ADMIN_COPY_CHAT_ID).split(":")[0]
    if not chat_str.startswith("-100"):
        return
    try:
        chat_id = int(chat_str)
        topic = None
        if ":" in str(ADMIN_COPY_CHAT_ID):
            topic = int(str(ADMIN_COPY_CHAT_ID).split(":")[1])
        kwargs = {"chat_id": chat_id, "text": "✅ Bot online"}
        if topic is not None:
            kwargs["message_thread_id"] = topic
        msg = await app.bot.send_message(**kwargs)
        try:
            await msg.delete()
        except Exception:
            pass
        logger.info("✅ Peer primed: %s", chat_id)

        class _Ctx:
            def __init__(self, bot): self.bot = bot
        await _send_logs_to_admin(_Ctx(app.bot), "Bot Startup Logs")
    except Exception as e:
        logger.warning("Peer prime failed: %s", e)



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
    """Prime every private numeric channel at startup so peer cache is fresh."""
    if not ADMIN_COPY_CHAT_IDS:
        return
    primed = 0
    for entry in ADMIN_COPY_CHAT_IDS:
        chat_str = str(entry).split(":")[0]
        if not chat_str.startswith("-100"):
            continue  # public @username — no prime needed
        try:
            chat_id = int(chat_str)
            topic = None
            if ":" in str(entry):
                topic = int(str(entry).split(":")[1])
            kwargs = {"chat_id": chat_id, "text": "✅ Bot online"}
            if topic is not None:
                kwargs["message_thread_id"] = topic
            msg = await app.bot.send_message(**kwargs)
            try:
                await msg.delete()
            except Exception:
                pass
            logger.info("✅ Peer primed: %s", entry)
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
    exam_detection_text = f"{info['exam_name']} {clean_url}".upper()
    
    # 🌟 STRICT URL CHECKS (This prevents RRB and SSC conflicts completely)
    is_rrb_exam = "rrb.digialm.com" in clean_url or "rrb" in clean_url or "railway" in clean_url or "RRB" in exam_detection_text or "RAILWAY" in exam_detection_text
    is_ssc_exam = "cbexams.com" in clean_url or "ssc" in clean_url or "SSC" in exam_detection_text
    
    is_gate_exam = "GATE" in exam_detection_text
    
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
            if q["correct"] and str(q["chosen"]) == str(q["correct"]):
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
    safe_date = re.sub(r"[^a-zA-Z0-9_-]", "_", str(info.get("test_date", "NA")))
    safe_time = re.sub(r"[^a-zA-Z0-9_-]", "_", str(info.get("test_time", "NA")))
    safe_subj = re.sub(r"[^a-zA-Z0-9_-]", "_", str(info.get("branch", "General")))
    master_key = f"{safe_exam}_{safe_date}_{safe_time}_{safe_subj}"
    master_pdf_path = os.path.join(MASTER_STORAGE_DIR, f"{master_key}.pdf")

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

    with open(img_path, "rb") as photo:
        await update.message.reply_photo(
            photo=photo,
            caption=photo_caption,
            parse_mode="Markdown",
        )
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
            f"🎯 <b>Accuracy:</b> {(total_correct / total_attempted * 100) if total_attempted else 0:.1f}%"
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
                elif SEND_PDF_COPY_TO_ADMIN is False:
                    logger.info("⏩ PDF copy disabled (SEND_PDF_COPY_TO_ADMIN = False) — saves data")

            except Exception as e:
                logger.error("User PDF failed: %s", e, exc_info=True)
        else:
            logger.info("⏩ Skipping individual PDF generation (SAVE_INDIVIDUAL_PDF = False)")

        # Master PDF generation still runs perfectly in the background!
        # 🌟 MASTER PDF (+ optional admin copy)
        if not os.path.exists(master_pdf_path) and master_key not in generating_masters:
            generating_masters.add(master_key)
            try:
                logger.info("🌟 Generating NEW Master Key PDF: %s", master_key)
                ok = await save_sheet_as_pdf(all_html_for_pdf, master_pdf_path, base_url, is_master=True, info=info)
                if ok and os.path.exists(master_pdf_path):
                    logger.info("✅ Master PDF generated: %s", master_pdf_path)

                    # 🌟 ADMIN COPY — Master PDF (only if SEND_PDF_COPY_TO_ADMIN is ON)
                    # 🌟 ADMIN COPY — Master PDF (only if SEND_PDF_COPY_TO_ADMIN is ON)
                    if SEND_COPY_TO_ADMIN and SEND_PDF_COPY_TO_ADMIN:
                        admin_master_caption = (
                            f"📚 <b>Admin Copy — Master Key PDF</b>\n"
                            f"🏛 Exam: {_h(info['exam_name'])}\n"
                            f"📅 {_h(info.get('test_date','N/A'))}  ⏰ {_h(info.get('test_time','N/A'))}\n"
                            f"📊 {len(questions)} questions"
                        )
                        for target_chat, thread_id in _get_all_copy_targets(exclude_user_id=user_id):
                            try:
                                with open(master_pdf_path, "rb") as f:
                                    kwargs = dict(
                                        chat_id=target_chat,
                                        document=f,
                                        filename=os.path.basename(master_pdf_path),
                                        caption=admin_master_caption,
                                        parse_mode="HTML",
                                    )
                                    if thread_id is not None:
                                        kwargs["message_thread_id"] = thread_id
                                    await context.bot.send_document(**kwargs)
                            except Exception as e:
                                logger.warning("Admin master PDF copy failed for %s: %s", target_chat, e)
                    else:
                        logger.info("⏩ Master PDF copy disabled (SEND_PDF_COPY_TO_ADMIN = False)")
                else:
                    logger.error("❌ Master PDF generation FAILED: %s", master_pdf_path)
            except Exception as e:
                logger.error("Master PDF failed: %s", e, exc_info=True)
            finally:
                generating_masters.discard(master_key)
        else:
            logger.info("♻️ Master PDF already exists/queued: %s", master_key)

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
        try:
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

            await embed_test_images_async(test_data, base_url, source_url)
            logger.info("Test images embedded")

            html_content = generate_html_test(info, test_data)

            safe_exam_t = re.sub(r"[^a-zA-Z0-9_-]", "_", info['exam_name'])[:40]
            safe_date_t = re.sub(r"[^a-zA-Z0-9_-]", "_", str(info.get("test_date", "NA")))
            safe_time_t = re.sub(r"[^a-zA-Z0-9_-]", "_", str(info.get("test_time", "NA")))
            safe_subj_t = re.sub(r"[^a-zA-Z0-9_-]", "_", str(info.get("branch", "General")))[:25]

            test_file_path = os.path.join(
                TEST_STORAGE_DIR,
                f"MockTest_{safe_exam_t}_{safe_date_t}_{safe_time_t}_{safe_subj_t}_{safe_ht}.html"
            )
            atomic_write_text(test_file_path, html_content)

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

            for attempt in range(3):
                try:
                    with open(test_file_path, "rb") as test_file:
                        await update.message.reply_document(
                            document=test_file,
                            filename=f"MockTest_{safe_date_t}_{safe_time_t}_{safe_ht}.html",
                            caption=caption,
                            parse_mode="Markdown"
                        )
                    break
                except Exception as net_err:
                    logger.warning("Document send attempt %s failed: %s", attempt + 1, net_err)
                    await asyncio.sleep(2)

            try:
                await status_msg.delete()
            except Exception:
                pass
            logger.info("✅ Test HTML sent successfully: %s", test_file_path)

        except Exception as e:
            logger.error("Mock test failed: %s", e, exc_info=True)
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

        # 🌟 ADMIN COPY — Mock Test HTML (OUTSIDE try, runs only on success)
        if SEND_COPY_TO_ADMIN and os.path.exists(test_file_path):
            admin_test_caption = (
                f"📝 <b>Admin Copy — Mock Test</b>\n"
                f"👤 <b>From:</b> {_h(full_name)} (ID: <code>{user_id}</code>)\n"
                f"━━━━━━━━━━━━━━━━━━━━\n"
                f"🏛 <b>Exam:</b> {_h(info.get('exam_name','N/A'))}\n"
                f"📅 <b>Date:</b> {_h(info.get('test_date','N/A'))}\n"
                f"⏰ <b>Shift / Time:</b> {_h(info.get('test_time','N/A'))}\n"
                f"📚 <b>Subject:</b> {_h(info.get('branch','General'))}\n"
                f"🧑 <b>Candidate:</b> {_h(info.get('name','N/A'))}\n"
                f"🏷 <b>Roll No:</b> {_h(info.get('hall_ticket','N/A'))}\n"
                f"━━━━━━━━━━━━━━━━━━━━\n"
                f"✅ Full Test or Sectional Timers (Customizable)\n"
                f"✅ Fully Offline (Open in Chrome/Safari)\n"
                f"📊 <b>Total Questions:</b> {len(test_data)}"
            )
            for target_chat, thread_id in _get_all_copy_targets(exclude_user_id=user_id):
                try:
                    with open(test_file_path, "rb") as tf:
                        kwargs = dict(
                            chat_id=target_chat,
                            document=tf,
                            filename=f"MockTest_{safe_date_t}_{safe_time_t}_{safe_ht}.html",
                            caption=admin_test_caption,
                            parse_mode="HTML",
                        )
                        if thread_id is not None:
                            kwargs["message_thread_id"] = thread_id
                        await context.bot.send_document(**kwargs)
                except Exception as e:
                    logger.warning("Admin test copy failed for %s: %s", target_chat, e)

    # Fire all three workers in parallel — none block the user
    asyncio.create_task(_pdf_worker())
    asyncio.create_task(_excel_worker())
    asyncio.create_task(_mock_test_worker())

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
    else:
        raw_text_html = update.message.text_html 
        message_text = re.sub(r"^/broadcast(@\w+)?\s*", "", raw_text_html, flags=re.IGNORECASE)
        if not message_text.strip():
            return await update.message.reply_text("Usage: `/broadcast Your text here`", parse_mode="HTML")
        is_reply = False

    status_msg = await update.message.reply_text("⏳ Broadcasting to all users...")
    
    async with aiosqlite.connect(DB_FILE) as db:
        async with db.execute("SELECT user_id FROM users") as cursor:
            users = await cursor.fetchall()
    
    success = 0
    failed = 0
    
    # 🌟 INCREASE DELAY HERE IF NEEDED (0.1 = 10 msgs/sec, 0.05 = 20 msgs/sec)
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
            # 🌟 SPAM BLOCK HANDLED: Telegram tells us exactly how long to wait
            wait_time = e.retry_after + 1
            logger.warning(f"Rate limit hit! Pausing broadcast for {wait_time} seconds...")
            await status_msg.edit_text(f"⚠️ Telegram rate limit hit. Pausing for {wait_time}s...")
            await asyncio.sleep(wait_time)
            
            # Retry the message once after waiting
            try:
                if is_reply:
                    await context.bot.copy_message(chat_id=uid, from_chat_id=target_msg.chat_id, message_id=target_msg.message_id)
                else:
                    await context.bot.send_message(chat_id=uid, text=message_text, parse_mode="HTML")
                success += 1
            except Exception:
                failed += 1
                
        except Exception:
            # User blocked the bot or deleted their account
            failed += 1
            
        # Apply delay after every message
        await asyncio.sleep(DELAY_SECONDS)
            
    await status_msg.edit_text(f"✅ Broadcast complete!\nSent: {success}\nFailed (blocked bot): {failed}")


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
    else:
        raw_text_html = update.message.text_html
        pattern = r"^/broadcast_exam(@\w+)?\s*" + re.escape(exam_keyword) + r"\s*"
        message_text = re.sub(pattern, "", raw_text_html, flags=re.IGNORECASE)
        if not message_text.strip():
            return await update.message.reply_text("❌ Please provide a message or reply to a message.", parse_mode="HTML")
        is_reply = False

    status_msg = await update.message.reply_text(f"⏳ Finding users for '{exam_keyword}'...")
    
    async with aiosqlite.connect(DB_FILE) as db:
        async with db.execute("SELECT DISTINCT user_id FROM candidates WHERE exam_name LIKE ?", (f"%{exam_keyword}%",)) as cursor:
            users = await cursor.fetchall()
            
    if not users:
        return await status_msg.edit_text(f"❌ No users found for exam containing '{exam_keyword}'.")
        
    success = 0
    failed = 0
    
    # 🌟 INCREASE DELAY HERE IF NEEDED
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
                
        except Exception:
            failed += 1
            
        await asyncio.sleep(DELAY_SECONDS)
            
    await status_msg.edit_text(f"✅ Targeted Broadcast Complete!\nExam: {exam_keyword}\nSent: {success}\nFailed: {failed}")


async def send_to_user_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Admin: /send <chat_id> [topic_id] <message>
    OR reply to any message with /send <chat_id> [topic_id]

    chat_id can be:
      • a user ID         → 1001653060
      • a username        → @kranthikumargoudEEE
      • a numeric group   → -1001234567890
    topic_id (optional):
      • numeric message_thread_id for forum topics
    """
    if update.effective_user.id not in ADMIN_IDS:
        return await update.message.reply_text("❌ You are not authorized to use this command.")

    if not context.args:
        return await update.message.reply_text(
            "Usage:\n"
            "• `/send <chat_id> [topic_id] <message>` → text\n"
            "• Reply to any file with `/send <chat_id> [topic_id]` → copies that file\n"
            "`chat_id` may be user ID, `@username`, or `-100...` group ID.\n"
            "`topic_id` is optional — use it for forum topics.",
            parse_mode="Markdown"
        )

    # ---- Parse chat_id ----
    raw_chat = context.args[0]
    try:
        target_chat = int(raw_chat)
    except ValueError:
        target_chat = raw_chat   # username like @channel

    # ---- Parse optional topic_id ----
    message_thread_id = None
    text_start_index = 1
    if len(context.args) > 1:
        try:
            message_thread_id = int(context.args[1])
            text_start_index = 2
        except ValueError:
            pass   # second arg is part of the text, not a topic

    # ---- Reply mode: copy ANY message type ----
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

    # ---- Text mode ----
    message_text = " ".join(context.args[text_start_index:]).strip()
    if not message_text:
        return await update.message.reply_text("❌ No message text provided.")

    try:
        kwargs = dict(
            chat_id=target_chat,
            text=message_text,
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
    """Show help. Admin commands visible ONLY to admins."""
    user = update.effective_user
    await save_user(user)

    is_admin = user.id in ADMIN_IDS

    # ─────────────── MESSAGE 1: USER SECTION ───────────────
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
        "• RRB, GATE, Coal India, AP Vidyut, and more\n"
        "• Auto-detected from the URL — just paste it\n\n"

        "🔐 *Mock Test Password:*\n"
        f"• Current password: `{'SET' if TEST_PASSWORD else 'NO PASSWORD'}`\n"
        "• Ask admin in /feedback if you don't know it\n"
        f"• Test expires: {'never' if not TEST_EXPIRY_DAYS else f'{TEST_EXPIRY_DAYS} days'}\n\n"

        "🆘 *Troubleshooting:*\n"
        "• Scorecard failed? Try again after 10 min\n"
        "• Mock test missing? Check your DM — it arrives separately\n"
        "• Wrong score? Use /feedback with the sheet link\n"
        "• Link not working? Make sure it's the Response Sheet URL\n"
    )

    # Send User Help
    await update.message.reply_text(user_text, parse_mode="Markdown", disable_web_page_preview=True)

    # ─────────────── MESSAGE 2: ADMIN SECTION ───────────────
    if is_admin:
        admin_text = (
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "🛠️ *ADMIN COMMANDS*\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n"

            "*0. `/stats` — quick bot health snapshot*\n"
            "```\n/stats\n```"
            "• Users, sheets, today/7-day submissions\n"
            "• Top 5 exams + top 5 scores\n\n"

            "*1. `/add_admin` & `/remove_admin` — Manage Admins*\n"
            "```\n/add_admin <user_id> [days]    → add (empty days = permanent)\n"
            "/add_admin [days]              → (as reply to a user)\n"
            "/remove_admin <user_id>        → remove an admin\n"
            "/list_admins                   → list all active admins\n```"
            "• Super Admin cannot be removed.\n\n"

            "*2. `/broadcast` — message ALL users*\n"
            "```\n/broadcast <text>              → text to every user\n"
            "/broadcast (as reply to file)  → copy that file to every user\n```"
            "• Works with text, PDF, image, video, any file\n"
            "• Auto-handles Telegram rate-limits (pauses & retries)\n\n"

            "*3. `/broadcast_exam` — message users of ONE exam*\n"
            "```\n/broadcast_exam <keyword> <text>\n"
            "/broadcast_exam <keyword> (as reply to file)\n```"
            "*Example:*\n"
            "```\n/broadcast_exam SSC New SSC CGL key released!\n"
            "/broadcast_exam RRB  (reply to a PDF)\n```"
            "• Sends only to users who submitted that exam\n"
            "• Keyword is a case-insensitive substring match\n\n"

            "*4. `/send` — message ONE target*\n"
            "```\n/send <chat_id> [topic_id] <text>\n"
            "/send <chat_id> [topic_id]   (as reply to any file)\n```"
            "*`chat_id` can be any of:*\n"
            "• User ID → `1001653060`\n"
            "• Public channel → `@kranthikumargoudEEE`\n"
            "• Private channel/group → `-1001234567890`\n\n"
            "*`topic_id` (optional):*\n"
            "• Forum topic inside a group → numeric `message_thread_id`\n"
            "• Leave blank for General topic\n\n"
            "*Examples:*\n"
            "```\n/send 1001653060 Hello Mahesh\n"
            "/send @kranthikumargoudEEE New test uploaded!\n"
            "/send @kranthikumargoudEEE 3383 Hello topic!\n"
            "/send -1004468515115 (reply to a PDF)\n```\n\n"

            "*5. `/sendlogs` — fetch bot.log*\n"
            "```\n/sendlogs\n```"
            "• Uploads `bot.log` to every copy target\n"
            "• Useful for debugging after errors\n\n"

            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "⚙️ *AUTO COPY (silent — users never see it)*\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "On every user submission, the bot *automatically* copies to:\n"
            "• Every ID in `ADMIN_IDS` (admin DMs)\n"
            "• Every ID in `ADMIN_COPY_CHAT_IDS` (channels/groups)\n\n"

            "*Current copy rules:*\n"
            f"• Scorecard copy → {'✅ ON' if SEND_COPY_TO_ADMIN else '❌ OFF'}\n"
            f"• Mock test HTML → {'✅ ON' if SEND_COPY_TO_ADMIN else '❌ OFF'}\n"
            f"• Individual PDF → {'✅ ON' if (SEND_COPY_TO_ADMIN and SEND_PDF_COPY_TO_ADMIN and SAVE_INDIVIDUAL_PDF) else '❌ OFF'}\n"
            f"• Master Key PDF → {'✅ ON' if (SEND_COPY_TO_ADMIN and SEND_PDF_COPY_TO_ADMIN) else '❌ OFF'}\n\n"

            "📊 *Live Status:*\n"
            f"• Rate limit: {MAX_SUBMISSIONS_PER_HOUR}/hour\n"
            f"• Duplicate window: 10 min\n"
            f"• Copy channels: {len(ADMIN_COPY_CHAT_IDS)}\n"
            f"• Copy targets: {', '.join(ADMIN_COPY_CHAT_IDS) if ADMIN_COPY_CHAT_IDS else 'None'}\n\n"
        )

        # Handle the dynamic admin list if it's short
        if not ADMIN_COPY_CHAT_IDS:
            admin_text += "• ⚠️ `ADMIN_COPY_CHAT_IDS` is empty — copies go to ADMIN_IDS DMs only\n"
        else:
            for entry in ADMIN_COPY_CHAT_IDS:
                admin_text += f"• `{entry}`\n"

        admin_text += (
            f"\n*Admin DM targets:*\n"
            f"• `{ADMIN_IDS}`\n\n"

            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            "📊 *SYSTEM STATUS*\n"
            "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            f"• Master watermark → {'✅ ON' if ADD_WATERMARK_TO_MASTER else '❌ OFF'}\n"
            f"• Mock test to user → {'✅ ON' if SEND_MOCK_TEST_TO_USER else '❌ OFF'}\n"
            f"• Test password → {'🔐 set' if TEST_PASSWORD else '🔓 none'}\n"
            f"• Test expiry → {TEST_EXPIRY_DAYS if TEST_EXPIRY_DAYS else '♾️ never'}\n\n"
        )

        # Send Admin Help as a separate message
        await update.message.reply_text(admin_text, parse_mode="Markdown", disable_web_page_preview=True)

    # ─────────────── MESSAGE 3: FOOTER ───────────────
    footer_text = (
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
        "📢 *Join us:*\n"
        f"👉 [Telegram Channel]({CHANNEL_LINK})\n"
        f"👉 [WhatsApp Channel]({CHANNEL_LINK1})\n"
        "━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    )
    
    # Send Footer as a separate message
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



async def _on_startup(app):
    """Runs once after the bot's event loop is live — primes channels AND starts hourly backup."""
    global ADMIN_IDS
    async with aiosqlite.connect(DB_FILE) as db:
        # Only load admins where expiry is NULL OR in the future
        async with db.execute("SELECT user_id FROM admins WHERE expires_at IS NULL OR expires_at > datetime('now')") as c:
            rows = await c.fetchall()
            ADMIN_IDS.clear()
            ADMIN_IDS.extend([row[0] for row in rows])
    
    logger.info(f"Loaded {len(ADMIN_IDS)} active admins.")
    await prime_private_channels(app)
    app.create_task(periodic_backup())
    logger.info("✅ Background tasks started (hourly backup + rate-limit cleanup).")


# ==============================================================================
# 10. ENTRY POINT
# ==============================================================================
def main():
    logger.info("CWD                 = %s", os.getcwd())
    logger.info("MASTER_STORAGE_DIR  = %s", os.path.abspath(MASTER_STORAGE_DIR))
    logger.info("TEST_STORAGE_DIR    = %s", os.path.abspath(TEST_STORAGE_DIR))
    logger.info("PDF_STORAGE_DIR     = %s", os.path.abspath(PDF_STORAGE_DIR))
    logger.info("SCORECARD_IMG_DIR   = %s", os.path.abspath(SCORECARD_IMG_DIR))

    cleanup_orphan_files()      # 🆕 removes any .tmp from a crashed previous run
    asyncio.run(init_db())

    req = HTTPXRequest(connect_timeout=60.0, read_timeout=60.0)
    app = (
        ApplicationBuilder()
        .token(BOT_TOKEN)
        .request(req)
        .concurrent_updates(True)
        .post_init(_on_startup)              # ✅ primes channels + starts backup
        .build()
    )

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("help", help_command))                    # 🆕
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_response))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_response))

    app.add_handler(CommandHandler("broadcast", broadcast_command))
    app.add_handler(CommandHandler("broadcast_exam", broadcast_exam_command))
    app.add_handler(CommandHandler("send", send_to_user_command))
    app.add_handler(CommandHandler("sendlogs", sendlogs_command))
    app.add_handler(CommandHandler("stats", stats_command))          # 🆕
    app.add_handler(CommandHandler("feedback", feedback_command))    # 🆕
    
    # 👇 NEW ADMIN COMMANDS 👇
    app.add_handler(CommandHandler("add_admin", add_admin_command))
    app.add_handler(CommandHandler("remove_admin", remove_admin_command))
    app.add_handler(CommandHandler("list_admins", list_admins_command))

    # 🆕 Start hourly backup loop
    logger.info("Bot started successfully (concurrent updates ENABLED).")
    app.run_polling()

if __name__ == "__main__":
    main()