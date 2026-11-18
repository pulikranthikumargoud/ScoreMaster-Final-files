import asyncio
import base64
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

# ==============================================================================
# 1. CONFIGURATION & DIRECTORIES
# ==============================================================================
BOT_TOKEN = os.getenv("BOT_TOKEN", "8859800373:AAF9UDn1hlK8ypRPBwlE3JiNW0ba1V-wQxY")

DB_FILE = os.getenv("DB_FILE", "exam_evaluator.db")
STORAGE_DIR = os.getenv("STORAGE_DIR", "downloaded_sheets")
PDF_STORAGE_DIR = os.getenv("PDF_STORAGE_DIR", "downloaded_pdfs")
SCORECARD_IMG_DIR = os.getenv("SCORECARD_IMG_DIR", "scorecard_images")

EXCEL_FILE = os.getenv("EXCEL_FILE", "exam_results.xlsx")
EXCEL_USER_HISTORY = os.getenv("EXCEL_USER_HISTORY", "user_submissions_history.xlsx")
UNRECOGNIZED_LOG = "unrecognized_submissions.txt"

CHANNEL_LINK = os.getenv("CHANNEL_LINK", "https://t.me/kranthikumargoudEEE")
CHANNEL_LINK1 = os.getenv("CHANNEL_LINK1", "https://whatsapp.com/channel/0029Vam4uhw4CrfcNOn3eV0K")

# Marking schemes: (Positive Marks, Negative Deduction)
EXAM_MARKING_SCHEMES = {
    "COAL INDIA": (1.0, 0.0),       # CIL: No negative marking
    "VIDYUT": (1.0, 0.0),           # AP Vidyut: No negative marking
    "SSC": (2.0, 0.50),             # SSC CGL/CHSL: +2.0, -0.50
    "RRB": (1.0, 0.333333),         # Railway RRB: +1.0, -1/3rd deduction
    "BARC": (3.0, 1.0),             # BARC: +3.0, -1.0
    "DEFAULT": (1.0, 0.0),          # Default fallback
}

for d in [STORAGE_DIR, PDF_STORAGE_DIR, SCORECARD_IMG_DIR]:
    os.makedirs(d, exist_ok=True)

excel_lock = asyncio.Lock()
history_lock = asyncio.Lock()
db_lock = asyncio.Lock()
playwright_semaphore = asyncio.Semaphore(2)

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)
logging.getLogger("httpx").setLevel(logging.WARNING)


# ==============================================================================
# 2. DATABASE INITIALIZATION
# ==============================================================================
async def init_db():
    if os.path.exists(DB_FILE):
        backup_db_name = f"backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
        shutil.copyfile(DB_FILE, backup_db_name)

    async with aiosqlite.connect(DB_FILE) as db:
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
    async with httpx.AsyncClient(timeout=15.0) as client:
        for img in soup.find_all("img"):
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
    return str(soup)


# ==============================================================================
# 4. ROBUST PDF GENERATOR (PASSAGE EXTRACTION & PAGE BLOAT FIX)
# ==============================================================================
PDF_PRINT_STYLES = """
<style>
    @page {
        size: A4 portrait;
        margin: 8mm 10mm 8mm 10mm;
    }
    @media print, all {
        html, body {
            width: 100% !important;
            min-width: 100% !important;
            margin: 0 !important;
            padding: 0 !important;
            font-size: 11px !important;
            line-height: 1.4 !important;
            background: #ffffff !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
        }

        table, tr, td, div, p {
            page-break-inside: auto !important;
            break-inside: auto !important;
        }

        .extracted-passage-box {
            display: block !important;
            width: 100% !important;
            background: #fafafa !important;
            border: 1px solid #d4d4d4 !important;
            border-left: 4px solid #1a73e8 !important;
            padding: 10px 14px !important;
            margin: 10px 0 14px 0 !important;
            font-size: 11.5px !important;
            line-height: 1.55 !important;
            color: #222 !important;
            word-break: normal !important;
            white-space: normal !important;
            page-break-inside: avoid !important;
        }

        .extracted-passage-box b, .extracted-passage-box strong {
            color: #0d47a1 !important;
            display: block !important;
            margin-bottom: 6px !important;
            font-size: 12px !important;
        }

        table.main-info, .main-info, table:first-of-type {
            width: 100% !important;
            border-collapse: collapse !important;
            margin-bottom: 8px !important;
            font-size: 11px !important;
        }

        .question-pnl, .questionRow, .main-table {
            width: 100% !important;
            border-bottom: 1px solid #dcdcdc !important;
            padding: 6px 0 !important;
            margin-bottom: 6px !important;
            display: block !important;
            clear: both !important;
        }

        .menu-tbl {
            float: right !important;
            width: 250px !important;
            border: 1px solid #c0c0c0 !important;
            background-color: #f9f9f9 !important;
            margin: 4px 0 6px 10px !important;
            font-size: 10px !important;
        }

        .watermark, div[style*="position: absolute"] {
            pointer-events: none !important;
            opacity: 0.15 !important;
        }

        img {
            max-width: 85% !important;
            height: auto !important;
        }

        button, .print-btn, #print, input[type="button"] {
            display: none !important;
        }
    }
</style>
"""

async def save_sheet_as_pdf(html_content: str, pdf_path: str, base_url: str = ""):
    try:
        html_ready = await embed_all_images(html_content, base_url)
        soup = BeautifulSoup(html_ready, "html.parser")

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
            if elem.has_attr("width"):
                del elem["width"]
            if elem.has_attr("style"):
                elem["style"] = re.sub(r'width\s*:\s*[^;]+;?', '', elem["style"], flags=re.I)

        if soup.head:
            soup.head.append(BeautifulSoup(PDF_PRINT_STYLES, "html.parser"))
        else:
            soup.insert(0, BeautifulSoup(PDF_PRINT_STYLES, "html.parser"))

        final_html = str(soup)

        async with playwright_semaphore:
            async with async_playwright() as p:
                browser = await p.chromium.launch(
                    headless=True,
                    args=["--disable-gpu", "--disable-dev-shm-usage", "--no-sandbox"],
                )
                page = await browser.new_page(viewport={"width": 1200, "height": 900})
                await page.set_content(final_html, wait_until="load", timeout=60000)
                await page.emulate_media(media="print")
                await page.wait_for_timeout(1500)

                pdf_bytes = await page.pdf(
                    format="A4",
                    landscape=False,
                    print_background=True,
                    prefer_css_page_size=True,
                    margin={"top": "8mm", "bottom": "8mm", "left": "10mm", "right": "10mm"},
                )
                await browser.close()

                target_path = pdf_path
                saved = False
                for _ in range(3):
                    try:
                        with open(target_path, "wb") as f:
                            f.write(pdf_bytes)
                        saved = True
                        logger.info(f"Complete PDF saved: {target_path}")
                        break
                    except PermissionError:
                        await asyncio.sleep(0.5)

                if not saved:
                    dir_name, file_name = os.path.split(pdf_path)
                    name_base, ext = os.path.splitext(file_name)
                    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                    snapshot_path = os.path.join(dir_name, f"{name_base}_snapshot_{ts}{ext}")
                    with open(snapshot_path, "wb") as f:
                        f.write(pdf_bytes)
                    logger.warning(f"File locked. Saved PDF snapshot to: '{snapshot_path}'")
    except Exception as e:
        logger.error(f"Playwright PDF error: {e}", exc_info=True)


# ==============================================================================
# 5. UNIVERSAL EXAM HEADER & NAME DETECTOR
# ==============================================================================
def extract_header_and_exam_name(soup: BeautifulSoup, source_url: str = ""):
    profile_table = soup.find("table", class_=re.compile(r"main-info|table|header|candidate", re.I)) or soup.find("table")
    header_html = None
    exam_name_text = None

    if profile_table:
        container = profile_table.parent
        while container and container.name not in ["body", "[document]"]:
            imgs = container.find_all("img")
            valid_imgs = [
                img for img in imgs
                if not any(k in (img.get("src", "") + img.get("alt", "")).lower()
                           for k in ["tick", "cross", "wrong", "correct", "blank", "dot"])
            ]

            if valid_imgs:
                banner_clone = BeautifulSoup(str(container), "html.parser")
                for t in banner_clone.find_all("table"):
                    t.decompose()
                for btn in banner_clone.find_all(["button", "a", "input"]):
                    if "print" in btn.get_text(strip=True).lower() or "print" in btn.get("onclick", "").lower():
                        btn.decompose()

                header_html = str(banner_clone)
                raw_banner_text = banner_clone.get_text("\n", strip=True)
                lines = [re.sub(r"\s+", " ", l).strip() for l in raw_banner_text.splitlines() if len(l.strip()) > 3]

                for line in lines:
                    if any(skip in line.lower() for skip in ["participant", "roll", "hall ticket", "print", "test center"]):
                        continue
                    if re.match(r"^[A-Z0-9\s,\.\(\)\-]{5,}$", line) and len(line.split()) >= 2:
                        exam_name_text = line
                        break
                    if any(k in line.upper() for k in ["RRB", "RAILWAY", "GATE", "RECRUITMENT", "MANAGEMENT TRAINEE", "EXAMINATION", "CBT"]):
                        exam_name_text = line
                        break

                if not exam_name_text and lines:
                    exam_name_text = lines[0]
                break

            for sibling in profile_table.find_previous_siblings():
                stxt = sibling.get_text(" ", strip=True)
                if "print" in stxt.lower() and len(stxt) < 10:
                    continue
                if len(stxt) > 5 and not any(k in stxt.lower() for k in ["impedance", "voltage", "circuit", "section", "q.", "participant"]):
                    exam_name_text = re.sub(r"\s+", " ", stxt).strip()
                    break

            if exam_name_text:
                break
            container = container.parent

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
    async with playwright_semaphore:
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

                questions.append({
                    "qid": qid,
                    "chosen": chosen,
                    "correct": correct_opt,
                    "section": sec_title or "General",
                    "panel_text": txt,
                })

    return candidate_info, questions


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
# 9. TELEGRAM HANDLERS
# ==============================================================================
async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user

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
        "Send your **DigiALM Response Sheet URL** (or upload your response HTML file) to evaluate your scorecard."
    )
    await update.message.reply_text(msg, parse_mode="Markdown")


async def handle_response(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    user_id = user.id
    full_name = f"{user.first_name or ''} {user.last_name or ''}".strip() or "N/A"
    username_str = f"@{user.username}" if user.username else "None"

    raw_text = update.message.text.strip() if update.message.text else ""
    is_document = bool(update.message.document)
    file_name = update.message.document.file_name if is_document else "N/A"

    # =========================================================================
    # 1. IMMEDIATE TERMINAL LOGGING (Captures ALL user messages)
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
    url_match = re.search(r"(https?://\S+|(?:\w+\.)?digialm\.com/\S+)", raw_text, re.I)

    # Save unfamiliar/unrecognized inputs to backlog file for inspection
    if not is_document and (not url_match or "digialm.com" not in raw_text.lower()):
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

        status_msg = await update.message.reply_text("⏳ Fetching response sheet...")
        try:
            async with httpx.AsyncClient(timeout=40.0, follow_redirects=True, verify=False) as client:
                res = await client.get(
                    source_url,
                    headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
                )
                if res.status_code != 200:
                    await status_msg.edit_text(f"❌ Failed to download sheet (HTTP {res.status_code}).")
                    return
                html_content = res.text
        except Exception as e:
            await status_msg.edit_text(f"❌ Network Error: {e}")
            return
    elif is_document:
        status_msg = await update.message.reply_text("⏳ Reading file...")
        try:
            doc = await context.bot.get_file(update.message.document.file_id)
            raw_b = await doc.download_as_bytearray()
            html_content = raw_b.decode("utf-8", errors="ignore")
        except Exception as e:
            await status_msg.edit_text(f"❌ Error: {e}")
            return
    else:
        await update.message.reply_text("❌ Please send a valid DigiALM response sheet URL or upload the HTML file.")
        return

    # =========================================================================
    # 3. SHEET PARSING & SCORING
    # =========================================================================
    info, questions = parse_digialm_html(html_content, source_url)
    if not questions:
        await status_msg.edit_text("❌ Could not parse questions from this sheet. Start again /start and send Response URL.")
        return

    # Multi-signal match across exam banner, branch/subject, and URL
    exam_detection_text = f"{info['exam_name']} {info.get('branch', '')} {source_url}".upper()
    is_gate_exam = "GATE" in exam_detection_text

    pos_mark, neg_mark = EXAM_MARKING_SCHEMES["DEFAULT"]
    if not is_gate_exam:
        for key, scheme in EXAM_MARKING_SCHEMES.items():
            if key in exam_detection_text:
                pos_mark, neg_mark = scheme
                if info["exam_name"] == "Competitive Exam":
                    info["exam_name"] = f"{key} Examination"
                break

    section_stats = {}
    total_correct = total_incorrect = total_attempted = 0
    total_score = 0.0
    total_max_marks = 0.0

    for q in questions:
        sec = q["section"]
        pnl_txt = q.get("panel_text", "")

        if is_gate_exam:
            is_two_mark = (
                "2 MARK" in sec.upper()
                or "2M" in sec.upper()
                or "MARKS : 2" in pnl_txt.upper()
                or "MARKS:2" in pnl_txt.upper()
            )
            q_pos = 2.0 if is_two_mark else 1.0

            is_nat_or_msq = (
                "NAT" in sec.upper()
                or "MSQ" in sec.upper()
                or "QUESTION TYPE : NAT" in pnl_txt.upper()
                or "QUESTION TYPE : MSQ" in pnl_txt.upper()
            )
            q_neg = 0.0 if is_nat_or_msq else (q_pos / 3.0)
        else:
            q_pos = pos_mark
            q_neg = neg_mark

        total_max_marks += q_pos

        if sec not in section_stats:
            section_stats[sec] = {"total": 0, "attempted": 0, "correct": 0, "wrong": 0, "score": 0.0}
        section_stats[sec]["total"] += 1

        if q["chosen"] not in ["--", ""]:
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
    # 4. DATABASE & ARCHIVE SAVING
    # =========================================================================
    async with db_lock:
        async with aiosqlite.connect(DB_FILE) as db:
            await db.execute("""
                INSERT OR REPLACE INTO candidates 
                (user_id, exam_name, hall_ticket, candidate_name, branch, total_score, total_questions, attempted, correct, incorrect, test_date, test_time, test_center, response_url, timestamp)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                user_id,
                info["exam_name"],
                info["hall_ticket"],
                info["name"],
                info["branch"],
                round(total_score, 2),
                len(questions),
                total_attempted,
                total_correct,
                total_incorrect,
                info["test_date"],
                info["test_time"],
                info["test_center"],
                source_url,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ))
            await db.execute("""
                INSERT INTO submission_logs
                (user_id, user_full_name, username, exam_name, hall_ticket, candidate_name, response_url, submitted_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                user_id,
                full_name,
                username_str,
                info["exam_name"],
                info["hall_ticket"],
                info["name"],
                source_url,
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            ))
            await db.commit()

    safe_exam = re.sub(r"[^a-zA-Z0-9_-]", "_", info["exam_name"])
    safe_ht = re.sub(r"[^a-zA-Z0-9_-]", "_", info["hall_ticket"])
    safe_name = re.sub(r"[^a-zA-Z0-9_-]", "_", info["name"])

    exam_html_dir = os.path.join(STORAGE_DIR, safe_exam)
    exam_pdf_dir = os.path.join(PDF_STORAGE_DIR, safe_exam)
    os.makedirs(exam_html_dir, exist_ok=True)
    os.makedirs(exam_pdf_dir, exist_ok=True)

    html_file = os.path.join(exam_html_dir, f"{safe_ht}_{safe_name}.html")
    pdf_file = os.path.join(exam_pdf_dir, f"{safe_ht}_{safe_name}.pdf")

    with open(html_file, "w", encoding="utf-8") as f:
        f.write(html_content)

    base_url = source_url.rsplit("/", 1)[0] + "/" if "/" in source_url.split("?")[0] else ""
    asyncio.create_task(save_sheet_as_pdf(html_content, pdf_file, base_url))
    
    await update_excel_sheet()
    await update_user_history_excel()

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

    # =========================================================================
    # 5. TERMINAL EVALUATION SUMMARY
    # =========================================================================
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
    print(f"🖼️ Scorecard   : {img_path}")
    print("─" * 75)
    print("📋 SECTION BREAKDOWN:")
    for s_name, s in section_stats.items():
        print(f"  • {s_name:<36} : {s['score']:>6.2f} Marks | R: {s['correct']:>2} | W: {s['wrong']:>2} | NA: {s['total'] - s['attempted']:>2}")
    print("═" * 75 + "\n")

    # =========================================================================
    # 6. SEND ONLY THE SCORECARD PHOTO WITH CHANNELS IN THE CAPTION
    # =========================================================================
    photo_caption = (
        f"📸 *Scorecard Snapshot* @KranthikumargoudEEE\n\n"
        f"📢 *Join our channel for latest updates:*\n"
        f"👉 [Click Here to Join Telegram Channel]({CHANNEL_LINK})\n"
        f"👉 [Click Here to Join Whatsapp Channel]({CHANNEL_LINK1})\n"
        f"{CHANNEL_LINK}"
    )

    if img_path and os.path.exists(img_path):
        with open(img_path, "rb") as photo:
            await update.message.reply_photo(
                photo=photo,
                caption=photo_caption,
                parse_mode="Markdown",
            )
        await status_msg.delete()
    else:
        await status_msg.edit_text("❌ Failed to generate scorecard image.")


# ==============================================================================
# 10. ENTRY POINT
# ==============================================================================
def main():
    asyncio.run(init_db())
    req = HTTPXRequest(connect_timeout=60.0, read_timeout=60.0)
    app = ApplicationBuilder().token(BOT_TOKEN).request(req).build()

    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(MessageHandler(filters.TEXT & (~filters.COMMAND), handle_response))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_response))

    logger.info("Bot started successfully: Single photo delivery with clickable channels in caption.")
    app.run_polling()


if __name__ == "__main__":
    main()