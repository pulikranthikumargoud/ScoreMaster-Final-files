# Exam Score Evaluator Bot

A high-performance, asynchronous Telegram bot that parses DigiALM and TCS iON Computer Based Test (CBT) response sheets. It computes section-wise scores with customizable negative marking, generates Retina-quality scorecard images, compiles print-ready A4 PDFs without text squishing, and automatically maintains Excel leaderboards and user audit trails.

---

## Features

- **Universal DigiALM / TCS iON Parsing**: Works seamlessly with URLs or uploaded `.html` files for major exams (Coal India, AP Vidyut / AP Genco, SSC, RRB, GATE, BARC, HPCL, etc.).
- **Dynamic Marking Schemes**: 
  - Standard exams: Configurable positive and negative scoring per question.
  - GATE: Automatic per-question mark detection (1-mark vs. 2-mark) and question type classification (deductions for MCQs, zero negative marking for NAT and MSQ).
- **Retina 2x Scorecard Images**: Renders clear summary scorecards featuring the official exam banner, candidate metadata, and section breakdown tables.
- **Robust PDF Generation**: Automatically restructures nested table elements and reading comprehension blocks into full-width views, preventing narrow vertical text distortion.
- **Multi-Tab Excel Reporting**:
  - `exam_results.xlsx`: Master ranking list and individual exam/branch worksheets.
  - `user_submissions_history.xlsx`: Audit history tracking every submission mapped to Telegram user details.
- **Windows File Lock Prevention**: Gracefully handles open Excel spreadsheets and PDF viewers by saving timestamped fallback snapshots rather than crashing with a `PermissionError`.

---

## Project Structure

```text
scorebot/
├── bot.py
├── requirements.txt
├── README.md
├── .gitignore
├── exam_evaluator.db             (Auto-generated SQLite database)
├── exam_results.xlsx             (Auto-generated rankings)
├── user_submissions_history.xlsx (Auto-generated submission audit log)
├── downloaded_sheets/            (Auto-generated raw HTML archive)
├── downloaded_pdfs/              (Auto-generated formatted PDFs)
└── scorecard_images/             (Auto-generated summary cards)


Dependencies (requirements.txt)

python-telegram-bot>=20.0
playwright>=1.40.0
beautifulsoup4>=4.12.0
pandas>=2.0.0
openpyxl>=3.1.0
aiosqlite>=0.19.0
httpx>=0.25.0

Local Setup (VS Code / Terminal)
1. Set Up Environment
Windows (PowerShell):

# Create virtual environment
python -m venv venv

# Activate virtual environment
.\venv\Scripts\Activate.ps1

# Install requirements
pip install --upgrade pip
pip install -r requirements.txt

# Install Playwright Chromium headless browser binaries
playwright install chromium

(If PowerShell restricts script execution, run: Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass before activating).

Linux / macOS (Terminal):

# Create virtual environment
python3 -m venv venv

# Activate virtual environment
source venv/bin/activate

# Install requirements
pip install --upgrade pip
pip install -r requirements.txt

# Install Playwright Chromium headless browser binaries
playwright install chromium


2. Configure Your Bot Token & Links
You can keep your configuration hardcoded in bot.py or set it via environment variables:
BOT_TOKEN = os.getenv("BOT_TOKEN", "YOUR_BOT_TOKEN_HERE")
CHANNEL_LINK = os.getenv("CHANNEL_LINK", "[https://t.me/your_telegram_channel](https://t.me/your_telegram_channel)")
CHANNEL_LINK1 = os.getenv("CHANNEL_LINK1", "[https://whatsapp.com/channel/your_whatsapp_channel](https://whatsapp.com/channel/your_whatsapp_channel)")


3. Run Locally
From Terminal:
Bash
# Windows
python bot.py

# Linux / macOS
python3 bot.py

From VS Code:
Open bot.py.

Press Ctrl + Shift + P (or Cmd + Shift + P on macOS) and select Python: Select Interpreter.

Choose the virtual environment interpreter (./venv/Scripts/python.exe or ./venv/bin/python).

Press F5 or click the Run play button in the top right.

Deployment on Render (Background Worker)
Because the bot uses long-polling (run_polling()), it must be deployed as a Background Worker on Render.

1. Setup .gitignore
Ensure your runtime files and binaries are not pushed to your Git repository:
venv/
__pycache__/
*.pyc
*.db
*.xlsx
downloaded_sheets/
downloaded_pdfs/
scorecard_images/
2. Configure the Render Background Worker
Log in to Render and navigate to the Dashboard.

Click New + > Background Worker.

Connect your GitHub repository.

Set the following build and run configurations:

Environment / Runtime: Python 3

Build Command:
pip install -r requirements.txt && playwright install chromium --with-deps

(The --with-deps flag installs necessary Linux rendering libraries on Ubuntu).

Start Command:
python bot.py

3. Set Environment Variables (Optional)If you want to manage your credentials outside the codebase, add these under the Environment tab:VariableRecommended ValueBOT_TOKENYour Telegram Bot Token from BotFatherCHANNEL_LINKhttps://t.me/your_channelCHANNEL_LINK1https://whatsapp.com/channel/your_channelPYTHONUNBUFFERED1Click Create Background Worker to trigger deployment.Configuration ReferenceAdding New Exams to the Marking SchemeTo support a new exam, add its title keyword and marking tuple (positive_marks, negative_deduction) to EXAM_MARKING_SCHEMES in bot.py:
EXAM_MARKING_SCHEMES = {
    "COAL INDIA": (1.0, 0.0),       # +1.0 for right, 0.0 for wrong
    "VIDYUT": (1.0, 0.0),           # +1.0 for right, 0.0 for wrong
    "SSC": (2.0, 0.50),             # +2.0 for right, -0.50 for wrong
    "RRB": (1.0, 0.333),            # +1.0 for right, -0.333 for wrong
    "BARC": (3.0, 1.0),             # +3.0 for right, -1.0 for wrong
    "ISRO": (3.0, 0.75),            # +3.0 for right, -0.75 for wrong
    "DEFAULT": (1.0, 0.0),          # Fallback if no keyword matches
}

(GATE exams are handled dynamically and do not need a fixed entry).