# Cold Email Automation System

**Automated Cold Email Outreach with Smart Follow-ups, Reply Detection & Resume Attachment**

This system sends personalized cold emails from your Gmail account to a list of contacts, automatically follows up at Day 2 and Day 5, detects replies, and stops when someone responds. Everything runs from your local machine — no cloud, no SaaS.

---

## What It Does

- Sends an initial email with your resume attached (Day 0)
- Sends a follow-up reminder in the same thread (Day 2)
- Sends a final follow-up in the same thread (Day 5)
- Detects if someone replied and skips them automatically
- Removes completed contacts from your CSV so you don't email them again
- Saves all results to `completed/completed_prospects.csv`
- Works with unlimited CSV files (`mail/1.csv`, `mail/2.csv`, etc.)

---

## Folder Structure (What Each File Does)

```
Cold-Email-Automation/
│
├── cold_email_automation.py     ← Main script. Run this to send emails.
├── excel_to_csv_converter.py    ← Convert Tracxn Excel exports to CSV format.
├── START.bat                    ← Double-click to run on Windows (no terminal needed).
├── requirements.txt             ← All Python packages needed.
│
├── credentials.json             ← YOUR Google OAuth file (see Step 2 below).
├── .env                         ← YOUR API keys (see Step 3 below).
│
├── mail/                        ← PUT YOUR CSV FILES HERE.
│   └── sample_contacts.csv      ← Example of correct CSV format.
│
├── tracking/                    ← Auto-created. Stores email history (JSON files).
├── completed/                   ← Auto-created. Stores finished/bounced contacts.
│
└── README.md                    ← This file.
```

> **Note:** `tracking/` and `completed/` folders are auto-created when you run the script.
> `token.json` is also auto-created after your first Gmail login.

---

## Step-by-Step Setup Guide

### Step 1 — Install Python

1. Download Python 3.10 or newer from https://www.python.org/downloads/
2. During installation, **check the box "Add Python to PATH"**
3. Verify it installed: open Command Prompt and type:
   ```
   python --version
   ```
   You should see something like `Python 3.11.x`

---

### Step 2 — Set Up Gmail OAuth (Google Cloud Console)

This lets the script send emails from your Gmail. You need a `credentials.json` file.

**Follow these steps exactly:**

1. Go to https://console.cloud.google.com/
2. Sign in with the **Gmail account you want to send emails from**
3. Click **"Select a project"** at the top → then **"New Project"**
   - Name it anything (e.g. `cold-email-bot`) → click **Create**
4. In the left sidebar, go to **APIs & Services → Library**
5. Search for **"Gmail API"** → click it → click **Enable**
6. Go to **APIs & Services → OAuth consent screen**
   - Choose **External** → click Create
   - App name: anything (e.g. `Email Bot`)
   - User support email: your Gmail address
   - Developer contact email: your Gmail address
   - Click **Save and Continue** through all steps
   - On the last screen, click **Back to Dashboard**
7. Go to **APIs & Services → Credentials**
8. Click **+ Create Credentials → OAuth Client ID**
   - Application type: **Desktop app**
   - Name: anything (e.g. `EmailBot`)
   - Click **Create**
9. A popup shows your credentials → click **Download JSON**
10. Rename the downloaded file to exactly `credentials.json`
11. **Replace** the `credentials.json` in this folder with your downloaded file

> **Important:** The `credentials.json` included here is a template placeholder — it will NOT work. You must replace it with your own downloaded file.

---

### Step 3 — Get Gemini API Keys

The script uses Google's Gemini AI to generate personalized email content.

1. Go to https://aistudio.google.com/app/apikey
2. Sign in with any Google account
3. Click **Create API Key**
4. Copy the key
5. Open the `.env` file in this folder with Notepad
6. Replace `YOUR_GEMINI_API_KEY_1` with your key
7. For better reliability, create 3-4 keys (repeat steps 3-4) and fill in `KEY_2`, `KEY_3`, `KEY_4`
   - Multiple keys prevent hitting the free rate limit during large sends

**The `.env` file should look like this when filled:**
```
GOOGLE_API_KEY_1=AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
GOOGLE_API_KEY_2=AIzaSyYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY
GOOGLE_API_KEY_3=AIzaSyZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ
GOOGLE_API_KEY_4=AIzaSyAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA
GOOGLE_API_KEY=AIzaSyXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
HUGGINGFACEHUB_API_TOKEN=hf_XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
```

> **HuggingFace token** is optional. If you don't have one, leave it as-is — the script will still work.
> Get one free at https://huggingface.co/settings/tokens if needed.

---

### Step 4 — Create Virtual Environment & Install Packages

Open Command Prompt **inside the `Cold-Email-Automation` folder**:
- Navigate in File Explorer to the folder
- Click the address bar at the top → type `cmd` → press Enter

Then run these commands one by one:

```cmd
python -m venv venv
```
```cmd
venv\Scripts\activate
```
```cmd
pip install -r requirements.txt
```

Wait for all packages to install (may take 5-10 minutes the first time).

> **Note:** Every time you open a new terminal to run the script, you must run `venv\Scripts\activate` first. The `START.bat` file does this automatically.

---

### Step 5 — Add Your Resume

Place your resume PDF in the root folder (same level as `cold_email_automation.py`).

The resume filename in the code is `Gaurav_Resume.pdf`.
**You need to either:**
- Rename your resume file to `Gaurav_Resume.pdf`, **OR**
- Open `cold_email_automation.py` and search for `Gaurav_Resume.pdf` and change it to your filename

---

### Step 6 — Customize Email Templates

Open `cold_email_automation.py` and find this section near line 200:

```python
EMAIL_TEMPLATES = {
    1: {
        "subject": "Remote Internship / Full-time",
        "body": """Hi {first_name}, ...your email body here..."""
    },
    2: {
        "subject": None,       # Sends as reply in same thread
        "body": """Hi {first_name}, ...your follow-up..."""
    },
    3: {
        "subject": None,       # Sends as reply in same thread
        "body": """Hello {first_name}, ...your final follow-up..."""
    }
}
```

Edit the email bodies to match your use case. You can use:
- `{first_name}` — replaced with the contact's first name
- `{company_name}` — replaced with the company name

**Timing (when follow-ups are sent):**
```python
TIMING_CONFIG = {
    1: 0,   # Email 1: immediately
    2: 2,   # Email 2: 2 days after Email 1
    3: 5    # Email 3: 5 days after Email 1 (3 days after Email 2)
}
```
Change the numbers to adjust delays.

---

### Step 7 — Prepare Your Contact List (CSV Files)

Put your contacts in the `mail/` folder as CSV files named `1.csv`, `2.csv`, `3.csv`, etc.

**Required columns:**
```csv
company_name,first_name,email
Google,John,john@google.com
Microsoft,Sarah,sarah@microsoft.com
```

**Rules:**
- File names must be numbers: `1.csv`, `2.csv`, `3.csv` ... up to `999.csv`
- The script auto-detects common column name variations (Company, Name, Email, etc.)
- The script auto-extracts first names if you give full names
- Duplicate emails are automatically removed across all CSV files

**If you have Excel files (Tracxn exports):**

Use the included converter to generate CSVs from Excel:

```cmd
python excel_to_csv_converter.py your_excel_folder_name
```

This reads the `People 2.1` sheet from each `.xlsx` file and outputs numbered CSV files.

---

### Step 8 — Run the System

**Option A — Double-click (Windows, easiest):**
Double-click `START.bat` in the folder. It activates the virtual environment and starts the script automatically.

**Option B — Command line:**
```cmd
venv\Scripts\activate
python cold_email_automation.py
```

---

## What Happens When You Run It

**1. Gmail Account Selection**
```
1. Continue with same Gmail account
2. Switch to different Gmail account
```
- First time: Choose option 1. A browser window opens to authorize your Gmail.
- After first login: Your session is saved in `token.json` automatically.
- If you want to switch to a different Gmail account: Choose option 2.

**2. Campaign Mode Selection**
```
1. Create drafts only (SAFE - review before sending)
2. Send emails directly (LIVE - sends immediately)
```
- **Always start with option 1 (drafts)** to review how your emails look in Gmail
- Switch to option 2 when you're confident everything looks right

**3. Processing**
The script processes each contact in your CSV files:
```
============================================================
Processing: Google - John (john@google.com)
============================================================
  New prospect — sending Email 1...
  Attachment added: Gaurav_Resume.pdf
  Email sent! ID: 18d4f5e6a7b8c9d0
  Email 1 sent successfully!
```

---

## Output Files (What Gets Created)

| File/Folder | What It Contains |
|---|---|
| `token.json` | Your Gmail login session (auto-created) |
| `tracking/email_tracking_<your_email>.json` | All active prospects and email history |
| `tracking/email_archive_<your_email>.json` | Completed/archived prospects |
| `completed/completed_prospects.csv` | All finished campaigns (replied, 3 emails sent) |
| `completed/bounced_emails.csv` | Emails that bounced or are invalid |

---

## How Follow-ups Work

The system tracks every email you send and automatically sends the right follow-up:

| Day | Action |
|---|---|
| Day 0 | Email 1 sent with resume attached |
| Day 2 | Email 2 sent as reply in same thread |
| Day 5 | Email 3 sent as reply in same thread |
| After Day 5 | Contact marked as completed, removed from CSV |

If a contact replies at any stage, they are **immediately skipped** — no more emails sent.

---

## Troubleshooting

**Problem: "No CSV files found"**
- Make sure your files are in the `mail/` folder
- File names must be numbers: `1.csv`, `2.csv` etc. (not `contacts.csv`)
- File extension must be lowercase `.csv`

**Problem: "Gmail authentication failed" or browser doesn't open**
- Delete `token.json` if it exists, then run again
- Make sure `credentials.json` is your real downloaded file (not the placeholder)
- Make sure the Gmail API is enabled in your Google Cloud project (Step 2, point 5)

**Problem: "Missing required column: email"**
- Open your CSV and make sure columns are: `company_name`, `first_name`, `email`
- Column names are case-insensitive but must contain these words

**Problem: pip install fails**
- Make sure you ran `venv\Scripts\activate` first
- Try: `pip install --upgrade pip` then retry

**Problem: Script runs but no emails appear in Gmail drafts/sent**
- Make sure you selected the correct mode (1 = drafts, 2 = live send)
- Check Gmail's Drafts or Sent folder
- Try running with a CSV containing just 1-2 test rows first

**Problem: Follow-up emails not sending**
- The system checks the date. If it's not yet 2 days after Email 1, it won't send Email 2.
- This is intentional — it respects the timing config.

**Problem: "Rate limit" or Gemini API errors**
- Add more API keys in `.env` (KEY_2, KEY_3, KEY_4)
- The script auto-rotates between keys and pauses when needed

---

## Security Notes

- `credentials.json` — Keep this private. Don't share it.
- `token.json` — Keep this private. It gives access to your Gmail.
- `.env` — Keep this private. It contains your API keys.
- None of these files should be uploaded to GitHub.

---

## Gmail Daily Sending Limits

| Account Type | Daily Limit |
|---|---|
| Free Gmail | ~500 emails/day |
| Google Workspace (paid) | ~2,000 emails/day |

The script has built-in rate limiting (pauses every 50 emails) to stay within limits.

---

## Quick Reference

```
SETUP (one time):
1. Install Python 3.10+
2. Download credentials.json from Google Cloud Console
3. Get Gemini API keys from aistudio.google.com
4. Fill in .env file with your API keys
5. Replace credentials.json with your downloaded file
6. Run: python -m venv venv
7. Run: venv\Scripts\activate
8. Run: pip install -r requirements.txt
9. Put your resume PDF in the root folder (rename to Gaurav_Resume.pdf)
10. Edit email templates in cold_email_automation.py

EVERY RUN:
1. Put CSV files in mail/ (named 1.csv, 2.csv, etc.)
2. Double-click START.bat (or run: venv\Scripts\activate && python cold_email_automation.py)
3. Choose Gmail account (option 1 to keep same, option 2 to switch)
4. Choose mode (option 1 for drafts, option 2 to send live)
5. Watch the console output
```

---

## License

MIT License — see LICENSE file.
