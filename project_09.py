"""
project_09.py
=====================================================================
CV Screening AI Agent -- SINGLE FILE VERSION (Streamlit)
Sab modules (config, languages, database, cv_parser, matcher,
email_sender, email_fetcher) isi ek file mein combine kar diye gaye hain,
taake alag-alag files manage na karni parein.

Run: streamlit run project_09.py
Note: Webhook listener (Email Forwarding method) is a SEPARATE small
Flask server and can't live inside a Streamlit app -- rakhein use
webhook_listener.py mein alag se agar chahiye.
=====================================================================
"""

# ============================== IMPORTS ==============================
import os
import re
import json
import smtplib
import imaplib
import hashlib
import sqlite3
import email as email_lib
from email.header import decode_header
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from datetime import datetime
from contextlib import contextmanager

import streamlit as st
import pandas as pd
from apscheduler.schedulers.background import BackgroundScheduler
from PyPDF2 import PdfReader
import docx
from dotenv import load_dotenv

load_dotenv()

# ============================================================
# SECTION 1: CONFIG
# ============================================================

def _get_setting(key: str, default: str = "") -> str:
    """Value pehle Streamlit Cloud 'Secrets' se dhoondta hai (agar wahan deploy
    hai), warna .env / normal environment variable se leta hai. Isse ek hi
    code local computer aur Streamlit Cloud dono par chal sakta hai."""
    try:
        if hasattr(st, "secrets") and key in st.secrets:
            return str(st.secrets[key])
    except Exception:
        pass
    return os.getenv(key, default)


# ---------------- COMPANY INFO ----------------
COMPANY_NAME = _get_setting("COMPANY_NAME", "Your Company Name")
COMPANY_EMAIL = _get_setting("COMPANY_EMAIL", "hr@yourcompany.com")

# ---------------- EMAIL RECEIVING (Method 1: IMAP Polling) ----------------
IMAP_HOST = _get_setting("IMAP_HOST", "imap.gmail.com")
IMAP_PORT = int(_get_setting("IMAP_PORT", "993"))
IMAP_EMAIL = _get_setting("IMAP_EMAIL", "")          # company email jahan CVs aati hain
IMAP_APP_PASSWORD = _get_setting("IMAP_APP_PASSWORD", "")  # Gmail App Password

# ---------------- EMAIL RECEIVING (Method 2: Gmail API) ----------------
GMAIL_API_ENABLED = _get_setting("GMAIL_API_ENABLED", "false").lower() == "true"
GMAIL_CREDENTIALS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials.json")
GMAIL_TOKEN_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "token.json")

# ---------------- EMAIL RECEIVING (Method 3: Forwarding + Webhook) ----------------
WEBHOOK_ENABLED = _get_setting("WEBHOOK_ENABLED", "false").lower() == "true"
WEBHOOK_SECRET = _get_setting("WEBHOOK_SECRET", "change-this-secret")
WEBHOOK_PORT = int(_get_setting("WEBHOOK_PORT", "5000"))

# ---------------- EMAIL SENDING (Selected/Rejected emails) ----------------
SMTP_HOST = _get_setting("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(_get_setting("SMTP_PORT", "587"))
SMTP_EMAIL = _get_setting("SMTP_EMAIL", "")          # from email address
SMTP_APP_PASSWORD = _get_setting("SMTP_APP_PASSWORD", "")

# ---------------- INBOX CHECK INTERVAL ----------------
CHECK_INTERVAL_MINUTES = 2   # har 2 minute baad inbox check hoga

# ---------------- AI MATCHING ----------------
# Agar ANTHROPIC_API_KEY set hai to LLM-based smart matching hogi,
# warna simple keyword-matching fallback use hogi.
ANTHROPIC_API_KEY = _get_setting("ANTHROPIC_API_KEY", "")
MATCH_SCORE_THRESHOLD = 70   # is se upar score = Selected

# ---------------- ADMIN LOGIN (default, first run pe use karein) ----------------
DEFAULT_ADMIN_USERNAME = _get_setting("DEFAULT_ADMIN_USERNAME", "admin")
DEFAULT_ADMIN_PASSWORD = _get_setting("DEFAULT_ADMIN_PASSWORD", "admin123")


DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "cv_agent.db")
ATTACHMENTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "attachments")

# ============================================================
# SECTION 2: LANGUAGES (Urdu / Pashto / English)
# ============================================================

TRANSLATIONS = {
    "en": {
        "login_title": "Login to Dashboard",
        "username": "Username",
        "password": "Password",
        "login_button": "Login",
        "login_error": "Invalid username or password",
        "logout": "Logout",
        "nav_dashboard": "Dashboard",
        "nav_post_job": "Post Job",
        "nav_candidates": "Candidates",
        "nav_selected": "Selected",
        "nav_rejected": "Rejected",
        "nav_settings": "Settings",
        "total_cvs": "Total CVs Received",
        "selected_count": "Selected",
        "rejected_count": "Rejected",
        "pending_count": "Pending Review",
        "job_title": "Job Title",
        "job_description": "Job Description",
        "job_requirements": "Required Skills / Qualification",
        "post_job_button": "Post Job",
        "job_posted_success": "Job posted successfully",
        "candidate_name": "Name",
        "candidate_email": "Email",
        "candidate_score": "Match Score",
        "candidate_status": "Status",
        "received_on": "Received On",
        "resend_email": "Resend Email",
        "check_inbox_now": "Check Inbox Now",
        "settings_title": "System Settings",
        "language": "Language",
    },
    "ur": {
        "login_title": "ڈیش بورڈ میں لاگ ان کریں",
        "username": "یوزر نیم",
        "password": "پاس ورڈ",
        "login_button": "لاگ ان",
        "login_error": "یوزر نیم یا پاس ورڈ غلط ہے",
        "logout": "لاگ آؤٹ",
        "nav_dashboard": "ڈیش بورڈ",
        "nav_post_job": "جاب پوسٹ کریں",
        "nav_candidates": "امیدوار",
        "nav_selected": "منتخب شدہ",
        "nav_rejected": "مسترد شدہ",
        "nav_settings": "سیٹنگز",
        "total_cvs": "کل موصول شدہ سی وی",
        "selected_count": "منتخب",
        "rejected_count": "مسترد",
        "pending_count": "زیر جائزہ",
        "job_title": "جاب کا عنوان",
        "job_description": "جاب کی تفصیل",
        "job_requirements": "مطلوبہ مہارتیں / قابلیت",
        "post_job_button": "جاب پوسٹ کریں",
        "job_posted_success": "جاب کامیابی سے پوسٹ ہوگئی",
        "candidate_name": "نام",
        "candidate_email": "ای میل",
        "candidate_score": "میچ اسکور",
        "candidate_status": "حیثیت",
        "received_on": "موصول ہونے کی تاریخ",
        "resend_email": "ای میل دوبارہ بھیجیں",
        "check_inbox_now": "ابھی ان باکس چیک کریں",
        "settings_title": "سسٹم کی سیٹنگز",
        "language": "زبان",
    },
    "ps": {
        "login_title": "ډشبورډ ته ننوتل",
        "username": "کارن نوم",
        "password": "پټ نوم",
        "login_button": "ننوتل",
        "login_error": "کارن نوم یا پټ نوم غلط دی",
        "logout": "وتل",
        "nav_dashboard": "ډشبورډ",
        "nav_post_job": "دنده خپروي",
        "nav_candidates": "کاندیدان",
        "nav_selected": "غوره شوي",
        "nav_rejected": "رد شوي",
        "nav_settings": "امستنې",
        "total_cvs": "ټول ترلاسه شوي CVs",
        "selected_count": "غوره شوي",
        "rejected_count": "رد شوي",
        "pending_count": "تر بیاکتنې لاندې",
        "job_title": "د دندې سرلیک",
        "job_description": "د دندې تفصیل",
        "job_requirements": "اړین مهارتونه / وړتیا",
        "post_job_button": "دنده خپره کړئ",
        "job_posted_success": "دنده په بریالیتوب سره خپره شوه",
        "candidate_name": "نوم",
        "candidate_email": "بریښنالیک",
        "candidate_score": "میچ سکور",
        "candidate_status": "حالت",
        "received_on": "د ترلاسه کیدو نیټه",
        "resend_email": "بریښنالیک بیا واستوئ",
        "check_inbox_now": "اوس ان باکس وګورئ",
        "settings_title": "د سیستم امستنې",
        "language": "ژبه",
    },
}


def t(key: str, lang: str = "en") -> str:
    """Given key aur language code, translated text wapis dega.
    Agar translation na mile to English fallback dega."""
    return TRANSLATIONS.get(lang, TRANSLATIONS["en"]).get(key, TRANSLATIONS["en"].get(key, key))


# ============================================================
# SECTION 3: DATABASE
# ============================================================




def _hash_password(password: str) -> str:
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


@contextmanager
def get_conn():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db():
    """Pehli dafa app chalane par tables banega + default admin user."""
    with get_conn() as conn:
        c = conn.cursor()
        c.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                created_at TEXT
            )
        """)
        c.execute("""
            CREATE TABLE IF NOT EXISTS jobs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                description TEXT,
                requirements TEXT,
                created_at TEXT,
                is_active INTEGER DEFAULT 1
            )
        """)
        c.execute("""
            CREATE TABLE IF NOT EXISTS candidates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                job_id INTEGER,
                name TEXT,
                email TEXT,
                phone TEXT,
                cv_path TEXT,
                cv_text TEXT,
                score REAL,
                status TEXT DEFAULT 'pending',   -- pending / selected / rejected
                source TEXT,                      -- imap / gmail_api / webhook
                received_at TEXT,
                FOREIGN KEY (job_id) REFERENCES jobs (id)
            )
        """)
        c.execute("""
            CREATE TABLE IF NOT EXISTS email_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                candidate_id INTEGER,
                email_type TEXT,   -- selected / rejected
                sent_at TEXT,
                status TEXT,       -- success / failed
                FOREIGN KEY (candidate_id) REFERENCES candidates (id)
            )
        """)

        # Default admin agar koi user nahi hai
        c.execute("SELECT COUNT(*) as cnt FROM users")
        if c.fetchone()["cnt"] == 0:
            c.execute(
                "INSERT INTO users (username, password_hash, created_at) VALUES (?, ?, ?)",
                (DEFAULT_ADMIN_USERNAME, _hash_password(DEFAULT_ADMIN_PASSWORD),
                 datetime.now().isoformat())
            )


def verify_user(username: str, password: str) -> bool:
    with get_conn() as conn:
        row = conn.execute("SELECT password_hash FROM users WHERE username = ?", (username,)).fetchone()
        if not row:
            return False
        return row["password_hash"] == _hash_password(password)


def add_job(title, description, requirements) -> int:
    with get_conn() as conn:
        cur = conn.execute(
            "INSERT INTO jobs (title, description, requirements, created_at) VALUES (?, ?, ?, ?)",
            (title, description, requirements, datetime.now().isoformat())
        )
        return cur.lastrowid


def get_active_jobs():
    with get_conn() as conn:
        return conn.execute("SELECT * FROM jobs WHERE is_active = 1 ORDER BY created_at DESC").fetchall()


def get_job_by_id(job_id):
    with get_conn() as conn:
        return conn.execute("SELECT * FROM jobs WHERE id = ?", (job_id,)).fetchone()


def update_job(job_id, title, description, requirements):
    with get_conn() as conn:
        conn.execute(
            "UPDATE jobs SET title = ?, description = ?, requirements = ? WHERE id = ?",
            (title, description, requirements, job_id)
        )


def delete_job(job_id):
    """Job ko permanently delete karta hai. Agar iske sath candidates linked
    hain, to unka job_id NULL kar deta hai taake candidate history na toote."""
    with get_conn() as conn:
        conn.execute("UPDATE candidates SET job_id = NULL WHERE job_id = ?", (job_id,))
        conn.execute("DELETE FROM jobs WHERE id = ?", (job_id,))


def deactivate_job(job_id):
    """Job ko 'inactive' kar deta hai — list se hat jati hai lekin data mehfooz rehta hai."""
    with get_conn() as conn:
        conn.execute("UPDATE jobs SET is_active = 0 WHERE id = ?", (job_id,))


def add_candidate(job_id, name, email, phone, cv_path, cv_text, source) -> int:
    with get_conn() as conn:
        cur = conn.execute("""
            INSERT INTO candidates (job_id, name, email, phone, cv_path, cv_text, status, source, received_at)
            VALUES (?, ?, ?, ?, ?, ?, 'pending', ?, ?)
        """, (job_id, name, email, phone, cv_path, cv_text, source, datetime.now().isoformat()))
        return cur.lastrowid


def candidate_exists(email: str, job_id) -> bool:
    """Duplicate CV/email dobara process na ho, isliye check karta hai."""
    with get_conn() as conn:
        row = conn.execute(
            "SELECT id FROM candidates WHERE email = ? AND job_id = ?", (email, job_id)
        ).fetchone()
        return row is not None


def update_candidate_score(candidate_id, score, status):
    with get_conn() as conn:
        conn.execute(
            "UPDATE candidates SET score = ?, status = ? WHERE id = ?",
            (score, status, candidate_id)
        )


def get_all_candidates():
    with get_conn() as conn:
        return conn.execute("""
            SELECT c.*, j.title as job_title FROM candidates c
            LEFT JOIN jobs j ON c.job_id = j.id
            ORDER BY c.received_at DESC
        """).fetchall()


def get_candidates_by_status(status):
    with get_conn() as conn:
        return conn.execute("""
            SELECT c.*, j.title as job_title FROM candidates c
            LEFT JOIN jobs j ON c.job_id = j.id
            WHERE c.status = ?
            ORDER BY c.received_at DESC
        """, (status,)).fetchall()


def get_dashboard_counts():
    with get_conn() as conn:
        total = conn.execute("SELECT COUNT(*) as cnt FROM candidates").fetchone()["cnt"]
        selected = conn.execute("SELECT COUNT(*) as cnt FROM candidates WHERE status='selected'").fetchone()["cnt"]
        rejected = conn.execute("SELECT COUNT(*) as cnt FROM candidates WHERE status='rejected'").fetchone()["cnt"]
        pending = conn.execute("SELECT COUNT(*) as cnt FROM candidates WHERE status='pending'").fetchone()["cnt"]
        return {"total": total, "selected": selected, "rejected": rejected, "pending": pending}


def log_email(candidate_id, email_type, status):
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO email_log (candidate_id, email_type, sent_at, status) VALUES (?, ?, ?, ?)",
            (candidate_id, email_type, datetime.now().isoformat(), status)
        )


def get_candidate_by_id(candidate_id):
    with get_conn() as conn:
        return conn.execute("SELECT * FROM candidates WHERE id = ?", (candidate_id,)).fetchone()


# ============================================================
# SECTION 4: CV PARSER
# ============================================================



def extract_text_from_pdf(file_path: str) -> str:
    text = ""
    try:
        reader = PdfReader(file_path)
        for page in reader.pages:
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
    except Exception as e:
        print(f"[cv_parser] PDF read error: {e}")
    return text


def extract_text_from_docx(file_path: str) -> str:
    text = ""
    try:
        d = docx.Document(file_path)
        text = "\n".join(p.text for p in d.paragraphs)
    except Exception as e:
        print(f"[cv_parser] DOCX read error: {e}")
    return text


def extract_text(file_path: str) -> str:
    """File extension ke hisab se sahi parser call karega."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".pdf":
        return extract_text_from_pdf(file_path)
    elif ext in (".docx", ".doc"):
        return extract_text_from_docx(file_path)
    else:
        return ""


def extract_email(text: str) -> str:
    match = re.search(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", text)
    return match.group(0) if match else ""


def extract_phone(text: str) -> str:
    # Pakistan aur general phone number patterns
    match = re.search(r"(\+?\d{1,3}[-.\s]?)?\(?\d{3,4}\)?[-.\s]?\d{3,4}[-.\s]?\d{3,4}", text)
    return match.group(0).strip() if match else ""


def extract_name(text: str, fallback_email: str = "") -> str:
    """Simple heuristic: CV ki pehli non-empty line usually naam hoti hai."""
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    if lines:
        first_line = lines[0]
        # Agar line bahut lambi hai (paragraph jaisi) to shayad naam nahi hai
        if len(first_line.split()) <= 5:
            return first_line
    if fallback_email:
        return fallback_email.split("@")[0]
    return "Candidate"


def parse_cv(file_path: str) -> dict:
    """Ek CV file se sara useful data nikal kar dictionary return karta hai."""
    text = extract_text(file_path)
    email = extract_email(text)
    phone = extract_phone(text)
    name = extract_name(text, email)
    return {
        "text": text,
        "email": email,
        "phone": phone,
        "name": name,
    }


# ============================================================
# SECTION 5: AI MATCHER
# ============================================================



def _keyword_match_score(cv_text: str, requirements: str) -> float:
    """Fallback method: kitne required keywords CV mein mojood hain, uska percentage."""
    req_words = set(re.findall(r"[a-zA-Z]+", requirements.lower()))
    req_words = {w for w in req_words if len(w) > 2}  # chote words (is, to, at) ignore
    cv_words = set(re.findall(r"[a-zA-Z]+", cv_text.lower()))

    if not req_words:
        return 0.0

    matched = req_words.intersection(cv_words)
    score = (len(matched) / len(req_words)) * 100
    return round(score, 2)


def _ai_match_score(cv_text: str, job_title: str, job_description: str, requirements: str) -> float:
    """Anthropic API se LLM based intelligent scoring."""
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

        prompt = f"""You are a recruitment screening assistant. Compare the candidate's CV
against the job requirements and give a match score from 0 to 100.

Job Title: {job_title}
Job Description: {job_description}
Required Skills/Qualifications: {requirements}

Candidate CV Text:
{cv_text[:4000]}

Respond ONLY with a JSON object in this exact format, nothing else:
{{"score": <number 0-100>, "reason": "<one short sentence>"}}"""

        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=200,
            messages=[{"role": "user", "content": prompt}]
        )
        raw_text = response.content[0].text.strip()
        raw_text = raw_text.replace("```json", "").replace("```", "").strip()
        data = json.loads(raw_text)
        return float(data.get("score", 0))
    except Exception as e:
        print(f"[matcher] AI matching failed, falling back to keyword match: {e}")
        return None


def calculate_match_score(cv_text: str, job_title: str, job_description: str, requirements: str) -> float:
    """Main function — yahi baaki app se call hoga."""
    if ANTHROPIC_API_KEY:
        score = _ai_match_score(cv_text, job_title, job_description, requirements)
        if score is not None:
            return score
    # Fallback
    return _keyword_match_score(cv_text, requirements)


def decide_status(score: float) -> str:
    return "selected" if score >= MATCH_SCORE_THRESHOLD else "rejected"


# ============================================================
# SECTION 6: EMAIL SENDER
# ============================================================




def _build_selected_email(candidate_name: str, job_title: str) -> tuple:
    subject = f"Congratulations! You have been shortlisted - {job_title}"
    body = f"""Dear {candidate_name},

Thank you for applying for the position of {job_title} at {COMPANY_NAME}.

We are pleased to inform you that you have been SELECTED after our initial review process.
Our HR team will contact you shortly via phone call to schedule your interview.

Please keep your phone accessible over the next few days.

Best regards,
HR Team
{COMPANY_NAME}
"""
    return subject, body


def _build_rejected_email(candidate_name: str, job_title: str) -> tuple:
    subject = f"Update on your application - {job_title}"
    body = f"""Dear {candidate_name},

Thank you for your interest in the position of {job_title} at {COMPANY_NAME}
and for taking the time to apply.

After careful review of your CV, we regret to inform you that we will not be moving
forward with your application at this time. This decision does not reflect on your
skills or potential — we simply had a highly competitive pool of candidates for this role.

We encourage you to apply for future openings that match your profile.

Best regards,
HR Team
{COMPANY_NAME}
"""
    return subject, body


def send_email(to_email: str, subject: str, body: str) -> bool:
    """Basic SMTP email sender. Returns True/False for success."""
    if not SMTP_EMAIL or not SMTP_APP_PASSWORD:
        print("[email_sender] SMTP credentials not configured in config.py / .env")
        return False

    try:
        msg = MIMEMultipart()
        msg["From"] = SMTP_EMAIL
        msg["To"] = to_email
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_EMAIL, SMTP_APP_PASSWORD)
            server.sendmail(SMTP_EMAIL, to_email, msg.as_string())
        return True
    except Exception as e:
        print(f"[email_sender] Failed to send email to {to_email}: {e}")
        return False


def notify_candidate(candidate_id: int):
    """Candidate ke status (selected/rejected) ke hisab se sahi email bhejta hai
    aur database mein log karta hai."""
    candidate = get_candidate_by_id(candidate_id)
    if not candidate:
        return False

    job_title = candidate["job_id"]  # will resolve title below via query if needed
    with get_conn() as conn:
        job_row = conn.execute("SELECT title FROM jobs WHERE id = ?", (candidate["job_id"],)).fetchone()
    job_title = job_row["title"] if job_row else "the position"

    if candidate["status"] == "selected":
        subject, body = _build_selected_email(candidate["name"], job_title)
        email_type = "selected"
    elif candidate["status"] == "rejected":
        subject, body = _build_rejected_email(candidate["name"], job_title)
        email_type = "rejected"
    else:
        return False

    success = send_email(candidate["email"], subject, body)
    log_email(candidate_id, email_type, "success" if success else "failed")
    return success


# ============================================================
# SECTION 7: EMAIL FETCHER (IMAP polling + Gmail API)
# ============================================================



ALLOWED_EXTENSIONS = (".pdf", ".docx", ".doc")


def _save_attachment(part, source_tag: str) -> str:
    filename = part.get_filename()
    if not filename:
        return ""
    filename = decode_header(filename)[0][0]
    if isinstance(filename, bytes):
        filename = filename.decode(errors="ignore")

    if not filename.lower().endswith(ALLOWED_EXTENSIONS):
        return ""

    os.makedirs(ATTACHMENTS_DIR, exist_ok=True)
    safe_name = f"{source_tag}_{datetime.now().strftime('%Y%m%d%H%M%S')}_{filename}"
    file_path = os.path.join(ATTACHMENTS_DIR, safe_name)
    with open(file_path, "wb") as f:
        f.write(part.get_payload(decode=True))
    return file_path


def process_new_cv(file_path: str, sender_email: str, source: str):
    """CV parse karo, active job se match karo, DB mein save karo, email bhejo."""
    parsed = parse_cv(file_path)
    candidate_email = parsed["email"] or sender_email
    if not candidate_email:
        print("[email_fetcher] Skipping CV, no email found:", file_path)
        return

    active_jobs = get_active_jobs()
    if not active_jobs:
        print("[email_fetcher] No active job posted, skipping matching for:", file_path)
        return

    # Sabse latest active job ke against match karte hain
    job = active_jobs[0]

    if candidate_exists(candidate_email, job["id"]):
        print(f"[email_fetcher] Candidate already processed: {candidate_email}")
        return

    candidate_id = add_candidate(
        job_id=job["id"],
        name=parsed["name"],
        email=candidate_email,
        phone=parsed["phone"],
        cv_path=file_path,
        cv_text=parsed["text"],
        source=source,
    )

    score = calculate_match_score(
        parsed["text"], job["title"], job["description"], job["requirements"]
    )
    status = decide_status(score)
    update_candidate_score(candidate_id, score, status)

    # Email automatically company ki taraf se candidate ko jayegi
    notify_candidate(candidate_id)
    print(f"[email_fetcher] Processed candidate {candidate_email} -> {status} ({score})")


# ---------------- METHOD 1: IMAP POLLING ----------------
def check_inbox_imap():
    """Har baar call hone par ek dafa inbox check karta hai (unseen emails)."""
    if not IMAP_EMAIL or not IMAP_APP_PASSWORD:
        print("[email_fetcher] IMAP credentials not configured, skipping IMAP check.")
        return

    try:
        mail = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT)
        mail.login(IMAP_EMAIL, IMAP_APP_PASSWORD)
        mail.select("inbox")

        status, messages = mail.search(None, "UNSEEN")
        if status != "OK":
            mail.logout()
            return

        for num in messages[0].split():
            status, data = mail.fetch(num, "(RFC822)")
            if status != "OK":
                continue
            msg = email_lib.message_from_bytes(data[0][1])
            sender = email_lib.utils.parseaddr(msg.get("From"))[1]

            for part in msg.walk():
                if part.get_content_maintype() == "multipart":
                    continue
                if part.get("Content-Disposition") is None:
                    continue
                file_path = _save_attachment(part, "imap")
                if file_path:
                    process_new_cv(file_path, sender, "imap")

        mail.logout()
    except Exception as e:
        print(f"[email_fetcher] IMAP check failed: {e}")


# ---------------- METHOD 2: GMAIL API ----------------
def check_inbox_gmail_api():
    """Gmail API se naye unread messages check karta hai (agar enabled ho).
    Requires: credentials.json (OAuth client) file GMAIL_CREDENTIALS_FILE par."""
    if not GMAIL_API_ENABLED:
        return

    try:
        from google.auth.transport.requests import Request
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow
        from googleapiclient.discovery import build
        import base64

        SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
        creds = None

        if os.path.exists(GMAIL_TOKEN_FILE):
            creds = Credentials.from_authorized_user_file(GMAIL_TOKEN_FILE, SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                if not os.path.exists(GMAIL_CREDENTIALS_FILE):
                    print("[email_fetcher] Gmail API credentials.json not found. See README.")
                    return
                flow = InstalledAppFlow.from_client_secrets_file(GMAIL_CREDENTIALS_FILE, SCOPES)
                creds = flow.run_local_server(port=0)
            with open(GMAIL_TOKEN_FILE, "w") as token_file:
                token_file.write(creds.to_json())

        service = build("gmail", "v1", credentials=creds)
        results = service.users().messages().list(userId="me", q="is:unread has:attachment").execute()
        messages = results.get("messages", [])

        for m in messages:
            msg = service.users().messages().get(userId="me", id=m["id"]).execute()
            headers = msg["payload"].get("headers", [])
            sender = ""
            for h in headers:
                if h["name"] == "From":
                    sender = h["value"]

            parts = msg["payload"].get("parts", [])
            for part in parts:
                filename = part.get("filename")
                if filename and filename.lower().endswith(ALLOWED_EXTENSIONS):
                    att_id = part["body"].get("attachmentId")
                    if not att_id:
                        continue
                    att = service.users().messages().attachments().get(
                        userId="me", messageId=m["id"], id=att_id
                    ).execute()
                    file_data = base64.urlsafe_b64decode(att["data"].encode("UTF-8"))
                    os.makedirs(ATTACHMENTS_DIR, exist_ok=True)
                    file_path = os.path.join(
                        ATTACHMENTS_DIR,
                        f"gmailapi_{datetime.now().strftime('%Y%m%d%H%M%S')}_{filename}"
                    )
                    with open(file_path, "wb") as f:
                        f.write(file_data)
                    process_new_cv(file_path, sender, "gmail_api")

            # Mark as read so we don't reprocess
            service.users().messages().modify(
                userId="me", id=m["id"], body={"removeLabelIds": ["UNREAD"]}
            ).execute()

    except Exception as e:
        print(f"[email_fetcher] Gmail API check failed: {e}")


def run_all_checks():
    """Yeh function scheduler har 2 minute baad call karega — dono methods check honge."""
    print(f"[email_fetcher] Checking inbox at {datetime.now().isoformat()}")
    check_inbox_imap()
    check_inbox_gmail_api()


# ============================================================
# SECTION 8: STREAMLIT APP
# ============================================================

import streamlit as st
import pandas as pd
from apscheduler.schedulers.background import BackgroundScheduler

# ---------------- PAGE CONFIG ----------------
st.set_page_config(
    page_title=f"{COMPANY_NAME} - CV Screening Agent",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------- INIT DATABASE ----------------
init_db()

# ---------------- PROFESSIONAL CSS ----------------
CUSTOM_CSS = """
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}

    .main {
        background-color: #f4f6f9;
    }

    .app-navbar {
        background: linear-gradient(90deg, #0f172a 0%, #1e3a8a 100%);
        padding: 18px 28px;
        border-radius: 10px;
        margin-bottom: 22px;
    }
    .app-navbar h1 {
        color: #ffffff;
        font-size: 26px;
        margin: 0;
        font-weight: 700;
    }
    .app-navbar p {
        color: #cbd5e1;
        margin: 2px 0 0 0;
        font-size: 14px;
    }

    .metric-card {
        background: #ffffff;
        border-radius: 12px;
        padding: 20px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        border-left: 5px solid #1e3a8a;
        text-align: center;
    }
    .metric-card h2 {
        font-size: 32px;
        margin: 0;
        color: #0f172a;
    }
    .metric-card p {
        margin: 4px 0 0 0;
        color: #64748b;
        font-size: 14px;
    }

    .status-selected {
        background-color: #dcfce7;
        color: #166534;
        padding: 3px 10px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
    }
    .status-rejected {
        background-color: #fee2e2;
        color: #991b1b;
        padding: 3px 10px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
    }
    .status-pending {
        background-color: #fef9c3;
        color: #854d0e;
        padding: 3px 10px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 600;
    }

    section[data-testid="stSidebar"] {
        background-color: #0f172a;
    }
    section[data-testid="stSidebar"] * {
        color: #e2e8f0 !important;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# ---------------- SESSION STATE DEFAULTS ----------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "lang" not in st.session_state:
    st.session_state.lang = "en"
if "scheduler_started" not in st.session_state:
    st.session_state.scheduler_started = False


# ---------------- BACKGROUND SCHEDULER (har 2 min inbox check) ----------------
def start_scheduler():
    if not st.session_state.scheduler_started:
        scheduler = BackgroundScheduler()
        scheduler.add_job(
            run_all_checks,
            "interval",
            minutes=CHECK_INTERVAL_MINUTES,
            id="inbox_check_job",
        )
        scheduler.start()
        st.session_state.scheduler_started = True


start_scheduler()

# ---------------- LANGUAGE SELECTOR (always visible, even on login) ----------------
lang_options = {"English": "en", "اردو (Urdu)": "ur", "پښتو (Pashto)": "ps"}
with st.sidebar:
    st.markdown("---")
    selected_lang_label = st.selectbox(
        "🌐 Language / زبان / ژبه",
        list(lang_options.keys()),
        index=list(lang_options.values()).index(st.session_state.lang),
    )
    st.session_state.lang = lang_options[selected_lang_label]

lang = st.session_state.lang


def T(key):
    return t(key, lang)


# ---------------- NAVBAR (Title/Subtitle always in English) ----------------
def render_navbar():
    st.markdown(f"""
        <div class="app-navbar">
            <h1>{COMPANY_NAME} — AI Recruitment Agent</h1>
            <p>Automated CV Screening &amp; Candidate Communication System</p>
        </div>
    """, unsafe_allow_html=True)


# ---------------- LOGIN PAGE ----------------
def login_page():
    render_navbar()
    col1, col2, col3 = st.columns([1, 1.2, 1])
    with col2:
        st.subheader(T("login_title"))
        username = st.text_input(T("username"))
        password = st.text_input(T("password"), type="password")
        if st.button(T("login_button"), use_container_width=True):
            if verify_user(username, password):
                st.session_state.logged_in = True
                st.rerun()
            else:
                st.error(T("login_error"))


# ---------------- DASHBOARD PAGE ----------------
def dashboard_page():
    counts = get_dashboard_counts()
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f'<div class="metric-card"><h2>{counts["total"]}</h2><p>{T("total_cvs")}</p></div>', unsafe_allow_html=True)
    with c2:
        st.markdown(f'<div class="metric-card"><h2>{counts["selected"]}</h2><p>{T("selected_count")}</p></div>', unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="metric-card"><h2>{counts["rejected"]}</h2><p>{T("rejected_count")}</p></div>', unsafe_allow_html=True)
    with c4:
        st.markdown(f'<div class="metric-card"><h2>{counts["pending"]}</h2><p>{T("pending_count")}</p></div>', unsafe_allow_html=True)

    st.write("")
    if st.button("🔄 " + T("check_inbox_now")):
        with st.spinner("Checking inbox..."):
            run_all_checks()
        st.success("Inbox checked.")
        st.rerun()

    st.write("")
    st.markdown("#### Recent Candidates")
    rows = get_all_candidates()
    if rows:
        df = pd.DataFrame([dict(r) for r in rows])[
            ["name", "email", "job_title", "score", "status", "source", "received_at"]
        ]
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("No candidates received yet.")


# ---------------- POST JOB PAGE ----------------
def post_job_page():
    st.markdown(f"#### {T('nav_post_job')}")

    # Agar edit mode mein hain, to purani values form mein pehle se bhar dein
    editing_job_id = st.session_state.get("editing_job_id", None)
    editing_job = get_job_by_id(editing_job_id) if editing_job_id else None

    with st.form("job_form", clear_on_submit=(editing_job is None)):
        title = st.text_input(T("job_title"), value=editing_job["title"] if editing_job else "")
        description = st.text_area(T("job_description"), value=editing_job["description"] if editing_job else "")
        requirements = st.text_area(
            T("job_requirements"),
            value=editing_job["requirements"] if editing_job else "",
            help="e.g. Python, 2 years experience, Bachelor's degree"
        )
        submit_label = "Update Job" if editing_job else T("post_job_button")
        submitted = st.form_submit_button(submit_label)
        if submitted:
            if title and requirements:
                if editing_job:
                    update_job(editing_job_id, title, description, requirements)
                    st.session_state.editing_job_id = None
                    st.success("Job updated successfully")
                else:
                    add_job(title, description, requirements)
                    st.success(T("job_posted_success"))
                st.rerun()
            else:
                st.error("Job title and requirements are required.")

    if editing_job:
        if st.button("Cancel Edit"):
            st.session_state.editing_job_id = None
            st.rerun()

    st.markdown("#### Active Jobs")
    jobs = get_active_jobs()
    if jobs:
        for j in jobs:
            j = dict(j)
            with st.container(border=True):
                col1, col2 = st.columns([4, 1])
                with col1:
                    st.markdown(f"**{j['title']}**")
                    st.caption(j["requirements"])
                with col2:
                    b1, b2 = st.columns(2)
                    with b1:
                        if st.button("✏️", key=f"edit_{j['id']}", help="Edit"):
                            st.session_state.editing_job_id = j["id"]
                            st.rerun()
                    with b2:
                        if st.button("🗑️", key=f"delete_{j['id']}", help="Delete"):
                            st.session_state[f"confirm_delete_{j['id']}"] = True

                if st.session_state.get(f"confirm_delete_{j['id']}", False):
                    st.warning(f"Delete '{j['title']}'? Linked candidate records will be kept but unlinked.")
                    c1, c2 = st.columns(2)
                    with c1:
                        if st.button("Yes, delete", key=f"confirm_yes_{j['id']}"):
                            delete_job(j["id"])
                            st.session_state[f"confirm_delete_{j['id']}"] = False
                            st.success("Job deleted.")
                            st.rerun()
                    with c2:
                        if st.button("Cancel", key=f"confirm_no_{j['id']}"):
                            st.session_state[f"confirm_delete_{j['id']}"] = False
                            st.rerun()
    else:
        st.info("No jobs posted yet.")


# ---------------- CANDIDATES / SELECTED / REJECTED PAGES ----------------
def candidates_table_page(status_filter=None, title_key="nav_candidates"):
    st.markdown(f"#### {T(title_key)}")
    rows = get_candidates_by_status(status_filter) if status_filter else get_all_candidates()
    if not rows:
        st.info("No records found.")
        return

    for r in rows:
        r = dict(r)
        badge_class = f"status-{r['status']}"
        with st.container(border=True):
            col1, col2, col3 = st.columns([3, 2, 1])
            with col1:
                st.markdown(f"**{r['name']}**  \n{r['email']}  \n📞 {r['phone'] or '-'}")
            with col2:
                st.markdown(f"Job: {r['job_title']}  \nScore: **{r['score']}**  \n"
                             f"<span class='{badge_class}'>{r['status'].upper()}</span>",
                             unsafe_allow_html=True)
            with col3:
                if st.button(T("resend_email"), key=f"resend_{r['id']}"):
                    notify_candidate(r["id"])
                    st.success("Email sent.")


# ---------------- SETTINGS PAGE ----------------
def settings_page():
    st.markdown(f"#### {T('settings_title')}")
    st.write(f"**Inbox check interval:** every {CHECK_INTERVAL_MINUTES} minutes")
    st.write(f"**IMAP configured:** {'Yes' if IMAP_EMAIL else 'No'}")
    st.write(f"**Gmail API enabled:** {'Yes' if GMAIL_API_ENABLED else 'No'}")
    st.write(f"**Webhook enabled:** {'Yes' if WEBHOOK_ENABLED else 'No'}")
    st.write(f"**AI Matching:** {'Claude API' if ANTHROPIC_API_KEY else 'Keyword fallback'}")
    st.write(f"**Match score threshold:** {MATCH_SCORE_THRESHOLD}")
    st.info("Settings ko update karne ke liye: Streamlit Cloud dashboard -> Manage app -> "
            "Settings -> Secrets mein values edit karein, phir app khud reboot ho jayegi.")


# ---------------- MAIN APP ROUTING ----------------
if not st.session_state.logged_in:
    login_page()
else:
    render_navbar()

    with st.sidebar:
        st.markdown(f"### {COMPANY_NAME}")
        page = st.radio("", [
            T("nav_dashboard"), T("nav_post_job"), T("nav_candidates"),
            T("nav_selected"), T("nav_rejected"), T("nav_settings")
        ])
        st.markdown("---")
        if st.button(T("logout")):
            st.session_state.logged_in = False
            st.rerun()

    if page == T("nav_dashboard"):
        dashboard_page()
    elif page == T("nav_post_job"):
        post_job_page()
    elif page == T("nav_candidates"):
        candidates_table_page(None, "nav_candidates")
    elif page == T("nav_selected"):
        candidates_table_page("selected", "nav_selected")
    elif page == T("nav_rejected"):
        candidates_table_page("rejected", "nav_rejected")
    elif page == T("nav_settings"):
        settings_page()
