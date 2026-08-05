# CV Screening AI Agent — Setup Guide

## 1. Install
```bash
cd cv_agent
python -m venv venv
source venv/bin/activate      # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

## 2. Configure
```bash
cp .env.example .env
```
`.env` file khol kar apni values bhar dein:
- **IMAP_EMAIL / IMAP_APP_PASSWORD** — Gmail App Password banayein (Google Account → Security →
  2-Step Verification → App Passwords). Yeh inbox polling ke liye zaroori hai (har 2 min check).
- **SMTP_EMAIL / SMTP_APP_PASSWORD** — usually same Gmail account, candidates ko email bhejne ke liye.
- **GMAIL_API_ENABLED=true** karne ke liye Google Cloud Console se `credentials.json`
  download karke `cv_agent/` folder mein rakhein (OAuth Desktop App type).
- **ANTHROPIC_API_KEY** — agar AI-smart matching chahiye (warna simple keyword matching chalegi).

## 3. Run the app
```bash
streamlit run app.py
```
Default login: `admin` / `admin123` (yeh `.env` mein change kar sakte hain).

## 4. (Optional) Run webhook listener — Method 3
Agar Email Forwarding + Webhook use karna hai (Zapier/Make.com se):
```bash
python webhook_listener.py
```
Forwarding service ko is endpoint par POST karna hoga:
`http://your-server:5000/webhook/cv`
Header: `X-Webhook-Secret: <WEBHOOK_SECRET from .env>`
Body (JSON): `{"sender_email": "...", "filename": "cv.pdf", "file_base64": "..."}`

## How inbox checking works
- App start hote hi ek background scheduler chalu ho jata hai jo **har 2 minute** baad
  IMAP inbox aur Gmail API (agar enabled) dono check karta hai.
- Naya CV milte hi: text parse hota hai → active job requirements se AI match hota hai →
  score ke hisab se Selected/Rejected decide hota hai → candidate ko **English mein** email
  chali jati hai → sab kuch dashboard mein dikhta hai.

## Notes
- Website title/subtitle hamesha English mein fixed hain.
- Sidebar se Urdu / Pashto / English switch ho sakta hai.
- Yeh ek working prototype hai — production mein deploy karne se pehle proper hosting,
  HTTPS, aur stronger authentication (password hashing already hai, lekin rate-limiting
  waghera add karna chahiye) zaroor lagayein.
