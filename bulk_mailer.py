#!/usr/bin/env python3
"""
bulk_mailer.py — GPT-personalised finance outreach

What it does
------------
- Reads an Excel/CSV file with columns: Company, Website, Email, (optional) Greeting, First Line, Status
- Scrapes company website (homepage / about) for a text blurb
- Uses OpenAI (if OPENAI_API_KEY set) to:
    * write a 3–5 sentence, company-specific opener that explains what they do & why you're a good fit
    * generate a short "{something}" phrase for the subject line
- Builds subject:
    "MSci Artificial Intelligence student willing to help {company} with {something}"
- Email body includes:
    * personalised opener
    * your resume paragraph
    * your London closing line
    * full signature (degree, phone, email, GitHub, LinkedIn)
- Attaches your resume file (PDF / DOCX)
- Sends via SMTP (Hotmail/Outlook or Gmail) with rate limiting
- Supports --dry_run to preview emails without sending

Usage (example)
---------------
python bulk_mailer.py --excel London_Finance_Firms.csv \
    --from_name "Hissan Omar" \
    --from_email "hissanomar786@hotmail.com" \
    --role "AI/Software Engineer Intern" \
    --resume "resume_H_Omar.pdf" \
    --dry_run

Environment
-----------
SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS
OPENAI_API_KEY  (required for GPT-based personalisation)
MAX_PER_RUN     (optional cap)
"""

import os
import re
import time
import csv
import argparse
import smtplib
import mimetypes
from email.message import EmailMessage
from urllib.parse import urljoin
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests
from bs4 import BeautifulSoup

from dotenv import load_dotenv
env_path = Path(__file__).parent / ".env"
load_dotenv(dotenv_path=env_path)

# Optional: OpenAI for stronger personalization
try:
    from openai import OpenAI
    openai_client = OpenAI()
except Exception:
    openai_client = None


# ---------- CONSTANT CONTENT (your personal branding) ----------
YOUR_DEGREE = "MSci Artificial Intelligence, King’s College London"
YOUR_PHONE = "07525483773"
YOUR_EMAIL = "hissanomar786@hotmail.com"
YOUR_GITHUB = "https://github.com/Hissan7"
YOUR_LINKEDIN = "https://www.linkedin.com/in/hissanomar"

RESUME_PARAGRAPH = (
    "I’m Hissan Omar, currently pursuing an MSci in Artificial Intelligence at King’s College London. "
    "My background spans AI, finance, and data analytics — from developing an Option Value Predictor using "
    "Monte Carlo simulations and the Black-Scholes model to building a Sentiment-Based Alpha Generator that "
    "links social media trends to stock returns."
)

LONDON_CLOSING = (
    "If there’s any way I could take a bit off your plate — whether through research, data processing, "
    "or analytics support — I’d love to contribute and learn from your team in London."
    "I'm open to a short chat, discussing possible tasks within your firm. I hope to hearing back from you soon !"
)
# --------------------------------------------------------------


def normalize_columns(df):
    mapping = {}
    for col in df.columns:
        lower = col.strip().lower()
        if lower in ("company", "company name", "organisation", "organization"):
            mapping[col] = "Company"
        elif lower in ("website", "site", "url", "homepage"):
            mapping[col] = "Website"
        elif lower in ("email", "e-mail", "contact"):
            mapping[col] = "Email"
        elif lower in ("greeting", "salutation"):
            mapping[col] = "Greeting"
        elif lower in ("first line", "first_line", "opener", "custom line"):
            mapping[col] = "First Line"
        elif lower in ("status",):
            mapping[col] = "Status"
        else:
            mapping[col] = col
    return df.rename(columns=mapping)


# ---- Robust Excel/CSV loader (handles messy XLSX styles) ----
def load_leads(path, sheet=None):
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception as e1:
        try:
            return pd.read_excel(path, sheet_name=sheet, engine="calamine")
        except Exception as e2:
            if str(path).lower().endswith(".csv"):
                return pd.read_csv(path)
            raise SystemExit(
                "Failed to read the Excel file.\n"
                f"openpyxl error: {e1}\n"
                f"calamine error: {e2}\n"
                "Fix by either: (1) pip install calamine, (2) upgrade openpyxl, "
                "(3) re-save the workbook, or (4) export to CSV and rerun with that file."
            )


def fetch_site_blurb(url, timeout=12):
    """Fetch short description from homepage or /about pages."""
    if not url or not isinstance(url, str):
        return ""
    if not url.startswith(("http://", "https://")):
        url = "http://" + url
    candidates = [url, urljoin(url, "/about"), urljoin(url, "/about-us"), urljoin(url, "/company")]
    texts = []
    for u in candidates:
        try:
            r = requests.get(u, timeout=timeout, headers={"User-Agent": "Mozilla/5.0"})
            if r.status_code != 200 or "text/html" not in r.headers.get("Content-Type", ""):
                continue
            soup = BeautifulSoup(r.text, "html.parser")
            md = soup.find("meta", attrs={"name": "description"})
            if md and md.get("content"):
                texts.append(md["content"].strip())
            # first p
            p = soup.find("p")
            if p:
                texts.append(p.get_text(" ", strip=True))
            joined = " ".join(t.strip() for t in texts if t and len(t.strip()) > 40)
            if joined:
                return re.sub(r"\s+", " ", joined)[:1200]
        except Exception:
            continue
    return ""


def extract_helper_phrase(text):
    """
    Create a short 'something' phrase for the subject line, e.g., 'sustainable finance', 'credit analytics'.
    Heuristic: grab clean noun-ish phrases, 2–4 words if possible.
    """
    if not text:
        return "analytics support"
    txt = re.sub(r"\s+", " ", text)

    candidates = []

    # Prepositions: 'in <phrase>', 'on <phrase>'
    for m in re.finditer(r"\b(in|on|around|for|within)\s+([A-Za-z][A-Za-z\- ]{3,50})", txt):
        phrase = m.group(2).strip().lower()
        phrase = re.sub(r"[^a-z\- ]+", "", phrase)
        if 3 <= len(phrase) <= 40:
            candidates.append(phrase)

    # Capitalised chunks: 'Inclusive Finance', 'Growth Equity'
    caps = re.findall(r"\b([A-Z][a-z]+(?:\s+[A-Z][a-z]+){0,3})\b", text)
    for c in caps:
        candidates.append(c.lower())

    # Keyword fallbacks
    keywords = [
        "sustainable finance", "inclusive finance", "credit analytics", "portfolio risk",
        "quant research", "deal sourcing", "market intelligence", "growth analytics",
        "data engineering", "automation", "fund operations", "fintech infrastructure"
    ]
    candidates.extend(keywords)

    cleaned = []
    for c in candidates:
        c = re.sub(r"\s+", " ", c.strip())
        if 3 <= len(c) <= 40:
            cleaned.append(c)

    if cleaned:
        cleaned.sort(key=lambda s: (abs(len(s.split()) - 3), len(s)))
        return cleaned[0]

    return "analytics support"


def gpt_company_personalization(company, website, blurb, your_role):
    """
    Use GPT to generate:
      - opener: 3–5 sentence paragraph (what they do + why you're relevant)
      - helper_phrase: short phrase for subject ({something})
    Falls back to heuristic if no API.
    """
    api_key = os.getenv("OPENAI_API_KEY")

    if not api_key or openai_client is None:
        print(f"[NO GPT] Using fallback opener for {company}")
        opener = fallback_opener(company, blurb, your_role)
        helper = extract_helper_phrase(blurb or opener)
        return opener, helper

    try:
        # print(f"[GPT] Generating personalized opener for {company}...")
        print(f"Creating personalized opener for {company}...")
        prompt = (
            "You are helping an MSci AI student cold-email finance/VC/asset-management firms.\n\n"
            "Given the company description and candidate profile, write:\n"
            "1) A 3–5 sentence paragraph that:\n"
            "   - briefly explains what the company does (using the blurb/website info),\n"
            "   - mentions 1–2 specific aspects of their focus (e.g., sector, geography, strategy),\n"
            "   - and clearly states why an AI/quant-focused intern is a good fit.\n"
            "2) A short 2–4 word phrase suitable for 'helping with {something}' in an email subject line.\n\n"
            "Output STRICTLY in this format:\n"
            "OPENER: <paragraph>\n"
            "PHRASE: <short phrase>\n\n"
            f"Company: {company}\n"
            f"Website: {website}\n"
            f"Blurb: {blurb}\n"
            f"Candidate role: {your_role} with AI/quant + data analytics background."
        )

        resp = openai_client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {"role": "system", "content": "You write concise, concrete, non-fluffy outreach copy."},
                {"role": "user", "content": prompt},
            ],
            temperature=0.6,
            max_tokens=300,
        )
        text = resp.choices[0].message.content.strip()

        opener_match = re.search(r"OPENER:\s*(.+?)\s*PHRASE:", text, flags=re.S | re.I)
        phrase_match = re.search(r"PHRASE:\s*(.+)", text, flags=re.S | re.I)

        if opener_match:
            opener = opener_match.group(1).strip()
        else:
            opener = fallback_opener(company, blurb, your_role)

        if phrase_match:
            helper = phrase_match.group(1).strip()
        else:
            helper = extract_helper_phrase(blurb or opener)

        helper = helper.strip().strip(".").lower()
        if not helper:
            helper = extract_helper_phrase(blurb or opener)

        return opener, helper

    except Exception as e:
        print(f"[GPT ERROR for {company}]: {e}")
        opener = fallback_opener(company, blurb, your_role)
        helper = extract_helper_phrase(blurb or opener)
        return opener, helper



def fallback_opener(company, blurb, your_role):
    if blurb:
        key = extract_helper_phrase(blurb)
        return (
            f"I’ve been looking into {company} and I really like your focus on {key}. "
            f"Given my background in AI-driven financial modelling and data analytics, "
            f"I’d love to support work in this area as a {your_role}."
        )
    return (
        f"I’ve been exploring the work that {company} does and I’m excited about your approach to building in finance. "
        f"With my background in AI and quantitative analysis, I’d love to support your team as a {your_role}."
    )


def subject_line(company, helper_phrase):
    company_simple = company.strip()
    phrase = helper_phrase.strip().lower()
    return f"Msci AI student willing to help {company_simple} with {phrase}"


def build_email_body(company, greeting, opener, website, your_name):
    if not greeting:
        greeting = f"Hello {company} team. I hope this message finds you in good health,"

    # Prefix the opener with your custom line
    if opener:
        # make the first letter of opener lowercase so it flows nicely:
        opener_text = opener[0].lower() + opener[1:]
        opener_block = f"I'm fascinated by how {opener_text}"
    else:
        opener_block = "I'm fascinated by how your work fits into the broader finance and technology landscape."

    body = f"""{greeting}

{opener_block}

{RESUME_PARAGRAPH}

{LONDON_CLOSING}

Best,
{your_name}
{YOUR_DEGREE}
Phone: {YOUR_PHONE}
Email: {YOUR_EMAIL}
GitHub: {YOUR_GITHUB}
LinkedIn: {YOUR_LINKEDIN}
Website noted: {website if website else "-"}
"""
    return body



def attach_files(msg: EmailMessage, attachments):
    for file_path in attachments or []:
        p = Path(file_path)
        if not p.exists():
            continue
        ctype, encoding = mimetypes.guess_type(str(p))
        if ctype is None or encoding is not None:
            ctype = "application/octet-stream"
        maintype, subtype = ctype.split("/", 1)
        with open(p, "rb") as f:
            msg.add_attachment(f.read(), maintype=maintype, subtype=subtype, filename=p.name)


def send_email(
    smtp_host,
    smtp_port,
    smtp_user,
    smtp_pass,
    sender,
    recipient,
    subject,
    body,
    cc=None,
    bcc=None,
    reply_to=None,
    attachments=None,
):
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = recipient
    if cc:
        msg["Cc"] = ", ".join(cc) if isinstance(cc, (list, tuple)) else cc
    if reply_to:
        msg["Reply-To"] = reply_to
    msg["Subject"] = subject
    msg.set_content(body)

    attach_files(msg, attachments)

    use_starttls = (str(smtp_port) == "587" or smtp_host.lower() in {"smtp.office365.com", "smtp-mail.outlook.com"})
    if use_starttls:
        with smtplib.SMTP(smtp_host, int(smtp_port)) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(smtp_user, smtp_pass)
            server.send_message(msg)
    else:
        with smtplib.SMTP_SSL(smtp_host, int(smtp_port)) as server:
            server.login(smtp_user, smtp_pass)
            server.send_message(msg)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--excel", required=True, help="Path to the Excel/CSV file of leads")
    parser.add_argument("--sheet", default=None, help="Sheet name (optional for Excel)")
    parser.add_argument("--from_name", required=True, help="Your name")
    parser.add_argument("--from_email", required=True, help="Email address to send from")
    parser.add_argument("--role", default="AI/Software Engineer Intern", help="Role you're pitching")
    parser.add_argument("--resume", default=None, help="Path to your resume file to attach (PDF/DOCX)")
    parser.add_argument("--rate_sec", type=float, default=20, help="Seconds to sleep between sends")
    parser.add_argument("--max", type=int, default=int(os.getenv("MAX_PER_RUN", "50")), help="Max emails to process this run")
    parser.add_argument("--dry_run", action="store_true", help="Do everything except actually send")
    parser.add_argument("--outbox", default="outbox_preview", help="Folder to save preview emails when dry_run")
    args = parser.parse_args()

    # SMTP creds
    smtp_host = os.getenv("SMTP_HOST", "smtp.gmail.com")
    smtp_port = os.getenv("SMTP_PORT", "465")
    smtp_user = os.getenv("SMTP_USER", args.from_email)
    smtp_pass = os.getenv("SMTP_PASS")
    if not args.dry_run and not smtp_pass:
        raise SystemExit("Missing SMTP_PASS env var. For Gmail, create an App Password and set SMTP_PASS.")

    Path(args.outbox).mkdir(parents=True, exist_ok=True)

    df = load_leads(args.excel, sheet=args.sheet)
    df = normalize_columns(df)

    required = ["Company", "Email"]
    for col in required:
        if col not in df.columns:
            raise SystemExit(f"Excel/CSV missing required column: {col}")

    sent_log_path = Path("send_log.csv")
    sent_log = []
    processed = 0

    for i, row in df.iterrows():
        if processed >= args.max:
            break

        company = str(row.get("Company", "")).strip()
        email = str(row.get("Email", "")).strip()
        website = str(row.get("Website", "")).strip()
        greeting = ""  # always force our own greeting

        # Handle opener_from_sheet safely (treat NaN as empty)
        if "First Line" in df.columns:
            raw_opener = row.get("First Line", "")
            if pd.isna(raw_opener):
                opener_from_sheet = ""
            else:
                opener_from_sheet = str(raw_opener).strip()
        else:
            opener_from_sheet = ""

        status = str(row.get("Status", "")).strip() if "Status" in df.columns else ""

        if not company or not email:
            continue
        if status.lower() in {"sent", "done", "bounced"}:
            continue

        blurb = ""

        if opener_from_sheet:
            print(f"[SHEET] Using opener from sheet for {company}")
            opener = opener_from_sheet
            helper_phrase = extract_helper_phrase(opener_from_sheet)
        else:
            # print(f"[SCRAPE+GPT] Fetching blurb and using GPT for {company}")
            print(f"{i}. Email has been created for {company}. Need to review and then send")

            blurb = fetch_site_blurb(website)
            opener, helper_phrase = gpt_company_personalization(company, website, blurb, args.role)

        # Hard safety: prevent 'nan' issues
        if opener is None or str(opener).lower() == "nan":
            print(f"[SAFETY] 'nan' opener for {company}, falling back.")
            opener = fallback_opener(company, blurb, args.role)

        subj = subject_line(company, helper_phrase)
        body = build_email_body(
            company=company,
            greeting=greeting,
            opener=opener,
            website=website,
            your_name=args.from_name,
        )

        timestamp = datetime.utcnow().isoformat(timespec="seconds") + "Z"

        if args.dry_run:
            safe_company = re.sub(r"[^A-Za-z0-9_.-]+", "_", company)[:60]
            filename = Path(args.outbox) / f"{safe_company}__{email}.txt"
            with open(filename, "w", encoding="utf-8") as f:
                f.write(f"TO: {email}\nSUBJECT: {subj}\n\n{body}")
            result = "PREVIEWED"
        else:
            try:
                send_email(
                    smtp_host=smtp_host,
                    smtp_port=smtp_port,
                    smtp_user=smtp_user,
                    smtp_pass=smtp_pass,
                    sender=f"{args.from_name} <{args.from_email}>",
                    recipient=email,
                    subject=subj,
                    body=body,
                    attachments=[args.resume] if args.resume else None,
                )
                result = "SENT"
                time.sleep(args.rate_sec)
            except Exception as e:
                result = f"ERROR: {e}"

        sent_log.append({
            "timestamp": timestamp,
            "company": company,
            "email": email,
            "website": website,
            "greeting": greeting,
            "first_line": opener,
            "subject": subj,
            "result": result,
        })
        processed += 1

    fieldnames = ["timestamp", "company", "email", "website", "greeting", "first_line", "subject", "result"]
    write_header = not sent_log_path.exists()
    with open(sent_log_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        if write_header:
            writer.writeheader()
        writer.writerows(sent_log)

    print(f"Processed {processed} rows. Log -> {sent_log_path.absolute()}")
    if args.dry_run:
        print(f"Preview emails saved in folder: {Path(args.outbox).absolute()}")
        print("When ready, remove --dry_run to actually send.")


if __name__ == "__main__":
    main()
