**## Cold Email Automation System ------------------------------------**

This is a Python-based system for sending highly personalised cold emails at scale, designed for students and early-career professionals applying to internships, analyst roles, or research positions.

This tool combines:

- A curated list of London-based finance firms
- Website scraping
- GPT-powered personalisation
- Automated email sending via SMTP

The result is tailored, company-specific outreach instead of generic mass emails.

**### Features**

1. Reads company data from CSV (company name, website, email)

2. Scrapes company websites to understand what they do

3. Uses OpenAI GPT to generate:
    - Company-specific opening paragraphs
    - Custom subject lines

4. Automatically attaches your CV / resume

5. Sends emails via Gmail or Outlook SMTP

6. Dry-run mode to preview emails before sending

7. Logs all send results to a CSV file

**### Prerequisites**

- Python 3.9+
- A Gmail or Outlook email account
- An OpenAI API key (with paid credits)
- Basic command line knowledge

**### Installation guide**

**#### 1. Clone the repository**

```bash
git clone https://github.com/<your-username>/cold-email-automation-system.git
cd cold-email-automation-system
```

**#### 2. Install the requirements**

```bash
pip install -r requirements.txt
```

**#### 3. Environment setup**

This project uses environment variables for security.
You must create a `.env` file locally.

1. First create `.env` :

```bash
cp env.example .env
```

2. Open `.env` and fill in your details : 

```json
# SMTP settings (Gmail example)
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=your_email@gmail.com
SMTP_PASS=your_app_password

# OpenAI
OPENAI_API_KEY=sk-xxxxxxxxxxxxxxxxxxxxxxxxxxxx
```

**#### 4. Email provider setup**

Gmail (recommended)

1. Enable 2-Step Verification
2. Create an App Password
3. Use that password as `SMTP_PASS`

**#### 5. Information about the input data**

The `London_Finance_Firms.csv` file contains a large set of companies with their contact emails. This csv file will be used with `bulk_mailer.py` to automatically create an email body using your OpenAI key, while wiring in your resume and contact details in their repsective locations in the email.

You may:

- Edit this file
- Add your own companies
- Replace it entirely with your own dataset

Please ensure this outreach tool is used efectively and with repsect to the companies on the csv provided. Improper use can lead to companies believing your emails are spam and we want to minimise any concerns in terms of bothering the companies repsective contact emails with irrelevant details. 

If you edit the csv file, only include publicly available contact emails.

**#### 6. Running the script**

A test run of the script **(highly recommended)** can be run with this command. It generates email review without sending anything by creating a file on your local workspace called `outbox_preview` which you can open and check the contents of the generated emails from : 

**THE TEST RUN**

```bash
python bulk_mailer.py \
  --excel London_Finance_Firms.csv \
  --from_name "Your Name" \
  --from_email "youremail@gmail.com" \
  --role "AI / Software Engineer Intern" \
  --resume "your_resume.pdf" \
  --dry_run
```

**THE LIVE RUN**

```bash
python bulk_mailer.py \
  --excel London_Finance_Firms.csv \
  --from_name "Your Name" \
  --from_email "youremail@gmail.com" \
  --role "your role" \
  --resume "your_resume.pdf"
```
This command allows for the emails to actually be sent. 
The command can also be modified with a `max` parameter that limits how many emails in the csv are sent. 

For example : 

```json
--max 10 #add this at the bottom of the command above 
```

As well as `outbox_preview`, all sends are logged to `send_log.csv` which shows the :

- Timestamp
- Company
- Email
- Subject
- Result (SENT or ERROR)

**### Further information**

**Included Company List Disclaimer**

This repository includes a curated list of London-based finance and investment firms
containing publicly available contact email addresses (e.g. info@, hello@).
This list is provided for educational and personal outreach purposes only.

Users are responsible for:
Complying with local regulations (GDPR, PECR, CAN-SPAM)
Sending respectful, relevant communications
Avoiding spam or excessive automation
The author does not endorse unsolicited mass emailing.

**Ethical Use**

This tool is designed for:
Internship applications
Research outreach
Thoughtful professional networking
It is not designed for spam, marketing blasts, or abuse.
Use responsibly.

**Troubleshooting**

Emails not appearing in “Sent”: 

Check `send_log.csv`
SMTP delivery ≠ mailbox “Sent” folder
BCC yourself if needed

**Authentication errors**

Ensure App Password is used
Confirm SMTP host/port
Verify `.env` is loaded correctly

**Future Improvements**

HTML email templates
Multi-language support
Reply tracking
CRM-style dashboard

**Author**

Hissan Omar
MSci Artificial Intelligence — King’s College London
GitHub: https://github.com/Hissan7
LinkedIn: https://www.linkedin.com/in/hissanomar