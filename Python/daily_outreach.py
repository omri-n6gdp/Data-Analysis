"""
=============================================================
  DAILY OUTREACH PIPELINE
  Run this from VS Code each morning (Tue/Wed/Thu only).
=============================================================
"""

import time
import smtplib
import logging
import requests
import anthropic
import openpyxl
from datetime import datetime, timedelta
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from bs4 import BeautifulSoup

# ═══════════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════════

from secrets_local import GMAIL_ADDRESS  # kept in secrets_local.py (not uploaded)
from secrets_local import GMAIL_APP_PASSWORD  # real value lives in secrets_local.py (not uploaded)
from secrets_local import ANTHROPIC_API_KEY  # real value lives in secrets_local.py (not uploaded)
EXCEL_PATH         = "/Users/omrishahar/Documents/Data Analysis/leads_master.xlsx"

EMAILS_PER_NICHE   = 5
DELAY_SECONDS      = 10

ACTIVE_NICHES = [
    "Gym",
    "Estate Agent",
    "Accountant",
    "Dentist",
    "Physio",
    "Corporate Events",
    "Letting Agent",
    "Tutor",
]

# ═══════════════════════════════════════════════════════════
#  NICHE CONFIG
# ═══════════════════════════════════════════════════════════

NICHE_CONFIG = {
    "Gym": {
        "angle": "class scheduling, membership retention, and capacity utilisation",
        "data_examples": "booking records, attendance logs, or membership history",
    },
    "Estate Agent": {
        "angle": "listing performance, days on market, and price achievement across their portfolio",
        "data_examples": "listing records, sale prices, or days on market history",
    },
    "Accountant": {
        "angle": "client data, outcome evidence, and business development patterns",
        "data_examples": "client records, engagement history, or service outcomes",
    },
    "Dentist": {
        "angle": "appointment patterns, no-show rates, and patient retention",
        "data_examples": "appointment records, booking times, or treatment history",
    },
    "Physio": {
        "angle": "patient retention, treatment completion rates, and appointment patterns",
        "data_examples": "appointment records, treatment plans, or patient return history",
    },
    "Corporate Events": {
        "angle": "seasonal demand patterns, lead-to-booking conversion, and pricing optimisation",
        "data_examples": "booking records, event types, or seasonal revenue history",
    },
    "Letting Agent": {
        "angle": "void periods, re-let times, and portfolio yield patterns",
        "data_examples": "tenancy records, void period logs, or rental yield history",
    },
    "Tutor": {
        "angle": "student retention, subject performance, and enquiry conversion rates",
        "data_examples": "session records, subject history, or student intake logs",
    },
}

# ═══════════════════════════════════════════════════════════
#  PROMPTS & TEMPLATE
# ═══════════════════════════════════════════════════════════

SYSTEM_PROMPT = """\
You write the middle paragraph of a cold outreach email. The email is written in first person — you are Omri, a Physics student at Imperial College London. Write as if you are Omri writing directly to the recipient. Do not refer to Omri in third person.

You only write the middle paragraph — 3 sentences maximum.

Rules:
- Sentence 1: Start with "I noticed on your website that..." and state one specific observation. Just describe what you saw, do not editorialise or imply the business has a problem.
- Sentence 2: Say what many businesses in this niche find difficult to see in their data — keep it factual and plain.
- Sentence 3: Say what insights I could give — simple and specific, no lists or bullet points.
- Write in first person throughout — never say "Omri" at any point.
- Sound human and fairly casual — not stiff, not corporate, not AI-sounding.
- No em dashes, no lists, no bullet points.
- Never use: "juggling", "complexity", "delve", "leverage", "unlock", "empower", "streamline", "quick data thing", "freelance data analysis".
- Do not include a subject line, greeting, or sign-off — only the middle paragraph.

GOOD EXAMPLES (write like these):

Example 1 — Gym:
"I noticed on your website that you run classes across three locations with different studio sizes. Many pilates and fitness studios find it difficult to see clearly which locations and time slots are underperforming relative to their capacity, and whether cancellation patterns differ by site. I could give insights into exactly where the gaps are and how to get more out of your existing schedule."

Example 2 — Dentist:
"I noticed on your website that you operate a same-day emergency service around the clock. Many emergency dental practices find it difficult to see clearly which times of day and days of the week drive the most demand, making staffing and capacity decisions harder than they need to be. I could give insights into exactly where those demand patterns sit and how to plan around them."

Example 3 — Estate Agent:
"I noticed on your website that you have active stock across North and East Finchley, including some properties that appear to have been on the market longer than the local average. Many independent agents find it difficult to see clearly whether that pattern is driven by pricing, listing quality, or timing, and what to adjust first. I could give insights into exactly what the data shows."

Example 4 — Letting Agent:
"I noticed on your website that you manage lettings across Plaistow, Barking and Stratford from two offices. Many letting agents find it difficult to see clearly how void periods and re-let times differ across their portfolio, and whether certain property types or locations are consistently underperforming. I could give insights into exactly where those patterns sit across your book."

BAD EXAMPLE (never write like this):
"I noticed on your website that you're running classes across multiple locations with different age groups, which means you're juggling a lot of scheduling complexity. Most gyms with this setup struggle to fill classes consistently and don't have a clear picture of which time slots and locations are actually driving member retention. Omri could help you analyse your booking patterns and attendance data to show which classes have the best member retention rates, which time slots consistently underperform, and where you could adjust your timetable to get better capacity utilisation across all your locations."

Why the bad example fails: it editorialises, refers to Omri in third person, uses a list structure, and adds unnecessary commentary. The good examples are plain, first person, and read like a human wrote them.
"""

TEMPLATE = """\
Hi,

I'm a Physics student at Imperial with experience handling large datasets and turning them into useful information.

{middle}

If you have any unused data (like {data_examples}) that you're not sure what to do with, feel free to reach out and I'd be happy to share an example of what I could extract, with no upfront commitment. My analysis isn't limited to the above either — whatever you feel would be most useful, I'm happy to take a look.

Thanks,
Omri Shahar
Physics, Imperial College London\
"""

# ═══════════════════════════════════════════════════════════
#  HELPERS
# ═══════════════════════════════════════════════════════════

def next_valid_send_day(base_date):
    """Return base_date + 7 days. Since you only send Tue/Wed/Thu,
    7 days later is always the same valid weekday."""
    return base_date + timedelta(days=7)

# ═══════════════════════════════════════════════════════════
#  LOGGING
# ═══════════════════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.FileHandler("outreach_log.txt"),
        logging.StreamHandler()
    ]
)
log = logging.getLogger()

# ═══════════════════════════════════════════════════════════
#  STEP 1 — Read leads from Excel
# ═══════════════════════════════════════════════════════════

def get_leads(excel_path, niche, n):
    wb = openpyxl.load_workbook(excel_path)
    if niche not in wb.sheetnames:
        log.warning(f"Sheet '{niche}' not found — skipping")
        return []

    ws = wb[niche]
    headers = [cell.value for cell in ws[1]]

    try:
        col = {h: i for i, h in enumerate(headers)}
        name_i    = col["Business Name"]
        email_i   = col["Email"]
        website_i = col["Website"]
        status_i  = col["Status"]
    except KeyError as e:
        log.error(f"Missing column in '{niche}': {e}")
        return []

    leads = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        if len(leads) >= n:
            break
        status = row[status_i].value
        email  = row[email_i].value
        if status == "To Contact" and email and email != "NOT FOUND":
            leads.append({
                "row":     row[0].row,
                "name":    row[name_i].value,
                "email":   email,
                "website": row[website_i].value,
                "niche":   niche,
            })

    wb.close()
    return leads

# ═══════════════════════════════════════════════════════════
#  STEP 2 — Fetch website
# ═══════════════════════════════════════════════════════════

def fetch_website(url):
    if not url or url == "NOT FOUND":
        return ""
    try:
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
        soup = BeautifulSoup(r.text, "html.parser")
        for tag in soup(["script", "style", "nav", "footer", "header"]):
            tag.decompose()
        return soup.get_text(separator=" ", strip=True)[:1500]
    except Exception as e:
        log.warning(f"Could not fetch {url}: {e}")
        return ""

# ═══════════════════════════════════════════════════════════
#  STEP 3 — Generate middle paragraph
# ═══════════════════════════════════════════════════════════

def generate_middle(lead, website_text, niche_config):
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    message = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=300,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": f"""Business name: {lead['name']}
Niche: {lead['niche']}
Key angles to consider: {niche_config['angle']}
Website content:
{website_text if website_text else '[Website could not be fetched — use general niche knowledge]'}

Write the middle paragraph now."""}]
    )
    return message.content[0].text.strip()

# ═══════════════════════════════════════════════════════════
#  STEP 4 — Assemble email
# ═══════════════════════════════════════════════════════════

def assemble_email(lead, middle, niche_config):
    body = TEMPLATE.format(
        middle=middle,
        data_examples=niche_config["data_examples"],
    )
    angles = {
        "Gym":              "class and membership data",
        "Estate Agent":     "listings data",
        "Accountant":       "client data",
        "Dentist":          "appointment patterns",
        "Physio":           "patient data",
        "Corporate Events": "booking patterns",
        "Letting Agent":    "portfolio data",
        "Tutor":            "student data",
    }
    subject = f"Something I noticed about your {angles.get(lead['niche'], 'data')}"
    return subject, body

# ═══════════════════════════════════════════════════════════
#  STEP 5 — Send
# ═══════════════════════════════════════════════════════════

def send_email(to_address, subject, body, smtp_server):
    msg = MIMEMultipart()
    msg["From"]    = GMAIL_ADDRESS
    msg["To"]      = to_address
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))
    smtp_server.sendmail(GMAIL_ADDRESS, to_address, msg.as_string())

# ═══════════════════════════════════════════════════════════
#  STEP 6 — Update Excel
# ═══════════════════════════════════════════════════════════

def update_excel(excel_path, niche, row_num, subject, body):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb[niche]
    headers = [cell.value for cell in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers)}

    today      = datetime.now()
    followup   = next_valid_send_day(today)

    ws.cell(row=row_num, column=col["Status"]).value          = "Sent"
    ws.cell(row=row_num, column=col["Email Sent Date"]).value = today.strftime("%Y-%m-%d")
    ws.cell(row=row_num, column=col["Follow Up Date"]).value  = followup.strftime("%Y-%m-%d")

    if "Notes" in col:
        ws.cell(row=row_num, column=col["Notes"]).value = f"Subject: {subject}\n\n{body}"

    wb.save(excel_path)
    wb.close()

# ═══════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════

def main():
    today_str = datetime.now().strftime("%Y-%m-%d")
    log.info(f"========== DAILY OUTREACH — {today_str} ==========")

    all_leads = []
    for niche in ACTIVE_NICHES:
        leads = get_leads(EXCEL_PATH, niche, EMAILS_PER_NICHE)
        log.info(f"{niche:20s} — {len(leads)} leads queued")
        all_leads.extend(leads)

    if not all_leads:
        log.info("No leads to contact today. Exiting.")
        return

    total = len(all_leads)
    log.info(f"\nTotal emails to send: {total}")
    log.info(f"Estimated time: {total * DELAY_SECONDS / 60:.1f} minutes\n")

    sent   = 0
    failed = 0

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        log.info("Gmail connected.\n")

        for i, lead in enumerate(all_leads, 1):
            niche_config = NICHE_CONFIG[lead["niche"]]
            try:
                website_text  = fetch_website(lead["website"])
                middle        = generate_middle(lead, website_text, niche_config)
                subject, body = assemble_email(lead, middle, niche_config)
                send_email(lead["email"], subject, body, server)
                update_excel(EXCEL_PATH, lead["niche"], lead["row"], subject, body)
                sent += 1
                log.info(f"[{i:02d}/{total}] SENT    {lead['niche']:20s}  {lead['name']:35s}  ->  {lead['email']}")
            except Exception as e:
                failed += 1
                log.error(f"[{i:02d}/{total}] FAILED  {lead['name']} — {e}")

            if i < total:
                time.sleep(DELAY_SECONDS)

    log.info(f"\n========== DONE ==========")
    log.info(f"Sent: {sent}  |  Failed: {failed}  |  Log: outreach_log.txt")

if __name__ == "__main__":
    main()
