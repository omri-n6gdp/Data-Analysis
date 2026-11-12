"""
=============================================================
  OUTREACH PIPELINE — daily sends + follow-ups + test mode
  Run this from VS Code each morning (Tue/Wed/Thu only).

  On launch you will be asked:
    [1] Test mode  — generates one email for a random uncontacted
                     business and sends it to an address you choose.
                     Nothing is written back to the Excel file.
    [2] Send mode  — runs the full daily outreach (5 per niche).
    [3] Follow-up  — sends follow-ups to every lead whose
                     Follow Up Date is today.
=============================================================
"""

import random
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

EMAILS_PER_NICHE   = 10
DELAY_SECONDS      = 0.1

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
        "intro":         "I work on short data analysis projects for gyms and fitness studios.",
        "data_examples": "booking records, attendance logs, membership history",
        "bullets": [
            "Which classes actually improve member retention",
            "What times or instructors drive the highest return visits",
            "Where demand exists but is not being captured",
        ],
        "subject": "Quick question about your membership data",
        "problems": (
            "common problems: classes that look busy but have high drop-off after the first visit, "
            "peak slots that are overbooked while off-peak slots stay empty, "
            "membership cancellations that cluster after a specific number of weeks, "
            "no visibility on which instructor or class type drives the most retention."
        ),
    },
    "Estate Agent": {
        "intro":         "I work on short data analysis projects for estate agents.",
        "data_examples": "listing records, sale prices, days on market history",
        "bullets": [
            "Which property types or areas are taking longest to sell",
            "Where your pricing is consistently above or below final sale price",
            "What seasonal patterns exist in your listings data",
        ],
        "subject": "Quick question about your listings data",
        "problems": (
            "common problems: properties sitting on the market without a clear pattern for why, "
            "no easy way to see whether overpricing or listing quality is driving slow sales, "
            "seasonal fluctuations that are felt but not measured, "
            "no visibility on which valuer or area consistently achieves asking price."
        ),
    },
    "Accountant": {
        "intro":         "I work on short data analysis projects for accountancy practices.",
        "data_examples": "client records, engagement history, invoicing data",
        "bullets": [
            "Which client segments generate the most repeat work",
            "Where clients tend to stop engaging and what usually comes before it",
            "What your data shows about service uptake over time",
        ],
        "subject": "Quick question about your client data",
        "problems": (
            "common problems: no clear picture of which services or client types are most profitable, "
            "clients who quietly stop engaging with no early warning signs, "
            "upsell opportunities that are visible in hindsight but missed in real time, "
            "no way to see which acquisition channel produces the longest-lasting clients."
        ),
    },
    "Dentist": {
        "intro":         "I work on short data analysis projects for dental practices.",
        "data_examples": "appointment records, booking history, treatment logs",
        "bullets": [
            "Which appointment types have the highest no-show rates",
            "What times and days drive the most demand",
            "Where patient retention drops off after the first visit",
        ],
        "subject": "Quick question about your appointment data",
        "problems": (
            "common problems: no-shows that vary by appointment type but look random in aggregate, "
            "recall rates that are lower than expected with no clear cause, "
            "peak demand that is guessed rather than measured, "
            "no visibility on which treatment pathways result in patients not returning."
        ),
    },
    "Physio": {
        "intro":         "I work on short data analysis projects for physiotherapy and allied health clinics.",
        "data_examples": "appointment records, treatment plans, patient return history",
        "bullets": [
            "Which treatment types have the best completion rates",
            "Where patients drop out before finishing a course of treatment",
            "What your appointment data shows about peak demand periods",
        ],
        "subject": "Quick question about your patient data",
        "problems": (
            "common problems: patients dropping out mid-course with no pattern identified, "
            "no visibility on which referral source produces the most completed treatments, "
            "appointment utilisation that varies across practitioners without explanation, "
            "no way to see whether certain conditions or treatment types have worse retention."
        ),
    },
    "Corporate Events": {
        "intro":         "I work on short data analysis projects for corporate events companies.",
        "data_examples": "booking records, client history, event type logs",
        "bullets": [
            "Which event types or clients generate the most repeat bookings",
            "Where in the year demand concentrates and where it drops",
            "What your enquiry data shows about conversion patterns",
        ],
        "subject": "Quick question about your bookings data",
        "problems": (
            "common problems: repeat client value that is felt but not quantified, "
            "no clear view of which event types convert enquiries most reliably, "
            "seasonal gaps that are anticipated but not planned around with data, "
            "no visibility on whether certain lead sources produce higher-value clients."
        ),
    },
    "Letting Agent": {
        "intro":         "I work on short data analysis projects for letting agents.",
        "data_examples": "tenancy records, void period logs, rental yield history",
        "bullets": [
            "Which properties or areas have the longest void periods",
            "Where re-let times are consistently running above your average",
            "What your portfolio data shows about yield by property type or location",
        ],
        "subject": "Quick question about your portfolio data",
        "problems": (
            "common problems: void periods that vary across the portfolio without an obvious reason, "
            "no easy way to see which property types or locations consistently underperform on yield, "
            "re-let times that feel longer than they should be but are hard to benchmark internally, "
            "no visibility on whether certain landlords or areas generate disproportionate management overhead."
        ),
    },
    "Tutor": {
        "intro":         "I work on short data analysis projects for tutoring businesses.",
        "data_examples": "session records, enquiry logs, student intake history",
        "bullets": [
            "Which subjects or year groups have the best student retention",
            "Where enquiries drop off before converting to bookings",
            "What your session data shows about long-term student outcomes",
        ],
        "subject": "Quick question about your student data",
        "problems": (
            "common problems: students who stop booking without any warning sign, "
            "enquiry conversion that varies by subject or tutor but looks random overall, "
            "no visibility on which students are at risk of dropping off before they do, "
            "no easy way to see which subjects or tutors produce the strongest retention."
        ),
    },
}

# ═══════════════════════════════════════════════════════════
#  PROMPT & TEMPLATE
# ═══════════════════════════════════════════════════════════

SYSTEM_PROMPT = """\
You write one short personalised sentence for a cold outreach email. The email is from Omri, a Physics student at Imperial College London. Write in first person. Never say "Omri".

The sentence connects something specific from the business's website to a real, common data problem in their niche. It is the only personalised part of the email. Everything else is fixed.

The sentence must follow this exact structure:
"I noticed [specific thing from their website], which usually means [common data problem this causes]."

Rules:
- The first half must name something real and specific from the website — a service, a feature, a scale detail, an operational fact. Not a vague description. Not a compliment.
- The second half must name the data problem that specific thing tends to cause. It should feel like something the recipient would recognise as true, not a generic claim.
- The problem must be something that data analysis can actually help with.
- One sentence only. No follow-on sentences.
- No em dashes. Not a single one. Use a comma before "which" instead.
- No lists, no bullet points, no bold text.
- Do not start with "I noticed on your website" — start with "I noticed" followed immediately by the specific thing.
- Sound plain and direct. Not corporate, not AI-sounding, not flattering.
- Never use: "juggling", "complexity", "delve", "leverage", "unlock", "empower", "streamline", "valuable", "actionable", "game-changing", "churn".
- If the website gives nothing useful, infer something plausible from the business name and niche. Do not mention the website was unavailable.

GOOD EXAMPLES:

Gym with multiple class types:
"I noticed you offer a mix of HIIT, yoga, and pilates sessions, which usually makes it hard to see clearly which formats are driving membership renewals versus one-off bookings."

Letting agent managing a large portfolio:
"I noticed you manage over 300 properties across several London boroughs, which typically makes it difficult to spot which areas or property types are consistently running above average void periods."

Dental practice with an online booking system:
"I noticed you take bookings directly through your website across several treatment types, which tends to create no-show patterns that vary by appointment type but look random when you look at the total."

Physio clinic with multiple practitioners:
"I noticed you have six practitioners listed across two specialisms, which usually makes it hard to see whether drop-out rates mid-course differ by practitioner or by condition type."

Estate agent covering a wide patch:
"I noticed you cover sales and lettings across five different postcodes, which tends to make it difficult to see which areas are consistently achieving asking price and which are not."

BAD EXAMPLES — never write like these:

"I noticed on your website that you run classes across three locations with different studio sizes." (purely observational, no problem stated)
"I noticed you have a great range of services, which shows real breadth." (vague and flattering)
"I noticed you could benefit from better data analysis." (presumptuous, no specifics)
"I noticed you offer drop-in classes — this means you're juggling scheduling complexity." (em dash, banned word)
"I noticed your website showcases a variety of treatment options." (vague, no real observation, no problem)
"""

def build_template(niche, personalised_line):
    cfg = NICHE_CONFIG[niche]
    bullet_lines = "\n".join(f"- {b}" for b in cfg["bullets"])
    return f"""\
Hi there,

I'm a Physics student at Imperial College London, and {cfg['intro']}

{personalised_line}

The kind of questions I can usually help with are things like:
{bullet_lines}

I work from whatever data you already have (things like {cfg['data_examples']}), so there is nothing to set up or change. It is genuinely a small project, not a big commitment, and if after seeing what I can do you do not find it useful, you have lost nothing.

If any of that sounds interesting, feel free to reply and I can show you a quick example of what the output looks like.

Thanks,
Omri Shahar
Physics, Imperial College London"""

FOLLOWUP_TEMPLATE = """\
Hi,

I just wanted to follow up on my email from last week in case it got buried.

I'm a Physics student at Imperial and I reached out because I noticed something specific about your business that made me think there could be something useful in your data. If you do have records sitting around that you are not making full use of, I would be happy to take a look with no upfront commitment.

Feel free to ignore this if it is not relevant, but if it is, I would love to hear from you.

Thanks,
Omri Shahar
Physics, Imperial College London"""

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
#  HELPERS
# ═══════════════════════════════════════════════════════════

def next_followup_date(base_date):
    return base_date + timedelta(days=7)

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

def generate_personalised_line(lead, website_text):
    client = anthropic.Anthropic(
        api_key=ANTHROPIC_API_KEY,
        base_url="https://api.anthropic.com"
    )
    cfg = NICHE_CONFIG[lead["niche"]]
    message = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=120,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": f"""Business name: {lead['name']}
Niche: {lead['niche']}
Known data problems for this niche: {cfg['problems']}
Website content:
{website_text if website_text else '[Not available]'}

Write the single personalised sentence now."""}]
    )
    return message.content[0].text.strip()

def assemble_email(lead, personalised_line):
    body    = build_template(lead["niche"], personalised_line)
    subject = NICHE_CONFIG[lead["niche"]]["subject"]
    return subject, body

def send_email(to_address, subject, body, smtp_server):
    msg = MIMEMultipart()
    msg["From"]    = GMAIL_ADDRESS
    msg["To"]      = to_address
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))
    smtp_server.sendmail(GMAIL_ADDRESS, to_address, msg.as_string())

# ═══════════════════════════════════════════════════════════
#  EXCEL HELPERS
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
        name_i, email_i, website_i, status_i = col["Business Name"], col["Email"], col["Website"], col["Status"]
    except KeyError as e:
        log.error(f"Missing column in '{niche}': {e}")
        return []
    leads = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        if len(leads) >= n:
            break
        if row[status_i].value == "To Contact" and row[email_i].value and row[email_i].value != "NOT FOUND":
            leads.append({
                "row":     row[0].row,
                "name":    row[name_i].value,
                "email":   row[email_i].value,
                "website": row[website_i].value,
                "niche":   niche,
            })
    wb.close()
    return leads

def get_random_uncontacted_lead(excel_path):
    all_leads = []
    for niche in ACTIVE_NICHES:
        all_leads.extend(get_leads(excel_path, niche, n=999))
    if not all_leads:
        return None
    return random.choice(all_leads)

def update_excel_sent(excel_path, niche, row_num, subject, body):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb[niche]
    headers = [cell.value for cell in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers)}
    today    = datetime.now()
    followup = next_followup_date(today)
    ws.cell(row=row_num, column=col["Status"]).value          = "Sent"
    ws.cell(row=row_num, column=col["Email Sent Date"]).value = today.strftime("%Y-%m-%d")
    ws.cell(row=row_num, column=col["Follow Up Date"]).value  = followup.strftime("%Y-%m-%d")
    if "Notes" in col:
        ws.cell(row=row_num, column=col["Notes"]).value = f"Subject: {subject}\n\n{body}"
    wb.save(excel_path)
    wb.close()

def get_followup_leads(excel_path, niche):
    wb = openpyxl.load_workbook(excel_path)
    if niche not in wb.sheetnames:
        return []
    ws = wb[niche]
    headers = [cell.value for cell in ws[1]]
    try:
        col = {h: i for i, h in enumerate(headers)}
        name_i, email_i, status_i, followup_i, notes_i = (
            col["Business Name"], col["Email"], col["Status"], col["Follow Up Date"], col["Notes"]
        )
    except KeyError as e:
        log.error(f"Missing column in '{niche}': {e}")
        return []
    today = datetime.now().strftime("%Y-%m-%d")
    leads = []
    for row in ws.iter_rows(min_row=2, values_only=False):
        status   = row[status_i].value
        followup = row[followup_i].value
        email    = row[email_i].value
        if followup:
            followup_str = followup.strftime("%Y-%m-%d") if hasattr(followup, "strftime") else str(followup)[:10]
        else:
            followup_str = ""
        if status == "Sent" and followup_str == today and email and email != "NOT FOUND":
            notes            = row[notes_i].value or ""
            original_subject = notes.split("\n")[0].replace("Subject: ", "") if notes.startswith("Subject: ") else ""
            leads.append({
                "row":              row[0].row,
                "name":             row[name_i].value,
                "email":            email,
                "niche":            niche,
                "original_subject": original_subject,
            })
    wb.close()
    return leads

def update_excel_followup(excel_path, niche, row_num):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb[niche]
    headers = [cell.value for cell in ws[1]]
    col = {h: i + 1 for i, h in enumerate(headers)}
    ws.cell(row=row_num, column=col["Status"]).value         = "Followed Up"
    ws.cell(row=row_num, column=col["Follow Up Date"]).value = datetime.now().strftime("%Y-%m-%d")
    wb.save(excel_path)
    wb.close()

# ═══════════════════════════════════════════════════════════
#  MODES
# ═══════════════════════════════════════════════════════════

def run_test_mode():
    print("\n--- TEST MODE ---")
    test_address = input("Send test email to: ").strip()
    if not test_address:
        print("No address entered. Aborting.")
        return

    print("\nPicking a random uncontacted business...")
    lead = get_random_uncontacted_lead(EXCEL_PATH)
    if not lead:
        print("No uncontacted leads found in the Excel file.")
        return

    print(f"Using: {lead['name']} ({lead['niche']}) — {lead['website']}")
    print("Fetching website...")
    website_text = fetch_website(lead["website"])
    print("Generating personalised line...")
    personalised_line = generate_personalised_line(lead, website_text)
    subject, body     = assemble_email(lead, personalised_line)

    print("\n--- EMAIL PREVIEW ---")
    print(f"Business:  {lead['name']} ({lead['niche']})")
    print(f"To:        {test_address}  (real address: {lead['email']})")
    print(f"Subject:   {subject}")
    print(f"\n{body}")
    print("---------------------\n")

    confirm = input("Send this test email? [y/N]: ").strip().lower()
    if confirm != "y":
        print("Cancelled.")
        return

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        send_email(test_address, subject, body, server)
    print(f"Test email sent to {test_address}.")
    print("(Excel file has NOT been updated.)")


def run_send_mode():
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

    sent = failed = 0

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        log.info("Gmail connected.\n")

        for i, lead in enumerate(all_leads, 1):
            try:
                website_text      = fetch_website(lead["website"])
                personalised_line = generate_personalised_line(lead, website_text)
                subject, body     = assemble_email(lead, personalised_line)
                send_email(lead["email"], subject, body, server)
                update_excel_sent(EXCEL_PATH, lead["niche"], lead["row"], subject, body)
                sent += 1
                log.info(f"[{i:02d}/{total}] SENT    {lead['niche']:20s}  {lead['name']:35s}  ->  {lead['email']}")
            except Exception as e:
                failed += 1
                log.error(f"[{i:02d}/{total}] FAILED  {lead['name']} — {e}")
            if i < total:
                time.sleep(DELAY_SECONDS)

    log.info(f"\n========== DONE ==========")
    log.info(f"Sent: {sent}  |  Failed: {failed}  |  Log: outreach_log.txt")


def run_followup_mode():
    today_str = datetime.now().strftime("%Y-%m-%d")
    log.info(f"========== FOLLOW-UP RUN — {today_str} ==========")

    all_leads = []
    for niche in ACTIVE_NICHES:
        leads = get_followup_leads(EXCEL_PATH, niche)
        if leads:
            log.info(f"{niche:20s} — {len(leads)} follow-ups due")
        all_leads.extend(leads)

    if not all_leads:
        log.info("No follow-ups due today.")
        return

    total = len(all_leads)
    log.info(f"\nTotal follow-ups to send: {total}\n")

    sent = failed = 0

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        log.info("Gmail connected.\n")

        for i, lead in enumerate(all_leads, 1):
            subject = f"Re: {lead['original_subject']}" if lead["original_subject"] else "Following up"
            try:
                msg = MIMEMultipart()
                msg["From"]    = GMAIL_ADDRESS
                msg["To"]      = lead["email"]
                msg["Subject"] = subject
                msg.attach(MIMEText(FOLLOWUP_TEMPLATE, "plain"))
                server.sendmail(GMAIL_ADDRESS, lead["email"], msg.as_string())
                update_excel_followup(EXCEL_PATH, lead["niche"], lead["row"])
                sent += 1
                log.info(f"[{i:02d}/{total}] SENT    {lead['niche']:20s}  {lead['name']:35s}  ->  {lead['email']}")
            except Exception as e:
                failed += 1
                log.error(f"[{i:02d}/{total}] FAILED  {lead['name']} — {e}")
            if i < total:
                time.sleep(DELAY_SECONDS)

    log.info(f"\n========== DONE ==========")
    log.info(f"Sent: {sent}  |  Failed: {failed}  |  Log: outreach_log.txt")

# ═══════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════

def main():
    print("\n========== OUTREACH PIPELINE ==========")
    print("[1] Test — generate a sample email and send to yourself")
    print("[2] Send — run today's daily outreach (5 per niche)")
    print("[3] Follow-up — send follow-ups due today")
    print("=======================================")
    choice = input("Choose [1/2/3]: ").strip()

    if choice == "1":
        run_test_mode()
    elif choice == "2":
        run_send_mode()
    elif choice == "3":
        run_followup_mode()
    else:
        print("Invalid choice. Exiting.")

if __name__ == "__main__":
    main()