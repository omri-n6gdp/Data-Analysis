"""
=============================================================
  FOLLOW-UP OUTREACH PIPELINE
  Run this on the same days as daily_outreach.py (Tue/Wed/Thu).
  It finds every lead where Follow Up Date is today and sends
  a short follow-up email.
=============================================================
"""

import time
import smtplib
import logging
import openpyxl
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ═══════════════════════════════════════════════════════════
#  CONFIG
# ═══════════════════════════════════════════════════════════

from secrets_local import GMAIL_ADDRESS  # kept in secrets_local.py (not uploaded)
from secrets_local import GMAIL_APP_PASSWORD  # real value lives in secrets_local.py (not uploaded)
EXCEL_PATH         = "/Users/omrishahar/Documents/Data Analysis/leads_master.xlsx"

DELAY_SECONDS = 10

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
#  FOLLOW-UP TEMPLATE
# ═══════════════════════════════════════════════════════════

FOLLOWUP_TEMPLATE = """\
Hi,

I just wanted to follow up on my email from last week in case it got buried.

I'm a Physics student at Imperial and I reached out because I noticed something in your data that I thought might be useful. If you do have any unused data sitting around and would like to know what insights could be pulled from it, I'd be happy to take a look with no upfront commitment.

Feel free to ignore this if it's not relevant, but if it is, I'd love to hear from you.

Thanks,
Omri Shahar
Physics, Imperial College London\
"""

# ═══════════════════════════════════════════════════════════
#  LOGGING
# ═══════════════════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.FileHandler("followup_log.txt"),
        logging.StreamHandler()
    ]
)
log = logging.getLogger()

# ═══════════════════════════════════════════════════════════
#  GET LEADS DUE FOR FOLLOW-UP TODAY
# ═══════════════════════════════════════════════════════════

def get_followup_leads(excel_path, niche):
    wb = openpyxl.load_workbook(excel_path)
    if niche not in wb.sheetnames:
        return []

    ws = wb[niche]
    headers = [cell.value for cell in ws[1]]

    try:
        col = {h: i for i, h in enumerate(headers)}
        name_i      = col["Business Name"]
        email_i     = col["Email"]
        status_i    = col["Status"]
        followup_i  = col["Follow Up Date"]
        notes_i     = col["Notes"]
    except KeyError as e:
        log.error(f"Missing column in '{niche}': {e}")
        return []

    today = datetime.now().strftime("%Y-%m-%d")
    leads = []

    for row in ws.iter_rows(min_row=2, values_only=False):
        status   = row[status_i].value
        followup = row[followup_i].value
        email    = row[email_i].value

        # Normalise followup date to string for comparison
        if followup:
            if hasattr(followup, "strftime"):
                followup_str = followup.strftime("%Y-%m-%d")
            else:
                followup_str = str(followup)[:10]
        else:
            followup_str = ""

        if status == "Sent" and followup_str == today and email and email != "NOT FOUND":
            # Extract original subject from Notes to use as Re: subject
            notes = row[notes_i].value or ""
            original_subject = ""
            if notes.startswith("Subject: "):
                original_subject = notes.split("\n")[0].replace("Subject: ", "")

            leads.append({
                "row":              row[0].row,
                "name":             row[name_i].value,
                "email":            email,
                "niche":            niche,
                "original_subject": original_subject,
            })

    wb.close()
    return leads

# ═══════════════════════════════════════════════════════════
#  UPDATE EXCEL AFTER FOLLOW-UP
# ═══════════════════════════════════════════════════════════

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
#  SEND
# ═══════════════════════════════════════════════════════════

def send_followup(to_address, subject, smtp_server):
    msg = MIMEMultipart()
    msg["From"]    = GMAIL_ADDRESS
    msg["To"]      = to_address
    msg["Subject"] = subject
    msg.attach(MIMEText(FOLLOWUP_TEMPLATE, "plain"))
    smtp_server.sendmail(GMAIL_ADDRESS, to_address, msg.as_string())

# ═══════════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════════

def main():
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

    sent   = 0
    failed = 0

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        log.info("Gmail connected.\n")

        for i, lead in enumerate(all_leads, 1):
            # Use Re: original subject so it threads in their inbox
            subject = f"Re: {lead['original_subject']}" if lead["original_subject"] else "Following up"

            try:
                send_followup(lead["email"], subject, server)
                update_excel_followup(EXCEL_PATH, lead["niche"], lead["row"])
                sent += 1
                log.info(f"[{i:02d}/{total}] SENT    {lead['niche']:20s}  {lead['name']:35s}  ->  {lead['email']}")
            except Exception as e:
                failed += 1
                log.error(f"[{i:02d}/{total}] FAILED  {lead['name']} — {e}")

            if i < total:
                time.sleep(DELAY_SECONDS)

    log.info(f"\n========== DONE ==========")
    log.info(f"Sent: {sent}  |  Failed: {failed}  |  Log: followup_log.txt")

if __name__ == "__main__":
    main()
