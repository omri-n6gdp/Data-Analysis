"""
TEST SCRIPT — sends one email to yourself to verify everything works.
"""

import smtplib
import requests
import anthropic
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from bs4 import BeautifulSoup

# ── CONFIG ────────────────────────────────────────────────
from secrets_local import GMAIL_ADDRESS  # kept in secrets_local.py (not uploaded)
from secrets_local import GMAIL_APP_PASSWORD  # real value lives in secrets_local.py (not uploaded)
from secrets_local import ANTHROPIC_API_KEY  # real value lives in secrets_local.py (not uploaded)

from secrets_local import TEST_RECIPIENT  # kept in secrets_local.py (not uploaded)
TEST_BUSINESS      = "Absolute Martial Arts"
TEST_WEBSITE       = "https://www.absolutemartialarts.co.uk"
TEST_NICHE         = "Gym"
TEST_NICHE_ANGLE   = "class scheduling, membership retention, and capacity utilisation"
TEST_DATA_EXAMPLES = "booking records, attendance logs, or membership history"
# ─────────────────────────────────────────────────────────

SYSTEM_PROMPT = """\
You write the middle paragraph of a cold outreach email. The email is written in first person — you are Omri, a Physics student at Imperial College London. Write as if you are Omri writing directly to the recipient. Do not refer to Omri in third person at any point.

You only write the middle paragraph — 3 sentences maximum.

Rules:
- Sentence 1: Start with "I noticed on your website that..." and state one specific observation. Just describe what you saw — do not editorialise, do not imply the business has a problem, do not add commentary.
- Sentence 2: Say what many businesses in this niche find difficult to see in their data — keep it factual and plain.
- Sentence 3: Say what insights I could give — simple and specific, no lists or bullet points.
- Write in first person throughout — never say "Omri" at any point.
- Sound human and fairly casual — not stiff, not corporate, not AI-sounding.
- No em dashes, no lists, no bullet points.
- Never use: "juggling", "complexity", "delve", "leverage", "unlock", "empower", "streamline", "quick data thing", "freelance data analysis".
- Do not include a subject line, greeting, or sign-off — only the middle paragraph.

GOOD EXAMPLES (write like these):

Example 1 — Gym:
"I noticed on your website that you run classes across three locations with different studio sizes. Many pilates studios find it difficult to see clearly which locations and time slots are underperforming relative to their capacity, and whether cancellation patterns differ by site. I could give insights into exactly where the gaps are and how to get more out of your existing schedule."

Example 2 — Dentist:
"I noticed on your website that you operate a same-day emergency service around the clock. Many emergency dental practices find it difficult to see clearly which times of day and days of the week drive the most demand, making staffing and capacity decisions harder than they need to be. I could give insights into exactly where those demand patterns sit and how to plan around them."

Example 3 — Estate Agent:
"I noticed on your website that you have been operating in Winchmore Hill and Southgate for over 30 years. Many long-established agencies find it difficult to make systematic use of the historical data they have built up over time, and what it could tell them about seasonal patterns, pricing trends, and property type performance. I could give insights into exactly what that data shows for your patch."

Example 4 — Letting Agent:
"I noticed on your website that you manage lettings across Plaistow, Barking and Stratford from two offices. Many letting agents find it difficult to see clearly how void periods and re-let times differ across their portfolio, and whether certain property types or locations are consistently underperforming. I could give insights into exactly where those patterns sit across your book."

BAD EXAMPLE (never write like this):
"I noticed on your website that you're running classes across multiple locations with different age groups, which means you're juggling a lot of scheduling complexity. Most gyms with this setup struggle to fill classes consistently and don't have a clear picture of which time slots and locations are actually driving member retention. Omri could help you analyse your booking patterns and attendance data to show which classes have the best member retention rates, which time slots consistently underperform, and where you could adjust your timetable to get better capacity utilisation across all your locations."

Why the bad example fails: it editorialises ("juggling complexity"), refers to Omri in third person, and uses a list structure. The good examples are plain, first person, and read like a human wrote them.
"""

TEMPLATE = """\
Hi,

I'm a Physics student at Imperial with experience handling large datasets and turning them into useful information.

{middle}

If you have any unused data (like {data_examples}) that you feel you are not making the most of, feel free to reach out and I would be happy to share an example of what I could extract, with no upfront commitment. My analysis is not limited to the above either, so whatever you feel you need, I am happy to take a look.

Thanks,
Omri Shahar
Physics, Imperial College London\
"""

def fetch_website(url):
    print(f"  Fetching website: {url}")
    try:
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
        soup = BeautifulSoup(r.text, "html.parser")
        for tag in soup(["script", "style", "nav", "footer", "header"]):
            tag.decompose()
        text = soup.get_text(separator=" ", strip=True)[:1500]
        print(f"  Website fetched — {len(text)} characters")
        return text
    except Exception as e:
        print(f"  Could not fetch website: {e}")
        return ""

def generate_middle(website_text):
    print("  Calling Claude API...")
    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
    message = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=300,
        system=SYSTEM_PROMPT,
        messages=[{"role": "user", "content": f"""Business name: {TEST_BUSINESS}
Niche: {TEST_NICHE}
Key angles to consider: {TEST_NICHE_ANGLE}
Website content:
{website_text}

Write the middle paragraph now."""}]
    )
    middle = message.content[0].text.strip()
    print(f"  Claude generated middle paragraph.")
    return middle

def send_test_email(subject, body):
    print(f"  Connecting to Gmail...")
    msg = MIMEMultipart()
    msg["From"]    = GMAIL_ADDRESS
    msg["To"]      = TEST_RECIPIENT
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_ADDRESS, TEST_RECIPIENT, msg.as_string())
    print(f"  Email sent to {TEST_RECIPIENT}")

def main():
    print("\n========== TEST RUN ==========\n")

    print("STEP 1 — Fetching website")
    website_text = fetch_website(TEST_WEBSITE)

    print("\nSTEP 2 — Generating email with Claude")
    middle = generate_middle(website_text)

    subject = "Something I noticed about your scheduling"
    body    = TEMPLATE.format(middle=middle, data_examples=TEST_DATA_EXAMPLES)

    print("\n--- EMAIL PREVIEW ---")
    print(f"To:      {TEST_RECIPIENT}")
    print(f"Subject: {subject}")
    print(f"\n{body}")
    print("---------------------\n")

    print("STEP 3 — Sending email")
    send_test_email(subject, body)

    print("\n========== TEST COMPLETE ==========")
    print("Check your inbox at", TEST_RECIPIENT)

if __name__ == "__main__":
    main()
