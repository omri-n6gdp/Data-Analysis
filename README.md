# Data Analysis: lead generation and outreach

Local-business lead scraping, a merged master lead list, and AI-written outreach emails sent via Gmail.

## What's here

- `Python/lead_scraper.py`: scrapes local-business leads into CSVs
- `Python/merge_into_master.py`, `merge_into_master.py`: merge new leads into the master list
- `Python/outreach.py`, `Python/daily_outreach.py`, `Python/followup_outreach.py`: write and send outreach emails (Claude writes them, Gmail sends them)
- `Python/report.py`: summary report of outreach so far
- `Python/test_email.py`: sends one test email
- Lead lists (`Emails/*.csv`) and the master spreadsheets (`*.xlsx`) are **not** in this public repo, because they hold other people's contact details.

## Keys (not in this repo)

The scripts read your Gmail address, app password, Anthropic API key and test recipient from `Python/secrets_local.py`.
That file stays on your computer and is never uploaded. To set it up on a new computer, create it with:

```python
GMAIL_ADDRESS = "you@gmail.com"
GMAIL_APP_PASSWORD = "your gmail app password"
ANTHROPIC_API_KEY = "your anthropic key"
TEST_RECIPIENT = "where test emails go"
```
