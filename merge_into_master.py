"""
MERGE NEW LEADS INTO MASTER EXCEL
==================================
Reads all leads_<niche>_new.csv and leads_no_email_<niche>_new.csv files
from the Data Analysis folder and appends new rows into the correct tabs
of leads_master.xlsx, deduplicating against existing data.

HOW TO RUN:
    python merge_into_master.py

REQUIREMENTS:
    pip install pandas openpyxl
"""

import os
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment

# ── CONFIG ────────────────────────────────────────────────────────────────────
DATA_DIR    = "/Users/omrishahar/Documents/Data Analysis"
MASTER_PATH = "/Users/omrishahar/Documents/Data Analysis/leads_master.xlsx"

# Maps niche label (from CSV filename) -> (main tab name, hunter.io tab name)
NICHE_TABS = {
    "gym":              ("Gym",              "Gym (Hunter.io)"),
    "estate_agent":     ("Estate Agent",     "Estate Agent (Hunter.io)"),
    "accountant":       ("Accountant",       "Accountant (Hunter.io)"),
    "dentist":          ("Dentist",          "Dentist (Hunter.io)"),
    "physio":           ("Physio",           "Physio (Hunter.io)"),
    "corporate_events": ("Corporate Events", "Events (Hunter.io)"),
    "letting_agent":    ("Letting Agent",    "Letting (Hunter.io)"),
    "tutor":            ("Tutor",            "Tutor (Hunter.io)"),
}

COLS = ["Business Name", "Email", "Website", "Address",
        "Niche", "Status", "Email Sent Date", "Follow Up Date", "Reply", "Notes"]

# Alternating row fills matching existing master formatting
FILL_ODD  = PatternFill("solid", fgColor="F5F5F5")   # light grey
FILL_EVEN = PatternFill("solid", fgColor="FFFFFF")    # white

# ── HELPERS ───────────────────────────────────────────────────────────────────

def normalise(s):
    """Lowercase, strip, collapse whitespace — for dedup comparisons."""
    return re.sub(r'\s+', ' ', str(s).lower().strip()) if s else ""

def get_existing_keys(ws):
    """Return sets of (normalised name, normalised email, normalised website) already in sheet."""
    names, emails, websites = set(), set(), set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]: names.add(normalise(row[0]))
        if row[1]: emails.add(normalise(row[1]))
        if row[2]: websites.add(normalise(row[2]))
    return names, emails, websites

def append_rows(ws, rows_df):
    """Append new rows to worksheet with alternating fill, matching master style."""
    next_row = ws.max_row + 1
    for i, (_, row) in enumerate(rows_df.iterrows()):
        fill = FILL_ODD if (next_row % 2 == 0) else FILL_EVEN
        for col_idx, col_name in enumerate(COLS, start=1):
            cell = ws.cell(row=next_row, column=col_idx)
            cell.value = row.get(col_name, "") or ""
            cell.fill  = fill
            cell.font  = Font(name="Calibri", size=11)
            cell.alignment = Alignment(wrap_text=(col_name == "Notes"))
        next_row += 1
    return next_row - (ws.max_row + 1 - len(rows_df))  # rows added

# ── MAIN ──────────────────────────────────────────────────────────────────────

def main():
    print(f"Loading master: {MASTER_PATH}")
    wb = load_workbook(MASTER_PATH)

    total_added = 0

    for niche, (main_tab, hunter_tab) in NICHE_TABS.items():
        main_csv    = os.path.join(DATA_DIR, f"leads_{niche}_new.csv")
        hunter_csv  = os.path.join(DATA_DIR, f"leads_no_email_{niche}_new.csv")

        for csv_path, tab_name in [(main_csv, main_tab), (hunter_csv, hunter_tab)]:
            if not os.path.exists(csv_path):
                print(f"  [{niche}] {os.path.basename(csv_path)} not found, skipping")
                continue

            df = pd.read_csv(csv_path, dtype=str).fillna("")

            # Ensure all expected columns present
            for col in COLS:
                if col not in df.columns:
                    df[col] = ""

            if tab_name not in wb.sheetnames:
                print(f"  WARNING: tab '{tab_name}' not found in master, skipping")
                continue

            ws = wb[tab_name]
            existing_names, existing_emails, existing_websites = get_existing_keys(ws)

            # Filter to genuinely new rows
            def is_new(row):
                name    = normalise(row.get("Business Name", ""))
                email   = normalise(row.get("Email", ""))
                website = normalise(row.get("Website", ""))
                if name    and name    in existing_names:    return False
                if email   and email   in existing_emails:   return False
                if website and website in existing_websites: return False
                return True

            new_rows = df[df.apply(is_new, axis=1)]
            skipped  = len(df) - len(new_rows)

            if new_rows.empty:
                print(f"  [{niche}] '{tab_name}': 0 new rows (all {len(df)} already exist)")
                continue

            added = append_rows(ws, new_rows)
            total_added += added
            print(f"  [{niche}] '{tab_name}': +{added} new rows added  ({skipped} duplicates skipped)")

    output_path = MASTER_PATH.replace(".xlsx", "_updated.xlsx")
    wb.save(output_path)
    print(f"\nDone. {total_added} total rows added.")
    print(f"Saved to: {output_path}")
    print(f"\nReview it, then rename to leads_master.xlsx when happy.")

if __name__ == "__main__":
    main()
