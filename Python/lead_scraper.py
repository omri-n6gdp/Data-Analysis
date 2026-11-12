"""
FREELANCE DATA ANALYSIS — LEAD SCRAPER (v13)
=============================================
8 niches, multiple slugs per niche.
Full England coverage: London + all major cities and towns.
(Scotland and Wales excluded.)

Saves per niche:
  leads_<niche>_new.csv          — confirmed emails, ready to send
  leads_no_email_<niche>_new.csv — has website but no email (use Hunter.io)

REQUIREMENTS:
    pip install requests beautifulsoup4 pandas tqdm

HOW TO RUN:
    python lead_scraper.py
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
import re
import time
import random
import os
from concurrent.futures import ThreadPoolExecutor, as_completed
from tqdm import tqdm

# ──────────────────────────────────────────────
# CONFIGURATION
# ──────────────────────────────────────────────

OUTPUT_DIR = "/Users/omrishahar/Documents/Data Analysis"

NICHES = {
    # "gym": (                                          # ✓ DONE
    #     ["health-clubs", "yoga", "martial-arts", "leisure-centres"],
    #     ["health club", "gym", "fitness", "sport", "leisure", "wellness",
    #      "yoga", "pilates", "crossfit", "boxing", "martial art", "personal train",
    #      "swim", "squash", "tennis", "dance", "studio"],
    # ),
    # "estate_agent": (                                 # ✓ DONE
    #     ["estate-agents"],
    #     ["estate agent", "estate agenc", "property agent", "property sale",
    #      "property service", "residential sale", "commercial property",
    #      "property valuat", "surveyor"],
    # ),
    # "accountant": (                                   # ✓ DONE
    #     ["accountants", "business-consultants"],
    #     ["accountant", "accounting", "bookkeep", "tax", "payroll",
    #      "chartered", "financial service", "audit", "vat",
    #      "business consultant", "management consultant"],
    # ),
    # "dentist": (                                      # ✓ DONE
    #     ["dentists", "dental-practices"],
    #     ["dentist", "dental", "orthodont", "implant", "cosmetic teeth",
    #      "oral", "teeth", "endodont", "dental practice", "dental surgery"],
    # ),
    # "physio": (                                       # ✓ DONE
    #     ["physiotherapists", "osteopaths", "chiropractors"],
    #     ["physio", "osteopath", "chiropract", "sports therapy",
    #      "rehabilitation", "sports injury", "massage therapy",
    #      "biomechanic", "manual therapy", "musculoskeletal"],
    # ),
    "corporate_events": (
        ["corporate-events"],
        ["event", "conference", "corporate", "venue", "function",
         "party plan", "exhibition", "seminar", "hospitality",
         "meeting room", "banquet"],
    ),
    "letting_agent": (
        ["letting-agents", "property-management"],
        ["letting", "lettings", "rental", "property management",
         "landlord", "tenant", "property let", "residential let",
         "property manager", "block management"],
    ),
    "tutor": (
        ["tutors", "science-tutors"],
        ["tutor", "tuition", "teaching", "coaching", "education",
         "revision", "learning", "academic", "exam prep",
         "a-level", "gcse", "science", "maths", "physics",
         "chemistry", "biology", "school"],
    ),
}

AREAS = [
    # ── All 33 London boroughs ──────────────────────────────────────────────
    "barking-london", "barnet-london", "bexley-london", "brent-london",
    "bromley-london", "camden-london", "city-of-london-london", "croydon-london",
    "ealing-london", "enfield-london", "greenwich-london", "hackney-london",
    "hammersmith-london", "haringey-london", "harrow-london", "havering-london",
    "hillingdon-london", "hounslow-london", "islington-london", "kensington-london",
    "kingston-london", "lambeth-london", "lewisham-london", "merton-london",
    "newham-london", "redbridge-london", "richmond-london", "southwark-london",
    "sutton-london", "tower-hamlets-london", "waltham-forest-london",
    "wandsworth-london", "westminster-london",

    # ── Key London neighbourhoods ───────────────────────────────────────────
    "shoreditch-london", "brixton-london", "clapham-london", "fulham-london",
    "chelsea-london", "notting-hill-london", "canary-wharf-london",
    "wimbledon-london", "stratford-london", "dalston-london", "peckham-london",
    "bethnal-green-london", "stoke-newington-london", "mayfair-london",
    "soho-london", "fitzrovia-london", "clerkenwell-london", "bermondsey-london",

    # ── Surrey ─────────────────────────────────────────────────────────────
    "guildford-surrey", "woking-surrey", "epsom-surrey", "redhill-surrey",
    "reigate-surrey", "leatherhead-surrey", "camberley-surrey", "farnham-surrey",
    "esher-surrey", "cobham-surrey", "dorking-surrey", "weybridge-surrey",
    "walton-on-thames-surrey", "staines-surrey", "caterham-surrey",
    "oxted-surrey", "haslemere-surrey", "godalming-surrey",

    # ── Hertfordshire ──────────────────────────────────────────────────────
    "st-albans-hertfordshire", "watford-hertfordshire", "hertford-hertfordshire",
    "stevenage-hertfordshire", "hemel-hempstead-hertfordshire",
    "welwyn-garden-city-hertfordshire", "potters-bar-hertfordshire",
    "bishops-stortford-hertfordshire", "hitchin-hertfordshire",
    "letchworth-hertfordshire", "harpenden-hertfordshire",
    "rickmansworth-hertfordshire", "cheshunt-hertfordshire",
    "hoddesdon-hertfordshire", "hatfield-hertfordshire", "bushey-hertfordshire",
    "borehamwood-hertfordshire", "radlett-hertfordshire",

    # ── Essex ──────────────────────────────────────────────────────────────
    "chelmsford-essex", "brentwood-essex", "romford-essex",
    "southend-on-sea-essex", "basildon-essex", "harlow-essex",
    "colchester-essex", "epping-essex", "witham-essex", "braintree-essex",
    "billericay-essex", "rayleigh-essex", "wickford-essex", "grays-essex",
    "thurrock-essex", "clacton-on-sea-essex", "saffron-walden-essex",
    "maldon-essex", "halstead-essex",

    # ── Kent ───────────────────────────────────────────────────────────────
    "sevenoaks-kent", "tunbridge-wells-kent", "maidstone-kent",
    "dartford-kent", "gravesend-kent", "orpington-kent", "tonbridge-kent",
    "canterbury-kent", "folkestone-kent", "dover-kent", "ashford-kent",
    "sittingbourne-kent", "faversham-kent", "whitstable-kent",
    "margate-kent", "ramsgate-kent", "deal-kent", "swanley-kent",
    "chatham-kent", "gillingham-kent", "rochester-kent", "herne-bay-kent",

    # ── Berkshire ──────────────────────────────────────────────────────────
    "reading-berkshire", "windsor-berkshire", "slough-berkshire",
    "maidenhead-berkshire", "bracknell-berkshire", "newbury-berkshire",
    "wokingham-berkshire", "ascot-berkshire", "thatcham-berkshire",
    "sandhurst-berkshire",

    # ── East Sussex ────────────────────────────────────────────────────────
    "brighton-east-sussex", "hove-east-sussex", "eastbourne-east-sussex",
    "hastings-east-sussex", "lewes-east-sussex", "uckfield-east-sussex",
    "bexhill-east-sussex", "seaford-east-sussex", "crowborough-east-sussex",

    # ── West Sussex ────────────────────────────────────────────────────────
    "crawley-west-sussex", "worthing-west-sussex", "chichester-west-sussex",
    "horsham-west-sussex", "bognor-regis-west-sussex",
    "haywards-heath-west-sussex", "burgess-hill-west-sussex",
    "littlehampton-west-sussex", "shoreham-by-sea-west-sussex",
    "east-grinstead-west-sussex",

    # ── Hampshire ──────────────────────────────────────────────────────────
    "southampton-hampshire", "portsmouth-hampshire", "winchester-hampshire",
    "basingstoke-hampshire", "andover-hampshire", "aldershot-hampshire",
    "fareham-hampshire", "gosport-hampshire", "eastleigh-hampshire",
    "farnborough-hampshire", "fleet-hampshire", "romsey-hampshire",
    "petersfield-hampshire", "havant-hampshire", "waterlooville-hampshire",
    "alton-hampshire", "lymington-hampshire",

    # ── Oxfordshire ────────────────────────────────────────────────────────
    "oxford-oxfordshire", "abingdon-oxfordshire", "banbury-oxfordshire",
    "bicester-oxfordshire", "didcot-oxfordshire", "witney-oxfordshire",
    "henley-on-thames-oxfordshire", "thame-oxfordshire",
    "chipping-norton-oxfordshire", "faringdon-oxfordshire",
    "wantage-oxfordshire",

    # ── Buckinghamshire ────────────────────────────────────────────────────
    "milton-keynes-buckinghamshire", "high-wycombe-buckinghamshire",
    "aylesbury-buckinghamshire", "amersham-buckinghamshire",
    "beaconsfield-buckinghamshire", "marlow-buckinghamshire",
    "chesham-buckinghamshire", "buckingham-buckinghamshire",

    # ── Cambridgeshire ─────────────────────────────────────────────────────
    "cambridge-cambridgeshire", "peterborough-cambridgeshire",
    "ely-cambridgeshire", "huntingdon-cambridgeshire",
    "st-ives-cambridgeshire", "wisbech-cambridgeshire",
    "march-cambridgeshire", "st-neots-cambridgeshire",

    # ── Suffolk ────────────────────────────────────────────────────────────
    "ipswich-suffolk", "bury-st-edmunds-suffolk", "lowestoft-suffolk",
    "felixstowe-suffolk", "sudbury-suffolk", "newmarket-suffolk",
    "haverhill-suffolk", "stowmarket-suffolk", "woodbridge-suffolk",

    # ── Norfolk ────────────────────────────────────────────────────────────
    "norwich-norfolk", "great-yarmouth-norfolk", "kings-lynn-norfolk",
    "dereham-norfolk", "thetford-norfolk", "cromer-norfolk",
    "sheringham-norfolk", "north-walsham-norfolk", "wymondham-norfolk",

    # ── Bedfordshire ───────────────────────────────────────────────────────
    "bedford-bedfordshire", "luton-bedfordshire", "dunstable-bedfordshire",
    "leighton-buzzard-bedfordshire", "biggleswade-bedfordshire",
    "sandy-bedfordshire",

    # ── Northamptonshire ───────────────────────────────────────────────────
    "northampton-northamptonshire", "kettering-northamptonshire",
    "corby-northamptonshire", "wellingborough-northamptonshire",
    "daventry-northamptonshire", "rushden-northamptonshire",

    # ── Gloucestershire ────────────────────────────────────────────────────
    "cheltenham-gloucestershire", "gloucester-gloucestershire",
    "stroud-gloucestershire", "cirencester-gloucestershire",
    "tewkesbury-gloucestershire", "dursley-gloucestershire",

    # ── Wiltshire ──────────────────────────────────────────────────────────
    "swindon-wiltshire", "salisbury-wiltshire", "chippenham-wiltshire",
    "trowbridge-wiltshire", "devizes-wiltshire", "melksham-wiltshire",
    "marlborough-wiltshire", "calne-wiltshire", "warminster-wiltshire",
    "westbury-wiltshire",

    # ── Somerset ───────────────────────────────────────────────────────────
    "bath-somerset", "taunton-somerset", "weston-super-mare-somerset",
    "yeovil-somerset", "bridgwater-somerset", "frome-somerset",
    "wells-somerset", "glastonbury-somerset", "shepton-mallet-somerset",

    # ── Bristol ────────────────────────────────────────────────────────────
    "bristol-bristol", "clifton-bristol", "bedminster-bristol",
    "filton-bristol", "westbury-on-trym-bristol",

    # ── Devon ──────────────────────────────────────────────────────────────
    "exeter-devon", "plymouth-devon", "torquay-devon", "barnstaple-devon",
    "paignton-devon", "newton-abbot-devon", "exmouth-devon",
    "sidmouth-devon", "tiverton-devon", "bideford-devon",
    "okehampton-devon", "totnes-devon",

    # ── Cornwall ───────────────────────────────────────────────────────────
    "truro-cornwall", "penzance-cornwall", "newquay-cornwall",
    "falmouth-cornwall", "bodmin-cornwall", "st-austell-cornwall",
    "camborne-cornwall", "redruth-cornwall", "helston-cornwall",
    "launceston-cornwall",

    # ── Dorset ─────────────────────────────────────────────────────────────
    "bournemouth-dorset", "poole-dorset", "weymouth-dorset",
    "dorchester-dorset", "blandford-forum-dorset", "wareham-dorset",
    "bridport-dorset", "sherborne-dorset", "christchurch-dorset",
    "wimborne-dorset",

    # ── Lincolnshire ───────────────────────────────────────────────────────
    "lincoln-lincolnshire", "grimsby-lincolnshire", "scunthorpe-lincolnshire",
    "boston-lincolnshire", "spalding-lincolnshire", "louth-lincolnshire",
    "grantham-lincolnshire", "stamford-lincolnshire", "sleaford-lincolnshire",
    "skegness-lincolnshire",

    # ── Nottinghamshire ────────────────────────────────────────────────────
    "nottingham-nottinghamshire", "mansfield-nottinghamshire",
    "newark-on-trent-nottinghamshire", "worksop-nottinghamshire",
    "retford-nottinghamshire", "hucknall-nottinghamshire",
    "beeston-nottinghamshire",

    # ── Derbyshire ─────────────────────────────────────────────────────────
    "derby-derbyshire", "chesterfield-derbyshire", "matlock-derbyshire",
    "ilkeston-derbyshire", "long-eaton-derbyshire", "swadlincote-derbyshire",
    "ashbourne-derbyshire", "buxton-derbyshire", "glossop-derbyshire",

    # ── Leicestershire ─────────────────────────────────────────────────────
    "leicester-leicestershire", "loughborough-leicestershire",
    "hinckley-leicestershire", "melton-mowbray-leicestershire",
    "market-harborough-leicestershire", "coalville-leicestershire",
    "oadby-leicestershire",

    # ── Staffordshire ──────────────────────────────────────────────────────
    "stoke-on-trent-staffordshire", "stafford-staffordshire",
    "lichfield-staffordshire", "tamworth-staffordshire",
    "burton-on-trent-staffordshire", "cannock-staffordshire",
    "rugeley-staffordshire", "leek-staffordshire",
    "newcastle-under-lyme-staffordshire",

    # ── Shropshire ─────────────────────────────────────────────────────────
    "shrewsbury-shropshire", "telford-shropshire", "oswestry-shropshire",
    "bridgnorth-shropshire", "ludlow-shropshire", "market-drayton-shropshire",

    # ── Herefordshire ──────────────────────────────────────────────────────
    "hereford-herefordshire", "ross-on-wye-herefordshire",
    "ledbury-herefordshire", "leominster-herefordshire",

    # ── Worcestershire ─────────────────────────────────────────────────────
    "worcester-worcestershire", "kidderminster-worcestershire",
    "redditch-worcestershire", "bromsgrove-worcestershire",
    "malvern-worcestershire", "droitwich-worcestershire",
    "evesham-worcestershire",

    # ── Warwickshire ───────────────────────────────────────────────────────
    "warwick-warwickshire", "stratford-upon-avon-warwickshire",
    "rugby-warwickshire", "nuneaton-warwickshire",
    "leamington-spa-warwickshire", "kenilworth-warwickshire",
    "bedworth-warwickshire",

    # ── West Midlands ──────────────────────────────────────────────────────
    "birmingham-west-midlands", "coventry-west-midlands",
    "wolverhampton-west-midlands", "walsall-west-midlands",
    "west-bromwich-west-midlands", "solihull-west-midlands",
    "dudley-west-midlands", "stourbridge-west-midlands",
    "halesowen-west-midlands", "sutton-coldfield-west-midlands",
    "smethwick-west-midlands",

    # ── Greater Manchester ──────────────────────────────────────────────────
    "manchester-greater-manchester", "salford-greater-manchester",
    "stockport-greater-manchester", "bolton-greater-manchester",
    "wigan-greater-manchester", "oldham-greater-manchester",
    "rochdale-greater-manchester", "bury-greater-manchester",
    "ashton-under-lyne-greater-manchester", "altrincham-greater-manchester",
    "sale-greater-manchester", "leigh-greater-manchester",
    "hyde-greater-manchester", "wilmslow-greater-manchester",

    # ── Merseyside ─────────────────────────────────────────────────────────
    "liverpool-merseyside", "birkenhead-merseyside", "southport-merseyside",
    "st-helens-merseyside", "wallasey-merseyside", "bootle-merseyside",
    "huyton-merseyside",

    # ── Cheshire ───────────────────────────────────────────────────────────
    "chester-cheshire", "macclesfield-cheshire", "warrington-cheshire",
    "crewe-cheshire", "congleton-cheshire", "nantwich-cheshire",
    "knutsford-cheshire", "northwich-cheshire", "ellesmere-port-cheshire",
    "runcorn-cheshire", "widnes-cheshire",

    # ── Lancashire ─────────────────────────────────────────────────────────
    "preston-lancashire", "blackpool-lancashire", "blackburn-lancashire",
    "burnley-lancashire", "lancaster-lancashire", "accrington-lancashire",
    "chorley-lancashire", "leyland-lancashire", "lytham-st-annes-lancashire",
    "clitheroe-lancashire", "nelson-lancashire", "morecambe-lancashire",

    # ── Cumbria ────────────────────────────────────────────────────────────
    "carlisle-cumbria", "barrow-in-furness-cumbria", "kendal-cumbria",
    "whitehaven-cumbria", "workington-cumbria", "keswick-cumbria",
    "penrith-cumbria", "windermere-cumbria",

    # ── West Yorkshire ─────────────────────────────────────────────────────
    "leeds-west-yorkshire", "bradford-west-yorkshire",
    "huddersfield-west-yorkshire", "halifax-west-yorkshire",
    "wakefield-west-yorkshire", "dewsbury-west-yorkshire",
    "batley-west-yorkshire", "keighley-west-yorkshire",
    "ilkley-west-yorkshire", "wetherby-west-yorkshire",

    # ── South Yorkshire ────────────────────────────────────────────────────
    "sheffield-south-yorkshire", "rotherham-south-yorkshire",
    "doncaster-south-yorkshire", "barnsley-south-yorkshire",

    # ── East Yorkshire ─────────────────────────────────────────────────────
    "hull-east-yorkshire", "beverley-east-yorkshire",
    "bridlington-east-yorkshire", "goole-east-yorkshire",

    # ── North Yorkshire ────────────────────────────────────────────────────
    "york-north-yorkshire", "harrogate-north-yorkshire",
    "scarborough-north-yorkshire", "northallerton-north-yorkshire",
    "ripon-north-yorkshire", "skipton-north-yorkshire",
    "whitby-north-yorkshire", "selby-north-yorkshire",

    # ── North East ─────────────────────────────────────────────────────────
    "newcastle-upon-tyne-tyne-and-wear", "sunderland-tyne-and-wear",
    "gateshead-tyne-and-wear", "south-shields-tyne-and-wear",
    "durham-county-durham", "middlesbrough-cleveland",
    "stockton-on-tees-cleveland", "darlington-county-durham",
    "hartlepool-cleveland", "bishop-auckland-county-durham",
    "consett-county-durham", "hexham-northumberland",
    "ashington-northumberland", "morpeth-northumberland",
    "blyth-northumberland",

    # ── Isle of Wight ──────────────────────────────────────────────────────
    "newport-isle-of-wight", "ryde-isle-of-wight", "cowes-isle-of-wight",
]

PAGES_PER_AREA = 5    # up from 2 — 2.5x more leads per area
SCRAPE_THREADS = 40   # Phase 1: parallel Thomson Local scraping
EMAIL_THREADS  = 60   # Phase 2: parallel email fetching (all different domains)

# ──────────────────────────────────────────────
# CONSTANTS
# ──────────────────────────────────────────────

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-GB,en;q=0.9",
}

BASE        = "https://www.thomsonlocal.com"
EMAIL_REGEX = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
EMAIL_BLACKLIST = [
    "example.com", "sentry.io", "wixpress.com", "squarespace.com",
    "shopify.com", "wordpress.com", "schema.org", "w3.org",
    ".png", ".jpg", ".gif", "noreply", "no-reply", "privacy@",
    "legal@", "webmaster@", "test@", "info@info", "email@",
    "@email", "your@", "name@", "user@", "domain.com",
]

# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────

def is_valid_email(e):
    return len(e) <= 60 and not any(b in e.lower() for b in EMAIL_BLACKLIST)

def is_correct_niche(lead, keywords, slugs):
    text = (lead.get("address", "") + " " + lead.get("name", "")).lower()
    if any(kw in text for kw in keywords):
        return True
    # Trust the slug — if we searched the right category, it belongs
    lead_slug = lead.get("slug", "")
    if any(s in lead_slug for s in slugs):
        return True
    return False

def fetch(url, timeout=10):
    for _ in range(2):
        try:
            r = requests.get(url, headers=HEADERS, timeout=timeout, allow_redirects=True)
            if r.status_code == 200:
                return r.text
        except Exception:
            pass
    return ""

def extract_emails(html):
    emails = set()
    soup = None
    try:
        soup = BeautifulSoup(html, "html.parser")
    except Exception:
        pass

    for e in EMAIL_REGEX.findall(html):
        if is_valid_email(e):
            emails.add(e.lower())

    if soup:
        for a in soup.find_all("a", href=True):
            if a["href"].startswith("mailto:"):
                e = a["href"].replace("mailto:", "").split("?")[0].strip().lower()
                if e and is_valid_email(e):
                    emails.add(e)

        for tag in ["header", "footer"]:
            section = soup.find(tag)
            if section:
                for e in EMAIL_REGEX.findall(str(section)):
                    if is_valid_email(e):
                        emails.add(e.lower())
                for a in section.find_all("a", href=True):
                    if a["href"].startswith("mailto:"):
                        e = a["href"].replace("mailto:", "").split("?")[0].strip().lower()
                        if e and is_valid_email(e):
                            emails.add(e)

        for cls in ["footer", "site-footer", "page-footer", "site-header", "contact", "contact-info"]:
            section = soup.find("div", class_=lambda c: c and cls in " ".join(c).lower())
            if section:
                for e in EMAIL_REGEX.findall(str(section)):
                    if is_valid_email(e):
                        emails.add(e.lower())

    for match in re.findall(
        r'[\w.+-]+\s*[\[\(]at[\]\)]\s*[\w.-]+\s*[\[\(]dot[\]\)]\s*[\w.]+',
        html, re.IGNORECASE
    ):
        cleaned = re.sub(r'\s*[\[\(]at[\]\)]\s*', '@', match, flags=re.IGNORECASE)
        cleaned = re.sub(r'\s*[\[\(]dot[\]\)]\s*', '.', cleaned, flags=re.IGNORECASE)
        if is_valid_email(cleaned.lower()):
            emails.add(cleaned.lower())

    priority_words = ["info", "hello", "contact", "admin", "owner", "manager", "enquir", "office"]
    priority = [e for e in emails if any(w in e for w in priority_words)]
    other    = [e for e in emails if e not in priority]
    return priority + other

def find_contact_page_url(base, homepage_html):
    try:
        soup = BeautifulSoup(homepage_html, "html.parser")
        for a in soup.find_all("a", href=True):
            href = a["href"].lower()
            text = a.get_text(strip=True).lower()
            if any(w in href or w in text for w in
                   ["contact", "enquir", "reach", "find-us", "get-in-touch", "email-us"]):
                full = (href if href.startswith("http")
                        else base.rstrip("/") + "/" + href.lstrip("/"))
                if base in full:
                    return full
    except Exception:
        pass
    return None

def find_email(website_url):
    if not website_url:
        return None
    base         = website_url.rstrip("/")
    homepage     = fetch(base)
    real_contact = find_contact_page_url(base, homepage)
    pages        = [base, base+"/contact", base+"/contact-us",
                    base+"/about", base+"/about-us", base+"/get-in-touch"]
    if real_contact and real_contact not in pages:
        pages.insert(1, real_contact)
    for i, url in enumerate(pages):
        html = homepage if i == 0 else fetch(url)
        if html:
            emails = extract_emails(html)
            if emails:
                return emails[0]
    return None

# ──────────────────────────────────────────────
# PHASE 1 — Scrape Thomson Local (parallelised)
# ──────────────────────────────────────────────

def scrape_slug_area(args):
    slug, area, pages = args
    leads = []
    for page in range(1, pages + 1):
        url = (f"{BASE}/search/{slug}/{area}/"
               if page == 1
               else f"{BASE}/search/{slug}/{area}/page-{page}/")
        try:
            r    = requests.get(url, headers=HEADERS, timeout=10)
            soup = BeautifulSoup(r.text, "html.parser")
            for header in soup.find_all("ul", class_="listingHeader"):
                content = header.find_next_sibling("div", class_="listContentWrapper")
                name, address = "Unknown", ""
                if content:
                    n = content.find("h2", class_="businessName")
                    name = n.get_text(strip=True) if n else "Unknown"
                    a = content.find("div", class_="address")
                    address = a.get_text(strip=True) if a else ""
                website = ""
                web_li = header.find("li", class_="website")
                if web_li:
                    a = web_li.find("a")
                    if a:
                        website = a.get("href", "")
                if name != "Unknown":
                    leads.append({"name": name, "website": website, "address": address, "slug": slug})
        except Exception:
            pass
        time.sleep(random.uniform(0.1, 0.3))
    return leads

# ──────────────────────────────────────────────
# PHASE 2 — Find email (threaded worker)
# ──────────────────────────────────────────────

def process_lead(args):
    i, total, lead, niche_label = args
    email = find_email(lead["website"]) if lead["website"] else None
    tqdm.write(f"  [{i+1}/{total}] {lead['name'][:45]:45} {'-> ' + email if email else ''}")
    return {
        "Business Name":   lead["name"],
        "Email":           email or "NOT FOUND",
        "Website":         lead["website"] or "NOT FOUND",
        "Address":         lead["address"],
        "Niche":           niche_label,
        "Status":          "To Contact",
        "Email Sent Date": "",
        "Follow Up Date":  "",
        "Reply":           "",
        "Notes":           "",
    }

# ──────────────────────────────────────────────
# SAVE
# ──────────────────────────────────────────────

def save_niche(results, niche_label):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with_email    = [r for r in results if r["Email"] != "NOT FOUND"]
    without_email = [r for r in results if r["Email"] == "NOT FOUND"
                                        and r["Website"] != "NOT FOUND"]
    out_email    = os.path.join(OUTPUT_DIR, f"leads_{niche_label}_new.csv")
    out_no_email = os.path.join(OUTPUT_DIR, f"leads_no_email_{niche_label}_new.csv")
    pd.DataFrame(with_email).to_csv(out_email, index=False)
    pd.DataFrame(without_email).to_csv(out_no_email, index=False)
    pct     = round(100 * len(with_email) / len(results)) if results else 0
    dropped = len(results) - len(with_email) - len(without_email)
    print(f"  leads_{niche_label}_new.csv          -> {len(with_email)} with emails ({pct}%)")
    print(f"  leads_no_email_{niche_label}_new.csv -> {len(without_email)} for Hunter.io")
    print(f"  Dropped {dropped} (no website or email)")
    return len(with_email), len(without_email)

# ──────────────────────────────────────────────
# RUN ONE NICHE
# ──────────────────────────────────────────────

def run_niche(niche_label, slugs, keywords):
    print(f"\n{'='*60}")
    print(f"NICHE: {niche_label.upper()} | slugs: {', '.join(slugs)}")
    print(f"{'='*60}")

    combos       = [(slug, area, PAGES_PER_AREA) for slug in slugs for area in AREAS]
    total_combos = len(combos)
    print(f"\nPHASE 1: {len(slugs)} slug(s) x {len(AREAS)} areas x {PAGES_PER_AREA} pages "
          f"= {total_combos} requests ({SCRAPE_THREADS} threads)...")

    all_raw = []

    with ThreadPoolExecutor(max_workers=SCRAPE_THREADS) as executor:
        futures = {executor.submit(scrape_slug_area, c): c for c in combos}
        with tqdm(total=total_combos, desc="  Scraping", unit="req",
                  bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]") as pbar:
            for future in as_completed(futures):
                slug, area, _ = futures[future]
                try:
                    leads = future.result()
                    all_raw.extend(leads)
                except Exception as e:
                    tqdm.write(f"  Scrape error ({slug}/{area}): {e}")
                pbar.update(1)
                pbar.set_postfix(raw_leads=len(all_raw), refresh=False)

    correct = [l for l in all_raw if is_correct_niche(l, keywords, slugs)]
    print(f"\n  Raw: {len(all_raw)} | After niche filter: {len(correct)} | Removed: {len(all_raw)-len(correct)}")

    seen_names, seen_websites, unique = set(), set(), []
    for l in correct:
        name_key = re.sub(r'\s+', ' ', l["name"].lower().strip())
        web_key  = l["website"].lower().rstrip("/") if l["website"] else None
        if name_key in seen_names:
            continue
        if web_key and web_key in seen_websites:
            continue
        seen_names.add(name_key)
        if web_key:
            seen_websites.add(web_key)
        unique.append(l)
    print(f"  After dedup: {len(unique)} unique businesses")

    if not unique:
        print("  Nothing to process.")
        return 0, 0

    print(f"\nPHASE 2: Finding emails ({EMAIL_THREADS} concurrent)...")
    args    = [(i, len(unique), lead, niche_label) for i, lead in enumerate(unique)]
    results = [None] * len(unique)

    emails_found = 0
    with ThreadPoolExecutor(max_workers=EMAIL_THREADS) as executor:
        futures = {executor.submit(process_lead, a): a[0] for a in args}
        with tqdm(total=len(unique), desc="  Emails  ", unit="biz",
                  bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}, {rate_fmt}]") as pbar:
            for future in as_completed(futures):
                try:
                    result       = future.result()
                    idx          = futures[future]
                    results[idx] = result
                    if result["Email"] != "NOT FOUND":
                        emails_found += 1
                except Exception as e:
                    tqdm.write(f"  Error: {e}")
                pbar.update(1)
                pbar.set_postfix(emails=emails_found, refresh=False)

    results = [r for r in results if r is not None]
    results.sort(key=lambda x: x["Business Name"])

    print(f"\nSaving {niche_label}...")
    return save_niche(results, niche_label)

# ──────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────

def main():
    print("=" * 60)
    print(f"LEAD SCRAPER v13")
    print(f"Niches:  {', '.join(NICHES.keys())}")
    print(f"Areas:   {len(AREAS)} locations — full England coverage")
    print(f"Pages:   {PAGES_PER_AREA} per area")
    print(f"Threads: {SCRAPE_THREADS} scrape / {EMAIL_THREADS} email")
    print(f"Output:  {OUTPUT_DIR}")
    print("=" * 60)

    summary = {}
    niche_list = list(NICHES.items())
    with tqdm(total=len(niche_list), desc="Overall   ", unit="niche",
              bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} niches [{elapsed}<{remaining}]") as overall:
        for niche_label, (slugs, keywords) in niche_list:
            overall.set_postfix(current=niche_label, refresh=True)
            counts = run_niche(niche_label, slugs, keywords)
            summary[niche_label] = counts
            overall.update(1)

    print(f"\n{'='*60}")
    print("ALL DONE — FINAL SUMMARY")
    print(f"{'='*60}")
    total_emails = 0
    for niche, (with_email, without_email) in summary.items():
        print(f"  {niche:<20} {with_email:>5} emails ready  |  {without_email:>5} for Hunter.io")
        total_emails += with_email
    print(f"\n  TOTAL emails ready to send: {total_emails}")
    print(f"  Files saved to: {OUTPUT_DIR}")
    print("=" * 60)

if __name__ == "__main__":
    main()