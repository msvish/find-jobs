#!/usr/bin/env python3
"""
careers_watch.py - watch plain company career pages listed in an Excel sheet
and report job postings that appeared since the last run.

Google Sheet (row 1 = headers), one row per company:
  CompanyName | CareersPage | selector | keywords | active | render   (only the first two are required)
  render    (optional) "yes" forces a headless browser for pages that load jobs with JavaScript
  selector  (optional) CSS selector around the job list, e.g. "#careers" or ".openings"
            -> cuts noise a lot; find it with right-click > Inspect on the page
  keywords  (optional) comma-separated; only report new items containing one, e.g. "engineer,developer"
  active    (optional) "no" to skip a row

Add rows whenever you find a company; the next run baselines it automatically.

Notifications (optional, via environment variables):
  GMAIL_USER, GMAIL_APP_PASSWORD, NOTIFY_TO  -> email
  SLACK_WEBHOOK                              -> Slack message
Every run rewrites new.html (all findings, newest first) from history.json.
"""
import csv, io, json, os, re, smtplib
from datetime import datetime
from zoneinfo import ZoneInfo
from email.mime.text import MIMEText
from pathlib import Path

import requests
from bs4 import BeautifulSoup

from report import render

BASE = Path(__file__).resolve().parent

# Local runs: put KEY=value lines in pipeline/.env (gitignored). In GitHub Actions they come from repo variables.
_env = BASE / ".env"
if _env.exists():
    for _line in _env.read_text().splitlines():
        if "=" in _line and not _line.lstrip().startswith("#"):
            _k, _v = _line.split("=", 1)
            os.environ.setdefault(_k.strip(), _v.strip().strip('"').strip("'"))

# From your sheet URL: https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit#gid=<SHEET_GID>
SHEET_ID = os.getenv("SHEET_ID", "").strip()
SHEET_GID = os.getenv("SHEET_GID", "").strip() or "0"
if not SHEET_ID:
    raise SystemExit("SHEET_ID is not set. In GitHub: Settings > Secrets and variables > Actions > Variables. "
                     "Locally: add SHEET_ID=... to pipeline/.env")
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={SHEET_GID}"
# One script, several watchers: each profile sets its own SHEET_ID, DATA_DIR, REGION and TIMEZONE.
DATA = BASE / os.getenv("DATA_DIR", "").strip() if os.getenv("DATA_DIR", "").strip() else BASE
DATA.mkdir(parents=True, exist_ok=True)
STATE = DATA / "state.json"
HISTORY = DATA / "history.json"
PAGE = DATA / "new.html"
TZ = ZoneInfo(os.getenv("TIMEZONE", "").strip() or "America/Los_Angeles")
MAX_RUNS = 200  # runs kept in history.json

HEADERS = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
           "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36"}
_DEFAULT_ROLES = (r"engineer|developer|programmer|intern(ship)?|analyst|scientist|architect|devops|sre|"
                  r"full[- ]?stack|front[- ]?end|back[- ]?end|software|opening|position|hiring|vacanc(y|ies)")
# ROLE_WORDS replaces the default list for a different field; JOB_WORDS adds to whichever list is used.
# Both are comma-separated regex pieces, e.g. ROLE_WORDS="financ\w*,tax(ation|es)?,fp&a"
_extra = lambda var: [w.strip() for w in os.getenv(var, "").split(",") if w.strip()]
JOB_WORDS = re.compile(r"(?<![A-Za-z])(" + "|".join(_extra("ROLE_WORDS") or [_DEFAULT_ROLES]) +
                       "".join("|" + w for w in _extra("JOB_WORDS")) + r")(?![A-Za-z])", re.I)
# ROLE_KEYWORDS (comma-separated) replaces JOB_WORDS entirely: only titles containing one of them count.
# Used for a profile in another field, e.g. ROLE_KEYWORDS="financial analyst,fp&a,tax analyst"
_rk = [k.strip() for k in os.getenv("ROLE_KEYWORDS", "").split(",") if k.strip()]
ROLE_KEYWORDS = re.compile(r"(?<![A-Za-z])(?:" + "|".join(re.escape(k).replace(r"\ ", r"\s+") for k in _rk) + r")(?![A-Za-z])",
                           re.I) if _rk else None
NOISE = re.compile(r"(cookie|privacy|terms of|copyright|©|subscribe|newsletter|log ?in|sign ?up)", re.I)
EMAIL = re.compile(r"[\w.+-]+@[\w-]+\.[a-z]{2,}(?:\.[a-z]{2,})?", re.I)


# --- Location filter --------------------------------------------------------------------
# REGION=US (default) keeps US roles, REGION=IN keeps India roles, REGION=ALL keeps everything.
# (Old setting US_ONLY=no still works and means ALL.) Roles whose location can't be read are kept
# (shown as "Location not listed") rather than silently dropped.
REGION = (os.getenv("REGION") or "US").strip().upper()
if (os.getenv("US_ONLY") or "").strip().lower() in ("no", "n", "false", "0"):
    REGION = "ALL"

_ABBR = ("AL AK AZ AR CA CO CT DE FL GA HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO MT NE NV NH NJ NM NY "
         "NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY DC").split()
_STATES = ("Alabama|Alaska|Arizona|Arkansas|California|Colorado|Connecticut|Delaware|Florida|Georgia|Hawaii|"
           "Idaho|Illinois|Indiana|Iowa|Kansas|Kentucky|Louisiana|Maine|Maryland|Massachusetts|Michigan|"
           "Minnesota|Mississippi|Missouri|Montana|Nebraska|Nevada|New Hampshire|New Jersey|New Mexico|"
           "New York|North Carolina|North Dakota|Ohio|Oklahoma|Oregon|Pennsylvania|Rhode Island|"
           "South Carolina|South Dakota|Tennessee|Texas|Utah|Vermont|Virginia|West Virginia|Wisconsin|Wyoming")
_CITIES = ("San Francisco|Los Angeles|San Diego|San Jose|Seattle|Portland|Austin|Dallas|Houston|Chicago|Boston|"
           "Atlanta|Miami|Denver|Phoenix|Philadelphia|Pittsburgh|Detroit|Ann Arbor|Minneapolis|Charlotte|Raleigh|"
           "Nashville|Salt Lake City|Brooklyn|Palo Alto|Mountain View|Menlo Park|Sunnyvale|Santa Clara|Redmond|"
           "Bellevue|Los Altos|Washington,? D\\.?C\\.?")
# Case-sensitive: country codes ("US", "USA"), "City, ST" abbreviations.
US_MARK = re.compile(r"(?<![A-Za-z])(?:U\.S\.A?\.?|USA|US)(?![A-Za-z])|,\s*(?:" + "|".join(_ABBR) + r")\b")
# Case-insensitive: "United States", full state names, common US cities.
US_MARK_CI = re.compile(r"United States(?: of America)?|\b(?:" + _STATES + r")\b|\b(?:" + _CITIES + r")\b", re.I)
FOREIGN_MARK = re.compile(
    r"\b(?:India|Canada|United Kingdom|England|Scotland|Wales|Ireland|Germany|France|Spain|Italy|Netherlands|"
    r"Poland|Portugal|Sweden|Norway|Denmark|Finland|Switzerland|Austria|Belgium|Israel|Dubai|UAE|"
    r"Saudi Arabia|Singapore|Hong Kong|Japan|Korea|China|Taiwan|Australia|New Zealand|Brazil|Argentina|"
    r"Colombia|Chile|Peru|Philippines|Vietnam|Indonesia|Malaysia|Thailand|Pakistan|Bangladesh|Nigeria|"
    r"Kenya|South Africa|Egypt|Turkey|Ukraine|Romania|Czech|Hungary|"
    r"London|Berlin|Munich|Paris|Madrid|Barcelona|Amsterdam|Dublin|Warsaw|Lisbon|Stockholm|Copenhagen|"
    r"Zurich|Tel Aviv|Tokyo|Seoul|Beijing|Shanghai|Taipei|Sydney|Melbourne|Toronto|Vancouver|Montreal|"
    r"Ottawa|Calgary|Bangalore|Bengaluru|Hyderabad|Pune|Chennai|Mumbai|Delhi|Gurgaon|Gurugram|Noida|"
    r"Manila|Sao Paulo|S[aã]o Paulo|Mexico City|Europe|Asia)\b|(?<!New )\bMexico\b|\b(?:EMEA|APAC|LATAM|DACH)\b",
    re.I)
IN_MARK = re.compile(
    r"\b(?:India|Bharat|Bangalore|Bengaluru|Hyderabad|Secunderabad|Pune|Chennai|Mumbai|Navi Mumbai|Thane|"
    r"New Delhi|Delhi|Delhi NCR|NCR|Gurgaon|Gurugram|Noida|Greater Noida|Faridabad|Ghaziabad|Kolkata|"
    r"Ahmedabad|Gandhinagar|Vadodara|Surat|Kochi|Cochin|Thiruvananthapuram|Trivandrum|Coimbatore|Madurai|"
    r"Mysore|Mysuru|Mangalore|Mangaluru|Hubli|Belgaum|Jaipur|Chandigarh|Mohali|Indore|Bhopal|Nagpur|Nashik|"
    r"Lucknow|Kanpur|Bhubaneswar|Visakhapatnam|Vizag|Vijayawada|Goa|Panaji|Guwahati|Patna|Ranchi|Dehradun|"
    r"Karnataka|Maharashtra|Telangana|Tamil Nadu|Kerala|Andhra Pradesh|Gujarat|Haryana|Uttar Pradesh|"
    r"West Bengal|Rajasthan|Punjab|Odisha|Madhya Pradesh)\b|(?<![A-Za-z])IND(?![A-Za-z])")
# PLACES narrows REGION to specific cities/states, e.g. PLACES="Bengaluru,Bangalore,Karnataka".
# Roles naming another place are dropped; roles with no location are still kept.
_pl = [p.strip() for p in os.getenv("PLACES", "").split(",") if p.strip()]
PLACES = re.compile(r"\b(?:" + "|".join(re.escape(p).replace(r"\ ", r"\s+") for p in _pl) + r")\b", re.I) if _pl else None
CODES = {"US": ("US", "USA"), "IN": ("IN", "IND")}  # Taleo-style "US-AL-Birmingham", "IN-KA-Bangalore"


def is_us(text):
    if US_MARK_CI.search(text):
        return True
    india = IN_MARK.search(text)  # ", IN" next to an Indian city is India, not Indiana
    return any(not (india and m.group().replace(",", "").strip() == "IN") for m in US_MARK.finditer(text))


def region_status(text):
    """'match' (a place in REGION is named), 'foreign' (only other places named) or 'unknown' (can't tell).
    The REGION check runs first, so "Bangalore or London" counts as India and "NYC or London" as US."""
    if PLACES:
        if PLACES.search(text):
            return "match"
        return "foreign" if IN_MARK.search(text) or is_us(text) or FOREIGN_MARK.search(text) else "unknown"
    if REGION == "IN":
        if IN_MARK.search(text):
            return "match"
        return "foreign" if is_us(text) or FOREIGN_MARK.search(text) else "unknown"
    if is_us(text):  # US (also used for display when REGION=ALL)
        return "match"
    return "foreign" if FOREIGN_MARK.search(text) else "unknown"


def guess_location(ctx, title=""):
    """Short location for display, read from the job card's own text segments; '' when none is recognisable."""
    for seg in re.split(r"\s*[|•·]\s*", ctx):
        seg = " ".join(seg.split())
        if not seg or seg == title or len(seg) > 60:
            continue
        m = re.fullmatch(r"([A-Z]{2,3})-([A-Z]{2})-(.+)", seg)  # Taleo style: US-AL-Birmingham
        if m and m.group(1) in CODES.get(REGION, ("US", "USA")):
            return f"{m.group(3)}, {m.group(2)}"
        if region_status(seg) == "match" or (REGION == "ALL" and FOREIGN_MARK.search(seg)):
            return seg
    return "Remote" if re.search(r"\bremote\b", f"{ctx} {title}", re.I) else ""


def load_sheet():
    r = requests.get(SHEET_URL, timeout=20)
    r.raise_for_status()
    if r.text.lstrip().startswith("<"):
        raise SystemExit("Got a login page instead of CSV - set the sheet's sharing to "
                         "'Anyone with the link: Viewer', and check SHEET_ID / SHEET_GID.")
    out = []
    for raw in csv.DictReader(io.StringIO(r.text)):
        # normalize headers: "CompanyName", "Company Name", "company_name" -> "companyname"
        d = {re.sub(r"[^a-z0-9]", "", str(k).lower()): (v or "").strip() for k, v in raw.items() if k}
        d["company"] = d.get("companyname") or d.get("company", "")
        d["url"] = d.get("careerspage") or d.get("careerpage") or d.get("url", "")
        if not d.get("url") or d.get("active", "").lower() == "no":
            continue
        d["keywords"] = [k.strip().lower() for k in d.get("keywords", "").split(",") if k.strip()]
        out.append(d)
    return out


def fetch_rendered(url):
    """Headless browser for JS-rendered pages. Needs: pip install playwright && playwright install chromium"""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        raise RuntimeError("page needs a browser to render - run: pip install playwright && playwright install chromium")
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(user_agent=HEADERS["User-Agent"])
        page.goto(url, wait_until="networkidle", timeout=60000)
        page.wait_for_timeout(3000)  # let late AJAX job lists (Taleo, etc.) finish filling in
        html = page.content()
        browser.close()
    return html


# Pages whose raw HTML is a shell: lots of text, but the job list arrives later via JavaScript
JS_SHELL_HINTS = ("taleo.net", "myworkdayjobs.com", "icims.com", "successfactors", "oraclecloud.com")


def fetch(url, render=False):
    if render or any(h in url.lower() for h in JS_SHELL_HINTS):
        return fetch_rendered(url)
    r = requests.get(url, headers=HEADERS, timeout=25)
    r.raise_for_status()
    html = r.text
    if len(BeautifulSoup(html, "html.parser").get_text(strip=True)) < 400:  # probably JS-rendered
        html = fetch_rendered(url)
    return html


def extract(html, selector):
    soup = BeautifulSoup(html, "html.parser")
    emails = set(EMAIL.findall(soup.get_text(" ")))
    emails |= {a["href"][7:].split("?")[0] for a in soup.select('a[href^="mailto:"]')}
    for t in soup(["script", "style", "noscript", "svg", "header", "footer", "nav"]):
        t.decompose()
    roots = soup.select(selector) if selector else []
    if not roots:  # selector missing or not found in the HTML -> use whole page
        roots = [soup.body or soup]
    items, ctx, cache = set(), {}, {}

    def clean(node):
        if id(node) not in cache:
            cache[id(node)] = " ".join(node.get_text(" ", strip=True).split())
        return cache[id(node)]

    def segments(node):  # same text, but pieces from different elements are joined with " | "
        if ("s", id(node)) not in cache:
            cache[("s", id(node))] = " | ".join(" ".join(t.split()) for t in node.stripped_strings)
        return cache[("s", id(node))]

    for root in roots:
        for el in root.find_all(["a", "h1", "h2", "h3", "h4", "h5", "li", "p", "td", "strong"]):
            text = clean(el)
            if 4 <= len(text) <= 160 and not NOISE.search(text):
                items.add(text)
                if text not in ctx:  # smallest surrounding block that holds more than the title: the job "card"
                    for anc in list(el.parents)[:4]:
                        if len(text) + 3 < len(clean(anc)) <= 400:
                            ctx[text] = segments(anc)
                            break
    return items, ctx, sorted(e for e in emails if not e.lower().endswith((".png", ".jpg", ".svg")))


def looks_like_title(t):
    """Job titles are short and don't end like a sentence; descriptions do."""
    return len(t.split()) <= 8 and "•" not in t and not t.rstrip().endswith((".", "!", "?"))


def drop_merged(cands, items, ctx):
    """Greenhouse-style cards yield 'Data Engineer' and 'Data Engineer Remote' (title + location line merged).
    Drop the merged one when its extra words are a separate piece of the same card."""
    def merged(a):
        for b in items:
            if b != a and a.startswith(b + " "):
                extra = a[len(b) + 1:].strip()
                if extra in {x.strip() for x in ctx.get(b, "").split("|")}:
                    return True
        return False
    return [a for a in cands if not merged(a)]


def notify(text, html):
    if os.getenv("SLACK_WEBHOOK"):
        requests.post(os.environ["SLACK_WEBHOOK"], json={"text": text}, timeout=15)
    if os.getenv("GMAIL_USER") and os.getenv("GMAIL_APP_PASSWORD"):
        msg = MIMEText(html, "html")
        msg["Subject"] = f"Career page watch - {datetime.now():%b %d}"
        msg["From"] = os.environ["GMAIL_USER"]
        msg["To"] = os.getenv("NOTIFY_TO") or os.environ["GMAIL_USER"]
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(os.environ["GMAIL_USER"], os.environ["GMAIL_APP_PASSWORD"])
            s.send_message(msg)


def main():
    now = datetime.now(TZ)
    rows = load_sheet()
    if not rows:
        return
    state = json.loads(STATE.read_text()) if STATE.exists() else {}
    history = json.loads(HISTORY.read_text()) if HISTORY.exists() else []
    findings, errors = [], []

    for row in rows:
        url, name = row["url"], row.get("company") or row["url"]
        try:
            # a URL like .../careers#open-roles hints at the job section's id
            selector = row.get("selector") or ("#" + url.split("#", 1)[1] if "#" in url else "")
            force_render = row.get("render", "").lower() in ("yes", "y", "true", "1")
            items, ctx, emails = extract(fetch(url, force_render), selector)
        except Exception as e:
            errors.append(f"{name}: {e}")
            continue
        prev = state.get(url)
        old = set(prev["items"]) if prev else set()
        is_role = ROLE_KEYWORDS.search if ROLE_KEYWORDS else JOB_WORDS.search
        cands = sorted(t for t in items if t not in old and looks_like_title(t) and is_role(t))
        if row["keywords"]:
            cands = [t for t in cands if any(k in t.lower() for k in row["keywords"])]
        cands = drop_merged(cands, items, ctx)
        roles = []
        for t in cands:
            m = next((x for x in (re.fullmatch(r"([A-Z]{2,3})-[A-Z]{2}-.+", seg.strip())
                                  for seg in ctx.get(t, "").split("|")) if x), None)
            status = region_status(f"{t} {ctx.get(t, '')}")
            if m and REGION in CODES and m.group(1) not in CODES[REGION]:  # explicit country code beats guessing
                status = "foreign"
            elif m and REGION in CODES and not PLACES:
                status = "match"
            if REGION != "ALL" and status == "foreign":
                continue  # wrong country: still in state.json, so it never comes back as "new" later
            roles.append({"title": t, "loc": guess_location(ctx.get(t, ""), t), "ok": status})
        if roles:
            findings.append({"company": name, "url": url, "first": prev is None, "emails": emails,
                             "items": roles[:25]})
        state[url] = {"company": name, "items": sorted(items), "checked": now.isoformat()}

    STATE.write_text(json.dumps(state, indent=1))
    if findings or errors:
        history.insert(0, {"checked": now.isoformat(), "findings": findings, "errors": errors})
        history = history[:MAX_RUNS]
        HISTORY.write_text(json.dumps(history, indent=1))
    # always re-render so "Last checked" stays current
    PAGE.write_text(render(history, len(rows), now, errors=errors), encoding="utf-8")

    if not findings and not errors:
        print("No changes.")
        return
    lines = []
    for f in findings:
        lines.append(f"[{'FIRST SCAN' if f['first'] else 'NEW'}] {f['company']} - {f['url']}")
        lines += [f"   - {i['title']}" + (f"  ({i['loc']})" if i.get("loc") else "") for i in f["items"]]
        if f["emails"]:
            lines.append(f"   apply via: {', '.join(f['emails'])}")
    if errors:
        lines += ["", "Errors:"] + [f"   {e}" for e in errors]
    text = "\n".join(lines)
    print(text)
    if any(not f["first"] for f in findings):
        notify(text, render([history[0]], len(rows), now, inline_css=True))


if __name__ == "__main__":
    main()