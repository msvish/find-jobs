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
STATE = BASE / "state.json"
HISTORY = BASE / "history.json"
PAGE = BASE / "new.html"
TZ = ZoneInfo("America/New_York")
MAX_RUNS = 200  # runs kept in history.json

HEADERS = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
           "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126 Safari/537.36"}
JOB_WORDS = re.compile(r"\b(engineer|developer|programmer|intern(ship)?|analyst|scientist|architect|"
                       r"devops|sre|full[- ]?stack|front[- ]?end|back[- ]?end|software|opening|"
                       r"position|hiring|vacanc(y|ies))\b", re.I)
NOISE = re.compile(r"(cookie|privacy|terms of|copyright|©|subscribe|newsletter|log ?in|sign ?up)", re.I)
EMAIL = re.compile(r"[\w.+-]+@[\w-]+\.[a-z]{2,}(?:\.[a-z]{2,})?", re.I)


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
    items = set()
    for root in roots:
        for el in root.find_all(["a", "h1", "h2", "h3", "h4", "h5", "li", "p", "td", "strong"]):
            text = " ".join(el.get_text(" ", strip=True).split())
            if 4 <= len(text) <= 160 and not NOISE.search(text):
                items.add(text)
    return items, sorted(e for e in emails if not e.lower().endswith((".png", ".jpg", ".svg")))


def looks_like_title(t):
    """Job titles are short and don't end like a sentence; descriptions do."""
    return len(t.split()) <= 8 and not t.rstrip().endswith((".", "!", "?"))


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
            items, emails = extract(fetch(url, force_render), selector)
        except Exception as e:
            errors.append(f"{name}: {e}")
            continue
        prev = state.get(url)
        old = set(prev["items"]) if prev else set()
        cands = sorted(t for t in items if t not in old and looks_like_title(t) and JOB_WORDS.search(t))
        if row["keywords"]:
            cands = [t for t in cands if any(k in t.lower() for k in row["keywords"])]
        if cands:
            findings.append({"company": name, "url": url, "first": prev is None, "emails": emails,
                             "items": [{"title": t} for t in cands[:25]]})
        state[url] = {"company": name, "items": sorted(items), "checked": now.isoformat()}

    STATE.write_text(json.dumps(state, indent=1))
    if findings or errors:
        history.insert(0, {"checked": now.isoformat(), "findings": findings, "errors": errors})
        history = history[:MAX_RUNS]
        HISTORY.write_text(json.dumps(history, indent=1))
    # always re-render so "Last checked" stays current
    PAGE.write_text(render(history, len(rows), now), encoding="utf-8")

    if not findings and not errors:
        print("No changes.")
        return
    lines = []
    for f in findings:
        lines.append(f"[{'FIRST SCAN' if f['first'] else 'NEW'}] {f['company']} - {f['url']}")
        lines += [f"   - {i['title']}" for i in f["items"]]
        if f["emails"]:
            lines.append(f"   apply via: {', '.join(f['emails'])}")
    if errors:
        lines += ["", "Errors:"] + [f"   {e}" for e in errors]
    text = "\n".join(lines)


if __name__ == "__main__":
    main()