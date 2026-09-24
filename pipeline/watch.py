#!/usr/bin/env python3
"""
careers_watch.py - watch plain company career pages listed in an Excel sheet
and report job postings that appeared since the last run.

Google Sheet (row 1 = headers), one row per company:
  CompanyName | CareersPage | selector | keywords | active   (only the first two are required)
  selector  (optional) CSS selector around the job list, e.g. "#careers" or ".openings"
            -> cuts noise a lot; find it with right-click > Inspect on the page
  keywords  (optional) comma-separated; only report new items containing one, e.g. "engineer,developer"
  active    (optional) "no" to skip a row

Add rows whenever you find a company; the next run baselines it automatically.

Notifications (optional, via environment variables):
  GMAIL_USER, GMAIL_APP_PASSWORD, NOTIFY_TO  -> email
  SLACK_WEBHOOK                              -> Slack message
A dated HTML report is always written to ./reports/.
"""
import csv, io, json, os, re, smtplib
from datetime import datetime
from email.mime.text import MIMEText
from pathlib import Path

import requests
from bs4 import BeautifulSoup

BASE = Path(__file__).resolve().parent
# From your sheet URL: https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit#gid=<SHEET_GID>
SHEET_ID = os.getenv("SHEET_ID", "1vRYUmUXby4ndyF3gUqdRR73GCCNVpKkaBvKgaL5k_bk")
SHEET_GID = os.getenv("SHEET_GID", "0")
SHEET_URL = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export?format=csv&gid={SHEET_GID}"
STATE = BASE / "state.json"
REPORTS = BASE / "reports"

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
    print(out)
    return out


def fetch_rendered(url):
    """Fallback for JS-rendered pages. Needs: pip install playwright && playwright install chromium"""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return None
    with sync_playwright() as p:
        browser = p.chromium.launch()
        page = browser.new_page(user_agent=HEADERS["User-Agent"])
        page.goto(url, wait_until="networkidle", timeout=45000)
        html = page.content()
        browser.close()
    return html


def fetch(url):
    r = requests.get(url, headers=HEADERS, timeout=25)
    r.raise_for_status()
    html = r.text
    if len(BeautifulSoup(html, "html.parser").get_text(strip=True)) < 400:  # probably JS-rendered
        html = fetch_rendered(url) or html
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


def notify(text, html):
    if os.getenv("SLACK_WEBHOOK"):
        requests.post(os.environ["SLACK_WEBHOOK"], json={"text": text}, timeout=15)
    if os.getenv("GMAIL_USER") and os.getenv("GMAIL_APP_PASSWORD"):
        msg = MIMEText(html, "html")
        msg["Subject"] = f"Career page watch - {datetime.now():%b %d}"
        msg["From"] = os.environ["GMAIL_USER"]
        msg["To"] = os.getenv("NOTIFY_TO", os.environ["GMAIL_USER"])
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
            s.login(os.environ["GMAIL_USER"], os.environ["GMAIL_APP_PASSWORD"])
            s.send_message(msg)


def main():
    rows = load_sheet()
    if not rows:
        return
    state = json.loads(STATE.read_text()) if STATE.exists() else {}
    findings, errors = [], []

    for row in rows:
        url, name = row["url"], row.get("company") or row["url"]
        try:
            # a URL like .../careers#open-roles hints at the job section's id
            selector = row.get("selector") or ("#" + url.split("#", 1)[1] if "#" in url else "")
            items, emails = extract(fetch(url), selector)
        except Exception as e:
            errors.append(f"{name}: {e}")
            continue
        prev = state.get(url)
        old = set(prev["items"]) if prev else set()
        cands = sorted(t for t in items if t not in old and JOB_WORDS.search(t))
        if row["keywords"]:
            cands = [t for t in cands if any(k in t.lower() for k in row["keywords"])]
        if cands:
            findings.append({"company": name, "url": url, "first": prev is None,
                             "items": cands[:25], "emails": emails})
        state[url] = {"company": name, "items": sorted(items), "checked": datetime.now().isoformat()}

    STATE.write_text(json.dumps(state, indent=1))

    if not findings and not errors:
        print("No changes.")
        return
    lines, html = [], ["<h2>Career page watch</h2>"]
    for f in findings:
        tag = "FIRST SCAN" if f["first"] else "NEW"
        lines.append(f"[{tag}] {f['company']} - {f['url']}")
        lines += [f"   - {t}" for t in f["items"]]
        if f["emails"]:
            lines.append(f"   apply via: {', '.join(f['emails'])}")
        html.append(f"<h3>{f['company']} <small>({tag})</small></h3><a href='{f['url']}'>{f['url']}</a><ul>"
                    + "".join(f"<li>{t}</li>" for t in f["items"]) + "</ul>"
                    + (f"<p>Apply via: {', '.join(f['emails'])}</p>" if f["emails"] else ""))
    if errors:
        lines += ["", "Errors:"] + [f"   {e}" for e in errors]
        html.append("<h3>Errors</h3><ul>" + "".join(f"<li>{e}</li>" for e in errors) + "</ul>")

    text, page = "\n".join(lines), "".join(html)
    print(text)
    REPORTS.mkdir(exist_ok=True)
    (REPORTS / f"{datetime.now():%Y-%m-%d_%H%M}.html").write_text(page)
    if findings:
        print(text, page)


if __name__ == "__main__":
    main()