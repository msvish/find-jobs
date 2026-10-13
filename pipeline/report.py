"""report.py - renders history.json into new.html, grouped by company.

The page shows only NEW roles: found after a company's first scan, within the last DAYS_SHOWN days.
First-scan lists and older finds stay in history.json (and the full snapshots in state.json) as the
backlog the nightly diff compares against; they are simply not displayed.

  Company name   Careers page link                 [N new]
    Role title                                      found date
    Location
    Apply by email: ...
"""
from datetime import datetime, timedelta
from html import escape
from pathlib import Path

CSS_FILE = Path(__file__).resolve().parent / "report.css"
DAYS_SHOWN = 2  # new roles older than this drop off the page (they stay in history.json)

JS = """
const q=document.getElementById('q');
q.addEventListener('input',()=>{const t=q.value.trim().toLowerCase();
document.querySelectorAll('.co').forEach(co=>{let any=false;const c=co.dataset.company;
co.querySelectorAll('.roles li').forEach(li=>{const m=!t||c.includes(t)||li.textContent.toLowerCase().includes(t);li.hidden=!m;any=any||m});
co.hidden=!any});});
"""


def _group(history, since):
    """{careers_url: company} holding only new roles (not first-scan) found on or after `since`."""
    companies = {}
    for run in history:  # history is newest first
        found = datetime.fromisoformat(run["checked"])
        if found < since:
            continue
        for f in run["findings"]:
            if f["first"]:
                continue  # baseline for comparison, not news
            co = companies.setdefault(f["url"], {"name": f["company"], "url": f["url"], "emails": f.get("emails", []),
                                                 "latest": found, "roles": [], "seen": set()})
            for i in f["items"]:
                if i["title"] not in co["seen"]:
                    co["seen"].add(i["title"])
                    co["roles"].append({**i, "found": found})
    return sorted(companies.values(), key=lambda c: c["latest"], reverse=True)


def _role(r):
    if r.get("loc"):
        where = f'<span class="loc{"" if r.get("us") == "us" else " unk"}">{escape(r["loc"])}</span>'
    elif r.get("us"):  # checked, but the page gave no location
        where = '<span class="loc unk">Location not listed</span>'
    else:  # older entries saved before locations were captured
        where = ""
    return (f'<li><div class="role"><span class="title">{escape(r["title"])}</span>{where}</div>'
            f'<time datetime="{r["found"].isoformat()}">{r["found"]:%b %-d}</time></li>')


def _company(co):
    n = len(co["roles"])
    emails = ""
    if co["emails"]:
        emails = '<p class="apply">Apply by email: ' + ", ".join(
            f'<a href="mailto:{escape(e)}">{escape(e)}</a>' for e in co["emails"]) + "</p>"
    return (f'<article class="co" data-company="{escape(co["name"].lower())}">'
            f'<div class="co-head"><h2>{escape(co["name"])}<a class="careers" href="{escape(co["url"])}" '
            f'target="_blank" rel="noopener">Careers page</a></h2><span class="tag new">{n} new</span></div>'
            f'<ul class="roles">{"".join(_role(r) for r in co["roles"])}</ul>{emails}</article>')


def render(history, watched, checked_at, inline_css=False, errors=None):
    """inline_css=True embeds the stylesheet (for email, where linked CSS won't load)."""
    style = f"<style>{CSS_FILE.read_text()}</style>" if inline_css else '<link rel="stylesheet" href="report.css">'
    companies = _group(history, checked_at - timedelta(days=DAYS_SHOWN))
    recent = sum(len(co["roles"]) for co in companies)
    headline = f"{recent} new role{'s' if recent != 1 else ''} this week" if recent else "No new roles this week"

    if errors is None:  # no explicit list: fall back to the newest history entry, if it is from today
        errors = history[0].get("errors", []) if history and history[0]["checked"][:10] == checked_at.isoformat()[:10] else []
    errors_html = ""
    if errors:
        n = len(errors)
        errors_html = (f'<details class="errors"><summary>{n} page{"s" if n != 1 else ""} couldn\'t be checked in the latest run</summary>'
                       "<ul>" + "".join(f"<li>{escape(e)}</li>" for e in errors) + "</ul></details>")

    body = "".join(_company(c) for c in companies) or \
        f'<p class="empty">No new roles in the last {DAYS_SHOWN} days. Everything already listed is kept for comparison.</p>'

    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex, nofollow">
<title>{escape(headline)} | Career watch</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Schibsted+Grotesk:wght@400;500;600;700&display=swap" rel="stylesheet">
{style}</head><body><div class="wrap">
<header><h1>{escape(headline)}</h1>
<p class="meta">Watching {watched} career page{'s' if watched != 1 else ''}. Last checked {checked_at:%b %-d at %-I:%M %p} PT.</p>
<div class="tools"><input type="search" id="q" placeholder="Filter by role or company" aria-label="Filter by role or company"></div></header>
{errors_html}<main>{body}</main></div><script>{JS}</script></body></html>"""
