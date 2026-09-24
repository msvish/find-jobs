"""report.py - renders history.json into new.html, grouped by company.

Page structure:
  Company name   Careers page link                 [N new]
    Role title                              found date
    Role title ...
    Apply by email: ...
Companies with the most recent finds are listed first; roles are newest first.
"""
from datetime import datetime, timedelta
from html import escape
from pathlib import Path

CSS_FILE = Path(__file__).resolve().parent / "report.css"

JS = """
const q=document.getElementById('q'),hf=document.getElementById('hf');
function apply(){const t=q.value.trim().toLowerCase();
document.querySelectorAll('.co').forEach(co=>{let any=false;const c=co.dataset.company;
co.querySelectorAll('.roles li').forEach(li=>{
const m=(!t||c.includes(t)||li.textContent.toLowerCase().includes(t))&&!(hf.checked&&li.classList.contains('first'));
li.hidden=!m;any=any||m});co.hidden=!any});}
q.addEventListener('input',apply);hf.addEventListener('change',apply);
"""


def _group(history):
    """Turn run-by-run history into {careers_url: company dict}, newest roles first."""
    companies = {}
    for run in history:  # history is newest first
        found = datetime.fromisoformat(run["checked"])
        for f in run["findings"]:
            co = companies.setdefault(f["url"], {"name": f["company"], "url": f["url"], "emails": f.get("emails", []),
                                                 "latest": found, "roles": [], "seen": set()})
            for i in f["items"]:
                if i["title"] not in co["seen"]:
                    co["seen"].add(i["title"])
                    co["roles"].append({**i, "found": found, "first": f["first"]})
    return sorted(companies.values(), key=lambda c: c["latest"], reverse=True)


def _company(co, week_ago):
    fresh = sum(1 for r in co["roles"] if not r["first"] and r["found"] >= week_ago)
    badge = f'<span class="tag new">{fresh} new this week</span>' if fresh else ""
    roles = "".join(
        f'<li class="{"first" if r["first"] else "new"}"><span>{escape(r["title"])}</span><time datetime="{r["found"].isoformat()}">'
        f'{"Listed " if r["first"] else ""}{r["found"]:%b %-d}</time></li>'
        for r in co["roles"])
    emails = ""
    if co["emails"]:
        emails = '<p class="apply">Apply by email: ' + ", ".join(
            f'<a href="mailto:{escape(e)}">{escape(e)}</a>' for e in co["emails"]) + "</p>"
    return (f'<article class="co" data-company="{escape(co["name"].lower())}">'
            f'<div class="co-head"><h2>{escape(co["name"])}<a class="careers" href="{escape(co["url"])}" '
            f'target="_blank" rel="noopener">Careers page</a></h2>{badge}</div><ul class="roles">{roles}</ul>{emails}</article>')


def render(history, watched, checked_at, inline_css=False):
    """inline_css=True embeds the stylesheet (for email, where linked CSS won't load)."""
    style = f"<style>{CSS_FILE.read_text()}</style>" if inline_css else '<link rel="stylesheet" href="report.css">'
    week_ago = checked_at - timedelta(days=7)
    companies = _group(history)
    recent = sum(1 for co in companies for r in co["roles"] if not r["first"] and r["found"] >= week_ago)
    headline = f"{recent} new role{'s' if recent != 1 else ''} this week" if recent else "No new roles this week"

    errors = ""
    if history and history[0].get("errors") and history[0]["checked"][:10] == checked_at.isoformat()[:10]:
        n = len(history[0]["errors"])
        errors = (f'<details class="errors"><summary>{n} page{"s" if n != 1 else ""} couldn\'t be checked today</summary>'
                  "<ul>" + "".join(f"<li>{escape(e)}</li>" for e in history[0]["errors"]) + "</ul></details>")

    body = "".join(_company(c, week_ago) for c in companies) or \
        '<p class="empty">Nothing yet. Roles appear here after the next check.</p>'

    return f"""<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{escape(headline)} | Career watch</title>
<link rel="preconnect" href="https://fonts.googleapis.com"><link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Schibsted+Grotesk:wght@400;500;600;700&display=swap" rel="stylesheet">
{style}</head><body><div class="wrap">
<header><h1>{escape(headline)}</h1>
<p class="meta">Watching {watched} career page{'s' if watched != 1 else ''}. Last checked {checked_at:%b %-d at %-I:%M %p} ET.</p>
<div class="tools"><input type="search" id="q" placeholder="Filter by role or company" aria-label="Filter by role or company">
<label><input type="checkbox" id="hf"> Hide first scans</label></div></header>
{errors}<main>{body}</main></div><script>{JS}</script></body></html>"""