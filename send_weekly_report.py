"""
Weekly job: look up last week's positions, add them as a new tab in the master
workbook, and email the workbook with a short HTML summary.

    python send_weekly_report.py                 # run + email
    python send_weekly_report.py --dry-run       # run, write report_preview.html, no email

Email settings come from environment variables (GitHub secrets in CI):
    SMTP_HOST, SMTP_PORT (587 or 465), SMTP_USER, SMTP_PASSWORD, MAIL_FROM (optional)
    MAIL_TO   comma-separated recipients
"""

from __future__ import annotations

import argparse
import os
import smtplib
import ssl
import subprocess
import sys
from email.message import EmailMessage
from email.utils import formataddr
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
from keyword_positions import format_range, last_week_range, load_keys_file, week_tab_name  # noqa: E402

load_keys_file(HERE / "keys.env")

BRAND = "#1F3864"
MASTER = HERE / "Fortinet_Keyword_Positions.xlsx"
INPUT = HERE / "Keyword_input.csv"


def run_lookup(start) -> Path:
    weekly = HERE / f"positions_{start:%Y%m%d}.xlsx"
    cmd = [sys.executable, str(HERE / "keyword_positions.py"), "--input", str(INPUT),
           "--week-start", start.isoformat(), "--output", str(weekly), "--master", str(MASTER)]
    print("$", " ".join(cmd), flush=True)
    if subprocess.run(cmd).returncode != 0:
        raise SystemExit("Lookup failed")
    return weekly


def movers(this_w: pd.DataFrame, last_w: pd.DataFrame | None, n: int = 5):
    """Biggest GSC-to-GSC moves (apples to apples only)."""
    if last_w is None:
        return [], []
    a = this_w[this_w["Source"] == "GSC"][["Keyword", "Page URL", "Average Position"]]
    b = last_w[last_w["Source"] == "GSC"][["Keyword", "Page URL", "Average Position"]]
    m = a.merge(b, on=["Keyword", "Page URL"], suffixes=("", "_prev"))
    m["chg"] = m["Average Position_prev"] - m["Average Position"]
    up = m[m.chg > 0].nlargest(n, "chg").to_dict("records")
    down = m[m.chg < 0].nsmallest(n, "chg").to_dict("records")
    return up, down


def build_html(start, end, tab: str) -> str:
    book = pd.read_excel(MASTER, sheet_name=None)
    w = book[tab]
    week_tabs = [k for k in book if k not in ("Summary", "Trend")]
    idx = week_tabs.index(tab)
    prev = book[week_tabs[idx + 1]] if idx + 1 < len(week_tabs) else None  # tabs are newest first

    src = w["Source"].astype(str)
    pos = pd.to_numeric(w["Average Position"], errors="coerce")
    gsc_avg = pd.to_numeric(w.loc[src == "GSC", "Average Position"], errors="coerce").mean()
    stats = [("Keywords", len(w)), ("Avg position (GSC)", f"{gsc_avg:.1f}" if pd.notna(gsc_avg) else "n/a"),
             ("In top 10", int((pos <= 10).sum())), ("Not ranking", int((src == "Not ranking").sum()))]
    tiles = "".join(
        f'<td style="padding:12px 16px;background:#F5F7FA;border-radius:6px">'
        f'<div style="font-size:11px;color:#667;text-transform:uppercase">{k}</div>'
        f'<div style="font-size:21px;font-weight:700;color:{BRAND}">{v}</div></td>' for k, v in stats)
    src_line = " · ".join(f"{s}: {int((src == s).sum())}" for s in ("GSC", "Ahrefs", "Semrush", "Not ranking"))

    up, down = movers(w, prev)

    def table(rows, color, title):
        if not rows:
            return ""
        trs = "".join(
            f'<tr><td style="padding:6px 10px">{r["Keyword"]}</td>'
            f'<td style="padding:6px 10px;text-align:right">{r["Average Position_prev"]:.1f}</td>'
            f'<td style="padding:6px 10px;text-align:right">{r["Average Position"]:.1f}</td>'
            f'<td style="padding:6px 10px;text-align:right;color:{color};font-weight:600">'
            f'{r["chg"]:+.1f}</td></tr>' for r in rows)
        return (f'<h3 style="margin:22px 0 6px;font-size:15px;color:{BRAND}">{title}</h3>'
                f'<table cellspacing="0" style="border-collapse:collapse;font-size:13px;width:100%">'
                f'<tr style="background:#EEF1F6;text-align:left"><th style="padding:6px 10px">Keyword</th>'
                f'<th style="padding:6px 10px;text-align:right">Last week</th>'
                f'<th style="padding:6px 10px;text-align:right">This week</th>'
                f'<th style="padding:6px 10px;text-align:right">Change</th></tr>{trs}</table>')

    return f"""<!DOCTYPE html><html><body style="margin:0;padding:22px;font-family:Segoe UI,Arial,sans-serif;color:#222">
<div style="max-width:760px;margin:0 auto">
<h2 style="margin:0 0 4px;color:{BRAND};font-size:20px">Fortinet weekly keyword positions</h2>
<p style="margin:0 0 16px;color:#667;font-size:14px">{format_range(start, end)}</p>
<table cellspacing="8" cellpadding="0"><tr>{tiles}</tr></table>
<p style="font-size:13px;color:#555;margin:8px 0 0">Sources: {src_line}</p>
{table(up, "#1E7B34", "Biggest gains (GSC, week over week)")}
{table(down, "#C00000", "Biggest drops (GSC, week over week)")}
<p style="margin:22px 0 0;font-size:13px;color:#555">The attached workbook has a tab for every week
(this week is <b>{tab}</b>), a <b>Trend</b> tab with each keyword across weeks, and a <b>Summary</b> tab.
GSC rows are weekly averages; Ahrefs and Semrush rows are point-in-time ranks, so compare like with like.</p>
<p style="margin:18px 0 0;color:#99a;font-size:12px;border-top:1px solid #e5e5e5;padding-top:10px">
Generated automatically every Tuesday.</p></div></body></html>"""


def send(html: str, start, end) -> None:
    env = lambda n: (os.environ.get(n) or "").strip()  # secrets often carry a trailing newline
    host, user, pw = env("SMTP_HOST"), env("SMTP_USER"), env("SMTP_PASSWORD").replace(" ", "")
    port = int(env("SMTP_PORT") or 587)
    to = [a.strip() for a in os.environ.get("MAIL_TO", "").split(",") if a.strip()]
    missing = [n for n, v in (("SMTP_HOST", host), ("SMTP_USER", user), ("SMTP_PASSWORD", pw), ("MAIL_TO", to)) if not v]
    if missing:
        raise SystemExit(f"Cannot send email, missing: {', '.join(missing)}")
    msg = EmailMessage()
    msg["Subject"] = f"Fortinet keyword positions: {format_range(start, end)}"
    msg["From"] = formataddr(("LeadWalnut SEO Reports", env("MAIL_FROM") or user))
    msg["To"] = ", ".join(to)
    msg.set_content("This report is HTML. The attached workbook has the full detail.")
    msg.add_alternative(html, subtype="html")
    msg.add_attachment(MASTER.read_bytes(), maintype="application",
                       subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       filename=MASTER.name)
    ctx = ssl.create_default_context()
    if port == 465:
        with smtplib.SMTP_SSL(host, port, context=ctx, timeout=60) as s:
            s.login(user, pw); s.send_message(msg)
    else:
        with smtplib.SMTP(host, port, timeout=60) as s:
            s.starttls(context=ctx); s.login(user, pw); s.send_message(msg)
    print("Sent to", ", ".join(to))


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true", help="No email; write report_preview.html")
    ap.add_argument("--week-start", help="YYYY-MM-DD (default: last complete Sun-Sat week)")
    a = ap.parse_args()
    if a.week_start:
        from datetime import date, timedelta
        start = date.fromisoformat(a.week_start); end = start + timedelta(days=6)
    else:
        start, end = last_week_range()
    print(f"Reporting week {start} to {end}", flush=True)
    run_lookup(start)
    html = build_html(start, end, week_tab_name(start, end))
    if a.dry_run:
        (HERE / "report_preview.html").write_text(html, encoding="utf-8")
        print("Dry run: wrote report_preview.html, no email sent")
        return 0
    send(html, start, end)
    return 0


if __name__ == "__main__":
    sys.exit(main())
