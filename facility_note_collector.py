"""
facility_note_collector.py

Daily collector for Japanese store/facility opening announcements, producing:
  1. A note.com-ready Markdown article (news-summary style)
  2. An appended row-set in an Excel tracking file
  3. An emailed copy of the Markdown article

Designed to be added alongside the existing collector.py in this repo and run
daily via GitHub Actions (see .github/workflows/facility_note_collector.yml).

ASSUMPTIONS (adjust to match your actual repo layout):
  - Excel tracker lives at: data/store_openings.xlsx (created if missing)
  - Markdown output goes to: output/note_YYYY-MM-DD.md
  - Email credentials/config come from environment variables (GitHub Secrets):
      SMTP_SERVER, SMTP_PORT, SMTP_USER, SMTP_PASS, MAIL_TO
  - Timezone: Asia/Tokyo (JST) is used for "today's date"

This script does NOT use an LLM. It is a best-effort structured scraper, so
its judgment (what counts as "genuinely today", filtering noise, etc.) is
much simpler than the interactive Claude skill. Treat its output as a first
draft to skim, not a guaranteed-accurate final article.
"""

import os
import re
import smtplib
from datetime import datetime, timezone, timedelta
from email.mime.text import MIMEText
from pathlib import Path

import requests
from bs4 import BeautifulSoup

try:
    from openpyxl import Workbook, load_workbook
except ImportError:
    raise SystemExit("Please add openpyxl to requirements.txt / pip install openpyxl")

JST = timezone(timedelta(hours=9))
TODAY = datetime.now(JST).strftime("%Y%m%d")
TODAY_HUMAN = datetime.now(JST).strftime("%Y年%m月%d日")

EXCEL_PATH = Path("data/store_openings.xlsx")
OUTPUT_DIR = Path("output")
OUTPUT_DIR.mkdir(exist_ok=True)
MD_PATH = OUTPUT_DIR / f"note_{datetime.now(JST).strftime('%Y-%m-%d')}.md"

HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; FacilityNoteCollector/1.0)"}


def fetch_shokusaihinkan():
    """
    Fetch today's roundup from 食彩品館NEWS, which publishes a same-day post
    at a predictable URL: https://shokusaihinkannews.com/YYYYMMDDopen/
    Returns a list of dicts: {name, url, region}
    """
    url = f"https://shokusaihinkannews.com/{TODAY}open/"
    items = []
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        if resp.status_code != 200:
            return items
        soup = BeautifulSoup(resp.text, "html.parser")
        # The per-item entries sit under headings like "### ◇東京" followed by
        # "・9/1[店舗名](URL)" lines. We look for all <a> tags whose surrounding
        # text starts with a date pattern like "9/1".
        for a in soup.find_all("a", href=True):
            text = a.get_text(strip=True)
            href = a["href"]
            prev_text = a.find_previous(string=True) or ""
            context = (str(prev_text) + text)
            if re.search(r"\d{1,2}/\d{1,2}", context) and "shokusaihinkannews.com" not in href:
                if text and len(text) > 1:
                    items.append({"name": text, "url": href, "source": "食彩品館NEWS"})
    except requests.RequestException:
        pass
    # de-dupe by url
    seen = set()
    deduped = []
    for it in items:
        if it["url"] not in seen:
            seen.add(it["url"])
            deduped.append(it)
    return deduped[:40]  # safety cap


def fetch_prtimes_keyword(keyword="新規開店"):
    """
    Fetch a PR TIMES keyword page and pull out headline links.
    Best-effort only; PR TIMES may rate-limit or change markup.
    """
    url = f"https://prtimes.jp/topics/keywords/{requests.utils.quote(keyword)}"
    items = []
    try:
        resp = requests.get(url, headers=HEADERS, timeout=15)
        if resp.status_code != 200:
            return items
        soup = BeautifulSoup(resp.text, "html.parser")
        for a in soup.select("a[href*='/main/html/rd/p/']"):
            title = a.get_text(strip=True)
            if title:
                items.append({"name": title, "url": a["href"], "source": "PR TIMES"})
    except requests.RequestException:
        pass
    return items[:20]


def build_markdown(shokusai_items, prtimes_items):
    lines = [f"# 【{TODAY_HUMAN}】新規オープン情報まとめ（自動収集版）", ""]
    lines.append("以下は自動収集された候補一覧です。手動での事実確認・重複排除前の下書きとして扱ってください。")
    lines.append("")
    lines.append("## 食彩品館NEWS 経由")
    if shokusai_items:
        for it in shokusai_items:
            lines.append(f"- [{it['name']}]({it['url']})")
    else:
        lines.append("- （本日は取得できませんでした）")
    lines.append("")
    lines.append("## PR TIMES「新規開店」経由")
    if prtimes_items:
        for it in prtimes_items:
            lines.append(f"- [{it['name']}]({it['url']})")
    else:
        lines.append("- （本日は取得できませんでした）")
    lines.append("")
    lines.append("---")
    lines.append(f"自動生成日時（JST）: {datetime.now(JST).strftime('%Y-%m-%d %H:%M')}")
    return "\n".join(lines)


def append_to_excel(shokusai_items, prtimes_items):
    if EXCEL_PATH.exists():
        wb = load_workbook(EXCEL_PATH)
        ws = wb.active
    else:
        EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)
        wb = Workbook()
        ws = wb.active
        ws.append(["取得日", "ソース", "店舗名/施設名", "URL"])

    for it in shokusai_items + prtimes_items:
        ws.append([TODAY_HUMAN, it["source"], it["name"], it["url"]])

    wb.save(EXCEL_PATH)


def send_email(markdown_text):
    server = os.environ.get("SMTP_SERVER")
    port = os.environ.get("SMTP_PORT", "587")
    user = os.environ.get("SMTP_USER")
    password = os.environ.get("SMTP_PASS")
    mail_to = os.environ.get("MAIL_TO")

    if not all([server, user, password, mail_to]):
        print("Email env vars not fully set; skipping email step.")
        return

    msg = MIMEText(markdown_text, "plain", "utf-8")
    msg["Subject"] = f"【{TODAY_HUMAN}】新規オープン情報まとめ（自動収集）"
    msg["From"] = user
    msg["To"] = mail_to

    with smtplib.SMTP(server, int(port)) as smtp:
        smtp.starttls()
        smtp.login(user, password)
        smtp.send_message(msg)
    print("Email sent.")


def main():
    shokusai_items = fetch_shokusaihinkan()
    prtimes_items = fetch_prtimes_keyword("新規開店")

    markdown_text = build_markdown(shokusai_items, prtimes_items)
    MD_PATH.write_text(markdown_text, encoding="utf-8")
    print(f"Wrote {MD_PATH}")

    append_to_excel(shokusai_items, prtimes_items)
    print(f"Updated {EXCEL_PATH}")

    send_email(markdown_text)


if __name__ == "__main__":
    main()
