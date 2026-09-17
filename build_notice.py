# -*- coding: utf-8 -*-
"""
CMS 공지사항 게시판 → notice.json

공지 게시판(anyboard)은 목록을 JavaScript(AJAX)로 불러오므로,
서버 원본 HTML만 받는 방식으로는 글이 안 보인다.
그래서 Playwright(브라우저)를 띄워 실제로 그려진 목록을 읽는다.
"""
import json, re, sys
from datetime import datetime, timezone, timedelta
from pathlib import Path

URL = "https://www.bwsenior.or.kr/main/sub.html?pageCode=40"
OUT = Path("notice.json")
KST = timezone(timedelta(hours=9))
MAX = 10   # JSON에 담을 최근 글 수 (화면 표시 개수는 notice.html 의 var MAX 에서 조절)

EXTRACT = r"""
() => {
  const rows = [...document.querySelectorAll('table tr')].filter(r => r.querySelector('td.jSubject'));
  return rows.map(r => {
    const cate = (r.querySelector('td.jCate') ? r.querySelector('td.jCate').textContent : '').trim();
    const sub  = r.querySelector('td.jSubject');
    const a    = sub.querySelector('a');
    const title = (a ? a.textContent : sub.textContent).replace(/\s+/g, ' ').trim();
    const date  = (r.querySelector('td.jDate') ? r.querySelector('td.jDate').textContent : '').trim();
    const url   = a ? a.href : '';
    return { cate, title, date, url };
  });
}
"""

def clean_title(t):
    t = re.sub(r"\s+", " ", t or "").strip()
    t = re.sub(r"\s*새글\s*$", "", t)
    t = re.sub(r"\s*\S+\.(jpg|jpeg|png|gif|hwp|hwpx|pdf|zip|xlsx?|docx?)\s*$", "", t, flags=re.I)
    return t.strip()

def scrape():
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        browser = pw.chromium.launch()
        page = browser.new_page(locale="ko-KR")
        page.goto(URL, wait_until="networkidle", timeout=40000)
        try:
            page.wait_for_selector("td.jSubject", timeout=15000)
        except Exception:
            browser.close()
            return []
        rows = page.evaluate(EXTRACT)
        browser.close()
    return rows or []

def main():
    try:
        rows = scrape()
    except Exception as e:
        print(f"::error::공지사항을 불러오지 못했습니다: {e}")
        sys.exit(1)

    items = []
    for r in rows:
        title = clean_title(r.get("title"))
        if not title:
            continue
        dtxt = r.get("date", "")
        m = re.search(r"(20\d\d)[.\-/](\d{1,2})[.\-/](\d{1,2})", dtxt)
        date = f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}" if m else ""
        items.append({
            "구분": (r.get("cate") or "").strip(),
            "제목": title,
            "날짜": date,
            "url": r.get("url") or URL,
        })
        if len(items) >= MAX:
            break

    if not items:
        print("::error::공지 글을 한 건도 읽지 못했습니다. 페이지 구조가 바뀌었을 수 있습니다. (기존 notice.json 유지)")
        sys.exit(1)

    doc = {
        "생성시각": datetime.now(KST).isoformat(timespec="seconds"),
        "출처": URL,
        "전체보기": URL,
        "items": items,
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"✓ {OUT} 생성 — 공지 {len(items)}건")

if __name__ == "__main__":
    main()
