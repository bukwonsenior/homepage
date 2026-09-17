# -*- coding: utf-8 -*-
"""
CMS 공지사항 게시판 → notice.json

담당자는 지금처럼 CMS(느티나무넷) 공지사항에 글을 올린다.
이 스크립트가 그 게시판을 매일 긁어서 최근 글 목록 JSON을 만든다.
홈페이지 공지 카드(notice.html)가 그 JSON에서 최근 글을 보여준다.
"""
import json, re, sys
from datetime import datetime, timezone, timedelta
from pathlib import Path
from urllib.parse import urljoin
import urllib.request

URL = "https://www.bwsenior.or.kr/main/sub.html?pageCode=40"
OUT = Path("notice.json")
KST = timezone(timedelta(hours=9))
MAX = 10   # JSON에 담을 최근 글 수 (화면 표시 개수는 notice.html 의 var MAX 에서 조절)

def fetch(url):
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0 (compatible; BukwonNoticeBot/1.0)",
        "Accept-Language": "ko",
    })
    with urllib.request.urlopen(req, timeout=20) as r:
        raw = r.read()
    for enc in ("utf-8", "euc-kr", "cp949"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "ignore")

def clean_title(t):
    t = re.sub(r"\s+", " ", t).strip()
    t = re.sub(r"\s*새글\s*$", "", t)
    # 첨부파일명 조각(예: xxx.jpg / .hwp)이 끝에 붙는 경우 제거
    t = re.sub(r"\s*\S+\.(jpg|jpeg|png|gif|hwp|hwpx|pdf|zip|xlsx?|docx?)\s*$", "", t, flags=re.I)
    return t.strip()

def parse(html):
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, "html.parser")
    items = []
    for tr in soup.select("table tbody tr"):
        sub = tr.select_one("td.jSubject")
        if not sub:
            continue
        a = sub.find("a")
        title = clean_title(a.get_text(" ", strip=True) if a else sub.get_text(" ", strip=True))
        if not title:
            continue
        cate = (tr.select_one("td.jCate").get_text(strip=True) if tr.select_one("td.jCate") else "")
        dtxt = (tr.select_one("td.jDate").get_text(strip=True) if tr.select_one("td.jDate") else "")
        m = re.search(r"(20\d\d)[.\-/](\d{1,2})[.\-/](\d{1,2})", dtxt)
        date = f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}" if m else ""
        href = a.get("href") if a else None
        url = urljoin(URL, href) if href else URL
        items.append({"구분": cate, "제목": title, "날짜": date, "url": url})
        if len(items) >= MAX:
            break
    return items

def main():
    try:
        html = fetch(URL)
    except Exception as e:
        print(f"::error::공지사항을 불러오지 못했습니다: {e}")
        sys.exit(1)

    items = parse(html)
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
