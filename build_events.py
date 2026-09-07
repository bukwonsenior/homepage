# -*- coding: utf-8 -*-
"""
CMS 이달의 행사 페이지 → events.json

담당자는 지금처럼 CMS(느티나무넷) 행사달력에 일정을 입력한다.
이 스크립트가 그 페이지를 매일 긁어서 날짜별 행사 JSON을 만든다.
홈페이지 행사 카드(events.html)가 그 JSON에서 '다가오는 행사'를 골라 보여준다.
"""
import json, re, sys
from datetime import datetime, date, timezone, timedelta
from pathlib import Path
import urllib.request

URL = "https://www.bwsenior.or.kr/main/sub.html?pageCode=12"
OUT = Path("events.json")
KST = timezone(timedelta(hours=9))

def fetch(url):
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0 (compatible; BukwonEventBot/1.0)",
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

def parse(html):
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, "html.parser")

    # 연/월: 헤더의 "YYYY.MM" (문서 내 유일)
    m = re.search(r"(20\d\d)\.(\d{2})", html)
    if not m:
        return {}
    year, month = int(m.group(1)), int(m.group(2))

    DAYNUM = re.compile(r"^\d{1,2}$")
    days = {}

    for cell in soup.select("div.mdDayBox"):
        # 텍스트 노드들을 순서대로: 날짜 숫자는 반복되고, 제목은 <a>/<i> 안에 있음
        toks = [t.strip() for t in cell.get_text("\n", strip=True).split("\n") if t.strip()]
        day = next((t for t in toks if DAYNUM.match(t)), None)
        if not day:
            continue  # 앞뒤 달의 빈 칸
        d = int(day)
        if not (1 <= d <= 31):
            continue
        # 제목: 순수 숫자가 아닌 토큰 (중복 제거, 순서 유지)
        titles = []
        for t in toks:
            if DAYNUM.match(t):
                continue
            if t not in titles:
                titles.append(t)
        if not titles:
            continue
        try:
            key = date(year, month, d).isoformat()
        except ValueError:
            continue
        # 같은 날짜가 여러 캘린더(PC/모바일)에 중복될 수 있으니 병합
        merged = days.setdefault(key, [])
        for t in titles:
            if t not in merged:
                merged.append(t)

    return days

def main():
    try:
        html = fetch(URL)
    except Exception as e:
        print(f"::error::행사 페이지를 불러오지 못했습니다: {e}")
        sys.exit(1)

    days = parse(html)
    if not days:
        print("::error::행사를 한 건도 읽지 못했습니다. 페이지 구조가 바뀌었을 수 있습니다. (기존 events.json 유지)")
        sys.exit(1)

    doc = {
        "생성시각": datetime.now(KST).isoformat(timespec="seconds"),
        "출처": URL,
        "달력주소": URL,
        "days": dict(sorted(days.items())),
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8")
    total = sum(len(v) for v in days.values())
    print(f"✓ {OUT} 생성 — {len(days)}일 {total}건 ({min(days)} ~ {max(days)})")

if __name__ == "__main__":
    main()
