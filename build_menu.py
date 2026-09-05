# -*- coding: utf-8 -*-
"""
CMS 이달의 식단 페이지 → menu.json

담당자는 지금처럼 CMS(느티나무넷)에 식단을 입력한다.
이 스크립트가 그 페이지를 매일 긁어서 날짜별 메뉴 JSON을 만든다.
홈페이지 식단 카드(embed/meal.html)가 그 JSON에서 '오늘 메뉴'를 골라 보여준다.
"""
import json, re, sys
from datetime import datetime, date, timezone, timedelta
from pathlib import Path
import urllib.request

URL  = "https://www.bwsenior.or.kr/main/sub.html?pageCode=13"
OUT  = Path("menu.json")
KST  = timezone(timedelta(hours=9))

def fetch(url):
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0 (compatible; BukwonMenuBot/1.0)",
        "Accept-Language": "ko",
    })
    with urllib.request.urlopen(req, timeout=20) as r:
        raw = r.read()
    # 인코딩 자동 판별 (EUC-KR 대비)
    for enc in ("utf-8", "euc-kr", "cp949"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", "ignore")

def parse(html):
    from bs4 import BeautifulSoup
    soup = BeautifulSoup(html, "html.parser")
    today = datetime.now(KST).date()
    days = {}

    for wl in soup.select("div.weekLine"):
        # 날짜: boxTop 안의 p 중 'MM.DD' 형식만
        dates = []
        for p in wl.select("div.boxTop p"):
            t = p.get_text(strip=True)
            m = re.match(r"(\d{2})\.(\d{2})", t)
            dates.append((int(m.group(1)), int(m.group(2))) if m else None)
        dates = [d for d in dates if d]

        # 메뉴: 중식 boxMenu (첫 번째). 라벨/빈칸 제외한 셀들
        boxmenus = wl.select("div.boxMenu")
        if not boxmenus:
            continue
        bm = boxmenus[0]
        cells = []
        for c in bm.find_all(recursive=False):
            txt = c.get_text("\n", strip=True)
            if txt and txt not in ("중식", "석식"):
                cells.append(txt)

        for (mm, dd), cell in zip(dates, cells):
            year = today.year
            if mm == 1 and today.month == 12:
                year += 1
            elif mm == 12 and today.month == 1:
                year -= 1
            items = [x.strip() for x in cell.split("\n") if x.strip()]
            if items:
                days[date(year, mm, dd).isoformat()] = items

    return days

def main():
    try:
        html = fetch(URL)
    except Exception as e:
        print(f"::error::식단 페이지를 불러오지 못했습니다: {e}")
        sys.exit(1)

    days = parse(html)
    if not days:
        print("::error::식단을 한 건도 읽지 못했습니다. 페이지 구조가 바뀌었을 수 있습니다. (기존 menu.json 유지)")
        sys.exit(1)

    doc = {
        "생성시각": datetime.now(KST).isoformat(timespec="seconds"),
        "출처": URL,
        "식단표주소": URL,
        "days": dict(sorted(days.items())),
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"✓ {OUT} 생성 — {len(days)}일치 식단 ({min(days)} ~ {max(days)})")

if __name__ == "__main__":
    main()
