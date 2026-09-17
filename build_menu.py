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
    bongsa = {}   # 날짜별 오늘의 봉사단

    for wl in soup.select("div.weekLine"):
        # boxTop 열 순서를 그대로(빈칸 포함) 읽어 열 번호를 맞춘다
        tops = [p.get_text(strip=True) for p in wl.select("div.boxTop p")]

        boxmenus = wl.select("div.boxMenu")
        if not boxmenus:
            continue
        lunch = boxmenus[0]          # 중식(첫 번째 boxMenu)
        vol = None                   # 봉사단('봉사단' 라벨로 시작하는 boxMenu)
        for bm in boxmenus:
            kids = bm.find_all(recursive=False)
            if kids and kids[0].get_text(strip=True) == "봉사단":
                vol = bm
                break
        lunch_cells = lunch.find_all(recursive=False)
        vol_cells = vol.find_all(recursive=False) if vol else []

        for j, t in enumerate(tops):
            m = re.match(r"(\d{2})\.(\d{2})", t)
            if not m:
                continue
            mm, dd = int(m.group(1)), int(m.group(2))
            year = today.year
            if mm == 1 and today.month == 12:
                year += 1
            elif mm == 12 and today.month == 1:
                year -= 1
            iso = date(year, mm, dd).isoformat()

            # 중식(같은 열 번호)
            if j < len(lunch_cells):
                cell = lunch_cells[j].get_text("\n", strip=True)
                items = [x.strip() for x in cell.split("\n") if x.strip() and x.strip() not in ("중식", "석식")]
                if items:
                    days[iso] = items

            # 봉사단(같은 열 번호)
            if j < len(vol_cells):
                name = vol_cells[j].get_text(" ", strip=True)
                if name and name != "봉사단":
                    bongsa[iso] = name

    return days, bongsa

def main():
    try:
        html = fetch(URL)
    except Exception as e:
        print(f"::error::식단 페이지를 불러오지 못했습니다: {e}")
        sys.exit(1)

    days, bongsa = parse(html)
    if not days:
        print("::error::식단을 한 건도 읽지 못했습니다. 페이지 구조가 바뀌었을 수 있습니다. (기존 menu.json 유지)")
        sys.exit(1)

    doc = {
        "생성시각": datetime.now(KST).isoformat(timespec="seconds"),
        "출처": URL,
        "식단표주소": URL,
        "days": dict(sorted(days.items())),
        "bongsa": dict(sorted(bongsa.items())),
    }
    OUT.parent.mkdir(parents=True, exist_ok=True)
    OUT.write_text(json.dumps(doc, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"✓ {OUT} 생성 — 식단 {len(days)}일 / 봉사단 {len(bongsa)}일 ({min(days)} ~ {max(days)})")

if __name__ == "__main__":
    main()
