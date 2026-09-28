# -*- coding: utf-8 -*-
"""
시설안내 엑셀(북원_시설안내.xlsx) → facility.json

엑셀 '시설안내' 시트 구조:
  1~3행: 안내문 (무시)
  4행: 헤더 → 게시 | 층코드 | 층명 | 실명 | 사진 | 역할 | 이용방법
  5행~: 데이터

사진 파일은 img/facility/ 폴더에 넣고, 엑셀에는 파일명만 적는다.
(예: f1_office.jpg → img/facility/f1_office.jpg)
"""
import json, re, sys
from pathlib import Path
import openpyxl

XLSX = Path("북원_시설안내.xlsx")
OUT  = Path("facility.json")
IMG_BASE = "img/facility/"

errors = []
def err(row, msg): errors.append(f"{row}행: {msg}")
def s(v): return "" if v is None else str(v).strip()

JUMIN = re.compile(r"\b\d{6}[-]\d{7}\b")

def main():
    if not XLSX.exists():
        print(f"::error::{XLSX} 파일이 없습니다.")
        sys.exit(1)

    wb = openpyxl.load_workbook(XLSX, data_only=True)
    if "시설안내" not in wb.sheetnames:
        print("::error::'시설안내' 시트가 없습니다.")
        sys.exit(1)
    ws = wb["시설안내"]

    head = [s(c.value) for c in ws[4]]
    need = ["게시", "층코드", "층명", "실명", "사진", "역할", "이용방법"]
    for h in need:
        if h not in head:
            errors.append(f"4행에 '{h}' 열이 없습니다.")
    if errors:
        for e in errors: print("  ✗ " + e)
        sys.exit(1)
    ix = {h: head.index(h) for h in need}

    # 데이터 파싱
    floors = {}  # 층코드 → { badge, name, rooms: [] }
    floor_order = []

    for r in range(5, ws.max_row + 1):
        row = [c.value for c in ws[r]]
        if not any(s(v) for v in row): continue

        pub = s(row[ix["게시"]]).upper()
        if pub in ("X", "×"): continue
        if pub != "O":
            err(r, f"게시 칸은 O 또는 X 만 (지금: '{s(row[ix['게시']])}')")
            continue

        code = s(row[ix["층코드"]])
        fname = s(row[ix["층명"]])
        rname = s(row[ix["실명"]])
        photo = s(row[ix["사진"]])
        role  = s(row[ix["역할"]])
        usage = s(row[ix["이용방법"]])

        if not code: err(r, "층코드가 비어 있습니다"); continue
        if not rname: err(r, "실명이 비어 있습니다"); continue

        # 개인정보 검사
        line = " ".join(s(v) for v in row)
        if JUMIN.search(line):
            err(r, "주민등록번호로 보이는 값이 있습니다")

        # 사진 경로 처리
        photo_url = ""
        if photo:
            if photo.startswith("http"):
                photo_url = photo
            else:
                photo_url = IMG_BASE + photo

        if code not in floors:
            # 층코드에서 배지 생성
            badge = fname if fname else code
            if code.startswith("f") and code[1:].isdigit():
                badge = code[1:] + "F"
            floors[code] = {"badge": badge, "name": fname, "rooms": []}
            floor_order.append(code)

        floors[code]["rooms"].append({
            "name": rname,
            "photo": photo_url,
            "role": role,
            "usage": usage
        })

    if errors:
        print("\n엑셀에서 고쳐야 할 곳이 있습니다.\n")
        for e in errors: print("  ✗ " + e)
        print(f"::error::엑셀 오류 {len(errors)}건")
        sys.exit(1)

    # JSON 출력
    result = []
    for code in floor_order:
        fl = floors[code]
        result.append({
            "code": code,
            "badge": fl["badge"],
            "name": fl["name"],
            "rooms": fl["rooms"]
        })

    OUT.write_text(json.dumps(result, ensure_ascii=False, indent=1), encoding="utf-8")

    total_rooms = sum(len(fl["rooms"]) for fl in result)
    print(f"✓ {OUT} 생성 — {len(result)}개 층, {total_rooms}개 실")

if __name__ == "__main__":
    main()
