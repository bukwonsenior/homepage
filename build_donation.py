# -*- coding: utf-8 -*-
"""
data/후원.xlsx -> data/donation.json
시트 '후원', 4행 머리글(게시/종류/이름·라벨/내용/아이콘), 5행부터.
종류: 소개 / 계좌 / 방법 / 종류 / 혜택 / 문의 / 안내
"""
import json, sys
from datetime import datetime, timezone, timedelta
from pathlib import Path
import openpyxl

SRC=Path("data/후원.xlsx"); OUT=Path("data/donation.json")
KST=timezone(timedelta(hours=9)); SHEET="후원"
KINDS={"소개","계좌","예금주","방법","종류","혜택","문의","안내"}
def s(v): return "" if v is None else str(v).strip()

def main():
    if not SRC.exists(): print(f"::error::{SRC} 없음"); sys.exit(1)
    wb=openpyxl.load_workbook(SRC,data_only=True)
    if SHEET not in wb.sheetnames: print(f"::error::'{SHEET}' 시트 없음"); sys.exit(1)
    ws=wb[SHEET]
    head=[s(c.value) for c in ws[4]]
    need=["게시","종류","이름/라벨","내용","아이콘"]; errors=[]
    for h in need:
        if h not in head: errors.append(f"4행에 '{h}' 칸 없음")
    if errors:
        for e in errors: print("::error::"+e)
        sys.exit(1)
    ix={h:head.index(h) for h in head}
    def col(row,n):
        i=ix.get(n); return s(row[i]) if (i is not None and i<len(row)) else ""
    doc={"생성시각":datetime.now(KST).isoformat(timespec="seconds"),"출처":SRC.name,
         "소개":"","문의":"","안내":"","예금주":"","계좌":[],"방법":[],"종류":[],"혜택":[]}
    for r in range(5,ws.max_row+1):
        row=[c.value for c in ws[r]]
        if not any(s(v) for v in row): continue
        if col(row,"게시").upper() in ("X","×"): continue
        kind=col(row,"종류"); name=col(row,"이름/라벨"); text=col(row,"내용"); icon=col(row,"아이콘")
        if kind not in KINDS:
            errors.append(f"{r}행: 종류는 {' / '.join(sorted(KINDS))} 중 하나 (지금:'{kind}')"); continue
        if kind=="소개": doc["소개"]=text
        elif kind=="문의": doc["문의"]=text
        elif kind=="안내": doc["안내"]=text
        elif kind=="예금주": doc["예금주"]=text or name
        elif kind=="계좌": doc["계좌"].append({"은행":name,"번호":text})
        elif kind in ("방법","종류","혜택"): doc[kind].append({"이름":name,"설명":text,"아이콘":icon})
    if errors:
        for e in errors: print("  ✗ "+e)
        print(f"::error::잘못된 값 {len(errors)}군데"); sys.exit(1)
    OUT.write_text(json.dumps(doc,ensure_ascii=False,indent=1),encoding="utf-8")
    print(f"✓ {OUT} — 계좌 {len(doc['계좌'])} / 방법 {len(doc['방법'])} / 종류 {len(doc['종류'])} / 혜택 {len(doc['혜택'])}")

if __name__=="__main__": main()
