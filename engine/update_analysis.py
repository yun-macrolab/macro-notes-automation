#!/usr/bin/env python3
"""Analysis 시트 갱신 (v2) — 자동 칸만 쓴다.

v2 원칙: **자동화는 점수를 쓰지 않는다.**
  이 스크립트가 쓰는 칸:  O/P(뉴스 참고 점수) · Q(뉴스 다이제스트) · R(하우스 뷰) · D3(결론) · I1(날짜)
  절대 쓰지 않는 칸:      D(한 줄 근거) · G/J(금주 점수) · M/N(태그) · E/H(비중) · FLAG · 메모
  (예외: carry-blank 는 새 주 시작 시 1회, 금주→전주 이월 후 사람 칸을 비운다)

v1 레이아웃 워크북에는 쓰기를 거부한다(SchemaError) — 행 의미가 달라 다른 칸에 값이 들어가기 때문.

사용법:
  python update_analysis.py <workbook.xlsx> <analysis.auto.json>
      → rollup 결과를 자동 칸에 기입. 이번 롤업에 없는 행의 낡은 자동 칸은 비운다.

  python update_analysis.py <workbook.xlsx> carry-blank [YYYY-MM-DD]
      → 새 주 시작용(파일 생성 직후 딱 1회). 금주(G/J)→전주(F/I)로 **항상** 덮어쓴 뒤
        금주 점수·근거·태그·자동 칸·메모를 비운다. 빈칸 금지 원칙상 매주 새로 채점하게 하기 위함.
        FLAG: 이벤트·TP 돌파는 OFF로 리셋, ρ60은 비움, 국면·전환 경계·근원 CPI 여부는 유지.

analysis.auto.json 형식 (rollup_analysis.py 가 생성):
{
  "date": "2026-09-21",
  "conclusion": "【금주 결론】 ...",     // 선택, D3
  "factors": [{"row": 8, "news_dur": -1.5, "news_curve": 0.5,
               "digest": "2건: NFP(+1.5); ...", "house_text": "상방 2 · 하방 1 (...)"}]
}
"""
import sys, json, datetime
import openpyxl
from safe_io import safe_save
import sheet_v2 as sv

SHEETS = ["News DB", "Analysis"]
PLACEHOLDER = "【금주 결론】 새 주 시작 — 전주 점수 이월 완료. 금주 점수는 주 1회 직접 채점한다(7행 모두)."


def set_date(ws, date_str):
    if not date_str:
        return
    try:
        ws["I1"] = datetime.datetime.strptime(date_str, "%Y-%m-%d")
        ws["I1"].number_format = "yyyy-mm-dd"
    except ValueError:
        ws["I1"] = date_str


def carry_blank(wb_path, date_str=None):
    wb = openpyxl.load_workbook(wb_path)
    ws = wb["Analysis"]
    sv.require_v2(ws, wb_path)
    C = sv.COL
    for row in sv.ROWS:
        ws.cell(row, C["dur_prev"]).value = ws.cell(row, C["dur"]).value    # 빈칸이어도 덮어쓴다
        ws.cell(row, C["cur_prev"]).value = ws.cell(row, C["cur"]).value
        for col in (*sv.HUMAN_COLS, *sv.AUTO_COLS, *sv.SHADOW.values(), sv.LOCK_COL):
            ws.cell(row, col).value = None             # 초안 장부·잠금도 새 주엔 초기화
    ws[sv.FLAG["confirmed"]] = "N"
    ws[sv.FLAG["event"]] = "OFF"
    ws[sv.FLAG["tp_break"]] = "OFF"
    ws[sv.FLAG["rho60"]].value = None
    for row in sv.MEMO_ROWS:
        ws.cell(row, 3).value = None
    set_date(ws, date_str)
    ws["D3"] = PLACEHOLDER
    safe_save(wb, wb_path, expect_sheets=SHEETS)
    print("carry-blank 완료: 금주(G/J)->전주(F/I) 이월, 사람 칸·자동 칸 비움, 이벤트/TP돌파 FLAG OFF. 저장: " + wb_path)


def apply(wb_path, cfg, out_path=None):
    wb = openpyxl.load_workbook(wb_path)
    ws = wb["Analysis"]
    sv.require_v2(ws, wb_path)
    C = sv.COL
    set_date(ws, cfg.get("date"))
    if cfg.get("conclusion"):
        ws["D3"] = cfg["conclusion"]
    for row in sv.ROWS:                       # 낡은 자동 칸 제거
        for col in sv.AUTO_COLS:
            ws.cell(row, col).value = None
    changed = []
    for f in cfg.get("factors", []):
        row = f.get("row")
        if row not in sv.ROWS:
            print("행 " + str(row) + " 무시(8~14만 허용)")
            continue
        ws.cell(row, C["news_dur"]).value = f.get("news_dur")
        ws.cell(row, C["news_cur"]).value = f.get("news_curve")
        ws.cell(row, C["digest"]).value = f.get("digest")
        ws.cell(row, C["house"]).value = f.get("house_text")
        changed.append(row)
    safe_save(wb, out_path or wb_path, expect_sheets=SHEETS)
    print("Analysis 자동 칸 갱신(O~R): 행 " + str(sorted(changed)) + " — 금주 점수(G/J)·근거·태그는 건드리지 않음")
    return changed


def main():
    if len(sys.argv) < 3:
        print(__doc__); sys.exit(1)
    wb_path, json_path = sys.argv[1], sys.argv[2]
    try:
        if json_path == "carry-blank":
            carry_blank(wb_path, sys.argv[3] if len(sys.argv) > 3 else None)
            return
        with open(json_path, encoding="utf-8") as fp:
            cfg = json.load(fp)
        apply(wb_path, cfg, sys.argv[3] if len(sys.argv) > 3 else None)
    except sv.SchemaError as e:
        print("실패: " + str(e), file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
