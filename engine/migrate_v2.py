#!/usr/bin/env python3
"""주간 워크북을 v1 → v2 레이아웃으로 1회 변환한다 (운영기준: 채권 주간 평가 시트 v2).

B~J 열 위치는 그대로 둔다(기존 수식·열 인덱스 보존). 바뀌는 것:
  - 행 의미: 12 수급·글로벌 기간 프리미엄 / 13 수급·국내 / 14 기술적·포지션  ('기타' 점수 행 폐지)
  - 비중 v2 (v1 비중은 K/L 열에 기록용으로 보존)
  - 16행 판정: ±0.5 5단계, 이벤트 ON이면 ±0.75, A′+근원CPI<3%면 U/W 한 단계 완화, 7행 미채점이면 '채점 미완'
  - M/N 태그, O/P 뉴스 참고(자동), Q 다이제스트(자동), R 하우스 뷰(자동), S~V 단일 테마 % 보조(숨김)
  - FLAG 블록(18~25행), 메모 칸(27~34행), W1 스키마 표식
  - News DB 표에 J열 '유형'(지표/이벤트/의견) 추가
  - **행 의미가 바뀌므로 8~14행의 점수·전주·근거(D,F,G,I,J)는 비운다.**

이미 v2인 파일에는 아무것도 하지 않는다(멱등). 과거 v1 주간 파일은 변환하지 말고 그대로 보존할 것.

사용법:
  python migrate_v2.py <workbook.xlsx>
  python migrate_v2.py <workbook.xlsx> --new-week YYYY-MM-DD
      → 변환 후 carry-blank까지 (v1 파일에서 막 복제된 새 주 파일용: D3·날짜·FLAG 초기화)
"""
import sys
from copy import copy
import openpyxl
from openpyxl.worksheet.table import TableColumn
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Alignment
from safe_io import safe_save
import sheet_v2 as sv

RELAX = 'AND(OR($C$19="A′",$C$19="A\'"),$C$24="Y")'
FLAG_ROWS = [
    (19, "국면", None, "A / A′ / B / C — S&P500 13주 수익률과 미 10년 13주 변화로 판정 (완충 ±2%, ±10bp면 직전 유지)"),
    (20, "조기경보 ρ60", None, "최근 60거래일 (S&P 일간수익률, 미10년 일간변화) 상관. A 국면에서 음(−)이면 A′ 경보"),
    (21, "전환 경계 (A→B)", "OFF", "Capex 하향 / 스프레드 추세 확대 / 반도체 수출 2개월 둔화 / 동반 하락 3주 — 둘 이상이면 ON"),
    (22, "이벤트", "OFF", "다음 주 안에 결과가 나오고 방향이 갈리는 비정례 이벤트면 ON → 판정 임계값 ±0.75"),
    (23, "TP 돌파", "OFF", "ACM TP10이 2023년 이후 최고치 경신 또는 1주 변화 ≥ 2σ → 기간 프리미엄 행 −2 고정"),
    (24, "미 근원 CPI 3% 미만", None, "Y / N — 국면 A′이고 Y이면 U/W 판정을 한 단계 완화"),
    (25, "판정 임계값", '=IF($C$22="ON",0.75,0.5)', "자동 계산 (이벤트 ON → 0.75)"),
]
MEMO_LABELS = ["미 10년(ACMY10) 주간 변화 (bp)", "└ 기대경로(RNY10) 몫 (bp)", "└ 기간 프리미엄(TP10) 몫 (bp, %)",
               "2s10s 중 TP 몫: Δ(TP10−TP02) (bp)", "2s10s 중 기대경로 몫: Δ(RNY10−RNY02) (bp)",
               "지난주 판정과 실제 결과 (국고10년 bp, 3s10s bp)", "3주 이상 같은 점수인 행의 '이번 주 새 정보' 유무"]
NOTE_16 = ("※ 판정(Total 기준): 강한 O/W(+1.0↑) / O/W(+0.5↑) / Neutral / U/W(−0.5↓) / 강한 U/W(−1.0↓). "
           "커브는 Flattening·Steepening.\n※ 이벤트 FLAG ON이면 ±0.75. 국면 A′ + 근원 CPI<3%이면 U/W 한 단계 완화. "
           "7행을 모두 채점해야 판정이 나온다(빈칸 금지, 0도 판단).\n※ 듀레이션 + = 금리 하락 요인, 커브 + = 플래트닝. "
           "O~R열은 자동 참고값이며 점수가 아니다.")


def style_from(dst, src, wrap=None):
    if src.has_style:
        dst.font, dst.border, dst.fill = copy(src.font), copy(src.border), copy(src.fill)
        dst.alignment, dst.number_format = copy(src.alignment), src.number_format
    if wrap is not None:
        dst.alignment = Alignment(horizontal="left" if wrap else "center", vertical="center", wrap_text=wrap)


def verdict_formula(total, hi, lo, count_rng, relaxable):
    strong_lo = f'IF({RELAX},"{lo}","강한 {lo}")' if relaxable else f'"강한 {lo}"'
    weak_lo = f'IF({RELAX},"Neutral","{lo}")' if relaxable else f'"{lo}"'
    return (f'=IF(COUNT({count_rng})<7,"채점 미완 "&COUNT({count_rng})&"/7",'
            f'IF({total}>=1,"강한 {hi}",IF({total}>=$C$25,"{hi}",'
            f'IF({total}<=-1,{strong_lo},IF({total}<=-$C$25,{weak_lo},"Neutral"))))'
            f'&IF($C$26="Y",""," (초안)"))')


def prev_formula(total, hi, lo, count_rng):
    return (f'=IF(COUNT({count_rng})=0,"",IF({total}>=1,"강한 {hi}",IF({total}>=0.5,"{hi}",'
            f'IF({total}<=-1,"강한 {lo}",IF({total}<=-0.5,"{lo}","Neutral")))))')


def migrate_rows(ws):
    C = sv.COL
    for r in sv.ROWS:                                  # v1 비중을 기록용으로 보존
        ws.cell(r, C["v1_dur_w"]).value = ws.cell(r, C["dur_w"]).value
        ws.cell(r, C["v1_cur_w"]).value = ws.cell(r, C["cur_w"]).value
    ws["K7"], ws["L7"] = "v1 평가비중", "v1 평가비중"
    for rng in ("B12:C12", "B13:C13"):
        if rng in {str(m) for m in ws.merged_cells.ranges}:
            ws.unmerge_cells(rng)
    ws.merge_cells("B12:B13")
    ws["B12"], ws["C12"], ws["C13"], ws["B14"] = "수급", "글로벌\n기간 프리미엄", "국내", "기술적·포지션"
    style_from(ws["B12"], ws["B10"]); style_from(ws["C12"], ws["C11"]); style_from(ws["C13"], ws["C11"])
    ws["C12"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["D7"] = "한 줄 근거 (듀: … / 커브: …)"
    for r in sv.ROWS:
        ws.cell(r, C["dur_w"]).value = sv.DUR_WEIGHTS[r]
        ws.cell(r, C["cur_w"]).value = sv.CUR_WEIGHTS[r]
        for col in (C["note"], C["dur_prev"], C["dur"], C["cur_prev"], C["cur"]):
            ws.cell(r, col).value = None               # 행 의미 변경 → 옛 값은 다른 뜻


def add_columns(ws):
    C = sv.COL
    heads = {C["dur_tag"]: "듀 태그", C["cur_tag"]: "커브 태그", C["news_dur"]: "뉴스 참고\n(듀)",
             C["news_cur"]: "뉴스 참고\n(커브)", C["digest"]: "뉴스 다이제스트 (자동)", C["house"]: "하우스 뷰 (자동 · 하우스당 1표)",
             C["dur_contrib"]: "기여(듀)", C["cur_contrib"]: "기여(커브)", C["dur_share"]: "테마비중(듀)", C["cur_share"]: "테마비중(커브)"}
    for col, text in heads.items():
        cell = ws.cell(7, col); cell.value = text
        style_from(cell, ws["J7"]); cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws["M5"], ws["O5"] = "태그 (사람)", "자동 참고 — 점수 아님"
    style_from(ws["M5"], ws["K5"]); style_from(ws["O5"], ws["K5"])
    for r in sv.ROWS:
        for col in (C["dur_tag"], C["cur_tag"], C["news_dur"], C["news_cur"]):
            style_from(ws.cell(r, col), ws.cell(r, C["cur"]), wrap=False)
        for col in (C["digest"], C["house"]):
            style_from(ws.cell(r, col), ws.cell(r, C["note"]), wrap=True)
        ws.cell(r, C["dur_contrib"]).value = f"=ABS(E{r}*G{r})"
        ws.cell(r, C["cur_contrib"]).value = f"=ABS(H{r}*J{r})"
        ws.cell(r, C["dur_share"]).value = (f'=IF(OR(M{r}="",SUM($S$8:$S$14)=0),0,'
                                           f'SUMIF($M$8:$M$14,M{r},$S$8:$S$14)/SUM($S$8:$S$14))')
        ws.cell(r, C["cur_share"]).value = (f'=IF(OR(N{r}="",SUM($T$8:$T$14)=0),0,'
                                           f'SUMIF($N$8:$N$14,N{r},$T$8:$T$14)/SUM($T$8:$T$14))')
    for tag_col, share in (("M", "U"), ("N", "V")):
        ws[f"{tag_col}15"] = (f'=IF(MAX({share}8:{share}14)>{sv.THEME_ALERT},"단일 테마 "&TEXT(MAX({share}8:{share}14),"0%")'
                              f'&"("&INDEX({tag_col}8:{tag_col}14,MATCH(MAX({share}8:{share}14),{share}8:{share}14,0))&")","분산")')
    for letter, width in (("M", 13), ("N", 13), ("O", 10), ("P", 10), ("Q", 44), ("R", 40)):
        ws.column_dimensions[letter].width = width
    for letter in "STUV":
        ws.column_dimensions[letter].hidden = True


def add_verdict(ws):
    ws["G16"] = verdict_formula("G15", "O/W", "U/W", "G8:G14", relaxable=True)
    ws["J16"] = verdict_formula("J15", "Flattening", "Steepening", "J8:J14", relaxable=False)
    ws["F16"] = prev_formula("F15", "O/W", "U/W", "F8:F14")
    ws["I16"] = prev_formula("I15", "Flattening", "Steepening", "I8:I14")
    ws["D16"] = NOTE_16


def add_flags_and_memo(ws):
    ws["B18"] = "FLAG — 점수가 아니라 다른 행의 부호·임계값을 바꾸는 조건 (채점보다 먼저 판정)"
    for row, label, default, hint in FLAG_ROWS:
        ws.cell(row, 2).value, ws.cell(row, 3).value, ws.cell(row, 4).value = label, default, hint
        ws.cell(row, 4).alignment = Alignment(wrap_text=True, vertical="center")
    for sqref, formula in (("C19", '"A,A′,B,C"'), ("C21:C23", '"ON,OFF"'), ("C24", '"Y,N"')):
        dv = DataValidation(type="list", formula1=formula, allow_blank=True)
        dv.add(sqref); ws.add_data_validation(dv)
    ws["B27"] = "메모 (채점하지 않음)"
    for row, label in zip(sv.MEMO_ROWS, MEMO_LABELS):
        ws.cell(row, 2).value = label


def add_draft_layout(ws):
    """점수 초안용 칸: C26 채점 확정, X~AC 숨김 장부(자동화가 마지막으로 쓴 값·잠금), 판정에 '(초안)' 접미.
    기존 점수·근거는 건드리지 않는다(멱등). 이미 v2인 파일의 업그레이드에도 쓴다."""
    ws["B26"] = "채점 확정"
    if ws[sv.FLAG["confirmed"]].value not in ("Y", "N"):
        ws[sv.FLAG["confirmed"]] = "N"
    ws["D26"] = "점수 조정을 마쳤으면 Y — 그 전까지 판정에 '(초안)'이 붙는다. 초안에 동의해 안 고쳤어도 Y를 찍어야 확정"
    ws["D26"].alignment = Alignment(wrap_text=True, vertical="center")
    dv = DataValidation(type="list", formula1='"Y,N"', allow_blank=True)
    dv.add(sv.FLAG["confirmed"]); ws.add_data_validation(dv)
    heads = {sv.SHADOW["dur"]: "초안(듀)", sv.SHADOW["cur"]: "초안(커브)", sv.SHADOW["dur_tag"]: "초안 태그(듀)",
             sv.SHADOW["cur_tag"]: "초안 태그(커브)", sv.SHADOW["note"]: "초안 요약", sv.LOCK_COL: "사람이 고친 칸"}
    for col, text in heads.items():
        ws.cell(7, col).value = text
    for letter in ("X", "Y", "Z", "AA", "AB", "AC"):
        ws.column_dimensions[letter].hidden = True
    ws["D7"] = "요약 · 근거 (자동 초안 → 고치면 그 주엔 보존)"
    add_verdict(ws)
    ws[sv.DRAFT_CELL] = sv.DRAFT_MARK


def apply_compact_layout(ws):
    """보고서형 서식: 요약 칸이 불릿 3줄에 맞게 행 높이를 균일하게, 줄바꿈·왼쪽 정렬, 태그 열 폭 확보.
    v1에서 물려받은 행 높이(최대 210pt)를 정리한다. 1회만 적용(W3 표식)."""
    for r in sv.ROWS:
        ws.row_dimensions[r].height = sv.ROW_HEIGHT
        for col in (sv.COL["note"], sv.COL["digest"], sv.COL["house"]):
            ws.cell(r, col).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.row_dimensions[3].height = sv.CONC_ROW_HEIGHT
    ws["D3"].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    for letter in ("M", "N"):
        ws.column_dimensions[letter].width = 17
    ws[sv.COMPACT_CELL] = sv.COMPACT_MARK


def migrate_news_db(nd):
    tbl = nd.tables["NewsDB"]
    start, end = tbl.ref.split(":")
    if end.startswith("J"):
        return
    end_row = int("".join(c for c in end if c.isdigit()))
    if any(nd.cell(r, 10).value is not None for r in range(4, end_row + 1)):
        raise RuntimeError("News DB J열이 비어 있지 않아 '유형' 열을 추가할 수 없습니다.")
    nd.cell(4, 10).value = "유형"
    for r in range(4, end_row + 1):
        style_from(nd.cell(r, 10), nd.cell(r, 9))
    tbl.tableColumns.append(TableColumn(id=10, name="유형"))
    tbl.ref = f"{start}:J{end_row}"
    if tbl.autoFilter is not None:
        tbl.autoFilter.ref = tbl.ref
    dv = DataValidation(type="list", formula1='"지표,이벤트,의견"', allow_blank=True)
    dv.add("J5:J999"); nd.add_data_validation(dv)
    nd.column_dimensions["J"].width = 9


def migrate(wb_path):
    wb = openpyxl.load_workbook(wb_path)
    ws = wb["Analysis"]
    if sv.is_v2(ws):
        if sv.has_draft_layout(ws):
            print("이미 v2 레이아웃 — 변경 없음: " + wb_path)
            return False
        add_draft_layout(ws)                       # 초안 도입 전의 v2 파일: 칸만 추가, 점수는 그대로
        safe_save(wb, wb_path, expect_sheets=["News DB", "Analysis"])
        print("v2 파일에 점수 초안 레이아웃 추가: " + wb_path)
        return True
    migrate_rows(ws)
    add_columns(ws)
    add_flags_and_memo(ws)
    add_draft_layout(ws)
    migrate_news_db(wb["News DB"])
    ws[sv.SCHEMA_CELL] = sv.SCHEMA_V2
    safe_save(wb, wb_path, expect_sheets=["News DB", "Analysis"])
    print("v2 변환 완료: " + wb_path)
    return True


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(1)
    migrate(sys.argv[1])
    if len(sys.argv) > 3 and sys.argv[2] == "--new-week":
        # 막 복제된 새 주 파일이면 D3·날짜·FLAG를 새 주 상태로 초기화한다.
        import update_analysis
        update_analysis.carry_blank(sys.argv[1], sys.argv[3])
