#!/usr/bin/env python3
"""News DB 주간 누적 → 행별 '뉴스 참고' 집계 (analysis.auto.json 생성, 워크북은 건드리지 않음).

v2 원칙: **이 스크립트의 숫자는 점수가 아니라 참고다.** 금주 점수(G/J)는 사람이 주 1회 쓴다.
update_analysis.py는 이 결과를 O/P(뉴스 참고 점수)·Q(다이제스트)·R(하우스 뷰) 열에만 기입한다.

[집계 규칙]
- 뉴스 참고(듀) = −(지표·이벤트 뉴스 강도 합)  → 캡 ±2.0 → 0.5 단위.  (News DB 강도 + = 금리상방)
- 뉴스 참고(커브) = 본문 키워드 투표(플래트닝 +, 스티프닝 −) × 0.5 → 캡 ±1.5
- 통화정책 2행과 수급·글로벌TP 행은 참고 점수를 내지 않는다(다이제스트만).
    정책은 '시장 프라이싱 − 내 경로' 갭으로, TP는 ACM z-score로 채점하는 행이라
    뉴스 방향을 더하면 이미 가격에 있는 것을 다시 세게 된다.
- 유형이 '의견'(애널리스트·하우스 뷰)인 행은 참고 점수에 넣지 않고 하우스 뷰로 따로 센다.
    하우스당 1표(같은 하우스는 최신 1건). 컨센서스는 더하는 항이 아니라 대조할 기준선이다.
- 요인 '기타'는 점수 행이 없다 → 이벤트 FLAG 후보로만 보고한다.
- A열 '예시' 행 제외, 이번 주(월~일) 밖 날짜 제외. 날짜 없으면 포함.

사용법:
  python rollup_analysis.py <workbook.xlsx> [--date YYYY-MM-DD] [--out analysis.auto.json]
"""
import sys, os, json, argparse, datetime
import openpyxl
import sheet_v2 as sv

FACTORS = {"펀더멘털", "통화정책", "수급", "기술적분석", "기타"}
CURVE_STEP = 0.5
FLATTEN_KW = ["플래트닝", "플랫트닝", "플랫", "평탄", "커브 축소", "커브축소",
              "장단기 축소", "수익률곡선 평탄"]
STEEPEN_KW = ["스티프닝", "스팁", "가팔", "커브 확대", "커브확대",
              "장단기 확대", "수익률곡선 가팔"]
DIR_KEY = {"금리상방": "up", "금리하방": "down", "중립": "neutral"}
DIR_KO = {"up": "상방", "down": "하방", "neutral": "중립"}


def week_bounds(d):
    mon = d - datetime.timedelta(days=d.weekday())
    return mon, mon + datetime.timedelta(days=6)


def parse_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            return datetime.datetime.strptime(str(v).strip()[:10], fmt).date()
        except ValueError:
            continue
    return None


def to_float(v):
    try:
        return None if v is None or v == "" else float(v)
    except (TypeError, ValueError):
        return None


def curve_vote(text):
    """ +1=플래트닝 신호, -1=스티프닝 신호, 0=없음/혼재 """
    t = text.replace("플랫폼", "")  # '플랫' 키워드가 '플랫폼'에 오탐하는 것 방지
    flat = any(k in t for k in FLATTEN_KW)
    steep = any(k in t for k in STEEPEN_KW)
    return 1 if flat and not steep else -1 if steep and not flat else 0


def read_week_rows(wb_path, today):
    """News DB에서 이번 주 실제 뉴스 행을 dict 목록으로."""
    mon, sun = week_bounds(today)
    wb = openpyxl.load_workbook(wb_path, read_only=True)
    out = []
    for r in wb["News DB"].iter_rows(min_row=5, max_col=10, values_only=True):
        r = tuple(r) + (None,) * (10 - len(r))
        if str(r[0]).strip() == "예시" or r[1] not in FACTORS:
            continue
        d = parse_date(r[0])
        if d is not None and not (mon <= d <= sun):
            continue
        out.append({"date": d, "factor": r[1], "region": r[2], "keyword": r[3] or "", "title": r[4] or "",
                    "content": r[5] or "", "direction": r[6], "intensity": r[7], "source": r[8] or "",
                    "kind": r[9]})
    wb.close()
    return out


def house_view(opinions):
    """하우스당 1표(최신 1건). → {'up': [...], 'down': [...], 'neutral': [...]}"""
    latest = {}
    for i, it in enumerate(opinions):
        house = (it.get("source") or "").strip() or "미상"
        key = (parse_date(it.get("date")) or datetime.date.min, i)
        if house not in latest or key >= latest[house][0]:
            latest = {**latest, house: (key, it)}
    view = {"up": [], "down": [], "neutral": []}
    for house, (_, it) in latest.items():
        k = DIR_KEY.get(it.get("direction"), "neutral")
        view = {**view, k: view[k] + [house]}
    return view


def house_text(view):
    if not any(view.values()):
        return None
    counts = " · ".join(f"{DIR_KO[k]} {len(view[k])}" for k in ("up", "down", "neutral") if view[k])
    names = " / ".join(f"{DIR_KO[k]}: {', '.join(view[k])}" for k in ("up", "down", "neutral") if view[k])
    return f"{counts} ({names})"


def fmt_item(it):
    s = to_float(it.get("intensity")) or 0.0
    return f"{it.get('keyword') or it.get('title')}({s:+g})"


def brief(it):
    """초안 요약을 쓰는 쪽(헤드리스 LLM)에 넘길 이번 주 뉴스 1건. 오늘 메모만이 아니라 주간 누적을 보게 한다."""
    d = parse_date(it.get("date"))
    return {"date": d.isoformat() if d else "", "kind": "의견" if sv.is_opinion(it) else (it.get("kind") or "지표"),
            "title": it.get("title"), "content": str(it.get("content") or "")[:400],
            "direction": it.get("direction"), "intensity": to_float(it.get("intensity")),
            "source": it.get("source") or ""}


def summarize_row(row, items, warns):
    facts = [it for it in items if not sv.is_opinion(it)]
    opinions = [it for it in items if sv.is_opinion(it)]
    news_dur = news_curve = None
    if facts and row not in sv.NO_NEWS_SCORE_ROWS:
        net = 0.0
        for it in facts:
            s = to_float(it.get("intensity"))
            if s is None:
                warns.append(f"[{sv.ROW_LABEL[row]}] 강도 결측 → 0 ({it.get('keyword')})")
                s = 0.0
            if (it.get("direction") == "금리상방" and s < 0) or (it.get("direction") == "금리하방" and s > 0):
                warns.append(f"[{sv.ROW_LABEL[row]}] 방향-부호 불일치 ({it.get('keyword')}, {s:+g})")
            net += s
        if abs(net) > sv.NEWS_CAP:
            warns.append(f"[{sv.ROW_LABEL[row]}] 참고 점수 캡 적용: 원값 {-net:+.1f}")
        news_dur = sv.round_half(sv.clamp(-net, sv.NEWS_CAP))
        votes = sum(curve_vote(f"{it['keyword']} {it['title']} {it['content']}") for it in facts)
        if votes:
            news_curve = sv.round_half(sv.clamp(votes * CURVE_STEP, sv.CURVE_CAP))
    view = house_view(opinions)
    digest = (f"{len(facts)}건: " + "; ".join(fmt_item(it) for it in facts)) if facts else None
    return {"row": row, "news_dur": news_dur, "news_curve": news_curve, "digest": digest,
            "house": view, "house_text": house_text(view), "n_facts": len(facts), "n_opinions": len(opinions)}


def aggregate(rows):
    """이번 주 뉴스 행 목록 → v2 참고 집계. 순수 함수(워크북 I/O 없음)."""
    warns, by_row, events = [], {}, []
    for it in rows:
        if it.get("factor") == "기타":
            events.append({"keyword": it.get("keyword"), "title": it.get("title"),
                           "direction": it.get("direction")})
            continue
        row = sv.map_row(it.get("factor"), it.get("region"))
        if row is None:
            warns.append(f"행 매핑 실패: 요인='{it.get('factor')}' 지역='{it.get('region')}'")
            continue
        if it.get("factor") in sv.REGION_SPLIT and it.get("region") not in ("글로벌", "국내"):
            warns.append(f"지역 없음 → 국내 행으로 집계 ({it.get('keyword')})")
        by_row = {**by_row, row: by_row.get(row, []) + [it]}
    opinions = [it for it in rows if sv.is_opinion(it)]
    return {"factors": [summarize_row(r, by_row[r], warns) for r in sorted(by_row)],
            "week_items": {str(r): [brief(it) for it in by_row[r]] for r in sorted(by_row)},
            "event_candidates": events,
            "opinion_n": len(opinions),
            "opinion_houses": sorted({(it.get("source") or "").strip() or "미상" for it in opinions}),
            "warns": warns}


def print_table(out, today):
    mon, sun = week_bounds(today)
    print(f"=== Rollup {today} (week {mon}~{sun}) — 뉴스 참고 집계 (점수 아님) ===")
    print(f"{'행':>3} {'요인':<14}{'지표':>4}{'의견':>4}{'참고(듀)':>9}{'참고(커브)':>10}  하우스 뷰")
    for f in out["factors"]:
        d = "-" if f["news_dur"] is None else f"{f['news_dur']:+.1f}"
        c = "-" if f["news_curve"] is None else f"{f['news_curve']:+.1f}"
        print(f"{f['row']:>3} {sv.ROW_LABEL[f['row']]:<14}{f['n_facts']:>4}{f['n_opinions']:>4}{d:>9}{c:>10}  "
              f"{f['house_text'] or ''}")
    if out["event_candidates"]:
        print("\n[이벤트 FLAG 후보 — 요인 '기타', 점수 행 없음]")
        for e in out["event_candidates"]:
            print(f"  - {e['keyword']}: {e['title']}")
    if out["opinion_n"]:
        print(f"\n의견으로 분류: {out['opinion_n']}건 ({', '.join(out['opinion_houses'])})")
    if out["warns"]:
        print("\n[경고/검토]")
        for w in out["warns"]:
            print("  - " + w)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("workbook")
    ap.add_argument("--date", default=None)
    ap.add_argument("--out", default=None)
    a = ap.parse_args()
    today = (datetime.datetime.strptime(a.date, "%Y-%m-%d").date() if a.date else datetime.date.today())
    out_path = a.out or os.path.join(os.getcwd(), "analysis.auto.json")

    out = aggregate(read_week_rows(a.workbook, today))
    payload = {"date": today.strftime("%Y-%m-%d"), "schema": 2,
               "factors": [{k: f[k] for k in ("row", "news_dur", "news_curve", "digest", "house_text")}
                           for f in out["factors"]],
               "week_items": out["week_items"],
               "event_candidates": out["event_candidates"],
               "opinion_n": out["opinion_n"], "opinion_houses": out["opinion_houses"]}
    with open(out_path, "w", encoding="utf-8") as fp:
        json.dump(payload, fp, ensure_ascii=False, indent=2)
    print_table(out, today)
    print(f"\n생성: {out_path}")


if __name__ == "__main__":
    main()
