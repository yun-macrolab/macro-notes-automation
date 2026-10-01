"""공개본 — 주간 기록에서 화면에 보이는 칸만 뽑고, 올리기 전에 개인정보를 검사한다(표준 라이브러리만 사용).

허용 목록에 없는 칸은 싣지 않는다(엑셀 문서 속성·숨김 장부 열·시트 행 번호 등은 처음부터 빠진다).
뉴스는 그 주간(Analysis I1 기준 월~일)에 날짜가 든 행만 싣는다.
화면(web/viewer.js)은 내 파일과 공개본을 같은 모양의 자료로 그린다.
쓰는 곳: 브라우저의 '공개본 만들기'(public_json), 배포 때 records/ 검사(web/build.py).
"""
import datetime as dt
import json
import math
import re
import unicodedata

SCHEMA = "macro-notes-public/1"
NOTICE = "개인이 매주 정리한 매크로 기록의 공개본입니다. 투자 판단이나 권유가 아닙니다."
TOP_KEYS = ("schema", "notice", "week", "generated_at", "conclusion", "state", "divergences", "memo", "news")
WEEK_KEYS = ("start", "end")
STATE_KEYS = ("rows", "dur", "cur", "confirmed", "edited_n", "flags")
ROW_KEYS = ("label", "dur", "cur", "dur_prev", "cur_prev", "dur_tag", "cur_tag",
            "note", "digest", "news_dur", "news_cur", "house", "locked")
BLOCK_KEYS = ("total", "verdict")
FLAG_KEYS = ("regime", "rho60", "transition", "event_on", "tp_break", "core_cpi_lt3", "relax")
NEWS_KEYS = ("date", "factor", "region", "keyword", "title", "content", "direction", "intensity", "source", "kind")
MEMO_KEYS = ("label", "value")
SCALAR = (str, int, float, bool, type(None))

# 누구에게나 해당하는 개인정보 모양. 실명·학교처럼 사람마다 다른 말은 검사어로 따로 받는다.
GENERIC = (
    ("이메일 주소", re.compile(r"[\w.+-]+@[\w-]+(\.[\w-]+)+")),
    ("전화번호", re.compile(r"(?<!\d)(?:0|\+82[-. ]?)\d{1,2}[-. ]?\d{3,4}[-. ]?\d{4}(?!\d)")),
    ("주민등록번호", re.compile(r"(?<!\d)\d{2}[01]\d[0-3]\d-[1-8]\d{6}(?!\d)")),
    # 인터넷 주소 속 /home/·/users/는 PC 경로로 보지 않는다(바로 앞이 주소에 쓰는 글자면 건너뛴다).
    ("PC 경로", re.compile(r"[a-z]:\\users\\|(?<![a-z0-9._%-])/(?:users|home)/[^/\s]+/|file:/", re.IGNORECASE)),
)

# 검사에 걸린 위치를 엑셀에서 찾을 수 있게 알려 줄 때 쓰는 칸 이름
FIELD_NAMES = {"date": "날짜", "factor": "요인", "region": "지역", "keyword": "키워드", "title": "제목",
               "content": "핵심내용", "direction": "방향", "source": "출처", "kind": "유형", "label": "이름",
               "value": "내용", "note": "근거", "digest": "뉴스 요약", "house": "하우스",
               "dur_tag": "듀레이션 태그", "cur_tag": "커브 태그"}


def _clean(value):
    """JSON에 넣을 수 없는 값(무한대·NaN)은 빈 값으로."""
    return None if isinstance(value, float) and not math.isfinite(value) else value


def _in_week(day, start, end):
    return isinstance(day, str) and bool(re.fullmatch(r"\d{4}-\d{2}-\d{2}", day)) and start <= day <= end


def build(data, generated_at=None):
    """workbook_view.read() 결과 → 공개본(dict). 주간 날짜가 없으면 ValueError."""
    week = data["week"]
    if not week.get("start"):
        raise ValueError("Analysis I1 칸에 주간 날짜가 없어 공개본을 만들 수 없습니다.")
    start, end = week["start"], week["end"]
    state = data["state"]
    rows = []
    for source in state["rows"]:
        row = {key: _clean(source.get(key)) for key in ROW_KEYS}
        row["locked"] = [x for x in (source.get("locked") or []) if x in ("dur", "cur")]
        rows.append(row)
    return {
        "schema": SCHEMA,
        "notice": NOTICE,
        "week": {"start": start, "end": end},
        "generated_at": generated_at or dt.datetime.now(dt.timezone.utc).isoformat(timespec="seconds"),
        "conclusion": data.get("conclusion") or "",
        "state": {
            "rows": rows,
            "dur": {key: _clean(state["dur"][key]) for key in BLOCK_KEYS},
            "cur": {key: _clean(state["cur"][key]) for key in BLOCK_KEYS},
            "confirmed": bool(state["confirmed"]),
            "edited_n": int(state["edited_n"]),
            "flags": {key: _clean(state["flags"].get(key)) for key in FLAG_KEYS},
        },
        "divergences": [str(x) for x in data.get("divergences") or []],
        "memo": [{key: m.get(key) or "" for key in MEMO_KEYS} for m in data.get("memo") or []],
        "news": [{key: _clean(n.get(key)) for key in NEWS_KEYS}
                 for n in data["news"] if _in_week(n.get("date"), start, end)],
    }


def _strings(value, where=""):
    if isinstance(value, str):
        yield where, value
    elif isinstance(value, dict):
        for key, item in value.items():
            yield from _strings(item, f"{where}.{key}" if where else key)
    elif isinstance(value, (list, tuple)):
        for index, item in enumerate(value):
            yield from _strings(item, f"{where}[{index}]")


def _normal(text):
    """전각·호환 문자는 보통 문자로, 대소문자는 같게, 보이지 않는 글자(폭 없는 공백 등)는 빼고 본다."""
    text = unicodedata.normalize("NFKC", text).casefold()
    return "".join(ch for ch in text if unicodedata.category(ch) != "Cf")


def _squash(text):
    """검사어는 띄어쓰기를 무시하고 찾는다(‘홍 길동’도 ‘홍길동’으로 걸린다)."""
    return re.sub(r"\s+", "", _normal(text))


def findings(record, words=()):
    """공개 전 검사 — [(위치, 종류, 걸린 말)]. words는 사용자가 정한 검사어(실명·학교 등)."""
    words = [(w.strip(), _squash(w)) for w in words if isinstance(w, str) and _squash(w)]
    found = []
    for where, text in _strings(record):
        normal = _normal(text)
        squashed = re.sub(r"\s+", "", normal)
        for word, key in words:
            if key in squashed:
                found.append((where, "검사어", word))
        for kind, pattern in GENERIC:
            match = pattern.search(normal)
            if match:
                found.append((where, kind, match.group(0)))
    return found


def place(record, where):
    """검사에 걸린 위치(news[3].content 같은 경로)를 엑셀에서 찾을 수 있는 말로 바꾼다."""
    match = re.fullmatch(r"(news|memo|state\.rows|divergences)\[(\d+)\](?:\.(\w+))?", where)
    if not match:
        return {"conclusion": "금주 결론(Analysis D3)"}.get(where, where)
    part, index = match.group(1), int(match.group(2))
    field = FIELD_NAMES.get(match.group(3), match.group(3))
    if part == "news":
        news = record["news"][index]
        return f"News DB {news.get('date')} ‘{(news.get('title') or '')[:30]}’의 {field}"
    if part == "memo":
        return f"메모 ‘{record['memo'][index].get('label') or index + 1}’의 {field}"
    if part == "state.rows":
        return f"점수표 ‘{record['state']['rows'][index].get('label')}’의 {field}"
    return "복기 포인트(점수에서 자동으로 만든 문장)"


def public_json(data_text, words_text="[]"):
    """브라우저 '공개본 만들기' — (성공 여부, JSON 문자열).
    검사에 걸리면 만들지 않고 위치·종류·걸린 말을 돌려준다(그 브라우저 화면에만 보인다)."""
    try:
        record = build(json.loads(data_text))
    except ValueError as error:
        return False, json.dumps({"error": str(error), "findings": []}, ensure_ascii=False)
    found = findings(record, json.loads(words_text))
    if found:
        return False, json.dumps({
            "error": "공개하면 안 될 것으로 보이는 내용이 있어 공개본을 만들지 않았습니다. 엑셀에서 고쳐 저장한 뒤 다시 만드세요.",
            "findings": [{"place": place(record, where), "kind": kind, "text": text} for where, kind, text in found],
        }, ensure_ascii=False)
    validate(record)
    return True, json.dumps(record, ensure_ascii=False, indent=2, allow_nan=False) + "\n"


def _exact(obj, keys, where):
    if not isinstance(obj, dict):
        raise ValueError(f"{where}: 객체가 아닙니다")
    if set(obj) != set(keys):
        extra, missing = sorted(set(obj) - set(keys)), sorted(set(keys) - set(obj))
        raise ValueError(f"{where}: 칸이 정해 둔 것과 다릅니다(더 있음 {extra}, 빠짐 {missing})")


def _flat(obj, keys, where, lists=()):
    """정해 둔 칸만 있고 값이 글자·숫자·참거짓·빈 값인 객체인지. lists에 든 칸은 글자 목록."""
    _exact(obj, keys, where)
    for key, value in obj.items():
        if key in lists:
            ok = isinstance(value, list) and all(isinstance(x, str) for x in value)
        else:
            ok = isinstance(value, SCALAR) and not (isinstance(value, float) and not math.isfinite(value))
        if not ok:
            raise ValueError(f"{where}.{key}: 값의 형식이 맞지 않습니다")


def _list(value, where):
    if not isinstance(value, list):
        raise ValueError(f"{where}: 목록이 아닙니다")
    return value


def validate(record, words=()):
    """저장소에 올라온 공개본이 정해 둔 모양이고 검사를 통과하는지. 아니면 ValueError(배포 중단).
    오류 문구에는 걸린 말을 싣지 않는다(공개 로그로 검사어나 개인정보가 새지 않게)."""
    _exact(record, TOP_KEYS, "공개본")
    if record["schema"] != SCHEMA:
        raise ValueError(f"schema가 {SCHEMA}가 아닙니다")
    if record["notice"] != NOTICE:
        raise ValueError("notice가 정해 둔 안내 문구와 다릅니다")
    _flat(record["week"], WEEK_KEYS, "week")
    try:
        start = dt.date.fromisoformat(record["week"]["start"])
        end = dt.date.fromisoformat(record["week"]["end"])
        dt.datetime.fromisoformat(record["generated_at"])
    except (TypeError, ValueError):
        raise ValueError("week.start·end 또는 generated_at이 날짜 형식이 아닙니다") from None
    if start.weekday() != 0 or end != start + dt.timedelta(days=6):
        raise ValueError("week는 월요일부터 그 주 일요일까지여야 합니다")
    if not isinstance(record["conclusion"], str):
        raise ValueError("conclusion: 글자가 아닙니다")
    state = record["state"]
    _exact(state, STATE_KEYS, "state")
    if not isinstance(state["confirmed"], bool) or isinstance(state["edited_n"], bool) \
            or not isinstance(state["edited_n"], int):
        raise ValueError("state.confirmed·edited_n의 형식이 맞지 않습니다")
    for index, row in enumerate(_list(state["rows"], "state.rows")):
        _flat(row, ROW_KEYS, f"state.rows[{index}]", lists=("locked",))
    for block in ("dur", "cur"):
        _flat(state[block], BLOCK_KEYS, f"state.{block}")
    _flat(state["flags"], FLAG_KEYS, "state.flags")
    for index, news in enumerate(_list(record["news"], "news")):
        _flat(news, NEWS_KEYS, f"news[{index}]")
        if not _in_week(news["date"], start.isoformat(), end.isoformat()):
            raise ValueError(f"news[{index}]: 날짜가 그 주간 밖이거나 YYYY-MM-DD가 아닙니다")
    for index, memo in enumerate(_list(record["memo"], "memo")):
        _flat(memo, MEMO_KEYS, f"memo[{index}]")
    if not all(isinstance(x, str) for x in _list(record["divergences"], "divergences")):
        raise ValueError("divergences: 글자 목록이 아닙니다")
    found = findings(record, words)
    if found:
        raise ValueError("개인정보로 보이는 내용(걸린 말은 싣지 않음): "
                         + ", ".join(f"{where}({kind})" for where, kind, _ in found[:5])
                         + (f" 외 {len(found) - 5}곳" if len(found) > 5 else ""))
    return True
