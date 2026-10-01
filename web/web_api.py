"""브라우저(Pyodide) 어댑터 — 로컬 서버 없이 app.py와 같은 경로로 DemoStore를 부른다.

화면(static/app.js)이 보내는 /api/* 요청을 web/bridge.js가 가로채 여기로 넘긴다.
응답 본문과 상태 코드, 오류 문구는 app.py와 같게 맞춘다. 브라우저 안에서만 도는 구조라
네트워크 경계가 없으므로 토큰은 화면 호환용으로만 돌려준다.
engine/과 core.py는 로컬 시연과 같은 파일을 그대로 쓴다.
"""
import json
import secrets
from urllib.parse import parse_qs, urlparse

from core import BASE, DemoStore

_store = None
_token = secrets.token_urlsafe(32)


def init(storage):
    """실행 기록 폴더(브라우저 저장소에 연결된 경로)로 DemoStore를 연다."""
    global _store
    _store = DemoStore(storage)


def _evidence():
    path = BASE / "evidence/results.json"
    if not path.exists():
        return {"cases": [], "notice": "검증 결과가 아직 없습니다."}
    return json.loads(path.read_text(encoding="utf-8"))


def _reply(status, data):
    return status, json.dumps(data, ensure_ascii=False)


def request(method, url, body=None):
    """(상태 코드, JSON 문자열)을 돌려준다. url은 '/api/...' 경로와 쿼리."""
    parsed = urlparse(url)
    route = parsed.path
    try:
        if method == "GET":
            if route == "/api/config":
                return _reply(200, {"token": _token, "mode": "browser_pyodide", "app": "macro-notes-demo-v1"})
            if route == "/api/history":
                return _reply(200, _store.history())
            if route == "/api/evidence":
                return _reply(200, _evidence())
            if route == "/api/state":
                return _reply(200, _store.state(parse_qs(parsed.query).get("id", [""])[0]))
            return _reply(404, {"error": "해당 페이지가 없습니다."})
        if method == "POST":
            if not body or len(body) > 4096:
                return _reply(400, {"error": "요청 크기가 올바르지 않습니다."})
            payload = json.loads(body)
            if not isinstance(payload, dict):
                raise ValueError("JSON 객체가 필요합니다.")
            if route == "/api/run":
                return _reply(200, _store.run(payload.get("scenario", "baseline")))
            if route == "/api/review":
                return _reply(200, _store.review(payload.get("id"), payload.get("score")))
            if route == "/api/reapply":
                return _reply(200, _store.reapply(payload.get("id")))
            return _reply(404, {"error": "지원하지 않는 동작입니다."})
        return _reply(405, {"error": "지원하지 않는 요청입니다."})
    except (ValueError, KeyError, TypeError, FileNotFoundError) as exc:
        return _reply(400, {"error": str(exc)})
    except Exception:
        return _reply(500, {"error": "실행 실패. 페이지를 새로고침한 뒤 다시 시도하세요."})


def workbook(run_id):
    """해당 실행의 합성 워크북 바이트. 실행 ID 검사는 DemoStore.directory가 한다."""
    return (_store.directory(run_id) / "demo.xlsx").read_bytes()
