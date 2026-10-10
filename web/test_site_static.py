"""배포되는 화면 파일의 정적 검사 — 브라우저 없이 글자만 본다(표준 라이브러리만, 네트워크·실제 기록 불필요).

이 사이트는 같은 주소 아래의 다른 화면과 브라우저 저장 공간을 함께 쓴다. 그래서 실제로 나가는 화면 파일
(web/build.py의 PAGES)에 위험한 줄이 실수로 들어가지 않는지 본다: 글자를 HTML로 넣는 낱말, 바깥에서 싣는 스크립트,
이 화면의 것이 아닌 저장 이름. 옛 시연 화면(static/app.js·static/index.html)은 배포 목록에 없어야 한다.

  python -X utf8 -m unittest discover -s web -p "test_*.py"

걸렸을 때: 뜻한 변화면 바로 아래의 표(SOURCES·ALLOW·CALLS)를 고친다. 실패 문구가 파일·낱말·횟수를 알려 준다.
"""
import os
from pathlib import Path
import re
import sys
import unittest

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "web"))

import build  # noqa: E402

# 배포 목록(build.PAGES)의 원본 파일 — 화면 파일을 더하면 여기에도 더한다(그 파일도 아래 검사를 지난다)
SOURCES = ("static/style.css", "web/index.html", "web/viewer.css", "web/viewer.js", "web/worker.mjs")
OLD_DEMO = ("static/app.js", "static/index.html")          # 글자를 HTML로 넣는 옛 시연 화면 — 배포하지 않는다

BANNED = (                                                 # 배포 파일 어디에서도 0번 — 아래 ALLOW로는 풀지 않는다
    "innerHTML", "outerHTML", "insertAdjacentHTML", "document.write", "DOMParser", "createContextualFragment",
    "setHTML", "srcdoc",
    "eval(", "new Function", "javascript:",
    'setAttribute("on', "setAttribute('on",
    "serviceWorker", "import(", "importScripts", "RTCPeerConnection", "window.open", "location.assign", "location.replace",
    'rel="dns-prefetch"', 'rel="prefetch"', 'rel="prerender"',
)

WATCHED = (                                                # 횟수가 ALLOW와 같아야 하는 낱말(없는 파일·낱말은 0번)
    'rel="preconnect"', "<script>", "<script", "<base",
    "fetch(", "XMLHttpRequest", "WebSocket", "EventSource", "sendBeacon", "new Image", "new Worker", "url(", "@import",
    "localStorage", "sessionStorage", "indexedDB", "document.cookie",
    "location.", "navigator.", "postMessage", "createElement",
)

ALLOW = {                                                  # 파일 → 낱말 → 허용 횟수
    "web/index.html": {"<script": 1},
    "web/viewer.js": {"fetch(": 2, "new Worker": 1, "localStorage": 2, "indexedDB": 1, "location.": 2, "postMessage": 1,
                      "createElement": 1},
    "web/worker.mjs": {"fetch(": 1, "postMessage": 6},
}

STORES = ("localStorage", "sessionStorage", "indexedDB")
CALLS = {                                                  # 저장 호출의 꼴(저장 공간, 함수, 첫 인자) — 이름은 상수로만 넘긴다
    ("localStorage", "getItem", "WORDS_KEY"), ("localStorage", "setItem", "WORDS_KEY"), ("indexedDB", "open", "DB.name"),
}
NAME_HEAD = "macro-notes-viewer"                           # 이 화면의 저장 이름은 모두 이 머리로 시작한다
NAME_DEFS = (r'const WORDS_KEY = "([^"]*)"', r'const DB = \{ name: "([^"]*)"')
WORKER = 'new Worker(new URL("web/worker.mjs", location.href)'


def read(name):
    return (ROOT / name).read_text(encoding="utf-8", errors="replace")


class PagesTest(unittest.TestCase):
    def test_deploy_list_is_the_checked_one(self):
        self.assertEqual(sorted(set(build.PAGES.values())), list(SOURCES),
                         "배포 목록(web/build.py의 PAGES)이 바뀌었다 — 이 파일의 SOURCES에도 적어 같은 검사를 지나게 한다")
        for name in SOURCES:
            self.assertTrue((ROOT / name).is_file(), name)

    def test_old_demo_screen_is_not_deployed(self):
        for name in OLD_DEMO:
            self.assertNotIn(name, build.PAGES.values(), name + ": 옛 시연 화면은 배포 목록에 넣지 않는다")


JS_TOKEN = re.compile(                  # 문자열 셋 | 줄 주석 | 블록 주석 | 정규식 리터럴(앞 글자로 알아본다)
    r'"(?:\\.|[^"\\\n])*"|\'(?:\\.|[^\'\\\n])*\'|`(?:\\.|[^`\\])*`|//[^\n]*|/\*.*?\*/'
    r'|(?:[=(,:!&|?\[>]|\breturn)\s*/(?![/*])(?:\\.|\[(?:\\.|[^\]\\\n])*\]|[^/\\\n\[])+/', re.S)
CSS_NOTE = re.compile(r"/\*.*?\*/", re.S)
HTML_NOTE = re.compile(r"<!--.*?-->", re.S)
INLINE = re.compile(r"(<script\b[^>]*>)(.*?)(</script\s*>)", re.S | re.I)
SRC = re.compile(r"<script\b[^>]*?\bsrc\s*=\s*[\"']?([^\"'\s>]*)", re.I)
MODULE = re.compile(r'^\s*import\b[^\n;]*?["\']([^"\']+)["\']', re.M)


def js_code(text):
    """주석을 뗀 스크립트. 문자열과 정규식은 그대로 둔다 — 그 안의 //(주소)와 /*는 주석이 아니다."""
    return JS_TOKEN.sub(lambda m: "" if m.group(0)[:2] in ("//", "/*") else m.group(0), text)


def code_only(name, text):
    """주석을 뗀 글자 — 설명 글에 나온 낱말로 검사가 걸리지 않게. .html은 <!-- -->와 인라인 스크립트 안의 주석만 뗀다."""
    ext = os.path.splitext(name)[1]
    if ext == ".html":
        return INLINE.sub(lambda m: m.group(1) + js_code(m.group(2)) + m.group(3), HTML_NOTE.sub("", text))
    if ext == ".css":
        return CSS_NOTE.sub("", text)
    return js_code(text) if ext in (".js", ".mjs") else text


def banned_in(name, text):
    code = code_only(name, text)
    return [w for w in BANNED if w in code]


def outside_scripts(name, text):
    """이 사이트 밖에서 싣는 스크립트의 주소 — .html의 script src와 스크립트의 import 주소."""
    code = code_only(name, text)
    if name.endswith(".html"):
        return [s for s in SRC.findall(code) if ":" in s or s.startswith("/") or ".." in s.split("/")]
    return [s for s in MODULE.findall(code) if ":" in s or not s.startswith(("./", "../"))]


def allow_gaps(name, text, allow):
    """허용 표와 실제가 다른 곳 → 문구 목록(파일·낱말·허용 횟수·지금 횟수). 늘어도 줄어도 걸린다 — 표가 묵지 않게."""
    code, row = code_only(name, text), allow.get(name, {})
    return ["%s: `%s` 허용 %d번, 지금 %d번" % (name, w, row.get(w, 0), code.count(w))
            for w in WATCHED if row.get(w, 0) != code.count(w)]


class WordsTest(unittest.TestCase):
    def test_no_banned_word_in_deployed_files(self):
        found = [(name, w) for name in SOURCES for w in banned_in(name, read(name))]
        self.assertEqual(found, [], "금지 낱말(BANNED)이 배포 파일에 있다 — 글자는 텍스트로만 넣는다. 허용 표로는 풀지 않는다")

    def test_scripts_come_from_this_site_only(self):
        found = [(name, s) for name in SOURCES for s in outside_scripts(name, read(name))]
        self.assertEqual(found, [], "이 사이트 밖의 스크립트를 싣는다")
        self.assertEqual(code_only("web/viewer.js", read("web/viewer.js")).count(WORKER), 1, "작업 스레드는 web/worker.mjs 하나")

    def test_watched_words_match_the_table(self):
        gaps = [g for name in SOURCES for g in allow_gaps(name, read(name), ALLOW)]
        self.assertEqual(gaps, [], "허용 표(ALLOW)와 다르다 — 뜻한 변화면 숫자를 고친다(새 저장 이름·새 바깥 요청은 표에 한 줄)")
        self.assertLessEqual(set(ALLOW), set(SOURCES), "ALLOW에 배포하지 않는 파일이 있다")

    def test_planted_lines_are_caught(self):
        js = ('// innerHTML 설명\nconst u = "https://example.org/a"; /* eval( */ box.outerHTML = u; fetch(u); // fetch(\n'
              'const QUOTES = /["\'`]/; // srcdoc 설명 — 정규식 안의 따옴표가 문자열을 열지 않는다\n')
        self.assertEqual(banned_in("x.js", js), ["outerHTML"])
        self.assertEqual(allow_gaps("x.js", js, {}), ["x.js: `fetch(` 허용 0번, 지금 1번"])
        self.assertEqual(allow_gaps("x.js", js, {"x.js": {"fetch(": 1}}), [])
        self.assertEqual(banned_in("x.html", '<!-- srcdoc --><a href="javascript:void 0">x</a>'), ["javascript:"])
        page = ('<script src="viewer.js" defer></script><script src="https://cdn.example/x.js"></script>'
                '<script src="../y.js"></script>')
        self.assertEqual(outside_scripts("x.html", page), ["https://cdn.example/x.js", "../y.js"])
        module = 'import { a } from "../pyodide/pyodide.mjs";\nimport "https://cdn.example/m.mjs";\nimport b from "bare";\n'
        self.assertEqual(outside_scripts("x.mjs", module), ["https://cdn.example/m.mjs", "bare"])


STORE_CALL = re.compile(r"\b(%s)\.(\w+)\(\s*([\w.]*)" % "|".join(STORES))


def store_gaps(name, text):
    """저장 호출이 정해 둔 꼴(CALLS)이 아닌 곳 → 문구 목록. 저장 공간을 통째로 넘기거나 훑는 꼴도 걸린다."""
    code = code_only(name, text)
    calls = STORE_CALL.findall(code)
    out = ["%s: %s.%s(%s …)" % ((name,) + c) for c in calls if c not in CALLS]
    loose = sum(code.count(s) for s in STORES) - len(calls)
    return out + (["%s: 호출 꼴이 아닌 저장 공간 사용 %d곳" % (name, loose)] if loose else [])


class StoreTest(unittest.TestCase):
    def test_storage_calls_use_this_screens_names_only(self):
        gaps = [g for name in SOURCES for g in store_gaps(name, read(name))]
        self.assertEqual(gaps, [], "저장 호출이 정해 둔 꼴(CALLS)이 아니다 — 다른 화면의 저장 이름을 읽거나 저장 공간을 훑지 않는다")
        code = code_only("web/viewer.js", read("web/viewer.js"))
        for pattern in NAME_DEFS:
            found = re.findall(pattern, code)
            self.assertEqual(len(found), 1, pattern)
            self.assertTrue(found[0].startswith(NAME_HEAD), pattern + ": 저장 이름은 " + NAME_HEAD + "로 시작한다")

    def test_a_planted_storage_call_is_caught(self):
        ok = "localStorage.getItem(WORDS_KEY); localStorage.setItem(WORDS_KEY, v); indexedDB.open(DB.name, 1);\n"
        self.assertEqual(store_gaps("x.js", ok), [])
        self.assertEqual(store_gaps("x.js", 'localStorage.getItem("other.key");\n'), ["x.js: localStorage.getItem( …)"])
        self.assertEqual(store_gaps("x.js", "localStorage.removeItem(WORDS_KEY);\n"),
                         ["x.js: localStorage.removeItem(WORDS_KEY …)"])
        self.assertEqual(store_gaps("x.js", "const all = Object.keys(localStorage); sessionStorage[k] = 1;\n"),
                         ["x.js: 호출 꼴이 아닌 저장 공간 사용 2곳"])
        self.assertEqual(store_gaps("x.js", "// localStorage.clear() 는 쓰지 않는다\n"), [])


if __name__ == "__main__":
    unittest.main()
