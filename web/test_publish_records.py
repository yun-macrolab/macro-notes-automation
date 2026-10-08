"""publish_records.py 검사 — 합성 기록 워크북 폴더와 가짜 GitHub(메모리)로 확인한다(네트워크 불필요).

  python -X utf8 -m unittest discover -s web -p "test_*.py"
"""
import base64
import contextlib
import datetime as dt
import io
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unittest
from unittest.mock import patch
import urllib.error

ROOT = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "engine"), str(ROOT / "web")]

import openpyxl  # noqa: E402

from core import DemoStore  # noqa: E402
import publish_records as pub  # noqa: E402
import sheet_v2 as sv  # noqa: E402

TOKEN = "github_pat_TEST_ONLY_0123456789"


@contextlib.contextmanager
def token_env():
    os.environ["MACRO_NOTES_TOKEN"] = TOKEN
    try:
        yield
    finally:
        os.environ.pop("MACRO_NOTES_TOKEN", None)


class Keyboard(io.StringIO):
    """사람이 터미널에서 직접 돌린 실행처럼 보이게 하는 입력."""

    def isatty(self):
        return True


class FakeGitHub:
    """Git Data API 중 publish_records가 쓰는 부분만 흉내 낸다."""

    def __init__(self, records=None):
        self.files = dict(records or {})        # {"records/x.json": 글}
        self.head = "c0"
        self.calls = []
        self.fail_patch = 0
        self.headers = []

    def __call__(self, request, timeout=None):
        method, url = request.get_method(), request.full_url
        path = url.split("/repos/yun-macrolab/macro-notes-automation", 1)[1]
        body = json.loads(request.data) if request.data else None
        self.calls.append((method, path.split("?")[0]))
        self.headers.append(dict(request.header_items()))
        if method == "GET" and path == "/git/ref/heads/main":
            out = {"object": {"sha": self.head}}
        elif method == "GET" and path.startswith("/git/commits/"):
            out = {"tree": {"sha": "t-" + path.rsplit("/", 1)[1]}}
        elif method == "GET" and path.startswith("/git/trees/"):
            out = {"tree": [{"path": name, "type": "blob", "sha": "b:" + name} for name in self.files]
                   + [{"path": "records", "type": "tree", "sha": "tr"}], "truncated": False}
        elif method == "GET" and path.startswith("/git/blobs/b:"):
            text = self.files[path[len("/git/blobs/b:"):]]
            out = {"content": base64.b64encode(text.encode("utf-8")).decode(), "encoding": "base64"}
        elif method == "POST" and path == "/git/trees":
            self.pending = {entry["path"]: entry["content"] for entry in body["tree"]}
            out = {"sha": "t-new"}
        elif method == "POST" and path == "/git/commits":
            assert body["author"]["email"] == "yun-macrolab@users.noreply.github.com"
            assert body["committer"] == body["author"]
            self.message = body["message"]
            out = {"sha": "c-new-" + str(len(self.calls))}
        elif method == "PATCH" and path == "/git/refs/heads/main":
            if self.fail_patch:
                self.fail_patch -= 1
                self.head = "c-moved"
                raise urllib.error.HTTPError(request.full_url, 422, "Update is not a fast forward", {},
                                             io.BytesIO(b'{"message": "Update is not a fast forward"}'))
            self.files.update(self.pending)
            self.head = body["sha"]
            out = {"object": {"sha": body["sha"]}}
        else:
            raise AssertionError(f"unexpected call {method} {path}")
        return contextlib.closing(io.BytesIO(json.dumps(out).encode("utf-8")))


class Publish(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory(prefix="publish_")
        store = DemoStore(cls.tmp.name)
        run_id = store.run()["run"]["id"]
        cls.demo = Path(cls.tmp.name) / run_id / "demo.xlsx"

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def setUp(self):
        self.work = Path(tempfile.mkdtemp(dir=self.tmp.name))
        self.folder = self.work / "주간"
        (self.folder / "2026-09").mkdir(parents=True)
        original = self.folder / "2026-09" / "기록_0921.xlsx"
        shutil.copy(self.demo, original)
        old = original.stat().st_mtime - 3600
        os.utime(original, (old, old))                                   # 복사본이 원본보다 늦게 저장된 상황
        self.variant("기록_0914.xlsx", {"I1": dt.datetime(2026, 9, 15), sv.DRAFT_CELL: None, "B26": None, "X7": None})
        self.variant("기록_v1.xlsx", {sv.SCHEMA_CELL: None, "B26": None, "X7": None})
        self.variant("기록_0921 - 복사본.xlsx", {"D3": "복사본에서 고친 결론"})
        (self.folder / "백업").mkdir()
        self.variant("백업/기록_0921.xlsx", {"D3": "백업 폴더의 결론"})
        (self.folder / "~$기록_0921.xlsx").write_bytes(b"lock")
        self.home, self.urlopen, self.wait = pub.HOME, pub.URLOPEN, pub.WAIT
        pub.HOME = self.work / "home"
        self.waits = []
        pub.WAIT = self.waits.append                                     # 가짜 기다림 — 검사는 실제로 기다리지 않는다
        os.environ.pop("MACRO_NOTES_TOKEN", None)

    def tearDown(self):
        pub.HOME, pub.URLOPEN, pub.WAIT = self.home, self.urlopen, self.wait

    def variant(self, name, cells):
        wb = openpyxl.load_workbook(self.demo)
        for ref, value in cells.items():
            wb["Analysis"][ref] = value
        wb.save(self.folder / name)
        old = (self.folder / name).stat().st_mtime - 1800                # 저장한 지 10분이 지난 파일(원본보다는 늦게)
        os.utime(self.folder / name, (old, old))

    def lock(self, name=str(Path("2026-09") / "기록_0921.xlsx")):
        """엑셀이 열어 둔 것처럼 '~$' 잠금 파일을 둔다."""
        file = self.folder / name
        lock = file.with_name("~$" + file.name)
        lock.write_bytes(b"lock")
        return lock

    def marker(self, age=0):
        """원본 작업이 끝났다는 표식을 age초 전에 쓴 것으로 둔다."""
        pub.HOME.mkdir(parents=True, exist_ok=True)
        marker = pub.HOME / "source-done"
        marker.write_text("2026-10-08T13:07:00", encoding="ascii")
        when = time.time() - age
        os.utime(marker, (when, when))
        return marker

    def log_text(self):
        return (pub.HOME / "publish.log").read_text(encoding="utf-8")

    def run_main(self, *args, github=None):
        pub.URLOPEN = github or FakeGitHub()
        out = io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(out):
            code = pub.main(["--folder", str(self.folder), *args])
        return code, out.getvalue()

    def test_collect(self):
        rows, picked = pub.collect(self.folder)
        self.assertEqual(sorted(picked), ["2026-09-14.json", "2026-09-21.json"])
        self.assertEqual(len(rows), 5)                                   # 잠금 파일(~$)은 뺀다
        self.assertEqual({(row["label"], row["reason"]) for row in rows if not row["ok"]},
                         {("기록_0921 - 복사본.xlsx", "copy"), (str(Path("백업") / "기록_0921.xlsx"), "copy"),
                          ("기록_v1.xlsx", "v1")})
        # 늦게 저장한 복사본·백업이 아니라 원본을 쓴다
        self.assertEqual(picked["2026-09-21.json"]["label"], str(Path("2026-09") / "기록_0921.xlsx"))
        self.assertIsNone(picked["2026-09-14.json"]["record"]["state"]["confirmed"])

    def test_root_folder_name_is_not_checked(self):
        root = self.work / "백업해 둔 주간기록"                           # 고른 폴더 자체의 이름은 보지 않는다
        shutil.copytree(self.folder, root)
        _, picked = pub.collect(root)
        self.assertEqual(sorted(picked), ["2026-09-14.json", "2026-09-21.json"])

    def test_uploads_new_weeks_in_one_commit(self):
        github = FakeGitHub()
        with token_env():
            code, out = self.run_main("--yes", github=github)
        self.assertEqual(code, 0, out)
        self.assertEqual(sorted(github.files), ["records/2026-09-14.json", "records/2026-09-21.json"])
        self.assertEqual([c for c in github.calls if c[0] != "GET"],
                         [("POST", "/git/trees"), ("POST", "/git/commits"), ("PATCH", "/git/refs/heads/main")])
        self.assertIn("records: 공개 기록 2주 올림", github.message)
        self.assertNotIn("복사본에서 고친 결론", github.files["records/2026-09-21.json"])
        self.assertNotIn("백업 폴더의 결론", github.files["records/2026-09-21.json"])
        self.assertTrue(all(h.get("Authorization") == f"Bearer {TOKEN}" for h in github.headers))
        self.assertNotIn(TOKEN, out)
        self.assertNotIn(TOKEN, (pub.HOME / "publish.log").read_text(encoding="utf-8"))
        self.assertEqual(json.loads((pub.HOME / "publish.json").read_text(encoding="utf-8"))["folder"], str(self.folder))

    def test_skips_unchanged_weeks(self):
        _, picked = pub.collect(self.folder)
        older = json.loads(picked["2026-09-21.json"]["text"])
        older["generated_at"] = "2026-01-01T00:00:00+00:00"              # 만든 시각만 다르면 같은 공개본
        github = FakeGitHub({"records/2026-09-21.json": json.dumps(older, ensure_ascii=False),
                             "records/README.md": "# records"})
        with token_env():
            code, out = self.run_main("--yes", github=github)
        self.assertEqual(code, 0, out)
        self.assertEqual(sorted(github.pending), ["records/2026-09-14.json"])   # 바뀐 주만
        self.assertIn("그대로", out)
        # 모두 올라간 뒤 다시 돌리면 아무것도 쓰지 않는다
        github.calls.clear()
        with token_env():
            code, out = self.run_main("--yes", github=github)
        self.assertEqual(code, 0)
        self.assertIn("올릴 것이 없습니다", out)
        self.assertTrue(all(method == "GET" for method, _ in github.calls))

    def test_dry_run_and_missing_token(self):
        github = FakeGitHub()
        code, out = self.run_main("--dry-run", github=github)
        self.assertEqual(code, 0)
        self.assertIn("미리 보기", out)
        self.assertTrue(all(method == "GET" for method, _ in github.calls))
        code, out = self.run_main("--yes", github=FakeGitHub())      # 예약 작업인데 토큰이 없으면 멈춘다
        self.assertEqual(code, 1)
        self.assertIn("토큰이 없습니다", out)

    def test_retries_once_when_main_moved(self):
        github = FakeGitHub()
        github.fail_patch = 1
        with token_env():
            code, out = self.run_main("--yes", github=github)
        self.assertEqual(code, 0, out)
        self.assertEqual([c for c in github.calls if c[0] == "PATCH"], [("PATCH", "/git/refs/heads/main")] * 2)
        self.assertEqual(sorted(github.files), ["records/2026-09-14.json", "records/2026-09-21.json"])

    def test_unexpected_error_is_logged(self):
        collect = pub.collect
        pub.collect = lambda *_: (_ for _ in ()).throw(PermissionError("폴더를 읽을 권한 없음"))
        try:
            code, out = self.run_main("--yes", github=FakeGitHub())
        finally:
            pub.collect = collect
        self.assertEqual(code, 1)                                       # 예약 작업 창(-Yes)이 실패 알림을 띄운다
        self.assertIn("예상하지 못한 문제", (pub.HOME / "publish.log").read_text(encoding="utf-8"))

    def test_out_folder(self):
        pub.HOME.mkdir(parents=True)
        (pub.HOME / "publish.json").write_text(json.dumps({"source_task": "DailyProducer"}), encoding="utf-8")
        code, out = self.run_main("--out", str(self.work / "out"), github=FakeGitHub())
        self.assertEqual(code, 0, out)
        self.assertEqual(sorted(p.name for p in (self.work / "out").iterdir()), ["2026-09-14.json", "2026-09-21.json"])
        self.assertEqual(pub.read_json(pub.HOME / "publish.json")["source_task"], "DailyProducer")

    def test_root_pattern_excludes_backups_and_other_projects(self):
        self.variant("기록_20260921-0927.xlsx", {"I1": "=TODAY()"})
        _, picked = pub.collect(self.folder, pattern="*_????????-????.xlsx")
        self.assertEqual(list(picked), ["2026-09-21.json"])

    def test_open_or_recently_saved_workbook_is_held_and_other_weeks_upload(self):
        file = self.folder / "2026-09" / "기록_0921.xlsx"
        lock = self.lock()
        github = FakeGitHub()
        with token_env():
            code, out = self.run_main("--yes", github=github)
        self.assertEqual(code, 0, out)                                   # 보류는 실패가 아니다
        self.assertEqual(sorted(github.files), ["records/2026-09-14.json"])   # 보류된 주만 빠진다
        self.assertEqual(self.waits, [pub.RECHECK_WAIT] * 3)             # 같은 실행 안에서 5분 뒤 다시 보기 3번
        self.assertEqual(pub.RECHECK_WAIT, 300)
        self.assertIn("→ 보류", out)
        log = self.log_text()
        self.assertIn("5분 뒤 다시 봄(3/3)", log)
        self.assertIn("올림 1주(2026-09-14) · 커밋 ", log.splitlines()[-1])
        self.assertTrue(log.splitlines()[-1].endswith("· 건너뜀 3 · 보류 1개(엑셀에서 열려 있음·저장 직후)"))
        self.assertNotIn("실패", log)
        lock.unlink()
        os.utime(file, None)
        rows, picked = pub.collect(self.folder)
        self.assertEqual([row["label"] for row in rows if row.get("reason") == "busy"], [str(Path("2026-09") / "기록_0921.xlsx")])
        self.assertEqual(sorted(picked), ["2026-09-14.json"])

    def test_only_held_files_exit_zero_and_are_logged(self):
        self.lock()
        self.lock("기록_0914.xlsx")
        github = FakeGitHub()
        with token_env():
            code, out = self.run_main("--yes", github=github)
        self.assertEqual(code, 0, out)
        self.assertFalse(github.calls)
        self.assertTrue(self.log_text().splitlines()[-1].endswith("올릴 공개본 없음 · 보류 2개(엑셀에서 열려 있음·저장 직후)"))
        self.assertNotIn("실패", self.log_text())

    def test_unchanged_weeks_with_held_file_log_line(self):
        _, picked = pub.collect(self.folder)
        github = FakeGitHub({"records/2026-09-14.json": picked["2026-09-14.json"]["text"]})
        self.lock()
        with token_env(), patch.object(pub, "RECHECKS", 0):
            code, out = self.run_main("--yes", github=github)
        self.assertEqual(code, 0, out)
        self.assertTrue(all(method == "GET" for method, _ in github.calls))
        # 건너뜀(복사본 2·v1 1)에는 보류를 넣지 않는다
        self.assertTrue(self.log_text().splitlines()[-1].endswith(
            "바뀐 주 없음 · 공개본 1주 · 건너뜀 3 · 보류 1개(엑셀에서 열려 있음·저장 직후)"))

    def test_blocked_file_stops_everything_even_with_held_file(self):
        self.variant("invalid.xlsx", {"D3": "private@example.com"})
        self.lock()
        github = FakeGitHub()
        with token_env():
            code, _ = self.run_main("--yes", github=github)
        self.assertEqual(code, 1)
        self.assertFalse(github.calls)
        self.assertEqual(self.waits, [])                                 # 실패가 있으면 기다리지 않고 멈춘다
        self.assertIn("실패", self.log_text())

    def test_held_week_is_not_taken_from_another_file(self):
        self.variant("기록_0921_둘째.xlsx", {"D3": "같은 주의 다른 파일"})
        _, picked = pub.collect(self.folder)
        self.assertEqual(sorted(picked), ["2026-09-14.json", "2026-09-21.json"])
        self.lock()                                                      # 보류한 파일이 그 주의 최신일 수 있다
        rows, picked = pub.collect(self.folder)
        self.assertEqual(sorted(picked), ["2026-09-14.json"])
        self.assertEqual([row["week"]["start"] for row in rows if row.get("reason") == "busy"], ["2026-09-21"])

    def test_held_copy_or_backup_is_not_counted_as_held(self):
        self.lock("기록_0921 - 복사본.xlsx")
        rows, picked = pub.collect(self.folder)
        self.assertFalse([row for row in rows if row.get("reason") == "busy"])
        self.assertEqual(sorted(picked), ["2026-09-14.json", "2026-09-21.json"])

    def test_files_left_out_anyway_are_not_held_when_changed_during_read(self):
        original = pub.workbook_view.public_from_workbook

        def changing(path, *args):                                       # 복사본과 예전 형식 파일이 읽는 중 바뀐다
            result = original(path, *args)
            if Path(path).name in ("기록_0921 - 복사본.xlsx", "기록_v1.xlsx"):
                os.utime(path, (time.time() - 900, time.time() - 900))
            return result

        github = FakeGitHub()
        with token_env(), patch.object(pub.workbook_view, "public_from_workbook", side_effect=changing):
            code, out = self.run_main("--yes", github=github)
        self.assertEqual(code, 0, out)
        self.assertEqual(self.waits, [])                                 # 올릴 일이 없는 파일 때문에 기다리지 않는다
        self.assertEqual(sorted(github.files), ["records/2026-09-14.json", "records/2026-09-21.json"])
        self.assertNotIn("보류", self.log_text())
        self.assertTrue(self.log_text().splitlines()[-1].endswith("· 건너뜀 3"))

    def test_workbook_open_in_excel_is_not_read(self):
        name = "기록_20260921-0927.xlsx"
        self.variant(name, {"D3": "엑셀에서 고치는 중"})
        self.lock(name)
        opened, load = [], openpyxl.load_workbook

        def watching(path, *args, **kwargs):
            opened.append(Path(path).name)
            return load(path, *args, **kwargs)

        with patch.object(openpyxl, "load_workbook", side_effect=watching):
            rows, picked = pub.collect(self.folder)
        self.assertIn("기록_0914.xlsx", opened)
        self.assertNotIn(name, opened)                                   # 읽는 사이 엑셀의 저장과 겹치지 않게 열지도 않는다
        self.assertEqual([(row["week"]["start"], row["held_by"]) for row in rows if row.get("reason") == "busy"],
                         [("2026-09-21", "open")])
        self.assertEqual(sorted(picked), ["2026-09-14.json"])

    def test_save_time_in_the_future_is_not_held(self):
        file = self.folder / "기록_0914.xlsx"
        later = time.time() + 30 * 86400                                 # 시계가 어긋나 저장 시각이 미래로 찍힌 파일
        os.utime(file, (later, later))
        rows, picked = pub.collect(self.folder)
        self.assertFalse([row for row in rows if row.get("reason") == "busy"])   # 영영 보류되지 않는다
        self.assertEqual(sorted(picked), ["2026-09-14.json", "2026-09-21.json"])
        soon = time.time() + 60                                          # 시계 오차만큼 앞선 것은 저장 직후로 본다
        os.utime(file, (soon, soon))
        _, picked = pub.collect(self.folder)
        self.assertEqual(sorted(picked), ["2026-09-21.json"])

    def test_lock_left_for_days_is_named_in_log(self):
        lock = self.lock()
        with token_env(), patch.object(pub, "RECHECKS", 0):
            code, out = self.run_main("--yes", github=FakeGitHub())
            self.assertEqual(code, 0, out)
            self.assertNotIn("잠금 파일", self.log_text())                # 방금 연 파일은 따로 적지 않는다
            old = time.time() - 3 * 86400 - 600                          # 엑셀이 죽으며 남긴 잠금 파일은 그 주를 계속 보류한다
            os.utime(lock, (old, old))
            code, out = self.run_main("--yes", github=FakeGitHub())
        self.assertEqual(code, 0, out)
        self.assertTrue(self.log_text().splitlines()[-1].endswith(
            "· 보류 1개(엑셀에서 열려 있음·저장 직후) · 잠금 파일이 하루 넘게 남은 주: 2026-09-21(3일 전부터)"))
        self.assertIn("→ 보류 · 엑셀에서 열려 있음(잠금 파일 3일 전부터", out)

    def test_held_line_says_why_and_when(self):
        self.lock()
        file = self.folder / "기록_0914.xlsx"
        now = time.time()
        os.utime(file, (now - 60, now - 60))                             # 1분 전에 저장하고 엑셀은 닫았다
        with patch.object(pub.time, "time", return_value=now):
            code, out = self.run_main("--dry-run", github=FakeGitHub())
        self.assertEqual(code, 0, out)
        lines = {line.split("  ")[2]: line for line in out.splitlines() if "→ 보류" in line}
        self.assertTrue(lines[str(Path("2026-09") / "기록_0921.xlsx")].endswith("→ 보류 · 엑셀에서 열려 있음"))
        self.assertTrue(lines["기록_0914.xlsx"].endswith("→ 보류 · 저장 직후(약 9분 뒤부터 올릴 수 있음)"))

    def test_token_problem_shows_before_waiting(self):
        self.lock()
        code, out = self.run_main("--yes", github=FakeGitHub())          # 토큰이 없으면 기다려도 올리지 못한다
        self.assertEqual(code, 1)
        self.assertIn("토큰이 없습니다", out)
        self.assertEqual(self.waits, [])
        with patch.object(pub, "read_token", side_effect=pub.PublishError("GitHub CLI 로그인이 필요합니다.")):
            code, out = self.run_main("--yes", github=FakeGitHub())
        self.assertEqual(code, 1)
        self.assertIn("로그인이 필요합니다", out)
        self.assertEqual(self.waits, [])                                 # 인증을 못 읽으면 15분 기다린 뒤가 아니라 바로 멈춘다

    def test_settle_is_ten_minutes_after_last_save(self):
        folder = self.work / "한 주"
        folder.mkdir()
        shutil.copy(self.demo, folder / "기록.xlsx")
        saved = (folder / "기록.xlsx").stat().st_mtime
        self.assertEqual((pub.SETTLE, pub.SETTLE_AFTER_SOURCE), (600, 15))
        # 평소엔 10분, 원본 작업 직후(표식)엔 15초
        for settle, age, held in (({}, 599, True), ({}, 601, False),
                                  ({"settle": pub.SETTLE_AFTER_SOURCE}, 14, True),
                                  ({"settle": pub.SETTLE_AFTER_SOURCE}, 16, False)):
            with patch.object(pub.time, "time", return_value=saved + age):
                rows, picked = pub.collect(folder, **settle)
            self.assertEqual([row.get("reason") for row in rows], ["busy"] if held else [None], age)
            self.assertEqual(len(picked), 0 if held else 1, age)

    def test_source_done_marker_shortens_settle_once_and_is_removed(self):
        file = self.folder / "기록_0914.xlsx"
        saved = time.time() - 60                                         # 원본 작업이 1분 전에 저장한 파일
        os.utime(file, (saved, saved))
        marker = self.marker()
        github = FakeGitHub()
        with token_env():
            code, out = self.run_main("--yes", github=github)
        self.assertEqual(code, 0, out)
        self.assertEqual(sorted(github.files), ["records/2026-09-14.json", "records/2026-09-21.json"])
        self.assertEqual(self.waits, [])
        self.assertFalse(marker.exists())                                # 이번 실행에만 쓰고 지운다
        self.assertIn("원본 작업 직후", self.log_text())
        # 표식이 없는 다음 실행은 같은 파일을 10분 규칙으로 보류한다
        rows, _ = pub.collect(self.folder)
        self.assertEqual([row["label"] for row in rows if row.get("reason") == "busy"], ["기록_0914.xlsx"])
        # 표식이 있어도 저장한 지 15초가 안 된 파일은 보류한다(미리 보기·--out은 표식을 지우지 않는다)
        marker = self.marker()
        os.utime(file, None)
        code, out = self.run_main("--out", str(self.work / "out"), github=FakeGitHub())
        self.assertEqual(code, 0, out)
        self.assertEqual(sorted(p.name for p in (self.work / "out").iterdir()), ["2026-09-21.json"])
        self.assertTrue(marker.exists())

    def test_stale_source_done_marker_is_ignored_and_removed(self):
        file = self.folder / "기록_0914.xlsx"
        saved = time.time() - 60
        os.utime(file, (saved, saved))
        marker = self.marker(age=301)                                    # 5분이 넘은 표식
        github = FakeGitHub()
        with token_env():
            code, out = self.run_main("--yes", github=github)
        self.assertEqual(code, 0, out)
        self.assertEqual(sorted(github.files), ["records/2026-09-21.json"])
        self.assertEqual(self.waits, [pub.RECHECK_WAIT] * 3)
        self.assertFalse(marker.exists())
        self.assertNotIn("원본 작업 직후", self.log_text())
        self.assertEqual(self.log_text().count("5분 넘게 묵어 무시"), 1)  # 왜 바로 올라가지 않았는지 로그에 남긴다

    def test_source_done_marker_is_fresh_up_to_five_minutes(self):
        file = self.folder / "기록_0914.xlsx"
        now = time.time()
        os.utime(file, (now - 60, now - 60))
        marker = self.marker()
        for age, fresh in ((299, True), (301, False)):
            os.utime(marker, (now - age, now - age))
            self.assertIs(pub.source_done(now), fresh, age)
        os.utime(marker, (now - 299, now - 299))
        github = FakeGitHub()
        with token_env(), patch.object(pub.time, "time", return_value=now):
            code, out = self.run_main("--yes", github=github)
        self.assertEqual(code, 0, out)
        self.assertEqual(self.waits, [])
        self.assertEqual(sorted(github.files), ["records/2026-09-14.json", "records/2026-09-21.json"])

    def test_source_done_marker_stays_valid_for_the_whole_run(self):
        self.lock()                                                      # 풀리지 않는 보류 → 세 번 다시 본다
        clock = [time.time()]
        self.marker(age=200)

        def wait(seconds):
            self.waits.append(seconds)
            clock[0] += seconds

        pub.WAIT = wait
        settles, collect = [], pub.collect

        def watching(folder, words, pattern, settle):
            settles.append(settle)
            return collect(folder, words, pattern, settle)

        with token_env(), patch.object(pub, "collect", side_effect=watching), \
                patch.object(pub.time, "time", side_effect=lambda: clock[0]):
            code, out = self.run_main("--yes", github=FakeGitHub())
        self.assertEqual(code, 0, out)
        self.assertEqual(self.waits, [300, 300, 300])
        self.assertEqual(settles, [pub.SETTLE_AFTER_SOURCE] * 4)         # 실행 시작 5분 전부터의 표식은 끝까지 쓴다

    def test_overlapping_run_leaves_marker_for_the_waiting_run(self):
        file = self.folder / "기록_0914.xlsx"
        saved = time.time() - 60
        os.utime(file, (saved, saved))
        seen = []

        def wait(seconds):                                               # 기다리는 사이 원본 작업이 끝나 표식을 쓰고 게시를 한 번 더 불렀다
            self.waits.append(seconds)
            marker = self.marker()
            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
                seen.append(pub.main(["--folder", str(self.folder), "--yes"]))
            seen.append(marker.exists())

        pub.WAIT = wait
        github = FakeGitHub()
        with token_env():
            code, out = self.run_main("--yes", github=github)
        self.assertEqual(seen, [1, True])                                # 겹친 실행은 멈추고 표식은 기다리던 실행 몫으로 남긴다
        self.assertEqual(code, 0, out)
        self.assertEqual(self.waits, [300])
        self.assertEqual(sorted(github.files), ["records/2026-09-14.json", "records/2026-09-21.json"])
        self.assertIn("실패: 다른 업로드가 실행 중입니다(보류된 파일을 기다리는 중이면 최대 15분)", self.log_text())

    def test_source_done_marker_written_while_waiting_is_used_on_next_look(self):
        file = self.folder / "기록_0914.xlsx"
        saved = time.time() - 60
        os.utime(file, (saved, saved))

        def wait(seconds):                                               # 기다리는 사이 원본 작업이 끝나 표식을 썼다
            self.waits.append(seconds)
            self.marker()

        pub.WAIT = wait
        github = FakeGitHub()
        with token_env():
            code, out = self.run_main("--yes", github=github)
        self.assertEqual(code, 0, out)
        self.assertEqual(self.waits, [300])
        self.assertEqual(sorted(github.files), ["records/2026-09-14.json", "records/2026-09-21.json"])
        self.assertFalse((pub.HOME / "source-done").exists())

    def test_source_done_marker_is_removed_after_failed_run(self):
        self.variant("invalid.xlsx", {"D3": "private@example.com"})
        marker = self.marker()
        with token_env():
            code, _ = self.run_main("--yes", github=FakeGitHub())
        self.assertEqual(code, 1)
        self.assertFalse(marker.exists())

    def test_recheck_uploads_when_released_on_second_look(self):
        lock = self.lock()

        def wait(seconds):
            self.waits.append(seconds)
            lock.unlink()                                                # 기다리는 사이 엑셀을 닫았다

        pub.WAIT = wait
        github = FakeGitHub()
        with token_env():
            code, out = self.run_main("--yes", github=github)
        self.assertEqual(code, 0, out)
        self.assertEqual(self.waits, [300])
        self.assertEqual(sorted(github.files), ["records/2026-09-14.json", "records/2026-09-21.json"])
        self.assertNotIn("보류", self.log_text().splitlines()[-1])

    def test_no_waiting_for_preview_out_or_manual_run(self):
        self.lock()
        marker = self.marker()
        with token_env():
            code, out = self.run_main("--dry-run", "--yes", github=FakeGitHub())
            self.assertEqual(code, 0, out)
            self.assertIn("미리 보기", out)
            code, out = self.run_main("--out", str(self.work / "out"), "--yes", github=FakeGitHub())
            self.assertEqual(code, 0, out)
            self.assertEqual(sorted(p.name for p in (self.work / "out").iterdir()), ["2026-09-14.json"])
            self.assertTrue(marker.exists())                             # 올리지 않는 실행은 표식을 남겨 둔다
            github = FakeGitHub()
            with patch.object(pub.sys, "stdin", Keyboard("n\n")):        # 사람이 직접 돌려 '올릴까요?'에 n
                code, out = self.run_main(github=github)
            self.assertEqual(code, 0, out)
            self.assertIn("올리지 않았습니다", out)
            self.assertTrue(all(method == "GET" for method, _ in github.calls))
        self.assertEqual(self.waits, [])

    def test_changed_during_read_is_rejected(self):
        original = pub.workbook_view.public_from_workbook
        def changing(path, *args):
            result = original(path, *args)
            os.utime(path, None)
            return result
        with patch.object(pub.workbook_view, "public_from_workbook", side_effect=changing):
            rows, picked = pub.collect(self.folder)
        self.assertFalse(picked)
        self.assertTrue(any(row.get("reason") == "busy" for row in rows))

    def test_week_of_file_changed_during_read_is_not_taken_from_another_file(self):
        self.variant("기록_0921_둘째.xlsx", {"D3": "같은 주의 다른 파일"})
        original = pub.workbook_view.public_from_workbook

        def changing(path, *args):
            result = original(path, *args)
            if path.endswith(str(Path("2026-09") / "기록_0921.xlsx")):
                os.utime(path, None)
            return result

        with patch.object(pub.workbook_view, "public_from_workbook", side_effect=changing):
            rows, picked = pub.collect(self.folder)
        self.assertEqual(sorted(picked), ["2026-09-14.json"])
        self.assertEqual([row["week"]["start"] for row in rows if row.get("reason") == "busy"], ["2026-09-21"])

    def test_gh_auth_captured_not_saved(self):
        with patch.object(pub.subprocess, "run") as run:
            run.return_value.returncode = 0
            run.return_value.stdout = TOKEN + "\n"
            self.assertEqual(pub.read_token("gh"), TOKEN)
            self.assertTrue(run.call_args.kwargs["capture_output"])
        self.assertFalse((pub.HOME / "token.txt").exists())

    def test_truncated_listing_is_rejected(self):
        github = pub.GitHub(pub.DEFAULT_REPO)
        with patch.object(github, "call", side_effect=[{"object": {"sha": "h"}},
                          {"tree": {"sha": "t"}}, {"tree": [], "truncated": True}]):
            with self.assertRaises(pub.PublishError):
                github.snapshot()

    def test_get_retries_transient_error_but_not_auth_failure(self):
        url = "https://api.github.com"
        fake = FakeGitHub()
        with patch.object(pub, "URLOPEN", side_effect=[urllib.error.URLError("offline"), fake(urllib.request.Request(url + "/repos/" + pub.DEFAULT_REPO + "/git/ref/heads/main"))]), patch.object(pub.time, "sleep"):
            self.assertEqual(pub.GitHub(pub.DEFAULT_REPO).call("GET", "/git/ref/heads/main")["object"]["sha"], "c0")
        with patch.object(pub, "URLOPEN", side_effect=urllib.error.HTTPError(url, 401, "bad auth", {}, io.BytesIO(b'{}'))) as call:
            with self.assertRaises(pub.PublishError):
                pub.GitHub(pub.DEFAULT_REPO).call("GET", "/git/ref/heads/main")
            self.assertEqual(call.call_count, 1)

    def test_lock_rejects_overlap_and_releases(self):
        with pub.publishing_lock():
            with self.assertRaises(pub.PublishError):
                with pub.publishing_lock():
                    self.fail("overlap")
        with pub.publishing_lock():
            pass

    def test_invalid_week_and_private_content_stop_before_github(self):
        for cells in ({"I1": None}, {"D3": "private@example.com"}):
            self.variant("invalid.xlsx", cells)
            github = FakeGitHub()
            code, _ = self.run_main("--yes", github=github)
            self.assertEqual(code, 1)
            self.assertFalse(github.calls)


# Publish-Records.ps1을 가짜 예약 작업 상태·가짜 기다림으로 한 번 돌린다(임시 폴더의 사본 + 흉내 publish_records.py).
HARNESS = r"""param([string]$Script, [string]$HomeDir, [string]$Seen)
$ErrorActionPreference = 'Stop'
Set-Variable -Name HOME -Value $HomeDir -Force
$global:queue = New-Object System.Collections.Queue
$Seen.Split(',') | ForEach-Object { $global:queue.Enqueue($_) }
$global:slept = @()
function Start-Sleep { param($Seconds) $global:slept += $Seconds }
function Get-ScheduledTask {
    param($TaskName, $ErrorAction)
    $state = if ($global:queue.Count -gt 1) { $global:queue.Dequeue() } else { $global:queue.Peek() }
    [pscustomobject]@{ State = $state }
}
& $Script -Yes
Write-Output "code=$LASTEXITCODE slept=$($global:slept -join '+')"
"""
STUB = 'import pathlib, sys\npathlib.Path(__file__).with_name("ran.txt").write_text(" ".join(sys.argv[1:]), encoding="ascii")\n'


class Launcher(unittest.TestCase):
    """Publish-Records.ps1 — 예약 기본값과, 원본 작업이 끝나기를 잠깐 기다리는 부분."""
    script = ROOT / "Publish-Records.ps1"

    def test_schedule_defaults(self):
        raw = self.script.read_bytes()
        self.assertTrue(all(byte < 128 for byte in raw))                 # 한글이 들어가면 BOM 없이 깨진다
        text = raw.decode("ascii")
        self.assertIn("[string]$At = '12:00', [int]$EveryMinutes = 360", text)   # 12:00 · 18:00 · 00:00 · 06:00
        limit = int(re.search(r"-ExecutionTimeLimit \(New-TimeSpan -Minutes (\d+)\)", text)[1])
        looks, pause = map(int, re.search(r"\$look -lt (\d+); \$look\+\+\) \{\s+Start-Sleep -Seconds (\d+)", text).groups())
        self.assertEqual((limit, looks * pause), (30, 30))
        # 원본 작업 기다림 + 다시 보기 3번을 담고도 남아야 한다(실행이 중간에 끊기지 않게)
        self.assertGreater(limit * 60, looks * pause + pub.RECHECK_WAIT * pub.RECHECKS + 600)

    @unittest.skipUnless(os.name == "nt" and shutil.which("powershell"), "Windows PowerShell에서만")
    def test_waits_briefly_for_source_task_before_deferring(self):
        with tempfile.TemporaryDirectory(prefix="launcher_") as tmp:
            tmp = Path(tmp)
            shutil.copy(self.script, tmp / "Publish-Records.ps1")
            (tmp / "publish_records.py").write_text(STUB, encoding="ascii")      # 진짜 게시는 돌지 않는다
            (tmp / "harness.ps1").write_text(HARNESS, encoding="ascii")
            home = tmp / "home" / ".macro-notes"
            home.mkdir(parents=True)
            (home / "publish.json").write_text(json.dumps({"source_task": "FakeProducer"}), encoding="ascii")
            env = {**os.environ, "PATH": str(Path(sys.executable).parent) + os.pathsep + os.environ["PATH"]}

            def launch(seen):
                (tmp / "ran.txt").unlink(missing_ok=True)
                result = subprocess.run(
                    ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(tmp / "harness.ps1"),
                     "-Script", str(tmp / "Publish-Records.ps1"), "-HomeDir", str(tmp / "home"), "-Seen", seen],
                    capture_output=True, text=True, timeout=120, env=env)
                log = home / "publish.log"
                return result.stdout.strip().splitlines()[-1:], (tmp / "ran.txt").exists(), \
                    log.read_text(encoding="utf-8-sig").count("deferred") if log.exists() else 0

            # 원본 작업이 게시를 깨우고 막 끝나는 중 — 두 번째 볼 때 끝나 있으면 그대로 올린다
            self.assertEqual(launch("Running,Ready"), (["code=0 slept=2+2"], True, 0))
            self.assertEqual((tmp / "ran.txt").read_text(encoding="ascii"), "--yes")
            self.assertEqual(launch("Ready"), (["code=0 slept=2"], True, 0))
            # 30초가 지나도 도는 중이면 지금처럼 보류한다
            self.assertEqual(launch("Running"), (["code=0 slept=" + "+".join(["2"] * 15)], False, 1))
            self.assertFalse((home / "publish-error.txt").exists())


if __name__ == "__main__":
    unittest.main()
