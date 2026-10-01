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
import shutil
import sys
import tempfile
import unittest
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
        self.variant("기록_0914.xlsx", {"I1": dt.datetime(2026, 9, 15), sv.DRAFT_CELL: None})
        self.variant("기록_v1.xlsx", {sv.SCHEMA_CELL: None})
        self.variant("기록_0921 - 복사본.xlsx", {"D3": "복사본에서 고친 결론"})
        (self.folder / "백업").mkdir()
        self.variant("백업/기록_0921.xlsx", {"D3": "백업 폴더의 결론"})
        (self.folder / "~$기록_0921.xlsx").write_bytes(b"lock")
        self.home, self.urlopen = pub.HOME, pub.URLOPEN
        pub.HOME = self.work / "home"
        os.environ.pop("MACRO_NOTES_TOKEN", None)

    def tearDown(self):
        pub.HOME, pub.URLOPEN = self.home, self.urlopen

    def variant(self, name, cells):
        wb = openpyxl.load_workbook(self.demo)
        for ref, value in cells.items():
            wb["Analysis"][ref] = value
        wb.save(self.folder / name)

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
        code, out = self.run_main("--out", str(self.work / "out"), github=FakeGitHub())
        self.assertEqual(code, 0, out)
        self.assertEqual(sorted(p.name for p in (self.work / "out").iterdir()), ["2026-09-14.json", "2026-09-21.json"])


if __name__ == "__main__":
    unittest.main()
