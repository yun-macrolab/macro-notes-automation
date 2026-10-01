"""web/public_record.py·build.py 공개 기록 검사 — 시연 엔진으로 만든 합성 기록 워크북으로 확인한다.

  python -X utf8 -m unittest discover -s web -p "test_*.py"
"""
import copy
import json
from pathlib import Path
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "engine"), str(ROOT / "web")]

from core import DemoStore  # noqa: E402
import build  # noqa: E402
import public_record as pr  # noqa: E402
import workbook_view  # noqa: E402


class PublicRecord(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory(prefix="public_record_")
        store = DemoStore(cls.tmp.name)
        run_id = store.run()["run"]["id"]
        store.review(run_id, 1.5)
        store.reapply(run_id)
        cls.data = workbook_view.read(str(Path(cls.tmp.name) / run_id / "demo.xlsx"))

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def record(self, **changes):
        data = copy.deepcopy(self.data)
        for key, value in changes.items():
            data[key] = value
        return pr.build(data, generated_at="2026-10-01T00:00:00+00:00")

    def test_only_listed_fields(self):
        record = self.record()
        self.assertEqual(set(record), set(pr.TOP_KEYS))
        self.assertEqual(record["week"], {"start": "2026-09-21", "end": "2026-09-27"})
        self.assertTrue(all(set(n) == set(pr.NEWS_KEYS) for n in record["news"]))      # sheet_row 없음
        self.assertTrue(all(set(r) == set(pr.ROW_KEYS) for r in record["state"]["rows"]))  # 숨김 장부(draft_*) 없음
        self.assertEqual(set(record["state"]["dur"]), set(pr.BLOCK_KEYS))
        self.assertEqual(record["state"]["rows"][0]["locked"], ["dur"])
        self.assertEqual(record["notice"], pr.NOTICE)
        self.assertTrue(pr.validate(record))
        json.dumps(record, allow_nan=False)

    def test_keeps_only_news_of_the_week(self):
        news = copy.deepcopy(self.data["news"])
        outside = dict(news[0], date="2026-09-14", title="지난주 뉴스")
        undated = dict(news[0], date="날짜 미상", title="날짜 없는 뉴스")
        record = self.record(news=news + [outside, undated])
        self.assertEqual(len(record["news"]), len(news))
        self.assertTrue(all("2026-09-21" <= n["date"] <= "2026-09-27" for n in record["news"]))

    def test_needs_week(self):
        with self.assertRaises(ValueError):
            self.record(week={"date": None, "start": None, "end": None})

    def test_findings(self):
        memo = [{"label": "연락", "value": "ｈｏｎｇ＠ｅｘａｍｐｌｅ．ｃｏｍ 010-1234-5678"},
                {"label": "파일", "value": r"C:\Users\hong\기록.xlsx"},
                {"label": "출처", "value": "https://www.example.com/home/news/1/ 홍 길동"}]
        record = self.record(memo=memo)
        kinds = {(where, kind) for where, kind, _ in pr.findings(record, ["홍길동"])}
        self.assertEqual(kinds, {("memo[0].value", "이메일 주소"), ("memo[0].value", "전화번호"),
                                 ("memo[1].value", "PC 경로"), ("memo[2].value", "검사어")})  # 인터넷 주소는 걸리지 않는다
        self.assertEqual(pr.findings(self.record()), [])

    def test_public_json(self):
        ok, text = pr.public_json(json.dumps(self.data, ensure_ascii=False))
        self.assertTrue(ok)
        record = json.loads(text)
        self.assertEqual(record["schema"], pr.SCHEMA)
        self.assertTrue(pr.validate(record))
        word = self.data["news"][0]["title"].split()[-1]
        ok, text = pr.public_json(json.dumps(self.data, ensure_ascii=False), json.dumps([word], ensure_ascii=False))
        self.assertFalse(ok)
        body = json.loads(text)
        self.assertTrue(body["findings"])
        self.assertTrue(all(f["kind"] == "검사어" and f["text"] == word for f in body["findings"]))
        self.assertIn("News DB", body["findings"][0]["place"])
        no_week = dict(self.data, week={"date": None, "start": None, "end": None})
        ok, text = pr.public_json(json.dumps(no_week, ensure_ascii=False))
        self.assertFalse(ok)
        self.assertIn("I1", json.loads(text)["error"])

    def test_validate_rejects(self):
        good = self.record()
        cases = {
            "extra top field": lambda r: r.update(author="홍길동"),
            "extra news field": lambda r: r["news"][0].update(sheet_row=6),
            "hidden ledger": lambda r: r["state"]["rows"][0].update(draft_dur=1.0),
            "missing field": lambda r: r["state"]["flags"].pop("regime"),
            "schema": lambda r: r.update(schema="other/1"),
            "notice": lambda r: r.update(notice="안내 없음"),
            "week not monday": lambda r: r["week"].update(start="2026-09-22", end="2026-09-28"),
            "news outside week": lambda r: r["news"][0].update(date="2026-09-14"),
            "nested value": lambda r: r["news"][0].update(title=["a"]),
            "email": lambda r: r.update(conclusion="문의 hong@example.com"),
        }
        for name, change in cases.items():
            with self.subTest(name):
                record = copy.deepcopy(good)
                change(record)
                with self.assertRaises(ValueError):
                    pr.validate(record)

    def test_validate_does_not_leak_words(self):
        record = self.record(conclusion="홍길동의 결론")
        with self.assertRaises(ValueError) as caught:
            pr.validate(record, ["홍길동"])
        self.assertNotIn("홍길동", str(caught.exception))
        self.assertIn("conclusion(검사어)", str(caught.exception))


class SiteRecords(unittest.TestCase):
    """build.py가 records/의 공개본을 검사해 싣고 목록을 만드는지."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory(prefix="site_records_")
        store = DemoStore(cls.tmp.name)
        run_id = store.run()["run"]["id"]
        data = workbook_view.read(str(Path(cls.tmp.name) / run_id / "demo.xlsx"))
        cls.record = pr.build(data, generated_at="2026-10-01T00:00:00+00:00")

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def folder(self, files):
        source = Path(tempfile.mkdtemp(dir=self.tmp.name))
        for name, content in files.items():
            (source / name).write_text(content if isinstance(content, str) else json.dumps(content, ensure_ascii=False),
                                       encoding="utf-8")
        return source

    def test_publishes_valid_records(self):
        older = copy.deepcopy(self.record)
        older["week"] = {"start": "2026-09-14", "end": "2026-09-20"}
        older["news"] = []
        source = self.folder({"README.md": "# records", "2026-09-21.json": self.record, "2026-09-14.json": older})
        target = Path(self.tmp.name) / "out" / "records"
        entries = build.write_records(target, build.load_records(source))
        self.assertEqual([e["start"] for e in entries], ["2026-09-21", "2026-09-14"])   # 최신 주가 먼저
        index = json.loads((target / "index.json").read_text(encoding="utf-8"))
        self.assertEqual(index["records"][0]["news"], len(self.record["news"]))
        self.assertEqual(json.loads((target / "2026-09-21.json").read_text(encoding="utf-8")), self.record)
        self.assertEqual(build.write_records(Path(self.tmp.name) / "empty", []), [])

    def test_rejects_bad_folders(self):
        cases = {
            "wrong name": {"2026-09-21 (1).json": self.record},
            "name differs from week": {"2026-09-28.json": self.record},
            "not json": {"2026-09-21.json": "{broken"},
            "workbook uploaded": {"기록.xlsx": "PK"},
            "check word": {"2026-09-21.json": dict(self.record, conclusion="홍길동")},
        }
        for name, files in cases.items():
            with self.subTest(name):
                with self.assertRaises(ValueError) as caught:
                    build.load_records(self.folder(files), ["홍길동"])
                self.assertNotIn("홍길동", str(caught.exception))


if __name__ == "__main__":
    unittest.main()
