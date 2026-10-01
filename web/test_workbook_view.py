"""web/workbook_view.py 검사 — 시연 엔진으로 만든 합성 기록 워크북을 읽어 본다(네트워크·실제 기록 불필요).

  python -X utf8 -m unittest discover -s web -p "test_*.py"
"""
import hashlib
import json
from pathlib import Path
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parent.parent
sys.path[:0] = [str(ROOT), str(ROOT / "engine"), str(ROOT / "web")]

import openpyxl  # noqa: E402

from core import DemoStore  # noqa: E402
import sheet_v2 as sv  # noqa: E402
import workbook_view  # noqa: E402


class ReadRecord(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = tempfile.TemporaryDirectory(prefix="workbook_view_")
        store = DemoStore(cls.tmp.name)
        run_id = store.run()["run"]["id"]
        store.review(run_id, 1.5)
        store.reapply(run_id)
        cls.path = Path(cls.tmp.name) / run_id / "demo.xlsx"

    @classmethod
    def tearDownClass(cls):
        cls.tmp.cleanup()

    def test_reads_news_scores_and_memo(self):
        data = workbook_view.read(str(self.path))
        self.assertEqual(len(data["news"]), 8)                      # '예시' 서식 행은 빠진다
        self.assertTrue(all(n["title"] and n["factor"] for n in data["news"]))
        dates = [n["date"] for n in data["news"]]
        self.assertEqual(dates, sorted(dates, reverse=True))        # 최신이 위로
        self.assertEqual(data["week"], {"date": "2026-09-21", "start": "2026-09-21", "end": "2026-09-27"})
        self.assertTrue(data["conclusion"].startswith(sv.CONC_HEAD))
        self.assertEqual(len(data["state"]["rows"]), 7)
        self.assertEqual(len(data["memo"]), 7)                       # 라벨만 있고 값은 빈 메모 칸
        first = data["state"]["rows"][0]
        self.assertEqual((first["dur"], "dur" in first["locked"]), (1.5, True))   # 사람이 고친 칸
        self.assertIsInstance(first["digest"], str)
        json.dumps(data, ensure_ascii=False)

    def test_scores_match_engine(self):
        data = workbook_view.read(str(self.path))
        state = sv.read_state(str(self.path))
        self.assertEqual(data["state"]["dur"]["total"], state["dur"]["total"])
        self.assertEqual(data["state"]["cur"]["verdict"], state["cur"]["verdict"])

    def test_does_not_modify_file(self):
        before = hashlib.sha256(self.path.read_bytes()).hexdigest()
        workbook_view.read_json(str(self.path))
        self.assertEqual(hashlib.sha256(self.path.read_bytes()).hexdigest(), before)

    def test_rejects_other_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            v1 = Path(tmp) / "v1.xlsx"
            wb = openpyxl.load_workbook(self.path)
            wb["Analysis"][sv.SCHEMA_CELL] = None
            wb.save(v1)
            no_news = Path(tmp) / "no_news.xlsx"
            wb = openpyxl.Workbook()
            wb.active.title = "Analysis"
            wb["Analysis"][sv.SCHEMA_CELL] = sv.SCHEMA_V2
            wb.save(no_news)
            broken = Path(tmp) / "broken.xlsx"
            broken.write_bytes(b"not a workbook")
            for path, phrase in ((v1, "v2 형식"), (no_news, "'News DB'"), (broken, "읽지 못했습니다")):
                ok, text = workbook_view.read_json(str(path))
                self.assertFalse(ok)
                self.assertIn(phrase, json.loads(text)["error"])


if __name__ == "__main__":
    unittest.main()
