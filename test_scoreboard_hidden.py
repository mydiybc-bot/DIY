"""首頁成績表隱藏已離職的人（2026-11-23）。"""
import json
import tempfile
import unittest
from pathlib import Path

from scoreboard import build_public_scoreboard, hidden_key, load_hidden_keys

CONTENT = {"sections": [{"id": "s1", "title": "單元一", "questions": [{}, {}]}]}
AUTH = {
    "employees": [
        {"employee_id": "A1", "employee_name": "王小明", "store": "3 台北南京店"},
        {"employee_id": "A2", "employee_name": "李小華", "store": "3 台北南京店"},
        {"employee_id": "A3", "employee_name": "王小明", "store": "8 新北新店店"},
    ]
}
REPORTS = [{"employee_id": "A1", "section_id": "s1", "question_count": 2, "created_at": "2026-07-01", "scores": ["Q1:10", "Q2:9"]}]


def names(board):
    return {(s["store"], p["name"]) for s in board["stores"] for p in s["people"]}


class HiddenTest(unittest.TestCase):
    def test_hidden_person_not_listed_and_counts_follow(self):
        board = build_public_scoreboard(CONTENT, AUTH, REPORTS, {}, {hidden_key("3 台北南京店", "王小明")})
        self.assertEqual(names(board), {("3 台北南京店", "李小華"), ("8 新北新店店", "王小明")})
        nanjing = next(s for s in board["stores"] if s["store"] == "3 台北南京店")
        self.assertEqual(nanjing["practiced"], 0)

    def test_same_name_other_store_unaffected(self):
        self.assertNotEqual(hidden_key("3 台北南京店", "王小明"), hidden_key("8 新北新店店", "王小明"))

    def test_key_uses_store_number_only(self):
        self.assertEqual(hidden_key("3 台北南京店", "王小明"), hidden_key("3 南京店", " 王小明 "))

    def test_bad_or_missing_file_hides_nobody(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(load_hidden_keys(Path(d) / "none.json"), set())
            bad = Path(d) / "bad.json"
            bad.write_text("{壞掉", encoding="utf-8")
            self.assertEqual(load_hidden_keys(bad), set())

    def test_repo_file_parses(self):
        keys = load_hidden_keys()
        payload = json.loads(Path("data/scoreboard_hidden.json").read_text(encoding="utf-8"))
        self.assertEqual(len(keys), len(payload["hidden"]))


if __name__ == "__main__":
    unittest.main()
