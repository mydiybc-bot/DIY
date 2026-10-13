"""2026-10-13:登入首頁各店人員成績表、員工帳號店別欄位。"""
import unittest

from auth_store import AuthStore
from scoreboard import UNASSIGNED, build_public_scoreboard

CONTENT = {"sections": [{"id": "a", "title": "單元A", "questions": [{}, {}, {}]}, {"id": "b", "title": "單元B", "questions": [{}, {}]}]}
AUTH = {"employees": [
    {"employee_id": "E1", "employee_name": "小明", "password": "1234", "store": "12 高雄SKM Park店"},
    {"employee_id": "E2", "employee_name": "小華", "password": "1234", "store": "2 台中草悟道店"},
    {"employee_id": "E3", "employee_name": "小美", "password": "1234", "store": "體驗設計部"},
    {"employee_id": "E4", "employee_name": "小強", "password": "1234"},
]}
REPORTS = [
    # 小明:單元A 先練到 Q2 中途結束,再接續完成(Q1、Q2 重複出現在兩筆,只算最新)
    {"employee_id": "E1", "role": "employee", "section_id": "a", "question_count": 2, "created_at": "2026-10-13T09:00:00",
     "details": [{"number": 1, "first_score": 10, "best_score": 10, "attempts": 1}, {"number": 2, "first_score": 6, "best_score": 10, "attempts": 3}]},
    {"employee_id": "E1", "role": "employee", "section_id": "a", "question_count": 3, "created_at": "2026-10-13T10:00:00",
     "details": [{"number": 1, "first_score": 10, "best_score": 10, "attempts": 1}, {"number": 2, "first_score": 6, "best_score": 10, "attempts": 3},
                 {"number": 3, "first_score": 10, "best_score": 10, "attempts": 1}]},
    # 小華:舊格式紀錄(只有 scores 字串)
    {"employee_id": "E2", "role": "employee", "section_id": "b", "question_count": 1, "created_at": "2026-10-01T10:00:00", "scores": ["Q1:10", "Q2:0(跳過)"]},
    # 管理員的紀錄不算
    {"employee_id": "", "role": "admin", "section_id": "a", "question_count": 3, "created_at": "2026-10-13T11:00:00", "scores": ["Q1:10"]},
]
PROGRESS = {"E2": {"a": {"current_index": 1, "scoreboard": []}}}


class ScoreboardTests(unittest.TestCase):
    def setUp(self):
        self.board = build_public_scoreboard(CONTENT, AUTH, REPORTS, PROGRESS)

    def test_store_order_and_unassigned_last(self):
        # 部門(主管、總部)不列;未設定店別仍列出提醒補填
        self.assertEqual([s["store"] for s in self.board["stores"]], ["2 台中草悟道店", "12 高雄SKM Park店", UNASSIGNED])

    def test_metrics_dedupe_latest(self):
        ming = self.board["stores"][1]["people"][0]
        self.assertEqual(ming["name"], "小明")
        self.assertTrue(ming["units"][0]["completed"])
        self.assertEqual(ming["completed_units"], 1)
        self.assertEqual(ming["first_try_pass_rate"], 67)  # 3 題中 2 題一次過
        self.assertEqual(ming["avg_attempts"], 1.7)       # (1+3+1)/3
        self.assertEqual(ming["average_score"], 10)
        self.assertEqual(ming["last_practice"], "2026-10-13")

    def test_old_reports_and_progress(self):
        hua = self.board["stores"][0]["people"][0]
        self.assertEqual(hua["units"][0], {"done": 1, "total": 3, "completed": False})
        self.assertIsNone(hua["first_try_pass_rate"])
        self.assertEqual(hua["average_score"], 5)
        self.assertEqual(self.board["stores"][0]["practiced"], 1)

    def test_no_private_fields(self):
        text = repr(self.board)
        self.assertNotIn("1234", text)
        self.assertNotIn("E1", text)

    def test_auth_store_keeps_store(self):
        validated = AuthStore.validate(None, {"admin_password": "admin123", "employees": AUTH["employees"]})
        self.assertEqual(validated["employees"][0]["store"], "12 高雄SKM Park店")
        self.assertEqual(validated["employees"][3]["store"], "")


if __name__ == "__main__":
    unittest.main()
