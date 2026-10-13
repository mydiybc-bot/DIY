"""2026-10-13:AI 只判斷抓到哪些重點、分數由系統公式計算;成績報表每題明細。"""
import unittest
from unittest import mock

import training_logic
from data_store import default_content
from training_logic import build_report_record, compute_points_score, create_session, respond, score_answer


class ScoringFormulaTests(unittest.TestCase):
    def test_formula(self):
        self.assertEqual(compute_points_score(4, 4, False), 10)
        self.assertEqual(compute_points_score(4, 4, True), 9)
        self.assertEqual(compute_points_score(3, 4, False), 8)  # 7.5 四捨五入
        self.assertEqual(compute_points_score(2, 3, False), 7)
        self.assertEqual(compute_points_score(1, 2, False), 5)
        self.assertEqual(compute_points_score(0, 4, False), 0)

    def _question(self):
        return default_content()["sections"][0]["questions"][0]

    def test_ai_points_decide_score_not_ai_score(self):
        q = self._question()
        rules = default_content()["rules"]
        fake = {"matched_points": [1, 2, 3, 4], "needs_better_expression": False, "uses_profanity": False, "coaching": "很好", "score": 7}
        with mock.patch.object(training_logic, "_call_openai", return_value=fake):
            r = score_answer(q, "x", rules)
        self.assertEqual(r["score"], 10)
        self.assertEqual(r["_scored_by"], "ai")
        fake = {"matched_points": [1, 3, 9], "needs_better_expression": False, "uses_profanity": False, "coaching": "再補"}
        with mock.patch.object(training_logic, "_call_openai", return_value=fake):
            r = score_answer(q, "x", rules)
        self.assertEqual(r["matched_count"], 2)
        self.assertEqual(r["score"], 5)

    def test_old_label_format_still_works(self):
        q = self._question()
        labels = [p["label"] for p in q["required_points"]]
        fake = {"matched_labels": labels, "needs_better_expression": True, "uses_profanity": False, "coaching": "c"}
        with mock.patch.object(training_logic, "_call_openai", return_value=fake):
            r = score_answer(q, "x", default_content()["rules"])
        self.assertEqual(r["score"], 9)

    def test_bad_ai_reply_falls_back_to_rules(self):
        q = self._question()
        with mock.patch.object(training_logic, "_call_openai", return_value={"coaching": "x"}):
            r = score_answer(q, q["answer"], default_content()["rules"])
        self.assertEqual(r["_scored_by"], "rules_fallback")

    def test_report_details(self):
        content = default_content()
        session = create_session(content["sections"][0]["id"], content)
        n = content["rules"]["max_attempts_before_answer"]
        seq = [5] * n + [10]
        results = iter(seq + [10] * 20)

        def fake_score(question, answer, rules):
            s = next(results)
            k = len(question["required_points"])
            return {"score": s, "matched": [], "matched_count": k if s == 10 else 1, "missing": [], "missing_count": 0 if s == 10 else k - 1,
                    "needs_better_expression": False, "uses_profanity": False, "feedback": "f", "coaching": "c", "score_protected": False}

        with mock.patch.object(training_logic, "score_answer", side_effect=fake_score):
            for _ in range(n + 1):
                respond(session, "答案")  # Q1:答 n 次沒過、看過答案,第 n+1 次過
            respond(session, "答案")  # Q2:一次過
            respond(session, "跳過這題")  # Q3:跳過
        rec = build_report_record(session)
        d = {x["number"]: x for x in rec["details"]}
        self.assertEqual(d[1]["first_score"], 5)
        self.assertEqual(d[1]["attempts"], n + 1)
        self.assertTrue(d[1]["revealed"])
        self.assertEqual(d[2]["first_score"], 10)
        self.assertTrue(d[3]["skipped"])
        self.assertIsNone(d[3]["first_score"])
        self.assertEqual(rec["first_try_pass_rate"], 33)
        self.assertEqual(rec["revealed_count"], 1)
        self.assertEqual(rec["skipped_count"], 1)
        self.assertEqual(rec["avg_attempts"], round((n + 2) / 2, 1))


if __name__ == "__main__":
    unittest.main()
