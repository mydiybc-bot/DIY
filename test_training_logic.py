import unittest
from io import BytesIO

from openpyxl import load_workbook
from data_store import default_content
from excel_tools import export_content_to_excel, export_reports_to_excel, import_content_from_excel
from training_data import TRAINING_SECTIONS
from training_logic import build_progress_snapshot, build_question_text, build_report_record, create_session, respond, score_answer, summarize_session


class TrainingLogicTests(unittest.TestCase):
    def test_every_reference_answer_scores_ten(self):
        rules = default_content()["rules"]
        for section in TRAINING_SECTIONS:
            for question in section["questions"]:
                result = score_answer(question, question["answer"], rules)
                self.assertEqual(result["score"], 10, msg=f"{section['id']} Q{question['number']}")

    def test_prompt_format(self):
        session = create_session("customer_service", default_content())
        question = session["questions"][0]
        self.assertTrue(build_question_text(question, session["rules"]).startswith("Q1"))
        self.assertTrue(build_question_text(question, session["rules"]).endswith("請回答"))

    def test_reveal_reference_answer_after_three_failed_attempts(self):
        session = create_session("sales", default_content())
        message = ""
        for _ in range(3):
            result = respond(session, "不知道")
            message = result["message"]
        self.assertIn("參考答案：", message)
        self.assertIn("請依照這個方向，再回答一次。", message)

    def test_move_to_next_question_after_ten_points(self):
        session = create_session("new_staff", default_content())
        first = session["questions"][0]
        result = respond(session, first["answer"])
        self.assertIn("這題通過。", result["message"])
        self.assertIn("Q2：", result["message"])

    def test_summary_contains_scores(self):
        session = create_session("customer_service", default_content())
        respond(session, session["questions"][0]["answer"])
        summary = summarize_session(session)
        self.assertIn("本次練習題目總數：1", summary)
        self.assertIn("Q1：10 分", summary)

    def test_custom_end_phrase_and_reveal_limit(self):
        content = default_content()
        content["rules"]["end_phrase"] = "停止練習"
        content["rules"]["max_attempts_before_answer"] = 2
        session = create_session("customer_service", content)
        first_attempt = respond(session, "不知道")
        second_attempt = respond(session, "還是不知道")
        stopped = respond(session, "停止練習")
        self.assertNotIn("參考答案：", first_attempt["message"])
        self.assertIn("參考答案：", second_attempt["message"])
        self.assertTrue(stopped["done"])

    def test_excel_round_trip(self):
        content = default_content()
        content["rules"]["end_phrase"] = "Excel結束"
        raw = export_content_to_excel(content)
        imported = import_content_from_excel(raw, type("StoreStub", (), {"validate": staticmethod(lambda payload: payload)})())
        self.assertEqual(imported["rules"]["end_phrase"], "Excel結束")
        self.assertEqual(imported["sections"][0]["questions"][0]["prompt"], content["sections"][0]["questions"][0]["prompt"])

    def test_report_record_and_export(self):
        session = create_session("customer_service", default_content())
        respond(session, session["questions"][0]["answer"])
        record = build_report_record(session)
        self.assertEqual(record["average_score"], 10.0)
        raw = export_reports_to_excel([{"created_at": "2026-03-22T10:00:00", "role": "employee", "employee_id": "E001", "employee_name": "王小明", **record}])
        self.assertGreater(len(raw), 1000)
        workbook = load_workbook(BytesIO(raw))
        sheet = workbook["reports"]
        headers = [cell.value for cell in sheet[1]]
        self.assertEqual(headers, ["練習時間", "員工編號", "員工姓名", "單元名稱", "題目數", "各題分數", "平均分數"])
        first_row = [cell.value for cell in sheet[2]]
        self.assertEqual(first_row[0], "2026-03-22T10:00:00")
        self.assertEqual(first_row[1], "E001")
        self.assertEqual(first_row[2], "王小明")

    def test_progress_snapshot_after_ten_points(self):
        session = create_session("customer_service", default_content())
        respond(session, session["questions"][0]["answer"])
        snapshot = build_progress_snapshot(session)
        self.assertEqual(snapshot["current_index"], 1)
        self.assertEqual(snapshot["scoreboard"][0]["best_score"], 10)

    def test_keyword_outline_does_not_score_ten(self):
        rules = default_content()["rules"]
        question = TRAINING_SECTIONS[0]["questions"][2]
        answer = "說明食譜配方是設定好的、先確認客人是否想調整配方、需要協助可再詢問、加購食材需到櫃檯結帳。請再回答一次。"
        result = score_answer(question, answer, rules)
        self.assertLess(result["score"], 10)
        self.assertTrue(result["needs_better_expression"])
        self.assertIn("完整句子", result["coaching"])

    def test_all_keywords_without_customer_facing_sentence_is_capped(self):
        rules = default_content()["rules"]
        question = TRAINING_SECTIONS[0]["questions"][2]
        answer = "食譜配方設定好，調整配方，需要協助再詢問，加購櫃檯結帳。"
        result = score_answer(question, answer, rules)
        self.assertLess(result["score"], 10)
        self.assertTrue(result["needs_better_expression"])

    def test_profanity_caps_score_even_if_content_is_correct(self):
        rules = default_content()["rules"]
        question = TRAINING_SECTIONS[0]["questions"][0]
        answer = f"{question['answer']} 媽的"
        result = score_answer(question, answer, rules)
        self.assertLess(result["score"], 10)
        self.assertTrue(result["uses_profanity"])
        self.assertIn("不能帶髒話", result["coaching"])

    def test_transcribed_swearing_is_also_blocked(self):
        rules = default_content()["rules"]
        question = TRAINING_SECTIONS[0]["questions"][0]
        answer = "好的那很抱歉讓你白跑一趟了 12歲以上才能進來喔 因為店內的烤箱啦刀子啦都很危險啊 那這邊有個小卡送你喔 不過幹你不來就算了啦"
        result = score_answer(question, answer, rules)
        self.assertLess(result["score"], 10)
        self.assertTrue(result["uses_profanity"])
        self.assertIn("不能帶髒話", result["coaching"])


if __name__ == "__main__":
    unittest.main()
