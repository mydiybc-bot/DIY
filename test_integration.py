from __future__ import annotations

import json
from pathlib import Path
import unittest
from http.cookiejar import CookieJar
from urllib.request import HTTPCookieProcessor, Request, build_opener


BASE_URL = "http://127.0.0.1:8765"
PROGRESS_FILE = Path("/Users/diybc/Desktop/Chatgpt Codex/training_progress.json")


def make_opener():
    return build_opener(HTTPCookieProcessor(CookieJar()))


def post_json(opener, path: str, payload: dict) -> tuple[int, dict]:
    req = Request(
        BASE_URL + path,
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
    )
    with opener.open(req) as response:
        return response.status, json.loads(response.read().decode("utf-8"))


def get_json(opener, path: str) -> tuple[int, dict]:
    with opener.open(BASE_URL + path) as response:
        return response.status, json.loads(response.read().decode("utf-8"))


class IntegrationTests(unittest.TestCase):
    def setUp(self):
        if PROGRESS_FILE.exists():
            PROGRESS_FILE.unlink()

    def test_admin_and_employee_core_flows(self):
        admin = make_opener()
        status, login = post_json(admin, "/api/login", {"role": "admin", "password": "admin123"})
        self.assertEqual(status, 200)
        self.assertEqual(login["role"], "admin")

        status, auth = get_json(admin, "/api/admin/auth")
        self.assertEqual(status, 200)
        self.assertTrue(auth["auth"]["employees"])

        first_employee = auth["auth"]["employees"][0]
        employee_id = first_employee["employee_id"]
        employee_name = first_employee["employee_name"]
        employee_password = first_employee["password"]

        status, content = get_json(admin, "/api/admin/content")
        self.assertEqual(status, 200)
        self.assertIn("sections", content["content"])

        status, reports_before = get_json(admin, f"/api/admin/reports?employee_id={employee_id}")
        initial_count = len(reports_before["reports"])

        employee = make_opener()
        status, employee_login = post_json(
            employee,
            "/api/login",
            {"role": "employee", "employee_id": employee_id, "password": employee_password},
        )
        self.assertEqual(status, 200)
        self.assertEqual(employee_login["employee_name"], employee_name)

        first_section = content["content"]["sections"][0]
        status, started = post_json(employee, "/api/start", {"section_id": first_section["id"]})
        self.assertEqual(status, 200)
        self.assertIn("Q1", started["message"])

        answer = first_section["questions"][0]["answer"]
        status, first_response = post_json(employee, "/api/respond", {"message": answer})
        self.assertEqual(status, 200)
        self.assertIn("評分：10/10", first_response["message"])

        end_phrase = content["content"]["rules"]["end_phrase"]
        status, finished = post_json(employee, "/api/respond", {"message": end_phrase})
        self.assertEqual(status, 200)
        self.assertTrue(finished["done"])

        status, reports_after = get_json(
            admin,
            f"/api/admin/reports?employee_id={employee_id}&section_id={first_section['id']}",
        )
        self.assertGreaterEqual(len(reports_after["reports"]), initial_count + 1)
        latest = reports_after["reports"][-1]
        self.assertEqual(latest["employee_id"], employee_id)
        self.assertEqual(latest["employee_name"], employee_name)
        self.assertEqual(latest["section_id"], first_section["id"])

    def test_employee_resumes_from_next_question_after_ten_points(self):
        admin = make_opener()
        status, login = post_json(admin, "/api/login", {"role": "admin", "password": "admin123"})
        self.assertEqual(status, 200)

        _, auth = get_json(admin, "/api/admin/auth")
        first_employee = auth["auth"]["employees"][0]
        employee_id = first_employee["employee_id"]
        employee_password = first_employee["password"]

        _, content = get_json(admin, "/api/admin/content")
        first_section = content["content"]["sections"][0]

        employee = make_opener()
        status, employee_login = post_json(
            employee,
            "/api/login",
            {"role": "employee", "employee_id": employee_id, "password": employee_password},
        )
        self.assertEqual(status, 200)

        status, started = post_json(employee, "/api/start", {"section_id": first_section["id"]})
        self.assertEqual(status, 200)

        answer = first_section["questions"][0]["answer"]
        status, first_response = post_json(employee, "/api/respond", {"message": answer})
        self.assertEqual(status, 200)
        self.assertIn("Q2", first_response["message"])

        employee_relogin = make_opener()
        status, employee_login_again = post_json(
            employee_relogin,
            "/api/login",
            {"role": "employee", "employee_id": employee_id, "password": employee_password},
        )
        self.assertEqual(status, 200)

        status, resumed = post_json(employee_relogin, "/api/start", {"section_id": first_section["id"]})
        self.assertEqual(status, 200)
        self.assertIn("下一題", resumed["message"])
        self.assertIn("Q2", resumed["message"])

    def test_public_config_responds(self):
        opener = make_opener()
        status, config = get_json(opener, "/api/config")
        self.assertEqual(status, 200)
        self.assertTrue(config["sections"])


if __name__ == "__main__":
    unittest.main()
