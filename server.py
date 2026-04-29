from __future__ import annotations

from email.parser import BytesParser
from email.policy import default as email_policy
import json
import os
import secrets
from datetime import datetime
from http import HTTPStatus
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import socket
from urllib.parse import parse_qs, urlparse

from auth_store import AuthStore
from data_store import TrainingContentStore
from excel_tools import export_content_to_excel, export_reports_to_excel, import_content_from_excel
from google_reviews_dashboard import load_google_reviews_dashboard
from progress_store import ProgressStore
from reimbursement_dashboard import load_reimbursement_dashboard
from report_store import ReportStore
from self_dashboard import load_self_dashboard
from training_logic import build_progress_snapshot, build_question_text, build_report_record, create_session, respond
from voice_transcription import TranscriptionError, transcribe_audio

ROOT = Path(__file__).parent
STATIC_DIR = ROOT / "static"
HOST = "0.0.0.0"
PORT = int(os.environ.get("PORT", "8765"))
SESSIONS: dict[str, dict] = {}
STORE = TrainingContentStore()
AUTH_STORE = AuthStore()
REPORT_STORE = ReportStore()
PROGRESS_STORE = ProgressStore()


class TrainingHandler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.serve_file("index.html", "text/html; charset=utf-8")
            return
        if parsed.path == "/dashboard":
            self.serve_file("dashboard.html", "text/html; charset=utf-8")
            return
        if parsed.path == "/pnl":
            self.serve_file("dashboard-pnl.html", "text/html; charset=utf-8")
            return
            
        if parsed.path == "/self-dashboard":
            self.serve_file("self-dashboard.html", "text/html; charset=utf-8")
            return
        if parsed.path == "/google-reviews":
            self.serve_file("google-reviews.html", "text/html; charset=utf-8")
            return
        if parsed.path == "/api/config":
            content = STORE.load()
            session = self._current_session()
            self.send_json(
                {
                    "sections": [
                        {"id": section["id"], "title": section["title"]}
                        for section in content["sections"]
                    ],
                    "rules": {
                        "end_phrase": content["rules"]["end_phrase"],
                    },
                    "role": session["role"] if session else None,
                    "employee_name": session.get("employee_name") if session else None,
                    "employee_id": session.get("employee_id") if session else None,
                }
            )
            return
        if parsed.path == "/api/reimbursements":
            self.send_json(load_reimbursement_dashboard())
            return
        if parsed.path == "/api/self-dashboard":
            filters = parse_qs(parsed.query)
            self.send_json(
                load_self_dashboard(
                    {
                        "start_date": _first_query_value(filters, "start_date"),
                        "end_date": _first_query_value(filters, "end_date"),
                        "store": _first_query_value(filters, "store"),
                        "coupon": _first_query_value(filters, "coupon"),
                    }
                )
            )
            return
        if parsed.path == "/api/google-reviews":
            self.send_json(load_google_reviews_dashboard())
            return
        if parsed.path == "/api/admin/content":
            session = self._require_auth(role="admin")
            if not session:
                return
            self.send_json({"ok": True, "content": STORE.load()})
            return
        if parsed.path == "/api/admin/auth":
            session = self._require_auth(role="admin")
            if not session:
                return
            self.send_json({"ok": True, "auth": AUTH_STORE.load()})
            return
        if parsed.path == "/api/admin/reports":
            session = self._require_auth(role="admin")
            if not session:
                return
            filters = parse_qs(parsed.query)
            reports = self._filter_reports(REPORT_STORE.load(), filters)
            self.send_json({"ok": True, "reports": reports})
            return
        if parsed.path == "/api/admin/export.xlsx":
            session = self._require_auth(role="admin")
            if not session:
                return
            raw = export_content_to_excel(STORE.load())
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="training_content.xlsx"')
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
            return
        if parsed.path == "/api/admin/reports.xlsx":
            session = self._require_auth(role="admin")
            if not session:
                return
            filters = parse_qs(parsed.query)
            raw = export_reports_to_excel(self._filter_reports(REPORT_STORE.load(), filters))
            self.send_response(HTTPStatus.OK)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="training_reports.xlsx"')
            self.send_header("Content-Length", str(len(raw)))
            self.end_headers()
            self.wfile.write(raw)
            return

        file_path = parsed.path.lstrip("/")
        if file_path.startswith("static/"):
            local_name = file_path.replace("static/", "", 1)
            self.serve_file(local_name, self._guess_content_type(local_name))
            return

        self.send_error(HTTPStatus.NOT_FOUND, "Not found")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/login":
            body = self.read_json()
            role = body.get("role", "employee")
            auth = AUTH_STORE.load()
            employee_id = None
            employee_name = None

            if role == "admin":
                password = body.get("password", "")
                if password != auth["admin_password"]:
                    self.send_json({"ok": False, "error": "密碼錯誤，請再試一次。"}, status=401)
                    return
            else:
                employee_id = str(body.get("employee_id", "")).strip()
                password = str(body.get("password", "")).strip()
                employee = next((item for item in auth["employees"] if item["employee_id"] == employee_id), None)
                if not employee or employee["password"] != password:
                    self.send_json({"ok": False, "error": "員工編號或密碼錯誤，請再試一次。"}, status=401)
                    return
                employee_name = employee["employee_name"]

            session_id = secrets.token_hex(16)
            SESSIONS[session_id] = {
                "authenticated": True,
                "role": role,
                "training": None,
                "employee_id": employee_id,
                "employee_name": employee_name,
            }
            self.send_json(
                {"ok": True, "role": role, "employee_id": employee_id, "employee_name": employee_name},
                cookies={"training_session": session_id},
            )
            return

        if parsed.path == "/api/logout":
            session_id = self._session_id()
            if session_id:
                SESSIONS.pop(session_id, None)
            self.send_json({"ok": True}, cookies={"training_session": ""}, expire_cookies=["training_session"])
            return

        if parsed.path == "/api/admin/import":
            session = self._require_auth(role="admin")
            if not session:
                return
            raw = self.read_body_bytes()
            try:
                validated = import_content_from_excel(raw, STORE)
                saved = STORE.save_validated(validated)
            except ValueError as exc:
                self.send_json({"ok": False, "error": str(exc)}, status=400)
                return
            self.send_json({"ok": True, "content": saved})
            return

        session = self._require_auth()
        if not session:
            return

        if parsed.path == "/api/start":
            body = self.read_json()
            section_id = body.get("section_id", "")
            content = STORE.load()
            progress = None
            resume_note = ""
            if session.get("role") == "employee" and session.get("employee_id"):
                progress = PROGRESS_STORE.get_section_progress(session["employee_id"], section_id)

            if progress:
                resumed_index = int(progress.get("current_index", 0))
                section_length = len(next(section for section in content["sections"] if section["id"] == section_id)["questions"])
                if resumed_index >= section_length:
                    PROGRESS_STORE.clear_section_progress(session["employee_id"], section_id)
                    progress = None
                    resume_note = "你已完成過本單元，這次會重新從第一題開始。\n"
                elif resumed_index > 0:
                    resume_note = f"已為你接續到上一題拿到 10 分後的下一題。\n"

            training = create_session(section_id, content, progress)
            session["training"] = training
            question = build_question_text(training["questions"][training["current_index"]], training["rules"])
            self.send_json(
                {
                    "ok": True,
                    "message": resume_note + question,
                    "title": training["section_title"],
                }
            )
            return

        if parsed.path == "/api/respond":
            body = self.read_json()
            user_message = body.get("message", "")
            try:
                result = self._handle_training_reply(session, user_message)
            except ValueError as exc:
                self.send_json({"ok": False, "error": str(exc)}, status=400)
                return
            self.send_json({"ok": True, **result})
            return

        if parsed.path == "/api/respond-audio":
            try:
                form = self.read_multipart_form()
            except ValueError as exc:
                self.send_json({"ok": False, "error": str(exc)}, status=400)
                return
            audio_field = form.get("audio")
            if not audio_field or not audio_field.get("data"):
                self.send_json({"ok": False, "error": "沒有收到語音檔案，請再試一次。"}, status=400)
                return

            try:
                transcript = transcribe_audio(
                    audio_bytes=audio_field["data"],
                    filename=audio_field.get("filename") or "reply.webm",
                    content_type=audio_field.get("content_type") or "application/octet-stream",
                )
                result = self._handle_training_reply(session, transcript)
            except TranscriptionError as exc:
                self.send_json({"ok": False, "error": str(exc)}, status=502)
                return
            except ValueError as exc:
                self.send_json({"ok": False, "error": str(exc)}, status=400)
                return

            self.send_json({"ok": True, "transcript": transcript, **result})
            return

        if parsed.path == "/api/admin/content":
            if session.get("role") != "admin":
                self.send_json({"ok": False, "error": "只有管理員可以修改後台。"}, status=403)
                return
            body = self.read_json()
            try:
                saved = STORE.save(body)
            except ValueError as exc:
                self.send_json({"ok": False, "error": str(exc)}, status=400)
                return
            self.send_json({"ok": True, "content": saved})
            return

        if parsed.path == "/api/admin/auth":
            if session.get("role") != "admin":
                self.send_json({"ok": False, "error": "只有管理員可以修改密碼。"}, status=403)
                return
            body = self.read_json()
            try:
                saved = AUTH_STORE.save(body)
            except ValueError as exc:
                self.send_json({"ok": False, "error": str(exc)}, status=400)
                return
            self.send_json({"ok": True, "auth": saved})
            return

        self.send_error(HTTPStatus.NOT_FOUND, "Not found")

    def serve_file(self, filename: str, content_type: str) -> None:
        path = STATIC_DIR / filename
        if not path.exists():
            self.send_error(HTTPStatus.NOT_FOUND, "Not found")
            return
        data = path.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store, max-age=0")
        self.send_header("Pragma", "no-cache")
        self.end_headers()
        self.wfile.write(data)

    def read_json(self) -> dict:
        length = int(self.headers.get("Content-Length", "0"))
        raw = self.rfile.read(length).decode("utf-8") if length else "{}"
        return json.loads(raw or "{}")

    def read_body_bytes(self) -> bytes:
        length = int(self.headers.get("Content-Length", "0"))
        return self.rfile.read(length) if length else b""

    def read_multipart_form(self) -> dict[str, dict[str, str | bytes]]:
        content_type = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in content_type:
            raise ValueError("語音上傳格式不正確，請重新錄音一次。")
        raw = self.read_body_bytes()
        message = BytesParser(policy=email_policy).parsebytes(
            f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("utf-8") + raw
        )
        parts: dict[str, dict[str, str | bytes]] = {}
        for part in message.iter_parts():
            name = part.get_param("name", header="content-disposition")
            if not name:
                continue
            parts[name] = {
                "filename": part.get_filename() or "",
                "content_type": part.get_content_type(),
                "data": part.get_payload(decode=True) or b"",
            }
        return parts

    def send_json(
        self,
        payload: dict,
        status: int = 200,
        cookies: dict | None = None,
        expire_cookies: list[str] | None = None,
    ) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store, max-age=0")
        self.send_header("Pragma", "no-cache")
        if cookies:
            for name, value in cookies.items():
                cookie = SimpleCookie()
                cookie[name] = value
                cookie[name]["path"] = "/"
                cookie[name]["httponly"] = True
                self.send_header("Set-Cookie", cookie.output(header="").strip())
        if expire_cookies:
            for name in expire_cookies:
                cookie = SimpleCookie()
                cookie[name] = ""
                cookie[name]["path"] = "/"
                cookie[name]["expires"] = "Thu, 01 Jan 1970 00:00:00 GMT"
                self.send_header("Set-Cookie", cookie.output(header="").strip())
        self.end_headers()
        self.wfile.write(data)

    def _guess_content_type(self, filename: str) -> str:
        if filename.endswith(".html"):
            return "text/html; charset=utf-8"
        if filename.endswith(".css"):
            return "text/css; charset=utf-8"
        if filename.endswith(".js"):
            return "application/javascript; charset=utf-8"
        return "text/plain; charset=utf-8"

    def _session_id(self) -> str | None:
        cookie_header = self.headers.get("Cookie")
        if not cookie_header:
            return None
        cookies = SimpleCookie()
        cookies.load(cookie_header)
        morsel = cookies.get("training_session")
        return morsel.value if morsel else None

    def _current_session(self) -> dict | None:
        session_id = self._session_id()
        return SESSIONS.get(session_id or "")

    def _require_auth(self, role: str | None = None) -> dict | None:
        session = self._current_session()
        if not session or not session.get("authenticated"):
            self.send_json({"ok": False, "error": "請先登入。"}, status=401)
            return None
        if role and session.get("role") != role:
            self.send_json({"ok": False, "error": "你沒有這個操作權限。"}, status=403)
            return None
        return session

    def _filter_reports(self, reports: list[dict], filters: dict[str, list[str]]) -> list[dict]:
        employee_id = (filters.get("employee_id") or [""])[0].strip()
        section_id = (filters.get("section_id") or [""])[0].strip()
        date_from = (filters.get("date_from") or [""])[0].strip()
        date_to = (filters.get("date_to") or [""])[0].strip()

        filtered = reports
        if employee_id:
            filtered = [item for item in filtered if item.get("employee_id", "") == employee_id]
        if section_id:
            filtered = [item for item in filtered if item.get("section_id", "") == section_id]
        if date_from:
            filtered = [item for item in filtered if item.get("created_at", "")[:10] >= date_from]
        if date_to:
            filtered = [item for item in filtered if item.get("created_at", "")[:10] <= date_to]
        return filtered

    def _handle_training_reply(self, session: dict, user_message: str) -> dict:
        training = session.get("training")
        if not training:
            raise ValueError("請先選擇練習單元。")

        result = respond(training, user_message)
        if session.get("role") == "employee" and session.get("employee_id"):
            PROGRESS_STORE.save_section_progress(
                session["employee_id"],
                training["section_id"],
                build_progress_snapshot(training),
            )
        if result.get("done"):
            report = build_report_record(training)
            report["created_at"] = datetime.now().isoformat(timespec="seconds")
            report["role"] = session.get("role", "employee")
            report["employee_id"] = session.get("employee_id", "")
            report["employee_name"] = session.get("employee_name", "")
            REPORT_STORE.append(report)
            session["training"] = None
        return result

    def log_message(self, format: str, *args) -> None:
        return


def main() -> None:
    server = ThreadingHTTPServer((HOST, PORT), TrainingHandler)
    local_url = f"http://127.0.0.1:{PORT}"
    lan_ip = detect_lan_ip()
    lan_url = f"http://{lan_ip}:{PORT}" if lan_ip else None
    print(f"Training site running on {local_url}")
    if lan_url:
        print(f"LAN access URL: {lan_url}")
    server.serve_forever()


def detect_lan_ip() -> str | None:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as sock:
            sock.connect(("8.8.8.8", 80))
            ip = sock.getsockname()[0]
        return ip if ip and not ip.startswith("127.") else None
    except OSError:
        return None


def _first_query_value(filters: dict[str, list[str]], key: str) -> str:
    return (filters.get(key) or [""])[0].strip()


if __name__ == "__main__":
    main()
