from __future__ import annotations

import json
from pathlib import Path

from storage_paths import data_file
from training_data import ADMIN_PASSWORD, EMPLOYEE_PASSWORD

AUTH_FILE = data_file("auth_config.json")


def default_auth() -> dict:
    return {
        "admin_password": ADMIN_PASSWORD,
        "employees": [
            {
                "employee_id": "E001",
                "employee_name": "示範員工",
                "password": EMPLOYEE_PASSWORD,
            }
        ],
    }


class AuthStore:
    def __init__(self, path: Path | None = None) -> None:
        self.path = path or AUTH_FILE

    def load(self) -> dict:
        if not self.path.exists():
            return default_auth()
        with self.path.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)
        validated = self.validate(payload)
        if validated != payload:
            with self.path.open("w", encoding="utf-8") as handle:
                json.dump(validated, handle, ensure_ascii=False, indent=2)
        return validated

    def save(self, payload: dict) -> dict:
        validated = self.validate(payload)
        with self.path.open("w", encoding="utf-8") as handle:
            json.dump(validated, handle, ensure_ascii=False, indent=2)
        return validated

    def validate(self, payload: dict) -> dict:
        if not isinstance(payload, dict):
            raise ValueError("密碼設定格式錯誤")
        admin_password = str(payload.get("admin_password", "")).strip()
        employees = payload.get("employees", [])
        legacy_employee_password = str(payload.get("employee_password", "")).strip()
        if (not isinstance(employees, list) or not employees) and legacy_employee_password:
            employees = [
                {
                    "employee_id": "E001",
                    "employee_name": "示範員工",
                    "password": legacy_employee_password,
                }
            ]
        if len(admin_password) < 4:
            raise ValueError("管理員密碼至少需要 4 個字元")
        if not isinstance(employees, list) or not employees:
            raise ValueError("至少需要一位員工帳號")

        cleaned_employees = []
        seen_ids = set()
        for item in employees:
            if not isinstance(item, dict):
                raise ValueError("員工資料格式錯誤")
            employee_id = str(item.get("employee_id", "")).strip()
            employee_name = str(item.get("employee_name", "")).strip()
            password = str(item.get("password", "")).strip()
            if not employee_id or not employee_name:
                raise ValueError("員工編號與姓名不能空白")
            if len(password) < 4:
                raise ValueError("員工密碼至少需要 4 個字元")
            if employee_id in seen_ids:
                raise ValueError("員工編號不可重複")
            seen_ids.add(employee_id)
            cleaned_employees.append(
                {
                    "employee_id": employee_id,
                    "employee_name": employee_name,
                    "password": password,
                }
            )
        return {
            "admin_password": admin_password,
            "employees": cleaned_employees,
        }
