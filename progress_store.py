from __future__ import annotations

import json
from pathlib import Path

from storage_paths import data_file, FILE_LOCK, atomic_write_json

PROGRESS_FILE = data_file("training_progress.json")


class ProgressStore:
    def __init__(self, path: Path | None = None) -> None:
        self.path = path or PROGRESS_FILE

    def load(self) -> dict:
        if not self.path.exists():
            return {}
        with self.path.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)
        return payload if isinstance(payload, dict) else {}

    def get_section_progress(self, employee_id: str, section_id: str) -> dict | None:
        payload = self.load()
        return payload.get(employee_id, {}).get(section_id)

    def save_section_progress(self, employee_id: str, section_id: str, progress: dict) -> dict:
        with FILE_LOCK:
            payload = self.load()
            employee_progress = payload.setdefault(employee_id, {})
            employee_progress[section_id] = progress
            atomic_write_json(self.path, payload)
        return progress

    def clear_section_progress(self, employee_id: str, section_id: str) -> None:
        with FILE_LOCK:
            payload = self.load()
            if employee_id not in payload:
                return
            payload[employee_id].pop(section_id, None)
            if not payload[employee_id]:
                payload.pop(employee_id, None)
            atomic_write_json(self.path, payload)
