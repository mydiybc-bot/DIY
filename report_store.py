from __future__ import annotations

import json
from pathlib import Path

from storage_paths import data_file, FILE_LOCK, atomic_write_json

REPORT_FILE = data_file("training_reports.json")


class ReportStore:
    def __init__(self, path: Path | None = None) -> None:
        self.path = path or REPORT_FILE

    def load(self) -> list[dict]:
        if not self.path.exists():
            return []
        with self.path.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)
        return payload if isinstance(payload, list) else []

    def append(self, record: dict) -> dict:
        with FILE_LOCK:
            reports = self.load()
            reports.append(record)
            atomic_write_json(self.path, reports)
        return record
