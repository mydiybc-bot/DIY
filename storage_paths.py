from __future__ import annotations

import os
from pathlib import Path


DEFAULT_DATA_DIR = Path(__file__).parent


def data_dir() -> Path:
    configured = os.environ.get("TRAINING_DATA_DIR", "").strip()
    if not configured:
        return DEFAULT_DATA_DIR
    path = Path(configured).expanduser()
    path.mkdir(parents=True, exist_ok=True)
    return path


def data_file(filename: str) -> Path:
    return data_dir() / filename
