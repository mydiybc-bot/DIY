from __future__ import annotations

import json
import os
import tempfile
import threading
from pathlib import Path


DEFAULT_DATA_DIR = Path(__file__).parent

# 全域檔案鎖：本服務為單一程序（ThreadingHTTPServer），用一把可重入鎖把所有
# 「讀檔→改→寫檔」序列化，避免多人同時操作造成資料互相覆蓋或遺失。
FILE_LOCK = threading.RLock()


def data_dir() -> Path:
    configured = os.environ.get("TRAINING_DATA_DIR", "").strip()
    if not configured:
        return DEFAULT_DATA_DIR
    path = Path(configured).expanduser()
    path.mkdir(parents=True, exist_ok=True)
    return path


def data_file(filename: str) -> Path:
    return data_dir() / filename


def atomic_write_json(path, payload) -> None:
    """先寫暫存檔再 os.replace 原子替換，避免寫到一半被讀到半截或檔案損毀。"""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(dir=str(path.parent), prefix=".tmp_", suffix=".json")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
    except BaseException:
        try:
            os.unlink(tmp)
        except OSError:
            pass
        raise


REV_FILE = data_file("revisions.json")


def get_rev(name: str) -> int:
    """取得某資源的版本號；找不到一律回 0（fail-safe）。"""
    with FILE_LOCK:
        try:
            with REV_FILE.open("r", encoding="utf-8") as handle:
                data = json.load(handle)
            return int(data.get(name, 0)) if isinstance(data, dict) else 0
        except (OSError, ValueError):
            return 0


def bump_rev(name: str) -> int:
    """版本號 +1 並寫回，回傳新版本號。"""
    with FILE_LOCK:
        try:
            with REV_FILE.open("r", encoding="utf-8") as handle:
                data = json.load(handle)
            if not isinstance(data, dict):
                data = {}
        except (OSError, ValueError):
            data = {}
        data[name] = int(data.get(name, 0)) + 1
        atomic_write_json(REV_FILE, data)
        return data[name]
