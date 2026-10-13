"""登入首頁的「各店人員成績表」(公開,不需登入)。

只回傳姓名、店別與練習統計,不含員編、密碼或作答內容。
只列門市(店別以店號開頭,例「4 台北士林店」)與「未設定店別」;北區、中南區、各部門、其他是主管或總部,
不需要練習,不列(2026-10-13 經營者指示)。
同一位員工同一單元同一題若出現在多筆練習紀錄(中途結束後接續練),以最新一筆為準,避免重複計算。
"""
from __future__ import annotations

import re

_STORE_NO = re.compile(r"^\s*(\d+)\s*")
_SCORE_ITEM = re.compile(r"^Q(\d+):(\d+)")
UNASSIGNED = "未設定店別"


def is_listed_store(store: str) -> bool:
    """首頁成績表要列的分組:門市(店號開頭)與未設定店別。"""
    return bool(_STORE_NO.match(store)) or store == UNASSIGNED


def _store_sort_key(store: str) -> tuple:
    match = _STORE_NO.match(store)
    if match:
        return (0, int(match.group(1)), store)
    if store == UNASSIGNED:
        return (2, 0, store)
    return (1, 0, store)


def _question_records(report: dict) -> list[dict]:
    """每題的 best_score / first_score / attempts;舊紀錄只有 scores 字串,只取得到最佳分。"""
    details = report.get("details")
    if isinstance(details, list):
        return [
            {
                "number": d.get("number"),
                "best_score": d.get("best_score", 0),
                "first_score": d.get("first_score"),
                "attempts": d.get("attempts", 0),
                "has_detail": True,
            }
            for d in details
            if isinstance(d, dict)
        ]
    records = []
    for text in report.get("scores", []) or []:
        match = _SCORE_ITEM.match(str(text))
        if match:
            records.append({"number": int(match.group(1)), "best_score": int(match.group(2)), "has_detail": False})
    return records


def build_public_scoreboard(content: dict, auth: dict, reports: list[dict], progress: dict) -> dict:
    sections = [
        {"id": s["id"], "title": s["title"], "total": len(s.get("questions", []))}
        for s in content.get("sections", [])
    ]
    section_total = {s["id"]: s["total"] for s in sections}

    employee_reports: dict[str, list[dict]] = {}
    for report in reports:
        if report.get("role", "employee") != "employee":
            continue
        employee_id = report.get("employee_id")
        if employee_id:
            employee_reports.setdefault(employee_id, []).append(report)

    groups: dict[str, list[dict]] = {}
    for employee in auth.get("employees", []):
        store_name = (employee.get("store") or "").strip() or UNASSIGNED
        if not is_listed_store(store_name):
            continue
        employee_id = employee.get("employee_id", "")
        own_reports = sorted(employee_reports.get(employee_id, []), key=lambda r: r.get("created_at", ""))

        latest: dict[tuple, dict] = {}
        completed: set[str] = set()
        for report in own_reports:
            section_id = report.get("section_id", "")
            if section_id in section_total and report.get("question_count", 0) >= section_total[section_id] > 0:
                completed.add(section_id)
            for record in _question_records(report):
                latest[(section_id, record["number"])] = record

        records = list(latest.values())
        detailed = [r for r in records if r["has_detail"]]
        answered = [r for r in detailed if (r.get("attempts") or 0) > 0]

        employee_progress = progress.get(employee_id, {}) if isinstance(progress, dict) else {}
        units = []
        for section in sections:
            total = section["total"]
            done = 0
            snapshot = employee_progress.get(section["id"])
            if isinstance(snapshot, dict):
                try:
                    done = max(0, min(int(snapshot.get("current_index", 0)), total))
                except (TypeError, ValueError):
                    done = 0
            finished = section["id"] in completed or (total > 0 and done >= total)
            units.append({"done": total if finished else done, "total": total, "completed": finished})

        groups.setdefault(store_name, []).append(
            {
                "name": employee.get("employee_name", ""),
                "units": units,
                "completed_units": sum(1 for u in units if u["completed"]),
                "first_try_pass_rate": round(sum(1 for r in detailed if r.get("first_score") == 10) / len(detailed) * 100)
                if detailed
                else None,
                "avg_attempts": round(sum(r["attempts"] for r in answered) / len(answered), 1) if answered else None,
                "average_score": round(sum(r["best_score"] for r in records) / len(records), 1) if records else None,
                "last_practice": own_reports[-1].get("created_at", "")[:10] if own_reports else "",
            }
        )

    stores = []
    for store in sorted(groups, key=_store_sort_key):
        people = sorted(
            groups[store],
            key=lambda p: (-p["completed_units"], -sum(u["done"] for u in p["units"]), p["name"]),
        )
        stores.append(
            {
                "store": store,
                "people": people,
                "practiced": sum(1 for p in people if p["last_practice"] or any(u["done"] for u in p["units"])),
            }
        )
    return {"sections": [{"title": s["title"], "total": s["total"]} for s in sections], "stores": stores}
