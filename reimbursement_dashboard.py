from __future__ import annotations

from collections import defaultdict
from datetime import datetime
import os
from pathlib import Path
import re

import xlrd


DEFAULT_SOURCE_DIR = Path(
    os.environ.get("REIMBURSEMENT_SOURCE_DIR", "/Users/diybc/Desktop/個人請款檔案/2026")
)
DATE_PATTERN = re.compile(r"(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日")
STOP_MARKER = "廠商款請統一填寫"


def load_reimbursement_dashboard(source_dir: Path | None = None) -> dict:
    directory = (source_dir or DEFAULT_SOURCE_DIR).expanduser()
    files = sorted(directory.glob("*.xls"))
    records: list[dict] = []
    file_summaries: list[dict] = []
    errors: list[dict] = []

    for path in files:
        try:
            parsed = _parse_workbook(path)
        except Exception as exc:  # pragma: no cover - dashboard should still render partial data
            errors.append({"file_name": path.name, "error": str(exc)})
            continue

        records.extend(parsed["records"])
        file_summaries.append(parsed["file_summary"])

    records.sort(key=lambda item: (item["date"], item["file_name"], item["row_number"]), reverse=True)
    file_summaries.sort(key=lambda item: (item["date"], item["file_name"]), reverse=True)

    latest_source_update = max((item["modified_at"] for item in file_summaries), default=None)

    return {
        "ok": True,
        "source_dir": str(directory),
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "source_updated_at": latest_source_update,
        "file_count": len(file_summaries),
        "record_count": len(records),
        "summary": _build_summary(records, file_summaries),
        "files": file_summaries,
        "records": records,
        "errors": errors,
    }


def _parse_workbook(path: Path) -> dict:
    workbook = xlrd.open_workbook(path)
    sheet = workbook.sheet_by_index(0)

    unit = _normalize_text(sheet.cell_value(1, 1))
    date_text = " ".join(_normalize_text(sheet.cell_value(1, column)) for column in range(sheet.ncols)).strip()
    date = _parse_date(date_text)

    records: list[dict] = []
    for row_index in range(3, sheet.nrows):
        summary = _normalize_text(sheet.cell_value(row_index, 0))
        if STOP_MARKER in summary:
            break

        code = _normalize_text(sheet.cell_value(row_index, 1))
        invoice = _normalize_text(sheet.cell_value(row_index, 2))
        amount = _coerce_amount(sheet.cell_value(row_index, 3))
        payee = _normalize_text(sheet.cell_value(row_index, 4))
        note = _normalize_text(sheet.cell_value(row_index, 5))

        if not any([summary, code, invoice, amount, payee, note]):
            continue

        if not any([summary, code, invoice, note]) and amount == 0:
            continue

        records.append(
            {
                "date": date,
                "month": date[:7] if date else "",
                "unit": unit,
                "summary": summary or "未填項目",
                "code": code,
                "invoice": invoice,
                "amount": amount,
                "payee": payee,
                "note": note,
                "file_name": path.name,
                "file_total": 0,
                "row_number": row_index + 1,
            }
        )

    file_total = _round_amount(sum(item["amount"] for item in records))
    for item in records:
        item["file_total"] = file_total

    return {
        "records": records,
        "file_summary": {
            "file_name": path.name,
            "date": date,
            "month": date[:7] if date else "",
            "unit": unit,
            "record_count": len(records),
            "nonzero_record_count": sum(1 for item in records if item["amount"] > 0),
            "zero_amount_count": sum(1 for item in records if item["amount"] == 0),
            "total_amount": file_total,
            "modified_at": datetime.fromtimestamp(path.stat().st_mtime).isoformat(timespec="seconds"),
        },
    }


def _build_summary(records: list[dict], file_summaries: list[dict]) -> dict:
    total_amount = _round_amount(sum(item["amount"] for item in records))
    nonzero_records = [item for item in records if item["amount"] > 0]
    zero_amount_count = sum(1 for item in records if item["amount"] == 0)

    by_month: dict[str, dict] = defaultdict(lambda: {"month": "", "amount": 0.0, "record_count": 0})
    by_unit: dict[str, dict] = defaultdict(lambda: {"unit": "", "amount": 0.0, "record_count": 0})
    by_category: dict[str, dict] = defaultdict(lambda: {"summary": "", "amount": 0.0, "record_count": 0})

    for item in records:
        month_key = item["month"]
        unit_key = item["unit"]
        category_key = item["summary"]

        by_month[month_key]["month"] = month_key
        by_month[month_key]["amount"] += item["amount"]
        by_month[month_key]["record_count"] += 1

        by_unit[unit_key]["unit"] = unit_key
        by_unit[unit_key]["amount"] += item["amount"]
        by_unit[unit_key]["record_count"] += 1

        by_category[category_key]["summary"] = category_key
        by_category[category_key]["amount"] += item["amount"]
        by_category[category_key]["record_count"] += 1

    months = sorted(
        (
            {"month": key, "amount": _round_amount(value["amount"]), "record_count": value["record_count"]}
            for key, value in by_month.items()
        ),
        key=lambda item: item["month"],
    )
    units = sorted(
        (
            {"unit": key, "amount": _round_amount(value["amount"]), "record_count": value["record_count"]}
            for key, value in by_unit.items()
        ),
        key=lambda item: item["amount"],
        reverse=True,
    )
    categories = sorted(
        (
            {"summary": key, "amount": _round_amount(value["amount"]), "record_count": value["record_count"]}
            for key, value in by_category.items()
        ),
        key=lambda item: item["amount"],
        reverse=True,
    )

    highest_month = max(months, key=lambda item: item["amount"], default=None)
    max_record = max(records, key=lambda item: item["amount"], default=None)

    return {
        "total_amount": total_amount,
        "record_count": len(records),
        "nonzero_record_count": len(nonzero_records),
        "zero_amount_count": zero_amount_count,
        "average_amount": _round_amount(total_amount / len(nonzero_records)) if nonzero_records else 0,
        "highest_month": highest_month,
        "largest_record": max_record,
        "months": months,
        "units": units,
        "categories": categories[:10],
        "latest_date": max((item["date"] for item in records if item["date"]), default=None),
        "file_count": len(file_summaries),
    }


def _normalize_text(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return f"{value:.2f}".rstrip("0").rstrip(".")
    return str(value).replace("\xa0", " ").strip()


def _coerce_amount(value: object) -> int | float:
    if value in ("", None):
        return 0
    if isinstance(value, (int, float)):
        return _round_amount(float(value))

    text = _normalize_text(value).replace(",", "")
    if not text:
        return 0

    try:
        return _round_amount(float(text))
    except ValueError:
        return 0


def _round_amount(value: float) -> int | float:
    rounded = round(value, 2)
    return int(rounded) if float(rounded).is_integer() else rounded


def _parse_date(value: str) -> str:
    matched = DATE_PATTERN.search(value)
    if not matched:
        return ""
    year, month, day = matched.groups()
    return f"{int(year):04d}-{int(month):02d}-{int(day):02d}"
