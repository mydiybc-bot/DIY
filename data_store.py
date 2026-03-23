from __future__ import annotations

from copy import deepcopy
import json
from pathlib import Path

from storage_paths import data_file
from training_data import TRAINING_RULES, TRAINING_SECTIONS

CONTENT_FILE = data_file("training_content.json")


def default_content() -> dict:
    return {
        "rules": deepcopy(TRAINING_RULES),
        "sections": deepcopy(TRAINING_SECTIONS),
    }


class TrainingContentStore:
    def __init__(self, path: Path | None = None) -> None:
        self.path = path or CONTENT_FILE

    def load(self) -> dict:
        if not self.path.exists():
            return default_content()

        with self.path.open("r", encoding="utf-8") as handle:
            payload = json.load(handle)
        return self.validate(payload)

    def save(self, payload: dict) -> dict:
        validated = self.validate(payload)
        with self.path.open("w", encoding="utf-8") as handle:
            json.dump(validated, handle, ensure_ascii=False, indent=2)
        return validated

    def save_validated(self, payload: dict) -> dict:
        with self.path.open("w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)
        return payload

    def validate(self, payload: dict) -> dict:
        if not isinstance(payload, dict):
            raise ValueError("資料格式錯誤")

        incoming_rules = payload.get("rules", {})
        incoming_sections = payload.get("sections", [])

        if not isinstance(incoming_rules, dict):
            raise ValueError("rules 格式錯誤")
        if not isinstance(incoming_sections, list) or not incoming_sections:
            raise ValueError("sections 格式錯誤")

        rules = deepcopy(TRAINING_RULES)
        for key, default_value in TRAINING_RULES.items():
            value = incoming_rules.get(key, default_value)
            if key == "max_attempts_before_answer":
                try:
                    numeric = int(value)
                except (TypeError, ValueError) as exc:
                    raise ValueError("max_attempts_before_answer 必須是數字") from exc
                rules[key] = max(1, numeric)
            else:
                text = str(value).strip()
                rules[key] = text if text else default_value

        sections: list[dict] = []
        for section_index, section in enumerate(incoming_sections, start=1):
            if not isinstance(section, dict):
                raise ValueError(f"第 {section_index} 個單元格式錯誤")

            title = str(section.get("title", "")).strip()
            section_id = str(section.get("id", "")).strip() or f"section_{section_index}"
            questions = section.get("questions", [])

            if not title:
                raise ValueError(f"第 {section_index} 個單元缺少標題")
            if not isinstance(questions, list) or not questions:
                raise ValueError(f"{title} 缺少題目")

            cleaned_questions = []
            for question_index, question in enumerate(questions, start=1):
                if not isinstance(question, dict):
                    raise ValueError(f"{title} 的第 {question_index} 題格式錯誤")

                prompt = str(question.get("prompt", "")).strip()
                answer = str(question.get("answer", "")).strip()
                required_points = question.get("required_points", [])

                if not prompt or not answer:
                    raise ValueError(f"{title} 的第 {question_index} 題缺少題目或標準答案")
                if not isinstance(required_points, list) or not required_points:
                    raise ValueError(f"{title} 的第 {question_index} 題缺少評分重點")

                cleaned_points = []
                for point_index, point in enumerate(required_points, start=1):
                    if not isinstance(point, dict):
                        raise ValueError(f"{title} 的第 {question_index} 題第 {point_index} 個重點格式錯誤")

                    label = str(point.get("label", "")).strip()
                    keywords = point.get("keywords", [])
                    if not label:
                        raise ValueError(f"{title} 的第 {question_index} 題有空白重點名稱")
                    if not isinstance(keywords, list):
                        raise ValueError(f"{title} 的第 {question_index} 題關鍵詞格式錯誤")

                    cleaned_keywords = [str(keyword).strip() for keyword in keywords if str(keyword).strip()]
                    if not cleaned_keywords:
                        raise ValueError(f"{title} 的第 {question_index} 題第 {point_index} 個重點缺少關鍵詞")

                    cleaned_points.append(
                        {
                            "label": label,
                            "keywords": cleaned_keywords,
                        }
                    )

                cleaned_questions.append(
                    {
                        "number": question_index,
                        "prompt": prompt,
                        "answer": answer,
                        "required_points": cleaned_points,
                    }
                )

            sections.append(
                {
                    "id": section_id,
                    "title": title,
                    "questions": cleaned_questions,
                }
            )

        return {"rules": rules, "sections": sections}
