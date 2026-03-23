from __future__ import annotations

from copy import deepcopy
import re

POLITE_TERMS = (
    "您好",
    "抱歉",
    "不好意思",
    "理解",
    "謝謝",
    "請",
    "麻煩",
    "我幫",
    "可以",
    "您",
    "辛苦",
    "沒關係",
    "我懂",
)


def build_catalog(sections: list[dict]) -> dict[str, dict]:
    return {section["id"]: deepcopy(section) for section in sections}


def create_session(section_id: str, content: dict, progress: dict | None = None) -> dict:
    catalog = build_catalog(content["sections"])
    if section_id not in catalog:
        raise ValueError("unknown training section")

    start_index = 0
    scoreboard = []
    if progress:
        try:
            start_index = max(0, int(progress.get("current_index", 0)))
        except (TypeError, ValueError):
            start_index = 0
        saved_scoreboard = progress.get("scoreboard", [])
        if isinstance(saved_scoreboard, list):
            scoreboard = deepcopy(saved_scoreboard)

    return {
        "section_id": section_id,
        "section_title": catalog[section_id]["title"],
        "questions": catalog[section_id]["questions"],
        "rules": deepcopy(content["rules"]),
        "current_index": start_index,
        "attempts": 0,
        "scoreboard": scoreboard,
    }


def normalize_text(text: str) -> str:
    lowered = text.lower()
    return re.sub(r"[\s，。、「」：:；;,.!?！？（）()～~\-\u3000]", "", lowered)


def build_question_text(question: dict, rules: dict) -> str:
    return f"Q{question['number']}：{question['prompt']} {rules['question_suffix']}"


def score_answer(question: dict, answer: str, rules: dict) -> dict:
    normalized_answer = normalize_text(answer)
    matched = []
    missing = []

    for point in question["required_points"]:
        keywords = point["keywords"]
        if any(normalize_text(keyword) in normalized_answer for keyword in keywords):
            matched.append(point["label"])
        else:
            missing.append(point["label"])

    matched_ratio = len(matched) / max(len(question["required_points"]), 1)
    polite_hits = sum(1 for term in POLITE_TERMS if term in answer)
    polite_score = 2 if polite_hits >= 2 else 1 if polite_hits == 1 else 0
    raw_score = round(matched_ratio * 8 + polite_score)
    score = min(10, max(0, raw_score))

    if not missing:
        score = 10

    if score == 10:
        feedback = rules["pass_feedback"]
        coaching = "很好，這題已經達標。"
    else:
        feedback = rules["retry_feedback"]
        coaching = "應該補上：" + "、".join(missing) + "。" + rules["retry_prompt"]

    return {
        "score": score,
        "matched": matched,
        "missing": missing,
        "feedback": feedback,
        "coaching": coaching,
    }


def summarize_session(session: dict) -> str:
    attempted = session["scoreboard"]
    rules = session["rules"]
    if not attempted:
        return (
            "本次練習題目總數：0\n"
            "每一題得分：尚未作答\n"
            "平均得分：0.0\n"
            f"{rules['summary_intro_if_empty']}"
        )

    scores = [item["best_score"] for item in attempted]
    score_lines = [f"Q{item['number']}：{item['best_score']} 分" for item in attempted]
    average = sum(scores) / len(scores)
    return (
        f"本次練習題目總數：{len(attempted)}\n"
        f"每一題得分：{'；'.join(score_lines)}\n"
        f"平均得分：{average:.1f}\n"
        f"{rules['summary_encouragement']}"
    )


def build_report_record(session: dict) -> dict:
    attempted = session["scoreboard"]
    scores = [item["best_score"] for item in attempted]
    average = round(sum(scores) / len(scores), 1) if scores else 0.0
    return {
        "section_id": session["section_id"],
        "section_title": session["section_title"],
        "question_count": len(attempted),
        "scores": [f"Q{item['number']}:{item['best_score']}" for item in attempted],
        "average_score": average,
    }


def build_progress_snapshot(session: dict) -> dict:
    return {
        "current_index": session["current_index"],
        "scoreboard": deepcopy(session["scoreboard"]),
    }


def _ensure_scoreboard_item(session: dict, question: dict) -> dict:
    for item in session["scoreboard"]:
        if item["number"] == question["number"]:
            return item

    item = {"number": question["number"], "best_score": 0, "attempts": 0}
    session["scoreboard"].append(item)
    return item


def respond(session: dict, answer: str) -> dict:
    clean_answer = answer.strip()
    rules = session["rules"]

    if clean_answer == rules["end_phrase"]:
        return {"done": True, "message": summarize_session(session)}

    question = session["questions"][session["current_index"]]
    session["attempts"] += 1
    result = score_answer(question, clean_answer, rules)

    record = _ensure_scoreboard_item(session, question)
    record["attempts"] += 1
    record["best_score"] = max(record["best_score"], result["score"])

    parts = [
        result["feedback"],
        f"評分：{result['score']}/10",
    ]

    if result["score"] < 10:
        parts.append(result["coaching"])
        if session["attempts"] >= rules["max_attempts_before_answer"]:
            parts.append(rules["reference_answer_intro"] + question["answer"])
            parts.append(rules["answer_reveal_prompt"])
        return {"done": False, "message": "\n".join(parts)}

    session["attempts"] = 0
    session["current_index"] += 1
    parts.append(rules["pass_message"])

    if session["current_index"] >= len(session["questions"]):
        parts.append("本單元題目已完成。")
        parts.append(summarize_session(session))
        return {"done": True, "message": "\n".join(parts)}

    next_question = session["questions"][session["current_index"]]
    parts.append(build_question_text(next_question, rules))
    return {"done": False, "message": "\n".join(parts)}
