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

CUSTOMER_FACING_TERMS = (
    "您",
    "你",
    "請問",
    "這邊",
    "幫您",
    "跟您",
    "如果",
    "可以",
    "我們",
    "我先",
    "我幫",
)

CONNECTIVE_TERMS = (
    "因為",
    "所以",
    "如果",
    "先",
    "再",
    "會",
    "可以",
    "請問",
    "幫您",
    "這邊",
)

META_RESPONSE_TERMS = (
    "請再回答一次",
    "應該補上",
    "參考答案",
    "公布答案",
    "評分",
)

PROFANITY_TERMS = (
    "靠北",
    "靠杯",
    "靠邀",
    "靠腰",
    "幹你",
    "干你",
    "幹爆",
    "干爆",
    "幹死",
    "干死",
    "媽的",
    "他媽的",
    "三小",
    "殺小",
    "白痴",
    "智障",
    "北七",
    "垃圾",
    "去死",
)

PROFANITY_PATTERNS = (
    r"幹(?!嘛|麻|話|線|部|事|活|員|道)",
    r"干(?!嘛|麻|話|線|部|事|活|員|道)",
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
        "last_result": None,
    }


def normalize_text(text: str) -> str:
    lowered = text.lower()
    return re.sub(r"[\s，。、「」：:；;,.!?！？（）()～~\-\u3000]", "", lowered)


def build_question_text(question: dict, rules: dict) -> str:
    return f"Q{question['number']}：{question['prompt']} {rules['question_suffix']}"


def _summarize_labels(labels: list[str], limit: int = 2) -> str:
    if not labels:
        return ""
    if len(labels) <= limit:
        return "、".join(labels)
    return "、".join(labels[:limit]) + "等重點"


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

    point_count = max(len(question["required_points"]), 1)
    matched_ratio = len(matched) / point_count
    polite_hits = sum(1 for term in POLITE_TERMS if term in answer)
    customer_hits = sum(1 for term in CUSTOMER_FACING_TERMS if term in answer)
    connective_hits = sum(1 for term in CONNECTIVE_TERMS if term in answer)
    punctuation_hits = sum(answer.count(term) for term in "。！？!?～~")
    delimiter_hits = sum(answer.count(term) for term in "、，,；;")
    copied_labels = sum(1 for point in question["required_points"] if normalize_text(point["label"]) in normalized_answer)
    has_meta_response = any(term in answer for term in META_RESPONSE_TERMS)
    uses_profanity = any(normalize_text(term) in normalized_answer for term in PROFANITY_TERMS) or any(
        re.search(pattern, answer) for pattern in PROFANITY_PATTERNS
    )
    minimum_length = max(18, point_count * 6)
    has_enough_length = len(normalized_answer) >= minimum_length
    looks_like_outline = copied_labels >= max(2, point_count - 1) or (
        delimiter_hits >= max(2, point_count - 1)
        and customer_hits == 0
        and connective_hits <= 1
        and punctuation_hits <= 1
    )

    keyword_score = round(matched_ratio * 6)
    tone_score = 2 if polite_hits >= 2 else 1 if (polite_hits == 1 or customer_hits >= 1) else 0
    expression_score = 0
    if has_enough_length:
        expression_score += 1
    if not looks_like_outline and not has_meta_response and (customer_hits >= 1 or connective_hits >= 2 or punctuation_hits >= 2):
        expression_score += 1

    raw_score = keyword_score + tone_score + expression_score
    score = min(10, max(0, raw_score))

    needs_better_expression = not has_enough_length or looks_like_outline or has_meta_response or (customer_hits == 0 and connective_hits <= 1)

    if not missing and not needs_better_expression and tone_score >= 1:
        score = 10
    elif not missing:
        score = min(score, 9)

    if uses_profanity:
        score = min(score, 2)

    if score == 10:
        feedback = rules["pass_feedback"]
        coaching = "很好，這題已經達標。"
    else:
        feedback = rules["retry_feedback"]
        coaching_parts = []
        if missing:
            coaching_parts.append("請補上：" + _summarize_labels(missing) + "。")
        if uses_profanity:
            coaching_parts.append("回答不能帶髒話或攻擊字眼，請改成尊重客人的說法。")
        if needs_better_expression:
            coaching_parts.append("請改成直接對客人說的完整句子，語氣自然一點。")
        if not missing and not uses_profanity and needs_better_expression:
            coaching_parts.append("離 10 分只差說得更自然、完整。")
        coaching_parts.append(rules["retry_prompt"])
        coaching = "".join(coaching_parts)

    return {
        "score": score,
        "matched": matched,
        "matched_count": len(matched),
        "missing": missing,
        "missing_count": len(missing),
        "needs_better_expression": needs_better_expression,
        "uses_profanity": uses_profanity,
        "feedback": feedback,
        "coaching": coaching,
        "score_protected": False,
    }


def protect_improved_attempt(previous: dict | None, current: dict) -> dict:
    if not previous:
        return current

    has_more_points = current["matched_count"] > previous["matched_count"]
    has_fewer_missing = current["missing_count"] < previous["missing_count"]
    expression_improved = previous["needs_better_expression"] and not current["needs_better_expression"]
    profanity_removed = previous["uses_profanity"] and not current["uses_profanity"]

    added_penalties = (
        current["uses_profanity"]
        or current["matched_count"] < previous["matched_count"]
        or (current["needs_better_expression"] and not previous["needs_better_expression"])
    )
    looks_like_progress = has_more_points or has_fewer_missing or expression_improved or profanity_removed

    if looks_like_progress and not added_penalties and current["score"] < previous["score"]:
        current["score"] = previous["score"]
        current["score_protected"] = True

    return current


def summarize_session(session: dict) -> str:
    attempted = session["scoreboard"]
    rules = session["rules"]
    title = session.get("section_title", "未知單元")
    if not attempted:
        return (
            f"本次練習主題：{title}\n"
            "本次練習題目總數：0\n"
            "每一題得分：尚未作答\n"
            "平均得分：0.0\n"
            f"{rules['summary_intro_if_empty']}"
        )

    scores = [item["best_score"] for item in attempted]
    score_lines = [f"Q{item['number']}：{item['best_score']} 分" for item in attempted]
    average = sum(scores) / len(scores)
    return (
        f"本次練習主題：{title}\n"
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
    previous_result = session.get("last_result")
    result = protect_improved_attempt(previous_result, result)
    session["last_result"] = {
        "score": result["score"],
        "matched_count": result["matched_count"],
        "missing_count": result["missing_count"],
        "needs_better_expression": result["needs_better_expression"],
        "uses_profanity": result["uses_profanity"],
    }

    record = _ensure_scoreboard_item(session, question)
    record["attempts"] += 1
    record["best_score"] = max(record["best_score"], result["score"])

    parts = [
        result["feedback"],
        f"評分：{result['score']}/10",
    ]

    if result["score"] < 10:
        if result["score_protected"]:
            parts.append("這次有依照建議補強，分數先不倒扣。")
        parts.append(result["coaching"])
        if session["attempts"] >= rules["max_attempts_before_answer"]:
            parts.append(rules["reference_answer_intro"] + question["answer"])
            parts.append(rules["answer_reveal_prompt"])
        return {"done": False, "message": "\n".join(parts)}

    session["attempts"] = 0
    session["last_result"] = None
    session["current_index"] += 1
    parts.append(rules["pass_message"])

    if session["current_index"] >= len(session["questions"]):
        parts.append("本單元題目已完成。")
        parts.append(summarize_session(session))
        return {"done": True, "message": "\n".join(parts)}

    next_question = session["questions"][session["current_index"]]
    parts.append(build_question_text(next_question, rules))
    return {"done": False, "message": "\n".join(parts)}
