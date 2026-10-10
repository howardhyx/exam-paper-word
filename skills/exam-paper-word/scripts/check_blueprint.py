#!/usr/bin/env python3
"""Audit subject-independent exam budgets and rubric totals; standard library only."""

import argparse
import json
import sys
from collections import defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path


ZERO = Decimal(0)


def mapping(value, where):
    if not isinstance(value, dict):
        raise ValueError(f"{where} must be an object")
    return value


def nonempty_list(value, where):
    if not isinstance(value, list) or not value:
        raise ValueError(f"{where} must be a nonempty array")
    return value


def label(value, where):
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{where} must be a nonempty string")
    return value


def number(value, where, positive=False):
    if isinstance(value, bool) or not isinstance(value, (str, int, float, Decimal)):
        raise ValueError(f"{where} must be numeric")
    try:
        result = Decimal(str(value))
    except InvalidOperation:
        raise ValueError(f"{where} must be numeric") from None
    if not result.is_finite() or result < ZERO or (positive and result == ZERO):
        raise ValueError(f"{where} must be finite and {'positive' if positive else 'nonnegative'}")
    return result


def unique_object(pairs):
    result = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key: {key}")
        result[key] = value
    return result


def reject_constant(value):
    raise ValueError(f"non-finite JSON number: {value}")


def audit(data, stage):
    data = mapping(data, "blueprint")
    errors, warnings = [], []
    expected_total = number(data.get("total_marks"), "total_marks", positive=True)
    allowed = nonempty_list(data.get("allowed_topics"), "allowed_topics")
    for index, topic in enumerate(allowed):
        label(topic, f"allowed_topics[{index}]")
    if len(set(allowed)) != len(allowed):
        raise ValueError("allowed_topics contains duplicate labels")
    allowed = set(allowed)
    allowed_types = data.get("allowed_types")
    if allowed_types is not None:
        allowed_types = nonempty_list(allowed_types, "allowed_types")
        allowed_types = {label(item, "allowed_types entry") for item in allowed_types}

    type_marks, topic_marks, difficulty_marks = (defaultdict(Decimal) for _ in range(3))
    type_counts = defaultdict(int)
    individual_marks = defaultdict(list)
    question_ids, unit_ids = set(), set()
    actual_total, minutes = ZERO, ZERO
    missing_time = False

    for q_index, item in enumerate(nonempty_list(data.get("questions"), "questions")):
        q = mapping(item, f"questions[{q_index}]")
        qid = label(q.get("id"), "question.id")
        if qid in question_ids:
            errors.append(f"duplicate question id: {qid}")
        question_ids.add(qid)
        qtype = label(q.get("type"), f"{qid}.type")
        if allowed_types is not None and qtype not in allowed_types:
            errors.append(f"{qid}: disallowed question type {qtype}")
        marks = number(q.get("marks"), f"{qid}.marks", positive=True)
        type_counts[qtype] += 1
        type_marks[qtype] += marks
        individual_marks[qtype].append(marks)
        actual_total += marks
        if "estimated_minutes" in q:
            minutes += number(q["estimated_minutes"], f"{qid}.estimated_minutes")
        else:
            missing_time = True
        unit_total = ZERO

        for u_index, item in enumerate(nonempty_list(q.get("units"), f"{qid}.units")):
            u = mapping(item, f"{qid}.units[{u_index}]")
            uid = label(u.get("id"), f"{qid}.unit.id")
            if uid in unit_ids:
                errors.append(f"duplicate scoring unit id: {uid}")
            unit_ids.add(uid)
            unit_marks = number(u.get("marks"), f"{uid}.marks", positive=True)
            unit_total += unit_marks
            difficulty = label(u.get("difficulty"), f"{uid}.difficulty")
            difficulty_marks[difficulty] += unit_marks
            allocation = mapping(u.get("topic_scores"), f"{uid}.topic_scores")
            if not allocation:
                raise ValueError(f"{uid}.topic_scores must not be empty")
            allocated_total = ZERO
            for topic, value in allocation.items():
                label(topic, f"{uid}.topic_scores label")
                allocated = number(value, f"{uid}.topic_scores.{topic}")
                if topic not in allowed:
                    errors.append(f"{uid}: unknown topic {topic}")
                allocated_total += allocated
                topic_marks[topic] += allocated
            if allocated_total != unit_marks:
                errors.append(f"{uid}: topic allocation {allocated_total} != unit marks {unit_marks}")

            rubric = u.get("rubric")
            if stage == "final" or rubric is not None:
                rubric_total = ZERO
                for r_index, entry in enumerate(nonempty_list(rubric, f"{uid}.rubric")):
                    entry = mapping(entry, f"{uid}.rubric[{r_index}]")
                    label(entry.get("criterion"), f"{uid}.rubric[{r_index}].criterion")
                    rubric_total += number(entry.get("marks"), f"{uid}.rubric[{r_index}].marks")
                if rubric_total != unit_marks:
                    errors.append(f"{uid}: rubric total {rubric_total} != unit marks {unit_marks}")

        if unit_total != marks:
            errors.append(f"{qid}: unit total {unit_total} != question marks {marks}")

    if actual_total != expected_total:
        errors.append(f"exam total {actual_total} != requested total {expected_total}")
    requirements = mapping(data.get("requirements", {}), "requirements")
    unknown = set(requirements) - {"types", "topics", "difficulty"}
    if unknown:
        raise ValueError(f"unsupported requirements: {sorted(unknown)}")
    distributions = {"types": type_marks, "topics": topic_marks, "difficulty": difficulty_marks}

    for kind, entries in requirements.items():
        for key, rule in mapping(entries, f"requirements.{kind}").items():
            rule = mapping(rule, f"requirements.{kind}.{key}")
            if kind == "topics" and key not in allowed:
                errors.append(f"requirement references unknown topic: {key}")
            permitted = {"count", "marks", "each_marks"} if kind == "types" else {"marks", "min", "max"}
            unknown = set(rule) - permitted
            if unknown:
                raise ValueError(f"unsupported constraint in {kind}.{key}: {sorted(unknown)}")
            actual = distributions[kind].get(key, ZERO)
            for field, value in rule.items():
                target = number(value, f"requirements.{kind}.{key}.{field}")
                if field == "count":
                    if target != target.to_integral_value():
                        raise ValueError(f"{key}.count must be an integer")
                    if type_counts.get(key, 0) != target:
                        errors.append(f"type {key}: count {type_counts.get(key, 0)} != {target}")
                elif field == "each_marks":
                    if any(value != target for value in individual_marks.get(key, [])):
                        errors.append(f"type {key}: some question marks differ from {target}")
                elif field == "marks" and actual != target:
                    errors.append(f"{kind}.{key}: marks {actual} != {target}")
                elif field == "min" and actual < target:
                    errors.append(f"{kind}.{key}: marks {actual} below minimum {target}")
                elif field == "max" and actual > target:
                    errors.append(f"{kind}.{key}: marks {actual} above maximum {target}")

    deviations = {}
    targets = mapping(data.get("targets", {}), "targets")
    if set(targets) - {"topics", "difficulty"}:
        raise ValueError("targets supports topics and difficulty only")
    for kind, entries in targets.items():
        deviations[kind] = {}
        for key, value in mapping(entries, f"targets.{kind}").items():
            if kind == "topics" and key not in allowed:
                raise ValueError(f"target references unknown topic: {key}")
            target = number(value, f"targets.{kind}.{key}")
            actual = distributions[kind].get(key, ZERO)
            deviations[kind][key] = {"target": target, "actual": actual, "delta": actual - target}

    duration = None
    if "duration_minutes" in data:
        duration = number(data["duration_minutes"], "duration_minutes", positive=True)
        if minutes > duration:
            warnings.append("estimated answer time exceeds exam duration; revise workload")
    if missing_time:
        warnings.append("some questions lack estimated_minutes; answer time total is incomplete")
    return {
        "ok": not errors, "stage": stage,
        "total_marks": actual_total, "requested_total_marks": expected_total,
        "type_counts": dict(type_counts), "type_marks": dict(type_marks),
        "topic_marks": {topic: topic_marks.get(topic, ZERO) for topic in sorted(allowed | set(topic_marks))},
        "difficulty_marks": dict(difficulty_marks),
        "estimated_minutes": minutes, "duration_minutes": duration,
        "errors": errors, "warnings": warnings, "soft_target_deviations": deviations,
        "scope": "Arithmetic/constraints only; content, difficulty estimates and Word layout still require review."
    }


def json_number(value):
    if isinstance(value, Decimal):
        return int(value) if value == value.to_integral_value() else float(value)
    raise TypeError(f"cannot serialize {type(value).__name__}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("blueprint", type=Path)
    parser.add_argument("--stage", choices=("plan", "final"), default="plan")
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    try:
        data = json.loads(args.blueprint.read_text(encoding="utf-8-sig"),
                          parse_float=Decimal, object_pairs_hook=unique_object,
                          parse_constant=reject_constant)
        result = audit(data, args.stage)
    except (OSError, ValueError, TypeError) as error:
        result = {"ok": False, "stage": args.stage, "errors": [str(error)]}
    rendered = json.dumps(result, ensure_ascii=False, indent=2, default=json_number)
    print(rendered)
    if args.report:
        try:
            args.report.parent.mkdir(parents=True, exist_ok=True)
            args.report.write_text(rendered + "\n", encoding="utf-8")
        except OSError as error:
            print(f"Could not write report: {error}", file=sys.stderr)
            return 1
    return 0 if result["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
