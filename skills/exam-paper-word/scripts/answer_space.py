#!/usr/bin/env python3
"""Plan and audit handwriting capacity. Standard library; no DOCX dependency."""

import argparse
import json
import math
from pathlib import Path


def numeric(value, name, *, positive=False, integer=False):
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError(f"{name} must be a number")
    if not math.isfinite(value) or value < 0 or (positive and value == 0):
        raise ValueError(f"{name} must be finite and {'positive' if positive else 'nonnegative'}")
    if integer and int(value) != value:
        raise ValueError(f"{name} must be an integer")
    return int(value) if integer else float(value)


def plan_space(spec, width_mm):
    """Return minimum and configured capacity; this is not a rendered measurement."""
    if not isinstance(spec, dict):
        raise ValueError("answer space must be an object")
    mode = spec.get("mode")
    if mode not in {"none", "blank", "lined", "grid", "drawing"}:
        raise ValueError("mode must be none, blank, lined, grid or drawing")
    if mode == "none":
        return {"mode": mode, "height_mm": 0, "minimum_height_mm": 0}
    width = numeric(width_mm, "width_mm", positive=True)
    basis = spec.get("basis")
    if not isinstance(basis, str) or not basis.strip():
        raise ValueError("basis must explain the expected handwritten work")
    drawing = numeric(spec.get("drawing_height_mm", 0), "drawing_height_mm")
    reserve = numeric(spec.get("reserve_lines", 0), "reserve_lines", integer=True)
    result = {"mode": mode, "width_mm": width, "drawing_height_mm": drawing}
    if mode in {"blank", "lined"}:
        lines = numeric(spec.get("estimated_lines"), "estimated_lines", positive=True, integer=True)
        line_height = numeric(spec.get("line_height_mm", 9), "line_height_mm", positive=True)
        if line_height < 5:
            raise ValueError("line_height_mm below 5 is too small for this handwriting component")
        minimum = (lines + reserve) * line_height + drawing
        allocated = numeric(spec.get("allocated_height_mm", minimum), "allocated_height_mm", positive=True)
        if allocated + 0.01 < minimum:
            raise ValueError(f"allocated height {allocated:g} mm is below required {minimum:g} mm")
        writing_height = allocated - drawing
        rows = math.floor((writing_height + 0.001) / line_height)
        result.update(writing_lines=rows, row_height_mm=writing_height / rows,
                      estimated_lines=lines, required_lines=lines + reserve,
                      height_mm=allocated, minimum_height_mm=minimum)
    elif mode == "grid":
        if drawing:
            raise ValueError("use a separate drawing space alongside a writing grid")
        chars = numeric(spec.get("estimated_chars"), "estimated_chars", positive=True, integer=True)
        reserve_chars = numeric(spec.get("reserve_chars", 0), "reserve_chars", integer=True)
        cell = numeric(spec.get("cell_size_mm", 8), "cell_size_mm", positive=True)
        if cell < 5:
            raise ValueError("cell_size_mm below 5 is too small for this handwriting component")
        columns = math.floor((width + 0.001) / cell)
        if columns < 1:
            raise ValueError("available width is smaller than one writing cell")
        min_rows = math.ceil((chars + reserve_chars) / columns)
        minimum = min_rows * cell
        allocated = numeric(spec.get("allocated_height_mm", minimum), "allocated_height_mm", positive=True)
        if allocated + 0.01 < minimum:
            raise ValueError(f"allocated grid height {allocated:g} mm is below required {minimum:g} mm")
        rows = math.ceil((allocated - 0.001) / cell)
        result.update(width_mm=columns * cell, columns=columns, writing_lines=rows,
                      row_height_mm=cell, grid_cells=columns * rows,
                      required_chars=chars + reserve_chars, height_mm=rows * cell,
                      minimum_height_mm=minimum)
    else:
        if drawing <= 0:
            raise ValueError("drawing mode requires positive drawing_height_mm")
        allocated = numeric(spec.get("allocated_height_mm", drawing), "allocated_height_mm", positive=True)
        if allocated + 0.01 < drawing:
            raise ValueError("allocated height is below required drawing height")
        result.update(drawing_height_mm=allocated, height_mm=allocated,
                      minimum_height_mm=drawing)
    return result


def audit_spaces(data, stage="plan"):
    """Audit explicit per-question spaces, without inferring them from score units."""
    if not isinstance(data, dict) or not isinstance(data.get("questions"), list) or not data["questions"]:
        raise ValueError("blueprint must contain a nonempty questions array")
    errors, warnings, reports = [], [], []
    used_ids = set()
    for question in data["questions"]:
        if not isinstance(question, dict):
            raise ValueError("question must be an object")
        qid = question.get("id", "?")
        spaces = question.get("answer_spaces")
        if spaces is None:
            (errors if stage == "final" else warnings).append(f"{qid}: answer_spaces missing")
            continue
        if not isinstance(spaces, list):
            errors.append(f"{qid}: answer_spaces must be an array (empty for no separate space)")
            continue
        for spec in spaces:
            sid = "?"
            try:
                if not isinstance(spec, dict):
                    raise ValueError("space must be an object")
                sid = spec.get("id")
                if not isinstance(sid, str) or not sid.strip():
                    raise ValueError("space.id must be a nonempty string")
                if sid in used_ids:
                    raise ValueError("duplicate answer space id")
                used_ids.add(sid)
                plan = plan_space(spec, spec.get("width_mm"))
                reports.append({"question_id": qid, "space_id": sid, "configured": plan})
                if stage != "final" or plan["mode"] == "none":
                    continue
                observed = spec.get("observed")
                if not isinstance(observed, dict):
                    raise ValueError("observed measurements from page review are required in final stage")
                height = numeric(observed.get("height_mm"), "observed.height_mm", positive=True)
                # Small allowance only for physical measurement / DOCX unit rounding.
                if height + 0.5 < plan["height_mm"]:
                    raise ValueError("observed height is below the configured allocation")
                width = numeric(observed.get("width_mm"), "observed.width_mm", positive=True)
                if abs(width - plan["width_mm"]) > 0.5:
                    raise ValueError("observed width differs from configured width")
                if plan["mode"] in {"blank", "lined"}:
                    actual_lines = numeric(observed.get("writing_lines"), "observed.writing_lines", integer=True)
                    if actual_lines < plan["writing_lines"]:
                        raise ValueError("observed handwriting capacity is below the allocated line count")
                elif plan["mode"] == "grid":
                    cells = numeric(observed.get("grid_cells"), "observed.grid_cells", integer=True)
                    if cells < plan["grid_cells"]:
                        raise ValueError("observed grid has fewer cells than allocated")
                if plan["drawing_height_mm"]:
                    actual_drawing = numeric(observed.get("drawing_height_mm"), "observed.drawing_height_mm")
                    if actual_drawing + 0.5 < plan["drawing_height_mm"]:
                        raise ValueError("observed drawing height is insufficient")
                reports[-1]["observed"] = observed
            except (ValueError, TypeError) as error:
                errors.append(f"{qid}/{sid}: {error}")
    return {"ok": not errors, "stage": stage, "spaces": reports, "errors": errors,
            "warnings": warnings,
            "scope": "Capacity arithmetic and supplied observations only; actual work and page review remain necessary."}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("blueprint", type=Path)
    parser.add_argument("--stage", choices=("plan", "final"), default="plan")
    parser.add_argument("--report", type=Path)
    args = parser.parse_args()
    try:
        data = json.loads(args.blueprint.read_text(encoding="utf-8-sig"))
        report = audit_spaces(data, args.stage)
    except (OSError, ValueError, TypeError) as error:
        report = {"ok": False, "errors": [str(error)]}
    rendered = json.dumps(report, ensure_ascii=False, indent=2)
    print(rendered)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(rendered + "\n", encoding="utf-8")
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
