#!/usr/bin/env python3
"""Local-only text diagnostics. Candidates do not certify facts or release readiness."""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path


PATTERNS = [
    ("bracketed", re.compile(r"[【\[（(〔]\s*(?:待补|待核|待确认|待定|补充|核实)[^】\]）)〕\n]{0,120}[】\]）)〕]")),
    ("x_marker", re.compile(r"(?<![A-Za-z])(?:X{2,}|x{2,})(?![A-Za-z])")),
    ("todo_marker", re.compile(r"(?<![A-Za-z0-9_])(?:TODO|TBD)(?![A-Za-z0-9_])", re.I)),
    ("blank_marker", re.compile(r"_{2,}")),
]


def diagnose(text: str, min_chars: int | None = None, max_chars: int | None = None) -> dict:
    if min_chars is not None and min_chars < 0:
        raise ValueError("min_chars must be nonnegative")
    if max_chars is not None and max_chars < 0:
        raise ValueError("max_chars must be nonnegative")
    if min_chars is not None and max_chars is not None and min_chars > max_chars:
        raise ValueError("min_chars must not exceed max_chars")
    findings = []
    occupied = []
    for kind, pattern in PATTERNS:
        for match in pattern.finditer(text):
            if any(a <= match.start() and match.end() <= b for a, b in occupied):
                continue
            occupied.append(match.span())
            line = text.count("\n", 0, match.start()) + 1
            last_newline = text.rfind("\n", 0, match.start())
            findings.append({"kind": kind, "line": line, "column": match.start() - last_newline,
                             "text": match.group(), "offset": match.start()})
    findings.sort(key=lambda item: item["offset"])
    char_count = len(re.sub(r"\s", "", text))
    han_count = len(re.findall(r"[\u3400-\u4dbf\u4e00-\u9fff\U00020000-\U0002ebef]", text))
    length_issues = []
    if min_chars is not None and char_count < min_chars:
        length_issues.append({"kind": "below_requested_minimum", "requested": min_chars, "actual": char_count})
    if max_chars is not None and char_count > max_chars:
        length_issues.append({"kind": "above_requested_maximum", "requested": max_chars, "actual": char_count})
    paragraphs = [p.strip() for p in re.split(r"\n\s*\n", text) if p.strip()]
    seen = {}
    repeated = []
    for index, paragraph in enumerate(paragraphs, start=1):
        if len(re.sub(r"\s", "", paragraph)) < 16:
            continue
        if paragraph in seen:
            repeated.append({"paragraph": index, "same_as": seen[paragraph], "text": paragraph})
        else:
            seen[paragraph] = index
    return {
        "schema_version": 1,
        "diagnostic_completed": True,
        "non_whitespace_characters": char_count,
        "han_characters": han_count,
        "requested_range": {"minimum": min_chars, "maximum": max_chars},
        "placeholder_candidates": findings,
        "length_issues": length_issues,
        "exact_repeated_paragraphs": repeated,
        "fact_check_performed": False,
        "scope_note": "候选提示须结合语境查看。待补草稿可正常保留占位；空白位置不能全部自动识别。字符统计不等于Word字数，诊断不判定政策、事实或可否报送。",
    }


def main() -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8")
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("input", type=Path, help="UTF-8 .md or .txt file")
    parser.add_argument("--min-chars", type=int)
    parser.add_argument("--max-chars", type=int)
    parser.add_argument("--json", action="store_true", help="Emit machine-readable JSON")
    parser.add_argument("--fail-on-placeholders", action="store_true",
                        help="Explicit check: exit 1 if candidate placeholders exist; do not use as a draft-writing gate")
    args = parser.parse_args()
    try:
        if args.input.suffix.lower() not in {".md", ".txt"}:
            raise ValueError("only UTF-8 .md and .txt inputs are supported; extract DOCX separately")
        text = args.input.read_text(encoding="utf-8-sig")
        result = diagnose(text, args.min_chars, args.max_chars)
    except (OSError, UnicodeError, ValueError) as exc:
        print(json.dumps({"diagnostic_completed": False, "error": str(exc)}, ensure_ascii=False))
        return 2
    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        print("非空白字符：{}；汉字：{}。".format(result["non_whitespace_characters"], result["han_characters"]))
        for finding in result["placeholder_candidates"]:
            print("候选占位 {}:{} {}".format(finding["line"], finding["column"], finding["text"]))
        for issue in result["length_issues"]:
            print("篇幅提示：{}，实际{}，要求{}。".format(issue["kind"], issue["actual"], issue["requested"]))
        for repeat in result["exact_repeated_paragraphs"]:
            print("重复候选：第{}段与第{}段相同。".format(repeat["paragraph"], repeat["same_as"]))
        print(result["scope_note"])
    return 1 if args.fail_on_placeholders and result["placeholder_candidates"] else 0


if __name__ == "__main__":
    raise SystemExit(main())
