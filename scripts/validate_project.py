#!/usr/bin/env python3
"""Check skill metadata, local links, examples, Python syntax and workflow YAML."""
from __future__ import annotations

import ast
import json
import re
import sys
from pathlib import Path
from urllib.parse import unquote, urlsplit

try:
    import yaml
except ImportError:
    raise SystemExit("Install requirements-dev.txt to run project validation (PyYAML required).")

from build_release import ROOT, collect_files
from gongwen_docx import DEFAULT_PROFILES, load_profile, validate_payload


def validate(root: Path = ROOT) -> tuple[int, list[str]]:
    root = root.resolve()
    files = collect_files(root, source_archive=True)
    errors = []
    text = files["SKILL.md"].decode("utf-8")
    front = re.match(r"\A---\n(.*?)\n---\n", text, re.S)
    metadata = yaml.safe_load(front.group(1)) if front else None
    if not isinstance(metadata, dict):
        raise ValueError("SKILL.md must contain YAML frontmatter")
    allowed = {"name", "description", "license", "allowed-tools", "metadata"}
    if set(metadata) - allowed:
        errors.append("SKILL.md contains unsupported frontmatter fields")
    name = metadata.get("name", "")
    if not isinstance(name, str) or len(name) > 64 or not re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", name):
        errors.append("Invalid skill name")
    description = metadata.get("description", "")
    if not isinstance(description, str) or not description.strip() or len(description) > 1024:
        errors.append("Skill description must contain 1–1024 characters")
    info = json.loads(files["assets/package-info.json"])
    if info["name"] != name:
        errors.append("Skill name differs from package-info.json")
    if not re.fullmatch(r"\d+\.\d+\.\d+", info["version"]):
        errors.append("Invalid package version")
    if f"## {info['version']} " not in files["CHANGELOG.md"].decode("utf-8"):
        errors.append("CHANGELOG.md has no entry for the current package version")
    if info.get("license") != "PolyForm-Noncommercial-1.0.0":
        errors.append("Package license must match retained upstream terms")
    if files["LICENSE"] != files["LICENSE-upstream.txt"]:
        errors.append("Retained PolyForm license text differs between LICENSE files")
    if info.get("tests_record") not in files:
        errors.append("Package validation record is missing")
    interface_data = yaml.safe_load(files["agents/openai.yaml"])
    interface = interface_data["interface"]
    if not 25 <= len(interface["short_description"]) <= 64:
        errors.append("UI short_description must contain 25–64 characters")
    if f"${name}" not in interface["default_prompt"]:
        errors.append("UI default_prompt must mention the skill")
    for key in ("icon_small", "icon_large"):
        if interface.get(key, "").removeprefix("./") not in files:
            errors.append(f"Missing UI icon: {key}")
    if set(interface_data.get("policy", {})) - {"allow_implicit_invocation"}:
        errors.append("Unexpected product policy field")
    checked = 0
    for relative, data in files.items():
        suffix = Path(relative).suffix
        if suffix == ".json":
            json.loads(data)
        elif suffix in {".yaml", ".yml"}:
            yaml.safe_load(data)
        elif suffix == ".py":
            ast.parse(data.decode("utf-8"), filename=relative, feature_version=(3, 9))
        elif suffix == ".md":
            # Ignore code fences, whose example links/paths are not Markdown links.
            prose = re.sub(r"```.*?```", "", data.decode("utf-8"), flags=re.S)
            for target in re.findall(r"\[[^\]\n]*\]\(([^)\n]+)\)", prose):
                parsed = urlsplit(target.strip("<>"))
                if parsed.scheme or parsed.netloc or not parsed.path:
                    continue
                candidate = ((root / relative).parent / unquote(parsed.path)).resolve()
                candidate.relative_to(root)
                checked += 1
                if not candidate.exists():
                    errors.append(f"Broken link in {relative}: {target}")
    example = validate_payload(json.loads(files["assets/example-document.json"]))
    load_profile(root / DEFAULT_PROFILES.relative_to(ROOT), example["profile"])
    return checked, errors


def main() -> int:
    try:
        links, errors = validate()
        for error in errors:
            print(error, file=sys.stderr)
        if errors:
            return 1
        print(f"Project valid; {links} local Markdown links checked; examples and metadata valid.")
        return 0
    except (OSError, ValueError, KeyError, TypeError, SyntaxError, yaml.YAMLError) as exc:
        print(f"Validation failed: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
