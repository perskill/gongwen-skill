#!/usr/bin/env python3
"""Build reproducible skill/source ZIPs from an explicit set of project files."""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import sys
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo


ROOT = Path(__file__).resolve().parents[1]
ROOT_FILES = (
    "SKILL.md", "README.md", "LICENSE", "LICENSE-upstream.txt", "NOTICE",
    "CHANGELOG.md", "CONTRIBUTING.md", "SECURITY.md", "requirements.txt",
    "requirements-dev.txt",
)
CONTENT_DIRS = ("agents", "assets", "references", "scripts", "tests", "docs")
SUFFIXES = {".md", ".py", ".json", ".yaml", ".yml", ".svg", ".txt", ".sha256"}
SKIP_DIRS = {"__pycache__", ".venv", ".git", "dist", "build"}
SOURCE_ONLY = {"requirements-dev.txt", "scripts/build_release.py",
               "scripts/validate_project.py", "tests/test_build_release.py"}


def collect_files(root: Path, source_archive: bool = False) -> dict[str, bytes]:
    """Exclude local environments, generated documents, hidden files and symlinks."""
    root = root.resolve()
    candidates = [root / name for name in ROOT_FILES]
    for directory in CONTENT_DIRS + ((".github",) if source_archive else ()):
        base = root / directory
        if not base.is_dir() or base.is_symlink():
            raise ValueError(f"Missing directory or unsupported symlink: {directory}")
        candidates.extend(sorted(base.rglob("*")))
    if source_archive:
        candidates.extend(root / name for name in (".gitignore", ".gitattributes"))
    files = {}
    for path in candidates:
        relative = path.relative_to(root)
        if not source_archive and relative.as_posix() in SOURCE_ONLY:
            continue
        if any(part in SKIP_DIRS for part in relative.parts):
            continue
        if any(part.startswith(".") and part not in {".github", ".gitignore", ".gitattributes"}
               for part in relative.parts):
            continue
        if path.is_symlink():
            raise ValueError(f"Symlinks cannot enter releases: {relative}")
        if path.is_dir():
            continue
        if relative.as_posix() not in ROOT_FILES and path.name not in {".gitignore", ".gitattributes"} and path.suffix not in SUFFIXES:
            continue
        path.resolve().relative_to(root)
        data = path.read_bytes()
        # Normalize repository text even when a checkout used Windows CRLF.
        data.decode("utf-8")
        files[relative.as_posix()] = data.replace(b"\r\n", b"\n")
    return dict(sorted(files.items()))


def write_zip(output: Path, prefix: str, files: dict[str, bytes]) -> None:
    manifest = "".join(f"{hashlib.sha256(data).hexdigest()}  {name}\n"
                       for name, data in sorted(files.items())).encode("utf-8")
    with ZipFile(output, "w", compression=ZIP_DEFLATED, compresslevel=9) as archive:
        for name, data in sorted({**files, "MANIFEST.sha256": manifest}.items()):
            entry = ZipInfo(f"{prefix}/{name}", date_time=(2026, 1, 1, 0, 0, 0))
            entry.create_system = 3
            entry.external_attr = 0o100644 << 16
            entry.compress_type = ZIP_DEFLATED
            archive.writestr(entry, data, compresslevel=9)
    digest = hashlib.sha256(output.read_bytes()).hexdigest()
    with output.with_suffix(output.suffix + ".sha256").open("w", encoding="utf-8", newline="\n") as stream:
        stream.write(f"{digest}  {output.name}\n")


def build(root: Path, output_dir: Path, source_archive: bool = False) -> list[Path]:
    root = root.resolve()
    files = collect_files(root)
    metadata = json.loads(files["assets/package-info.json"])
    name, version = metadata["name"], metadata["version"]
    if not isinstance(name, str) or not re.fullmatch(r"[a-z0-9]+(?:-[a-z0-9]+)*", name) or len(name) > 64:
        raise ValueError("Invalid skill name")
    if not isinstance(version, str) or not re.fullmatch(r"\d+\.\d+\.\d+", version):
        raise ValueError("Release version must use MAJOR.MINOR.PATCH")
    packages = [(name, f"{name}-v{version}.zip", files)]
    if source_archive:
        packages.append(("gongwen-skill", f"gongwen-skill-github-v{version}.zip", collect_files(root, True)))
    output_dir.mkdir(parents=True, exist_ok=True)
    outputs = []
    for prefix, filename, contents in packages:
        target = output_dir / filename
        write_zip(target, prefix, contents)
        outputs.append(target)
    return outputs


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, default=ROOT / "dist")
    parser.add_argument("--source-archive", action="store_true", help="Also include the GitHub project archive")
    args = parser.parse_args()
    try:
        for output in build(ROOT, args.output_dir, args.source_archive):
            print(output.name)
        return 0
    except (OSError, ValueError, KeyError, TypeError) as exc:
        print(f"Build failed: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
