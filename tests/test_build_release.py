"""Verify installable archives, exclusion rules, checksums and reproducibility."""
import hashlib
import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("build_release", ROOT / "scripts" / "build_release.py")
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)


class ReleaseTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory(prefix="发布 测试 ")
        self.addCleanup(self.temporary.cleanup)
        self.base = Path(self.temporary.name)

    def test_archives_are_installable_and_checksums_cover_every_file(self):
        info = json.loads((ROOT / "assets/package-info.json").read_text(encoding="utf-8"))
        outputs = MODULE.build(ROOT, self.base, source_archive=True)
        self.assertEqual(len(outputs), 2)
        for path in outputs:
            digest, filename = path.with_suffix(".zip.sha256").read_text(encoding="utf-8").split()
            self.assertEqual(filename, path.name)
            self.assertEqual(digest, hashlib.sha256(path.read_bytes()).hexdigest())
            with ZipFile(path) as archive:
                prefix = info["name"] if path.name.startswith(info["name"]) else "gongwen-skill"
                names = set(archive.namelist())
                self.assertIn(f"{prefix}/SKILL.md", names)
                self.assertIn(f"{prefix}/LICENSE-upstream.txt", names)
                self.assertIn(f"{prefix}/assets/docx_profiles.json", names)
                lines = archive.read(f"{prefix}/MANIFEST.sha256").decode("utf-8").splitlines()
                recorded = set()
                for line in lines:
                    checksum, relative = line.split("  ", 1)
                    name = f"{prefix}/{relative}"
                    recorded.add(name)
                    self.assertEqual(checksum, hashlib.sha256(archive.read(name)).hexdigest())
                self.assertEqual(recorded, names - {f"{prefix}/MANIFEST.sha256"})
                self.assertEqual(f"{prefix}/.github/workflows/ci.yml" in names, prefix == "gongwen-skill")
                self.assertFalse(any("__pycache__" in name or ".venv/" in name for name in names))

    def test_rebuild_is_identical(self):
        first = MODULE.build(ROOT, self.base / "one", True)
        second = MODULE.build(ROOT, self.base / "two", True)
        self.assertEqual([p.read_bytes() for p in first], [p.read_bytes() for p in second])

    def test_local_files_are_excluded_and_crlf_is_normalized(self):
        root = self.base / "project"
        root.mkdir()
        for name in MODULE.ROOT_FILES:
            (root / name).write_bytes(b"text\r\n")
        for name in MODULE.CONTENT_DIRS:
            (root / name).mkdir()
        for relative in ("scripts/__pycache__/local.py", "assets/.env", "assets/draft.docx", "secrets.txt", ".env"):
            path = root / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_bytes(b"must not ship")
        (root / "scripts" / "valid.py").write_bytes(b"print('ok')\r\n")
        files = MODULE.collect_files(root)
        self.assertEqual(set(files), (set(MODULE.ROOT_FILES) - MODULE.SOURCE_ONLY) | {"scripts/valid.py"})
        self.assertEqual(files["scripts/valid.py"], b"print('ok')\n")

    def test_installed_skill_runs_its_tests_from_a_different_directory(self):
        output = MODULE.build(ROOT, self.base / "packages")[0]
        with ZipFile(output) as archive:
            archive.extractall(self.base / "installed")
        installed = self.base / "installed" / "draft-official-documents"
        result = subprocess.run(
            [sys.executable] + (["-S"] if sys.flags.no_site else []) +
            ["-m", "unittest", "discover", "-s", str(installed / "tests")],
            cwd=self.base, capture_output=True, text=True, encoding="utf-8", timeout=60)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("Ran 37 tests", result.stderr)


if __name__ == "__main__":
    unittest.main()
