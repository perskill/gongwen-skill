import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "check_draft.py"
spec = importlib.util.spec_from_file_location("check_draft", SCRIPT)
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)


class DraftDiagnosticTests(unittest.TestCase):
    def test_common_placeholder_forms_are_candidates(self):
        result = module.diagnose("【待补：数量】（待核：事项）〔待定〕 XXX工作，XX万元，___。TODO TBD")
        self.assertEqual(len(result["placeholder_candidates"]), 8)
        self.assertFalse(result["fact_check_performed"])

    def test_nested_marker_not_double_counted(self):
        self.assertEqual(len(module.diagnose("【待补：XX万元】")["placeholder_candidates"]), 1)

    def test_positions_preserve_chinese(self):
        result = module.diagnose("第一行\n金额XX万元")
        item = result["placeholder_candidates"][0]
        self.assertEqual((item["line"], item["column"]), (2, 3))

    def test_word_fragments_are_not_x_markers(self):
        self.assertFalse(module.diagnose("EXXON xxyy TODAY TBDcode")["placeholder_candidates"])

    def test_no_range_implied(self):
        self.assertEqual(module.diagnose("短文。")["length_issues"], [])

    def test_counts_and_explicit_range(self):
        result = module.diagnose("中文 12\n。", min_chars=8)
        self.assertEqual(result["non_whitespace_characters"], 5)
        self.assertEqual(result["han_characters"], 2)
        self.assertEqual(result["length_issues"][0]["kind"], "below_requested_minimum")

    def test_invalid_range_rejected(self):
        with self.assertRaises(ValueError):
            module.diagnose("正文", 20, 10)

    def test_repeated_paragraph_retained_as_candidate(self):
        para = "这是一段用于检查全文重复提示的较长虚构文字。"
        result = module.diagnose(para + "\n\n" + para)
        self.assertEqual(result["exact_repeated_paragraphs"][0]["same_as"], 1)

    def test_draft_placeholder_is_not_default_failure(self):
        with tempfile.TemporaryDirectory(prefix="draft test ") as directory:
            path = Path(directory) / "待补 草稿.txt"
            path.write_text("\ufeff开展了XXX工作，入库XX万元。", encoding="utf-8")
            run = subprocess.run([sys.executable, str(SCRIPT), str(path), "--json"], capture_output=True, text=True, encoding="utf-8")
            self.assertEqual(run.returncode, 0, run.stderr)
            self.assertEqual(len(json.loads(run.stdout)["placeholder_candidates"]), 2)
            run = subprocess.run([sys.executable, str(SCRIPT), str(path), "--fail-on-placeholders", "--json"], capture_output=True, text=True, encoding="utf-8")
            self.assertEqual(run.returncode, 1)

    def test_missing_input_reports_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            run = subprocess.run([sys.executable, str(SCRIPT), str(Path(directory) / "missing.txt"), "--json"], capture_output=True, text=True, encoding="utf-8")
            self.assertEqual(run.returncode, 2)
            self.assertFalse(json.loads(run.stdout)["diagnostic_completed"])


if __name__ == "__main__":
    unittest.main()
