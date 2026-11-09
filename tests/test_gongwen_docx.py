"""Regression tests for known conversion failures and portable DOCX output.

Run from the skill root: python -m unittest discover -s tests -v
The content verification tests need only the Python standard library.
Build tests skip explicitly when python-docx is unavailable.
"""
from __future__ import annotations

import copy
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location("gongwen_docx", ROOT / "scripts" / "gongwen_docx.py")
MODULE = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(MODULE)
W = MODULE.W


def document_payload(blocks):
    return MODULE.validate_payload({"schema_version": 1, "profile": "internal-report", "blocks": blocks})


def package(path, blocks, extra=None):
    root = ET.Element(W + "document")
    body = ET.SubElement(root, W + "body")
    for block in blocks:
        if block["type"] == "table":
            tbl = ET.SubElement(body, W + "tbl")
            for row in block["rows"]:
                tr = ET.SubElement(tbl, W + "tr")
                for text in row:
                    tc = ET.SubElement(tr, W + "tc")
                    p = ET.SubElement(tc, W + "p")
                    ET.SubElement(ET.SubElement(p, W + "r"), W + "t").text = text
        else:
            p = ET.SubElement(body, W + "p")
            r = ET.SubElement(p, W + "r")
            if block["type"] == "page_break":
                ET.SubElement(r, W + "br", {W + "type": "page"})
            else:
                ET.SubElement(r, W + "t").text = block["text"]
    if extra is not None:
        extra(body)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
        archive.writestr("_rels/.rels", '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
        archive.writestr("word/document.xml", ET.tostring(root, encoding="utf-8", xml_declaration=True))


class ContentTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.path = Path(self.temporary.name) / "检查样稿.docx"

    def check(self, source, actual):
        payload = document_payload(source)
        package(self.path, actual)
        return MODULE.verify_document(payload, self.path)

    def test_entities_and_short_text_preserved(self):
        blocks = [{"type": "paragraph", "text": "A&B < 3，金额：0元"}, {"type": "paragraph", "role": "signature", "text": "办公室"}]
        self.assertTrue(self.check(blocks, blocks)["passed"])

    def test_prefix_only_truncation_fails(self):
        text = "这是长度超过十六个字符的一段完整文字，后半段包含重要条件及不予认定的结论。"
        result = self.check([{"type": "paragraph", "text": text}], [{"type": "paragraph", "text": text[:16]}])
        self.assertFalse(result["passed"])

    def test_deleted_short_line_and_signoff_fail(self):
        blocks = [{"type": "paragraph", "text": "这是主要说明。"}, {"type": "paragraph", "text": "金额：0元"}, {"type": "paragraph", "role": "signature", "text": "办公室"}]
        self.assertFalse(self.check(blocks, blocks[:1])["passed"])

    def test_reordered_paragraphs_fail(self):
        blocks = [{"type": "paragraph", "text": "先完成核对。"}, {"type": "paragraph", "text": "再提请审议。"}]
        self.assertFalse(self.check(blocks, list(reversed(blocks)))["passed"])

    def test_duplicate_paragraph_deleted_fails(self):
        block = {"type": "paragraph", "text": "同一事项需按规定办理。"}
        self.assertFalse(self.check([block, block], [block])["passed"])

    def test_inserted_content_fails(self):
        block = {"type": "paragraph", "text": "尚未开展。"}
        self.assertFalse(self.check([block], [block, {"type": "paragraph", "text": "已完成培训。"}])["passed"])

    def test_negation_decimal_and_fullwidth_change_fail(self):
        original = [{"type": "paragraph", "text": "不得转出；金额1.50万元，编号１２。"}]
        for text in ("得转出；金额1.50万元，编号１２。", "不得转出；金额150万元，编号１２。", "不得转出；金额1.50万元，编号12。"):
            with self.subTest(text=text):
                self.assertFalse(self.check(original, [{"type": "paragraph", "text": text}])["passed"])

    def test_table_order_and_dimensions_fail(self):
        expected = [{"type": "table", "rows": [["事项", "金额"], ["甲", "0"], ["乙", "2"]]}]
        variants = [
            [{"type": "table", "rows": [["事项", "金额"], ["乙", "2"], ["甲", "0"]]}],
            [{"type": "table", "rows": [["事项金额"], ["甲0"], ["乙2"]]}],
        ]
        for actual in variants:
            self.assertFalse(self.check(expected, actual)["passed"])

    def test_table_empty_cells_and_entities_pass(self):
        blocks = [{"type": "table", "rows": [["A&B", ""], ["0", "不适用"]]}]
        self.assertTrue(self.check(blocks, blocks)["passed"])

    def test_explicit_page_break_checked(self):
        blocks = [{"type": "paragraph", "text": "正文"}, {"type": "page_break"}, {"type": "paragraph", "text": "附件"}]
        self.assertTrue(self.check(blocks, blocks)["passed"])
        self.assertFalse(self.check(blocks, [blocks[0], blocks[2]])["passed"])

    def test_page_break_moved_with_same_count_fails(self):
        blocks = [{"type": "paragraph", "text": "正文第一段。"}, {"type": "page_break"}, {"type": "paragraph", "text": "附件甲：核对清单。"}, {"type": "paragraph", "text": "附件乙：统计表。"}]
        actual = [blocks[0], blocks[2], blocks[1], blocks[3]]
        result = self.check(blocks, actual)
        self.assertEqual(result["expected_page_breaks"], result["actual_page_breaks"])
        self.assertFalse(result["passed"])
        self.assertFalse(result["page_breaks_equal"])
        self.assertNotEqual(result["expected_page_break_positions"], result["actual_page_break_positions"])

    def test_visible_run_symbols_and_unknown_leaves_rejected(self):
        blocks = [{"type": "paragraph", "text": "应纳税额为0元。"}]
        for tag in ("sym", "noBreakHyphen", "softHyphen", "ptab", "separator", "continuationSeparator", "unknownVisibleNode"):
            with self.subTest(tag=tag):
                def inject(body):
                    paragraph = body.find(W + "p")
                    run = ET.SubElement(paragraph, W + "r")
                    ET.SubElement(run, W + tag, {W + "font": "Wingdings", W + "char": "F0FC"})
                package(self.path, blocks, inject)
                with self.assertRaisesRegex(MODULE.DocumentError, "未支持"):
                    MODULE.verify_document(document_payload(blocks), self.path)

    def test_manual_page_break_inside_table_rejected(self):
        blocks = [{"type": "table", "rows": [["甲", "乙"]]}]
        def inject(body):
            run = body.find(".//" + W + "tc/" + W + "p/" + W + "r")
            ET.SubElement(run, W + "br", {W + "type": "page"})
        package(self.path, blocks, inject)
        with self.assertRaisesRegex(MODULE.DocumentError, "显式分页符"):
            MODULE.verify_document(document_payload(blocks), self.path)

    def test_explicit_page_break_before_property_checked(self):
        blocks = [{"type": "paragraph", "text": "正文"}, {"type": "paragraph", "text": "附件"}]
        def inject(body):
            paragraph = body.findall(W + "p")[1]
            ET.SubElement(ET.SubElement(paragraph, W + "pPr"), W + "pageBreakBefore")
        package(self.path, blocks, inject)
        result = MODULE.verify_document(document_payload(blocks), self.path)
        self.assertFalse(result["passed"])
        self.assertEqual(result["actual_page_break_positions"], [1])

    def test_missing_and_corrupt_files_error(self):
        payload = document_payload([{"type": "paragraph", "text": "正文"}])
        with self.assertRaises(MODULE.DocumentError):
            MODULE.verify_document(payload, self.path)
        self.path.write_bytes(b"not a document")
        with self.assertRaises(MODULE.DocumentError):
            MODULE.verify_document(payload, self.path)

    def test_empty_document_rejected(self):
        package(self.path, [])
        with self.assertRaises(MODULE.DocumentError):
            MODULE.extract_document(self.path)

    def test_revision_textbox_and_content_control_rejected(self):
        blocks = [{"type": "paragraph", "text": "正文"}]
        for tag in ("ins", "del", "rPrChange", "pPrChange", "txbxContent", "sdt", "altChunk", "drawing", "fldSimple", "object"):
            with self.subTest(tag=tag):
                package(self.path, blocks, lambda b: ET.SubElement(b, W + tag))
                with self.assertRaises(MODULE.DocumentError):
                    MODULE.extract_document(self.path)

    def test_nested_and_merged_tables_rejected(self):
        blocks = [{"type": "table", "rows": [["甲", "乙"]]}]
        package(self.path, blocks, lambda b: ET.SubElement(b.find(".//" + W + "tc"), W + "tbl"))
        with self.assertRaises(MODULE.DocumentError):
            MODULE.extract_document(self.path)
        def merge(body):
            cell = body.find(".//" + W + "tc")
            ET.SubElement(ET.SubElement(cell, W + "tcPr"), W + "gridSpan", {W + "val": "2"})
        package(self.path, blocks, merge)
        with self.assertRaises(MODULE.DocumentError):
            MODULE.extract_document(self.path)

    def test_unsupported_input_not_silently_ignored(self):
        for block in ({"type": "image", "text": "test"}, {"type": "paragraph", "text": "正文", "hidden": "遗漏"}, {"type": "table", "rows": [["甲"], ["乙", "丙"]]}, {"type": "paragraph", "text": 0}):
            with self.subTest(block=block):
                with self.assertRaises(MODULE.DocumentError):
                    document_payload([block])

    def test_cli_failure_is_nonzero_json(self):
        source = Path(self.temporary.name) / "source.json"
        source.write_text(json.dumps({"schema_version": 1, "blocks": [{"type": "paragraph", "text": "正文"}]}), encoding="utf-8")
        completed = subprocess.run([sys.executable, str(ROOT / "scripts" / "gongwen_docx.py"), "verify", "--source", str(source), "--docx", str(self.path)], capture_output=True, text=True, encoding="utf-8")
        self.assertEqual(completed.returncode, 2)
        self.assertFalse(json.loads(completed.stdout)["ok"])

    def test_render_zero_exit_without_pdf_is_failure(self):
        package(self.path, [{"type": "paragraph", "text": "正文"}])
        fake = subprocess.CompletedProcess(args=[], returncode=0, stdout="reported success", stderr="")
        with patch.object(MODULE.subprocess, "run", return_value=fake):
            with self.assertRaises(MODULE.DocumentError):
                MODULE.render_document(self.path, Path(self.temporary.name) / "render", soffice="fake-soffice")


@unittest.skipUnless(importlib.util.find_spec("docx"), "python-docx unavailable; standard-library verification tests still run")
class BuildTests(unittest.TestCase):
    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.directory = Path(self.temporary.name)

    def payload(self, profile="internal-report"):
        value = {"schema_version": 1, "profile": profile, "blocks": [
            {"type": "title", "text": "虚构教学样稿"},
            {"type": "paragraph", "role": "recipient", "text": "有关单位："},
            {"type": "heading", "level": 1, "text": "一、工作条件"},
            {"type": "paragraph", "text": "A&B 尚未完成；金额1.50万元，编号１２。\n第二行\t保留制表符。"},
            {"type": "paragraph", "text": "重复段落"}, {"type": "paragraph", "text": "重复段落"},
            {"type": "table", "rows": [["项目", "结果"], ["A&B", "0"], ["空位", ""], ["多行", "第一行\n第二行"]]},
            {"type": "page_break"}, {"type": "paragraph", "role": "attachment", "text": "附件：核对清单"},
            {"type": "paragraph", "role": "signature", "text": "办公室"},
            {"type": "paragraph", "role": "date", "text": "2026年10月29日"},
        ]}
        return MODULE.validate_payload(value)

    def test_build_all_profiles_and_verify(self):
        for name in ("internal-report", "research-report", "government-body"):
            with self.subTest(profile=name):
                payload = self.payload(name)
                output = self.directory / (name + ".docx")
                profile = MODULE.load_profile(MODULE.DEFAULT_PROFILES, name)
                result = MODULE.build_document(payload, output, profile)
                self.assertTrue(result["ok"])
                self.assertTrue(MODULE.verify_document(payload, output)["passed"])
                self.assertTrue(MODULE.verify_layout(output, profile)["passed"])
                self.assertEqual(result["visual_review"], "not_performed")

    def test_existing_file_preserved_without_overwrite(self):
        output = self.directory / "existing.docx"
        output.write_bytes(b"keep this original")
        payload = self.payload()
        profile = MODULE.load_profile(MODULE.DEFAULT_PROFILES, payload["profile"])
        with self.assertRaises(MODULE.DocumentError):
            MODULE.build_document(payload, output, profile)
        self.assertEqual(output.read_bytes(), b"keep this original")

    def test_failed_postcheck_never_replaces_original(self):
        output = self.directory / "existing.docx"
        output.write_bytes(b"keep this original")
        payload = self.payload()
        profile = MODULE.load_profile(MODULE.DEFAULT_PROFILES, payload["profile"])
        with patch.object(MODULE, "verify_document", return_value={"passed": False}):
            with self.assertRaises(MODULE.DocumentError):
                MODULE.build_document(payload, output, profile, overwrite=True)
        self.assertEqual(output.read_bytes(), b"keep this original")
        self.assertEqual(list(self.directory.glob(".gongwen-*")), [])

    def test_missing_python_docx_is_clear(self):
        import builtins
        original = builtins.__import__
        def missing(name, *args, **kwargs):
            if name == "docx" or name.startswith("docx."):
                raise ImportError("test missing")
            return original(name, *args, **kwargs)
        payload = self.payload()
        profile = MODULE.load_profile(MODULE.DEFAULT_PROFILES, payload["profile"])
        with patch("builtins.__import__", side_effect=missing):
            with self.assertRaisesRegex(MODULE.DocumentError, "需要 python-docx"):
                MODULE.build_document(payload, self.directory / "new.docx", profile)

    def test_missing_profile_notes_never_replaces_original(self):
        config = json.loads(MODULE.DEFAULT_PROFILES.read_text(encoding="utf-8"))
        del config["profiles"]["internal-report"]["notes"]
        config_path = self.directory / "missing-notes.json"
        config_path.write_text(json.dumps(config, ensure_ascii=False), encoding="utf-8")
        with self.assertRaisesRegex(MODULE.DocumentError, "notes"):
            MODULE.load_profile(config_path, "internal-report")
        output = self.directory / "existing.docx"
        output.write_bytes(b"keep this original")
        source = self.directory / "source.json"
        source.write_text(json.dumps(self.payload(), ensure_ascii=False), encoding="utf-8")
        completed = subprocess.run([sys.executable, str(ROOT / "scripts" / "gongwen_docx.py"), "build", "--source", str(source), "--output", str(output), "--profiles", str(config_path), "--overwrite"], capture_output=True, text=True, encoding="utf-8")
        self.assertEqual(completed.returncode, 2)
        self.assertFalse(json.loads(completed.stdout)["ok"])
        self.assertEqual(output.read_bytes(), b"keep this original")
        with self.assertRaisesRegex(MODULE.DocumentError, "notes"):
            MODULE.build_document(self.payload(), output, config["profiles"]["internal-report"], overwrite=True)
        self.assertEqual(output.read_bytes(), b"keep this original")

    def test_success_response_serialized_before_replacement(self):
        output = self.directory / "existing.docx"
        output.write_bytes(b"keep this original")
        payload = self.payload()
        profile = MODULE.load_profile(MODULE.DEFAULT_PROFILES, payload["profile"])
        original = json.dumps
        def fail_success_response(value, *args, **kwargs):
            if isinstance(value, dict) and value.get("operation") == "build":
                raise TypeError("simulated response serialization failure")
            return original(value, *args, **kwargs)
        with patch.object(MODULE.json, "dumps", side_effect=fail_success_response):
            with self.assertRaisesRegex(TypeError, "serialization failure"):
                MODULE.build_document(payload, output, profile, overwrite=True)
        self.assertEqual(output.read_bytes(), b"keep this original")
        self.assertEqual(list(self.directory.glob(".gongwen-*")), [])


if __name__ == "__main__":
    unittest.main()
