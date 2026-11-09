#!/usr/bin/env python3
"""Portable, explicit-content DOCX construction and verification (Python 3.9+).

No network or platform SDK is used. Only build requires python-docx.
Verification covers the main document body, not policy correctness or visual fit.
"""
from __future__ import annotations

import argparse
import difflib
import importlib.util
import json
import math
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
from typing import Any, Dict, List, Optional, Tuple
import xml.etree.ElementTree as ET
import zipfile

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{" + NS["w"] + "}"
DEFAULT_PROFILES = Path(__file__).resolve().parent.parent / "assets" / "docx_profiles.json"
ROLES = {"body", "recipient", "signature", "date", "attachment", "note"}
MAX_XML_BYTES = 64 * 1024 * 1024


class DocumentError(Exception):
    """Input, unsupported format, dependency, or output failure."""


def _keys(value: Dict[str, Any], allowed: set, label: str) -> None:
    unknown = set(value) - allowed
    if unknown:
        raise DocumentError("%s 含未支持字段：%s" % (label, ", ".join(sorted(unknown))))


def _text(value: Any, label: str, empty: bool = False) -> str:
    if not isinstance(value, str):
        raise DocumentError(label + " 必须是字符串；数字也须显式写成字符串")
    value = value.replace("\r\n", "\n").replace("\r", "\n")
    if not empty and not value.strip():
        raise DocumentError(label + " 不得为空；待补信息请明确标记")
    if any((ord(c) < 32 and c not in "\n\t") or 0xD800 <= ord(c) <= 0xDFFF or ord(c) in (0xFFFE, 0xFFFF) for c in value):
        raise DocumentError(label + " 含 XML 不支持的控制字符")
    return value


def validate_payload(value: Any) -> Dict[str, Any]:
    if not isinstance(value, dict):
        raise DocumentError("源文件必须是 JSON 对象")
    _keys(value, {"schema_version", "profile", "blocks"}, "源文件")
    if type(value.get("schema_version")) is not int or value["schema_version"] != 1:
        raise DocumentError("schema_version 必须为整数 1")
    profile = value.get("profile", "internal-report")
    if not isinstance(profile, str) or not profile.strip():
        raise DocumentError("profile 必须是非空字符串")
    blocks = value.get("blocks")
    if not isinstance(blocks, list) or not blocks:
        raise DocumentError("blocks 必须为非空数组")
    cleaned = []
    for i, block in enumerate(blocks):
        label = "blocks[%d]" % i
        if not isinstance(block, dict):
            raise DocumentError(label + " 必须是对象")
        kind = block.get("type")
        if kind in ("title", "heading", "paragraph"):
            allowed = {"type", "text"} | ({"level"} if kind == "heading" else {"role"} if kind == "paragraph" else set())
            _keys(block, allowed, label)
            item = {"type": kind, "text": _text(block.get("text"), label + ".text")}
            if kind == "heading":
                if type(block.get("level")) is not int or block["level"] not in (1, 2, 3):
                    raise DocumentError(label + ".level 必须为整数 1、2 或 3")
                item["level"] = block["level"]
            if kind == "paragraph":
                role = block.get("role", "body")
                if not isinstance(role, str) or role not in ROLES:
                    raise DocumentError(label + ".role 不受支持")
                item["role"] = role
            cleaned.append(item)
        elif kind == "table":
            _keys(block, {"type", "rows", "header_rows"}, label)
            rows = block.get("rows")
            if not isinstance(rows, list) or not rows or not isinstance(rows[0], list) or not rows[0]:
                raise DocumentError(label + ".rows 必须为非空二维数组")
            width = len(rows[0])
            out_rows = []
            for ri, row in enumerate(rows):
                if not isinstance(row, list) or len(row) != width:
                    raise DocumentError(label + ".rows 行宽不一致")
                out_rows.append([_text(cell, "%s.rows[%d][%d]" % (label, ri, ci), True) for ci, cell in enumerate(row)])
            header_rows = block.get("header_rows", 1)
            if type(header_rows) is not int or header_rows not in (0, 1):
                raise DocumentError(label + ".header_rows 必须为整数 0 或 1")
            cleaned.append({"type": kind, "rows": out_rows, "header_rows": header_rows})
        elif kind == "page_break":
            _keys(block, {"type"}, label)
            cleaned.append({"type": kind})
        else:
            raise DocumentError(label + ".type 不受支持；不能静默忽略内容")
    if not any(b["type"] != "page_break" for b in cleaned):
        raise DocumentError("blocks 不能仅包含分页符")
    return {"schema_version": 1, "profile": profile, "blocks": cleaned}


def load_payload(path: Path) -> Dict[str, Any]:
    try:
        return validate_payload(json.loads(path.read_text(encoding="utf-8-sig")))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise DocumentError("无法读取源 JSON：%s" % exc) from exc


def expected_blocks(payload: Dict[str, Any]) -> List[Dict[str, Any]]:
    result = []
    for block in payload["blocks"]:
        if block["type"] == "page_break":
            result.append({"type": "page_break"})
        elif block["type"] == "table":
            result.append({"type": "table", "rows": block["rows"]})
        else:
            result.append({"type": "paragraph", "text": block["text"]})
    return result


def _xml(data: bytes, part: str) -> ET.Element:
    if len(data) > MAX_XML_BYTES or b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
        raise DocumentError("XML 过大或包含不支持的 DTD/实体声明：" + part)
    try:
        return ET.fromstring(data)
    except ET.ParseError as exc:
        raise DocumentError("XML 损坏：%s：%s" % (part, exc)) from exc


def _open_parts(path: Path) -> Dict[str, bytes]:
    if not path.is_file():
        raise DocumentError("DOCX 文件不存在：" + str(path))
    try:
        with zipfile.ZipFile(path) as archive:
            infos = archive.infolist()
            if len(infos) > 10000 or sum(i.file_size for i in infos) > 200 * 1024 * 1024:
                raise DocumentError("DOCX 展开体积或条目数超出本工具范围")
            names = {i.filename for i in infos}
            required = {"[Content_Types].xml", "_rels/.rels", "word/document.xml"}
            if not required <= names:
                raise DocumentError("文件不是完整 DOCX；缺少 " + ", ".join(sorted(required - names)))
            if len(names) != len(infos):
                raise DocumentError("DOCX 含重复 ZIP 条目，拒绝歧义内容")
            damaged = archive.testzip()
            if damaged is not None:
                raise DocumentError("DOCX ZIP 条目损坏：" + damaged)
            parts = {}
            for name in names:
                if name in required or name == "word/styles.xml" or name == "word/settings.xml" or re.fullmatch(r"word/footer\d+\.xml", name):
                    info = archive.getinfo(name)
                    if info.file_size > MAX_XML_BYTES:
                        raise DocumentError("XML 过大：" + name)
                    parts[name] = archive.read(name)
            content_types = _xml(parts["[Content_Types].xml"], "[Content_Types].xml")
            main_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
            if not any(n.get("PartName") == "/word/document.xml" and n.get("ContentType") == main_type for n in content_types):
                raise DocumentError("DOCX 未声明受支持的 Word 正文内容类型")
            relationships = _xml(parts["_rels/.rels"], "_rels/.rels")
            if not any(n.get("Type", "").endswith("/officeDocument") and n.get("Target", "").lstrip("/") == "word/document.xml" and n.get("TargetMode") != "External" for n in relationships):
                raise DocumentError("DOCX 根关系未指向本地 word/document.xml")
            return parts
    except (OSError, zipfile.BadZipFile, RuntimeError, EOFError) as exc:
        raise DocumentError("无法解析 DOCX：%s" % exc) from exc


def _paragraph_tokens(paragraph: ET.Element) -> List[Tuple[str, str]]:
    """Read supported visible run leaves; never silently drop an unknown leaf."""
    tokens = []
    page_before = paragraph.find("w:pPr/w:pageBreakBefore", NS)
    if page_before is not None and page_before.get(W + "val", "true") not in ("0", "false", "off"):
        tokens.append(("page_break", ""))
    markers = {W + name for name in ("bookmarkStart", "bookmarkEnd", "proofErr", "commentRangeStart", "commentRangeEnd", "permStart", "permEnd")}

    def walk(container: ET.Element) -> None:
        for node in container:
            if node.tag in (W + "pPr", W + "rPr") or node.tag in markers:
                continue
            if node.tag in (W + "r", W + "hyperlink"):
                walk(node)
            elif node.tag == W + "t" and len(node) == 0:
                tokens.append(("text", node.text or ""))
            elif node.tag == W + "tab" and len(node) == 0:
                tokens.append(("text", "\t"))
            elif node.tag == W + "cr" and len(node) == 0:
                tokens.append(("text", "\n"))
            elif node.tag == W + "br" and len(node) == 0:
                kind = node.get(W + "type", "textWrapping")
                if kind == "page":
                    tokens.append(("page_break", ""))
                elif kind == "textWrapping":
                    tokens.append(("text", "\n"))
                else:
                    raise DocumentError("不支持的显式换栏/换行类型：" + kind)
            elif node.tag == W + "lastRenderedPageBreak" and len(node) == 0:
                # Renderer metadata records automatic pagination, not a manual break.
                continue
            else:
                raise DocumentError("段落含本工具未支持的可见或未知内容节点：" + node.tag.rsplit("}", 1)[-1])

    walk(paragraph)
    return tokens


def _paragraph_blocks(paragraph: ET.Element) -> List[Dict[str, Any]]:
    blocks = []
    chunks = []
    for kind, value in _paragraph_tokens(paragraph):
        if kind == "page_break":
            if chunks:
                blocks.append({"type": "paragraph", "text": "".join(chunks)})
                chunks = []
            blocks.append({"type": "page_break"})
        elif value:
            chunks.append(value)
    if chunks:
        blocks.append({"type": "paragraph", "text": "".join(chunks)})
    return blocks


def _paragraph_text(paragraph: ET.Element) -> str:
    tokens = _paragraph_tokens(paragraph)
    if any(kind == "page_break" for kind, _ in tokens):
        raise DocumentError("表格单元格含显式分页符，本工具不支持自动比对")
    return "".join(value for _, value in tokens)


def extract_document(path: Path) -> Dict[str, Any]:
    parts = _open_parts(path)
    root = _xml(parts["word/document.xml"], "word/document.xml")
    body = root.find("w:body", NS)
    if body is None:
        raise DocumentError("DOCX 缺少正文 body")
    unsupported = {W + t for t in ("ins", "del", "moveFrom", "moveTo", "txbxContent", "altChunk", "drawing", "pict", "object", "fldSimple", "fldChar", "instrText", "footnoteReference", "endnoteReference", "sdt", "cellIns", "cellDel", "cellMerge")}
    findings = sorted({node.tag.rsplit("}", 1)[-1] for node in body.iter() if node.tag in unsupported or (node.tag.startswith(W) and node.tag.endswith("Change")) or node.tag.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/math}")})
    if findings:
        raise DocumentError("正文含本工具不支持的结构（请先用原生工具处理，再核对）：" + ", ".join(findings))
    blocks = []
    for child in body:
        if child.tag == W + "p":
            blocks.extend(_paragraph_blocks(child))
        elif child.tag == W + "tbl":
            rows = []
            for row in child.findall("w:tr", NS):
                if row.find("w:trPr/w:gridBefore", NS) is not None or row.find("w:trPr/w:gridAfter", NS) is not None:
                    raise DocumentError("表格含省略网格单元格，本工具不支持自动比对")
                cells = []
                for cell in row.findall("w:tc", NS):
                    if cell.find(".//w:tbl", NS) is not None:
                        raise DocumentError("正文含嵌套表格，本工具不支持")
                    if cell.find("w:tcPr/w:gridSpan", NS) is not None or cell.find("w:tcPr/w:vMerge", NS) is not None:
                        raise DocumentError("正文含合并单元格，本工具不支持自动比对")
                    if any(n.tag not in (W + "tcPr", W + "p") for n in cell):
                        raise DocumentError("表格含不支持的单元格内容")
                    cells.append("\n".join(_paragraph_text(p) for p in cell.findall("w:p", NS)))
                if not cells:
                    raise DocumentError("表格含空行结构")
                rows.append(cells)
            if not rows or any(len(r) != len(rows[0]) for r in rows):
                raise DocumentError("表格为空或维度不一致")
            blocks.append({"type": "table", "rows": rows})
        elif child.tag != W + "sectPr":
            raise DocumentError("正文含不支持的顶层结构：" + child.tag.rsplit("}", 1)[-1])
    if not any(block["type"] != "page_break" for block in blocks):
        raise DocumentError("DOCX 正文未提取到内容，不能判为完整")
    page_breaks = sum(block["type"] == "page_break" for block in blocks)
    return {"blocks": blocks, "page_breaks": page_breaks, "scope": "main_document_body", "parts": parts}


def verify_document(payload: Dict[str, Any], path: Path) -> Dict[str, Any]:
    actual = extract_document(path)
    expected = expected_blocks(payload)
    encode = lambda b: json.dumps(b, ensure_ascii=False, sort_keys=True)
    a, b = [encode(x) for x in expected], [encode(x) for x in actual["blocks"]]
    differences = []
    for tag, i, j, k, l in difflib.SequenceMatcher(None, a, b, autojunk=False).get_opcodes():
        if tag != "equal":
            differences.append({"kind": tag, "expected_range": [i, j], "actual_range": [k, l], "expected_preview": a[i:j][:2], "actual_preview": b[k:l][:2]})
    expected_breaks = sum(x["type"] == "page_break" for x in payload["blocks"])
    expected_positions = [i for i, block in enumerate(expected) if block["type"] == "page_break"]
    actual_positions = [i for i, block in enumerate(actual["blocks"]) if block["type"] == "page_break"]
    return {
        "passed": a == b and expected_breaks == actual["page_breaks"],
        "scope": "main_document_body",
        "expected_blocks": len(a), "actual_blocks": len(b),
        "content_equal": a == b,
        "page_breaks_equal": expected_positions == actual_positions,
        "expected_page_breaks": expected_breaks, "actual_page_breaks": actual["page_breaks"],
        "expected_page_break_positions": expected_positions, "actual_page_break_positions": actual_positions,
        "difference_count": len(differences), "differences": differences[:20],
        "comparison": "complete_ordered_blocks_including_manual_page_break_positions_and_table_cells; only CRLF/CR source line endings normalized",
        "visual_review": "not_performed",
    }


def validate_profile(profile: Any) -> Dict[str, Any]:
    """Validate every field consumed by construction and its success response."""
    try:
        if not isinstance(profile, dict):
            raise DocumentError("版式配置须为对象")
        required = {"label", "notes", "page", "body", "title", "headings", "table", "page_numbers"}
        _keys(profile, required, "版式配置")
        if not required <= set(profile):
            raise DocumentError("版式配置缺少字段：" + ", ".join(sorted(required - set(profile))))
        _text(profile["label"], "label")
        _text(profile["notes"], "notes", empty=True)
        _keys(profile["headings"], {"1", "2", "3"}, "headings")
        groups = {
            "page": ({"width_mm", "height_mm", "top_mm", "bottom_mm", "left_mm", "right_mm", "footer_mm"}, set()),
            "body": ({"font", "size_pt", "line_pt", "first_line_chars"}, {"bold"}),
            "title": ({"font", "size_pt", "line_pt"}, {"bold"}),
            "table": ({"font", "size_pt", "line_pt"}, {"bold"}),
            "page_numbers": ({"font", "size_pt", "position"}, set()),
        }
        values = dict(profile)
        for level in ("1", "2", "3"):
            values["heading" + level] = profile["headings"][level]
            groups["heading" + level] = ({"font", "size_pt"}, {"bold", "line_pt"})
        for group, (mandatory, optional) in groups.items():
            item = values[group]
            if not isinstance(item, dict):
                raise DocumentError(group + " 须为对象")
            _keys(item, mandatory | optional, group)
            if not mandatory <= set(item):
                raise DocumentError(group + " 缺少字段：" + ", ".join(sorted(mandatory - set(item))))
            if "font" in item:
                _text(item["font"], group + ".font")
            if "bold" in item and type(item["bold"]) is not bool:
                raise DocumentError(group + ".bold 须为布尔值")
            for key, value in item.items():
                if key in ("font", "bold", "position"):
                    continue
                low, high = (5, 72) if key == "size_pt" else (5, 144) if key == "line_pt" else (0, 10) if key == "first_line_chars" else (0, 1000)
                if type(value) not in (int, float) or not math.isfinite(value) or not low <= value <= high:
                    raise DocumentError("无效版式数值：%s.%s" % (group, key))
        if profile["page_numbers"]["position"] not in ("center", "outside", "none"):
            raise DocumentError("页码位置须为 center、outside 或 none")
        page = profile["page"]
        if page["left_mm"] + page["right_mm"] >= page["width_mm"] or page["top_mm"] + page["bottom_mm"] >= page["height_mm"]:
            raise DocumentError("页边距使版心宽或高小于等于零")
        return profile
    except (ValueError, KeyError, TypeError, AttributeError) as exc:
        raise DocumentError("无效版式配置：%s" % exc) from exc


def load_profile(path: Path, name: str) -> Dict[str, Any]:
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
        if not isinstance(data, dict) or type(data.get("schema_version")) is not int or data["schema_version"] != 1 or not isinstance(data.get("profiles"), dict) or name not in data["profiles"]:
            raise DocumentError("配置版本不符或不存在版式：" + name)
        return validate_profile(data["profiles"][name])
    except (OSError, UnicodeError, ValueError) as exc:
        raise DocumentError("无法读取版式配置：%s" % exc) from exc


def _style_name(block: Dict[str, Any]) -> str:
    if block["type"] == "title":
        return "GongwenTitle"
    if block["type"] == "heading":
        return "Heading %d" % block["level"]
    return "Gongwen" + block.get("role", "body").capitalize()


def build_document(payload: Dict[str, Any], output: Path, profile: Dict[str, Any], overwrite: bool = False) -> Dict[str, Any]:
    payload = validate_payload(payload)
    profile = validate_profile(profile)
    output_path_text = str(output.absolute())
    try:
        from docx import Document
        from docx.enum.style import WD_STYLE_TYPE
        from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        from docx.shared import Mm, Pt, RGBColor
    except ImportError as exc:
        raise DocumentError("build 需要 python-docx；其他检查命令可用。未自动安装依赖。") from exc
    if output.exists() and not overwrite:
        raise DocumentError("目标已存在；保留原文件，确认替换时加 --overwrite")
    if output.suffix.lower() != ".docx":
        raise DocumentError("输出文件须以 .docx 结尾")
    output.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    doc.core_properties.author = ""
    doc.core_properties.last_modified_by = ""
    doc.core_properties.title = next((b["text"] for b in payload["blocks"] if b["type"] == "title"), "")
    section = doc.sections[0]
    page = profile["page"]
    for attr, key in (("page_width", "width_mm"), ("page_height", "height_mm"), ("top_margin", "top_mm"), ("bottom_margin", "bottom_mm"), ("left_margin", "left_mm"), ("right_margin", "right_mm"), ("footer_distance", "footer_mm")):
        setattr(section, attr, Mm(page[key]))

    def font_for(style, config):
        style.font.name = config["font"]
        style.font.size = Pt(config["size_pt"])
        style.font.bold = bool(config.get("bold", False))
        style.font.italic = False
        style.font.color.rgb = RGBColor(0, 0, 0)
        rpr = style.element.get_or_add_rPr()
        fonts = rpr.find(qn("w:rFonts"))
        if fonts is None:
            fonts = OxmlElement("w:rFonts")
            rpr.insert(0, fonts)
        for attr in ("ascii", "hAnsi", "eastAsia", "cs"):
            fonts.set(qn("w:" + attr), config["font"])
        for attr in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            fonts.attrib.pop(qn("w:" + attr), None)

    def make_style(name, config, indent=0, keep_next=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY):
        style = doc.styles[name] if name in doc.styles else doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        font_for(style, config)
        pf = style.paragraph_format
        pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
        pf.line_spacing = Pt(config.get("line_pt", profile["body"]["line_pt"]))
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.first_line_indent = Pt(config["size_pt"] * indent)
        pf.keep_with_next = keep_next
        pf.widow_control = True
        pf.alignment = align
        ind = style.element.get_or_add_pPr().find(qn("w:ind"))
        if ind is not None:
            ind.set(qn("w:firstLineChars"), str(round(indent * 100)))
        return style

    make_style("Normal", profile["body"], profile["body"]["first_line_chars"])
    make_style("GongwenBody", profile["body"], profile["body"]["first_line_chars"])
    make_style("GongwenTitle", profile["title"], keep_next=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.styles["GongwenTitle"].paragraph_format.space_after = Pt(profile["body"]["line_pt"] / 2)
    for level in (1, 2, 3):
        make_style("Heading %d" % level, profile["headings"][str(level)], indent=profile["body"]["first_line_chars"], keep_next=True)
    for role in ROLES - {"body"}:
        align = WD_ALIGN_PARAGRAPH.RIGHT if role in ("signature", "date") else WD_ALIGN_PARAGRAPH.LEFT
        make_style("Gongwen" + role.capitalize(), profile["body"], align=align)
    make_style("GongwenTable", profile["table"], align=WD_ALIGN_PARAGRAPH.LEFT)
    header_config = dict(profile["table"], bold=True)
    make_style("GongwenTableHeader", header_config, align=WD_ALIGN_PARAGRAPH.LEFT)
    for block in payload["blocks"]:
        kind = block["type"]
        if kind == "page_break":
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        elif kind == "table":
            rows = block["rows"]
            table = doc.add_table(rows=len(rows), cols=len(rows[0]))
            table.style = "Table Grid"
            table.autofit = False
            cell_width = Mm((page["width_mm"] - page["left_mm"] - page["right_mm"]) / len(rows[0]))
            for column in table.columns:
                column.width = cell_width
            for ri, row in enumerate(rows):
                trpr = table.rows[ri]._tr.get_or_add_trPr()
                trpr.append(OxmlElement("w:cantSplit"))
                if ri < block["header_rows"]:
                    trpr.append(OxmlElement("w:tblHeader"))
                for ci, value in enumerate(row):
                    cell = table.cell(ri, ci)
                    cell.width = cell_width
                    paragraph = cell.paragraphs[0]
                    paragraph.style = "GongwenTableHeader" if ri < block["header_rows"] else "GongwenTable"
                    paragraph.add_run(value)
        else:
            doc.add_paragraph(block["text"], style=_style_name(block))

    page_numbers = profile["page_numbers"]
    if page_numbers["position"] != "none":
        make_style("GongwenPageNumber", dict(page_numbers, line_pt=page_numbers["size_pt"] + 2), align=WD_ALIGN_PARAGRAPH.CENTER)
        doc.settings.odd_and_even_pages_header_footer = page_numbers["position"] == "outside"
        footers = [(section.footer, WD_ALIGN_PARAGRAPH.RIGHT), (section.even_page_footer, WD_ALIGN_PARAGRAPH.LEFT)] if page_numbers["position"] == "outside" else [(section.footer, WD_ALIGN_PARAGRAPH.CENTER)]
        for footer, align in footers:
            p = footer.paragraphs[0]
            p.style = "GongwenPageNumber"
            p.alignment = align
            if align == WD_ALIGN_PARAGRAPH.RIGHT:
                p.paragraph_format.right_indent = Pt(page_numbers["size_pt"])
            elif align == WD_ALIGN_PARAGRAPH.LEFT:
                p.paragraph_format.left_indent = Pt(page_numbers["size_pt"])
            p.add_run("— ")
            run = p.add_run()
            for typ in ("begin", "separate", "end"):
                if typ == "separate":
                    instruction = OxmlElement("w:instrText")
                    instruction.set(qn("xml:space"), "preserve")
                    instruction.text = " PAGE "
                    run._r.append(instruction)
                field = OxmlElement("w:fldChar")
                field.set(qn("w:fldCharType"), typ)
                run._r.append(field)
                if typ == "separate":
                    text = OxmlElement("w:t")
                    text.text = "1"
                    run._r.append(text)
            p.add_run(" —")
    update = OxmlElement("w:updateFields")
    update.set(qn("w:val"), "true")
    doc.settings.element.append(update)
    fd, temporary = tempfile.mkstemp(prefix=".gongwen-", suffix=".docx", dir=str(output.parent))
    os.close(fd)
    temp_path = Path(temporary)
    try:
        doc.save(temp_path)
        verification = verify_document(payload, temp_path)
        if not verification["passed"]:
            raise DocumentError("生成后全文核对失败，目标文件未替换：" + json.dumps(verification, ensure_ascii=False))
        layout = verify_layout(temp_path, profile)
        if not layout["passed"]:
            raise DocumentError("生成后版式参数核对失败，目标文件未替换：" + json.dumps(layout, ensure_ascii=False))
        if output.exists() and not overwrite:
            raise DocumentError("输出期间目标已存在；未覆盖")
        result = {"ok": True, "operation": "build", "docx_path": output_path_text, "profile": payload["profile"], "content": verification, "layout": layout, "visual_review": "not_performed", "warnings": [profile["notes"], "字体名称已设置；字体实际可用性、分页和页面效果须渲染检查。"]}
        # Nothing that consumes configuration or serializes the success response
        # may fail after the final atomic replacement.
        json.dumps(result, ensure_ascii=False).encode("utf-8")
        os.replace(temp_path, output)
        return result
    except BaseException:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass
        raise


def verify_layout(path: Path, profile: Dict[str, Any]) -> Dict[str, Any]:
    parts = _open_parts(path)
    root = _xml(parts["word/document.xml"], "word/document.xml")
    styles = _xml(parts.get("word/styles.xml", b"<styles/>"), "word/styles.xml")
    issues = []
    sections = root.findall(".//w:sectPr", NS)
    if len(sections) != 1:
        issues.append("本脚本基础版式应只有一个节")
    else:
        sec = sections[0]
        for key, xml_path, attr in (("width_mm", "w:pgSz", "w"), ("height_mm", "w:pgSz", "h"), ("top_mm", "w:pgMar", "top"), ("bottom_mm", "w:pgMar", "bottom"), ("left_mm", "w:pgMar", "left"), ("right_mm", "w:pgMar", "right"), ("footer_mm", "w:pgMar", "footer")):
            node = sec.find(xml_path, NS)
            value = node.get(W + attr) if node is not None else None
            expected = round(profile["page"][key] / 25.4 * 1440)
            if value is None or abs(int(value) - expected) > 1:
                issues.append("页面参数不符：" + key)
    style_map = {s.get(W + "styleId"): s for s in styles.findall("w:style", NS)}
    expected_styles = {"GongwenBody": profile["body"], "GongwenTitle": profile["title"], "GongwenTable": profile["table"]}
    for level in (1, 2, 3):
        expected_styles["Heading%d" % level] = profile["headings"][str(level)]
    for sid, config in expected_styles.items():
        style = style_map.get(sid)
        font = style.find("w:rPr/w:rFonts", NS) if style is not None else None
        size = style.find("w:rPr/w:sz", NS) if style is not None else None
        if font is None or font.get(W + "eastAsia") != config["font"]:
            issues.append("中文字体配置不符：" + sid)
        if size is None or size.get(W + "val") != str(round(config["size_pt"] * 2)):
            issues.append("字号配置不符：" + sid)
    body_style = style_map.get("GongwenBody")
    spacing = body_style.find("w:pPr/w:spacing", NS) if body_style is not None else None
    indent = body_style.find("w:pPr/w:ind", NS) if body_style is not None else None
    if spacing is None or spacing.get(W + "lineRule") != "exact" or spacing.get(W + "line") != str(round(profile["body"]["line_pt"] * 20)):
        issues.append("正文固定行距配置不符")
    if indent is None or indent.get(W + "firstLineChars") != str(round(profile["body"]["first_line_chars"] * 100)):
        issues.append("正文段首字符缩进配置不符")
    settings = _xml(parts.get("word/settings.xml", b"<settings/>"), "word/settings.xml")
    outside = profile["page_numbers"]["position"] == "outside"
    if outside and settings.find("w:evenAndOddHeaders", NS) is None:
        issues.append("未启用奇偶页不同页脚")
    footer_alignments = []
    for name, data in parts.items():
        if re.fullmatch(r"word/footer\d+\.xml", name):
            footer = _xml(data, name)
            for paragraph in footer.findall("w:p", NS):
                if any("PAGE" in (n.text or "") for n in paragraph.iter(W + "instrText")):
                    alignment = paragraph.find("w:pPr/w:jc", NS)
                    footer_alignments.append(alignment.get(W + "val") if alignment is not None else None)
    if outside and not {"left", "right"} <= set(footer_alignments):
        issues.append("未同时找到左右对齐的页码")
    if profile["page_numbers"]["position"] == "center" and "center" not in footer_alignments:
        issues.append("未找到居中页码")
    return {"checked": True, "passed": not issues, "scope": "generated_base_profile_parameters", "issues": issues, "limitations": "不认证全部有效样式、字体安装、视觉分页、红头印章或 GB/T 9704 全项合规"}


def doctor() -> Dict[str, Any]:
    return {"ok": True, "operation": "doctor", "python": sys.version.split()[0], "python_executable": sys.executable, "capabilities": {"content_verification": True, "docx_build": importlib.util.find_spec("docx") is not None, "soffice": shutil.which("soffice") or shutil.which("libreoffice"), "pdftoppm": shutil.which("pdftoppm"), "fontconfig": shutil.which("fc-match")}, "network_used": False, "dependencies_installed": False, "font_availability": "not_verified"}


def render_document(path: Path, output_dir: Path, png: bool = False, soffice: Optional[str] = None, pdftoppm: Optional[str] = None) -> Dict[str, Any]:
    _open_parts(path)
    binary = soffice or shutil.which("soffice") or shutil.which("libreoffice")
    if not binary:
        raise DocumentError("未找到 LibreOffice/soffice；可指定 --soffice 路径。未安装任何软件。")
    raster = pdftoppm or shutil.which("pdftoppm")
    if png and not raster:
        raise DocumentError("--png 需要 pdftoppm；可指定 --pdftoppm 路径，或先只生成 PDF")
    if output_dir.exists() and any(output_dir.iterdir()):
        raise DocumentError("渲染目录须为不存在或空目录，以免混入旧页面")
    output_dir.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix="gongwen-render-") as temporary:
        temp = Path(temporary)
        target_dir = temp / "pdf"
        target_dir.mkdir()
        profile_uri = (temp / "lo-profile").resolve().as_uri()
        command = [str(binary), "-env:UserInstallation=" + profile_uri, "--headless", "--convert-to", "pdf", "--outdir", str(target_dir), str(path.resolve())]
        try:
            completed = subprocess.run(command, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120)
        except (OSError, subprocess.TimeoutExpired) as exc:
            raise DocumentError("渲染工具执行失败：%s" % exc) from exc
        generated = target_dir / (path.stem + ".pdf")
        valid_pdf = generated.is_file() and generated.stat().st_size >= 5
        if valid_pdf:
            with generated.open("rb") as handle:
                valid_pdf = handle.read(5) == b"%PDF-"
        if completed.returncode != 0 or not valid_pdf:
            raise DocumentError("LibreOffice 未生成有效 PDF：" + (completed.stderr + completed.stdout)[-1000:])
        final_pdf = output_dir / (path.stem + ".pdf")
        shutil.copy2(generated, final_pdf)
    pages = []
    if png:
        prefix = output_dir / "page"
        try:
            completed = subprocess.run([str(raster), "-r", "120", "-png", str(final_pdf), str(prefix)], capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=120)
        except (OSError, subprocess.TimeoutExpired) as exc:
            raise DocumentError("PDF 已生成，但 PNG 渲染失败：%s" % exc) from exc
        pages = sorted(output_dir.glob("page-*.png"), key=lambda p: int(p.stem.rsplit("-", 1)[1]))
        if completed.returncode != 0 or not pages:
            raise DocumentError("PDF 已生成，但未生成 PNG 页面：" + completed.stderr[-1000:])
    return {"ok": True, "operation": "render", "pdf_path": str(final_pdf.resolve()), "pages": [str(p.resolve()) for p in pages], "visual_review": "required", "fonts": "may_be_substituted; inspect_each_page", "network_used": False}


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("doctor", help="检查本机能力，不安装软件")
    build = sub.add_parser("build", help="从完整结构化 JSON 新建 DOCX")
    build.add_argument("--source", type=Path, required=True)
    build.add_argument("--output", type=Path, required=True)
    build.add_argument("--profiles", type=Path, default=DEFAULT_PROFILES)
    build.add_argument("--profile")
    build.add_argument("--overwrite", action="store_true")
    verify = sub.add_parser("verify", help="与完整 JSON 比较正文、表格及显式分页符")
    verify.add_argument("--source", type=Path, required=True)
    verify.add_argument("--docx", type=Path, required=True)
    verify.add_argument("--check-layout", action="store_true", help="仅用于本脚本基础版式，不用于认证单位模板")
    verify.add_argument("--profiles", type=Path, default=DEFAULT_PROFILES)
    verify.add_argument("--profile", help="与 build --profile 相同的覆盖值")
    extract = sub.add_parser("extract", help="提取受支持的正文结构，不接受修订/文本框等")
    extract.add_argument("--docx", type=Path, required=True)
    render = sub.add_parser("render", help="使用本机 LibreOffice 生成 PDF，可附 PNG")
    render.add_argument("--docx", type=Path, required=True)
    render.add_argument("--output-dir", type=Path, required=True)
    render.add_argument("--png", action="store_true")
    render.add_argument("--soffice")
    render.add_argument("--pdftoppm")
    args = parser.parse_args(argv)
    try:
        if args.command == "doctor":
            result = doctor()
        elif args.command == "build":
            payload = load_payload(args.source)
            if args.profile:
                payload["profile"] = args.profile
            result = build_document(payload, args.output, load_profile(args.profiles, payload["profile"]), args.overwrite)
        elif args.command == "verify":
            payload = load_payload(args.source)
            if args.profile:
                payload["profile"] = args.profile
            content = verify_document(payload, args.docx)
            layout = verify_layout(args.docx, load_profile(args.profiles, payload["profile"])) if args.check_layout else {"checked": False}
            result = {"ok": content["passed"] and layout.get("passed", True), "operation": "verify", "docx_path": str(args.docx.resolve()), "content": content, "layout": layout, "visual_review": "not_performed"}
        elif args.command == "extract":
            result = extract_document(args.docx)
            result.pop("parts")
            result.update({"ok": True, "operation": "extract"})
        else:
            result = render_document(args.docx, args.output_dir, args.png, args.soffice, args.pdftoppm)
        print(json.dumps(result, ensure_ascii=False, indent=2))
        return 0 if result["ok"] else 1
    except (DocumentError, OSError, ValueError, KeyError, TypeError, zipfile.BadZipFile) as exc:
        print(json.dumps({"ok": False, "operation": args.command, "error": {"type": type(exc).__name__, "message": str(exc)}}, ensure_ascii=False, indent=2))
        return 2


if __name__ == "__main__":
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    sys.exit(main())
