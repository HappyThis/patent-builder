#!/usr/bin/env python3
"""Utilities for reading and editing DOCX OOXML without external dependencies."""

from __future__ import annotations

import json
import re
import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Tuple
import xml.etree.ElementTree as ET

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
XML_NS = "http://www.w3.org/XML/1998/namespace"
NS = {"w": W_NS, "m": M_NS}

ET.register_namespace("w", W_NS)
ET.register_namespace("m", M_NS)

PLACEHOLDER_PATTERNS = [
    re.compile(r"\{\{\s*([^{}]+?)\s*\}\}"),
    re.compile(r"【\s*([^【】]+?)\s*】"),
    re.compile(r"\[\[\s*([^\[\]]+?)\s*\]\]"),
]
LABEL_COLON_PATTERN = re.compile(r"^([^：:]{1,48})[：:]$")
SIMPLE_SUBSCRIPT_MAP = {
    "0": "₀",
    "1": "₁",
    "2": "₂",
    "3": "₃",
    "4": "₄",
    "5": "₅",
    "6": "₆",
    "7": "₇",
    "8": "₈",
    "9": "₉",
    "a": "ₐ",
    "e": "ₑ",
    "h": "ₕ",
    "i": "ᵢ",
    "j": "ⱼ",
    "k": "ₖ",
    "l": "ₗ",
    "m": "ₘ",
    "n": "ₙ",
    "o": "ₒ",
    "p": "ₚ",
    "r": "ᵣ",
    "s": "ₛ",
    "t": "ₜ",
    "u": "ᵤ",
    "v": "ᵥ",
    "x": "ₓ",
}


@dataclass
class StyleInfo:
    style_id: str
    name: str
    style_type: str
    based_on: Optional[str]
    is_default: bool
    font_ascii: Optional[str]
    font_east_asia: Optional[str]
    size_pt: Optional[float]
    bold: Optional[bool]
    italic: Optional[bool]
    line_spacing: Optional[float]
    space_before_pt: Optional[float]
    space_after_pt: Optional[float]
    first_line_indent_pt: Optional[float]


def qn(tag: str) -> str:
    prefix, local = tag.split(":", 1)
    namespace_map = {"w": W_NS, "m": M_NS}
    if prefix not in namespace_map:
        raise ValueError(f"Unsupported namespace prefix: {prefix}")
    return f"{{{namespace_map[prefix]}}}{local}"


def now_utc_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def twips_to_pt(value: Optional[str]) -> Optional[float]:
    if value in (None, ""):
        return None
    try:
        return round(int(value) / 20.0, 2)
    except ValueError:
        return None


def half_points_to_pt(value: Optional[str]) -> Optional[float]:
    if value in (None, ""):
        return None
    try:
        return round(int(value) / 2.0, 2)
    except ValueError:
        return None


def normalize_key(text: str) -> str:
    lowered = text.strip().lower()
    return re.sub(r"[^0-9a-z\u4e00-\u9fff]+", "", lowered)


def normalize_equation_text(text: str) -> str:
    def repl(match: re.Match[str]) -> str:
        base = match.group(1)
        suffix = match.group(2)
        converted = []
        for ch in suffix:
            sub = SIMPLE_SUBSCRIPT_MAP.get(ch.lower())
            if sub is None:
                return match.group(0)
            converted.append(sub)
        return base + "".join(converted)

    return re.sub(r"\b([A-Za-z])_([A-Za-z0-9]+)\b", repl, text)


def paragraph_text(paragraph: ET.Element) -> str:
    parts = []
    for node in paragraph.findall(".//w:t", NS):
        if node.text:
            parts.append(node.text)
    return "".join(parts)


def extract_placeholder_keys(text: str) -> List[str]:
    keys: List[str] = []
    for pattern in PLACEHOLDER_PATTERNS:
        for match in pattern.finditer(text):
            key = match.group(1).strip()
            if key:
                keys.append(key)
    label_match = LABEL_COLON_PATTERN.match(text.strip())
    if label_match:
        label = label_match.group(1).strip()
        if 1 <= len(label) <= 40:
            keys.append(label)
    return keys


def extract_strict_placeholder_keys(text: str) -> List[str]:
    keys: List[str] = []
    for pattern in PLACEHOLDER_PATTERNS:
        for match in pattern.finditer(text):
            key = match.group(1).strip()
            if key:
                keys.append(key)
    return keys


def read_docx_members(docx_path: Path) -> Dict[str, bytes]:
    with zipfile.ZipFile(docx_path, "r") as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def read_xml_member(docx_path: Path, member: str) -> ET.Element:
    with zipfile.ZipFile(docx_path, "r") as zf:
        data = zf.read(member)
    return ET.fromstring(data)


def write_docx_members(src_docx: Path, dst_docx: Path, updates: Dict[str, bytes]) -> None:
    with zipfile.ZipFile(src_docx, "r") as zin:
        with zipfile.ZipFile(dst_docx, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = updates.get(item.filename)
                if data is None:
                    data = zin.read(item.filename)
                zout.writestr(item, data)


def body_element(document_root: ET.Element) -> ET.Element:
    body = document_root.find("w:body", NS)
    if body is None:
        raise ValueError("Invalid DOCX: missing w:body")
    return body


def body_paragraphs(document_root: ET.Element) -> List[ET.Element]:
    return body_element(document_root).findall("w:p", NS)


def all_paragraphs(document_root: ET.Element) -> List[ET.Element]:
    return document_root.findall(".//w:p", NS)


def paragraph_style_id(paragraph: ET.Element) -> Optional[str]:
    ppr = paragraph.find("w:pPr", NS)
    if ppr is None:
        return None
    pstyle = ppr.find("w:pStyle", NS)
    if pstyle is None:
        return None
    return pstyle.get(qn("w:val"))


def _parse_rpr(style: ET.Element) -> Tuple[Optional[str], Optional[str], Optional[float], Optional[bool], Optional[bool]]:
    rpr = style.find("w:rPr", NS)
    if rpr is None:
        return None, None, None, None, None
    rfonts = rpr.find("w:rFonts", NS)
    ascii_font = rfonts.get(qn("w:ascii")) if rfonts is not None else None
    east_asia_font = rfonts.get(qn("w:eastAsia")) if rfonts is not None else None
    size_node = rpr.find("w:sz", NS)
    size_pt = half_points_to_pt(size_node.get(qn("w:val")) if size_node is not None else None)
    bold = rpr.find("w:b", NS) is not None
    italic = rpr.find("w:i", NS) is not None
    return ascii_font, east_asia_font, size_pt, bold, italic


def _parse_ppr(style: ET.Element) -> Tuple[Optional[float], Optional[float], Optional[float], Optional[float]]:
    ppr = style.find("w:pPr", NS)
    if ppr is None:
        return None, None, None, None
    spacing = ppr.find("w:spacing", NS)
    line_spacing = twips_to_pt(spacing.get(qn("w:line")) if spacing is not None else None)
    space_before = twips_to_pt(spacing.get(qn("w:before")) if spacing is not None else None)
    space_after = twips_to_pt(spacing.get(qn("w:after")) if spacing is not None else None)
    ind = ppr.find("w:ind", NS)
    first_line = twips_to_pt(ind.get(qn("w:firstLine")) if ind is not None else None)
    return line_spacing, space_before, space_after, first_line


def parse_styles(styles_root: Optional[ET.Element]) -> Tuple[Dict[str, StyleInfo], Dict[str, str]]:
    if styles_root is None:
        return {}, {}

    styles: Dict[str, StyleInfo] = {}
    name_to_id: Dict[str, str] = {}

    for style in styles_root.findall("w:style", NS):
        style_id = style.get(qn("w:styleId"))
        style_type = style.get(qn("w:type"), "")
        if not style_id or style_type != "paragraph":
            continue

        name_node = style.find("w:name", NS)
        name = name_node.get(qn("w:val")) if name_node is not None else style_id
        based_on_node = style.find("w:basedOn", NS)
        based_on = based_on_node.get(qn("w:val")) if based_on_node is not None else None
        is_default = style.get(qn("w:default")) == "1"

        ascii_font, east_asia_font, size_pt, bold, italic = _parse_rpr(style)
        line_spacing, space_before, space_after, first_line = _parse_ppr(style)

        info = StyleInfo(
            style_id=style_id,
            name=name,
            style_type=style_type,
            based_on=based_on,
            is_default=is_default,
            font_ascii=ascii_font,
            font_east_asia=east_asia_font,
            size_pt=size_pt,
            bold=bold,
            italic=italic,
            line_spacing=line_spacing,
            space_before_pt=space_before,
            space_after_pt=space_after,
            first_line_indent_pt=first_line,
        )
        styles[style_id] = info
        if name:
            name_to_id[name] = style_id

    return styles, name_to_id


def parse_layout_rules(document_root: ET.Element) -> Dict[str, Optional[float]]:
    body = body_element(document_root)
    sect = body.find("w:sectPr", NS)
    if sect is None:
        candidate = body.findall("w:p/w:pPr/w:sectPr", NS)
        sect = candidate[-1] if candidate else None

    if sect is None:
        return {}

    rules: Dict[str, Optional[float]] = {}
    pg_mar = sect.find("w:pgMar", NS)
    if pg_mar is not None:
        rules = {
            "margin_top_pt": twips_to_pt(pg_mar.get(qn("w:top"))),
            "margin_bottom_pt": twips_to_pt(pg_mar.get(qn("w:bottom"))),
            "margin_left_pt": twips_to_pt(pg_mar.get(qn("w:left"))),
            "margin_right_pt": twips_to_pt(pg_mar.get(qn("w:right"))),
            "header_pt": twips_to_pt(pg_mar.get(qn("w:header"))),
            "footer_pt": twips_to_pt(pg_mar.get(qn("w:footer"))),
        }

    pg_sz = sect.find("w:pgSz", NS)
    if pg_sz is not None:
        rules.update(
            {
                "page_width_pt": twips_to_pt(pg_sz.get(qn("w:w"))),
                "page_height_pt": twips_to_pt(pg_sz.get(qn("w:h"))),
            }
        )

    return rules


def style_info_to_dict(style: StyleInfo) -> Dict[str, object]:
    return {
        "style_id": style.style_id,
        "name": style.name,
        "style_type": style.style_type,
        "based_on": style.based_on,
        "is_default": style.is_default,
        "font_ascii": style.font_ascii,
        "font_east_asia": style.font_east_asia,
        "size_pt": style.size_pt,
        "bold": style.bold,
        "italic": style.italic,
        "line_spacing_pt": style.line_spacing,
        "space_before_pt": style.space_before_pt,
        "space_after_pt": style.space_after_pt,
        "first_line_indent_pt": style.first_line_indent_pt,
    }


def choose_default_paragraph_style(style_map: Dict[str, StyleInfo], name_to_id: Dict[str, str]) -> Optional[str]:
    for name in ("缺省文本", "Normal", "正文", "Body Text"):
        if name in name_to_id:
            return name_to_id[name]
    for info in style_map.values():
        if info.is_default:
            return info.style_id
    return None


def choose_heading_style(style_map: Dict[str, StyleInfo], name_to_id: Dict[str, str], level: int = 1) -> Optional[str]:
    candidates = [
        f"Heading {level}",
        f"标题 {level}",
        f"Heading{level}",
        f"标题{level}",
    ]
    for name in candidates:
        if name in name_to_id:
            return name_to_id[name]

    pattern = re.compile(rf"(heading|标题)\s*{level}", re.IGNORECASE)
    for info in style_map.values():
        if pattern.search(info.name):
            return info.style_id
    return None


def create_paragraph(text: str, style_id: Optional[str] = None) -> ET.Element:
    paragraph = ET.Element(qn("w:p"))
    if style_id:
        ppr = ET.SubElement(paragraph, qn("w:pPr"))
        pstyle = ET.SubElement(ppr, qn("w:pStyle"))
        pstyle.set(qn("w:val"), style_id)

    lines = text.split("\n") if text else [""]
    for idx, line in enumerate(lines):
        run = ET.SubElement(paragraph, qn("w:r"))
        text_node = ET.SubElement(run, qn("w:t"))
        if line.startswith(" ") or line.endswith(" "):
            text_node.set(f"{{{XML_NS}}}space", "preserve")
        text_node.text = line
        if idx < len(lines) - 1:
            ET.SubElement(run, qn("w:br"))
    return paragraph


def create_equation_paragraph(text: str, style_id: Optional[str] = None, center: bool = True) -> ET.Element:
    paragraph = ET.Element(qn("w:p"))
    ppr = None
    if style_id or center:
        ppr = ET.SubElement(paragraph, qn("w:pPr"))
    if style_id and ppr is not None:
        pstyle = ET.SubElement(ppr, qn("w:pStyle"))
        pstyle.set(qn("w:val"), style_id)
    if center and ppr is not None:
        jc = ET.SubElement(ppr, qn("w:jc"))
        jc.set(qn("w:val"), "center")

    omath_para = ET.SubElement(paragraph, qn("m:oMathPara"))
    omath = ET.SubElement(omath_para, qn("m:oMath"))
    mr = ET.SubElement(omath, qn("m:r"))
    mt = ET.SubElement(mr, qn("m:t"))
    normalized_text = normalize_equation_text(text)
    if text.startswith(" ") or text.endswith(" "):
        mt.set(f"{{{XML_NS}}}space", "preserve")
    mt.text = normalized_text
    return paragraph


def set_paragraph_text(paragraph: ET.Element, text: str) -> None:
    keep = []
    for child in list(paragraph):
        if child.tag == qn("w:pPr"):
            keep.append(child)
    for child in list(paragraph):
        paragraph.remove(child)
    for child in keep:
        paragraph.append(child)

    lines = text.split("\n") if text else [""]
    for idx, line in enumerate(lines):
        run = ET.SubElement(paragraph, qn("w:r"))
        text_node = ET.SubElement(run, qn("w:t"))
        if line.startswith(" ") or line.endswith(" "):
            text_node.set(f"{{{XML_NS}}}space", "preserve")
        text_node.text = line
        if idx < len(lines) - 1:
            ET.SubElement(run, qn("w:br"))


def append_before_sectpr(body: ET.Element, paragraph: ET.Element) -> None:
    children = list(body)
    for idx, child in enumerate(children):
        if child.tag == qn("w:sectPr"):
            body.insert(idx, paragraph)
            return
    body.append(paragraph)


def dump_json(path: Path, payload: Dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)


def style_name(style_id: Optional[str], style_map: Dict[str, StyleInfo]) -> Optional[str]:
    if not style_id:
        return None
    info = style_map.get(style_id)
    return info.name if info else None


def infer_heading_level(style_name_value: Optional[str], text: str) -> Optional[int]:
    if style_name_value:
        m = re.search(r"heading\s*(\d+)", style_name_value, re.IGNORECASE)
        if m:
            return int(m.group(1))
        m = re.search(r"标题\s*(\d+)", style_name_value)
        if m:
            return int(m.group(1))

    numbered = re.match(r"^(\d+(?:\.\d+)*)", text.strip())
    if numbered:
        return numbered.group(1).count(".") + 1

    if re.match(r"^[一二三四五六七八九十]+[、.].+", text.strip()):
        return 1
    return None


def is_heading_candidate(text: str, style_name_value: Optional[str]) -> bool:
    if not text.strip():
        return False

    canonical = (
        "技术领域",
        "背景技术",
        "发明内容",
        "附图说明",
        "具体实施方式",
        "技术方案",
        "关键点",
        "保护点",
        "有益效果",
    )
    if any(word in text for word in canonical):
        return True

    if style_name_value and re.search(r"(heading|标题)", style_name_value, re.IGNORECASE):
        return True

    short_text = len(text) <= 30
    if short_text and text.endswith(("：", ":")):
        return False

    return short_text and bool(re.match(r"^(\d+(?:\.\d+)*)\s*\S+", text))


def parse_markdown_sections(path: Path) -> List[Dict[str, object]]:
    sections: List[Dict[str, object]] = []
    current_title: Optional[str] = None
    bucket: List[str] = []

    def flush() -> None:
        nonlocal bucket, current_title
        if not current_title:
            bucket = []
            return
        paragraphs: List[str] = []
        paragraph_acc: List[str] = []
        for line in bucket:
            stripped = line.rstrip()
            if stripped == "":
                if paragraph_acc:
                    paragraphs.append(" ".join(part.strip() for part in paragraph_acc).strip())
                    paragraph_acc = []
                continue
            paragraph_acc.append(stripped)
        if paragraph_acc:
            paragraphs.append(" ".join(part.strip() for part in paragraph_acc).strip())
        sections.append({"title": current_title, "paragraphs": [p for p in paragraphs if p]})
        bucket = []

    with path.open("r", encoding="utf-8") as fh:
        for raw in fh:
            line = raw.rstrip("\n")
            if line.startswith("## "):
                flush()
                current_title = line[3:].strip()
                continue
            if line.startswith("### "):
                flush()
                current_title = line[4:].strip()
                continue
            if line.startswith("# "):
                continue
            bucket.append(line)

    flush()
    return [s for s in sections if s.get("title")]


def section_aliases(title: str) -> List[str]:
    aliases = [title]
    mapping = {
        "发明内容": ["技术方案描述", "技术方案"],
        "有益效果": ["优点", "效果"],
        "关键点和保护点": ["技术方案的关键点和保护点", "保护点"],
        "背景技术": ["现有方案及其缺陷"],
    }
    for key, vals in mapping.items():
        if key in title:
            aliases.extend(vals)
    return aliases


def estimate_match_score(target_norm: str, candidate_norm: str) -> int:
    if not target_norm or not candidate_norm:
        return 0
    if target_norm == candidate_norm:
        return 100
    if target_norm in candidate_norm or candidate_norm in target_norm:
        overlap = min(len(target_norm), len(candidate_norm))
        return 70 + min(overlap, 20)

    overlap = len(set(target_norm) & set(candidate_norm))
    return overlap
