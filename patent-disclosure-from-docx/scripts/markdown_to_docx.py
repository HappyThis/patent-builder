#!/usr/bin/env python3
"""Convert Markdown into a better-formatted standalone DOCX."""

from __future__ import annotations

import argparse
import re
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

from ooxml_docx import qn

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC_NS = "http://purl.org/dc/elements/1.1/"
DCTERMS_NS = "http://purl.org/dc/terms/"
XSI_NS = "http://www.w3.org/2001/XMLSchema-instance"
XML_NS = "http://www.w3.org/XML/1998/namespace"

ET.register_namespace("w", W_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("cp", CP_NS)
ET.register_namespace("dc", DC_NS)
ET.register_namespace("dcterms", DCTERMS_NS)
ET.register_namespace("xsi", XSI_NS)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Convert a Markdown file to DOCX.")
    parser.add_argument("input", help="Input Markdown path")
    parser.add_argument("output", help="Output DOCX path")
    return parser.parse_args()


def flush_paragraph(lines: List[str], blocks: List[Dict[str, object]]) -> None:
    if not lines:
        return
    text = " ".join(part.strip() for part in lines if part.strip()).strip()
    if text:
        blocks.append({"type": "paragraph", "text": text})
    lines.clear()


def parse_markdown_blocks(markdown_text: str) -> List[Dict[str, object]]:
    blocks: List[Dict[str, object]] = []
    paragraph_lines: List[str] = []
    in_code_block = False
    code_lines: List[str] = []

    for raw_line in markdown_text.splitlines():
        line = raw_line.rstrip()

        if line.startswith("```"):
            flush_paragraph(paragraph_lines, blocks)
            if in_code_block:
                blocks.append({"type": "code", "text": "\n".join(code_lines)})
                code_lines = []
                in_code_block = False
            else:
                in_code_block = True
            continue

        if in_code_block:
            code_lines.append(raw_line.rstrip("\n"))
            continue

        heading_match = re.match(r"^(#{1,3})\s+(.+?)\s*$", line)
        if heading_match:
            flush_paragraph(paragraph_lines, blocks)
            blocks.append(
                {
                    "type": "heading",
                    "level": len(heading_match.group(1)),
                    "text": heading_match.group(2).strip(),
                }
            )
            continue

        if not line.strip():
            flush_paragraph(paragraph_lines, blocks)
            continue

        if re.match(r"^([-*]|\d+\.)\s+", line.strip()):
            flush_paragraph(paragraph_lines, blocks)
            blocks.append({"type": "list", "text": line.strip()})
            continue

        paragraph_lines.append(line)

    flush_paragraph(paragraph_lines, blocks)
    if code_lines:
        blocks.append({"type": "code", "text": "\n".join(code_lines)})

    return blocks


def add_section_properties(body: ET.Element) -> None:
    sect_pr = ET.SubElement(body, qn("w:sectPr"))

    pg_sz = ET.SubElement(sect_pr, qn("w:pgSz"))
    pg_sz.set(qn("w:w"), "11906")
    pg_sz.set(qn("w:h"), "16838")

    pg_mar = ET.SubElement(sect_pr, qn("w:pgMar"))
    pg_mar.set(qn("w:top"), "1440")
    pg_mar.set(qn("w:right"), "1800")
    pg_mar.set(qn("w:bottom"), "1440")
    pg_mar.set(qn("w:left"), "1800")
    pg_mar.set(qn("w:header"), "720")
    pg_mar.set(qn("w:footer"), "720")
    pg_mar.set(qn("w:gutter"), "0")


def add_run_properties(
    run: ET.Element,
    *,
    font_ascii: str,
    font_east_asia: str,
    size_half_points: int,
    bold: bool,
) -> None:
    rpr = ET.SubElement(run, qn("w:rPr"))
    rfonts = ET.SubElement(rpr, qn("w:rFonts"))
    rfonts.set(qn("w:ascii"), font_ascii)
    rfonts.set(qn("w:hAnsi"), font_ascii)
    rfonts.set(qn("w:eastAsia"), font_east_asia)

    lang = ET.SubElement(rpr, qn("w:lang"))
    lang.set(qn("w:val"), "zh-CN")
    lang.set(qn("w:eastAsia"), "zh-CN")

    size = ET.SubElement(rpr, qn("w:sz"))
    size.set(qn("w:val"), str(size_half_points))
    size_cs = ET.SubElement(rpr, qn("w:szCs"))
    size_cs.set(qn("w:val"), str(size_half_points))

    if bold:
        ET.SubElement(rpr, qn("w:b"))
        ET.SubElement(rpr, qn("w:bCs"))


def add_text_runs(
    paragraph: ET.Element,
    text: str,
    *,
    font_ascii: str,
    font_east_asia: str,
    size_half_points: int,
    bold: bool,
) -> None:
    lines = text.split("\n") if text else [""]
    for idx, line in enumerate(lines):
        run = ET.SubElement(paragraph, qn("w:r"))
        add_run_properties(
            run,
            font_ascii=font_ascii,
            font_east_asia=font_east_asia,
            size_half_points=size_half_points,
            bold=bold,
        )
        text_node = ET.SubElement(run, qn("w:t"))
        if line.startswith(" ") or line.endswith(" "):
            text_node.set(f"{{{XML_NS}}}space", "preserve")
        text_node.text = line
        if idx < len(lines) - 1:
            ET.SubElement(run, qn("w:br"))


def create_formatted_paragraph(
    text: str,
    *,
    align: str,
    font_ascii: str,
    font_east_asia: str,
    size_half_points: int,
    bold: bool,
    space_before: int,
    space_after: int,
    line: int,
    first_line_chars: Optional[int] = None,
    left: Optional[int] = None,
    hanging: Optional[int] = None,
) -> ET.Element:
    paragraph = ET.Element(qn("w:p"))
    ppr = ET.SubElement(paragraph, qn("w:pPr"))

    jc = ET.SubElement(ppr, qn("w:jc"))
    jc.set(qn("w:val"), align)

    spacing = ET.SubElement(ppr, qn("w:spacing"))
    spacing.set(qn("w:before"), str(space_before))
    spacing.set(qn("w:after"), str(space_after))
    spacing.set(qn("w:line"), str(line))
    spacing.set(qn("w:lineRule"), "auto")

    if first_line_chars is not None or left is not None or hanging is not None:
        ind = ET.SubElement(ppr, qn("w:ind"))
        if first_line_chars is not None:
            ind.set(qn("w:firstLineChars"), str(first_line_chars))
        if left is not None:
            ind.set(qn("w:left"), str(left))
        if hanging is not None:
            ind.set(qn("w:hanging"), str(hanging))

    add_text_runs(
        paragraph,
        text,
        font_ascii=font_ascii,
        font_east_asia=font_east_asia,
        size_half_points=size_half_points,
        bold=bold,
    )
    return paragraph


def create_title_paragraph(text: str) -> ET.Element:
    return create_formatted_paragraph(
        text,
        align="center",
        font_ascii="Times New Roman",
        font_east_asia="黑体",
        size_half_points=32,
        bold=True,
        space_before=0,
        space_after=240,
        line=360,
    )


def create_heading_paragraph(text: str, level: int) -> ET.Element:
    size = 28 if level == 2 else 26
    before = 200 if level == 2 else 160
    after = 120 if level == 2 else 80
    return create_formatted_paragraph(
        text,
        align="left",
        font_ascii="Times New Roman",
        font_east_asia="黑体",
        size_half_points=size,
        bold=True,
        space_before=before,
        space_after=after,
        line=360,
    )


def create_body_paragraph(text: str) -> ET.Element:
    return create_formatted_paragraph(
        text,
        align="both",
        font_ascii="Times New Roman",
        font_east_asia="宋体",
        size_half_points=24,
        bold=False,
        space_before=0,
        space_after=0,
        line=420,
        first_line_chars=200,
    )


def create_list_paragraph(text: str) -> ET.Element:
    return create_formatted_paragraph(
        text,
        align="both",
        font_ascii="Times New Roman",
        font_east_asia="宋体",
        size_half_points=24,
        bold=False,
        space_before=0,
        space_after=0,
        line=420,
        left=540,
        hanging=540,
    )


def create_code_paragraph(text: str) -> ET.Element:
    return create_formatted_paragraph(
        text,
        align="left",
        font_ascii="Courier New",
        font_east_asia="等宽更纱黑体 SC",
        size_half_points=21,
        bold=False,
        space_before=60,
        space_after=60,
        line=300,
        left=420,
    )


def build_document_xml(blocks: List[Dict[str, object]]) -> bytes:
    document = ET.Element(qn("w:document"))
    body = ET.SubElement(document, qn("w:body"))

    for idx, block in enumerate(blocks):
        block_type = str(block.get("type", "paragraph"))
        text = str(block.get("text", ""))

        if block_type == "heading":
            level = int(block.get("level", 1))
            if idx == 0 and level == 1:
                paragraph = create_title_paragraph(text)
            else:
                paragraph = create_heading_paragraph(text, level=min(max(level, 2), 3))
        elif block_type == "list":
            paragraph = create_list_paragraph(text)
        elif block_type == "code":
            paragraph = create_code_paragraph(text)
        else:
            paragraph = create_body_paragraph(text)

        body.append(paragraph)

    add_section_properties(body)
    return ET.tostring(document, encoding="utf-8", xml_declaration=True)


def styles_xml_bytes() -> bytes:
    content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>
        <w:lang w:val="zh-CN" w:eastAsia="zh-CN"/>
        <w:sz w:val="24"/>
        <w:szCs w:val="24"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:jc w:val="both"/>
        <w:spacing w:before="0" w:after="0" w:line="420" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:jc w:val="both"/>
      <w:spacing w:before="0" w:after="0" w:line="420" w:lineRule="auto"/>
      <w:ind w:firstLineChars="200"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/>
      <w:lang w:val="zh-CN" w:eastAsia="zh-CN"/>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="Heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:jc w:val="center"/>
      <w:spacing w:before="0" w:after="240" w:line="360" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="黑体"/>
      <w:b/>
      <w:bCs/>
      <w:sz w:val="32"/>
      <w:szCs w:val="32"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="Heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:jc w:val="left"/>
      <w:spacing w:before="200" w:after="120" w:line="360" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="黑体"/>
      <w:b/>
      <w:bCs/>
      <w:sz w:val="28"/>
      <w:szCs w:val="28"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="Heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:jc w:val="left"/>
      <w:spacing w:before="160" w:after="80" w:line="360" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="黑体"/>
      <w:b/>
      <w:bCs/>
      <w:sz w:val="26"/>
      <w:szCs w:val="26"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Code">
    <w:name w:val="Code"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:jc w:val="left"/>
      <w:spacing w:before="60" w:after="60" w:line="300" w:lineRule="auto"/>
      <w:ind w:left="420"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:eastAsia="等宽更纱黑体 SC"/>
      <w:sz w:val="21"/>
      <w:szCs w:val="21"/>
    </w:rPr>
  </w:style>
</w:styles>
"""
    return content.encode("utf-8")


def content_types_xml_bytes() -> bytes:
    content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""
    return content.encode("utf-8")


def package_rels_xml_bytes() -> bytes:
    content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""
    return content.encode("utf-8")


def document_rels_xml_bytes() -> bytes:
    content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
"""
    return content.encode("utf-8")


def app_xml_bytes() -> bytes:
    content = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Codex</Application>
</Properties>
"""
    return content.encode("utf-8")


def core_xml_bytes(title: str) -> bytes:
    now = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    content = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="{CP_NS}" xmlns:dc="{DC_NS}" xmlns:dcterms="{DCTERMS_NS}" xmlns:xsi="{XSI_NS}">
  <dc:title>{escape(title)}</dc:title>
  <dc:creator>Codex</dc:creator>
  <cp:lastModifiedBy>Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{now}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{now}</dcterms:modified>
</cp:coreProperties>
"""
    return content.encode("utf-8")


def infer_title(blocks: List[Dict[str, object]], input_path: Path) -> str:
    for block in blocks:
        if block.get("type") == "heading" and int(block.get("level", 1)) == 1:
            title = str(block.get("text", "")).strip()
            if title:
                return title
    return input_path.stem


def write_docx(output_path: Path, blocks: List[Dict[str, object]], title: str) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml_bytes())
        zf.writestr("_rels/.rels", package_rels_xml_bytes())
        zf.writestr("docProps/core.xml", core_xml_bytes(title))
        zf.writestr("docProps/app.xml", app_xml_bytes())
        zf.writestr("word/document.xml", build_document_xml(blocks))
        zf.writestr("word/styles.xml", styles_xml_bytes())
        zf.writestr("word/_rels/document.xml.rels", document_rels_xml_bytes())


def main() -> int:
    args = parse_args()
    input_path = Path(args.input)
    output_path = Path(args.output)

    markdown_text = input_path.read_text(encoding="utf-8")
    blocks = parse_markdown_blocks(markdown_text)
    if not blocks:
        raise SystemExit("Input Markdown is empty or contains no exportable content.")

    title = infer_title(blocks, input_path)
    write_docx(output_path, blocks, title)
    print(f"[OK] Output DOCX: {output_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
