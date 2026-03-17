#!/usr/bin/env python3
"""Render a final DOCX by combining template + intermediate draft artifacts."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import xml.etree.ElementTree as ET

from ooxml_docx import (
    NS,
    all_paragraphs,
    append_before_sectpr,
    body_element,
    body_paragraphs,
    choose_default_paragraph_style,
    choose_heading_style,
    create_equation_paragraph,
    create_paragraph,
    dump_json,
    estimate_match_score,
    extract_placeholder_keys,
    normalize_key,
    paragraph_style_id,
    paragraph_text,
    parse_markdown_sections,
    parse_styles,
    qn,
    read_xml_member,
    section_aliases,
    style_name,
    write_docx_members,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Render DOCX from draft json/markdown and a template.")
    parser.add_argument("--template", required=True, help="Template DOCX path")
    parser.add_argument("--draft-json", help="Draft JSON path")
    parser.add_argument("--draft-md", help="Draft Markdown path")
    parser.add_argument("--rules", help="Template rules JSON path")
    parser.add_argument("--output", required=True, help="Output DOCX path")
    parser.add_argument("--mode", choices=["conservative", "aggressive"], default="conservative")
    parser.add_argument("--report-json", help="Render report JSON path")
    parser.add_argument(
        "--clean-template-guides",
        action="store_true",
        help="Remove template guidance/instruction paragraphs based on rules model.",
    )
    return parser.parse_args()


def normalize_blocks(raw_items: object) -> List[Dict[str, str]]:
    if isinstance(raw_items, str):
        raw_items = [raw_items]
    if not isinstance(raw_items, list):
        return []

    blocks: List[Dict[str, str]] = []
    for item in raw_items:
        if isinstance(item, str):
            text = item.strip()
            if text:
                blocks.append({"type": "text", "text": text})
            continue
        if not isinstance(item, dict):
            continue
        block_type = str(item.get("type", "text")).strip().lower() or "text"
        if block_type not in {"text", "equation"}:
            block_type = "text"
        text_value = item.get("text", item.get("value", ""))
        if not isinstance(text_value, str):
            continue
        text = text_value.strip()
        if text:
            blocks.append({"type": block_type, "text": text})
    return blocks


def load_sections(draft_json: Optional[Path], draft_md: Optional[Path]) -> List[Dict[str, object]]:
    if draft_json and draft_json.exists():
        payload = json.loads(draft_json.read_text(encoding="utf-8"))
        sections = payload.get("sections", []) if isinstance(payload, dict) else []
        normalized: List[Dict[str, object]] = []
        for section in sections:
            if not isinstance(section, dict):
                continue
            title = section.get("title")
            if not isinstance(title, str) or not title.strip():
                continue
            blocks = normalize_blocks(section.get("blocks", section.get("paragraphs", [])))
            normalized.append({"title": title.strip(), "blocks": blocks})
        if normalized:
            return normalized

    if draft_md and draft_md.exists():
        parsed = parse_markdown_sections(draft_md)
        return [
            {
                "title": item.get("title", ""),
                "blocks": normalize_blocks(item.get("paragraphs", [])),
            }
            for item in parsed
            if item.get("title")
        ]

    return []


def paragraph_is_heading(paragraph: ET.Element, style_map: Dict[str, object]) -> bool:
    sid = paragraph_style_id(paragraph)
    sname = style_name(sid, style_map) if style_map else None
    text = paragraph_text(paragraph).strip()
    if not text:
        return False
    if sname and ("Heading" in sname or "标题" in sname):
        return True
    if len(text) <= 26 and not text.endswith(("：", ":")):
        return True
    return False


def find_heading_index(body: ET.Element, title: str, style_map: Dict[str, object]) -> Optional[int]:
    target_norm = normalize_key(title)
    aliases = [normalize_key(x) for x in section_aliases(title)]
    best: Tuple[int, Optional[int]] = (0, None)

    for idx, child in enumerate(list(body)):
        if child.tag != qn("w:p"):
            continue
        if not paragraph_is_heading(child, style_map):
            continue
        text = paragraph_text(child).strip()
        candidate_norm = normalize_key(text)
        if not candidate_norm:
            continue

        score = estimate_match_score(target_norm, candidate_norm)
        for alias in aliases:
            score = max(score, estimate_match_score(alias, candidate_norm))

        if score > best[0]:
            best = (score, idx)

    return best[1] if best[0] >= 70 else None


def create_block_node(block: Dict[str, str], style_id: Optional[str]) -> ET.Element:
    block_type = block.get("type", "text")
    text = block.get("text", "")
    if block_type == "equation":
        return create_equation_paragraph(text, style_id=style_id)
    return create_paragraph(text, style_id=style_id)


def insert_blocks_after(body: ET.Element, index: int, blocks: List[Dict[str, str]], style_id: Optional[str]) -> int:
    inserted = 0
    insert_pos = index + 1
    for block in blocks:
        node = create_block_node(block, style_id=style_id)
        body.insert(insert_pos, node)
        insert_pos += 1
        inserted += 1
    return inserted


def build_placeholder_map(document_root: ET.Element) -> Dict[str, List[ET.Element]]:
    mapping: Dict[str, List[ET.Element]] = {}
    for paragraph in all_paragraphs(document_root):
        text = paragraph_text(paragraph).strip()
        if not text:
            continue
        keys = extract_placeholder_keys(text)
        for key in keys:
            nk = normalize_key(key)
            if not nk:
                continue
            mapping.setdefault(nk, []).append(paragraph)
    return mapping


def replace_paragraph_with_blocks(
    body: ET.Element,
    target_paragraph: ET.Element,
    blocks: List[Dict[str, str]],
    fallback_style_id: Optional[str],
) -> int:
    children = list(body)
    try:
        index = children.index(target_paragraph)
    except ValueError:
        return 0

    style_id = paragraph_style_id(target_paragraph) or fallback_style_id
    body.remove(target_paragraph)
    insert_pos = index
    for block in blocks:
        body.insert(insert_pos, create_block_node(block, style_id=style_id))
        insert_pos += 1
    return len(blocks)


def load_guide_texts_from_rules(rules_path: Optional[Path]) -> List[str]:
    if not rules_path or not rules_path.exists():
        return []

    payload = json.loads(rules_path.read_text(encoding="utf-8"))
    texts: List[str] = []

    for item in payload.get("content_requirements", []):
        if isinstance(item, dict):
            text = item.get("text")
            if isinstance(text, str) and text.strip():
                texts.append(text.strip())

    for item in payload.get("section_tree", []):
        if not isinstance(item, dict):
            continue
        title = item.get("title")
        if not isinstance(title, str):
            continue
        title = title.strip()
        if not title:
            continue
        # Keep concise section headings and only collect long guidance-like titles.
        if len(title) > 20 or any(ch in title for ch in ("，", "。", "；", "（", "）", ":", "：", "、")):
            texts.append(title)

    normalized = []
    seen = set()
    for text in texts:
        if text not in seen:
            seen.add(text)
            normalized.append(text)
    return normalized


def prune_template_guides(body: ET.Element, guide_texts: List[str], protected_texts: Optional[List[str]] = None) -> int:
    if not guide_texts:
        return 0

    targets = {text.strip() for text in guide_texts if text.strip()}
    protected = {text.strip() for text in (protected_texts or []) if text and text.strip()}
    removed = 0
    for child in list(body):
        if child.tag != qn("w:p"):
            continue
        text = paragraph_text(child).strip()
        if text in protected:
            continue
        if text in targets:
            body.remove(child)
            removed += 1
    return removed


def prune_empty_body_paragraphs(body: ET.Element) -> int:
    removed = 0
    for child in list(body):
        if child.tag != qn("w:p"):
            continue
        if paragraph_text(child).strip():
            continue
        if child.findall(".//m:oMathPara", NS) or child.findall(".//m:oMath", NS):
            continue
        if child.findall(".//w:br", NS):
            continue
        if child.findall(".//w:fldChar", NS):
            continue
        if child.findall(".//w:drawing", NS):
            continue
        if child.findall(".//w:object", NS):
            continue
        body.remove(child)
        removed += 1
    return removed


def render(
    document_root: ET.Element,
    style_map: Dict[str, object],
    sections: List[Dict[str, object]],
    mode: str,
    guide_texts: Optional[List[str]] = None,
) -> Dict[str, object]:
    body = body_element(document_root)
    placeholder_map = build_placeholder_map(document_root)

    heading_style_1 = choose_heading_style(style_map, {v.name: k for k, v in style_map.items()}, level=1)
    heading_style_2 = choose_heading_style(style_map, {v.name: k for k, v in style_map.items()}, level=2)
    body_style = choose_default_paragraph_style(style_map, {v.name: k for k, v in style_map.items()})

    report = {
        "mode": mode,
        "sections_total": len(sections),
        "placeholder_replaced": [],
        "heading_inserted": [],
        "appended": [],
        "unresolved": [],
        "guidance_paragraphs_removed": 0,
        "empty_paragraphs_removed": 0,
    }

    append_bucket: List[Dict[str, object]] = []

    for section in sections:
        title = str(section.get("title", "")).strip()
        blocks = normalize_blocks(section.get("blocks", section.get("paragraphs", [])))
        if not title or not blocks:
            report["unresolved"].append(
                {"title": title or "(empty)", "reason": "missing title or blocks"}
            )
            continue

        aliases = [normalize_key(x) for x in section_aliases(title)]
        target_keys = [normalize_key(title)] + aliases

        replaced = False
        for key in target_keys:
            if key in placeholder_map and placeholder_map[key]:
                inserted = replace_paragraph_with_blocks(body, placeholder_map[key][0], blocks, body_style)
                report["placeholder_replaced"].append({"title": title, "key": key, "blocks_inserted": inserted})
                replaced = True
                break

        if replaced:
            continue

        heading_idx = find_heading_index(body, title, style_map)
        if heading_idx is not None:
            inserted = insert_blocks_after(body, heading_idx, blocks, body_style)
            report["heading_inserted"].append(
                {"title": title, "heading_index": heading_idx, "blocks_inserted": inserted}
            )
            continue

        append_bucket.append({"title": title, "blocks": blocks})

    if append_bucket:
        if mode == "conservative":
            append_before_sectpr(body, create_paragraph("自动生成内容（待确认插入位置）", heading_style_1))

        for section in append_bucket:
            append_before_sectpr(body, create_paragraph(section["title"], heading_style_2 or heading_style_1))
            for block in section["blocks"]:
                append_before_sectpr(body, create_block_node(block, body_style))
            report["appended"].append({"title": section["title"], "blocks": len(section["blocks"])})

    if guide_texts:
        protected_titles = [str(s.get("title", "")).strip() for s in sections if str(s.get("title", "")).strip()]
        report["guidance_paragraphs_removed"] = prune_template_guides(
            body,
            guide_texts,
            protected_texts=protected_titles,
        )

    report["empty_paragraphs_removed"] = prune_empty_body_paragraphs(body)

    return report


def main() -> int:
    args = parse_args()

    template = Path(args.template)
    output = Path(args.output)
    draft_json = Path(args.draft_json) if args.draft_json else None
    draft_md = Path(args.draft_md) if args.draft_md else None
    rules_path = Path(args.rules) if args.rules else None

    sections = load_sections(draft_json, draft_md)
    if not sections:
        raise SystemExit("No draft sections found. Provide --draft-json or --draft-md with content.")

    document_root = read_xml_member(template, "word/document.xml")
    try:
        styles_root = read_xml_member(template, "word/styles.xml")
    except KeyError:
        styles_root = None
    style_map, _ = parse_styles(styles_root)

    guide_texts: List[str] = []
    if args.clean_template_guides:
        guide_texts = load_guide_texts_from_rules(rules_path)

    render_report = render(document_root, style_map, sections, args.mode, guide_texts=guide_texts)
    render_report["template"] = str(template)
    render_report["output"] = str(output)

    xml_bytes = ET.tostring(document_root, encoding="utf-8", xml_declaration=True)
    output.parent.mkdir(parents=True, exist_ok=True)
    write_docx_members(template, output, {"word/document.xml": xml_bytes})

    if args.report_json:
        dump_json(Path(args.report_json), render_report)

    print(f"[OK] Output DOCX: {output}")
    if args.report_json:
        print(f"[OK] Render report: {args.report_json}")
    print(
        "[INFO] replaced={} inserted={} appended={} unresolved={}".format(
            len(render_report["placeholder_replaced"]),
            len(render_report["heading_inserted"]),
            len(render_report["appended"]),
            len(render_report["unresolved"]),
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
