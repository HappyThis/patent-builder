#!/usr/bin/env python3
"""Validate generated disclosure DOCX against extracted rule model."""

from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path
from typing import Dict, List, Optional

from ooxml_docx import (
    all_paragraphs,
    body_paragraphs,
    dump_json,
    extract_strict_placeholder_keys,
    normalize_key,
    paragraph_style_id,
    paragraph_text,
    parse_styles,
    read_xml_member,
    style_name,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Validate generated DOCX against template rules.")
    parser.add_argument("--output", required=True, help="Generated DOCX path")
    parser.add_argument("--rules", help="template-rules.json path")
    parser.add_argument("--template", help="Template DOCX path (for style drift comparison)")
    parser.add_argument("--report-json", help="Validation report JSON path")
    parser.add_argument("--strict", action="store_true", help="Treat warnings as blocking")
    return parser.parse_args()


def collect_required_sections(rules_payload: Optional[Dict[str, object]]) -> List[str]:
    if not isinstance(rules_payload, dict):
        return []
    hard_rules = rules_payload.get("hard_rules", {})
    if not isinstance(hard_rules, dict):
        return []
    sections = hard_rules.get("required_sections", [])
    if not isinstance(sections, list):
        return []
    return [s for s in sections if isinstance(s, str) and s.strip()]


def collect_output_texts(doc_root) -> List[str]:
    texts = []
    for p in body_paragraphs(doc_root):
        text = paragraph_text(p).strip()
        if text:
            texts.append(text)
    return texts


def first_match_index(texts: List[str], section: str) -> Optional[int]:
    norm_target = normalize_key(section)
    for idx, text in enumerate(texts):
        norm_text = normalize_key(text)
        if norm_target and (norm_target in norm_text or norm_text == norm_target):
            return idx
    return None


def detect_placeholders(doc_root) -> List[Dict[str, object]]:
    leftovers = []
    for idx, p in enumerate(all_paragraphs(doc_root)):
        text = paragraph_text(p).strip()
        if not text:
            continue
        keys = extract_strict_placeholder_keys(text)
        if keys:
            leftovers.append(
                {
                    "paragraph_index": idx,
                    "text": text,
                    "keys": keys,
                }
            )
    return leftovers


def dominant_body_style(doc_root, style_map: Dict[str, object]) -> Optional[str]:
    counter: Counter[str] = Counter()
    for p in body_paragraphs(doc_root):
        text = paragraph_text(p).strip()
        if not text:
            continue
        sid = paragraph_style_id(p)
        sname = style_name(sid, style_map) if sid else None
        if sname and ("Heading" in sname or "标题" in sname):
            continue
        if sid:
            counter[sid] += 1
    if not counter:
        return None
    return counter.most_common(1)[0][0]


def main() -> int:
    args = parse_args()

    output_path = Path(args.output)
    out_root = read_xml_member(output_path, "word/document.xml")
    try:
        out_styles_root = read_xml_member(output_path, "word/styles.xml")
    except KeyError:
        out_styles_root = None
    out_style_map, _ = parse_styles(out_styles_root)

    rules_payload = None
    if args.rules:
        rules_payload = json.loads(Path(args.rules).read_text(encoding="utf-8"))

    required_sections = collect_required_sections(rules_payload)
    output_texts = collect_output_texts(out_root)

    missing_sections = []
    section_positions = {}
    for section in required_sections:
        pos = first_match_index(output_texts, section)
        if pos is None:
            missing_sections.append(section)
        else:
            section_positions[section] = pos

    ordered_positions = [section_positions[s] for s in required_sections if s in section_positions]
    order_issue = ordered_positions != sorted(ordered_positions)

    placeholders = detect_placeholders(out_root)

    style_drift = None
    if args.template:
        tpl_root = read_xml_member(Path(args.template), "word/document.xml")
        try:
            tpl_styles_root = read_xml_member(Path(args.template), "word/styles.xml")
        except KeyError:
            tpl_styles_root = None
        tpl_style_map, _ = parse_styles(tpl_styles_root)

        tpl_dom = dominant_body_style(tpl_root, tpl_style_map)
        out_dom = dominant_body_style(out_root, out_style_map)
        if tpl_dom and out_dom and tpl_dom != out_dom:
            style_drift = {
                "template_dominant_style_id": tpl_dom,
                "output_dominant_style_id": out_dom,
                "template_style_name": style_name(tpl_dom, tpl_style_map),
                "output_style_name": style_name(out_dom, out_style_map),
            }

    warnings = []
    blocking = []

    if placeholders:
        blocking.append("Unresolved placeholders remain in output DOCX.")
    if missing_sections:
        blocking.append("Required sections are missing in output DOCX.")
    if order_issue:
        warnings.append("Required sections appear out of template order.")
    if style_drift:
        warnings.append("Dominant body style differs from template.")

    report = {
        "output": str(output_path),
        "rules": str(args.rules) if args.rules else None,
        "template": str(args.template) if args.template else None,
        "summary": {
            "blocking_count": len(blocking),
            "warning_count": len(warnings),
            "required_sections": len(required_sections),
            "missing_sections": len(missing_sections),
            "placeholder_leftovers": len(placeholders),
        },
        "blocking": blocking,
        "warnings": warnings,
        "details": {
            "missing_sections": missing_sections,
            "section_positions": section_positions,
            "order_issue": order_issue,
            "placeholder_leftovers": placeholders,
            "style_drift": style_drift,
        },
    }

    if args.report_json:
        dump_json(Path(args.report_json), report)

    print(f"[INFO] blocking={len(blocking)} warnings={len(warnings)}")
    if missing_sections:
        print(f"[WARN] missing sections: {missing_sections}")
    if placeholders:
        print(f"[WARN] unresolved placeholders: {len(placeholders)}")
    if style_drift:
        print("[WARN] style drift detected")

    should_fail = bool(blocking) or (args.strict and bool(warnings))
    if should_fail:
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
