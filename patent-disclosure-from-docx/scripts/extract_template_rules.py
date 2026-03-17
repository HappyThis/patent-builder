#!/usr/bin/env python3
"""Extract template rules from DOCX and generate JSON/Markdown artifacts."""

from __future__ import annotations

import argparse
from collections import Counter
from pathlib import Path
from typing import Dict, List, Optional
import xml.etree.ElementTree as ET

from ooxml_docx import (
    NS,
    body_paragraphs,
    dump_json,
    estimate_match_score,
    extract_placeholder_keys,
    infer_heading_level,
    is_heading_candidate,
    normalize_key,
    now_utc_iso,
    paragraph_style_id,
    paragraph_text,
    parse_layout_rules,
    parse_styles,
    read_xml_member,
    style_info_to_dict,
    style_name,
)

REQUIREMENT_KEYWORDS = (
    "应",
    "应当",
    "需要",
    "需",
    "必须",
    "不得",
    "避免",
    "请",
    "写明",
    "说明",
    "列举",
    "提供",
    "不少于",
    "至少",
)

CANONICAL_SECTION_HINTS = (
    "技术领域",
    "背景技术",
    "技术方案",
    "发明内容",
    "附图说明",
    "实施方式",
    "有益效果",
    "保护点",
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Extract rule model from DOCX template.")
    parser.add_argument("--template", required=True, help="Template DOCX path")
    parser.add_argument("--out-json", required=True, help="Output JSON path")
    parser.add_argument("--out-md", required=True, help="Output Markdown path")
    return parser.parse_args()


def nearest_section(section_tree: List[Dict[str, object]], paragraph_idx: int) -> Optional[str]:
    found = None
    for section in section_tree:
        idx = section.get("paragraph_index")
        if isinstance(idx, int) and idx <= paragraph_idx:
            found = section.get("title")
        else:
            break
    return found


def pick_required_sections(section_tree: List[Dict[str, object]]) -> List[str]:
    seen = set()
    ordered: List[str] = []

    canonical_hits = []
    canonical_titles = (
        "技术领域",
        "背景技术",
        "技术方案描述(重点)",
        "技术方案描述",
        "技术方案",
        "发明内容",
        "技术方案的关键点和保护点",
        "有助于理解本技术方案的技术资料(没有的话，可以不填)",
        "有助于理解本技术方案的技术资料",
    )

    for section in section_tree:
        title = str(section.get("title", "")).strip()
        if not title:
            continue
        if title in canonical_titles:
            canonical_hits.append(title)
            continue
        if len(title) <= 20 and any(hint in title for hint in CANONICAL_SECTION_HINTS):
            canonical_hits.append(title)

    source = canonical_hits if canonical_hits else [
        str(s.get("title", ""))
        for s in section_tree
        if isinstance(s.get("title"), str) and len(str(s.get("title")).strip()) <= 20
    ]
    for title in source:
        if not title:
            continue
        key = normalize_key(title)
        if key in seen:
            continue
        if "注意事项" in title:
            continue
        seen.add(key)
        ordered.append(title)

    return ordered[:12]


def build_markdown_report(payload: Dict[str, object]) -> str:
    lines: List[str] = []
    lines.append("# Template Rule Extraction Report")
    lines.append("")
    lines.append(f"- Source template: `{payload.get('source_template')}`")
    lines.append(f"- Generated at (UTC): `{payload.get('generated_at_utc')}`")
    lines.append("")

    confidence = payload.get("confidence", {})
    lines.append("## Confidence")
    lines.append("")
    lines.append(f"- section_detection: `{confidence.get('section_detection')}`")
    lines.append(f"- anchor_detection: `{confidence.get('anchor_detection')}`")
    lines.append(f"- overall: `{confidence.get('overall')}`")
    lines.append("")

    lines.append("## Required Sections")
    lines.append("")
    required = payload.get("hard_rules", {}).get("required_sections", [])
    if required:
        for title in required:
            lines.append(f"- {title}")
    else:
        lines.append("- (none inferred)")
    lines.append("")

    lines.append("## Section Tree")
    lines.append("")
    for item in payload.get("section_tree", []):
        level = item.get("level")
        title = item.get("title")
        style = item.get("style_name") or item.get("style_id")
        idx = item.get("paragraph_index")
        lines.append(f"- L{level} | P{idx} | {title} | style={style}")
    lines.append("")

    lines.append("## Anchor Hints")
    lines.append("")
    anchors = payload.get("anchors", [])
    if anchors:
        for a in anchors[:50]:
            lines.append(
                f"- {a.get('type')} | key={a.get('key')} | P{a.get('paragraph_index')} | section={a.get('section_context')}"
            )
    else:
        lines.append("- (no anchors found)")
    lines.append("")

    lines.append("## Requirement Snippets")
    lines.append("")
    reqs = payload.get("content_requirements", [])
    if reqs:
        for req in reqs[:80]:
            lines.append(f"- P{req.get('paragraph_index')} | {req.get('section_context')} | {req.get('text')}")
    else:
        lines.append("- (none extracted)")

    return "\n".join(lines) + "\n"


def main() -> int:
    args = parse_args()

    template_path = Path(args.template)
    out_json = Path(args.out_json)
    out_md = Path(args.out_md)

    document_root = read_xml_member(template_path, "word/document.xml")
    try:
        styles_root = read_xml_member(template_path, "word/styles.xml")
    except KeyError:
        styles_root = None

    style_map, _ = parse_styles(styles_root)

    records: List[Dict[str, object]] = []
    usage_counter: Counter[str] = Counter()
    for idx, paragraph in enumerate(body_paragraphs(document_root)):
        text = paragraph_text(paragraph).strip()
        if not text:
            continue
        sid = paragraph_style_id(paragraph)
        sname = style_name(sid, style_map)
        if sid:
            usage_counter[sid] += 1
        records.append(
            {
                "paragraph_index": idx,
                "text": text,
                "style_id": sid,
                "style_name": sname,
            }
        )

    section_tree: List[Dict[str, object]] = []
    stack: List[Dict[str, object]] = []
    for rec in records:
        text = str(rec["text"])
        sname = rec.get("style_name")
        if not is_heading_candidate(text, sname if isinstance(sname, str) else None):
            continue

        level = infer_heading_level(sname if isinstance(sname, str) else None, text)
        if level is None:
            level = 1

        while stack and int(stack[-1]["level"]) >= level:
            stack.pop()

        node = {
            "title": text,
            "level": level,
            "paragraph_index": rec["paragraph_index"],
            "style_id": rec.get("style_id"),
            "style_name": rec.get("style_name"),
            "path": [s["title"] for s in stack] + [text],
        }
        section_tree.append(node)
        stack.append(node)

    anchors: List[Dict[str, object]] = []
    requirements: List[Dict[str, object]] = []

    for rec in records:
        text = str(rec["text"])
        pidx = int(rec["paragraph_index"])
        section_context = nearest_section(section_tree, pidx)

        keys = extract_placeholder_keys(text)
        for raw_key in keys:
            anchors.append(
                {
                    "type": "label-colon" if text.endswith(("：", ":")) else "placeholder",
                    "key": raw_key,
                    "key_norm": normalize_key(raw_key),
                    "raw": text,
                    "paragraph_index": pidx,
                    "section_context": section_context,
                }
            )

        if any(k in text for k in REQUIREMENT_KEYWORDS):
            requirements.append(
                {
                    "text": text,
                    "paragraph_index": pidx,
                    "section_context": section_context,
                }
            )

    required_sections = pick_required_sections(section_tree)

    style_usage = []
    for sid, count in usage_counter.most_common():
        info = style_map.get(sid)
        style_usage.append(
            {
                "style_id": sid,
                "style_name": info.name if info else None,
                "count": count,
            }
        )

    section_score = min(1.0, len(section_tree) / 8.0)
    anchor_score = min(1.0, len(anchors) / 6.0)
    overall = round((section_score * 0.6 + anchor_score * 0.4), 3)

    payload: Dict[str, object] = {
        "version": "1.0",
        "source_template": str(template_path),
        "generated_at_utc": now_utc_iso(),
        "layout_rules": parse_layout_rules(document_root),
        "style_catalog": [style_info_to_dict(info) for info in style_map.values()],
        "style_usage": style_usage,
        "section_tree": section_tree,
        "anchors": anchors,
        "content_requirements": requirements,
        "hard_rules": {
            "required_sections": required_sections,
            "no_unresolved_placeholders": True,
            "preserve_section_order": True,
        },
        "soft_rules": {
            "prefer_complete_technical_detail": True,
            "prefer_include_alternative_implementations": True,
            "prefer_explicit_beneficial_effects": True,
        },
        "fallback_policy": {
            "default_mode": "conservative",
            "when_missing_anchor": "append unresolved sections in generated block and flag for review",
            "when_low_confidence": "avoid risky inline insertion and emit warning",
        },
        "confidence": {
            "section_detection": round(section_score, 3),
            "anchor_detection": round(anchor_score, 3),
            "overall": overall,
            "notes": [
                "Score is heuristic; confirm required sections manually for uncommon templates.",
                "Low anchor score usually means no explicit placeholder exists in template.",
            ],
        },
    }

    out_json.parent.mkdir(parents=True, exist_ok=True)
    out_md.parent.mkdir(parents=True, exist_ok=True)
    dump_json(out_json, payload)
    out_md.write_text(build_markdown_report(payload), encoding="utf-8")

    print(f"[OK] Rule model JSON: {out_json}")
    print(f"[OK] Rule model report: {out_md}")
    print(f"[INFO] sections={len(section_tree)} anchors={len(anchors)} requirements={len(requirements)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
