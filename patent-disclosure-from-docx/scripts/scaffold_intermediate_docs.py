#!/usr/bin/env python3
"""Scaffold collaboration and drafting intermediates from template rules."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Dict, List


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Generate intermediate files for proposal loop, stage-2 and stage-3."
    )
    parser.add_argument("--rules", required=True, help="Path to template-rules.json")
    parser.add_argument("--out-dir", required=True, help="Output directory")
    parser.add_argument("--force", action="store_true", help="Overwrite existing files")
    return parser.parse_args()


def unique_titles(titles: List[str]) -> List[str]:
    seen = set()
    out: List[str] = []
    for title in titles:
        key = title.strip().lower()
        if not key or key in seen:
            continue
        seen.add(key)
        out.append(title)
    return out


def ensure_write(path: Path, content: str, force: bool) -> None:
    if path.exists() and not force:
        raise FileExistsError(f"Refusing to overwrite existing file without --force: {path}")
    path.write_text(content, encoding="utf-8")


def build_solution_spec_md(required_sections: List[str]) -> str:
    lines: List[str] = []
    lines.append("# 技术方案协作文档")
    lines.append("")
    lines.append("## 1. 发明信息（自动推断，可覆写）")
    lines.append("- 发明名称（默认由已批准技术方案自动生成）：")
    lines.append("- 适用技术领域（默认自动生成）：")
    lines.append("- 保护对象（方法/系统/装置/介质，默认自动判定）：")
    lines.append("")
    lines.append("## 2. 目标与约束")
    lines.append("- 目标问题：")
    lines.append("- 业务场景：")
    lines.append("- 约束条件（性能、成本、兼容、合规）：")
    lines.append("")
    lines.append("## 3. 技术方案总览")
    lines.append("- 方案总思路：")
    lines.append("- 关键流程（输入 -> 处理 -> 输出）：")
    lines.append("- 核心模块：")
    lines.append("")
    lines.append("## 4. 关键创新与保护点")
    lines.append("- 区别于现有方案的关键点（按重要性排序）：")
    lines.append("- 每个关键点带来的有益效果：")
    lines.append("")
    lines.append("## 5. 可替代方案与边界")
    lines.append("- 可替代实施方式：")
    lines.append("- 失效边界或不适用场景：")
    lines.append("")
    lines.append("## 6. 与模板章节映射")
    for title in required_sections:
        lines.append(f"### {title}")
        lines.append("- 需要覆盖的技术要点：")
        lines.append("- 证据或示例：")
        lines.append("")
    return "\n".join(lines).strip() + "\n"


def build_proposal_session_md(required_sections: List[str]) -> str:
    lines: List[str] = []
    lines.append("# 方案推荐与确认会话")
    lines.append("")
    lines.append("## 1. 会话状态")
    lines.append("- 当前轮次：1")
    lines.append("- 是否批准进入写作：否")
    lines.append("- 已确认写作方向：")
    lines.append("- 用户确认语句（原文）：")
    lines.append("")
    lines.append("## 2. 用户输入")
    lines.append("- 原始想法（用户提供）：")
    lines.append("- 输入摘要（系统整理）：")
    lines.append("- 明确约束（成本/时延/精度/部署/合规）：")
    lines.append("")
    lines.append("## 3. 候选写作方向（可由用户选择或自定义）")
    lines.append("### 方向候选 A")
    lines.append("- 方向名称：")
    lines.append("- 方向定位：")
    lines.append("- 匹配理由：")
    lines.append("- 潜在保护重点：")
    lines.append("- 风险与信息缺口：")
    lines.append("")
    lines.append("### 方向候选 B")
    lines.append("- 方向名称：")
    lines.append("- 方向定位：")
    lines.append("- 匹配理由：")
    lines.append("- 潜在保护重点：")
    lines.append("- 风险与信息缺口：")
    lines.append("")
    lines.append("## 4. 当前方向下技术方案（持续深化）")
    lines.append("- 当前方向：")
    lines.append("- 方向来源（候选/用户自定义）：")
    lines.append("- 要解决的核心技术问题：")
    lines.append("- 核心机理：")
    lines.append("- 输入 -> 处理 -> 输出：")
    lines.append("- 关键模块：")
    lines.append("- 关键保护点：")
    lines.append("- 预期有益效果：")
    lines.append("- 实施风险：")
    lines.append("- 未决问题：")
    lines.append("")
    lines.append("## 5. 自动推断信息（由 skill 填充）")
    lines.append("- 推断发明标题：")
    lines.append("- 推断保护对象：")
    lines.append("- 推断适用场景：")
    lines.append("- 推断依据：")
    lines.append("")
    lines.append("## 6. 本轮待确认问题（最多 3 个）")
    lines.append("- Q1：你当前最想写的方向是什么？你也可以直接给出自定义方向。")
    lines.append("")
    lines.append("## 7. 与模板章节映射（由 skill 自动填充）")
    for title in required_sections:
        lines.append(f"- {title}：")
    lines.append("")
    lines.append("## 8. 轮次记录")
    lines.append("- round=1 | action=init | note=已进入方向式对话模式，等待用户提供想法并选择/自定义方向")
    lines.append("")
    lines.append("## 9. 批准闸门")
    lines.append("- 在用户明确同意前，不进入 disclosure-draft 写作。")
    lines.append("- 用户明确表示“可以写作了/开始生成doc”等语句后，自动推断标题/保护对象并进入写作。")
    return "\n".join(lines).strip() + "\n"


def build_draft_md(required_sections: List[str], requirement_map: Dict[str, List[str]]) -> str:
    lines: List[str] = []
    lines.append("# 交底书草稿")
    lines.append("")
    lines.append("- 填写说明：将每个章节写成可直接放入交底书的正文段落，避免只写功能描述。")
    lines.append("- 语言要求：直接、干练、工程化，避免“我建议”“优选地”“本发明”等表达。")
    lines.append("- 排版要求：长段落拆分，步骤单独成段，公式单独成行。")
    lines.append("- 公式要求：最终 DOCX 中应以 Word 公式对象表示，不要只写成普通文本。")
    lines.append("")

    for title in required_sections:
        lines.append(f"## {title}")
        hints = requirement_map.get(title, [])
        if hints:
            lines.append("")
            lines.append("参考要求：")
            for hint in hints[:3]:
                lines.append(f"- {hint}")
            lines.append("")
        lines.append("[在此填写本章节正文。建议至少 1-3 段。]")
        lines.append("")

    return "\n".join(lines).strip() + "\n"


def build_solution_spec_json() -> Dict[str, object]:
    return {
        "title": "",
        "title_auto_inferred": True,
        "title_inference_basis": [],
        "application_domain": "",
        "application_domain_auto_inferred": True,
        "protection_targets": [],
        "protection_targets_auto_inferred": True,
        "problem_statement": "",
        "constraints": [],
        "overall_approach": "",
        "workflow": [],
        "modules": [
            {
                "name": "",
                "responsibility": "",
                "inputs": [],
                "outputs": [],
            }
        ],
        "innovations": [
            {
                "point": "",
                "difference_vs_prior_art": "",
                "benefit": "",
            }
        ],
        "alternatives": [],
        "boundaries": [],
        "evidence_or_examples": [],
    }


def build_proposal_session_json(required_sections: List[str]) -> Dict[str, object]:
    return {
        "session": {
            "round": 1,
            "approved_for_drafting": False,
            "approved_direction": "",
            "approval_statement": "",
        },
        "user_input": {
            "raw": "",
            "summary": "",
            "constraints": [],
        },
        "direction_candidates": [
            {
                "name": "",
                "positioning": "",
                "fit_rationale": "",
                "potential_protection_points": [],
                "risks_or_gaps": [],
            },
            {
                "name": "",
                "positioning": "",
                "fit_rationale": "",
                "potential_protection_points": [],
                "risks_or_gaps": [],
            },
        ],
        "selected_direction": {
            "name": "",
            "source": "",
            "status": "pending",
        },
        "current_technical_solution": {
            "core_problem": "",
            "core_mechanism": "",
            "workflow_steps": [],
            "key_modules": [],
            "protection_points": [],
            "expected_benefits": [],
            "implementation_risks": [],
            "open_items": [],
            "confidence_notes": [],
        },
        "iteration_state": {
            "ready_to_write": False,
            "pending_questions": [],
            "resolved_items": [],
        },
        "auto_inference": {
            "invention_title": "",
            "protection_targets": [],
            "application_scenario": "",
            "mapping_notes": [],
            "inference_basis": [],
        },
        "template_section_targets": required_sections,
        "questions_for_user": [
            "请先说说你对这份交底书内容的想法，我会给出可写作方向，你可选或自定义。"
        ],
        "round_history": [
            {
                "round": 1,
                "action": "init",
                "note": "initialized proposal loop artifacts with direction-based iterative mode",
            }
        ],
    }


def build_draft_json(required_sections: List[str]) -> Dict[str, object]:
    return {
        "title": "交底书草稿",
        "sections": [
            {
                "title": title,
                "blocks": [],
                "tables": [],
            }
            for title in required_sections
        ],
    }


def main() -> int:
    args = parse_args()
    rules_path = Path(args.rules)
    out_dir = Path(args.out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    payload = json.loads(rules_path.read_text(encoding="utf-8"))

    required_sections = payload.get("hard_rules", {}).get("required_sections") or []
    if not required_sections:
        required_sections = [
            item.get("title", "")
            for item in payload.get("section_tree", [])
            if isinstance(item, dict)
        ]
    required_sections = unique_titles([s for s in required_sections if isinstance(s, str)])

    requirement_map: Dict[str, List[str]] = {title: [] for title in required_sections}
    for req in payload.get("content_requirements", []):
        if not isinstance(req, dict):
            continue
        ctx = req.get("section_context")
        text = req.get("text")
        if isinstance(ctx, str) and isinstance(text, str) and ctx in requirement_map:
            requirement_map[ctx].append(text)

    solution_md = out_dir / "solution-spec.md"
    solution_json = out_dir / "solution-spec.json"
    proposal_md = out_dir / "proposal-session.md"
    proposal_json = out_dir / "proposal-session.json"
    draft_md = out_dir / "disclosure-draft.md"
    draft_json = out_dir / "disclosure-draft.json"

    ensure_write(solution_md, build_solution_spec_md(required_sections), args.force)
    ensure_write(
        solution_json,
        json.dumps(build_solution_spec_json(), ensure_ascii=False, indent=2) + "\n",
        args.force,
    )
    ensure_write(proposal_md, build_proposal_session_md(required_sections), args.force)
    ensure_write(
        proposal_json,
        json.dumps(build_proposal_session_json(required_sections), ensure_ascii=False, indent=2) + "\n",
        args.force,
    )
    ensure_write(draft_md, build_draft_md(required_sections, requirement_map), args.force)
    ensure_write(
        draft_json,
        json.dumps(build_draft_json(required_sections), ensure_ascii=False, indent=2) + "\n",
        args.force,
    )

    print(f"[OK] {solution_md}")
    print(f"[OK] {solution_json}")
    print(f"[OK] {proposal_md}")
    print(f"[OK] {proposal_json}")
    print(f"[OK] {draft_md}")
    print(f"[OK] {draft_json}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
