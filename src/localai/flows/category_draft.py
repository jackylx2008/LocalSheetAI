from __future__ import annotations

from collections import defaultdict
import re
from typing import Any

from localai.context import AppContext
from localai.modules.category_prompt_builder import DEFAULT_SEED_CATEGORIES, build_category_draft_prompt
from localai.modules.category_skill_writer import write_category_draft
from localai.modules.config_loader import as_int
from localai.flows.excel_prepare import prepare


def run(ctx: AppContext) -> dict[str, Any]:
    flow_config = ctx.flow_config("excel_ai")
    client = None
    try:
        client, llama_status, targets, issue_rows = prepare(ctx, ensure_llama=True)
        if client is None:
            raise RuntimeError("llama.cpp client is required for category draft generation")

        seed_categories = flow_config.get("seed_categories") or DEFAULT_SEED_CATEGORIES
        if isinstance(seed_categories, str):
            seed_categories = [item.strip() for item in seed_categories.split(",") if item.strip()]
        grouped = defaultdict(list)
        for row in issue_rows:
            grouped[(row.workbook_path, row.sheet_name)].append(row)

        max_rows_per_sheet = as_int(flow_config.get("draft_max_rows_per_sheet"), 20)
        max_context_chars = as_int(flow_config.get("draft_max_context_chars"), 900)
        max_tokens = as_int(flow_config.get("draft_max_tokens"), 4096)
        sections = ["# 分类口径草案"]
        for (_workbook_path, _sheet_name), rows in grouped.items():
            prompt = build_category_draft_prompt(
                rows,
                seed_categories=list(seed_categories),
                max_rows_per_sheet=max_rows_per_sheet,
                max_context_chars=max_context_chars,
            )
            sections.append(_normalize_draft_section(client.chat(prompt, max_tokens=max_tokens)))
        sections.append(_fallback_section())

        answer = "\n\n".join(section.strip() for section in sections if section.strip())
        output_path = ctx.resolve_path(str(flow_config.get("draft_output_path", "./output/category_skills_draft.md")))
        write_category_draft(answer, output_path)

        return {
            "status": "draft_generated",
            "llamacpp": llama_status,
            "output_path": str(output_path),
            "excel": {
                "target_count": len(targets),
                "row_count": len(issue_rows),
                "targets": [
                    {
                        "pattern": target.pattern,
                        "workbook": str(target.workbook_path),
                        "sheet": target.sheet_name,
                    }
                    for target in targets
                ],
            },
        }
    finally:
        if client is not None:
            client.shutdown_server()


def _normalize_draft_section(markdown: str) -> str:
    text = markdown.strip()
    text = re.sub(r"(?m)^# 分类口径草案\s*", "", text).strip()
    text = re.sub(r"(?ms)^##\s*通用兜底分类\b.*?(?=^##\s|\Z)", "", text).strip()
    return text


def _fallback_section() -> str:
    return (
        "## 通用兜底分类\n\n"
        "- 缺少问题描述\n"
        "- 其他问题"
    )
