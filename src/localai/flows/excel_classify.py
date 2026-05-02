from __future__ import annotations

from dataclasses import dataclass
from typing import Any
import re

from localai.context import AppContext
from localai.logging_config import get_logger
from localai.flows.excel_prepare import prepare
from localai.modules.category_policy import SheetCategoryPolicy, load_category_policies
from localai.modules.config_loader import as_int
from localai.modules.excel_classification_writer import RowClassification, write_classified_workbooks
from localai.modules.excel_row_extractor import ExcelIssueRow, ExcelIssueExtractConfig


logger = get_logger(__name__)


@dataclass(frozen=True)
class ClassificationStats:
    total_rows: int
    ai_rows: int
    missing_problem_description_rows: int
    fallback_rows: int


def run(ctx: AppContext) -> dict[str, Any]:
    flow_config = ctx.flow_config("excel_ai")
    client = None
    try:
        client, llama_status, targets, issue_rows = prepare(ctx, ensure_llama=True)
        if client is None:
            raise RuntimeError("llama.cpp client is required for Excel classification")

        draft_path = ctx.resolve_path(str(flow_config.get("draft_output_path", "./output/category_skills_draft.md")))
        policies = load_category_policies(draft_path)
        max_tokens = as_int(flow_config.get("classify_max_tokens"), 32)
        max_context_chars = as_int(flow_config.get("classify_max_context_chars"), 1000)

        classifications: list[RowClassification] = []
        fallback_rows = 0
        ai_rows = 0

        for index, row in enumerate(issue_rows, start=1):
            policy = policies.get(row.sheet_name)
            if policy is None:
                raise RuntimeError(f"No category policy found for sheet: {row.sheet_name}")

            if row.missing_problem_description:
                category = "缺少问题描述"
            else:
                ai_rows += 1
                answer = client.chat(
                    _build_classify_prompt(row, policy, max_context_chars=max_context_chars),
                    max_tokens=max_tokens,
                )
                category, used_fallback = _normalize_category(answer, policy.labels)
                fallback_rows += int(used_fallback)

            classifications.append(
                RowClassification(
                    workbook_path=row.workbook_path,
                    sheet_name=row.sheet_name,
                    row_number=row.row_number,
                    category=category,
                )
            )

            if index == 1 or index % 25 == 0 or index == len(issue_rows):
                logger.info("Classified Excel issue rows: %s/%s", index, len(issue_rows))

        extract_config = ExcelIssueExtractConfig.from_config(flow_config)
        header_text = str(flow_config.get("category_header") or "AI整理归类项").strip() or "AI整理归类项"
        output_paths = write_classified_workbooks(
            classifications,
            ctx.output_dir,
            header_row=extract_config.header_row,
            header_text=header_text,
        )

        stats = ClassificationStats(
            total_rows=len(issue_rows),
            ai_rows=ai_rows,
            missing_problem_description_rows=sum(row.missing_problem_description for row in issue_rows),
            fallback_rows=fallback_rows,
        )
        return {
            "status": "classified",
            "llamacpp": llama_status,
            "category_draft": str(draft_path),
            "outputs": [str(path) for path in output_paths],
            "excel": {
                "target_count": len(targets),
                "row_count": len(issue_rows),
                "ai_classified_rows": stats.ai_rows,
                "missing_problem_description_rows": stats.missing_problem_description_rows,
                "fallback_rows": stats.fallback_rows,
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


def _build_classify_prompt(
    row: ExcelIssueRow,
    policy: SheetCategoryPolicy,
    max_context_chars: int,
) -> str:
    labels = [*policy.labels, "其他问题"]
    context = _compact_row_context(row, max_context_chars)
    return (
        "你是工程质量问题分类器。请严格根据分类口径，为这一行问题选择一个且仅一个分类标签。\n"
        "只能输出标签本身，不要输出解释、标点、编号或 Markdown。\n\n"
        f"可选标签：{', '.join(labels)}\n\n"
        "分类口径：\n"
        f"{policy.markdown}\n\n"
        "待分类行：\n"
        f"工作簿：{row.workbook_path.name}\n"
        f"Sheet：{row.sheet_name}\n"
        f"行号：{row.row_number}\n"
        f"问题描述及原因初步判定：{row.problem_text}\n"
        f"整行上下文：{context}\n\n"
        "最终分类标签："
    )


def _compact_row_context(row: ExcelIssueRow, max_context_chars: int) -> str:
    parts = [f"{key}={value}" for key, value in row.row_context.items() if value]
    context = "；".join(parts) if parts else "无"
    if len(context) <= max_context_chars:
        return context
    return context[: max_context_chars - 1] + "…"


def _normalize_category(answer: str, labels: list[str]) -> tuple[str, bool]:
    allowed = [*labels, "其他问题"]
    text = _clean_answer(answer)
    for label in allowed:
        if text == label:
            return label, False
    for label in allowed:
        if label in text:
            return label, False
    return "其他问题", True


def _clean_answer(answer: str) -> str:
    text = answer.strip()
    text = re.sub(r"^```(?:text|markdown)?", "", text).strip()
    text = re.sub(r"```$", "", text).strip()
    text = text.splitlines()[0].strip() if text else ""
    text = text.strip("：:，,。.;；\"'`* \t")
    text = re.sub(r"^\d+[.、)\s]+", "", text).strip()
    return text
