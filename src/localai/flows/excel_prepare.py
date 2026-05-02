from __future__ import annotations

from dataclasses import replace
from typing import Any

from localai.context import AppContext
from localai.logging_config import get_logger
from localai.modules.config_loader import as_int
from localai.modules.excel_row_extractor import ExcelIssueExtractConfig, ExcelIssueRow, extract_issue_rows
from localai.modules.excel_target_resolver import ExcelTarget, resolve_excel_targets
from localai.modules.llamacpp_client import LlamaCppClient, LlamaCppConfig


logger = get_logger(__name__)


def prepare(ctx: AppContext, ensure_llama: bool = True) -> tuple[LlamaCppClient | None, dict[str, Any], list[ExcelTarget], list[ExcelIssueRow]]:
    flow_config = ctx.flow_config("excel_ai")
    input_dir = ctx.resolve_path(ctx.app_config.get("input_path", "./input/"))
    targets = resolve_excel_targets(input_dir, flow_config)

    client: LlamaCppClient | None = None
    llama_status: dict[str, Any] = {"checked": False}
    if ensure_llama:
        llama_config = LlamaCppConfig.from_config(ctx.config, ctx.project_root)
        llama_config = replace(
            llama_config,
            timeout_sec=as_int(flow_config.get("timeout_sec"), llama_config.timeout_sec),
        )
        client = LlamaCppClient(llama_config)
        health, models = client.ensure_server()
        client.assert_model_available(models)
        llama_status = {
            "checked": True,
            "base_url": llama_config.base_url,
            "model": llama_config.model,
            "server_path": llama_config.server_path,
            "extra_dll_dirs": llama_config.extra_dll_dirs,
            "n_gpu_layers": llama_config.n_gpu_layers,
            "ctx_size": llama_config.ctx_size,
            "health": health,
            "available_models": client.model_ids(models),
        }

    extract_config = ExcelIssueExtractConfig.from_config(flow_config)
    issue_rows: list[ExcelIssueRow] = []
    for target in targets:
        rows = extract_issue_rows(target.workbook_path, target.sheet_name, extract_config)
        issue_rows.extend(rows)
        logger.info(
            "Excel issues extracted; workbook=%s sheet=%s rows=%s missing_problem_description=%s",
            target.workbook_path,
            target.sheet_name,
            len(rows),
            sum(1 for row in rows if row.missing_problem_description),
        )

    return client, llama_status, targets, issue_rows


def run(ctx: AppContext, ensure_llama: bool = True, include_samples: bool = False) -> dict[str, Any]:
    client = None
    try:
        client, llama_status, targets, issue_rows = prepare(ctx, ensure_llama=ensure_llama)
        result: dict[str, Any] = {
            "status": "prepared",
            "llamacpp": llama_status,
            "excel": {
                "target_count": len(targets),
                "row_count": len(issue_rows),
                "targets": [
                    {
                        "pattern": target.pattern,
                        "workbook": str(target.workbook_path),
                        "sheet": target.sheet_name,
                        "row_count": sum(
                            1
                            for row in issue_rows
                            if row.workbook_path == target.workbook_path and row.sheet_name.strip() == target.sheet_name.strip()
                        ),
                    }
                    for target in targets
                ],
                "missing_problem_description_count": sum(row.missing_problem_description for row in issue_rows),
            },
        }
        if include_samples:
            result["excel"]["samples"] = [
                {
                    "workbook": str(row.workbook_path),
                    "sheet": row.sheet_name,
                    "row_number": row.row_number,
                    "problem_text": row.problem_text,
                    "missing_problem_description": row.missing_problem_description,
                    "row_context": row.row_context,
                }
                for row in issue_rows[:10]
            ]
        return result
    finally:
        if client is not None:
            client.shutdown_server()
