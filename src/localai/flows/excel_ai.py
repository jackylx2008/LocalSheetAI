from __future__ import annotations

from dataclasses import replace
from pathlib import Path
from typing import Any

from localai.context import AppContext
from localai.logging_config import get_logger
from localai.modules.config_loader import as_int
from localai.modules.excel_reader import (
    ExcelSheetReadConfig,
    find_excel_file,
    read_sheet_snapshot,
    resolve_sheet_targets,
)
from localai.modules.llamacpp_client import LlamaCppClient, LlamaCppConfig


logger = get_logger(__name__)


def run(ctx: AppContext, analyze: bool = False, include_preview: bool = False) -> dict[str, Any]:
    flow_config = ctx.flow_config("excel_ai")
    excel_config = ExcelSheetReadConfig.from_config(flow_config)
    sheet_targets = resolve_sheet_targets(excel_config, flow_config)
    input_dir = ctx.resolve_path(ctx.app_config.get("input_path", "./input/"))

    logger.info("Step 1/2: preparing configured llama.cpp model service")
    llama_config = LlamaCppConfig.from_config(ctx.config, ctx.project_root)
    llama_config = replace(
        llama_config,
        timeout_sec=as_int(flow_config.get("timeout_sec"), llama_config.timeout_sec),
    )
    logger.info("llama.cpp base_url=%s", llama_config.base_url)
    logger.info("llama.cpp configured model=%s", llama_config.model)
    logger.info("llama.cpp server_path=%s", llama_config.server_path)
    logger.info("llama.cpp model_path=%s", llama_config.model_path)
    client = LlamaCppClient(llama_config)
    try:
        health, models = client.ensure_server()
        client.assert_model_available(models)
        available_model_ids = client.model_ids(models)
        logger.info("llama.cpp service is ready; health=%s; models=%s", health, available_model_ids)

        logger.info("Step 2/2: opening configured Excel workbook(s) and reading sheet(s)")
        logger.info("Excel input_dir=%s", input_dir)
        logger.info("Excel configured targets=%s", [(target.input_file, target.sheet_name) for target in sheet_targets])

        snapshots = []
        for target in sheet_targets:
            target_config = replace(excel_config, input_file=target.input_file, sheet_name=target.sheet_name)
            workbook_path = find_excel_file(input_dir, target_config.input_file)
            logger.info("Excel workbook selected=%s", workbook_path)
            snapshot = read_sheet_snapshot(workbook_path, target_config)
            snapshots.append(snapshot)
            logger.info(
                "Excel sheet loaded; requested=%s actual=%s max_row=%s max_column=%s rows_read=%s cols_read=%s truncated=%s",
                snapshot.requested_sheet_name or "<first sheet>",
                snapshot.actual_sheet_name,
                snapshot.max_row,
                snapshot.max_column,
                snapshot.rows_read,
                snapshot.cols_read,
                snapshot.truncated,
            )

        result: dict[str, Any] = {
            "status": "prepared" if not analyze else "analyzed",
            "llamacpp": {
                "base_url": llama_config.base_url,
                "model": llama_config.model,
                "server_path": llama_config.server_path,
                "extra_dll_dirs": llama_config.extra_dll_dirs,
                "n_gpu_layers": llama_config.n_gpu_layers,
                "ctx_size": llama_config.ctx_size,
                "health": health,
                "available_models": available_model_ids,
            },
            "excel": {
                "target_count": len(snapshots),
                "targets": [_snapshot_summary(snapshot) for snapshot in snapshots],
            },
        }

        if include_preview:
            result["excel"]["previews"] = [
                {
                    "workbook": str(snapshot.workbook_path),
                    "sheet": snapshot.actual_sheet_name,
                    "preview": snapshot.text[:1000],
                }
                for snapshot in snapshots
            ]

        if not analyze:
            logger.info("Prepared configured model and Excel sheet. Analysis step is skipped.")
            return result

        prompt = build_prompt(
            instruction=str(flow_config.get("prompt", "")).strip(),
            sheet_blocks=[
                {
                    "workbook_name": snapshot.workbook_path.name,
                    "sheet_name": snapshot.actual_sheet_name,
                    "sheet_text": snapshot.text,
                    "truncated": snapshot.truncated,
                }
                for snapshot in snapshots
            ],
        )
        answer = client.chat(prompt, max_tokens=as_int(flow_config.get("max_tokens"), 1024))

        output_path = ctx.resolve_path(str(flow_config.get("output_path", "./output/excel_ai_response.md")))
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(answer, encoding="utf-8")

        result["output_path"] = str(output_path)
        result["answer"] = answer
        return result
    finally:
        client.shutdown_server()


def build_prompt(
    instruction: str,
    sheet_blocks: list[dict[str, Any]],
) -> str:
    instruction_text = instruction or "请阅读下面的 Excel 工作表内容，并给出简明分析。"
    blocks = []
    for item in sheet_blocks:
        truncated_note = "内容已按配置截断，请只基于可见数据分析。" if item["truncated"] else "内容未触发截断。"
        blocks.append(
            f"工作簿文件名：{item['workbook_name']}\n"
            f"工作表：{item['sheet_name']}\n"
            f"截断说明：{truncated_note}\n"
            "```tsv\n"
            f"{item['sheet_text']}\n"
            "```"
        )
    return (
        f"{instruction_text}\n\n"
        "下面是从配置工作簿和工作表抽取的 TSV 文本，第一批非空行通常包含表头或说明：\n\n"
        + "\n\n".join(blocks)
    )


def _snapshot_summary(snapshot: Any) -> dict[str, Any]:
    return {
        "workbook": str(snapshot.workbook_path),
        "sheet": snapshot.actual_sheet_name,
        "requested_sheet": snapshot.requested_sheet_name,
        "available_sheets": snapshot.available_sheet_names,
        "max_row": snapshot.max_row,
        "max_column": snapshot.max_column,
        "rows_read": snapshot.rows_read,
        "cols_read": snapshot.cols_read,
        "truncated": snapshot.truncated,
    }
