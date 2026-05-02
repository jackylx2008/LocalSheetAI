from __future__ import annotations

from typing import Any

from localai.context import AppContext


def run(ctx: AppContext) -> dict[str, Any]:
    return {
        "flow": "scan",
        "input_path": str(ctx.resolve_path(ctx.app_config.get("input_path", "./input/"))),
        "config": ctx.flow_config("scan"),
    }
