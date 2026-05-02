from __future__ import annotations

from typing import Any

from localai.context import AppContext


def run(ctx: AppContext) -> dict[str, Any]:
    return {
        "flow": "export",
        "output_dir": str(ctx.output_dir),
        "config": ctx.flow_config("export"),
    }
