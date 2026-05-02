from __future__ import annotations

import argparse
import sys
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parent
SRC_PATH = str(PROJECT_ROOT / "src")
if SRC_PATH not in sys.path:
    sys.path.insert(0, SRC_PATH)

from localai.entrypoints import bootstrap_context, print_json
from localai.flows import category_draft, excel_ai, excel_classify, excel_prepare


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Prepare Excel issue rows and draft category rules with local AI. "
            "The llama-server process started by this command is stopped before exit."
        )
    )
    parser.add_argument("--config", default="config.yaml", help="Path to config.yaml.")
    parser.add_argument(
        "--mode",
        choices=["prepare", "draft", "classify", "legacy"],
        default="prepare",
        help="prepare extracts structured rows; draft generates category_skills_draft.md; classify writes AI categories to output; legacy keeps the old TSV analysis flow.",
    )
    parser.add_argument(
        "--analyze",
        action="store_true",
        help="Legacy mode only: send the configured sheet content to the local model for analysis.",
    )
    parser.add_argument(
        "--show-preview",
        action="store_true",
        help="Print sample rows in prepare mode, or a TSV preview in legacy mode.",
    )
    parser.add_argument(
        "--no-llama",
        action="store_true",
        help="Prepare mode only: skip local AI service checks and only extract Excel issue rows.",
    )
    args = parser.parse_args()

    ctx = bootstrap_context(__file__, args.config)
    if args.mode == "draft":
        result = category_draft.run(ctx)
    elif args.mode == "classify":
        result = excel_classify.run(ctx)
    elif args.mode == "legacy":
        result = excel_ai.run(ctx, analyze=args.analyze, include_preview=args.show_preview)
    else:
        result = excel_prepare.run(ctx, ensure_llama=not args.no_llama, include_samples=args.show_preview)
    print_json(result)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
