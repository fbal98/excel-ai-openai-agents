"""Command‑line interface for the Autonomous Excel Assistant."""

import argparse
import asyncio
import os

from dotenv import load_dotenv
from agents import Runner

from .agent_core import excel_assistant_agent
from .context import AppContext


def parse_args() -> argparse.Namespace:
    """Parse CLI arguments."""
    parser = argparse.ArgumentParser(description="Autonomous Excel Assistant")
    parser.add_argument(
        "--input-file",
        type=str,
        required=False,
        help="Path to input Excel file (optional; a new workbook is created if omitted)",
    )
    parser.add_argument(
        "--output-file",
        type=str,
        required=False,
        help="Path to save the output Excel file (ignored in --live mode if omitted)",
    )
    parser.add_argument(
        "--instruction",
        type=str,
        required=True,
        help="Instruction for the agent (natural language)",
    )
    parser.add_argument(
        "--live",
        action="store_true",
        help="Edit the workbook in‑process via xlwings so changes appear in real time.",
    )
    return parser.parse_args()


async def main() -> None:
    load_dotenv()
    if not os.getenv("OPENAI_API_KEY"):
        raise RuntimeError("OPENAI_API_KEY not set in environment or .env file.")

    args = parse_args()

    # ------------------------------------------------------------------ #
    #  Select Excel manager implementation                               #
    # ------------------------------------------------------------------ #
    if args.live:
        try:
            from .live_excel_ops import LiveExcelManager as Manager
        except ImportError as exc:
            raise RuntimeError("xlwings is required for --live mode; pip install xlwings") from exc
    else:
        from .excel_ops import ExcelManager as Manager  # type: ignore

    # ------------------------------------------------------------------ #
    #  Initialise context & run agent                                    #
    # ------------------------------------------------------------------ #
    excel_manager = Manager(file_path=args.input_file) if args.input_file else Manager()
    app_context = AppContext(excel_manager=excel_manager)

    print(f"Running agent (live={args.live}) with instruction: {args.instruction}")
    result = await Runner.run(
        excel_assistant_agent,
        input=args.instruction,
        context=app_context,
        max_turns=25,
    )
    print(f"Agent finished. Final output: {result.final_output}")

    # ------------------------------------------------------------------ #
    #  Persist workbook if not in live mode                              #
    # ------------------------------------------------------------------ #
    if not args.live:
        if not args.output_file:
            raise ValueError("--output-file is required in batch mode.")
        try:
            excel_manager.save_workbook(args.output_file)
            print(f"Workbook saved to {args.output_file}")
        except Exception as exc:
            print(f"Failed to save workbook: {exc}")

    print("Agent result:")
    print(result)


if __name__ == "__main__":
    asyncio.run(main())