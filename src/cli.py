"""Command‑line interface for the Autonomous Excel Assistant (single realtime mode)."""

import argparse
import asyncio
import logging
import os
import sys
import time
from typing import Any, Optional

from dotenv import load_dotenv
from agents import Runner
from agents.stream_events import RunItemStreamEvent
from openai.types.responses import ResponseTextDeltaEvent

from .agent_core import excel_assistant_agent
from .context import AppContext
from .excel_ops import ExcelManager  # unified manager


# ──────────────────────────────
#  Argument parsing
# ──────────────────────────────
def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Autonomous Excel Assistant (realtime)")
    p.add_argument("--input-file", type=str, help="Path to an existing workbook.")
    p.add_argument("--output-file", type=str, help="Path to save at the end (optional).")
    p.add_argument("--instruction", type=str, default=None, help="Instruction for the agent.")
    p.add_argument("--interactive", "-i", action="store_true", help="Interactive chat mode.")
    p.add_argument("--stream", action="store_true", help="Stream the agent's reasoning/output.")
    p.add_argument("-v", "--verbose", action="store_true", help="Verbose logging.")
    return p.parse_args()


# ──────────────────────────────
#  Streaming helper
# ──────────────────────────────
async def handle_streaming(result, verbose: bool):
    logger = logging.getLogger(__name__)
    final_output, last_tool, last_message = "", None, None
    async for event in result.stream_events():
        if event.type == "raw_response_event" and isinstance(event.data, ResponseTextDeltaEvent):
            final_output += event.data.delta
            if verbose:
                print(event.data.delta, end="", flush=True)
        elif event.type == "run_item_stream_event":
            item = event  # type: RunItemStreamEvent
            if item.item.type == "tool_call_item":
                fn = getattr(item.item, "function_name", getattr(item.item, "name", "tool"))
                last_tool = fn
                if verbose:
                    logger.info(f"🛠️  {fn}")
            elif item.item.type == "tool_call_output_item" and verbose:
                logger.info(f"📊 Result from {last_tool}: {item.item.output}")
            elif item.item.type == "message_output_item":
                msg = getattr(item.item, "text", item.item.content)
                last_message = msg
                if verbose:
                    logger.info(f"💬 {msg}")
    class _Result:  # simple wrapper to mimic Runner result
        def __init__(self, text): self.final_output = text
    return _Result(last_message or final_output)


# ──────────────────────────────
#  Main entry
# ──────────────────────────────
async def main() -> None:
    print("DEBUG: Starting main()")
    load_dotenv()
    print("DEBUG: load_dotenv() called")
    if not os.getenv("OPENAI_API_KEY"):
        print("DEBUG: OPENAI_API_KEY not found, raising error")
        raise RuntimeError("OPENAI_API_KEY not set.")
    print("DEBUG: OPENAI_API_KEY found")

    args = parse_args()
    print(f"DEBUG: Args parsed: {args}")
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format="%(message)s")
    logger = logging.getLogger(__name__)
    print("DEBUG: Logging configured")

    # Init unified manager
    try:
        print("DEBUG: Initializing ExcelManager...")
        excel_mgr = ExcelManager(file_path=args.input_file)
        print("DEBUG: ExcelManager initialized")
        ctx = AppContext(excel_manager=excel_mgr)
        print("DEBUG: AppContext initialized")
    except Exception as e:
        print(f"DEBUG: Error during ExcelManager/AppContext init: {e}")
        raise

    if args.input_file:
        logger.info(f"📂 Opened workbook: {args.input_file}")
    else:
        logger.info("🆕 Started new workbook.")
    print("DEBUG: Workbook info logged")

    # ────────── Interactive mode ──────────
    print(f"DEBUG: Checking interactive mode (args.interactive={args.interactive})")
    if args.interactive:
        print("DEBUG: Entering interactive mode block")
        chat: list[dict[str, str]] = []
        print("Hello! How can I help you today?")
        while True:
            try:
                print("DEBUG: Waiting for input...")
                user = input("> ")
                print(f"DEBUG: Received input: {user}")
            except (EOFError, KeyboardInterrupt):
                print("\nExiting.")
                break
            if user.lower() in {"exit", "quit"}:
                break
            chat.append({"role": "user", "content": user})
            print("DEBUG: Calling Runner.run...")
            res = await Runner.run(excel_assistant_agent, input=chat, context=ctx, max_turns=25)
            print("DEBUG: Runner.run completed")
            reply = res.final_output
            chat.append({"role": "assistant", "content": reply})
            print(reply)
        print("DEBUG: Exiting interactive loop")
        return

    # ────────── One‑shot / scripted mode ──────────
    print("DEBUG: Entering one-shot mode block")
    logger.info(f"\n💡 Instruction: {args.instruction}")
    start = time.monotonic()
    try:
        if args.stream:
            streamed = Runner.run_streamed(excel_assistant_agent, input=args.instruction, context=ctx, max_turns=25)
            result = await handle_streaming(streamed, args.verbose)
        else:
            result = await Runner.run(excel_assistant_agent, input=args.instruction, context=ctx, max_turns=25)
    except Exception as e:
        logger.error(f"❌ Agent error: {e}")
        sys.exit(1)
    elapsed = time.monotonic() - start
    logger.info(f"✅ Done in {elapsed:.1f}s\n\n📤 {result.final_output}\n")

    # ────────── Optional explicit save ──────────
    if args.output_file:
        try:
            excel_mgr.save_as(args.output_file)
            logger.info(f"💾 Workbook saved to {args.output_file}")
        except Exception as e:
            logger.error(f"❌ Failed to save workbook: {e}")


# Execute main() when run directly
if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nExiting.")
        sys.exit(0)