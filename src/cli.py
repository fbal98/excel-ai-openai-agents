"""Command‑line interface for the Autonomous Excel Assistant (single realtime mode)."""

from __future__ import annotations

import argparse
import asyncio
import logging
import os
import re
import sys
import time
from typing import Any, Optional
import uuid

from dotenv import load_dotenv
from agents import Runner, ItemHelpers, trace
from agents.stream_events import RunItemStreamEvent
from openai.types.responses import ResponseTextDeltaEvent

from .agent_core import excel_assistant_agent
from .context import AppContext
from .excel_ops import ExcelManager


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
    p.add_argument("--trace-off", action="store_true", help="Disable OpenAI tracing for this run.")
    return p.parse_args()


# ──────────────────────────────
#  Streaming helper
# ──────────────────────────────
async def handle_streaming(result, verbose: bool):
    """
    Consume a RunResultStreaming, display live progress, and
    return an object with `.final_output`.
    """
    logger = logging.getLogger(__name__)
    final_output: str = ""
    last_message: Optional[str] = None

    try:
        async for event in result.stream_events():
            # Raw token‑level deltas
            if event.type == "raw_response_event" and isinstance(event.data, ResponseTextDeltaEvent):
                delta: str = event.data.delta
                final_output += delta
                if last_message is None:
                    # Keep a rolling copy so non‑verbose runs still get the full text
                    last_message = final_output
                if verbose:
                    print(delta, end="", flush=True)
                continue

            if event.type != "run_item_stream_event":
                continue

            item = event.item  # type: RunItemStreamEvent
            if item.type == "tool_call_item":
                if verbose:
                    print(f"🛠️  {item.function_name}")
            elif item.type == "tool_call_output_item":
                if verbose:
                    ok = "✔" if "error" not in item.output else "✖"
                    print(f"   ↳ {ok} {item.output}")
            elif item.type == "message_output_item":
                msg_text = ItemHelpers.text_message_output(item)
                last_message = msg_text
                if verbose:
                    print(f"💬 {msg_text}")
    except Exception as e:
        logger.error("Streaming error: %s", e)
        final_output += f"\n\n[Streaming error: {e}]"

    class _Result:
        def __init__(self, text: str):  # noqa: D401
            self.final_output = text

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

    # Use async context manager for ExcelManager to ensure proper resource management
    try:
        print("DEBUG: Initializing ExcelManager...")
        # Make manager visible by default, allow attaching
        async with ExcelManager(file_path=args.input_file, visible=True, attach_existing=True) as excel_mgr:
            print("DEBUG: ExcelManager initialized")
            ctx = AppContext(excel_manager=excel_mgr)
            print("DEBUG: AppContext initialized")

            # Initial workbook‑shape scan
            if ctx.update_shape():
                logger.debug("Initial workbook shape scanned (v%s).", ctx.shape.version)
            else:
                logger.warning("Initial workbook shape scan failed; proceeding without shape info.")

            if args.input_file:
                logger.info("📂 Opened workbook: %s", args.input_file)
            else:
                logger.info("🆕 Started new workbook.")

            # ────────── Interactive mode ──────────
            print(f"DEBUG: Checking interactive mode (args.interactive={args.interactive})")
            if args.interactive:
                print("DEBUG: Entering interactive mode block")
                chat: list[dict[str, str]] = []
                print("Hello! How can I help you today?")
                print("(Enter your message, use multiple lines if needed. Submit with an empty line)")
                while True:
                    try:
                        print("DEBUG: Waiting for multi-line input...")
                        lines = []
                        while True:
                            line = input("> " if not lines else "... ")  # Different prompt for continuation lines
                            if not line:
                                break
                            lines.append(line)
                        user = "\n".join(lines)
                        print(f"DEBUG: Received multi-line input: {user}")
                        
                        if not user:  # Skip if only empty line was entered
                            continue
                            
                        if user.lower() in {"exit", "quit"}:
                            break
                            
                        chat.append({"role": "user", "content": user})
                        print("DEBUG: Calling Runner.run...")
                        try:
                            # Always call directly, bypassing the trace context manager
                            res = await Runner.run(
                                excel_assistant_agent,
                                input=chat,
                                context=ctx,
                                max_turns=25,
                                # trace_id=str(uuid.uuid4()) # Trace ID not needed if not tracing
                            )
                            print("DEBUG: Runner.run completed")
                            
                            # Ensure all Excel changes are applied before giving feedback to the user
                            try:
                                print("DEBUG: Ensuring Excel changes are applied...")
                                await excel_mgr.ensure_changes_applied()
                                print("DEBUG: Excel changes applied.")
                            except Exception as e:
                                print(f"DEBUG: Error ensuring Excel changes: {e}")
                            
                            reply = res.final_output
                            chat.append({"role": "assistant", "content": reply})
                            # ---- Buffer‑window memory: keep last 4 user‑assistant pairs ----
                            if len(chat) > 8:
                                chat = chat[-8:]
                            print(reply)
                        except Exception as e:
                            # Basic error reporting for interactive mode
                            error_msg = f"Error processing your request: {e}"
                            print(error_msg)
                            # Add a placeholder response in chat history
                            chat.append({"role": "assistant", "content": f"Sorry, I encountered an error: {e}"})
                            if len(chat) > 8:
                                chat = chat[-8:]
                    except (EOFError, KeyboardInterrupt):
                        print("\nExiting.")
                        break
                print("DEBUG: Exiting interactive loop")
                return

            # ────────── One‑shot / scripted mode ──────────
            print("DEBUG: Entering one-shot mode block")
            logger.info(f"\n💡 Instruction: {args.instruction}")
            start = time.monotonic()
            try:
                if args.stream:
                    # Always call directly, bypassing the trace context manager
                    streamed = await Runner.run_streamed(
                        excel_assistant_agent,
                        input=args.instruction,
                        context=ctx,
                        max_turns=25,
                        # trace_id=str(uuid.uuid4()) # Trace ID not needed if not tracing
                    )
                    result = await handle_streaming(streamed, args.verbose)
                else:
                    # Always call directly, bypassing the trace context manager
                    result = await Runner.run(
                        excel_assistant_agent,
                        input=args.instruction,
                        context=ctx,
                        max_turns=25,
                        # trace_id=str(uuid.uuid4()) # Trace ID not needed if not tracing
                    )
            except Exception as e:
                # Basic error reporting for one-shot mode
                logger.error(f"❌ Agent error: {e}")
                sys.exit(1)
            elapsed = time.monotonic() - start
            logger.info(f"✅ Done in {elapsed:.1f}s\n\n📤 {result.final_output}\n")

            # ────────── Optional explicit save ──────────
            if args.output_file:
                try:
                    await excel_mgr.ensure_changes_applied()
                    saved_path = await excel_mgr.save_with_confirmation(args.output_file)
                    logger.info("💾 Workbook saved to %s", saved_path)
                except Exception as e:
                    logger.error("❌ Failed to save workbook: %s", e)
                    try:
                        saved_path = await excel_mgr.save_with_confirmation()
                        logger.info("💾 Workbook saved to fallback location: %s", saved_path)
                    except Exception:
                        logger.error("All save attempts failed.")
    except Exception as e:
        logger.error("❌ Fatal error: %s", e)
        sys.exit(1)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nExiting.")
        sys.exit(0)