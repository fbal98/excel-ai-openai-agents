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
#  Rate‑limit helpers
# ──────────────────────────────
def extract_retry_time(error: Exception) -> float:
    """
    Determine how long to wait before retrying after a rate‑limit error.

    • Prefer the OpenAI‑supplied ``error.retry_after`` (seconds).
    • Fallback: parse legacy "try again in Xs" strings.
    • Last resort: return a 10‑second default.
    """
    retry_after = getattr(error, "retry_after", None)
    if retry_after:
        try:
            return float(retry_after)
        except (TypeError, ValueError):
            pass

    match = re.search(r"try again in (\d+(?:\.\d+)?)s", str(error), flags=re.IGNORECASE)
    if match:
        return float(match.group(1))

    return 10.0  # Default back‑off


async def handle_rate_limit(func, *args, **kwargs):
    """
    Execute *func* with automatic one‑shot retry on HTTP‑429 responses.
    """
    logger = logging.getLogger(__name__)

    try:
        return await func(*args, **kwargs)
    except Exception as e:
        # Look for OpenAI style rate‑limit indicators
        if any(token in str(e).lower() for token in {"rate_limit_exceeded", "429"}):
            wait_time = extract_retry_time(e)
            logger.warning("⏱️  Rate limit exceeded. Sleeping %.1fs then retrying…", wait_time)
            await asyncio.sleep(wait_time)
            try:
                return await func(*args, **kwargs)
            except Exception as e2:
                logger.error("❌ Second attempt failed after back‑off: %s", e2)
                raise
        raise


# ──────────────────────────────
#  Main entry
# ──────────────────────────────
async def main() -> None:
    load_dotenv()
    if not os.getenv("OPENAI_API_KEY"):
        raise RuntimeError("OPENAI_API_KEY not set.")

    args = parse_args()
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(message)s",
    )
    logger = logging.getLogger(__name__)

    # Use async context manager for ExcelManager to ensure proper resource management
    try:
        async with ExcelManager(file_path=args.input_file, visible=True, attach_existing=True) as excel_mgr:
            ctx = AppContext(excel_manager=excel_mgr)

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
            if args.interactive:
                chat: list[dict[str, str]] = []
                print("Hello! How can I help you today?")
                print("(Enter multi‑line messages; finish with an empty line.)")

                while True:
                    try:
                        # Collect multi‑line input
                        lines: list[str] = []
                        while True:
                            line = input("> " if not lines else "... ")
                            if not line:
                                break
                            lines.append(line)
                        user_msg = "\n".join(lines)

                        if not user_msg:
                            continue
                        if user_msg.lower() in {"exit", "quit"}:
                            break

                        chat.append({"role": "user", "content": user_msg})

                        if args.trace_off:
                            res = await handle_rate_limit(
                                Runner.run,
                                excel_assistant_agent,
                                input=chat,
                                context=ctx,
                                max_turns=25,
                            )
                        else:
                            async with trace("Excel Assistant Run"):
                                res = await handle_rate_limit(
                                    Runner.run,
                                    excel_assistant_agent,
                                    input=chat,
                                    context=ctx,
                                    max_turns=25,
                                )

                        # Ensure Excel has applied changes before replying
                        await excel_mgr.ensure_changes_applied()

                        reply = res.final_output
                        chat.append({"role": "assistant", "content": reply})
                        # Keep last 4 user‑assistant pairs
                        if len(chat) > 8:
                            chat = chat[-8:]
                        print(reply)
                    except (EOFError, KeyboardInterrupt):
                        print("\nExiting.")
                        break
                return  # End interactive mode

            # ────────── One‑shot / scripted mode ──────────
            logger.info("💡 Instruction: %s", args.instruction)
            start = time.monotonic()

            try:
                if args.stream:
                    if args.trace_off:
                        streamed = await handle_rate_limit(
                            Runner.run_streamed,
                            excel_assistant_agent,
                            input=args.instruction,
                            context=ctx,
                            max_turns=25,
                        )
                    else:
                        async with trace("Excel Assistant Run"):
                            streamed = await handle_rate_limit(
                                Runner.run_streamed,
                                excel_assistant_agent,
                                input=args.instruction,
                                context=ctx,
                                max_turns=25,
                            )
                    result = await handle_streaming(streamed, args.verbose)
                else:
                    if args.trace_off:
                        result = await handle_rate_limit(
                            Runner.run,
                            excel_assistant_agent,
                            input=args.instruction,
                            context=ctx,
                            max_turns=25,
                        )
                    else:
                        async with trace("Excel Assistant Run"):
                            result = await handle_rate_limit(
                                Runner.run,
                                excel_assistant_agent,
                                input=args.instruction,
                                context=ctx,
                                max_turns=25,
                            )
            except Exception as e:
                if any(tok in str(e).lower() for tok in {"rate_limit_exceeded", "429"}):
                    logger.error("❌ Rate‑limit error after retry: %s", e)
                    sys.exit(1)
                logger.error("❌ Agent error: %s", e)
                sys.exit(1)

            elapsed = time.monotonic() - start
            logger.info("✅ Done in %.1fs\n\n📤 %s\n", elapsed, result.final_output)

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