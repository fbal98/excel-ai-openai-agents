"""Command‑line interface for the Autonomous Excel Assistant (single realtime mode)."""

import argparse
import asyncio
import logging
import os
import sys
import time
import re
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
    
    try:
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
    except Exception as e:
        err_str = str(e).lower()
        if "rate_limit_exceeded" in err_str or "429" in err_str:
            logger.warning(f"⚠️ Rate limit exceeded during streaming: {e}")
            # Add the error message to the output so the user knows what happened
            rate_limit_msg = "\n\n[Streaming halted due to OpenAI rate limit. Some content may be missing.]"
            if last_message:
                last_message += rate_limit_msg
            else:
                final_output += rate_limit_msg
        else:
            logger.error(f"Error during streaming: {e}")
            # Add a generic error message
            error_msg = f"\n\n[Streaming error: {e}]"
            if last_message:
                last_message += error_msg
            else:
                final_output += error_msg
    
    class _Result:  # simple wrapper to mimic Runner result
        def __init__(self, text): self.final_output = text
    return _Result(last_message or final_output)


# ──────────────────────────────
#  Rate limit helper
# ──────────────────────────────
def extract_retry_time(error_message: str) -> float:
    """Extract retry time from OpenAI rate limit error message or return a default time."""
    # Try to extract "Please try again in X.XXXs" pattern
    pattern = r"try again in (\d+\.\d+)s"
    match = re.search(pattern, str(error_message))
    if match:
        return float(match.group(1))
    
    # If that fails, use a default backoff time
    return 10.0  # Default to 10 seconds


async def handle_rate_limit(func, *args, **kwargs):
    """
    Wrapper to handle rate limit errors with automatic retry.
    If a rate limit error is encountered, wait and retry once.
    
    Args:
        func: Async function to call
        *args, **kwargs: Arguments to pass to the function
        
    Returns:
        Result from successful function call
        
    Raises:
        The original exception if retry also fails
    """
    logger = logging.getLogger(__name__)
    
    try:
        return await func(*args, **kwargs)
    except Exception as e:
        err_str = str(e).lower()
        
        # Check if it's a rate limit error (429)
        if "rate_limit_exceeded" in err_str or "429" in err_str:
            # Extract wait time or use default
            wait_time = extract_retry_time(err_str)
            logger.warning(f"⏱️ Rate limit exceeded. Waiting {wait_time:.1f}s before retry...")
            
            # Sleep for the specified time
            await asyncio.sleep(wait_time)
            
            # Try one more time
            try:
                logger.info("🔄 Retrying API call after rate limit backoff...")
                return await func(*args, **kwargs)
            except Exception as e2:
                logger.error(f"❌ Second attempt also failed after rate limit backoff: {e2}")
                raise  # Propagate the error
        else:
            # Not a rate limit error, propagate immediately
            raise


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

            # --- Perform initial workbook shape scan using the AppContext helper ---
            print("DEBUG: Performing initial workbook shape scan...")
            initial_scan_success = ctx.update_shape() # This now handles logging internally
            if not initial_scan_success:
                logger.warning("Initial workbook shape scan failed. Proceeding without initial shape info (will retry on first write).")
            else:
                # Log success (already done inside update_shape, but can add CLI specific msg if needed)
                print(f"DEBUG: Initial shape scanned (v{ctx.shape.version if ctx.shape else 'N/A'})")
                # Optional: Perform an initial state dump if needed
                # ctx.dump_state_to_json("initial_state_dump.json")
            # --- End initial scan ---

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
                            res = await handle_rate_limit(Runner.run, excel_assistant_agent, input=chat, context=ctx, max_turns=25)
                            print("DEBUG: Runner.run completed")
                            
                            # Ensure all Excel changes are applied before giving feedback to the user
                            try:
                                print("DEBUG: Ensuring Excel changes are applied...")
                                excel_mgr.ensure_changes_applied()
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
                            err_str = str(e).lower()
                            if "rate_limit_exceeded" in err_str or "429" in err_str:
                                error_msg = "Sorry, I hit the OpenAI API rate limit and couldn't process your request even after waiting."
                                error_msg += "\nThis can happen with large Excel files or frequent requests."
                                error_msg += "\nPlease try again in a few minutes or consider simplifying your Excel data."
                                print(error_msg)
                                # Don't add error to chat history
                            else:
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
                    streamed = await handle_rate_limit(Runner.run_streamed, excel_assistant_agent, input=args.instruction, context=ctx, max_turns=25)
                    result = await handle_streaming(streamed, args.verbose)
                else:
                    result = await handle_rate_limit(Runner.run, excel_assistant_agent, input=args.instruction, context=ctx, max_turns=25)
            except Exception as e:
                err_str = str(e).lower()
                if "rate_limit_exceeded" in err_str or "429" in err_str:
                    logger.error(f"❌ Rate limit error: {e}")
                    logger.error("The OpenAI API rate limit was exceeded, and our retry attempt also failed.")
                    logger.error("You may want to try again in a few minutes or with a smaller Excel file.")
                else:
                    logger.error(f"❌ Agent error: {e}")
                sys.exit(1)
            elapsed = time.monotonic() - start
            logger.info(f"✅ Done in {elapsed:.1f}s\n\n📤 {result.final_output}\n")

            # ────────── Optional explicit save ──────────
            if args.output_file:
                try:
                    # Ensure all changes are visible
                    excel_mgr.ensure_changes_applied()
                    
                    # Use the more robust save method
                    saved_path = excel_mgr.save_with_confirmation(args.output_file)
                    logger.info(f"💾 Workbook saved to {saved_path}")
                except Exception as e:
                    logger.error(f"❌ Failed to save workbook: {e}")
                    # Try one more time with a default path
                    try:
                        saved_path = excel_mgr.save_with_confirmation()
                        logger.info(f"💾 Workbook saved to alternative location: {saved_path}")
                    except:
                        logger.error("All save attempts failed.")
                        
    except Exception as e:
        print(f"DEBUG: Error during execution: {e}")
        logger.error(f"❌ Error: {e}")
        sys.exit(1)


# Execute main() when run directly
if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nExiting.")
        sys.exit(0)