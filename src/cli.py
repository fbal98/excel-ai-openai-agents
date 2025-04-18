"""Command‑line interface for the Autonomous Excel Assistant."""

import argparse
import asyncio
import os
import logging
import sys
import time
from typing import Dict, Any, Optional

from dotenv import load_dotenv
from agents import Runner, RunResultStreaming, StreamEvent
from agents.stream_events import RunItemStreamEvent, RawResponsesStreamEvent
from openai.types.responses import ResponseTextDeltaEvent

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
        required=False,
        default=None,
        help="Instruction for the agent (natural language)",
    )
    parser.add_argument(
        "--interactive", "-i",
        action="store_true",
        help="Start interactive chat mode (conversational CLI)",
    )
    parser.add_argument(
        "--live",
        action="store_true",
        help="Edit the workbook in‑process via xlwings so changes appear in real time.",
    )
    parser.add_argument(
        "-v", "--verbose",
        action="store_true",
        help="Enable verbose logging, including debug messages."
    )
    parser.add_argument(
        "--stream",
        action="store_true",
        help="Stream the agent's responses and thought process in real time.",
    )
    return parser.parse_args()


async def handle_streaming(result: RunResultStreaming, verbose: bool) -> Any:
    """Process streaming results from the agent."""
    logger = logging.getLogger(__name__)
    logger.info("🔄 Streaming mode enabled. Showing agent progress in real time...\n")
    
    # Keep track of the last tool call for better logging
    last_tool: Optional[str] = None
    final_output = ""
    last_message = None
    
    try:
        async for event in result.stream_events():
            if event.type == "raw_response_event" and isinstance(event.data, ResponseTextDeltaEvent):
                # Collect the final output regardless of verbose mode
                final_output += event.data.delta
                
                # Show raw token output only in verbose mode
                if verbose:
                    print(event.data.delta, end="", flush=True)
            
            elif event.type == "run_item_stream_event":
                item_event = event  # type: RunItemStreamEvent
                item = item_event.item
                
                if item.type == "tool_call_item":
                    # Get tool information in a more flexible way
                    tool_name = "unknown tool"
                    
                    # Try different ways to access the tool info based on API version
                    try:
                        if hasattr(item, 'function_name'):
                            tool_name = item.function_name
                        elif hasattr(item, 'name'):
                            tool_name = item.name
                        elif hasattr(item, 'function') and isinstance(item.function, dict):
                            tool_name = item.function.get('name', 'unknown tool')
                            
                        # For debug purposes in verbose mode only
                        if verbose and tool_name == "unknown tool":
                            logger.debug(f"Debug tool item: {dir(item)}")
                            
                        # Try to get arguments
                        args_str = ""
                        
                        if hasattr(item, 'arguments') and item.arguments:
                            if isinstance(item.arguments, dict):
                                args_str = ", ".join([f"{k}={v}" for k, v in item.arguments.items()])
                            else:
                                args_str = str(item.arguments)
                        elif hasattr(item, 'args') and item.args:
                            if isinstance(item.args, dict):
                                args_str = ", ".join([f"{k}={v}" for k, v in item.args.items()])
                            else:
                                args_str = str(item.args)
                        elif hasattr(item, 'function') and isinstance(item.function, dict) and 'arguments' in item.function:
                            args = item.function.get('arguments')
                            if isinstance(args, dict):
                                args_str = ", ".join([f"{k}={v}" for k, v in args.items()])
                            else:
                                args_str = str(args)
                            
                        last_tool = tool_name
                        
                        # Log the tool call
                        if args_str:
                            logger.info(f"🛠️  [TOOL] {tool_name}: {args_str}")
                        else:
                            logger.info(f"🛠️  [TOOL] {tool_name}")
                    except Exception as e:
                        if verbose:
                            logger.debug(f"Error parsing tool call: {e}")
                        logger.info(f"🛠️  Tool called")
                
                elif item.type == "tool_call_output_item":
                    output = "unknown output"
                    try:
                        if hasattr(item, 'output'):
                            output = item.output
                        elif hasattr(item, 'content'):
                            output = item.content
                            
                        if verbose:
                            if last_tool:
                                logger.info(f"📊 Result from {last_tool}: {output}")
                            else:
                                logger.info(f"📊 Tool result: {output}")
                    except Exception as e:
                        if verbose:
                            logger.debug(f"Error parsing tool output: {e}")
                
                elif item.type == "message_output_item":
                    try:
                        # Get the message content
                        content = ""
                        if hasattr(item, 'text'):
                            content = item.text
                        elif hasattr(item, 'content'):
                            content = item.content
                        
                        # Save this as the potential final message
                        if content:
                            last_message = content
                            if verbose:
                                logger.info(f"💭 Agent: {content}")
                    except Exception as e:
                        if verbose:
                            logger.debug(f"Error parsing message: {e}")
    except Exception as e:
        logger.error(f"Error processing stream events: {e}")
        if verbose:
            import traceback
            logger.error(traceback.format_exc())
    
    # Prefer the last message content if we captured it
    result_output = last_message or final_output or "Task completed."
    
    # The streaming API doesn't have get_result, so use a class that mimics the result
    class SimpleResult:
        def __init__(self, final_output):
            self.final_output = final_output
            
    return SimpleResult(result_output)


async def main() -> None:
    load_dotenv()
    if not os.getenv("OPENAI_API_KEY"):
        raise RuntimeError("OPENAI_API_KEY not set in environment or .env file.")

    args = parse_args()

    # Configure logging
    log_level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=log_level, format="%(message)s")

    # ──────────────────────────────────────────────────────────────────
    #  Suppress noisy HTTP request logs from OpenAI/HTTPX libraries
    #  so the console shows only our [TOOL] lines and key agent output
    # ──────────────────────────────────────────────────────────────────
    for noisy in (
        "openai",             # OpenAI client logger
        "openai._client",     # low‑level client
        "openai._requester",  # HTTP requester
        "httpx",              # transport used by OpenAI SDK
        "httpcore",           # lower‑level HTTP core
        "httpcore.http11",    # HTTP/1.1 wire logs
    ):
        logging.getLogger(noisy).setLevel(logging.WARNING)

    logger = logging.getLogger(__name__)
    start_time = time.monotonic()

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
    # Inform about workbook
    if args.input_file:
        logger.info("📂 Loaded workbook: %s", args.input_file)
    else:
        logger.info("🆕 Created new workbook")
    if args.live:
        logger.info("📊 Live mode enabled. Changes will appear in Excel in real time.")

    # Interactive chat mode
    if args.interactive:
        print("Hello! How can I help you today?")
        chat_history = []
        while True:
            try:
                user_input = input("> ")
            except (EOFError, KeyboardInterrupt):
                print("\nExiting interactive mode.")
                break
            if user_input.strip().lower() in ("exit", "quit"):
                print("Exiting interactive mode.")
                break
            chat_history.append({"role": "user", "content": user_input})
            result = await Runner.run(
                excel_assistant_agent,
                input=chat_history,
                context=app_context,
                max_turns=25,
            )
            reply = result.final_output
            print(reply)
            chat_history.append({"role": "assistant", "content": reply})
        sys.exit(0)

    logger.info("\n💡 Instruction: %s", args.instruction)
    logger.info("🤖 Running agent (live=%s)...", args.live)
    
    try:
        if args.stream:
            # Use streaming mode
            result_streaming = Runner.run_streamed(
                excel_assistant_agent,
                input=args.instruction,
                context=app_context,
                max_turns=25,
            )
            result = await handle_streaming(result_streaming, args.verbose)
        else:
            # Use regular mode
            result = await Runner.run(
                excel_assistant_agent,
                input=args.instruction,
                context=app_context,
                max_turns=25,
            )
    except Exception as e:
        logger.error("❌ Agent error: %s", e)
        sys.exit(1)
        
    elapsed = time.monotonic() - start_time
    logger.info("✅ Agent completed in %.1f seconds.", elapsed)
    logger.info("\n📤 Final Output:\n%s\n", result.final_output)

    # ------------------------------------------------------------------ #
    #  Persist workbook if not in live mode                              #
    # ------------------------------------------------------------------ #
    if not args.live:
        if not args.output_file:
            logger.error("--output-file is required in batch mode.")
            sys.exit(1)
        try:
            excel_manager.save_workbook(args.output_file)
            logger.info("📁 Workbook saved to %s", args.output_file)
        except Exception as exc:
            logger.error("❌ Failed to save workbook: %s", exc)
            sys.exit(1)

    if args.verbose:
        # Show full result details in verbose mode
        logger.debug("Agent full result: %r", result)


if __name__ == "__main__":
    import sys
    try:
        asyncio.run(main())
    except Exception as e:
        # Top-level error handling
        print(f"❌ {e}", file=sys.stderr)
        sys.exit(1)