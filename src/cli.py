#!/usr/bin/env python
"""
Interactive command-line interface for the Excel AI Assistant.
Starts without an active workbook. Use commands like :open, :new, :close.
"""

import asyncio
import contextlib
import itertools
import logging
import os
import sys
import shlex # For parsing slash commands
from typing import Optional, Any

# Third-party imports
try:
    # Use prompt_toolkit if available for better UX
    from prompt_toolkit import PromptSession
    from prompt_toolkit.history import FileHistory
    from prompt_toolkit.auto_suggest import AutoSuggestFromHistory
    from prompt_toolkit.styles import Style
    PROMPT_TOOLKIT_AVAILABLE = True
except ImportError:
    PROMPT_TOOLKIT_AVAILABLE = False
    logging.getLogger(__name__).info(
        "Optional dependency `prompt_toolkit` not found. Falling back to basic input(). "
        "Install with: pip install prompt_toolkit for a richer CLI experience."
    )

try:
    # Use getch for single-character input (optional, e.g., for step collapse)
    from getch import getch as _sync_getch # Alias to avoid name clash
    GETCH_AVAILABLE = True
except ImportError:
    GETCH_AVAILABLE = False

# Third-party imports (optional)
try:
    from dotenv import load_dotenv
    DOTENV_AVAILABLE = True
except ImportError:
    DOTENV_AVAILABLE = False

# Local project imports
from agents import Runner, Agent
from agents.result import RunResultStreaming, RunResult # Keep RunResult for potential future use
from agents.stream_events import StreamEvent
from agents.exceptions import AgentsException
from .excel_ops import ExcelConnectionError # Added ExcelConnectionError

from .agent_core import excel_assistant_agent
from .context import AppContext, WorkbookShape # Ensure WorkbookShape is imported
from .excel_ops import ExcelManager
from .stream_renderer import format_event # Import the event formatter
from .constants import SHOW_COST

# --- Logging Setup ---
log_level_name = os.getenv("LOG_LEVEL", "INFO").upper()
log_level = getattr(logging, log_level_name, logging.INFO)
log_file = os.getenv("EXCEL_AI_LOG_FILE", "excel_ai.log")

for _h in logging.root.handlers[:]:
    logging.root.removeHandler(_h)

_file_handler = logging.FileHandler(log_file, mode="a", encoding="utf-8")
_file_handler.setLevel(log_level)
_file_handler.setFormatter(
    logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s")
)

_console_handler = logging.StreamHandler(sys.stderr)
_console_handler.setLevel(logging.WARNING) # Default: only warnings/errors to console
_console_handler.setFormatter(
    logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
)

logging.basicConfig(level=log_level, handlers=[_file_handler, _console_handler])
logger = logging.getLogger(__name__)

# --- Configuration ---
HISTORY_FILE = ".excel_ai_history"

# --- CLI Styling (Optional, requires prompt_toolkit) ---
cli_style = Style.from_dict(
    {
        "prompt": "bold cyan",
        "prompt.no-workbook": "bold yellow", # Style for when no workbook is open
        "input": "",
        # Colors handled by stream_renderer ANSI codes now
        # "output.agent": "green",
        # "output.tool": "yellow",
        # "output.error": "bold red",
        "output.info": "cyan",
        "output.warning": "yellow",
        "output.success": "green",
        "output.cost": "italic #888888", # Keep for cost display
        "spinner": "magenta",
    }
)

# --- Helper Functions ---

async def _spinner(prefix="⌛ Thinking", interval=0.2):
    """Simple async spinner using UTF-8 characters."""
    try:
        is_tty = sys.stdout.isatty()
        supports_utf8 = sys.stdout.encoding.lower() == "utf-8"

        if not is_tty or not supports_utf8:
            print(f"{prefix} ...", end="", flush=True)
            while True:
                await asyncio.sleep(interval * 5) # Keep sleeping if not a TTY

        spinner_chars = "|/-\\" # Simpler spinner
        for char in itertools.cycle(spinner_chars):
            # Magenta spinner prefix using ANSI code
            print(f"\r\033[95m{prefix}\033[0m {char} ", end="", flush=True)
            await asyncio.sleep(interval)
    except asyncio.CancelledError:
        # Clean up the spinner line
        clear_len = len(prefix) + 1 + 1 # prefix + char + space
        if sys.stdout.isatty(): # Only clear if TTY
            print("\r" + " " * clear_len + "\r", end="", flush=True)
        else:
            print() # Newline if not TTY
        raise # Re-raise CancelledError
    except Exception as e:
        logger.error(f"Spinner error: {e}", exc_info=True)
        clear_len = len(prefix) + 1 + 1
        if sys.stdout.isatty(): # Only clear if TTY
            print("\r" + " " * clear_len + "\r", end="", flush=True)
        else:
            print() # Newline if not TTY


async def run_agent_streamed(agent: Agent, user_input: str, ctx: AppContext):
    """
    Runs the agent using Runner.run_streamed and renders events live.
    Handles the spinner and event formatting using stream_renderer.py.
    """
    if not ctx.excel_manager or not ctx.excel_manager.book:
        print("\n\033[93m⚠️ No active workbook. Use ':open <path>' or ':new' first.\033[0m")
        return # Cannot run agent without a workbook

    thinking_task = None
    spinner_prefix = "🤖 Thinking"
    result_stream: Optional[RunResultStreaming] = None # Initialize

    try:
        # --- Start Streaming Run ---
        result_stream = Runner.run_streamed(
            agent, input=user_input, context=ctx
        )

        # --- Start Spinner ---
        if sys.stdout.isatty():
            thinking_task = asyncio.create_task(_spinner(prefix=spinner_prefix))
        else:
            print(f"{spinner_prefix}...", flush=True) # Simple message for non-TTY

        # --- Process Events ---
        first_meaningful_event_received = False
        async for ev in result_stream.stream_events():
            logger.debug(f"Received StreamEvent: {ev}") # Log raw event for debugging
            formatted_output = format_event(ev) # Use the renderer
            if formatted_output:
                logger.debug(f"Formatted Output: {formatted_output}") # Log formatted output

                # --- Stop Spinner on First Output ---
                if not first_meaningful_event_received and thinking_task and not thinking_task.done():
                    first_meaningful_event_received = True
                    thinking_task.cancel()
                    with contextlib.suppress(asyncio.CancelledError):
                        await thinking_task # Wait for spinner to clean up

                # --- Print Formatted Event ---
                print(formatted_output)
                sys.stdout.flush() # Ensure output is visible immediately

        # --- Ensure Spinner is Stopped After Loop ---
        if thinking_task and not thinking_task.done():
            thinking_task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await thinking_task

        # --- Handle Final Output (if stream didn't print anything) ---
        # This might happen if the agent finishes without a final message chunk
        # or if the final event wasn't handled by format_event (unlikely with current renderer)
        if not first_meaningful_event_received:
            # Try flushing the buffer one last time
            final_flush = format_event(None) # Pass None to trigger final flush if needed
            if final_flush:
                print(final_flush)
            elif result_stream and result_stream.final_output:
                # If flush didn't work but we have final output, print it minimally
                final_output_str = result_stream.final_output
                if isinstance(final_output_str, str):
                    # Basic agent final output formatting (fallback)
                    prefix = "✔️ 🤖 Agent: "
                    indent = " " * len(prefix)
                    lines = final_output_str.splitlines()
                    formatted_lines = [f"\033[92m{prefix if i == 0 else indent}{line}\033[0m" for i, line in enumerate(lines)]
                    print("\n".join(formatted_lines))
                else:
                    print(f"\033[94m🤔 Run finished. Final output (non-string): {result_stream.final_output}\033[0m")

            else:
                print("\033[94m🤔 Run finished without generating visible streaming output.\033[0m")


        # --- Display Cost ---
        # Cost info should be available in context.state after the run completes
        if SHOW_COST and hasattr(ctx, 'state') and 'last_run_cost' in ctx.state:
            cost = ctx.state.get('last_run_cost', 0.0)
            usage_info = ctx.state.get('last_run_usage', {})
            tokens = usage_info.get('total_tokens', 'N/A')
            # Use ANSI codes for style
            print(f"\033[90m\033[3m💰 Cost: ${cost:.4f} ({tokens} tokens)\033[0m", file=sys.stderr)


    except asyncio.CancelledError:
        if thinking_task and not thinking_task.done():
            thinking_task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await thinking_task
        print("\n\033[93m🚫 Operation cancelled by user (Ctrl+C).\033[0m")
    except AgentsException as e:
        if thinking_task and not thinking_task.done():
            thinking_task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await thinking_task
        print(f"\n\033[91m❌ Agent Error: {e}\033[0m")
        logger.error(f"Agent execution error: {e}", exc_info=True)
    except Exception as e:
        if thinking_task and not thinking_task.done():
            thinking_task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await thinking_task
        print(f"\n\033[91m❌ Unexpected Error: {e}\033[0m")
        logger.error("Unexpected error during agent run", exc_info=True)
    finally:
        # Ensure spinner is always cancelled
        if thinking_task and not thinking_task.done():
            thinking_task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await thinking_task


async def main():
    """Main async function for the CLI."""
    # --- Load Environment Variables (Optional) ---
    if DOTENV_AVAILABLE:
        logger.info("Attempting to load environment variables from .env file...")
        if load_dotenv():
            logger.info(".env file loaded successfully.")
        else:
            logger.info("No .env file found or it was empty.")
    else:
        logger.info("Optional dependency `python-dotenv` not found. Skipping .env file loading. Install with: pip install python-dotenv")

    # --- API Key Check ---
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        print("❌ Error: OPENAI_API_KEY environment variable not set.", file=sys.stderr)
        print("Please set the environment variable before running.", file=sys.stderr)
        sys.exit(1)

    # --- Argument Parsing (Basic) ---
    import argparse
    parser = argparse.ArgumentParser(description="Excel AI Assistant CLI")
    parser.add_argument(
        "file_path", # Keep argument for potential future use or direct open, but don't use on startup
        nargs="?",
        help="Optional: Path to an Excel workbook (not opened automatically).",
        default=None,
    )

    parser.add_argument(
        "--attach",
        action="store_true",
        help="When opening/creating, attempt to attach to an existing running Excel instance.",
    )
    parser.add_argument(
        "--kill-others",
        action="store_true",
        help="When opening/creating, attempt to close other running Excel instances first.",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Show DEBUG log output in the console in addition to the log file.",
    )
    args = parser.parse_args()

    # Adjust console logging based on verbosity
    if args.verbose:
        # Set console handler to DEBUG if verbose is requested
        _console_handler.setLevel(logging.DEBUG)
        # Ensure handler is attached if it wasn't already
        root_logger = logging.getLogger()
        if _console_handler not in root_logger.handlers:
            root_logger.addHandler(_console_handler)
        logger.info("Verbose logging to console enabled.")
    else:
        # Keep console handler at WARNING level (default)
        _console_handler.setLevel(logging.WARNING)
        # Ensure it's attached (might have been removed if toggled)
        root_logger = logging.getLogger()
        if _console_handler not in root_logger.handlers:
            root_logger.addHandler(_console_handler)


    print("\033[1m\033[96m🚀 Excel AI Assistant CLI\033[0m") # Bold Cyan Title
    print("\033[90mType instructions, or use commands: :open <path>, :new, :close, exit\033[0m") # Grey help text
    print("\033[93m⚠️ No workbook loaded. Use :open <path> or :new to start.\033[0m") # Initial warning

    # --- Initialize Context (without Excel initially) ---
    excel_manager: Optional[ExcelManager] = None
    app_context = AppContext(excel_manager=None) # Start with no manager

    # --- Input Loop ---
    if PROMPT_TOOLKIT_AVAILABLE:
        session = PromptSession(
            history=FileHistory(HISTORY_FILE),
            auto_suggest=AutoSuggestFromHistory(),
            # Style will be updated dynamically based on workbook state
        )
        async def get_input(prompt: str, current_style: Style):
            return await session.prompt_async(prompt, style=current_style)
    else:
        # Fallback: Run synchronous input in a thread
        # Style needs ANSI codes for fallback
        async def get_input(prompt: str, current_style: Optional[Style]): # style unused here
            return await asyncio.to_thread(input, prompt)

    while True:
        user_input_str = ""
        try:
            # Determine prompt style based on whether a workbook is open
            prompt_prefix = "💬 User: "
            if PROMPT_TOOLKIT_AVAILABLE:
                # Use the style map directly with prompt_toolkit
                prompt_text = prompt_prefix
                # Determine style key based on workbook state
                style_key = "prompt" if excel_manager else "prompt.no-workbook"
                # We pass the whole style object, prompt_toolkit selects the right key
                current_style = cli_style
            else:
                # Apply ANSI codes for fallback
                prompt_color = "\033[1m\033[96m" if excel_manager else "\033[1m\033[93m" # Cyan if open, Yellow if not
                prompt_text = f"{prompt_color}{prompt_prefix}\033[0m"
                current_style = None # Not used by fallback

            user_input_str = await get_input(prompt_text, current_style)
            user_input_str = user_input_str.strip()

            if not user_input_str:
                continue

            # --- Command Handling ---
            if user_input_str.startswith(":"):
                command_parts = shlex.split(user_input_str[1:])
                command = command_parts[0].lower() if command_parts else ""
                cmd_args = command_parts[1:]

                if command == "open":
                    if not cmd_args:
                        print("\033[91m❌ Usage: :open <file_path.xlsx>\033[0m")
                        continue
                    file_path_to_open = cmd_args[0]
                    # Ensure path exists or provide feedback
                    # if not os.path.exists(file_path_to_open):
                    #     print(f"\033[93m⚠️ File not found: {file_path_to_open}. A new file will be created if possible.\033[0m")

                    print(f"\033[94m🔄 Closing current workbook (if open) and opening '{file_path_to_open}'...\033[0m")
                    if excel_manager:
                        try:
                            await excel_manager.close()
                            excel_manager = None
                            app_context.excel_manager = None
                            app_context.shape = None
                            app_context.state = {} # Reset state on close/open
                            app_context.actions = [] # Reset actions
                        except Exception as e:
                            print(f"\033[91m❌ Error closing previous workbook: {e}\033[0m")
                            logger.error(f"Error closing previous workbook: {e}", exc_info=True)
                            # Continue trying to open the new one

                    try:
                        excel_manager = ExcelManager(
                            file_path=file_path_to_open,
                            visible=True,
                            attach_existing=args.attach,
                            kill_others=args.kill_others
                        )
                        await excel_manager.open()
                        app_context.excel_manager = excel_manager
                        # Perform initial shape scan
                        try:
                            shape_updated = app_context.update_shape(tool_name=":open") # Pass context
                            if shape_updated and app_context.shape: # Check if shape exists
                                print(f"\033[92m✔️ Workbook '{excel_manager.file_path}' opened successfully (Shape v{app_context.shape.version}).\033[0m")
                                logger.info(f"Workbook opened via :open. Path: {excel_manager.file_path}. Shape v{app_context.shape.version}")
                            else:
                                print(f"\033[93m⚠️ Workbook '{excel_manager.file_path}' opened, but could not get initial shape (may indicate connection issue or empty book).\033[0m")
                                logger.warning(f"Workbook opened via :open, but initial shape scan failed or returned empty. Path: {excel_manager.file_path}")
                        except ExcelConnectionError as ce:
                            print(f"\033[91m❌ Workbook '{excel_manager.file_path}' opened, but connection failed during shape scan: {ce}\033[0m")
                            logger.error(f"Connection error during shape scan after opening workbook '{excel_manager.file_path}': {ce}")
                            # Attempt to close the potentially problematic manager
                            try: await excel_manager.close()
                            except: pass
                            excel_manager = None # Ensure manager is None on connection failure
                            app_context.excel_manager = None
                            app_context.shape = None


                    except ExcelConnectionError as ce:
                        # Catch connection error during the manager.open() call itself
                        print(f"\033[91m❌ Failed to establish connection for workbook '{file_path_to_open}': {ce}\033[0m")
                        logger.error(f"Connection error during ExcelManager.open for '{file_path_to_open}': {ce}", exc_info=True)
                        excel_manager = None
                        app_context.excel_manager = None
                        app_context.shape = None
                    except Exception as e:
                        print(f"\033[91m❌ Error opening workbook '{file_path_to_open}': {e}\033[0m")
                        logger.error(f"Error opening workbook '{file_path_to_open}': {e}", exc_info=True)
                        excel_manager = None # Ensure manager is None on failure
                        app_context.excel_manager = None
                        app_context.shape = None

                elif command == "new":
                    print("\033[94m🔄 Closing current workbook (if open) and creating a new one...\033[0m")
                    if excel_manager:
                        try:
                            await excel_manager.close()
                            excel_manager = None
                            app_context.excel_manager = None
                            app_context.shape = None
                            app_context.state = {} # Reset state
                            app_context.actions = [] # Reset actions
                        except Exception as e:
                            print(f"\033[91m❌ Error closing previous workbook: {e}\033[0m")
                            logger.error(f"Error closing previous workbook: {e}", exc_info=True)
                            # Continue trying to open the new one

                    try:
                        excel_manager = ExcelManager(
                            file_path=None, # No initial path for new
                            visible=True,
                            attach_existing=args.attach,
                            kill_others=args.kill_others
                        )
                        await excel_manager.open()
                        app_context.excel_manager = excel_manager
                        # Perform initial shape scan
                        try:
                            shape_updated = app_context.update_shape(tool_name=":new") # Pass context
                            if shape_updated and app_context.shape: # Check if shape exists
                                print(f"\033[92m✔️ New workbook '{excel_manager.file_path}' created successfully (Shape v{app_context.shape.version}).\033[0m")
                                logger.info(f"New workbook created via :new. Path: {excel_manager.file_path}. Shape v{app_context.shape.version}")
                            else:
                                print(f"\033[93m⚠️ New workbook '{excel_manager.file_path}' created, but could not get initial shape (may indicate connection issue or empty book).\033[0m")
                                logger.warning(f"New workbook created via :new, but initial shape scan failed or returned empty. Path: {excel_manager.file_path}")
                        except ExcelConnectionError as ce:
                            print(f"\033[91m❌ New workbook created, but connection failed during shape scan: {ce}\033[0m")
                            logger.error(f"Connection error during shape scan after creating workbook: {ce}")
                            # Attempt to close the potentially problematic manager
                            try: await excel_manager.close()
                            except: pass
                            excel_manager = None # Ensure manager is None on connection failure
                            app_context.excel_manager = None
                            app_context.shape = None

                    except ExcelConnectionError as ce:
                        # Catch connection error during the manager.open() call itself
                        print(f"\033[91m❌ Failed to establish connection for new workbook: {ce}\033[0m")
                        logger.error(f"Connection error during ExcelManager.open for ':new': {ce}", exc_info=True)
                        excel_manager = None
                        app_context.excel_manager = None
                        app_context.shape = None
                    except Exception as e:
                        print(f"\033[91m❌ Error creating new workbook: {e}\033[0m")
                        logger.error(f"Error creating new workbook: {e}", exc_info=True)
                        excel_manager = None # Ensure manager is None on failure
                        app_context.excel_manager = None
                        app_context.shape = None

                elif command == "close":
                    if excel_manager:
                        print("\033[94m🔄 Closing current workbook...\033[0m")
                        try:
                            await excel_manager.close()
                            print("\033[92m✔️ Workbook closed.\033[0m")
                            logger.info("Workbook closed via :close command.")
                        except Exception as e:
                            print(f"\033[91m❌ Error closing workbook: {e}\033[0m")
                            logger.error(f"Error closing workbook via :close: {e}", exc_info=True)
                        finally:
                            excel_manager = None
                            app_context.excel_manager = None
                            app_context.shape = None
                            app_context.state = {} # Reset state
                            app_context.actions = [] # Reset actions
                            print("\033[93m⚠️ No workbook loaded. Use :open <path> or :new to start.\033[0m")
                    else:
                        print("\033[93m⚠️ No workbook is currently open.\033[0m")

                elif command.lower() in ["exit", "quit"]:
                    break # Exit CLI loop

                elif command == "clear":
                    # Basic clear screen (might not work on all terminals)
                    print("\033[H\033[J", end="")

                elif command == "help":
                    print("\nAvailable commands:")
                    print("  :open <path>  - Close current workbook and open/create one at <path>.")
                    print("  :new          - Close current workbook and create a new blank one.")
                    print("  :close        - Close the current workbook.")
                    print("  :shape        - Show the current workbook structure known to the agent.")
                    print("  :clear        - Clear the terminal screen.")
                    print("  :help         - Show this help message.")
                    print("  exit / quit   - Exit the CLI.")
                    print("Enter Excel instructions directly otherwise.")

                elif command == "shape":
                    if app_context.shape:
                        from .agent_core import _format_workbook_shape # Use the formatter
                        shape_str = _format_workbook_shape(app_context.shape)
                        print("\n\033[94mCurrent Workbook Shape:\033[0m")
                        print(shape_str)
                    elif app_context.excel_manager:
                        print("\033[93m⚠️ Workbook is open, but shape information is not available (try running an instruction or check logs).\033[0m")
                    else:
                        print("\033[93m⚠️ No workbook open to show shape.\033[0m")

                else:
                    print(f"\033[91m❌ Unknown command: ':{command}'. Type ':help' for options.\033[0m")

            # --- Regular Instruction Handling ---
            elif user_input_str.lower() in ["exit", "quit"]:
                break
            else:
                # Check if workbook is open before running agent
                if not excel_manager or not excel_manager.book:
                    print("\033[93m⚠️ Please open or create a workbook first using ':open <path>' or ':new'.\033[0m")
                    continue

                # Run the agent with the user input using the streaming function
                await run_agent_streamed(
                    excel_assistant_agent,
                    user_input_str,
                    app_context
                )

        except EOFError: # Handle Ctrl+D
            break
        except KeyboardInterrupt: # Handle Ctrl+C during input prompt
            print("\n\033[93mUse 'exit' or 'quit' to leave. (Ctrl+C cancels current operation)\033[0m")
            continue # Continue loop after Ctrl+C during input
        except Exception as e:
            print(f"\n\033[91m❌ Error in CLI loop: {e}\033[0m")
            logger.error("Error in main CLI loop", exc_info=True)
            # Optional: break or continue

    # --- Cleanup ---
    if excel_manager:
        print("\n\033[94m👋 Closing active workbook before exiting...\033[0m")
        try:
            await excel_manager.close()
        except Exception as e:
            logger.error(f"Error during final cleanup close: {e}", exc_info=True)

    print("\n\033[96m👋 Exiting Excel AI Assistant. Goodbye!\033[0m")


# --- Entry Point ---
if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n\033[93m🚫 Exiting due to user interrupt.\033[0m")
        sys.exit(0)
    except Exception as e:
        logger.critical(f"CLI critical error during startup/shutdown: {e}", exc_info=True)
        print(f"\n\033[91m❌ Critical Error: {e}\033[0m", file=sys.stderr)
        sys.exit(1)