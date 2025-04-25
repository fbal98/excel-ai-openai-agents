#!/usr/bin/env python
"""
Interactive command-line interface for the Excel AI Assistant.
Starts without an active workbook. Use commands like :open, :new, :close.
Supports switching LLM providers via the :provider command.
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
    from getch import getch as _sync_getch # Alias to avoid name clash
    GETCH_AVAILABLE = True
except ImportError:
    GETCH_AVAILABLE = False

try:
    from dotenv import load_dotenv
    DOTENV_AVAILABLE = True
except ImportError:
    DOTENV_AVAILABLE = False

# Local project imports
from agents import Runner, Agent
from agents.result import RunResultStreaming # Keep RunResult for potential future use
from agents.stream_events import StreamEvent
from agents.exceptions import AgentsException
from agents import RunContextWrapper # Import RunContextWrapper

from .excel_ops import ExcelConnectionError, ExcelManager
from .model_config import set_active_provider, get_active_provider, list_available_providers # Import provider functions
from .model_integration import create_excel_assistant_agent # Import agent factory function
from .context import AppContext, WorkbookShape
from .stream_renderer import format_event
from .constants import SHOW_COST
from .agent_core import _dynamic_instructions, _format_workbook_shape # Import necessary functions from agent_core

# --- Logging Setup ---
log_level_name = os.getenv("LOG_LEVEL", "INFO").upper()
log_level = getattr(logging, log_level_name, logging.INFO)
log_file = os.getenv("EXCEL_AI_LOG_FILE", "excel_ai.log")

# Clear existing root handlers to avoid duplicates if re-run
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
    logging.Formatter("%(levelname)s: %(message)s") # Simpler console format
)

logging.basicConfig(level=log_level, handlers=[_file_handler, _console_handler])
logger = logging.getLogger(__name__) # Get logger for this module

# --- Configuration ---
HISTORY_FILE = ".excel_ai_history"

# --- CLI Styling (Optional, requires prompt_toolkit) ---
if PROMPT_TOOLKIT_AVAILABLE:
    cli_style = Style.from_dict(
    {
        "prompt": "bold cyan",
        "prompt.no-workbook": "bold yellow",
        "input": "",
        "output.info": "cyan",
        "output.warning": "yellow",
        "output.success": "green",
        "output.cost": "italic #888888", # Grey italics for cost
        "spinner": "magenta",
        "error": "bold red", # For error messages
    }
)

# --- Helper Functions ---

async def _spinner(prefix="⌛ Thinking", interval=0.2):
    """Simple async spinner using UTF-8 characters."""
    spinner_task = None
    try:
        is_tty = sys.stdout.isatty()
        supports_utf8 = sys.stdout.encoding.lower() == "utf-8"

        if not is_tty or not supports_utf8:
            print(f"{prefix} ...", end="", flush=True)
            while True:
                await asyncio.sleep(interval * 5)

        spinner_chars = "|/-\\"
        for char in itertools.cycle(spinner_chars):
            # Magenta spinner prefix using ANSI code
            print(f"\r\033[95m{prefix}\033[0m {char} ", end="", flush=True)
            await asyncio.sleep(interval)
    except asyncio.CancelledError:
        # Clean up the spinner line
        clear_len = len(prefix) + 1 + 1 # prefix + char + space
        if sys.stdout.isatty():
            print("\r" + " " * clear_len + "\r", end="", flush=True)
        else:
            print() # Newline if not TTY
        # No need to re-raise here, cancellation is handled by caller
    except Exception as e:
        logger.error(f"Spinner error: {e}", exc_info=True)
        clear_len = len(prefix) + 1 + 1
        if sys.stdout.isatty():
            print("\r" + " " * clear_len + "\r", end="", flush=True)
        else:
            print()


async def run_agent_streamed(agent: Agent, user_input: str, ctx: AppContext):
    """
    Runs the agent using Runner.run_streamed and renders events live.
    Handles the spinner and event formatting using stream_renderer.py.
    Displays cost information after the run if enabled.
    """
    if not ctx.excel_manager or not ctx.excel_manager.book:
        print("\n\033[93m⚠️ No active workbook. Use ':open <path>' or ':new' first.\033[0m")
        return

    thinking_task = None
    spinner_prefix = f"🤖 {get_active_provider().capitalize()} Thinking" # Show current provider
    result_stream: Optional[RunResultStreaming] = None
    first_meaningful_event_received = False

    logger.info(f"--- Running Agent ---")
    logger.info(f"Input: '{user_input}'")
    logger.info(f"Context State BEFORE run: {ctx.state}")
    logger.info(f"Context Shape BEFORE run: v{ctx.shape.version if ctx.shape else 'N/A'}")
    # Log last few actions if needed: logger.info(f"Context Actions BEFORE run: {ctx.actions[-3:]}")

    logger.info(f"--- Running Agent ---")
    logger.info(f"Input: '{user_input}'")
    logger.info(f"Context State BEFORE run: {ctx.state}")
    logger.info(f"Context Shape BEFORE run: v{ctx.shape.version if ctx.shape else 'N/A'}")
    # Log last few actions if needed: logger.info(f"Context Actions BEFORE run: {ctx.actions[-3:]}")

    try:
        # --- Start Streaming Run ---
        result_stream = Runner.run_streamed(
            agent, input=user_input, context=ctx
        )

        # --- Start Spinner ---
        if sys.stdout.isatty():
            thinking_task = asyncio.create_task(_spinner(prefix=spinner_prefix))
        else:
            print(f"{spinner_prefix}...", flush=True)

        # --- Process Events ---
        async for ev in result_stream.stream_events():
            logger.debug(f"Raw StreamEvent: {ev}")
            formatted_output = format_event(ev)
            if formatted_output:
                logger.debug(f"Formatted Output: {formatted_output}")

                # --- Stop Spinner on First Meaningful Output ---
                if not first_meaningful_event_received and thinking_task and not thinking_task.done():
                    first_meaningful_event_received = True
                    thinking_task.cancel()
                    with contextlib.suppress(asyncio.CancelledError):
                        await thinking_task # Wait for spinner cleanup

                # --- Print Formatted Event ---
                print(formatted_output, end='') # format_event usually includes newline
                sys.stdout.flush()

        # --- Ensure Spinner is Stopped After Loop ---
        if thinking_task and not thinking_task.done():
            thinking_task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await thinking_task

        # --- Handle Final Output (if stream didn't print anything meaningful) ---
        if not first_meaningful_event_received:
            final_flush = format_event(None) # Trigger final flush
            if final_flush:
                print(final_flush, end='')
                sys.stdout.flush()
            elif result_stream and result_stream.final_output:
                 final_output_str = result_stream.final_output
                 if isinstance(final_output_str, str):
                     prefix = "✔️ 🤖 Agent: "
                     indent = " " * len(prefix)
                     lines = final_output_str.strip().splitlines()
                     formatted_lines = [f"\033[92m{prefix if i == 0 else indent}{line}\033[0m" for i, line in enumerate(lines)]
                     print("\n".join(formatted_lines)) # Print with newline formatting
                 else:
                     print(f"\n\033[94m🤔 Run finished. Final output (non-string): {result_stream.final_output}\033[0m")
            else:
                 print("\n\033[94m🤔 Run finished without generating visible streaming output.\033[0m")


        # --- After run completes (before cost display) ---
        logger.info(f"--- Agent Run Finished ---")
        logger.info(f"Context State AFTER run: {ctx.state}")
        logger.info(f"Context Shape AFTER run: v{ctx.shape.version if ctx.shape else 'N/A'}")
        logger.info(f"Context Actions AFTER run: {ctx.actions[-5:]}") # Log last 5 actions

        # --- After run completes (before cost display) ---
        logger.info(f"--- Agent Run Finished ---")
        logger.info(f"Context State AFTER run: {ctx.state}")
        logger.info(f"Context Shape AFTER run: v{ctx.shape.version if ctx.shape else 'N/A'}")
        logger.info(f"Context Actions AFTER run: {ctx.actions[-5:]}") # Log last 5 actions

        # --- Display Cost ---
        # Cost info should be available in context.state after run_and_cost completes
        if SHOW_COST and hasattr(ctx, 'state') and 'last_run_cost' in ctx.state:
            cost = ctx.state.get('last_run_cost', 0.0)
            usage_info = ctx.state.get('last_run_usage', {})
            tokens = usage_info.get('total_tokens', 'N/A')
            model_used = usage_info.get('model_name', 'N/A')
            cost_style_prefix = "\033[90m\033[3m" # Grey italic ANSI codes
            cost_style_suffix = "\033[0m"
            if PROMPT_TOOLKIT_AVAILABLE and cli_style:
                 # Experimental: Try applying style class via print if possible (might need Rich)
                 print(f"{cost_style_prefix}💰 Cost: ${cost:.4f} ({tokens} tokens, Model: {model_used}){cost_style_suffix}", file=sys.stderr)
            else:
                 print(f"{cost_style_prefix}💰 Cost: ${cost:.4f} ({tokens} tokens, Model: {model_used}){cost_style_suffix}", file=sys.stderr)


    except asyncio.CancelledError:
        # This is expected if the user hits Ctrl+C during the run
        if thinking_task and not thinking_task.done():
             thinking_task.cancel()
             with contextlib.suppress(asyncio.CancelledError): await thinking_task
        print("\n\033[93m🚫 Agent run cancelled by user (Ctrl+C).\033[0m")
        # Do not re-raise, let the main loop handle Ctrl+C during input
    except AgentsException as e:
        if thinking_task and not thinking_task.done():
            thinking_task.cancel()
            with contextlib.suppress(asyncio.CancelledError): await thinking_task
        print(f"\n\033[91m❌ Agent Error: {e}\033[0m") # Red error ANSI
        logger.error(f"Agent execution error: {e}", exc_info=True)
    except ExcelConnectionError as e: # Catch specific Excel errors
        if thinking_task and not thinking_task.done():
            thinking_task.cancel()
            with contextlib.suppress(asyncio.CancelledError): await thinking_task
        print(f"\n\033[91m❌ Excel Connection Error: {e}\033[0m")
        logger.error(f"Excel connection error during agent run: {e}", exc_info=True)
    except Exception as e:
        if thinking_task and not thinking_task.done():
            thinking_task.cancel()
            with contextlib.suppress(asyncio.CancelledError): await thinking_task
        print(f"\n\033[91m❌ Unexpected Error during agent run: {e}\033[0m")
        logger.error("Unexpected error during agent run", exc_info=True)
    finally:
        # Ensure spinner is always cancelled cleanly
        if thinking_task and not thinking_task.done():
            thinking_task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await thinking_task


async def main():
    """Main async function for the CLI."""

    # --- Load Environment Variables ---
    if DOTENV_AVAILABLE:
        logger.info("Attempting to load environment variables from .env file...")
        if load_dotenv(override=True): # override=True ensures .env takes precedence
            logger.info(".env file loaded successfully.")
            # Re-evaluate active provider after loading .env
            global _active_provider
            from .model_config import _active_provider as current_provider_state # Access internal state carefully
            _active_provider = os.getenv("DEFAULT_MODEL_PROVIDER", "openai").lower() # Re-read from env
            logger.info(f"Re-evaluated active provider from .env: {_active_provider}")
        else:
            logger.info("No .env file found or it was empty.")
    else:
        logger.info("Optional dependency `python-dotenv` not found. Skipping .env file loading.")

    # --- Basic Sanity Check (Check if default provider is configured) ---
    # This check runs after .env load attempt
    from .model_config import _is_provider_configured # Import checker
    initial_provider = get_active_provider()
    if not _is_provider_configured(initial_provider):
        print(f"\n\033[93m⚠️ Warning: Default provider '{initial_provider}' is not fully configured.\033[0m")
        print(f"\033[93m   Check {initial_provider.upper()}_API_KEY and {initial_provider.upper()}_MODEL environment variables.\033[0m")
        print("\033[93m   You may need to use the ':provider' command to switch to a configured provider.\033[0m")
        # Allow continuing, but agent creation might fail


    # --- Argument Parsing ---
    import argparse
    parser = argparse.ArgumentParser(description="Excel AI Assistant CLI")
    parser.add_argument("file_path", nargs="?", help="Optional: Path to an Excel workbook (not opened automatically).", default=None,)
    parser.add_argument("--attach", action="store_true", help="When opening/creating, attempt to attach to an existing running Excel instance.")
    parser.add_argument("--kill-others", action="store_true", help="When opening/creating, attempt to close other running Excel instances first.")
    parser.add_argument("--verbose", "-v", action="store_true", help="Show DEBUG log output in the console.")
    args = parser.parse_args()

    # Adjust console logging based on verbosity
    if args.verbose:
        _console_handler.setLevel(logging.DEBUG)
        logger.info("Verbose logging to console enabled.")
    else:
        _console_handler.setLevel(logging.WARNING)

    # --- Initialize Agent using factory ---
    # Do this *after* parsing args and setting log levels
    try:
        excel_assistant_agent = create_excel_assistant_agent() # Use the factory
        logger.info(f"Initial agent instance created for provider: {get_active_provider()}")
    except Exception as e:
        print(f"\n\033[91m❌ Critical Error initializing agent: {e}\033[0m", file=sys.stderr)
        logger.critical(f"Failed to create initial agent: {e}", exc_info=True)
        # Ask user if they want to configure provider or exit? For now, just exit.
        print("\n\033[93m   Please check your environment variables for the active provider ('{get_active_provider()}') or try switching providers using the ':provider' command after starting.\033[0m")
        sys.exit(1)


    print("\n\033[1m\033[96m🚀 Excel AI Assistant CLI\033[0m") # Bold Cyan Title
    print(f"\033[90mType Excel instructions, or use commands (:help for list). Provider: \033[1m{get_active_provider()}\033[0m") # Grey help text
    print("\033[93m⚠️ No workbook loaded. Use :open <path> or :new to start.\033[0m") # Initial warning

    # --- Initialize Context (without Excel initially) ---
    excel_manager: Optional[ExcelManager] = None
    app_context = AppContext(excel_manager=None) # Start with no manager

    # --- Input Loop ---
    if PROMPT_TOOLKIT_AVAILABLE:
        session = PromptSession(
            history=FileHistory(HISTORY_FILE),
            auto_suggest=AutoSuggestFromHistory(),
        )
        async def get_input(prompt: str, current_style: Style):
            return await session.prompt_async(prompt, style=current_style)
    else:
        async def get_input(prompt: str, current_style: Optional[Style]):
            return await asyncio.to_thread(input, prompt)

    while True:
        user_input_str = ""
        try:
            # Determine prompt style based on workbook state and provider
            prompt_prefix = f"💬 ({get_active_provider()}) User: " # Show provider in prompt
            if PROMPT_TOOLKIT_AVAILABLE:
                style_key = "prompt" if excel_manager else "prompt.no-workbook"
                # Pass the whole style object, prompt_toolkit selects the right key based on class
                current_style = cli_style
                # Construct prompt text without explicit ANSI codes for prompt_toolkit
                prompt_text = prompt_prefix
            else:
                # Apply ANSI codes for fallback
                prompt_color = "\033[1m\033[96m" if excel_manager else "\033[1m\033[93m" # Cyan if open, Yellow if not
                prompt_text = f"{prompt_color}{prompt_prefix}\033[0m"
                current_style = None # Not used by fallback input()

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
                    print(f"\033[94m🔄 Closing current workbook (if open) and opening '{file_path_to_open}'...\033[0m")
                    if excel_manager:
                        try:
                            await excel_manager.close()
                        except Exception as e:
                            logger.error(f"Error closing previous workbook during :open: {e}", exc_info=True)
                        finally: # Ensure cleanup regardless of close success
                            excel_manager = None
                            app_context.excel_manager = None
                            app_context.shape = None
                            app_context.state = {}
                            app_context.actions = []
                    try:
                        excel_manager = ExcelManager(file_path=file_path_to_open, visible=True, attach_existing=args.attach, kill_others=args.kill_others)
                        await excel_manager.open()
                        app_context.excel_manager = excel_manager
                        shape_updated = app_context.update_shape(tool_name=":open") # Update shape and context
                        if shape_updated and app_context.shape:
                            print(f"\033[92m✔️ Workbook '{excel_manager.file_path}' opened (Shape v{app_context.shape.version}).\033[0m")
                        else:
                            print(f"\033[93m⚠️ Workbook '{excel_manager.file_path}' opened, but shape scan failed or empty.\033[0m")
                    except (ExcelConnectionError, Exception) as e:
                        print(f"\033[91m❌ Error opening workbook '{file_path_to_open}': {e}\033[0m")
                        logger.error(f"Error during :open '{file_path_to_open}': {e}", exc_info=True)
                        excel_manager = None # Ensure cleanup on failure
                        app_context.excel_manager = None
                        app_context.shape = None

                elif command == "new":
                    print("\033[94m🔄 Closing current workbook (if open) and creating a new one...\033[0m")
                    if excel_manager:
                         try: await excel_manager.close()
                         except Exception as e: logger.error(f"Error closing previous workbook during :new: {e}", exc_info=True)
                         finally:
                            excel_manager = None
                            app_context.excel_manager = None
                            app_context.shape = None
                            app_context.state = {}
                            app_context.actions = []
                    try:
                        excel_manager = ExcelManager(file_path=None, visible=True, attach_existing=args.attach, kill_others=args.kill_others)
                        await excel_manager.open()
                        app_context.excel_manager = excel_manager
                        shape_updated = app_context.update_shape(tool_name=":new")
                        if shape_updated and app_context.shape:
                            print(f"\033[92m✔️ New workbook '{excel_manager.file_path}' created (Shape v{app_context.shape.version}).\033[0m")
                        else:
                            print(f"\033[93m⚠️ New workbook '{excel_manager.file_path}' created, but shape scan failed or empty.\033[0m")
                    except (ExcelConnectionError, Exception) as e:
                        print(f"\033[91m❌ Error creating new workbook: {e}\033[0m")
                        logger.error(f"Error during :new: {e}", exc_info=True)
                        excel_manager = None
                        app_context.excel_manager = None
                        app_context.shape = None

                elif command == "close":
                    if excel_manager:
                        print("\033[94m🔄 Closing current workbook...\033[0m")
                        try:
                            await excel_manager.close()
                            print("\033[92m✔️ Workbook closed.\033[0m")
                        except Exception as e:
                            print(f"\033[91m❌ Error closing workbook: {e}\033[0m")
                            logger.error(f"Error closing workbook via :close: {e}", exc_info=True)
                        finally: # Always clean up context state
                            excel_manager = None
                            app_context.excel_manager = None
                            app_context.shape = None
                            app_context.state = {}
                            app_context.actions = []
                            print("\033[93m⚠️ No workbook loaded. Use :open <path> or :new to start.\033[0m")
                    else:
                        print("\033[93m⚠️ No workbook is currently open.\033[0m")

                elif command.lower() in ["exit", "quit"]:
                    print("\n\033[1m\033[95m--- Current System Prompt (for active agent) ---\033[0m")
                    try:
                        temp_wrapper = RunContextWrapper(context=app_context)
                        # Use the currently active agent instance
                        current_prompt = _dynamic_instructions(temp_wrapper, excel_assistant_agent)
                        print(current_prompt)
                    except Exception as e:
                        print(f"\033[91mError generating system prompt: {e}\033[0m")
                    print("\033[1m\033[95m--- End System Prompt ---\033[0m\n")
                    break # Exit CLI loop

                elif command == "clear":
                    print("\033[H\033[J", end="") # Basic clear screen

                elif command == "help":
                    print("\nAvailable commands:")
                    print("  :open <path>    - Close current workbook and open/create one at <path>.")
                    print("  :new            - Close current workbook and create a new blank one.")
                    print("  :close          - Close the current workbook.")
                    print("  :shape          - Show the current workbook structure known to the agent.")
                    print("  :provider [name]- Switch LLM provider (e.g., :provider gemini) or show current/available.")
                    print("  :clear          - Clear the terminal screen.")
                    print("  :help           - Show this help message.")
                    print("  exit / quit     - Exit the CLI.")
                    print("\nEnter Excel instructions directly otherwise.")

                elif command == "shape":
                    if app_context.shape:
                        shape_str = _format_workbook_shape(app_context.shape)
                        print("\n\033[94mCurrent Workbook Shape (as seen by agent):\033[0m")
                        print(shape_str)
                    elif app_context.excel_manager:
                        print("\033[93m⚠️ Workbook open, but shape info missing (try an instruction? Check logs).\033[0m")
                    else:
                        print("\033[93m⚠️ No workbook open to show shape.\033[0m")

                elif command == "provider":
                    # Handle provider command
                    if not cmd_args:
                        # Show current provider and available providers
                        current = get_active_provider()
                        providers = list_available_providers()
                        print(f"\n\033[94mCurrent model provider: \033[1m{current}\033[0m")
                        print("Available providers (Checked from .env):")
                        for provider, configured in providers.items():
                            status = "\033[92m✓ Configured\033[0m" if configured else "\033[91m✗ Not Configured\033[0m"
                            print(f"  {provider}: {status}")
                        print("\nTo switch providers: :provider <name> (e.g., :provider gemini)")
                    else:
                        # Switch provider
                        new_provider = cmd_args[0].lower()
                        try:
                            print(f"\033[94m🔄 Attempting to switch provider to '{new_provider}'...\033[0m")
                            set_active_provider(new_provider) # Set the state in model_config
                            # Re-initialize the agent with the new model configuration
                            excel_assistant_agent = create_excel_assistant_agent() # << RECREATE AGENT
                            print(f"\033[92m✔️ Provider switched to: {new_provider}. Agent re-initialized.\033[0m")
                            logger.info(f"Provider switched to '{new_provider}' via command and agent refreshed.")
                        except ValueError as e: # Catch errors from set_active_provider (e.g., unsupported)
                            print(f"\033[91m❌ Error switching provider: {e}\033[0m")
                        except Exception as e: # Catch errors from create_excel_assistant_agent
                            print(f"\033[91m❌ Failed to initialize agent for provider '{new_provider}': {e}\033[0m")
                            logger.error(f"Failed to recreate agent after switching to {new_provider}", exc_info=True)
                            # Optional: Attempt to switch back? For now, just report error.
                            print("\033[93m   Agent may be in an unstable state. Consider restarting or switching to a known good provider.\033[0m")

                else:
                    print(f"\033[91m❌ Unknown command: ':{command}'. Type ':help' for options.\033[0m")

            # --- Regular Instruction Handling ---
            elif user_input_str.lower() in ["exit", "quit"]: # Handle plain exit/quit too
                 print("\n\033[1m\033[95m--- Current System Prompt (for active agent) ---\033[0m")
                 try:
                     temp_wrapper = RunContextWrapper(context=app_context)
                     current_prompt = _dynamic_instructions(temp_wrapper, excel_assistant_agent)
                     print(current_prompt)
                 except Exception as e: print(f"\033[91mError generating system prompt: {e}\033[0m")
                 print("\033[1m\033[95m--- End System Prompt ---\033[0m\n")
                 break
            else:
                # Check if workbook is open before running agent
                if not excel_manager or not excel_manager.book:
                    print("\033[93m⚠️ Please open or create a workbook first using ':open <path>' or ':new'.\033[0m")
                    continue

                # Run the agent with the user input using the streaming function
                # This uses the *currently active* excel_assistant_agent instance
                await run_agent_streamed(
                    excel_assistant_agent,
                    user_input_str,
                    app_context # Pass the application context
                )

        except EOFError: # Handle Ctrl+D
            print("\n\033[94m👋 EOF received, exiting.\033[0m")
            break
        except KeyboardInterrupt: # Handle Ctrl+C during input prompt
            print("\n\033[93m🚫 Input cancelled (Ctrl+C). Type 'exit' or ':help'.\033[0m")
            continue # Continue loop after Ctrl+C during input
        except Exception as e: # Catch unexpected errors in the main loop
            print(f"\n\033[91m❌ Unexpected Error in CLI loop: {e}\033[0m")
            logger.error("Error in main CLI loop", exc_info=True)
            # Consider if loop should continue or break on unexpected errors
            await asyncio.sleep(1) # Small delay before next prompt

    # --- Cleanup ---
    if excel_manager:
        print("\n\033[94m👋 Closing active workbook before exiting...\033[0m")
        try:
            await excel_manager.close()
        except Exception as e:
            logger.error(f"Error during final workbook close: {e}", exc_info=True)

    print("\n\033[96m👋 Exiting Excel AI Assistant. Goodbye!\033[0m")


# --- Entry Point ---
if __name__ == "__main__":
    try:
        # Check Python version
        if sys.version_info < (3, 8):
             print("❌ Error: Python 3.8 or higher is required.", file=sys.stderr)
             sys.exit(1)
        asyncio.run(main())
    except KeyboardInterrupt:
        # This handles Ctrl+C before the main loop starts or after it exits cleanly
        print("\n\033[93m🚫 Exiting due to user interrupt.\033[0m")
        sys.exit(0)
    except Exception as e:
        # Catch critical errors during startup/shutdown (outside main loop)
        logger.critical(f"CLI critical error: {e}", exc_info=True)
        print(f"\n\033[91m❌ Critical Error: {e}\033[0m", file=sys.stderr)
        sys.exit(1)