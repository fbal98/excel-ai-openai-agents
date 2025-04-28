#!/usr/bin/env python
"""
Interactive command-line interface for the Excel AI Assistant.
Starts without an active workbook. Use **slash-commands** like /open, /new, /close.
Supports switching LLM providers via the /provider command.
"""

import asyncio
import contextlib
import itertools
import logging
import os
import sys
import shlex  # For parsing slash commands
from typing import Optional, Any

# Attempt to import litellm for advanced error handling
try:
    import litellm
    LITELLM_AVAILABLE = True
except ImportError:
    LITELLM_AVAILABLE = False

# Attempt to import prompt_toolkit for a nicer CLI
try:
    from prompt_toolkit import PromptSession
    from prompt_toolkit.history import FileHistory
    from prompt_toolkit.auto_suggest import AutoSuggestFromHistory
    from prompt_toolkit.styles import Style
    from prompt_toolkit.completion import Completer, Completion
    from prompt_toolkit.document import Document
    PROMPT_TOOLKIT_AVAILABLE = True
except ImportError:
    PROMPT_TOOLKIT_AVAILABLE = False

try:
    from dotenv import load_dotenv
    DOTENV_AVAILABLE = True
except ImportError:
    DOTENV_AVAILABLE = False

# Attempt to import openai errors
try:
    import openai
except ImportError:
    class DummyOpenAIError(Exception):
        pass
    openai = type("obj", (object,), {"InternalServerError": DummyOpenAIError})

from agents import Runner, Agent
from agents.result import RunResultStreaming
from agents.stream_events import StreamEvent
from agents.exceptions import UserError
from agents import RunContextWrapper

from .excel_ops import ExcelConnectionError, ExcelManager
from .model_config import (
    set_active_provider,
    get_active_provider,
    list_available_providers
)
from .model_integration import create_excel_assistant_agent
from .context import AppContext
from .stream_renderer import format_event
from .constants import SHOW_COST
from .agent_core import _dynamic_instructions, _format_workbook_shape

log_level_name = os.getenv("LOG_LEVEL", "INFO").upper()
log_level = getattr(logging, log_level_name, logging.INFO)
log_file = os.getenv("EXCEL_AI_LOG_FILE", "excel_ai.log")

# Clear existing handlers to avoid duplication in repeated runs
for h in logging.root.handlers[:]:
    logging.root.removeHandler(h)

file_handler = logging.FileHandler(log_file, mode="a", encoding="utf-8")
file_handler.setLevel(log_level)
file_handler.setFormatter(logging.Formatter("%(asctime)s - %(name)s - %(levelname)s - %(message)s"))

console_handler = logging.StreamHandler(sys.stderr)
console_handler.setLevel(logging.WARNING)  # Only warnings+ to console by default
console_handler.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))

logging.basicConfig(level=log_level, handlers=[file_handler, console_handler])
logger = logging.getLogger(__name__)

HISTORY_FILE = ".excel_ai_history"

# A nice style for prompt_toolkit if available
if PROMPT_TOOLKIT_AVAILABLE:
    cli_style = Style.from_dict({
        "prompt": "bold cyan",
        "prompt.no-workbook": "bold yellow",
        "input": "",
        "output.info": "cyan",
        "output.warning": "yellow",
        "output.success": "green",
        "output.cost": "italic #888888",
        "spinner": "magenta",
        "error": "bold red",
        "completion-menu.completion": "bg:#1e1e1e #bcbcbc",
        "completion-menu.completion.current": "bg:#005f5f #ffffff bold",
        "completion-menu.meta": "#6c6c6c italic",
        "scrollbar.background": "bg:#262626",
        "scrollbar.button": "bg:#3a3a3a",
    })

CLI_COMMANDS = {
    "open": "Open an Excel workbook",
    "new": "Create a new blank workbook",
    "close": "Close the current workbook",
    "clear": "Clear the terminal screen",
    "provider": "Switch or show LLM provider",
    "history": "Show or clear conversation history",
    "cost": "Show cost of last agent run",
    "reset-chat": "Reset conversation history",
    "shape": "Show workbook shape",
    "help": "Show this help message",
    "exit": "Exit the CLI",
    "quit": "Exit the CLI",
}

# Slash command completer for prompt_toolkit
if PROMPT_TOOLKIT_AVAILABLE:
    class SlashCommandCompleter(Completer):
        def __init__(self, commands):
            self.commands = commands
            self.max_cmd_len = max(len(f"/{cmd}") for cmd in commands) if commands else 0

        def get_completions(self, document: Document, complete_event):
            text = document.text_before_cursor
            if not text.startswith("/") or " " in text:
                return
            needle = text[1:].lower()

            # Some icons for visual flair
            command_styles = {
                "open": ("📂", "36"),
                "new": ("✨", "36"),
                "close": ("🔒", "36"),
                "clear": ("🧹", "35"),
                "provider": ("⚙️", "34"),
                "history": ("📜", "33"),
                "cost": ("💰", "33"),
                "reset-chat": ("🔄", "35"),
                "shape": ("📊", "34"),
                "help": ("❓", "32"),
                "exit": ("🚪", "31"),
                "quit": ("🚪", "31"),
            }

            from prompt_toolkit.formatted_text import ANSI
            for name, desc in self.commands.items():
                if name.startswith(needle):
                    icon, color = command_styles.get(name, ("•", "37"))
                    cmd_display = f"/{name}"
                    padding = self.max_cmd_len - len(cmd_display) + 2
                    ansi_display = ANSI(
                        f"\033[1;{color}m{icon} {cmd_display}\033[0m"
                        + " " * padding
                        + f"\033[2;37m{desc}\033[0m"
                    )
                    yield Completion(
                        text=name,
                        start_position=-len(needle),
                        display=ansi_display,
                    )

def _normalize_content(content: Any) -> str:
    """
    Convert model response content into a plain string so it is always valid.
    Sometimes the content can be a list of dicts from certain providers.
    """
    if isinstance(content, list):
        parts = []
        for item in content:
            if isinstance(item, dict):
                parts.append(str(item.get("text", item.get("content", ""))).strip())
            else:
                parts.append(str(item).strip())
        return "\n".join(x for x in parts if x)
    return str(content)

async def _spinner(prefix="⌛ Thinking", interval=0.2):
    """
    Simple async spinner used during agent runs.
    """
    try:
        is_tty = sys.stdout.isatty()
        supports_utf8 = sys.stdout.encoding.lower() == "utf-8"

        if not is_tty or not supports_utf8:
            print(f"{prefix} ...", end="", flush=True)
            while True:
                await asyncio.sleep(interval * 5)

        spinner_chars = "|/-\\"
        for char in itertools.cycle(spinner_chars):
            print(f"\r\033[95m{prefix}\033[0m {char} ", end="", flush=True)
            await asyncio.sleep(interval)
    except asyncio.CancelledError:
        clear_len = len(prefix) + 2
        if sys.stdout.isatty():
            print("\r" + " " * clear_len + "\r", end="", flush=True)
        else:
            print()
        raise
    except Exception as e:
        logger.error(f"Spinner error: {e}", exc_info=True)
        clear_len = len(prefix) + 2
        if sys.stdout.isatty():
            print("\r" + " " * clear_len + "\r", end="", flush=True)
        else:
            print()

async def _run_agent_with_retry(agent: Agent, input_data: list, ctx: AppContext,
                                thinking_task=None, is_retry=False) -> Optional[RunResultStreaming]:
    """
    Helper function that runs the agent in streaming mode with an optional retry
    if certain known errors occur.
    """
    retry_suffix = " (retry)" if is_retry else ""

    # ── Lightweight shape refresh for retries as well ────────────────────
    try:
        if ctx.update_shape(tool_name="pre_turn_refresh"):
            ctx.pending_write_count = 0
            logger.debug("Pre-turn shape refresh OK (v%s)", ctx.shape.version)
    except Exception as e:
        logger.warning("Pre-turn shape refresh failed: %s", e)
    # --------------------------------------------------------------------

    result_stream = None
    first_event_received = False

    try:
        result_stream = Runner.run_streamed(agent, input=input_data, context=ctx)
        if thinking_task is None or thinking_task.done():
            spinner_prefix = f"🤖 {get_active_provider().capitalize()} Thinking{retry_suffix}"
            if sys.stdout.isatty():
                thinking_task = asyncio.create_task(_spinner(prefix=spinner_prefix))
            else:
                print(f"{spinner_prefix}...", flush=True)

        async for ev in result_stream.stream_events():
            formatted_output = format_event(ev)
            if formatted_output and not first_event_received:
                first_event_received = True
                if thinking_task and not thinking_task.done():
                    thinking_task.cancel()
                    with contextlib.suppress(asyncio.CancelledError):
                        await thinking_task
            if formatted_output:
                print(formatted_output, end="")
                sys.stdout.flush()

        if thinking_task and not thinking_task.done():
            thinking_task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await thinking_task

        if not first_event_received and result_stream:
            if result_stream.final_output:
                fo_str = result_stream.final_output
                if isinstance(fo_str, str):
                    prefix = "✔️ 🤖 Agent: "
                    lines = fo_str.strip().splitlines()
                    indent = " " * len(prefix)
                    out = []
                    for i, line in enumerate(lines):
                        if i == 0:
                            out.append(f"\033[92m{prefix}{line}\033[0m")
                        else:
                            out.append(f"\033[92m{indent}{line}\033[0m")
                    print("\n".join(out))
                else:
                    print(f"\n\033[94m🤔 Run finished. Output: {fo_str}\033[0m")
            else:
                print("\n\033[94m🤔 Run finished without visible streaming output.\033[0m")

        # Save conversation history
        if result_stream:
            try:
                conversation_history = result_stream.to_input_list()
                filtered_history = []
                for msg in conversation_history:
                    # Gracefully handle items with no role
                    role = msg.get("role") or "assistant"
                    # Skip SDK tool-call stubs and explicit tool messages
                    if (role == "assistant" and msg.get("name", "").endswith("_tool") and not msg.get("content")) \
                            or role == "tool":
                        continue
                    norm_msg = {
                        "role": role,
                        "content": _normalize_content(msg.get("content", "")),
                    }
                    if "name" in msg:
                        norm_msg["name"] = msg["name"]
                    filtered_history.append(norm_msg)
                ctx.state["conversation_history"] = filtered_history
                logger.info(f"Saved conversation history{retry_suffix} with {len(filtered_history)} messages")
            except Exception as e:
                logger.error(f"Error saving conversation history{retry_suffix}: {e}", exc_info=True)
                if "conversation_history" not in ctx.state:
                    ctx.state["conversation_history"] = []

        return result_stream

    except Exception as e:
        if thinking_task and not thinking_task.done():
            thinking_task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await thinking_task
        raise

async def run_agent_streamed(agent: Agent, user_input: str, ctx: AppContext) -> Optional[RunResultStreaming]:
    """
    Runs the agent using run_streamed, printing events as they appear.
    """
    from .model_config import get_active_provider

    # Check if current provider is gemini; if so, fallback to non-streamed run to avoid JSONDecodeError
    if get_active_provider() == "gemini":
        # Fallback to non-streamed calls
        # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        current_user_message = {"role": "user", "content": user_input}
        conversation_history = ctx.state.get("conversation_history", [])
        if not isinstance(conversation_history, list):
            conversation_history = []
        input_data = conversation_history + [current_user_message]

        print("\033[90m(Using non-streaming run for Gemini)\033[0m")
        # Directly call Runner.run
        try:
            result_run = await Runner.run(agent, input=input_data, context=ctx)

            # Cost/Usage logic for fallback
            usage = getattr(ctx, "usage", None)
            if usage and hasattr(agent, "model"):
                from .costs import dollars_for_usage
                cost_val = dollars_for_usage(usage, model_name_from_agent=agent.model)
                total_tokens = (usage.input_tokens or 0) + (usage.output_tokens or 0)
                ctx.state["last_run_cost"] = cost_val
                ctx.state["last_run_usage"] = {
                    "input_tokens": usage.input_tokens or 0,
                    "output_tokens": usage.output_tokens or 0,
                    "total_tokens": total_tokens,
                    "model_name": agent.model,
                }

            # Print final output if any
            if result_run and result_run.final_output:
                msg = str(result_run.final_output).strip()
                if msg:
                    print(f"\n\033[92m✔️ 🤖 Agent: {msg}\033[0m")
            return None  # We didn't produce a streaming result

        except Exception as e:
            print(f"\033[91m❌ Unexpected error (Gemini fallback): {e}\033[0m")
            logger.error("Error in gemini fallback path", exc_info=True)
            return None

    # If not gemini, proceed with streaming.
    """
    Runs the agent using run_streamed, printing events as they appear.
    """
    if not ctx.excel_manager or not ctx.excel_manager.book:
        print("\n\033[93m⚠️ No active workbook. Use '/open <path>' or '/new' first.\033[0m")
        return None

    thinking_task = None
    first_event_received = False
    spinner_prefix = f"🤖 {get_active_provider().capitalize()} Thinking"

    # ── Lightweight shape refresh before every turn ──────────────────────
    try:
        if ctx.update_shape(tool_name="pre_turn_refresh"):
            ctx.pending_write_count = 0  # keep debounce logic in sync
            logger.debug("Pre-turn shape refresh OK (v%s)", ctx.shape.version)
    except Exception as e:
        logger.warning("Pre-turn shape refresh failed: %s", e)
    # --------------------------------------------------------------------

    result_stream: Optional[RunResultStreaming] = None

    current_user_message = {"role": "user", "content": user_input}
    conversation_history = ctx.state.get("conversation_history", [])
    if not isinstance(conversation_history, list):
        conversation_history = []
    input_data = conversation_history + [current_user_message]

    try:
        if sys.stdout.isatty():
            thinking_task = asyncio.create_task(_spinner(prefix=spinner_prefix))
        else:
            print(f"{spinner_prefix}...", flush=True)

        result_stream = Runner.run_streamed(agent, input=input_data, context=ctx)

        async for ev in result_stream.stream_events():
            formatted_output = format_event(ev)
            if formatted_output and not first_event_received:
                first_event_received = True
                if thinking_task and not thinking_task.done():
                    thinking_task.cancel()
                    with contextlib.suppress(asyncio.CancelledError):
                        await thinking_task
            if formatted_output:
                print(formatted_output, end="")
                sys.stdout.flush()

        if thinking_task and not thinking_task.done():
            thinking_task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await thinking_task

        # Handle final output if no events were streamed
        if not first_event_received and result_stream:
            final_out = result_stream.final_output
            if final_out:
                if isinstance(final_out, str):
                    prefix = "✔️ 🤖 Agent: "
                    lines = final_out.strip().splitlines()
                    indent = " " * len(prefix)
                    out = []
                    for i, line in enumerate(lines):
                        if i == 0:
                            out.append(f"\033[92m{prefix}{line}\033[0m")
                        else:
                            out.append(f"\033[92m{indent}{line}\033[0m")
                    print("\n".join(out))
                else:
                    print(f"\n\033[94m🤔 Run finished. Output: {final_out}\033[0m")
            else:
                print("\n\033[94m🤔 Run finished without generating visible streaming output.\033[0m")

        # Save conversation history
        if result_stream:
            try:
                conversation_history = result_stream.to_input_list()
                filtered_history = []
                for msg in conversation_history:
                    role = msg.get("role") or "assistant"  # Fallback role
                    # Skip SDK tool-call stubs and explicit tool messages
                    if (role == "assistant" and msg.get("name", "").endswith("_tool") and not msg.get("content")) \
                            or role == "tool":
                        continue
                    norm_msg = {
                        "role": role,
                        "content": _normalize_content(msg.get("content", "")),
                    }
                    if "name" in msg:
                        norm_msg["name"] = msg["name"]
                    filtered_history.append(norm_msg)
                ctx.state["conversation_history"] = filtered_history
            except Exception as e:
                logger.error(f"Error saving conversation history: {e}", exc_info=True)
                ctx.state["conversation_history"] = ctx.state.get("conversation_history", [])

        # Info logs after the run
        if SHOW_COST:
            if "last_run_cost" in ctx.state:
                cost = ctx.state.get("last_run_cost", 0.0)
                usage_info = ctx.state.get("last_run_usage", {})
                tokens = usage_info.get("total_tokens", "N/A")
                model_used = usage_info.get("model_name", "N/A")
                cost_style = "\033[90m\033[3m"
                print(f"{cost_style}💰 Cost: ${cost:.4f} ({tokens} tokens, Model: {model_used})\033[0m", file=sys.stderr)
            else:
                # Debug output to identify why costs might not be showing
                logger.warning("SHOW_COST is True but 'last_run_cost' not found in context state")
                print("\033[93m⚠️ Cost info missing. Check logs for details.\033[0m", file=sys.stderr)

        return result_stream

    except asyncio.CancelledError:
        if thinking_task and not thinking_task.done():
            thinking_task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await thinking_task
        print("\n\033[93m🚫 Agent run cancelled.\033[0m")
        return None
    except UserError as e:
        if thinking_task and not thinking_task.done():
            thinking_task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await thinking_task
        print(f"\n\033[91m❌ Agent User Error: {e}\033[0m")
        return None
    except openai.InternalServerError as e:
        if thinking_task and not thinking_task.done():
            thinking_task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await thinking_task
        print("\n\033[91m❌ OpenAI Internal Server Error.\033[0m")
        print(f"\033[93mDetails: {e}\033[0m")
        ctx.state["last_error"] = {"type": "openai_server_error", "message": str(e)}
        return None
    except Exception as e:
        if thinking_task and not thinking_task.done():
            thinking_task.cancel()
            with contextlib.suppress(asyncio.CancelledError):
                await thinking_task
        print(f"\n\033[91m❌ Unexpected Error: {e}\033[0m")
        logger.error("Error during agent run", exc_info=True)
        return None

async def main():
    if DOTENV_AVAILABLE:
        if load_dotenv(override=True):
            logger.info(".env loaded.")
            provider = os.getenv("DEFAULT_MODEL_PROVIDER", "openai").lower()
            try:
                set_active_provider(provider)
            except Exception:
                pass

    from .model_config import _is_provider_configured
    current_provider = get_active_provider()
    if not _is_provider_configured(current_provider):
        print(f"\033[93m⚠️ Warning: provider '{current_provider}' may not be fully configured.\033[0m")

    import argparse
    parser = argparse.ArgumentParser(description="Excel AI Assistant CLI")
    parser.add_argument("file_path", nargs="?", default=None, help="Optional workbook path (not auto-opened).")
    parser.add_argument("--attach", action="store_true", help="Attach to an existing Excel instance if found.")
    parser.add_argument("--kill-others", action="store_true", help="Kill other Excel instances first.")
    parser.add_argument("--verbose", "-v", action="store_true", help="Increase console log level to DEBUG.")
    args = parser.parse_args()

    if args.verbose:
        console_handler.setLevel(logging.DEBUG)
        logger.info("Verbose console logging enabled.")
    else:
        console_handler.setLevel(logging.WARNING)

    excel_assistant_agent = None
    try:
        excel_assistant_agent = create_excel_assistant_agent()
    except Exception as e:
        print(f"\033[91m❌ Unable to initialize agent for provider '{get_active_provider()}': {e}\033[0m")
        logger.critical("Failed agent creation", exc_info=True)

    print("\n\033[1m\033[96m🚀 Excel AI Assistant CLI\033[0m")
    print(f"\033[90mType instructions or slash-commands (/help). Provider: {get_active_provider()}\033[0m")
    print("\033[93m⚠️ No workbook loaded. Use /open <path> or /new.\033[0m")

    excel_manager = None
    app_context = AppContext(excel_manager=None)
    app_context.state["conversation_history"] = []

    if PROMPT_TOOLKIT_AVAILABLE:
        slash_completer = SlashCommandCompleter(CLI_COMMANDS)
        session = PromptSession(
            history=FileHistory(HISTORY_FILE),
            auto_suggest=AutoSuggestFromHistory(),
            completer=slash_completer,
            complete_while_typing=True,
            style=cli_style,
        )
        async def get_input_ptk(prompt_text: str):
            return await session.prompt_async(prompt_text)
        get_input = get_input_ptk
    else:
        async def get_input_basic(prompt_text: str):
            return await asyncio.to_thread(input, prompt_text)
        get_input = get_input_basic

    while True:
        try:
            prompt_prefix = f"💬 ({get_active_provider()}) User: "
            user_inp = await get_input(prompt_prefix)
            user_inp = user_inp.strip()
            if not user_inp:
                continue

            if user_inp.startswith("/"):
                parts = []
                try:
                    parts = shlex.split(user_inp[1:])
                except ValueError as e:
                    print(f"\033[91m❌ Parse Error: {e}\033[0m")
                    continue
                cmd = parts[0].lower() if parts else ""
                cmd_args = parts[1:]

                if cmd not in CLI_COMMANDS:
                    if cmd == "":
                        continue
                    print(f"\033[91m❌ Unknown command '/{cmd}'. Use '/help'.\033[0m")
                    continue

                if cmd in ["exit", "quit"]:
                    print("\n\033[95m--- Current System Prompt ---\033[0m")
                    if excel_assistant_agent:
                        try:
                            wrapper = RunContextWrapper(context=app_context)
                            sys_prompt = _dynamic_instructions(wrapper, excel_assistant_agent)
                            print(sys_prompt)
                        except Exception as err:
                            print(f"Error showing system prompt: {err}")
                    print("\033[95m--- End System Prompt ---\n")
                    break

                elif cmd == "open":
                    if not cmd_args:
                        print("\033[91mUsage: /open <file_path>\033[0m")
                        continue
                    path_to_open = cmd_args[0]
                    print(f"\033[94mClosing any open workbook, then opening '{path_to_open}'...\033[0m")
                    if excel_manager:
                        try:
                            await excel_manager.close()
                        except:
                            pass
                        conversation_history = app_context.state.get("conversation_history", [])
                        excel_manager = None
                        app_context.excel_manager = None
                        app_context.shape = None
                        app_context.state = {}
                        if conversation_history:
                            app_context.state["conversation_history"] = conversation_history
                    try:
                        excel_manager = ExcelManager(
                            file_path=path_to_open,
                            visible=True,
                            attach_existing=args.attach,
                            kill_others=args.kill_others
                        )
                        await excel_manager.open()
                        app_context.excel_manager = excel_manager
                        app_context.state["conversation_history"] = []
                        updated = app_context.update_shape(tool_name="/open")
                        if updated and app_context.shape:
                            print(f"\033[92m✔️ Opened '{excel_manager.file_path}' (Shape v{app_context.shape.version}).\033[0m")
                        else:
                            print(f"\033[93m⚠️ Opened '{excel_manager.file_path}', shape not updated.\033[0m")
                    except Exception as e:
                        print(f"\033[91m❌ Error opening '{path_to_open}': {e}\033[0m")
                        excel_manager = None
                        app_context.excel_manager = None

                elif cmd == "new":
                    print("\033[94mCreating a new blank workbook...\033[0m")
                    if excel_manager:
                        try:
                            await excel_manager.close()
                        except:
                            pass
                        conversation_history = app_context.state.get("conversation_history", [])
                        excel_manager = None
                        app_context.excel_manager = None
                        app_context.shape = None
                        app_context.state = {}
                        if conversation_history:
                            app_context.state["conversation_history"] = conversation_history
                    try:
                        excel_manager = ExcelManager(
                            file_path=None,
                            visible=True,
                            attach_existing=args.attach,
                            kill_others=args.kill_others
                        )
                        await excel_manager.open()
                        app_context.excel_manager = excel_manager
                        app_context.state["conversation_history"] = []
                        updated = app_context.update_shape(tool_name="/new")
                        if updated and app_context.shape:
                            print(f"\033[92m✔️ New workbook '{excel_manager.file_path}' (Shape v{app_context.shape.version}).\033[0m")
                        else:
                            print("\033[93m⚠️ New workbook created, shape not updated.\033[0m")
                    except Exception as e:
                        print(f"\033[91m❌ Error creating new workbook: {e}\033[0m")
                        excel_manager = None
                        app_context.excel_manager = None

                elif cmd == "close":
                    if excel_manager:
                        print("\033[94mClosing current workbook...\033[0m")
                        try:
                            await excel_manager.close()
                            print("\033[92m✔️ Workbook closed.\033[0m")
                        except Exception as e:
                            print(f"\033[91m❌ Error closing workbook: {e}\033[0m")
                        finally:
                            conversation_history = app_context.state.get("conversation_history", [])
                            excel_manager = None
                            app_context.excel_manager = None
                            app_context.shape = None
                            app_context.state = {}
                            if conversation_history:
                                app_context.state["conversation_history"] = conversation_history
                            print("\033[93m⚠️ No workbook loaded.\033[0m")
                    else:
                        print("\033[93mNo workbook to close.\033[0m")

                elif cmd == "clear":
                    # Clears the terminal
                    print("\033[H\033[J", end="")

                elif cmd == "help":
                    print("\n\033[96mAvailable commands:\033[0m")
                    max_len = max(len(f"/{c}") for c in CLI_COMMANDS)
                    for c, desc in CLI_COMMANDS.items():
                        disp = f"/{c}"
                        pad = " " * (max_len - len(disp) + 2)
                        print(f"  \033[1m\033[94m{disp}\033[0m{pad}\033[90m{desc}\033[0m")

                elif cmd == "history":
                    if cmd_args and cmd_args[0].lower() == "clear":
                        app_context.state["conversation_history"] = []
                        print("\033[92mConversation history cleared.\033[0m")
                    else:
                        history = app_context.state.get("conversation_history", [])
                        if not history:
                            print("\033[93mNo conversation history.\033[0m")
                        else:
                            print(f"\033[94m{len(history)} messages in conversation.\033[0m")
                            role_counts = {}
                            for m in history:
                                r = m.get("role", "unknown")
                                role_counts[r] = role_counts.get(r, 0) + 1
                            for r, ct in role_counts.items():
                                print(f"  - {r}: {ct}")

                elif cmd == "cost":
                    c = app_context.state.get("last_run_cost")
                    usage = app_context.state.get("last_run_usage", {})
                    tokens = usage.get("total_tokens", "N/A")
                    used_model = usage.get("model_name", get_active_provider())
                    
                    # Debug output of all state variables to help troubleshoot
                    logger.debug(f"State contents: {', '.join(app_context.state.keys())}")
                    if "usage" in app_context.state:
                        logger.debug(f"Usage details in state: {app_context.state['usage']}")
                    
                    if c is None:
                        print("\033[93mNo cost info yet.\033[0m")
                        print("\033[93mActive provider: " + get_active_provider() + "\033[0m")
                        
                        # Show any usage stats that might be available directly
                        if hasattr(app_context, 'usage') and app_context.usage:
                            input_tokens = getattr(app_context.usage, "input_tokens", 0) or 0
                            output_tokens = getattr(app_context.usage, "output_tokens", 0) or 0
                            print(f"\033[93mFound usage directly on context: Input={input_tokens}, Output={output_tokens}\033[0m")
                    else:
                        print(f"\033[94mCost: ${c:.4f} ({tokens} tokens, Model: {used_model})\033[0m")

                elif cmd == "reset-chat":
                    if excel_manager:
                        app_context.state["conversation_history"] = []
                        print("\033[92mConversation history reset.\033[0m")
                    else:
                        print("\033[93mNo active workbook. Nothing to reset.\033[0m")

                elif cmd == "shape":
                    if app_context.shape:
                        shape_str = _format_workbook_shape(app_context.shape)
                        print("\033[94mWorkbook Shape:\033[0m")
                        print(shape_str)
                    else:
                        print("\033[93mNo shape available.\033[0m")

                elif cmd == "provider":
                    if not cmd_args:
                        current = get_active_provider()
                        all_p = list_available_providers()
                        print(f"\n\033[94mCurrent provider: {current}\033[0m")
                        print("Available providers:")
                        for p, configured in all_p.items():
                            s = "\033[92m✓" if configured else "\033[91m✗"
                            print(f"  {p}: {s}\033[0m configured")
                        print("Use '/provider <name>' to switch.")
                    else:
                        newp = cmd_args[0].lower()
                        try:
                            set_active_provider(newp)
                            excel_assistant_agent = create_excel_assistant_agent()
                            print(f"\033[92mSwitched to provider '{newp}'.\033[0m")
                        except Exception as e:
                            print(f"\033[91mFailed to switch provider: {e}\033[0m")
                            excel_assistant_agent = None

            elif user_inp.lower() in ["exit", "quit"]:
                print("\n\033[95m--- Current System Prompt ---\033[0m")
                if excel_assistant_agent:
                    try:
                        wrapper = RunContextWrapper(context=app_context)
                        sys_prompt = _dynamic_instructions(wrapper, excel_assistant_agent)
                        print(sys_prompt)
                    except Exception as err:
                        print(f"Error showing system prompt: {err}")
                print("\033[95m--- End System Prompt ---\n")
                break
            else:
                # Normal text - run the agent if possible
                if not excel_manager or not excel_manager.book:
                    print("\033[93mNo workbook open. Use /open or /new.\033[0m")
                    continue
                if not excel_assistant_agent:
                    print("\033[91mAgent not initialized. Use /provider to configure.\033[0m")
                    continue
                await run_agent_streamed(excel_assistant_agent, user_inp, app_context)

        except EOFError:
            print("\n\033[94mEOF reached, exiting.\033[0m")
            break
        except KeyboardInterrupt:
            print("\n\033[93mCancelled.\033[0m")
            continue
        except Exception as e:
            print(f"\n\033[91mUnexpected error: {e}\033[0m")
            logger.error("CLI loop error", exc_info=True)
            await asyncio.sleep(0.5)

    if excel_manager:
        print("\033[94mClosing workbook before exit...\033[0m")
        try:
            await excel_manager.close()
        except Exception as e:
            logger.error(f"Error closing on exit: {e}", exc_info=True)

    print("\033[96mGoodbye!\033[0m")
    sys.exit(0)

if __name__ == "__main__":
    try:
        if sys.version_info < (3, 8):
            print("Python 3.8+ required.", file=sys.stderr)
            sys.exit(1)
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\n\033[93mExiting by Ctrl+C.\033[0m")
        sys.exit(0)
    except Exception as e:
        logger.critical(f"Startup error: {e}", exc_info=True)
        print(f"\033[91mCritical error: {e}\033[0m", file=sys.stderr)
        sys.exit(1)