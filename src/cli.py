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
from agents import RunContextWrapper, FunctionTool  # Added FunctionTool
from openai.types.responses import ResponseTextDeltaEvent

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

def patch_tool_schemas(agent: Agent):
    """
    Ensures that the JSON schema for each FunctionTool's parameters
    explicitly includes 'additionalProperties': False, as required by OpenAI API.
    """
    if not agent or not agent.tools:
        logger.debug("Patching schemas skipped: No agent or no tools found.")
        return

    logger.debug(f"Patching schemas for {len(agent.tools)} tools...")
    patched_count = 0
    for tool in agent.tools:
        if isinstance(tool, FunctionTool):
            # Ensure params_json_schema exists and is a dictionary
            if hasattr(tool, 'params_json_schema') and isinstance(tool.params_json_schema, dict):
                schema = tool.params_json_schema
                # Check if 'type' is 'object', as additionalProperties only applies to objects
                if schema.get("type") == "object":
                    if schema.get("additionalProperties") is not False:
                        schema["additionalProperties"] = False
                        logger.info(f"Patched schema for tool '{tool.name}': Set additionalProperties=False.")
                        patched_count += 1
                else:
                     logger.debug(f"Skipping schema patch for tool '{tool.name}': Schema type is not 'object' (type: {schema.get('type')}).")
            else:
                logger.debug(f"Skipping schema patch for tool '{tool.name}': No valid params_json_schema dictionary found.")
        else:
            logger.debug(f"Skipping schema patch for tool '{getattr(tool, 'name', 'Unnamed Tool')}': Not a FunctionTool instance.")
    if patched_count > 0:
        logger.info(f"Schema patching complete. Patched {patched_count} tool schemas.")
    else:
        logger.debug("Schema patching complete. No schemas required patching.")

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
        # Usage info must be set on the agent's model_settings, not in run_streamed
        result_stream = Runner.run_streamed(agent, input=input_data, context=ctx)
        if thinking_task is None or thinking_task.done():
            spinner_prefix = f"🤖 {get_active_provider().capitalize()} Thinking{retry_suffix}"
            if sys.stdout.isatty():
                thinking_task = asyncio.create_task(_spinner(prefix=spinner_prefix))
            else:
                print(f"{spinner_prefix}...", flush=True)

        async for ev in result_stream.stream_events():
            # ── Raw delta tokens ────────────────────────────────────────────
            if ev.type == "raw_response_event" and isinstance(ev.data, ResponseTextDeltaEvent):
                print(ev.data.delta, end="", flush=True)
                if not first_event_received:
                    first_event_received = True
                    if thinking_task and not thinking_task.done():
                        thinking_task.cancel()
                        with contextlib.suppress(asyncio.CancelledError):
                            await thinking_task
                continue

            # ── Higher-level events (tools, thoughts, etc.) ────────────────
            # ── Raw delta tokens ────────────────────────────────────────────
            with contextlib.suppress(asyncio.CancelledError):
                if not first_event_received:
                    first_event_received = True
                    if thinking_task and not thinking_task.done():
                        thinking_task.cancel()
                        with contextlib.suppress(asyncio.CancelledError):
                            await thinking_task
                continue

            # ── Higher-level events (tools, thoughts, etc.) ────────────────
            formatted_output = format_event(ev)
            if formatted_output:
                if not first_event_received:
                    first_event_received = True
                    if thinking_task and not thinking_task.done():
                        thinking_task.cancel()
                        with contextlib.suppress(asyncio.CancelledError):
                            await thinking_task
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

            # ---- Persist conversation history for non-streaming runs ----
            try:
                hist = ctx.state.get("conversation_history", [])
                if not isinstance(hist, list):
                    hist = []
                # Add the current user message unless it is a duplicate
                if not hist or hist[-1] != current_user_message:
                    hist.append(current_user_message)
                # Add the assistant reply (if any)
                if result_run and result_run.final_output:
                    assistant_msg = str(result_run.final_output).strip()
                    if assistant_msg:
                        hist.append({"role": "assistant", "content": assistant_msg})
                ctx.state["conversation_history"] = hist
                logger.debug(f"Saved non-streaming conversation history with {len(hist)} messages")
            except Exception as hist_err:
                logger.error(
                    f"Error updating conversation history (Gemini fallback): {hist_err}",
                    exc_info=True,
                )

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

        # Usage info must be set on the agent's model_settings, not in run_streamed
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
                
                # Extract and store cost information from result_stream
                try:
                    # Get usage info from the run (for OpenAI streaming runs)
                    usage = None
                    
                    # Debug all available attributes on result_stream
                    logger.debug(f"Available result_stream attributes: {dir(result_stream)}")
                    
                    # Try to get usage from result_stream.usage first
                    if hasattr(result_stream, 'usage') and result_stream.usage:
                        usage = result_stream.usage
                        logger.info(f"Got usage from result_stream.usage: {usage}")
                    # Then try context.usage
                    elif hasattr(ctx, 'usage') and ctx.usage:
                        usage = ctx.usage
                        logger.info(f"Got usage from context.usage: {usage}")
                    # Try to get usage from raw_responses if available
                    elif hasattr(result_stream, 'raw_responses') and result_stream.raw_responses:
                        for resp in result_stream.raw_responses:
                            if hasattr(resp, 'usage') and resp.usage:
                                usage = resp.usage
                                logger.info(f"Got usage from raw_responses: {usage}")
                                break
                    # Try to access _usage if it exists (some SDK versions use private attributes)
                    elif hasattr(result_stream, '_usage') and result_stream._usage:
                        usage = result_stream._usage
                        logger.info(f"Got usage from result_stream._usage: {usage}")
                    else:
                        logger.warning(f"No usage information found. Agent model_settings: {agent.model_settings}")
                        
                    if usage:
                        # Get the model name from the agent
                        model_name = None
                        if hasattr(agent, 'model') and isinstance(agent.model, str):
                            model_name = agent.model
                            
                        # Calculate cost
                        from .costs import dollars_for_usage
                        cost = dollars_for_usage(usage, model_name_from_agent=model_name)
                        
                        # Store in context state
                        input_tokens = getattr(usage, "input_tokens", 0) or 0
                        output_tokens = getattr(usage, "output_tokens", 0) or 0
                        total_tokens = input_tokens + output_tokens
                        
                        ctx.state["last_run_cost"] = cost
                        ctx.state["last_run_usage"] = {
                            "input_tokens": input_tokens,
                            "output_tokens": output_tokens,
                            "total_tokens": total_tokens,
                            "model_name": model_name or get_active_provider()
                        }
                        logger.info(f"Stored cost (${cost:.6f}) and usage ({total_tokens} tokens) in context state")
                except Exception as cost_err:
                    logger.error(f"Error calculating/storing cost: {cost_err}", exc_info=True)
            except Exception as e:
                logger.error(f"Error saving conversation history: {e}", exc_info=True)
                ctx.state["conversation_history"] = ctx.state.get("conversation_history", [])

        # Info logs after the run
        # Cost printout removed: only show cost at session end or on explicit /cost command

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
        if excel_assistant_agent:
             patch_tool_schemas(excel_assistant_agent) # Apply patch after initial creation
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

                elif cmd in ["exit", "quit"]:
                    print("\n\033[95m--- Current System Prompt ---\033[0m")
                    if excel_assistant_agent:
                        try:
                            wrapper = RunContextWrapper(context=app_context)
                            sys_prompt = _dynamic_instructions(wrapper, excel_assistant_agent)
                            print(sys_prompt)
                        except Exception as err:
                            print(f"Error showing system prompt: {err}")
                    print("\033[95m--- End System Prompt ---\n")
                    # Print cost summary at session end if available
                    last_cost = app_context.state.get("last_run_cost")
                    usage = app_context.state.get("last_run_usage", {})
                    tokens = usage.get("total_tokens", "N/A")
                    model_used = usage.get("model_name", get_active_provider())
                    if last_cost is not None:
                        print(f"\033[96m💰 Session cost: ${last_cost:.4f} ({tokens} tokens, Model: {model_used})\033[0m")
                    else:
                        print("\033[96m💰 Session cost: $0.0000 (no usage recorded)\033[0m")
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
                        
                        # Try to manually look for usage info and calculate cost on the spot
                        usage_found = False
                        if hasattr(app_context, 'usage') and app_context.usage:
                            usage_found = True
                            input_tokens = getattr(app_context.usage, "input_tokens", 0) or 0
                            output_tokens = getattr(app_context.usage, "output_tokens", 0) or 0
                            total_tokens = input_tokens + output_tokens
                            print(f"\033[93mFound usage directly on context: Input={input_tokens}, Output={output_tokens}, Total={total_tokens}\033[0m")
                            
                            # Try to calculate cost right now
                            try:
                                from .costs import dollars_for_usage
                                model_name = None
                                if hasattr(excel_assistant_agent, 'model') and isinstance(excel_assistant_agent.model, str):
                                    model_name = excel_assistant_agent.model
                                    print(f"\033[93mModel from agent: {model_name}\033[0m")
                                
                                cost = dollars_for_usage(app_context.usage, model_name_from_agent=model_name)
                                print(f"\033[92mCalculated cost now: ${cost:.4f}\033[0m")
                                
                                # Store it for future reference
                                app_context.state["last_run_cost"] = cost
                                app_context.state["last_run_usage"] = {
                                    "input_tokens": input_tokens,
                                    "output_tokens": output_tokens,
                                    "total_tokens": total_tokens,
                                    "model_name": model_name or get_active_provider()
                                }
                                print("\033[92mStored cost information in context state\033[0m")
                            except Exception as calc_err:
                                print(f"\033[91mError calculating cost: {calc_err}\033[0m")
                        
                        if not usage_found:
                            print("\033[93mNo usage information found. Try running a query first.\033[0m")
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
                            excel_assistant_agent = create_excel_assistant_agent() # Recreate agent
                            if excel_assistant_agent:
                                 patch_tool_schemas(excel_assistant_agent) # PATCH AGAIN after recreation
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