"""Commandâ€‘line interface for the Autonomous Excel Assistant (single realâ€‘time mode)."""

from __future__ import annotations

import argparse
import asyncio
import logging
import os
import sys
import time
from collections import deque

# Environment switch – set OPENAI_SHOW_COST=0 to disable cost displays
SHOW_COST = os.getenv("OPENAI_SHOW_COST", "1") == "1"
from typing import Optional, List, Dict

from dotenv import load_dotenv
from agents import Runner, ItemHelpers
from agents.stream_events import RunItemStreamEvent
from openai.types.responses import ResponseTextDeltaEvent

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
#  RichÂ powerâ€‘ups (colour, panels, prompts â€¦)
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
from rich.console import Console, Group
from rich.layout import Layout
from rich.logging import RichHandler

# ------------------------------------------------------------------
# Stop log "echoâ€ť to stderr; only warnings+ should bypass Live panel
# ------------------------------------------------------------------
RichHandler.level = logging.WARNING
from rich.panel import Panel
from rich.prompt import Prompt
from rich.live import Live
import threading, queue, getchlib

# ------------------------------------------------------------------
# 1. Global helpers for the realâ€‘time input buffer
# ------------------------------------------------------------------
_KEY_QUEUE: "queue.Queue[str]" = queue.Queue()

class _InputBuffer:
    """Holds the text the user is currently typing."""
    def __init__(self) -> None:
        self._chars: list[str] = []
        self.lock = threading.Lock()

    def append(self, ch: str) -> None:
        with self.lock:
            self._chars.append(ch)

    def backspace(self) -> None:
        with self.lock:
            if self._chars:
                self._chars.pop()

    def clear(self) -> str:
        with self.lock:
            out = "".join(self._chars)
            self._chars.clear()
            return out

    def __str__(self) -> str:
        with self.lock:
            return "".join(self._chars)

_input_buffer = _InputBuffer()

# ------------------------------------------------------------------
# 2. Background thread that reads one key at a time
# ------------------------------------------------------------------
def _keyboard_loop(render_input_cb, live) -> None:
    while True:
        # Read a key *without* echoing it to the terminal. When echo is enabled,
        # the character is briefly printed outside the Live layout and then
        # cleared on each refresh, which made the input appear to "disappearâ€ť.
        ch = getchlib.getkey(echo=False)
        if ch in ("\x03", "\x04"):          # Ctrlâ€‘C / Ctrlâ€‘D â€“ exit app
            _KEY_QUEUE.put_nowait("__EXIT__")
            break
        if ch in ("\r", "\n"):              # ENTER â€“ submit current buffer
            _KEY_QUEUE.put_nowait(_input_buffer.clear())
        elif ch in ("\x7f", "\b"):          # Backspace
            _input_buffer.backspace()
        elif ch.isprintable():
            _input_buffer.append(ch)

        # Reâ€‘paint the input panel *immediately* from any thread
        # Pass the latest buffer contents explicitly to avoid stale snapshots.
        # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        # Rich â‰ĄÂ 13.4 ships `Console.call_from_thread`, which safely queues a
        # drawâ€‘call onto the main render loop.  On older Rich versions the
        # attribute is missing â†’ calling it raises `AttributeError` inside this
        # background thread, killing the thread and causing the user's typed
        # characters to "disappearâ€ť.  We therefore:
        #   1. Prefer the modern helper when present.
        #   2. Gracefully fall back to calling the renderer directly when absent.
        # Because `_keyboard_loop` already executes off the main thread, the
        # direct call is safe and keeps the input panel responsive.
        # â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
        try:
            live.console.call_from_thread(render_input_cb, str(_input_buffer))
        except AttributeError:
            # Fallback path for RichÂ <Â 13.4
            render_input_cb(str(_input_buffer))
from rich.text import Text

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
#  Project imports
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
from .agent_core import excel_assistant_agent
from .context import AppContext
from .excel_ops import ExcelManager


console = Console(highlight=False)

# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
#  Argument parsing
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="Autonomous Excel Assistant (realâ€‘time)")
    p.add_argument("--input-file", type=str, help="Path to an existing workbook.")
    p.add_argument("--output-file", type=str, help="Path to save at the end (optional).")
    p.add_argument("--instruction", type=str, help="Oneâ€‘shot instruction for the agent.")
    p.add_argument("--interactive", "-i", action="store_true", help="Interactive chat mode.")
    p.add_argument("--stream", action="store_true", help="Stream the agent's reasoning/output.")
    p.add_argument("-v", "--verbose", action="store_true", help="Verbose (DEBUG) logging.")
    p.add_argument("--trace-off", action="store_true", help="Disable OpenAI tracing for this run.")
    p.add_argument("--hide-stats", action="store_true",
                   help="Suppress token & cost summary after each run.")
    return p.parse_args()


# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
#  Layout-based logging handler
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
class LiveLogHandler(logging.Handler):
    """
    Custom handler that keeps a rolling log in memory
    and displays it in a fixed top panel of the Layout.
    """

    def __init__(self, live: Live, layout: Layout, max_records: int = 100) -> None:
        super().__init__()
        self.records = deque(maxlen=max_records)
        self.live = live
        self.layout = layout

    def emit(self, record: logging.LogRecord) -> None:
        # Keep only INFO / DEBUG in the top panel; warnings+ will go to console
        if record.levelno > logging.INFO:
            return

        msg = self.format(record)
        level_style = {
            logging.DEBUG: "dim cyan",
            logging.INFO: "green",
        }.get(record.levelno, "white")

        # Deâ€‘dupe consecutive identical INFO/DEBUG lines
        if not self.records or self.records[-1].plain != msg:
            self.records.append(Text(msg, style=level_style))

        # Update the "logs" layout panel
        def _refresh() -> None:
            self.layout["logs"].update(
                Panel(
                    Text("\n").join(self.records),
                    title="Logs",
                    border_style="blue",
                    padding=(0, 1),
                    height=8,
                )
            )
            self.live.update(self.layout, refresh=True)

        # Attempt a thread-safe call if supported
        try:
            self.live.console.call_from_thread(_refresh)
        except AttributeError:
            _refresh()

    def close(self) -> None:
        try:
            # Final update before stopping
            self.layout["logs"].update(
                Panel(
                    Text("\n").join(self.records),
                    title="Logs",
                    border_style="blue",
                    padding=(0, 1),
                    height=8,
                )
            )
            self.live.update(self.layout, refresh=True)
            self.live.stop()
        except Exception:
            pass
        super().close()


# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
#  Stream helper
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
async def handle_streaming(result, verbose: bool):
    """
    Prettyâ€‘print streaming events if requested. In nonâ€‘interactive,
    we show the agent's final output or incremental tokens.
    """
    logger = logging.getLogger(__name__)
    final_output = ""
    last_msg: Optional[str] = None

    try:
        async for event in result.stream_events():
            # Raw LLM deltas
            if (
                event.type == "raw_response_event"
                and isinstance(event.data, ResponseTextDeltaEvent)
            ):
                delta = event.data.delta
                final_output += delta
                if last_msg is None:
                    last_msg = final_output
                if verbose:
                    console.print(delta, end="")
                continue

            # Detailed events from the agent
            if event.type != "run_item_stream_event":
                continue

            item: RunItemStreamEvent = event.item
            if item.type == "tool_call_item" and verbose:
                console.print(f"[cyan]đź› ď¸Ź  {item.function_name}[/]")
            elif item.type == "tool_call_output_item" and verbose:
                ok = "âś”" if "error" not in item.output else "âś–"
                colour = "green" if ok == "âś”" else "red"
                console.print(f"   â†ł [{colour}]{ok} {item.output}[/{colour}]")
            elif item.type == "message_output_item":
                text = ItemHelpers.text_message_output(item)
                last_msg = text
                if verbose:
                    console.print(f"[magenta]đź’¬ {text}[/]")
    except Exception as e:  # pragma: no cover
        logger.error("Streaming error: %s", e, exc_info=True)
        final_output += f"\n\n[Streaming error: {e}]"

    class _Result:
        def __init__(self, text: str):
            self.final_output = text

    return _Result(last_msg or final_output)


# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
#  Main
# â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
async def main() -> None:
    load_dotenv()
    if not os.getenv("OPENAI_API_KEY"):
        console.print("[bold red]âťŚ OPENAI_API_KEY not set â€“ aborting.[/]")
        sys.exit(1)

    args = parse_args()

    # Build a Layout with top logs (fixed height) & bottom chat
    layout = Layout(name="root")
    # Top â†’ Logs (fixed), middle â†’ Chat (flex), bottom â†’ Input (fixed)
    layout.split_column(
        Layout(name="logs",  size=8),
        Layout(name="stats", size=3),     # NEW
        Layout(name="chat",  ratio=1),
        Layout(name="input", size=3),
    )

    # Start a Live display with our layout (alternate screen avoids flicker)
    live = Live(
        layout,
        console=console,
        refresh_per_second=10,
        screen=True,
        redirect_stdout=False,   # donâ€™t let prints escape the layout
    )
    live.__enter__()  # We'll manually manage the context

    # Build a logging handler to feed the top panel
    live_handler = LiveLogHandler(live=live, layout=layout)
    rich_handler = RichHandler(rich_tracebacks=True, console=console, markup=True)
    # Only show WARNING+ in the scrolling console
    rich_handler.setLevel(logging.WARNING)

    # Remove *all* preâ€‘existing handlers to avoid duplicates, then install ours
    root = logging.getLogger()
    root.handlers.clear()
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(message)s",
        handlers=[live_handler, rich_handler],
        force=True,        # PythonÂ 3.13+; on 3.11/3.12 we did .handlers.clear()
    )
    logger = logging.getLogger(__name__)

    if args.verbose:
        logger.debug("Verbose logging enabled (DEBUG level).")

    # We'll store chat messages in a list for the bottom panel
    chat_history: List[Dict[str, str]] = []

    def render_chat_panel() -> None:
        """
        Update the 'chat' panel in the layout.

        Each utterance is rendered in its own bordered Panel titled
        "Youâ€ť or "Assistantâ€ť, stacked vertically so the interface resembles
        a classic chat transcript while the Logs panel remains fixed above.
        """
        message_panels: List[Panel] = []
        for msg in chat_history:
            role = msg.get("role", "system")
            text = msg.get("content", "")

            if role == "user":
                panel = Panel(
                    Text(text),
                    title="You",
                    border_style="cyan",
                    padding=(0, 1),
                )
            elif role == "assistant":
                panel = Panel(
                    Text(text),
                    title="Assistant",
                    border_style="magenta",
                    padding=(0, 1),
                )
            else:  # e.g. system or other roles
                panel = Panel(
                    Text(text),
                    title=role.capitalize(),
                    border_style="grey58",
                    padding=(0, 1),
                )

            message_panels.append(panel)

        # Keep only the most recent 50 messages to avoid unbounded growth
        message_panels = message_panels[-50:]

        # Stack the panels in the chat area
        layout["chat"].update(Group(*message_panels))
        render_stats_panel()  # keep stats in sync
        live.update(layout, refresh=True)

    # Stats panel render helper
    def render_stats_panel() -> None:
        """
        Update the small 'stats' panel with token and cost info.
        Hidden when --hide-stats is set or OPENAI_SHOW_COST=0.
        """
        if args.hide_stats or not SHOW_COST:
            layout["stats"].update(Text(""))  # keep height stable
            return

        u    = ctx.state.get("last_run_usage", {})
        cost = ctx.state.get("last_run_cost", 0.0)
        body = (
            f"Tokens: {u.get('total_tokens',0)} "
            f\"(in={u.get('input_tokens',0)}, out={u.get('output_tokens',0)})\\n\"
            f\"Cost:   ${cost:,.4f}\"
        )
        layout["stats"].update(
            Panel(Text(body), title="Stats", border_style="yellow", padding=(0,1))
        )

    # new input panel render helper
    def render_input_panel(content: str = "") -> None:
        """
        Keep the user's current input visible *inside* the layout.
        """
        layout["input"].update(
            Panel(
                Text(content if content else " ", overflow="fold"),
                title="You",
                border_style="cyan",
                padding=(0, 1),
            )
        )
        live.update(layout, refresh=True)

    # â”€â”€â”€â”€â”€â”€â”€â”€â”€ Excel session â”€â”€â”€â”€â”€â”€â”€â”€â”€
    try:
        async with ExcelManager(
            file_path=args.input_file,
            visible=True,
            attach_existing=True,
            single_workbook=True,
        ) as excel_mgr:
            ctx = AppContext(excel_manager=excel_mgr)
            ctx.update_shape()  # first shape scan
            render_stats_panel()           # initial blank box

            if args.input_file:
                console.print(f"đź“‚ Opened workbook: [bold]{args.input_file}[/]")
            else:
                console.print("đź†• Started new workbook.")

            # ============ INTERACTIVE ============
            if args.interactive:
                chat_history.append(
                    {
                        "role": "assistant",
                        "content": "Hello! How can I help you today?",
                    }
                )
                render_chat_panel()
                render_input_panel("")   # clear input box
                render_input_panel("")  # seed empty input box
                console.print(
                    "(Type your question. Blank line to submit. 'exit' to quit.)"
                )

                # --- NEW realâ€‘time input loop --------------------
                kb_thread = threading.Thread(target=_keyboard_loop, args=(render_input_panel, live), daemon=True)
                kb_thread.start()

                while True:
                    try:
                        # Block until ENTER or Ctrlâ€‘D / Ctrlâ€‘C
                        user_msg = _KEY_QUEUE.get()
                        if user_msg == "__EXIT__":
                            console.print("\n[bold red]Goodâ€‘bye đź‘‹[/]")
                            return
                        if not user_msg.strip():
                            render_input_panel("")   # clear input box
                            continue
                        if user_msg.lower() in {"exit", "quit"}:
                            break

                        # Add user's message
                        chat_history.append({"role": "user", "content": user_msg})
                        render_chat_panel()
                        render_input_panel("")  # reset input box

                        # Invoke the agent
                        res = await Runner.run(
                            excel_assistant_agent,
                            input=user_msg,
                            context=ctx,
                            max_turns=25,
                        )
                        await excel_mgr.ensure_changes_applied()
                        reply = res.final_output

                        # Add agent's answer
                        chat_history.append({"role": "assistant", "content": reply})
                        render_chat_panel()

                    except (KeyboardInterrupt, EOFError):
                        console.print("\n[bold red]Goodâ€‘bye đź‘‹[/]")
                        break
                    except Exception as e:
                        logger.error("Error during agent run: %s", e, exc_info=True)
                        console.print(f"[bold red]Error:[/] {e}")

                return

            # ============ ONE-SHOT ============
            if not args.instruction:
                console.print(
                    "[bold red]âš  No instruction provided. Use --instruction or -i for interactive mode.[/]"
                )
                sys.exit(1)

            logger.info("đź’ˇ Instruction: %s", args.instruction)
            start = time.monotonic()

            # Add the user instruction to chat panel for clarity
            chat_history.append({"role": "user", "content": args.instruction})
            render_chat_panel()

            try:
                if args.stream:
                    streamed = await Runner.run_streamed(
                        excel_assistant_agent,
                        input=args.instruction,
                        context=ctx,
                        max_turns=25,
                    )
                    result = await handle_streaming(streamed, args.verbose)
                else:
                    result = await Runner.run(
                        excel_assistant_agent,
                        input=args.instruction,
                        context=ctx,
                        max_turns=25,
                    )
            except Exception as e:
                logger.error("âťŚ Agent error: %s", e, exc_info=True)
                sys.exit(1)

            await excel_mgr.ensure_changes_applied()

            # ── Plain cost summary ───────────────────────────────
            if SHOW_COST and not args.hide_stats:
                u     = ctx.state.get("last_run_usage", {})
                cost  = ctx.state.get("last_run_cost", 0.0)
                console.print(
                    f"[bold green]Tokens:[/] {u.get('total_tokens', 0)} "
                    f"(in={u.get('input_tokens',0)}, out={u.get('output_tokens',0)})"
                )
                console.print(f"[bold yellow]Cost:[/] ${cost:,.4f}")

            reply = result.final_output
            elapsed = time.monotonic() - start

            logger.info("Done in %.1fs", elapsed)

            # Show final output in chat
            chat_history.append({"role": "assistant", "content": reply})
            render_chat_panel()

            # ---------- Optional save ----------
            if args.output_file:
                try:
                    saved = await excel_mgr.save_with_confirmation(args.output_file)
                    console.print(f"[bold green]đź’ľ Workbook saved to {saved}[/]")
                except Exception as e:
                    logger.error("Failed to save workbook: %s", e, exc_info=True)
                    console.print(f"[bold red]âťŚ Save failed:[/] {e}")

    except Exception as e:  # pragma: no cover
        logger.error("Fatal error: %s", e, exc_info=True)
        console.print(f"[bold red]âťŚ Fatal error:[/] {e}")
        sys.exit(1)
    finally:
        # Stop live logging
        live_handler.close()
        live.__exit__(None, None, None)


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        console.print("\nExiting.")
        sys.exit(0)