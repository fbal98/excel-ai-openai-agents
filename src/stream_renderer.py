"""
Stream-renderer for the Excel-AI CLI.

`format_event(event) -> str | None`
-----------------------------------
Turns a single streaming *event* emitted by the OpenAI **Agents** SDK into a
human-friendly one-liner:

• Assistant message chunks  → «🤖 Agent: …» (re-assembled before printing)
• Tool call start           → «🛠️ Tool: name(arg=val, …)»
• Tool result / end         → «🛠️ Tool ✔ {...}» or «🛠️ Tool ✗ {...}»

Anything else (thoughts, unknown events) is ignored.
The helper is deliberately defensive – it duck-types the event object so minor
SDK updates won’t break the CLI.
"""

from __future__ import annotations

from typing import Any, Dict, Optional

# Keep try/except for RICH_AVAILABLE check, but Text is not used here directly
try:
    from rich.text import Text # Keep import for type checking if needed elsewhere
    RICH_AVAILABLE = True
except ImportError:
    RICH_AVAILABLE = False

# ─────────────────────────────────────────────────────────────────────────────
#  Helpers
# ─────────────────────────────────────────────────────────────────────────────


def _event_to_dict(ev: Any) -> Dict[str, Any]:
    """Best-effort conversion of an arbitrary SDK event into a plain dict."""
    if isinstance(ev, dict):
        return ev
    if hasattr(ev, "__dict__"):
        return vars(ev)
    for attr in ("model_dump", "to_dict", "dict"):
        fn = getattr(ev, attr, None)
        if callable(fn):
            try:
                return fn()  # type: ignore[func-returns-value]
            except Exception:  # pragma: no cover
                pass
    return {"unserialisable_event": repr(ev)}


# ─────────────────────────────────────────────────────────────────────────────
#  Module-level state – we buffer assistant token deltas until a flush point.
# ─────────────────────────────────────────────────────────────────────────────
_ASSISTANT_BUFFER: list[str] = []


def _flush_assistant_buffer(is_final: bool = False) -> Optional[str]:
    """
    Return buffered assistant text as a formatted string and clear the buffer.
    Prepends success checkmark if final.
    """
    if not _ASSISTANT_BUFFER:
        return None
    joined = "".join(_ASSISTANT_BUFFER).strip()
    _ASSISTANT_BUFFER.clear()

    # No need to format if empty
    if not joined:
        return None

    prefix = "✔️ " if is_final else ""
    # ANSI Green for agent message
    agent_prefix = f"\033[92m{prefix}🤖 Agent: "
    reset_code = "\033[0m"
    indent = " " * (len(prefix) + len("🤖 Agent: ")) # Indent based on visible prefix length

    # Handle multi-line responses
    lines = joined.splitlines()
    formatted_lines = []
    first_line = True
    for line in lines:
        if first_line:
            # Apply color only to the first line with prefix
            formatted_lines.append(f"{agent_prefix}{line}{reset_code}")
            first_line = False
        else:
            # Apply color to subsequent lines, maintaining indentation
            formatted_lines.append(f"\033[92m{indent}{line}{reset_code}" if line.strip() else "")

    final_text = "\n".join(formatted_lines)
    return final_text


# ─────────────────────────────────────────────────────────────────────────────
#  Public API
# ─────────────────────────────────────────────────────────────────────────────
def format_event(ev: Any) -> Optional[str]:  # noqa: D401
    """
    Convert *ev* to a printable string **without** trailing newline,
    using ANSI codes for basic coloring.

    Returns ``None`` when the event should not be shown to the user.
    """
    evd = _event_to_dict(ev)

    kind = (
        evd.get("kind")
        or evd.get("event_type")
        or evd.get("type")
        or evd.get("category")
    )

    # ── Assistant message chunks ────────────────────────────────────────────
    if kind in {"message_output_item", "assistant_message", "assistant_chunk"}:
        chunk = evd.get("text") or evd.get("delta") or evd.get("content") or ""
        if chunk:
            # Preserve newlines in the original chunks
            _ASSISTANT_BUFFER.append(str(chunk))
        is_final_chunk = evd.get("last") or evd.get("is_finished")
        if is_final_chunk:
            # Pass is_final=True to add checkmark
            return _flush_assistant_buffer(is_final=True)
        return None  # keep buffering

    # Any *other* event flushes pending assistant text first
    # Ensure is_final is False for intermediate flushes
    pending_output = _flush_assistant_buffer(is_final=False)

    # ── Tool call started ───────────────────────────────────────────────────
    if kind in {"tool_call", "tool_start"}:
        name = evd.get("name") or evd.get("tool_name") or evd.get("tool") or "<unknown>"
        args = evd.get("args") or evd.get("arguments") or {}
        if isinstance(args, dict):
            arg_str = ", ".join(f"{k}={repr(v)}" for k, v in args.items())
        else:
            arg_str = str(args)
        # ANSI Yellow for tool start
        line = f"\033[93m🛠️ Tool: {name}({arg_str})\033[0m"

        # Combine pending output (if any) with the new line
        if pending_output:
            return f"{pending_output}\n{line}"
        else:
            return line

    # ── Tool result / end ───────────────────────────────────────────────────
    if kind in {"tool_end", "tool_result"}:
        result = evd.get("result") or evd.get("output") or evd.get("data")
        success = True
        if isinstance(result, dict) and (result.get("success") is False or result.get("error")):
            success = False
        status_icon = "✔" if success else "✗"
        color_code = "\033[92m" if success else "\033[91m" # Green or Red
        reset_code = "\033[0m"
        line = f"{color_code}🛠️ Tool {status_icon} {result}{reset_code}"

        # Combine pending output (if any) with the new line
        if pending_output:
            return f"{pending_output}\n{line}"
        else:
            return line

    # ── Thoughts are now shown in cyan with a 💭 prefix ──────────────
    if kind in {"assistant_thought", "thought"}:
        thought_text = evd.get("text") or evd.get("content") or ""
        if thought_text:
            return f"\033[96m💭 Thought: {thought_text}\033[0m"
        return pending_output  # Show any flushed text if no thought text

    # ── Unknown event – just return any pending assistant text ──────────────
    # If we reached here, it's an unknown event type
    return pending_output # Return any flushed text