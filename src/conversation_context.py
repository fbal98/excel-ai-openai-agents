# src/conversation_context.py
from __future__ import annotations
import copy
import difflib, json, logging, tiktoken
from typing import Any, Dict, List, Literal, Optional, TYPE_CHECKING
from agents import RunResult # Changed to RunResult as per the error message
from agents.results import RunResultBase

if TYPE_CHECKING:
    from .context import AppContext, WorkbookShape # Avoid circular import

logger = logging.getLogger(__name__)

Role = Literal["assistant", "system", "user"] # CHANGED

def _token_len(text: str, model: str = "gpt-4o-mini") -> int:
    try:
        enc = tiktoken.encoding_for_model(model)
        return len(enc.encode(text))
    except Exception: # Handle cases where model name might not be found
        logger.warning(f"Could not get tiktoken encoding for model '{model}', falling back to simple length.")
        return len(text)

class ConversationContext:
    """Utilities for emitting, deduplicating & pruning context messages."""
    DEFAULT_MAX_TOKENS = 1500 # Max tokens for conversation history before pruning
    KEEP_LAST_N_MSGS = 8    # Keep this many recent messages unsummarized

    @staticmethod
    def emit(ctx: 'AppContext', *, role: Role, content: str) -> None:
        """Append a message to conversation_history, deduplicating if needed."""
        hist: List[Dict[str, str]] = ctx.state.setdefault("conversation_history", [])
        # Deduplicate identical consecutive messages
        if hist and hist[-1].get("role") == role and hist[-1].get("content") == content:
            logger.debug("Skipping duplicate consecutive message emission.")
            return
        hist.append({"role": role, "content": content})
        logger.debug(f"Emitted {role} message (len: {len(content)}): {content[:100]}...")

    @staticmethod
    def emit_shape_delta(ctx: 'AppContext', old_shape: Optional['WorkbookShape'], new_shape: 'WorkbookShape') -> None:
        """Generate and emit a shape delta message."""
        if not new_shape:
            logger.warning("emit_shape_delta called with no new_shape.")
            return
        diff = ConversationContext._shape_diff(old_shape, new_shape)
        ConversationContext.emit(ctx, role="assistant", content=diff)

    @staticmethod
    def emit_progress_line(ctx: 'AppContext', line: str) -> None:
        """Wrap text in progress tags and emit."""
        ConversationContext.emit(
            ctx,
            role="assistant",
            content=f"<progress_summary>\n{line}\n</progress_summary>"
        )

    @staticmethod
    def emit_tool_failure(ctx: 'AppContext', tool_name: str, error_msg: str) -> None:
        """Emit a specific tool failure message."""
        ConversationContext.emit(
            ctx,
            role="assistant",
            content=f"<tool_failure tool={tool_name}>{error_msg}</tool_failure>"
        )

    @staticmethod
    def maybe_prune(ctx: 'AppContext', *, model: str = "gpt-4o-mini") -> None:
        """Prune conversation_history if it exceeds token limits."""
        hist: List[Dict[str, str]] = ctx.state.get("conversation_history", [])
        if not hist:
            return

        # Get the model name from context state if available (set by CLI/runner)
        # This helps use the correct tokenizer for the active model.
        model_name_from_state = ctx.state.get("last_run_usage", {}).get("model_name")
        if model_name_from_state and model_name_from_state != "Unknown":
            model = model_name_from_state

        total = sum(_token_len(m.get("content", ""), model) for m in hist)
        if total < ConversationContext.DEFAULT_MAX_TOKENS:
            return

        if len(hist) <= ConversationContext.KEEP_LAST_N_MSGS:
            logger.debug("History is short but over token limit; not pruning.")
            return

        # Summarise oldest messages, keeping the tail untouched
        prune_index = len(hist) - ConversationContext.KEEP_LAST_N_MSGS
        head_to_summarize = hist[:prune_index]
        tail_to_keep = hist[prune_index:]

        # Only proceed if there are messages to summarize
        if not head_to_summarize:
            logger.debug("Pruning condition met, but no messages old enough to summarize.")
            return

        summary_text = ConversationContext._summarise_chunks(
            [m.get("content", "") for m in head_to_summarize]
        )

        # Replace the head with the summary message
        ctx.state["conversation_history"] = [{"role": "system", "content": summary_text}] + tail_to_keep

        pruned_total = sum(_token_len(m.get("content", ""), model) for m in ctx.state["conversation_history"])
        logger.info(
            "Pruned conversation_history from %d messages (%d tokens) to %d messages (%d tokens) using model '%s' for counting.",
            len(hist), total, len(ctx.state["conversation_history"]), pruned_total, model
        )

    @staticmethod
    def update_history_from_result(
        ctx: "AppContext",
        result: RunResultBase,
        original_input: Any,
    ) -> None:
        """
        Merge the latest turn into ctx.state['conversation_history'].

        * Adds the user message for this turn (if not already present).
        * Appends assistant `message_output_item`s.
        * Skips tool-call chatter to keep history tidy.
        """
        hist: list[dict[str, str]] = ctx.state.setdefault("conversation_history", [])

        # ── 1 ▸ ensure current user message present ──────────────────────
        if isinstance(original_input, str):
            user_msg = {"role": "user", "content": original_input}
        elif isinstance(original_input, list):
            # assume last user message is the current one
            user_msg = next((
                m for m in reversed(original_input)
                if isinstance(m, dict) and m.get("role") == "user"
            ), None)
        else:
            user_msg = None

        if user_msg and (not hist or hist[-1] != user_msg):
            hist.append(user_msg)

        # ── 2 ▸ add assistant outputs ────────────────────────────────────
        from .stream_renderer import _normalize_content # local import avoids cycle

        new_items: list[dict[str, str]] = []
        for item in result.new_items:
            if item.type != "message_output_item":
                continue                       # skip tool chatter
            item_dict = item.to_dict()
            content = _normalize_content(item_dict.get("content"))
            if content:
                new_items.append({
                    "role": item_dict.get("role", "assistant"),
                    "content": content,
                })

        if new_items:
            hist.extend(new_items)


    @staticmethod
    def _shape_diff(old: Optional['WorkbookShape'], new: 'WorkbookShape') -> str:
        """Generate a textual diff of the workbook shape."""
        # Deep copy to avoid modifying original shapes if they are mutable objects
        old_dict = copy.deepcopy(old.__dict__) if old else None
        new_dict = copy.deepcopy(new.__dict__)

        # Ensure version is handled correctly
        new_version = new_dict.get("version", 1) # Use getter with default
        if old_dict is None:
            return f"<workbook_shape v={new_version}>\n{json.dumps(new_dict, indent=2)}\n</workbook_shape>"

        old_version = old_dict.get("version", 0) # Use getter with default

        # Remove version key for diffing, as it changes always
        if 'version' in old_dict: del old_dict['version']
        if 'version' in new_dict: del new_dict['version']

        try:
            old_json = json.dumps(old_dict, indent=2, sort_keys=True)
            new_json = json.dumps(new_dict, indent=2, sort_keys=True)

            diff_lines = list(difflib.unified_diff(
                old_json.splitlines(),
                new_json.splitlines(),
                fromfile=f'shape_v{old_version}',
                tofile=f'shape_v{new_version}',
                lineterm=""
            ))

            # If diff is empty (only version changed), return full shape
            if not diff_lines or all(line.startswith(('---', '+++', '@@')) for line in diff_lines):
                 logger.debug("Shape diff was empty or header-only, returning full shape instead.")
                 return f"<workbook_shape v={new_version}>\n{new_json}\n</workbook_shape>"

            # Include only context lines and changes (+/-)
            filtered_diff = [line for line in diff_lines if not line.startswith(('---', '+++', '@@'))]

            return f"<workbook_shape_delta v={new_version} from_v={old_version}>\n" + "\n".join(filtered_diff) + "\n</workbook_shape_delta>"
        except Exception as e:
            logger.error(f"Error generating shape diff: {e}", exc_info=True)
            # Fallback to full shape if diff fails
            return f"<workbook_shape v={new_version}>\n{json.dumps(new_dict, indent=2)}\n</workbook_shape>"


    @staticmethod
    def _summarise_chunks(chunks: List[str]) -> str:
        """Simple heuristic summarizer for old conversation parts."""
        # Limit the number of chunks summarized to avoid overly long summaries
        max_chunks_to_summarize = 30
        chunks_to_process = chunks[-max_chunks_to_summarize:]

        # Take first N chars of each chunk, prioritizing non-empty ones
        summary_lines = []
        for chunk in chunks_to_process:
            stripped_chunk = chunk.strip()
            if stripped_chunk: # Only include non-empty content
                 # Limit length and remove excessive newlines
                summary_line = ' '.join(stripped_chunk[:200].splitlines())
                summary_lines.append(summary_line)

        if not summary_lines:
            return "<summary>Previous conversation context was empty.</summary>"

        # Join non-empty summary lines
        joined_summary = "\n".join(summary_lines)
        return f"<summary>\nSummary of earlier conversation:\n{joined_summary}\n</summary>"