"""
Shared runtime context passed to every tool and hook.

• Keeps a handle to the live ExcelManager so tools can act on the workbook.
• Tracks the latest lightweight snapshot ("shape”) of the workbook so the
  agent can reference sheet/range metadata in its prompt.
"""

from __future__ import annotations

import json
import logging
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
#  Lightweight workbook description passed to the agent prompt
# ---------------------------------------------------------------------------
@dataclass
class WorkbookShape:
    """Sparse summary of the workbook’s structure."""

    sheets: Dict[str, str] = field(default_factory=dict)   # sheet → used‑range (e.g. A1:G42)
    headers: Dict[str, List[str]] = field(default_factory=dict)  # sheet → header list
    names: Dict[str, str] = field(default_factory=dict)    # name → refersTo
    version: int = 1                                       # increments on every change

    # Equality helper so update_shape() can test "changed or not”
    def __eq__(self, other: object) -> bool:  # noqa: D401
        if not isinstance(other, WorkbookShape):
            return NotImplemented
        return (
            self.sheets == other.sheets
            and self.headers == other.headers
            and self.names == other.names
        )


# ---------------------------------------------------------------------------
#  Safe JSON encoder for arbitrary Excel objects in the state dump
# ---------------------------------------------------------------------------
class _SafeEncoder(json.JSONEncoder):
    def default(self, obj):  # noqa: D401, N802
        try:
            return super().default(obj)
        except TypeError:
            return str(obj)


# ---------------------------------------------------------------------------
#  Main context handed to every tool call
# ---------------------------------------------------------------------------
class AppContext:
    """
    Aggregates run‑time data required by the agent, hooks and tools.
    The CLI creates one instance per session and passes it to `Runner.run(...)`.
    """

    def __init__(
        self,
        *,
        excel_manager: Any | None = None,
        state: Optional[Dict[str, Any]] = None,
        shape: Optional[WorkbookShape] = None,
        max_actions: int = 200,
    ) -> None:
        self.excel_manager = excel_manager
        self.state: Dict[str, Any] = state or {}
        self.shape: Optional[WorkbookShape] = shape
        self.actions: List[Dict[str, Any]] = []
        self.max_actions = max_actions

        # Debounce / self‑regulation helpers (see hooks.py)
        self.pending_write_count: int = 0
        self.consecutive_errors: int = 0
        self.last_error_key: tuple[str, str] = ("", "")

    # ---------------------------------------------------------------------
    #  Action history – handy for the agent’s self‑reflection
    # ---------------------------------------------------------------------
    def record_action(
        self,
        *,
        tool: Dict[str, Any] | str,
        args: Dict[str, Any],
        result: Any,
        ok: bool,
    ) -> None:
        if len(self.actions) >= self.max_actions:
            self.actions.pop(0)
        self.actions.append(
            {"tool": tool, "args": args, "result": result, "ok": ok}
        )

    # ---------------------------------------------------------------------
    #  Workbook‑shape refresh
    # ---------------------------------------------------------------------
    def update_shape(self, *, tool_name: Optional[str] = None) -> bool:
        """
        Refresh ``self.shape`` by calling ``excel_manager.quick_scan_shape()``.
        Returns **True** only when a *real* change was detected.

        The scan is skipped gracefully when:

        • ``excel_manager`` is missing.
        • The manager has no ``quick_scan_shape`` attribute.
        • The scan itself raises an exception.
        """
        if self.excel_manager is None:
            logger.debug("update_shape skipped – excel_manager is None")
            return False

        if not hasattr(self.excel_manager, "quick_scan_shape"):
            logger.debug("update_shape skipped – excel_manager lacks quick_scan_shape()")
            return False

        try:
            new_shape: WorkbookShape | None = self.excel_manager.quick_scan_shape()
        except ExcelConnectionError as exc: # Catch specific connection error
             logger.error("update_shape skipped – Connection Error during quick_scan_shape: %s", exc)
             # Potentially reset consecutive errors if connection loss caused scan failure?
             # self.consecutive_errors = 0 # Resetting might hide underlying issues
             return False # Shape update failed due to connection
        except Exception as exc:  # pragma: no cover – depends on xlwings runtime
            logger.warning("quick_scan_shape() failed with unexpected error: %s", exc)
            return False # Shape update failed


        if new_shape is None:
            # This might happen if quick_scan_shape returns None on error internally now
            logger.debug("quick_scan_shape() returned None – potentially due to internal error or no change.")
            return False

        # Increment version whenever the shape *really* changes
        if self.shape is None or new_shape != self.shape:
            new_version = (self.shape.version + 1) if self.shape else 1
            new_shape.version = new_version
            self.shape = new_shape
            logger.debug("Workbook shape updated to v%s", new_version)
            return True

        logger.debug("Workbook shape unchanged – still v%s", self.shape.version)
        return False

    # ---------------------------------------------------------------------
    #  Persist agent & workbook state for debugging
    # ---------------------------------------------------------------------
    def dump_state_to_json(self, file_path: str = "state_dump.json") -> None:
        try:
            with open(file_path, "w", encoding="utf-8") as fp:
                json.dump(
                    {
                        "shape": self.shape.__dict__ if self.shape else None,
                        "state": self.state,
                        "actions": self.actions,
                    },
                    fp,
                    cls=_SafeEncoder,
                    indent=2,
                )
            logger.debug("State dumped to %s", file_path)
        except Exception as exc:
            logger.warning("Failed to dump state to %s: %s", file_path, exc)


# ---------------------------------------------------------------------------
#  Legacy stub kept for backward‑compat (rarely used, but imported elsewhere)
# ---------------------------------------------------------------------------
class context:  # noqa: N801 – legacy camelCase name
    logger = logger