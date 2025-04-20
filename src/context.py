import logging
from dataclasses import dataclass, field, asdict
from typing import Optional, Dict, List, Any
from .excel_ops import ExcelManager
from .constants import WRITE_TOOLS # Add this import

logger = logging.getLogger(__name__)

@dataclass
class WorkbookShape:
    """Represents a lightweight snapshot of the workbook structure."""
    sheets: Dict[str, str] = field(default_factory=dict)  # sheet_name -> used_range (A1:XN)
    headers: Dict[str, List[str]] = field(default_factory=dict) # sheet_name -> list of header strings (row 1)
    names: Dict[str, str] = field(default_factory=dict)   # named_range_name -> A1 reference
    version: int = 0

@dataclass
class AppContext:
    excel_manager: ExcelManager
    # Generic bag for planner / retry metadata, agent state, etc.
    state: dict = field(default_factory=dict) # Holds 'summary', etc.
    shape: Optional[WorkbookShape] = None     # Holds the current workbook shape snapshot
    actions: List[Dict[str, Any]] = field(default_factory=list)  # Rolling ledger of recent tool calls
    max_actions: int = 50  # Keep only the last *N* actions

    def record_action(self, *, tool: Dict[str, Any] | str, args: Dict[str, Any], result: Any, ok: bool) -> None:
        """
        Append a record to the action ledger and truncate to `max_actions`.

        Args:
            tool:   Tool name.
            args:   Arguments passed to the tool.
            result: Raw result returned from the tool.
            ok:     True if the tool succeeded, else False.
        """
        self.actions.append(
            {"tool": str(tool), "args": args, "result": result, "ok": ok}
        )
        if len(self.actions) > self.max_actions:
            self.actions = self.actions[-self.max_actions:]

    def update_shape(self, *, tool_name: Optional[str] = None) -> bool:
        """
        Attempts to refresh the workbook shape state by calling excel_manager.quick_scan_shape,
        unless tool_name indicates a read-only tool was used.
        Handles version incrementing, logging, and error fallback.

        Returns:
            bool: True if the shape was successfully updated or skipped, False only on scan failure.
        """
        previous_shape = self.shape # Keep track of previous shape

        # --- Add this check ---
        if tool_name and tool_name not in WRITE_TOOLS:
            logger.debug(f"Tool '{tool_name}' is not a write tool. Skipping shape scan, keeping previous shape (v{previous_shape.version if previous_shape else 0}).")
            # No actual update occurred, but returning True signifies no error state.
            # Returning False could be misinterpreted as a scan failure.
            return True # Indicates the context state is considered valid, even if unchanged.
        # --- End check ---

        logger.debug(f"Executing shape scan (tool='{tool_name or 'Initial/Manual'}')...")
        try:
            # Delegate scanning to the ExcelManager
            new_shape = self.excel_manager.quick_scan_shape()

            # Increment version based on previous state
            if previous_shape:
                new_shape.version = previous_shape.version + 1
            else:
                new_shape.version = 1 # Initial version

            self.shape = new_shape # Update the shape on success
            logger.info(
                 f"Workbook shape updated (v{new_shape.version}): "
                 f"{len(new_shape.sheets)} sheets, "
                 f"{sum(len(h) for h in new_shape.headers.values())} headers found."
            )
            return True
        except Exception as e:
            logger.warning(
                f"Failed to update workbook shape: {e}. "
                f"Using previous shape (v{previous_shape.version if previous_shape else 0})."
            )
            # Keep the old shape if the scan fails
            return False

    def dump_state_to_json(self, file_path: str = "state_dump.json"):
        """Dumps the current shape and agent state to a JSON file."""
        import json
        try:
            full_state = {
                "shape": asdict(self.shape) if self.shape else None,
                "agent_state": self.state,  # Contains "summary", etc.
                "actions": self.actions,    # Rolling action ledger
            }
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(full_state, f, indent=2)
            logger.debug(f"Full state (shape + agent_state) dumped to {file_path}")
        except Exception as e:
            logger.error(f"Failed to dump state to {file_path}: {e}")