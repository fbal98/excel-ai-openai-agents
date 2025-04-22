# CLAUDE.md - Project Guide (Excel AI Agent)

Quick reference for working on the `excel-ai-openai-agents` project.

## Goal

AI agent using `xlwings` to modify Excel workbooks based on natural language. Features CLI and TUI.

## Core Modules

* **Entry Points**: `src/cli.py`, `src/__main__.py` (Handles args, init, starts UI or agent)
* **Agent Definition**: `src/agent_core.py` (`Agent`, system prompt, tool list, model, dynamic instructions)
* **Tools**: `src/tools.py` (`@function_tool` decorated functions for agent actions)
* **Excel Interface**: `src/excel_ops.py` (`ExcelManager` class wrapping `xlwings`)
* **Runtime Context**: `src/context.py` (`AppContext` holding state, `ExcelManager`, `WorkbookShape`)
* **Agent Hooks**: `src/hooks.py` (Manage state, metrics, shape updates, error limits)
* **UI**: `src/ui/` (TUI modules using `prompt_toolkit` & `rich`)
* **Constants**: `src/constants.py`, `src/debounce_constants.py` (Tool lists, config)

## Quick Commands

* **Run TUI**: `python -m src --input-file <path> [options]`
* **Run CLI (one-shot)**: `python -m src --instruction "..." --no-ui [options]`
* **Run CLI (interactive)**: `python -m src -i --no-ui [options]`
* **Install**: `pip install -r requirements.txt`

## Coding Guidelines

* **Style**: Standard Python style (snake_case, CamelCase, UPPER_CASE). Max 100 chars/line.
* **Types**: Full type annotations required (`typing`).
* **Docs**: Docstrings for public APIs.
* **Errors**: Use `try/except`, `logging`, and specific tool error returns (see below).
* **Logging**: Standard `logging`. TUI uses `UILogHandler`.

## Tool Development (`src/tools.py`)

* **Naming**: Must end with `_tool`.
* **Decorator**: Must use `@function_tool`.
* **Signature**: `def my_tool(ctx: RunContextWrapper[AppContext], arg1: type, ...) -> ToolResult:`
    * Access context: `ctx.context` (`AppContext`)
    * Access Excel: `ctx.context.excel_manager`
* **Return**: Must conform to `ToolResult` (`TypedDict` from `tools.py`):
    * **Success**: `{"success": True}` or `{"success": True, "data": ...}`
    * **Failure**: `{"success": False, "error": "Reason"}`
    * `_ensure_toolresult` helper normalizes simpler returns (like `True`, `None`, `dict` with just `"error"`).
* **Interaction**: Use `ctx.context.excel_manager` methods for Excel actions.
* **Validation**: Validate inputs early, return error dict on failure.
* **Bulk Ops**: Prefer bulk tools (`set_cell_values`, `insert_table`) over loops.

## State & Context (`src/context.py`)

* `AppContext`: Central object passed around. Holds `ExcelManager`, `state` dict, `shape` (`WorkbookShape`), `actions` list, `metrics`.
* `WorkbookShape`: Summary of sheets, headers, names. Fed into agent prompt.
* **Shape Updates**: Managed by `hooks.py` (`ActionLoggingHooks`). Triggered by `WRITE_TOOLS` / `STRUCTURAL_WRITE_TOOLS`, debounced.

## Excel Interaction (`src/excel_ops.py`)

* `ExcelManager`: Handles all `xlwings` interaction. Use its methods inside tools.
* Supports snapshots & revert. Async context manager (`__aenter__`, `__aexit__`).

## Prompting (`src/agent_core.py`)

* `SYSTEM_PROMPT`: Defines agent behavior, tool usage rules.
* `_dynamic_instructions`: Adds current `<workbook_shape>` and `<progress_summary>` to the prompt.

## UI (`src/ui/`)

* Built with `prompt_toolkit` and `rich`.
* Key files: `app.py` (layout/app logic), `log_handler.py` (TUI logging), `renderers.py` (message styling).