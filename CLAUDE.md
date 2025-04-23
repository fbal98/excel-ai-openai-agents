# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

# Excel AI Agent - Project Guide

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

* **Style**: Standard Python style. Max 100 chars/line. Group imports: stdlib → third-party → local.
* **Naming**: snake_case (functions), CamelCase (classes), UPPER_CASE (constants). Tools must end with `_tool`.
* **Types**: Full type annotations required (`typing`). Return `ToolResult` TypedDict from tools.
* **Errors**: Early validation with descriptive errors. Return `{"success": False, "error": "Reason"}` from tools.
* **Docs**: Google-style docstrings for public APIs. Triple double-quotes.

## Tool Development (`src/tools.py`)

* **Signature**: `def my_tool(ctx: RunContextWrapper[AppContext], arg1: type, ...) -> ToolResult:`
* **Return**: `{"success": True}` or `{"success": True, "data": ...}` for success, `{"success": False, "error": "Reason"}` for failure.
* **Validation**: Validate inputs early, return error dict on failure.
* **Bulk Ops**: Prefer bulk tools (`set_cell_values`, `insert_table`) over loops.

## State & Context (`src/context.py`)

* `AppContext`: Central object passed around. Holds `ExcelManager`, `state` dict, `shape`, `actions`, `metrics`.
* `WorkbookShape`: Summary of sheets, headers, names. Fed into agent prompt.
* **Shape Updates**: Managed by `hooks.py`, triggered by write tools, debounced.

## Known Issues & Solutions

* **Excel Workbook Comparison**: In `ExcelManager.__aenter__` (`excel_ops.py`), never compare workbook objects directly (`wb != self.book`). Instead, compare workbook names to avoid "Object does not exist" errors (macOS). Always check `len(app.books)` before accessing any workbook. This prevents OSERROR -1728.