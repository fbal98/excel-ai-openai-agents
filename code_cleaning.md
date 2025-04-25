# Production-Readiness Plan for Excel AI Agent

This document merges feedback from multiple analyses to provide a comprehensive plan for hardening the Excel AI Agent codebase for production use. It covers packaging, configuration, error handling, testing, architecture, performance, security, and documentation.

## 1. Packaging, Dependencies & Distribution

* **Adopt `pyproject.toml`:** Replace `requirements.txt` with a `pyproject.toml` file following PEP 621 standards for metadata, dependencies, and build configuration.
* **Pin Dependencies:** Use a lock file (e.g., `poetry.lock`, `pdm.lock`, or `requirements.lock` generated via `pip-tools`) to ensure reproducible builds with exact dependency versions.
* **Define Extras:** Split dependencies into optional groups (e.g., `[excel]`, `[cli]`, `[dev]`, `[litellm]`) so users install only what's needed. Move optional imports like `prompt_toolkit` and `getchlib` to relevant extras.
* **Standard Project Layout:** Structure the code within a `src/` directory (e.g., `src/excel_ai`).
* **CLI Entry Point:** Define a script entry point in `pyproject.toml` (e.g., `[project.scripts] excel-ai = "excel_ai.cli:main"`) for easy CLI invocation after installation (`pipx install .` or similar).
* **Document Platform Specifics:** Clearly state any platform limitations (Windows, macOS, Linux), especially regarding `xlwings` behavior and dependencies (`pywintypes`, AppleScript, LibreOffice setup).

## 2. Configuration & Secrets Management

* **Centralized Settings:** Use a library like `pydantic-settings` to manage configuration via a singleton `Settings` object. Load settings from both environment variables and potentially a configuration file (YAML/TOML).
* **Early Validation:** Validate required configuration (API keys, model names) *at startup*. Raise clear, user-friendly errors if essential settings are missing or invalid. Don't allow the application to proceed in a partially configured state.
* **Load `.env` First:** Ensure `.env` files are loaded before any modules that might read environment variables at import time.
* **Inject Configuration:** Pass the `Settings` object or specific configurations to components that need them, rather than having modules read directly from `os.getenv` or global state. Avoid mutable global config state.
* **Secure Secret Handling:**
    * Emphasize secure management of API keys (e.g., `.env` in `.gitignore`, using deployment environment secrets management).
    * Mask secrets automatically in logs using logging filters.
* **Configurable Prompt Version:** Allow the system prompt version (e.g., `system_v1.txt`, `v2.txt`) to be selected via configuration, not hardcoded.
* **Optional Config Check Command:** Consider adding a CLI command (e.g., `excel-ai check-config` or `:check-env`) to validate and display the current configuration state.

## 3. Error Handling & Resilience

* **Unified Tool Error Handling:** Wrap all tool executions (functions decorated with `@function_tool`) consistently. Catch exceptions *within* the tool function itself, not just relying on agent-level retries. Ensure all outcomes (success or failure) are reported back to the agent via a standardized `ToolResult` dictionary format (e.g., `{"success": bool, "result": ..., "error": ...}`).
* **Strict `ToolResult` Contract:** Implement checks (perhaps via a decorator) to ensure tool functions strictly adhere to the `ToolResult` format. Raise a specific `ToolContractViolation` or similar error if the format is invalid to prevent silent failures or data loss.
* **Specific Exception Handling:**
    * Replace broad `except Exception:` blocks with catches for more specific errors where possible (`FileNotFoundError`, `ValueError`, custom exceptions).
    * Wrap fragile `xlwings` COM/AppleScript calls. Catch specific `pywintypes.com_error` (Windows) or RPC/AppleScript exceptions. Re-raise unknown errors instead of swallowing them into a generic `ExcelConnectionError`.
    * Define custom exception classes for common application errors (e.g., `ExcelConnectionError`, `InvalidToolUsageError`, `ConfigurationError`).
* **`xlwings` Robustness:**
    * Implement retry logic with exponential backoff specifically for known intermittent COM errors (like `RPC_E_SERVERCALL_RETRYLATER`).
    * Address state consistency: If an Excel operation fails mid-way, ensure the `AppContext` state accurately reflects the actual workbook state, or explicitly inform the agent about partial failures. Consider updating context *after* operations succeed.
* **CLI Robustness:** Handle errors gracefully during CLI startup (e.g., agent creation failure, config issues) and within command execution. Provide clear error messages to the user instead of just crashing or exiting.
* **Agent Error Limits:** Refine the `MAX_CONSECUTIVE_ERRORS` logic in hooks. Use error codes or categories instead of full error messages as keys to prevent minor variations from resetting the counter. Consider limiting error message size stored in context.
* **Circuit Breakers:** Wrap calls to external dependencies (like LLM pricing endpoints) with circuit breakers to prevent repeated failures from overwhelming the system and provide fallbacks (e.g., cached pricing data).

## 4. Logging & Monitoring

* **Application-Level Configuration:** Configure logging handlers, formatters, and levels *once* at the application entry point (e.g., in `cli.py` or a dedicated `logging_config.py`), not within library modules.
* **Library Logging:** Library modules should only get a logger instance (`logger = logging.getLogger(__name__)`) and use it.
* **Configurable Logging:** Allow logging level and format (e.g., plain text vs. JSON) to be configured via CLI arguments or the settings file.
* **Structured Logging:** Use a library like `structlog` for JSON output (e.g., enabled by `EXCEL_AI_JSON_LOG=1`) to facilitate automated log processing.
* **Contextual Information:** Include relevant context (e.g., active sheet, tool name, relevant state variables) in log messages, especially for errors.
* **Metrics Hook:** Implement a mechanism (e.g., hooks within the agent runner) to collect and export operational metrics (like token usage, tool calls, errors, costs) to monitoring systems (e.g., Prometheus via an exporter, OpenTelemetry). Flush cost metrics instead of just storing them in volatile state.

## 5. Testing

* **Mandatory Automated Tests:** Implement a comprehensive test suite using `pytest`. This is non-negotiable for production readiness.
* **Unit Tests:** Test individual functions, helper utilities, configuration loading, and context management. Mock external dependencies like LLM APIs and `xlwings`.
* **Integration Tests:**
    * Test tool functions against a mocked `ExcelGateway` (see Architecture) or a real (headless) Excel instance if feasible.
    * Test the agent's core flow for completing simple tasks end-to-end, possibly mocking LLM responses based on expected tool call sequences.
    * Test the CLI argument parsing and command execution flow.
* **CI Pipeline:** Set up a Continuous Integration (CI) pipeline (e.g., GitHub Actions, GitLab CI) to run on every commit/PR.
* **CI Quality Gates:** Gate merges on passing tests (`pytest`), linting (`ruff --select I,S,E9,F`), static type checking (`mypy --strict`), and potentially dependency consistency checks (`pip-tools compile --dry-run`).
* **CI Matrix (Optional):** Consider running tests on different platforms (Windows, macOS, Linux) if cross-platform support is critical, using stubbed implementations where needed (e.g., `ExcelGateway` stub on Linux).

## 6. Code Structure & Architecture

* **Port/Adapter Pattern (`ExcelGateway`):** Introduce an abstract base class (`ExcelGateway`) defining the interface for Excel operations (e.g., `open`, `get_cell`, `set_cell`, `insert_table`).
    * Implement a concrete adapter using `xlwings` for production.
    * Implement a mock/in-memory adapter for testing.
    * Refactor all tool functions and `AppContext` to depend on the `ExcelGateway` interface, not directly on `xlwings`. This decouples logic and enables testing without Excel.
* **Modularization:** Break down large files (`excel_ops.py`, `cli.py`) into smaller, more focused modules with clear responsibilities.
    * Example split for `excel_ops`: `manager.py` (core connection/lifecycle), `tables.py`, `styling.py`, `reading.py`.
    * Example split for `cli`: `main.py` (entry point, setup), `commands.py` (command logic), `ui.py` (spinner, display logic).
* **Async Discipline:**
    * Ensure all tool functions are `async def` for uniformity, even if they don't perform async I/O internally.
    * Offload blocking I/O calls (like `xlwings` COM/AppleScript interactions within the gateway) to a thread pool using `asyncio.to_thread` (or `trio.to_thread.run_sync`) to avoid blocking the event loop.
    * Remove discouraged synchronous wrappers like `ExcelManager.close_sync`.
* **Pure Hooks:** Refactor agent hooks (like `SummaryHooks`) to be pure functions. Instead of mutating context directly, they should return a dictionary representing the *changes* (a patch) to be applied by the runner.
* **Docstrings:** Ensure comprehensive docstrings for public functions/methods, detailing parameters, return values, and potential exceptions raised.
* **Constants:** Consolidate constants into fewer, well-organized files (e.g., merge `debounce_constants.py` into `constants.py` if appropriate).
* **Tool Decoration:** Ensure all functions intended as agent tools are correctly decorated (`@function_tool`). Consider adding an assertion or check during agent setup to verify this.

## 7. Performance & Resource Management

* **Excel Instance Management:**
    * Prefer creating and managing dedicated Excel instances (`xw.App(add_book=False)`) over attaching to existing ones (`xw.App(visible=True)` attach can be unpredictable).
    * Implement robust cleanup logic to ensure the Excel instance is closed cleanly on exit or error (`app.quit()`). Handle potential errors during quit.
    * Ensure reliable cleanup of temporary files (e.g., snapshots created with `tempfile`).
* **Shape Scanning Optimization:**
    * Make shape scanning frequency configurable (e.g., `SHAPE_SCAN_EVERY_N_WRITES`) or allow disabling it (e.g., via a `--fast` flag or config setting).
    * Implement more aggressive compaction/summarization for workbook shape representation sent to the LLM (`_format_workbook_shape`), especially for large sheets/headers.
    * Consider caching read results for short periods.
    * Potentially implement smarter change detection to avoid full rescans if changes are localized or small.
* **Prompt Token Usage:**
    * Review and condense the system prompt to reduce fixed token cost and latency. Ensure all instructions are necessary and effective.
    * Implement strict length caps on dynamically generated parts of the prompt (workbook shape, named ranges) to prevent exceeding context limits with large files.

## 8. Security

* **Input Sanitization/Guardrails:** Implement actual guardrails (as hinted by `input_guardrail`/`output_guardrail` mentions) to check agent instructions or user input for potentially malicious operations (e.g., attempts to access local file system outside designated areas, read environment variables, execute arbitrary code).
* **Restrict Tool Access (Optional):** If deployed in multi-user or sensitive environments, evaluate if Role-Based Access Control (RBAC) or privilege levels are needed to restrict access to certain powerful tools.
* **Secure Log Storage:** Ensure logs do not inadvertently store sensitive data (credentials, PII, full file contents unless explicitly required and secured).

## 9. Documentation

* **API Documentation:** Generate API documentation automatically from docstrings using tools like `mkdocs` with `mkdocstrings` and the `mkdocs-material` theme. Publish this documentation (e.g., to GitHub Pages).
* **User Documentation:** Provide clear instructions on installation, configuration, usage, and troubleshooting.
* **Tool Schema (Optional):** Consider defining and potentially publishing a schema (e.g., JSON Schema or OpenAPI-like) for each tool's parameters, which could be useful for integrations or future UI development.

## Sequence of Work (Suggested High-Level Order)

1.  **Foundation:** Implement `pyproject.toml`, lock dependencies, set up basic CI structure.
2.  **Configuration:** Refactor to use a centralized `Settings` object (`pydantic-settings`), validate early.
3.  **Architecture & Testability:** Introduce `ExcelGateway` abstraction, refactor tools/context, set up `pytest` with mock gateway tests.
4.  **Error Handling:** Implement unified `ToolResult` handling, custom exceptions, stricter `xlwings` error wrapping.
5.  **Async/Structure:** Ensure async uniformity, offload blocking calls, refactor large modules.
6.  **Testing Expansion:** Add integration tests, CI quality gates (linting, typing).
7.  **Logging & Metrics:** Centralize config, add structured logging, implement metrics hook.
8.  **Security & Performance:** Implement guardrails, optimize shape scanning, refine resource management.
9.  **CLI & UX:** Refactor CLI, improve user feedback.
10. **Documentation:** Generate API docs, improve user guides.

By systematically addressing these areas, the Excel AI Agent can transition from a functional prototype to a robust, maintainable, and reliable production-grade tool.