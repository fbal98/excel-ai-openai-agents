# Plan: Enhance CLI Feedback

**Goal:** Improve the user experience of the Excel AI Assistant CLI by adding clearer visual indicators for loading/thinking states, success/error feedback, and progress during tool execution.

**Affected Files:**

*   `src/cli.py`: Handles the main CLI loop, initial spinner, and overall error catching.
*   `src/stream_renderer.py`: Formats and renders individual events from the agent stream, including tool calls and agent messages.

**Proposed Changes:**

1.  **Enhance Loading/Thinking Indicator (`src/cli.py`)**
    *   Replace the current `_spinner` function with a more robust solution (e.g., using `rich` or `halo`).
    *   Ensure the spinner starts reliably when `run_agent_streamed` is called and stops cleanly before the first meaningful output is printed.

2.  **Implement Success/Error Feedback**
    *   **Overall Success (`src/stream_renderer.py`):** Prepend "✔️ " to the *final* agent message. Requires reliably identifying the final message chunk.
    *   **Tool Success/Error (`src/stream_renderer.py`):** Keep the existing icons ("🛠️ ✔/✗"). Consider adding color (e.g., green for ✔, red for ✗) via ANSI codes or a library like `rich`.
    *   **General Errors (`src/cli.py`):** Keep the "❌" icon. Ensure consistency with tool error formatting (e.g., using color if adopted).

3.  **Implement Progress Bars for Tool Calls (`src/stream_renderer.py`)**
    *   When a `tool_start` event is processed:
        *   Print the standard "🛠️ Tool: name(...)" line.
        *   Immediately *start* an *indeterminate* progress indicator/spinner on the *next line* (e.g., using `rich` or `halo`).
    *   When the corresponding `tool_end` event is processed:
        *   *Stop* and *remove* the progress indicator from the screen.
        *   Print the standard "🛠️ Tool ✔/✗ ..." result line.
    *   **Consideration:** Requires managing the state of the active progress indicator between `tool_start` and `tool_end` events, especially if concurrent tool calls are possible.

**Optional Enhancement:**

*   **Adopt `rich` library:** Using `rich` throughout (`rich.print`, `rich.spinner`, `rich.progress`, console markup for colors) could unify styling, simplify implementation, and provide a more polished look.