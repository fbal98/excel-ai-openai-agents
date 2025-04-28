# Autonomous Excel Assistant

A command-line tool that uses an AI agent to perform spreadsheet operations on Excel workbooks autonomously.

## Features
- Natural-language instructions to manipulate spreadsheets via the OpenAI API.
- Supports batch mode (via `openpyxl`) and live mode (via `xlwings`).
- Streaming mode to see agent progress in real-time.
- Verbose mode for debug logging and full result inspection.
- Emoji-enhanced, structured logging for clear status updates.
- Bulk operations for tables, formulas, and cell styling.
- **Dynamic Context**: Workbook structure (`<workbook_shape>`) and progress updates (`<progress_summary>`) are injected directly into the conversation history as assistant messages, keeping the LLM informed without cluttering the main system prompt.

## Requirements
- Python 3.9 or higher
- An OpenAI API key set in the environment (`OPENAI_API_KEY`) or in a `.env` file
- Dependencies (see `requirements.txt`):
  - `openai-agents`, `python-dotenv`, `openpyxl`, `xlwings`

## Installation
```bash
pip install -r requirements.txt
```

## Usage
```bash
python -m src.cli --instruction "<your instruction>" [options]
```

### Options
- `--instruction` (required): Natural-language command for the AI agent.
- `--input-file` (optional): Path to an existing Excel workbook. If omitted, a new workbook is created.
- `--output-file` (required in batch mode): Path to save the modified workbook (ignored in live mode).
- `--live`: Enable live editing in Excel via `xlwings`. Changes appear in real time.
- `--stream`: Enable streaming to see the agent's progress as it works.
- `-v`, `--verbose`: Enable verbose logging (debug-level messages and full result dump).

### Examples
**Batch mode**
```bash
export OPENAI_API_KEY=...
python -m src.cli \
  --instruction "Add a summary sheet with totals for all numeric columns." \
  --input-file data.xlsx \
  --output-file output.xlsx
```

**Live mode**
```bash
export OPENAI_API_KEY=...
python -m src.cli \
  --instruction "Highlight all negative values in red." \
  --input-file data.xlsx \
  --live
```

**Streaming mode**
```bash
export OPENAI_API_KEY=...
python -m src.cli \
  --instruction "Create a financial dashboard with charts." \
  --input-file financials.xlsx \
  --output-file dashboard.xlsx \
  --stream
```

**Verbose output**
```bash
python -m src.cli -v \
  --instruction "Sort sheet 'Sales' by column 'Revenue' descending." \
  --input-file sales.xlsx \
  --output-file sorted.xlsx
```

## Logging
Uses Python's `logging` module with emoji prefixes:
- 📂 Loaded or 🆕 Created workbook
- 📊 Live mode enabled
- 🔄 Streaming mode enabled
- 💡 Instruction detail
- 🤖 Running agent
- 🛠️ Tool execution
- ✅ Agent completion time
- 📤 Final AI-generated output
- 📁 Workbook saved
- ❌ Errors

## Operation Examples
- Create sheets, tables, and data structures
- Format cells with colors, fonts, borders
- Merge/unmerge cells and adjust row heights and column widths
- Add formulas and calculations
- Find, filter, and process data
- Verify operations for accuracy

## Limitations
- Live mode requires `xlwings` and a running Excel instance.
- Some style and inspection operations may not be supported in live mode.
- Maximum 25 agent turns per instruction for complexity control.