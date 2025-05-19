# Autonomous Excel Assistant

Autonomous Excel Assistant leverages the OpenAI Agents SDK to modify Excel spreadsheets from natural-language instructions. It can run against a saved workbook in batch mode or connect to a live Excel session for real-time changes.

## Quick Start
1. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```
2. **Set your API key**
   ```bash
   export OPENAI_API_KEY=your-key
   ```
3. **Run an instruction**
   ```bash
   python -m src.cli --instruction "Add totals for each column" --input-file data.xlsx --output-file out.xlsx
   ```

The agent will open the workbook, apply your instruction, and save the result. Use `--live` to connect directly to Excel for instant updates.

## Features
- Natural-language commands executed via the OpenAI API
- Batch mode using `openpyxl` and live mode using `xlwings`
- Streaming output to watch agent progress
- Verbose logging for debugging
- Emoji-enhanced status messages
- Bulk tools for tables, formulas and cell styling
- **Dynamic Context**: workbook shape and progress summaries automatically shared with the LLM

## Usage
```
python -m src.cli --instruction "<your instruction>" [options]
```

### Options
- `--instruction` (required): Command for the AI agent
- `--input-file`: Existing workbook path (a new one is created if omitted)
- `--output-file`: Destination workbook in batch mode
- `--live`: Apply changes in real time using Excel
- `--stream`: Show the agent's progress live
- `-v`, `--verbose`: Print detailed logs

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
Status messages use emoji to make progress clear:
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
- Merge or unmerge cells and adjust row heights and column widths
- Add formulas and calculations
- Find, filter, and process data
- Verify operations for accuracy

## Limitations
- Live mode requires `xlwings` and a running Excel instance
- Some style and inspection operations may not work in live mode
- Maximum 25 agent turns per instruction keep complexity manageable

## Learn More
See the documentation in `agents-sdk-docs/` for advanced usage, tool development, and design details.

## Contributing
Issues and pull requests are welcome! If you encounter a problem or have an improvement in mind, open an issue to discuss it. New tools, bug fixes, and documentation updates all help the community.

