# Autonomous Excel Assistant (Python Edition)

This project is an autonomous agent system for manipulating Excel files using natural language, powered by OpenAI Agents SDK and openpyxl.

## Features
- Python-based agent architecture
- Excel file manipulation (read, write, style, formulas, etc.)
- Verification tool: `get_range_values_tool` to fetch and inspect cell ranges post-operations
- CLI interface

## Stack
- Python 3.9+
- openpyxl
- openai-agents (see `agents-sdk-docs/`)
- python-dotenv

## Setup
1. Create a virtual environment: `python -m venv .venv`
2. Install dependencies: `pip install -r requirements.txt`
3. Set your OpenAI API key in `.env`

## Usage
Run the CLI:
```sh
python src/cli.py --input-file input.xlsx --output-file output.xlsx --instruction "<your instruction>"
```

- Note: Tool wrappers rely on ExcelManager methods, which return None on success; always catch exceptions and return True on success in tool implementations.
