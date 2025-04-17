# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands
- Run the application: `python -m src.cli --instruction "<your instruction>" [options]`
- Run tests: `python test_agent_run.py`
- Install dependencies: `pip install -r requirements.txt`

## Code Style Guidelines
- **Imports**: Group by stdlib, third-party, then internal modules
- **Naming**: snake_case for functions/variables, CamelCase for classes, UPPER_CASE for constants
- **Tool functions**: Always suffix with `_tool` and use function_tool decorator
- **Type Annotations**: Required for all parameters and return values
- **Docstrings**: Required for functions, classes, and modules
- **Error Handling**: Use structured try/except blocks with descriptive messages
- **Line Length**: Maximum 100 characters
- **Excel Operations**: Prefer bulk operations over single-cell operations when possible
- **Logging**: Use emoji prefixes (📂, 🤖, ✅, etc.) for structured logging