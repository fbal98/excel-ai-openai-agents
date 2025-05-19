# AGENTS.md

This repository contains an AI assistant for Excel powered by the OpenAI Agents SDK.
These notes guide automated contributions.

## Style
- Follow the rules in `CLAUDE.md` (PEP8-like style, max 100 characters per line,
  grouped imports, Google-style docstrings).
- All `.py`, `.md` and `.txt` files must end with a single trailing newline.

## Development
- Keep modules focused and avoid mixing configuration with logic.
- Functions should perform a single task and stay concise. Helper functions are
  preferred over long monoliths.
- New tools placed under `src/tools/` must end with `_tool` and return a
  `ToolResult` typed dictionary.

## Tests
Run Python compilation checks before committing any change:

```bash
python -m py_compile $(git ls-files '*.py' | grep -v hello_world_jupyter.py)
```

Fix any errors before creating a pull request.
