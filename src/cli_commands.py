"""
Central registry for CLI commands.

Designed for future refactor: each command is a `Command` dataclass instance
holding its name, description, and async handler.

Currently **unused by cli.py** – but available for upcoming migration.
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Awaitable, List, Dict, Optional

@dataclass(frozen=True, slots=True)
class Command:
    name: str
    description: str
    handler: Callable[[List[str]], Awaitable[None]]

# Global registry -------------------------------------------------------------
COMMAND_REGISTRY: Dict[str, Command] = {}

def register(command: Command) -> None:
    """Register a new command. Raises on duplicates."""
    if command.name in COMMAND_REGISTRY:
        raise ValueError(f"Duplicate CLI command registered: {command.name}")
    COMMAND_REGISTRY[command.name] = command

def get(name: str) -> Optional[Command]:
    """Fetch a command by name or return None."""
    return COMMAND_REGISTRY.get(name)