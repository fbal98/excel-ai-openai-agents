"""
Main entry point when running the package with `python -m src`
"""

import asyncio
import sys
from .cli import main

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nExiting.")
        sys.exit(0) 