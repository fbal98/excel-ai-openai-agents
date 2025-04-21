"""Central helper for ensuring every callable is an `agents` FunctionTool.

Usage
-----
>>> from .tool_wrapper import ensure_tool, with_retry
>>> wrapped = ensure_tool(my_tool_function)
"""

import inspect
from functools import wraps
from agents import function_tool, FunctionTool


def with_retry(max_retries: int = 1):
    """
    Decorator that retries a tool *max_retries* times (default **1**) whenever
    the wrapped function returns a mapping that looks like
    ``{"error": ...}``.

    The wrapper **preserves the original function's metadata** (``__name__``,
    ``__doc__``, and signature) via :pyfunc:`functools.wraps`, which is
    required so that OpenAI's function‑tool mapping in the *agents* SDK
    continues to work.

    The decorator automatically produces an *async* or *sync* wrapper to match
    the wrapped function’s nature.

    Parameters
    ----------
    max_retries:
        Number of *additional* attempts after the initial call. A value of 1
        therefore means *two total* executions (initial + one retry).

    Notes
    -----
    • Works for both **async** and **sync** callables.
    • Assumes the wrapped tool is idempotent.
    """

    def _decorator(fn):
        # Decide at decoration time which kind of wrapper we need.
        if inspect.iscoroutinefunction(fn):

            @wraps(fn)
            async def _wrapper(ctx, *a, **kw):   # type: ignore[override]
                attempts = 0
                while True:
                    result = await fn(ctx, *a, **kw)
                    # Treat anything without {"error": ...} as success.
                    if not (isinstance(result, dict) and result.get("error")):
                        return result
                    if attempts >= max_retries:
                        return result
                    attempts += 1

            _wrapper.__signature__ = inspect.signature(fn)
            _wrapper.__signature__ = inspect.signature(fn)
            return _wrapper

        else:

            @wraps(fn)
            def _wrapper(ctx, *a, **kw):         # type: ignore[override]
                attempts = 0
                while True:
                    result = fn(ctx, *a, **kw)
                    if not (isinstance(result, dict) and result.get("error")):
                        return result
                    if attempts >= max_retries:
                        return result
                    attempts += 1

            return _wrapper

    return _decorator


# The ensure_tool function is removed as explicit decoration is now handled in tools.py
# using @function_tool.