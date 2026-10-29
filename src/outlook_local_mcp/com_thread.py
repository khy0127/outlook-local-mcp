"""
com_thread.py

Runs every Outlook COM call on one dedicated STA thread.

The MCP SDK executes synchronous tools on a pool of worker threads, but COM
proxies to OUTLOOK.EXE belong to the thread that created them and every
thread must call CoInitialize first. Funnelling all calls through a single
worker keeps that simple and also serialises access, so two tool calls can
never interleave operations on the same mailbox.
"""

from __future__ import annotations

import concurrent.futures
from typing import Any, Callable, TypeVar

T = TypeVar("T")

# Outlook can block behind a modal dialog; don't let a tool call hang forever.
_CALL_TIMEOUT_SECONDS = 600


def _init_com() -> None:
    import pythoncom

    pythoncom.CoInitialize()


_executor = concurrent.futures.ThreadPoolExecutor(
    max_workers=1, thread_name_prefix="outlook-com", initializer=_init_com
)


def run(fn: Callable[..., T], *args: Any, **kwargs: Any) -> T:
    """Execute fn(*args, **kwargs) on the COM thread and return its result.

    Failures are re-raised as ToolError so the model sees the reason (e.g. a
    refused move) instead of the SDK's generic "Error executing tool".
    """
    from mcp.server.mcpserver.exceptions import ToolError

    future = _executor.submit(fn, *args, **kwargs)
    try:
        return future.result(timeout=_CALL_TIMEOUT_SECONDS)
    except concurrent.futures.TimeoutError as exc:
        raise ToolError(
            f"Outlook did not respond within {_CALL_TIMEOUT_SECONDS}s. "
            "Check whether Outlook is showing a dialog box that needs attention."
        ) from exc
    except Exception as exc:  # noqa: BLE001
        raise ToolError(str(exc)) from exc
