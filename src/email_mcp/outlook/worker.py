"""OutlookWorker — a dedicated STA thread that owns the Outlook handle.

Why this exists: COM objects are apartment-bound. A long-running MCP server
dispatches tool handlers across an asyncio event loop / threadpool, but the
Outlook handle must only be touched from the single thread (apartment) that
created it. So the server creates ONE worker thread, initializes it as an STA,
opens an :class:`OutlookSession` there, and marshals every tool call onto it.

The CLI does NOT use this — it is one-shot per process and runs the session on
its own main thread directly.

Usage (from the async server):

    worker = OutlookWorker()
    worker.start()                     # blocks until the session is attached (or fails)
    result = await asyncio.wrap_future(worker.submit(lambda s: ops(s, ...)))
    worker.shutdown()
"""

from __future__ import annotations

import queue
import threading
from concurrent.futures import Future
from typing import Any, Callable

from email_mcp.outlook.session import OutlookSession

_Job = tuple[Callable[[OutlookSession], Any], "Future[Any]"]
_SHUTDOWN: Any = object()


class OutlookWorker:
    def __init__(self) -> None:
        self._queue: "queue.Queue[_Job | Any]" = queue.Queue()
        self._thread: threading.Thread | None = None
        self._ready = threading.Event()
        self._startup_error: BaseException | None = None
        self._session: OutlookSession | None = None

    def start(self, timeout: float = 30.0) -> None:
        """Launch the worker thread and block until its session attaches (or errors)."""
        if self._thread is not None:
            return
        self._thread = threading.Thread(target=self._run, name="outlook-sta", daemon=True)
        self._thread.start()
        if not self._ready.wait(timeout):
            raise TimeoutError("Outlook worker did not become ready in time")
        if self._startup_error is not None:
            raise self._startup_error

    def _run(self) -> None:
        import pythoncom

        # Explicit STA: Outlook automation expects a single-threaded apartment.
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        try:
            self._session = OutlookSession()
            # Attach eagerly so start() surfaces COM problems synchronously.
            self._session.app  # noqa: B018 - triggers lazy attach
        except BaseException as e:  # pragma: no cover - depends on live Outlook
            self._startup_error = e
            self._ready.set()
            pythoncom.CoUninitialize()
            return
        self._ready.set()

        try:
            while True:
                job = self._queue.get()
                if job is _SHUTDOWN:
                    break
                fn, fut = job  # type: ignore[misc]
                if fut.set_running_or_notify_cancel():
                    try:
                        fut.set_result(fn(self._session))
                    except BaseException as e:  # noqa: BLE001 - marshal all errors back
                        fut.set_exception(e)
        finally:
            self._session = None
            pythoncom.CoUninitialize()

    def submit(self, fn: Callable[[OutlookSession], Any]) -> "Future[Any]":
        """Queue ``fn(session)`` to run on the worker thread; returns a Future."""
        if self._thread is None:
            raise RuntimeError("OutlookWorker.start() was not called")
        fut: "Future[Any]" = Future()
        self._queue.put((fn, fut))
        return fut

    def shutdown(self, join_timeout: float = 5.0) -> None:
        if self._thread is None:
            return
        self._queue.put(_SHUTDOWN)
        self._thread.join(join_timeout)
        self._thread = None
