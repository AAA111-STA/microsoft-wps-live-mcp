"""Keep every COM reference on one daemon STA thread, including across MCP calls."""
from __future__ import annotations

import asyncio
from collections import OrderedDict
from concurrent.futures import Future
import queue
import sys
import threading
import uuid

from .native import NativeOffice, OfficeError


class OfficeWorker:
    def __init__(self, data_dir=None, timeout=90, factory=NativeOffice, autosave_interval=1.0):
        self.data_dir = data_dir
        self.timeout = timeout
        self.factory = factory
        self.autosave_interval = autosave_interval
        self.jobs = queue.Queue()
        self.history = OrderedDict()
        self.progress = {}
        self.lock = threading.Lock()
        self.current = None
        self.thread = None

    def _run(self):
        if sys.platform == "win32":
            import pythoncom
            pythoncom.CoInitialize()
        engine = None
        try:
            while True:
                try:
                    job = self.jobs.get(timeout=self.autosave_interval)
                except queue.Empty:
                    if engine is not None and hasattr(engine, "autosave_tick"):
                        engine.autosave_tick()
                    if sys.platform == "win32":
                        pythoncom.PumpWaitingMessages()
                    continue
                if job is None:
                    if engine is not None and hasattr(engine, "autosave_tick"):
                        engine.autosave_tick()
                    return
                identifier, future, method, args, kwargs = job
                try:
                    if engine is None:
                        engine = self.factory(self.data_dir)
                    result = getattr(engine, method)(*args, **kwargs)
                    if hasattr(engine, "enforce_autosave"):
                        result = engine.enforce_autosave(result)
                except Exception as exc:
                    if engine is not None and hasattr(engine, "pause_after_error"):
                        engine.pause_after_error(args)
                    future.set_exception(exc)
                else:
                    future.set_result(result)
        finally:
            # Never close or quit a user's applications on MCP disconnect.
            engine = None
            if sys.platform == "win32":
                pythoncom.CoUninitialize()

    def submit(self, method, *args, **kwargs):
        with self.lock:
            if self.current and not self.history[self.current].done():
                raise OfficeError(f"Office operation {self.current} is still running. "
                                  "Use office_operation_status; do not replay a modification.")
            if self.thread is None:
                self.thread = threading.Thread(target=self._run, name="office-mcp-sta", daemon=True)
                self.thread.start()
            identifier = uuid.uuid4().hex
            future = Future()
            self.current = identifier
            self.history[identifier] = future
            if "_on_progress" in kwargs:
                callback = kwargs["_on_progress"]

                def report(event):
                    with self.lock:
                        self.progress[identifier] = dict(event)
                    if callback:
                        callback(event)

                kwargs["_on_progress"] = report
            while len(self.history) > 100:
                expired, _ = self.history.popitem(last=False)
                self.progress.pop(expired, None)
            self.jobs.put((identifier, future, method, args, kwargs))
            return identifier, future

    async def call(self, method, *args, **kwargs):
        identifier, future = self.submit(method, *args, **kwargs)
        try:
            result = await asyncio.wait_for(asyncio.shield(asyncio.wrap_future(future)), self.timeout)
            return {**result, "operation_id": identifier}
        except asyncio.TimeoutError:
            return {"state": "pending", "operation_id": identifier, "confirmed": False,
                    "message": "Native application is still processing. Query office_operation_status; "
                               "do not replay inserts or other modifications."}
        except OfficeError:
            raise
        except Exception as exc:
            raise OfficeError(f"Native Office operation failed ({identifier}): {exc}. "
                              "A multi-step edit may have partially applied. Inspect the live document "
                              "before retrying; no success or disk save is implied.") from exc

    def operation_status(self, identifier):
        with self.lock:
            future = self.history.get(identifier)
        if future is None:
            raise OfficeError("Unknown operation ID or result expired. Inspect the document before retrying.")
        if not future.done():
            return {"operation_id": identifier, "state": "pending", "confirmed": False,
                    "progress": self.progress.get(identifier)}
        try:
            result = future.result()
        except Exception as exc:
            return {"operation_id": identifier, "state": "failed", "confirmed": False,
                    "message": str(exc), "inspect_before_retry": True, "progress": self.progress.get(identifier)}
        return {"operation_id": identifier, "state": "completed", "confirmed": True, "result": result}

    def stop(self):
        self.jobs.put(None)
        if self.thread is not None and self.thread is not threading.current_thread():
            self.thread.join(timeout=3)

