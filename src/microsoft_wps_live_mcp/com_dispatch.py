"""Use late-bound wrappers throughout the object tree, without changing global makepy caches."""

import functools
import time
from types import MethodType


def retry_busy(operation):
    import pythoncom
    import pywintypes
    deadline = time.monotonic() + 8
    while True:
        try:
            return operation()
        except pywintypes.com_error as exc:
            # These two RPC errors mean the application rejected the call before
            # execution. Other errors/timeouts must never replay a modification.
            if exc.hresult not in (-2147418111, -2147417846) or time.monotonic() >= deadline:
                raise
            pythoncom.PumpWaitingMessages()
            time.sleep(0.1)


def dispatch(obj, userName=None, returnCLSID=None, typeinfo=None, clsctx=None):
    import pythoncom
    import win32com.client.dynamic as dynamic

    class LiveDispatch(dynamic.CDispatch):
        def __getattr__(self, name):
            return retry_busy(lambda: super(LiveDispatch, self).__getattr__(name))

        def __setattr__(self, name, value):
            return retry_busy(lambda: super(LiveDispatch, self).__setattr__(name, value))

        def _wrap_dispatch_(self, ob, userName=None, returnCLSID=None):
            return dispatch(ob, userName, returnCLSID)

        def _make_method_(self, name):
            method = super()._make_method_(name)
            if method is not None:
                # pywin32's dynamic method builder otherwise returns makepy wrappers.
                # Its generated function has a private globals dictionary, so this does
                # not monkeypatch win32com or other applications' Python processes.
                method.__func__.__globals__["Dispatch"] = dispatch
                original = method.__func__

                @functools.wraps(original)
                def guarded(instance, *args, **kwargs):
                    return retry_busy(lambda: original(instance, *args, **kwargs))

                self._builtMethods_[method.__name__] = guarded
                return MethodType(guarded, self)
            return method

    return dynamic.Dispatch(obj, userName=userName, createClass=LiveDispatch,
                            typeinfo=typeinfo, clsctx=clsctx or pythoncom.CLSCTX_SERVER)

