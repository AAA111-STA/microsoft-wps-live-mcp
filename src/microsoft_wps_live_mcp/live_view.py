"""Visible, paced editing in the native document window, shared by all Office tasks."""
from __future__ import annotations

import datetime as dt
import sys
import time


def live_error(message):
    from .native import OfficeError
    return OfficeError(message)


class LiveView:
    def _paint(self, record, pause_ms=0):
        if getattr(self, "com", None) is not None:
            self.com.PumpWaitingMessages()
        try:
            if record.kind == "word" and record.engine == "microsoft":
                record.app.ScreenRefresh()
        except Exception:
            pass
        # Native Office has its own UI thread. Yield time for it to paint while
        # this STA also pumps incoming COM messages; never suppress screen updates.
        deadline = time.monotonic() + pause_ms / 1000
        while time.monotonic() < deadline:
            if getattr(self, "com", None) is not None:
                self.com.PumpWaitingMessages()
            time.sleep(min(0.04, max(0, deadline - time.monotonic())))

    @staticmethod
    def _window_handle(window, record=None):
        if sys.platform != "win32":
            return None
        import win32gui
        for name in ("Hwnd", "HWND"):
            try:
                value = getattr(window, name)
                hwnd = int(value() if callable(value) else value)
                if win32gui.IsWindow(hwnd):
                    return win32gui.GetAncestor(hwnd, 2) or hwnd
            except Exception:
                pass
        if record is not None and getattr(record, "kind", None) == "ppt":
            h = win32gui.FindWindow("PPTFrameClass", None)
            if h and win32gui.IsWindow(h):
                return h
        if record is not None:
            # Some PowerPoint/WPS type libraries expose HWND but reject its
            # invocation. A unique matching top-level document caption is a
            # fallback; don't pick between ambiguous windows with the same name.
            from pathlib import Path
            import os
            import win32api
            import win32process
            name = str(record.doc.Name).casefold()
            stem = Path(name).stem
            application_dir = os.path.normcase(str(record.app.Path))
            matches = []

            def visit(hwnd, _):
                caption = win32gui.GetWindowText(hwnd).casefold()
                if win32gui.IsWindowVisible(hwnd) and (name in caption or stem in caption):
                    process = None
                    try:
                        pid = win32process.GetWindowThreadProcessId(hwnd)[1]
                        process = win32api.OpenProcess(0x410, False, pid)
                        image = win32process.GetModuleFileNameEx(process, 0)
                        if os.path.normcase(str(Path(image).parent)) == application_dir:
                            matches.append(hwnd)
                    except Exception:
                        pass
                    finally:
                        if process is not None:
                            process.Close()
                return True

            win32gui.EnumWindows(visit, None)
            if len(matches) == 1:
                return matches[0]
            if len(matches) > 1:
                # Office can show two files with the same basename. Activation
                # plus the native window's Active state disambiguates the actual
                # foreground window; don't select an arbitrary caption match.
                foreground = win32gui.GetForegroundWindow()
                if foreground in matches:
                    try:
                        if bool(window.Active):
                            return foreground
                    except Exception:
                        pass
        return None

    def _ensure_live_window(self, record, foreground=False):
        if not record.live_enabled:
            return None
        app, doc = record.app, record.doc
        try:
            if not bool(app.Visible):
                app.Visible = True
            try:
                if not bool(app.ScreenUpdating):
                    app.ScreenUpdating = True
            except Exception:
                pass  # PowerPoint has no Application.ScreenUpdating property.
            if int(doc.Windows.Count) == 0:
                if record.kind == "ppt":
                    doc.NewWindow()
                else:
                    raise live_error("The selected document has no native editing window.")
            window = doc.Windows.Item(1)
            try:
                if not bool(window.Visible):
                    window.Visible = True
            except Exception:
                pass
            activate = foreground or getattr(self, "live_active_document", None) != record.id
            try:
                active = getattr(app, {"word": "ActiveDocument", "excel": "ActiveWorkbook", "ppt": "ActivePresentation"}[record.kind])
                activate = foreground or str(active.FullName) != str(doc.FullName)
            except Exception:
                pass
            if activate:
                window.Activate()
            self.live_active_document = record.id
            hwnd = self._window_handle(window, record)
            if foreground and not hwnd and record.kind == "ppt" and record.engine == "microsoft":
                try:
                    app.Activate()
                    hwnd = self._window_handle(window, record)
                except Exception:
                    pass
            visible = None
            if sys.platform == "darwin":
                try:
                    if hasattr(app, "Activate"):
                        app.Activate()
                except Exception:
                    pass
                visible = True
            elif hwnd:
                import win32con
                import win32gui
                if win32gui.IsIconic(hwnd):
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                visible = bool(win32gui.IsWindowVisible(hwnd))
                if not visible:
                    win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
                    visible = bool(win32gui.IsWindowVisible(hwnd))
                if foreground:
                    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                    if right - left < 800 or bottom - top < 500:
                        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
                    try:
                        import ctypes
                        ctypes.windll.user32.ShowWindow(hwnd, win32con.SW_SHOW)
                        ctypes.windll.user32.SwitchToThisWindow(hwnd, True)
                    except Exception:
                        pass
                    try:
                        win32gui.SetForegroundWindow(hwnd)
                    except Exception:
                        pass  # Windows may refuse foreground focus; don't simulate keys.
            if foreground and record.kind == "ppt":
                try:
                    window.View.ZoomToFit = True
                except Exception:
                    pass
            if visible is False:
                raise live_error("The native editing window could not be made visible.")
            record.live_state = {"enabled": True, "state": "ready", "window_open": True,
                                 "visibility_confirmed": visible is True, "hwnd": hwnd,
                                 "updated_at": dt.datetime.now(dt.timezone.utc).isoformat()}
            self.__dict__.setdefault("live_windows", {})[record.id] = window
            self._paint(record)
            return window
        except Exception as exc:
            record.live_state = {"enabled": True, "state": "unavailable", "error": str(exc)}
            raise live_error("Visible editing could not be prepared. No new edit was started: " + str(exc)) from exc

    def _focus_live(self, record, window, focus):
        if not focus:
            return
        if record.kind == "word":
            start = focus.get("start", focus.get("position"))
            if start is None:
                return
            finish = focus.get("end", start)
            size = int(record.doc.Content.End)
            if not isinstance(start, int) or not isinstance(finish, int) or not 0 <= start <= finish <= size:
                raise live_error("Word live focus requires valid start/end native character offsets.")
            span = record.doc.Range(start, finish)
            if record.engine == "wps":
                # Avoid selecting WPS end-of-cell/end-of-document marker ranges;
                # a caret at the first character navigates without a table selection.
                record.doc.Range(start, start).Select()
            else:
                window.ScrollIntoView(span, True)
            record.live_state["focus"] = {"start": start, "end": finish}
        elif record.kind == "excel":
            from .native import a1_size
            sheet = focus.get("sheet", 1)
            address = focus.get("address", focus.get("range"))
            if address is None:
                return
            a1_size(address)
            cell = self._sheet(record, sheet).Range(address)
            record.app.Goto(cell, True)
            record.live_state["focus"] = {"sheet": str(cell.Worksheet.Name), "address": address}
        else:
            slide = focus.get("slide")
            if slide is None:
                return
            self._slide(record, slide)
            try:
                if int(window.ViewType) not in (1, 9):
                    window.ViewType = 9  # Navigate to a slide only when explicitly focusing a slide.
            except Exception:
                pass
            window.View.GotoSlide(slide)
            record.live_state["focus"] = {"slide": slide}
            if focus.get("shape") is not None:
                page = self._slide(record, slide)
                if not 1 <= focus["shape"] <= int(page.Shapes.Count):
                    raise live_error("shape must refer to an existing 1-based shape index.")
                shape = page.Shapes.Item(focus["shape"])
                shape.Select()
                record.live_state["focus"]["shape"] = focus["shape"]

    def live_checkpoint(self, record, focus=None, pause_ms=None, label=None):
        if not record.live_enabled:
            return {"enabled": False}
        try:
            window = self.__dict__.get("live_windows", {}).get(record.id)
            hwnd = (record.live_state or {}).get("hwnd")
            if window is not None and hwnd and sys.platform == "win32":
                import win32gui
                if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd) or win32gui.IsIconic(hwnd):
                    window = None
            if window is None:
                window = self._ensure_live_window(record)
            record.live_state.pop("focus_error", None)
            record.live_state["updated_at"] = dt.datetime.now(dt.timezone.utc).isoformat()
            try:
                self._focus_live(record, window, focus)
            except Exception as exc:
                record.live_state["focus_error"] = str(exc)
            if label:
                record.live_state["label"] = label
            self._paint(record, record.live_pause_ms if pause_ms is None else pause_ms)
        except Exception as exc:
            # A completed/saved edit must remain distinguishable from a display error.
            record.live_state = {"enabled": True, "state": "unavailable", "error": str(exc)}
        return dict(record.live_state or {"enabled": True, "state": "unknown"})

    def show_document(self, document, focus=None):
        record = self.resolve(document)
        record.live_enabled = True
        self._ensure_live_window(record, foreground=True)
        self.live_checkpoint(record, focus, pause_ms=0)
        return self._info(record)

    def _focus_for_object(self, record, obj):
        # Best effort: navigate to a changed native object without guessing indices.
        current = obj
        for _ in range(6):
            try:
                if record.kind == "ppt":
                    index = int(current.SlideIndex)
                    return {"slide": index}
                if record.kind == "excel":
                    return {"sheet": str(current.Worksheet.Name), "address": str(current.Address)}
                return {"start": int(current.Start), "end": int(current.End)}
            except Exception:
                pass
            try:
                current = current.Parent
            except Exception:
                break
        return None

    def live_steps(self, document, stages, pause_ms=450, _on_progress=None, _cancel_event=None):
        from .object_model import validate_steps
        if not isinstance(stages, list) or not 1 <= len(stages) <= 100:
            raise live_error("stages must contain 1 to 100 named stages.")
        if isinstance(pause_ms, bool) or not isinstance(pause_ms, int) or not 0 <= pause_ms <= 2000:
            raise live_error("pause_ms must be an integer between 0 and 2000.")
        for stage in stages:
            if (not isinstance(stage, dict) or set(stage) - {"label", "steps", "focus"}
                    or not isinstance(stage.get("label"), str) or not stage["label"].strip()
                    or len(stage["label"]) > 200 or not isinstance(stage.get("steps"), list)
                    or not stage["steps"] or ("focus" in stage and not isinstance(stage["focus"], dict))):
                raise live_error("Each stage needs label and nonempty steps, with optional focus object.")
            validate_steps(stage["steps"])
        record = self.resolve(document)
        record.live_enabled = True
        self._ensure_live_window(record, foreground=True)
        aliases, completed = {}, []
        for index, stage in enumerate(stages, 1):
            if _cancel_event is not None and _cancel_event.is_set():
                return {**self._info(record), "state": "cancelled", "completed_stages": len(completed),
                        "stages": completed, "disk_saved": bool(record.doc.Saved)}
            try:
                result = self.api_batch(record.id, stage["steps"], True, _aliases=aliases)
            except Exception as exc:
                # Native busy/save/preparation failures may escape api_batch.
                # Preserve confirmed prior stages without making more COM reads
                # into a busy app or losing the history behind a generic error.
                record.autosave_paused = True
                record.autosave_state = {"state": "needs_inspection", "message": str(exc)}
                return {"document_id": record.id, "state": "partial_error", "disk_saved": False,
                        "completed_stages": len(completed), "failed_stage": index, "stages": completed,
                        "error": str(exc), "inspect_before_retry": True,
                        "message": "Prior stages remain saved. Inspect the current stage before continuing; no replay was attempted."}
            view = self.live_checkpoint(record, stage.get("focus"), pause_ms=0, label=stage["label"])
            event = {"stage": index, "total": len(stages), "label": stage["label"],
                     "state": result["state"], "disk_saved": result.get("disk_saved", False), "live_view": view}
            completed.append({**event, "results": result.get("results", [])})
            if _on_progress:
                _on_progress(event)
            if result["state"] != "completed":
                return {**result, "completed_stages": index - 1, "failed_stage": index, "stages": completed}
            if view.get("state") == "unavailable":
                return {**result, "state": "display_error", "completed_stages": index,
                        "stages": completed, "message": "Stage was saved; visible display failed. Inspect before continuing."}
            self._paint(record, pause_ms)
        return {**self._info(record), "state": "completed", "disk_saved": bool(record.doc.Saved),
                "completed_stages": len(completed), "stages": completed}
