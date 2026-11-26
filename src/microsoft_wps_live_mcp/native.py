"""Stateful native COM operations. All calls must run on one serialized STA thread."""
from __future__ import annotations

from contextlib import contextmanager
from dataclasses import dataclass
import datetime as dt
import hashlib
import json
import math
import os
from pathlib import Path
import re
import shutil
import sys
import time
import uuid
import zipfile
from types import SimpleNamespace

from mcp.server.mcpserver.exceptions import ToolError

from .discovery import (COLLECTIONS, EXTENSIONS, MACOS_APPS, PROGIDS, canonical,
                        capabilities, macos_app_installed, native_window_apps,
                        registration, vendor_of)
from .com_dispatch import dispatch, retry_busy
from .object_model import ObjectModel
from .features import SuiteFeatures
from .live_view import LiveView
from .ppt_diagram import PptDiagramMixin


class OfficeError(ToolError, RuntimeError):
    """An actionable error, safe to show to an MCP client."""


def file_hash(path: Path) -> str | None:
    if not path.is_file():
        return None
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def json_value(value):
    if isinstance(value, (tuple, list)):
        return [json_value(v) for v in value]
    if isinstance(value, (dt.datetime, dt.date)):
        return value.isoformat()
    if isinstance(value, float) and not math.isfinite(value):
        return str(value)
    if value is None or isinstance(value, (str, bool, int, float)):
        return value
    return str(value)


def rgb(value: str) -> int:
    if not re.fullmatch(r"#?[0-9a-fA-F]{6}", value):
        raise OfficeError("Color must be six hexadecimal digits, for example #4472C4.")
    raw = value.lstrip("#")
    return int(raw[0:2], 16) | int(raw[2:4], 16) << 8 | int(raw[4:6], 16) << 16


def a1_size(address: str) -> tuple[int, int]:
    match = re.fullmatch(r"\$?([A-Za-z]{1,3})\$?([1-9][0-9]*)(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]*))?", address)
    if not match:
        raise OfficeError("Use a single A1 range such as A1 or A1:C10; expressions and external references are rejected.")

    def column(text):
        number = 0
        for char in text.upper():
            number = number * 26 + ord(char) - 64
        return number

    c1, r1 = column(match[1]), int(match[2])
    c2, r2 = column(match[3] or match[1]), int(match[4] or match[2])
    if not (1 <= c1 <= c2 <= 16384 and 1 <= r1 <= r2 <= 1048576):
        raise OfficeError("Range is reversed or outside Excel worksheet bounds.")
    rows, cols = r2 - r1 + 1, c2 - c1 + 1
    if rows * cols > 10000:
        raise OfficeError("At most 10000 cells per request; split the range.")
    return rows, cols


@dataclass
class Document:
    id: str
    app: object
    doc: object
    engine: str
    kind: str
    path: str | None
    disk_hash: str | None
    opened_by_server: bool = False
    managed: bool = False
    autosave_paused: bool = False
    autosave_state: dict | None = None
    recovery_copy: str | None = None
    live_enabled: bool = False
    live_pause_ms: int = 180
    live_state: dict | None = None


class NativeOffice(ObjectModel, SuiteFeatures, LiveView, PptDiagramMixin):
    def __init__(self, data_dir: str | Path | None = None):
        if sys.platform not in ("win32", "darwin"):
            raise OfficeError("Native editing requires Windows or macOS with Microsoft Office or WPS Office installed.")
        if sys.platform == "win32":
            import pythoncom
            self.com = pythoncom
        else:
            self.com = None

        self.documents: dict[str, Document] = {}
        self.apps: list[tuple[str, str, object]] = []
        base = os.environ.get("MICROSOFT_WPS_LIVE_MCP_DATA") or os.environ.get("OFFICE_LIVE_MCP_DATA")
        if data_dir is None:
            if sys.platform == "win32":
                user_data = Path(os.environ.get("LOCALAPPDATA", Path.home() / ".local" / "share"))
            elif sys.platform == "darwin":
                user_data = Path.home() / "Library" / "Application Support"
            else:
                user_data = Path.home() / ".local" / "share"
            legacy_data = user_data / "office-live-mcp"
            data_dir = base or (legacy_data if legacy_data.is_dir() else user_data / "microsoft-wps-live-mcp")
        self.data_dir = Path(data_dir).expanduser().resolve()

    @staticmethod
    def check_engine(engine):
        if engine not in ("auto", "microsoft", "wps"):
            raise OfficeError("engine must be auto, microsoft, or wps.")

    @staticmethod
    def _app_key(app, vendor, kind):
        if sys.platform == "darwin" or getattr(app, "macos", False):
            return vendor, kind, getattr(app, "app_name", f"macos-{vendor}-{kind}")
        # Office can return different IDispatch/IUnknown proxies for the same app
        # through ROT, GetActiveObject and native-window accessibility.
        try:
            import win32process
            try:
                value = app.Hwnd
            except Exception:
                value = app.ActiveWindow.Hwnd
            hwnd = int(value() if callable(value) else value)
            pid = win32process.GetWindowThreadProcessId(hwnd)[1]
            if pid:
                return vendor, kind, pid
        except Exception:
            pass
        if vendor == "microsoft" and kind == "ppt":
            try:
                # HWND is advertised but unimplemented in some PowerPoint
                # interfaces; use its verified native document window fallback.
                hwnd = LiveView._window_handle(app.ActiveWindow,
                    SimpleNamespace(app=app, doc=app.ActivePresentation))
                if hwnd:
                    return vendor, kind, win32process.GetWindowThreadProcessId(hwnd)[1]
            except Exception:
                pass
        # PyIDispatch repr contains both the Python wrapper and underlying COM pointer.
        # Only the latter is stable when no application window is available.
        raw = repr(app._oleobj_)
        match = re.search(r"obj at (0x[0-9A-Fa-f]+)", raw)
        return vendor, kind, match[1] if match else raw

    def _apps(self, kind):
        if sys.platform == "darwin":
            from .macos import MacAppProxy, is_process_running
            found = []
            for vendor in ("microsoft", "wps"):
                if macos_app_installed(vendor, kind):
                    proxy = MacAppProxy(vendor, kind)
                    if is_process_running(proxy.app_name):
                        found.append((vendor, proxy))
            return found
        found = [(vendor, app) for vendor, k, app in self.apps if k == kind]
        for vendor, kinds in PROGIDS.items():
            for progid in kinds[kind]:
                try:
                    app = dispatch(self.com.GetActiveObject(progid).QueryInterface(self.com.IID_IDispatch))
                    actual = vendor_of(app, vendor)
                    key = self._app_key(app, actual, kind)
                    # A hidden PowerPoint application has no HWND and may issue a new
                    # automation proxy on every query. Prefer the retained connection.
                    if kind == "ppt" and actual == "microsoft" and not isinstance(key[2], int) and any(v == actual for v, _ in found):
                        continue
                    found.append((actual, app))
                except Exception:
                    pass
        # ROT also exposes document monikers from other native application instances.
        rot = self.com.GetRunningObjectTable()
        ctx = self.com.CreateBindCtx(0)
        enum = rot.EnumRunning()
        while True:
            batch = enum.Next(1)
            if not batch:
                break
            moniker = batch[0]
            try:
                name = moniker.GetDisplayName(ctx, None)
                suffix = Path(name).suffix.lower()
                if EXTENSIONS.get(suffix) != kind:
                    continue
                obj = dispatch(rot.GetObject(moniker).QueryInterface(self.com.IID_IDispatch))
                app = obj.Application
                vendor = vendor_of(app)
                if kind == "ppt" and vendor == "microsoft":
                    key = self._app_key(app, vendor, kind)
                    if not isinstance(key[2], int) and any(v == vendor for v, _ in found):
                        continue
                if vendor != "unknown":
                    found.append((vendor, app))
            except Exception:
                continue
        try:
            for app in native_window_apps(kind):
                vendor = vendor_of(app)
                if vendor != "unknown":
                    found.append((vendor, app))
        except Exception:
            pass
        unique = []
        for vendor, app in found:
            try:
                getattr(app, COLLECTIONS[kind]).Count
                if any(self._app_key(app, vendor, kind) == self._app_key(old, v, kind) for v, old in unique):
                    continue
                unique.append((vendor, app))
            except Exception:
                continue
        return unique

    def _path(self, doc) -> str | None:
        try:
            name = str(doc.FullName)
            if Path(name).is_absolute():
                return canonical(name)
        except Exception:
            pass
        return None

    def _wait_document_ready(self, doc, timeout=10):
        # Some WPS builds return an IDispatch before the new/opened document's
        # type information is ready. Retry reads on the same object only; never
        # replay Documents.Add/Open and accidentally create duplicate documents.
        deadline = time.monotonic() + timeout
        last_error = None
        while True:
            try:
                doc.Name
                doc.Saved
                return doc
            except Exception as exc:
                last_error = exc
            if time.monotonic() >= deadline:
                raise OfficeError(f"Native document did not become ready: {last_error}. Creation/open was not replayed.")
            self.com.PumpWaitingMessages()
            time.sleep(0.1)
            doc = dispatch(doc._oleobj_)

    def _remember(self, app, doc, vendor, kind, opened=False):
        path = self._path(doc)
        for record in self.documents.values():
            try:
                same_app = self._app_key(app, vendor, kind) == self._app_key(record.app, record.engine, record.kind)
                if same_app and ((path is not None and path == record.path)
                                 or getattr(record.doc, "_oleobj_", record.doc) == getattr(doc, "_oleobj_", doc)):
                    record.doc.Name  # do not reuse a closed native document reference
                    return record
            except Exception:
                pass
        record = Document(uuid.uuid4().hex, app, doc, vendor, kind, path,
                          file_hash(Path(path)) if path else None, opened)
        record.managed = opened
        self.documents[record.id] = record
        return record

    def _info(self, record):
        doc = record.doc
        path = self._path(doc)
        if path != record.path:
            # A user may use Save As or give an untitled document its first filename.
            record.path = path
            record.disk_hash = file_hash(Path(path)) if path else None
        try:
            read_only = bool(doc.ReadOnly)
        except Exception:
            read_only = False
        return {"document_id": record.id, "name": str(doc.Name), "path": record.path,
                "kind": record.kind, "engine": record.engine, "saved": bool(doc.Saved),
                "read_only": read_only, "opened_by_server": record.opened_by_server,
                "automatic_save": record.managed and not record.autosave_paused,
                "autosave_state": record.autosave_state, "recovery_copy": record.recovery_copy,
                "live_view": dict(record.live_state or {"enabled": record.live_enabled})}

    def list_open(self, engine="auto"):
        self.check_engine(engine)
        if sys.platform == "darwin":
            return self._list_open_macos(engine)
        result = []
        warnings = []
        seen = set()
        # (vendor, kind) pairs whose running application was successfully
        # attached and whose document collection Count read back via COM.
        # This is real automation evidence, not just registry registration.
        reachable = set()
        for kind in COLLECTIONS:
            for vendor, app in self._apps(kind):
                if engine != "auto" and vendor != engine:
                    continue
                try:
                    collection = getattr(app, COLLECTIONS[kind])
                    count = int(collection.Count)
                except Exception as exc:
                    warnings.append({"engine": vendor, "kind": kind, "error": str(exc)})
                    continue
                reachable.add((vendor, kind))
                for index in range(1, count + 1):
                    prior_ids = set(self.documents)
                    try:
                        doc = collection.Item(index)
                        record = self._remember(app, doc, vendor, kind)
                        info = self._info(record)
                        if record.id not in seen:
                            result.append(info)
                            seen.add(record.id)
                    except Exception as exc:
                        # A busy/unready external document must not hide all other
                        # documents, become managed, or crash environment checks.
                        for identifier in set(self.documents) - prior_ids:
                            self.documents.pop(identifier, None)
                        warnings.append({"engine": vendor, "kind": kind, "index": index, "error": str(exc)})
        return {"documents": result, "discovery_warnings": warnings,
                "reachable_engines": sorted(f"{v}/{k}" for v, k in reachable)}

    def _list_open_macos(self, engine="auto"):
        from .macos import MacAppProxy, MacDocProxy, get_app_name, is_process_running, list_open_docs
        result, warnings, seen, reachable = [], [], set(), set()
        for kind in COLLECTIONS:
            for vendor in ("microsoft", "wps"):
                if engine != "auto" and vendor != engine:
                    continue
                if macos_app_installed(vendor, kind) and is_process_running(get_app_name(vendor, kind)):
                    reachable.add((vendor, kind))
        for item in list_open_docs():
            if engine != "auto" and item["engine"] != engine:
                continue
            kind, vendor, full_name = item["kind"], item["engine"], item["full_name"]
            target = canonical(full_name) if full_name and Path(full_name).is_absolute() else None
            record = None
            if target:
                for existing in self.documents.values():
                    if existing.path == target:
                        record = existing
                        break
                if record is None and Path(full_name).is_file():
                    try:
                        app = MacAppProxy(vendor, kind)
                        doc = MacDocProxy(full_name, kind, vendor)
                        record = self._remember(app, doc, vendor, kind, opened=False)
                    except Exception as exc:
                        warnings.append({"engine": vendor, "kind": kind, "error": str(exc)})
            if record is not None and record.id not in seen:
                result.append(self._info(record))
                seen.add(record.id)
            elif record is None:
                result.append({"document_id": None, "name": item["name"], "path": target,
                               "kind": kind, "engine": vendor, "saved": True, "read_only": False,
                               "opened_by_server": False, "automatic_save": False})
        for record in self.documents.values():
            if record.id not in seen and (engine == "auto" or record.engine == engine):
                result.append(self._info(record))
                seen.add(record.id)
        return {"documents": result, "discovery_warnings": warnings,
                "reachable_engines": sorted(f"{v}/{k}" for v, k in reachable)}

    def status(self):
        result = capabilities()
        result["data_dir"] = str(self.data_dir)
        result.update(self.list_open())
        self._annotate_connection(result)
        return result

    def _annotate_connection(self, result):
        """Mark each engine/kind with a real COM connectivity verdict.

        Non-intrusive: reflects only what list_open already observed by
        attaching to running applications. No application is started here.
        Use the --probe command (or probe()) to launch-and-verify closed apps.
        """
        reachable = {tuple(pair.split("/", 1)) for pair in result.get("reachable_engines", [])}
        docs = result.get("documents", [])
        warnings = result.get("discovery_warnings", [])
        warn_by = {}
        for w in warnings:
            key = (w.get("engine"), w.get("kind"))
            if key[0] is not None:
                warn_by.setdefault(key, []).append(w.get("error", ""))
        engines = result.get("engines", {})
        for vendor, kinds in engines.items():
            for kind, info in kinds.items():
                if not info.get("registered"):
                    info["connection"] = "unavailable"
                    continue
                if (vendor, kind) in reachable:
                    info["connection"] = "verified_running"
                elif (vendor, kind) in warn_by:
                    info["connection"] = "error"
                    info["connection_error"] = "; ".join(warn_by[(vendor, kind)])[:500]
                else:
                    info["connection"] = "registered_not_running"
        result["note"] = (
            "connection=verified_running: a running app was attached and its document collection Count read via COM "
            "(real automation proven). connection=registered_not_running: ProgID is registered but no instance is "
            "running; status never starts apps. connection=error: attach failed. "
            "Run `--probe` to launch-and-verify closed apps. Registration alone is not a connection test."
        )

    def probe(self, engine="auto", kind=None):
        """Verify COM automation for every registered engine/kind.

        Attaches to already-running instances without launching. For a kind
        with no running instance, launches the registered automation server,
        reads identifying properties, and quits only the instance this probe
        started. Pre-existing user applications are never touched or quit.
        """
        self.check_engine(engine)
        if kind is not None and kind not in COLLECTIONS:
            raise OfficeError("kind must be one of: " + ", ".join(COLLECTIONS))
        if sys.platform == "darwin":
            return self._probe_macos(engine, kind)
        engines = {}
        quit_warnings = []
        for vendor, kinds in PROGIDS.items():
            if engine != "auto" and vendor != engine:
                continue
            engines[vendor] = {}
            for k, progids in kinds.items():
                if kind is not None and k != kind:
                    continue
                entry, warn = self._probe_one(vendor, k, progids)
                engines[vendor][k] = entry
                quit_warnings.extend(warn)
        # Release any remaining COM wrappers promptly so launched apps that were
        # Quit can actually exit, instead of lingering on held IDispatch refs.
        import gc
        gc.collect()
        return {
            "platform": sys.platform,
            "engines": engines,
            "probe_quit_warnings": quit_warnings,
            "note": ("--probe verifies COM by attaching to running apps or launching registered ones; "
                     "only instances started by this probe are quit. Pre-existing apps are left untouched. "
                     "Launched apps are hidden (Visible=False); Office click-to-run may take a few seconds "
                     "to finalize process exit after Quit."),
        }

    def _probe_macos(self, engine="auto", kind=None):
        from .macos import get_app_name, is_process_running
        engines = {}
        for vendor, kinds in MACOS_APPS.items():
            if engine != "auto" and vendor != engine:
                continue
            engines[vendor] = {}
            for k, targets in kinds.items():
                if kind is not None and k != kind:
                    continue
                installed = macos_app_installed(vendor, k)
                app_name = get_app_name(vendor, k)
                running = is_process_running(app_name) if installed else False
                conn = "verified_running" if running else ("registered_not_running" if installed else "unavailable")
                engines[vendor][k] = {
                    "registered": installed,
                    "running": running,
                    "app_name": app_name,
                    "targets": list(targets),
                    "connection": conn,
                }
        return {
            "platform": sys.platform,
            "engines": engines,
            "probe_quit_warnings": [],
            "note": "macOS uses AppleScript and OOXML document bridge; no COM registration required.",
        }

    def _probe_one(self, vendor, kind, progids):
        registered_progids = [p for p in progids if registration(p)]
        entry = {"registered": bool(registered_progids), "progids": list(progids)}
        quit_warnings = []
        if not registered_progids:
            entry["connection"] = "unavailable"
            return entry, quit_warnings
        # Prefer attaching to a running instance; never launch a duplicate.
        try:
            running = next((a for v, a in self._apps(kind) if v == vendor), None)
        except Exception:
            running = None
        if running is not None:
            entry.update(self._probe_read(running, kind))
            entry["running"] = True
            entry["started_by_probe"] = False
            entry["connection"] = "verified_running"
            return entry, quit_warnings
        # No running instance; launch each registered ProgID until one connects.
        attempts = []
        for progid in registered_progids:
            try:
                app = dispatch(retry_busy(lambda: self.com.CoCreateInstanceEx(
                    progid, None, self.com.CLSCTX_SERVER, None, (self.com.IID_IDispatch,))[0]))
            except Exception as exc:
                attempts.append({"progid": progid, "launched": False, "error": str(exc)[:300]})
                continue
            # Word/PowerPoint COM servers default to visible. Hide immediately so
            # a probe never flashes a window at the user.
            try:
                app.Visible = False
            except Exception:
                pass
            actual = vendor_of(app, vendor)
            if actual != vendor:
                attempts.append({"progid": progid, "launched": True, "vendor_mismatch": actual})
                self._quit_launched(app, vendor, kind, progid, quit_warnings)
                del app
                continue
            entry.update(self._probe_read(app, kind))
            entry["running"] = False
            entry["started_by_probe"] = True
            entry["launched_progid"] = progid
            entry["connection"] = "verified_launched"
            entry["launch_attempts"] = attempts
            # Quit now, while the only COM reference is this local variable;
            # holding it across the rest of probe() keeps the app process alive.
            self._quit_launched(app, vendor, kind, progid, quit_warnings)
            del app
            return entry, quit_warnings
        entry["launch_attempts"] = attempts
        entry["connection"] = "launch_failed"
        return entry, quit_warnings

    @staticmethod
    def _quit_launched(app, vendor, kind, progid, quit_warnings):
        try:
            try:
                app.DisplayAlerts = False
            except Exception:
                pass
            app.Quit()
        except Exception as exc:
            quit_warnings.append({"engine": vendor, "kind": kind, "progid": progid,
                                  "quit_error": str(exc)[:300]})

    @staticmethod
    def _probe_read(app, kind):
        info = {}
        for prop in ("Name", "Version", "Build"):
            try:
                info[prop.lower()] = str(getattr(app, prop))
            except Exception:
                continue
        try:
            info["document_count"] = int(getattr(app, COLLECTIONS[kind]).Count)
        except Exception as exc:
            info["document_count_error"] = str(exc)[:200]
        return info

    def _validate_path(self, filename, exists=True):
        path = Path(filename).expanduser()
        if not path.is_absolute():
            raise OfficeError("Use an absolute local document path.")
        path = path.resolve()
        if path.suffix.lower() not in EXTENSIONS:
            raise OfficeError("Supported native formats are .docx, .xlsx and .pptx. Convert other formats explicitly first.")
        if exists:
            if not path.is_file():
                raise OfficeError(f"File not found: {path}")
            try:
                with zipfile.ZipFile(path) as archive:
                    names = archive.namelist()
                    if any(n.lower().endswith("vbaproject.bin") for n in names):
                        raise OfficeError("Macro-containing documents are not accepted by this adapter.")
            except zipfile.BadZipFile as exc:
                raise OfficeError("The file is encrypted, damaged, or is not an OOXML document.") from exc
        elif path.exists():
            raise OfficeError("Destination exists; create and Save As never overwrite an existing file.")
        elif not path.parent.is_dir():
            raise OfficeError("Destination directory does not exist.")
        return path

    def _choose_app(self, kind, engine):
        vendors = [engine] if engine != "auto" else ["microsoft", "wps"]
        if sys.platform == "darwin":
            from .macos import MacAppProxy
            for vendor in vendors:
                if macos_app_installed(vendor, kind):
                    for actual, app_kind, app in self.apps:
                        if actual == vendor and app_kind == kind:
                            return app, vendor
                    app = MacAppProxy(vendor, kind)
                    self.apps.append((vendor, kind, app))
                    return app, vendor
            raise OfficeError(f"No usable {engine} {kind} application found on macOS. Install Microsoft Office or WPS Office.")
        errors = []
        for vendor in vendors:
            # Reuse applications activated by this server. External open files
            # are attached separately in open(), before any activation occurs.
            for actual, app_kind, app in self.apps:
                if actual == vendor and app_kind == kind:
                    try:
                        getattr(app, COLLECTIONS[kind]).Count
                        return app, vendor
                    except Exception:
                        continue
            for progid in PROGIDS[vendor][kind]:
                if not registration(progid):
                    continue
                try:
                    # Bypass compatibility ROT entries. WPS may advertise Office
                    # identities there, or return an app whose document proxies
                    # are stale. Activate the registered automation server directly;
                    # never change registrations or replay document creation.
                    app = dispatch(retry_busy(lambda: self.com.CoCreateInstanceEx(
                        progid, None, self.com.CLSCTX_SERVER, None,
                        (self.com.IID_IDispatch,))[0]))
                    actual = vendor_of(app)
                    if actual != vendor:
                        errors.append(f"{progid} belongs to {actual}, not {vendor}")
                        continue
                    self.apps.append((vendor, kind, app))
                    return app, vendor
                except Exception as exc:
                    errors.append(f"{progid}: {exc}")
        raise OfficeError("No usable " + engine + " " + kind + " automation interface. Install desktop Office "
                          "or a WPS build exposing COM, and run in the same Windows user session. " + "; ".join(errors))

    @contextmanager
    def _opening_settings(self, app):
        previous = None
        try:
            previous = app.AutomationSecurity
            app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        except Exception:
            pass
        try:
            yield
        finally:
            if previous is not None:
                app.AutomationSecurity = previous

    @contextmanager
    def _excel_events(self, record):
        if sys.platform == "darwin" or getattr(record.app, "macos", False):
            yield
            return
        previous = record.app.EnableEvents
        record.app.EnableEvents = False
        try:
            yield
        finally:
            record.app.EnableEvents = previous

    def open(self, filename, engine="auto", visible=True):
        self.check_engine(engine)
        path = self._validate_path(filename)
        kind = EXTENSIONS[path.suffix.lower()]
        target = canonical(path)
        if sys.platform == "darwin":
            return self._open_macos(path, kind, target, engine, visible)
        matches = []
        for vendor, app in self._apps(kind):
            collection = getattr(app, COLLECTIONS[kind])
            for i in range(1, int(collection.Count) + 1):
                doc = collection.Item(i)
                if self._path(doc) == target:
                    matches.append(self._remember(app, doc, vendor, kind))
        if len(matches) > 1:
            raise OfficeError("This file is open in multiple application instances. Close the extra copies before editing.")
        if matches:
            record = matches[0]
            if engine != "auto" and record.engine != engine:
                raise OfficeError(f"File is already open in {record.engine}. Use that engine or close it first.")
            record.managed = True
            record.live_enabled = visible
            if visible:
                self._ensure_live_window(record, foreground=True)
            return {**self._info(record), "already_open": True}
        app, vendor = self._choose_app(kind, engine)
        with self._opening_settings(app):
            if kind == "word":
                doc = app.Documents.Open(FileName=str(path), ConfirmConversions=False, ReadOnly=False,
                                         AddToRecentFiles=False, Visible=visible,
                                         PasswordDocument="", WritePasswordDocument="")
            elif kind == "excel":
                previous = app.EnableEvents
                app.EnableEvents = False
                try:
                    doc = app.Workbooks.Open(Filename=str(path), UpdateLinks=0, ReadOnly=False,
                                             Password="", WriteResPassword="", IgnoreReadOnlyRecommended=True,
                                             AddToMru=False)
                finally:
                    app.EnableEvents = previous
            else:
                doc = app.Presentations.Open(str(path), False, False, visible)
        doc = self._wait_document_ready(doc)
        record = self._remember(app, doc, vendor, kind, opened=True)
        record.live_enabled = visible
        if visible:
            self._ensure_live_window(record, foreground=True)
        return {**self._info(record), "already_open": False}

    def create(self, filename, engine="auto", visible=True):
        self.check_engine(engine)
        path = self._validate_path(filename, exists=False)
        kind = EXTENSIONS[path.suffix.lower()]
        if sys.platform == "darwin":
            return self._create_macos(path, kind, engine, visible)
        app, vendor = self._choose_app(kind, engine)
        if kind == "word":
            doc = app.Documents.Add()
        elif kind == "excel":
            previous = app.EnableEvents
            app.EnableEvents = False
            try:
                doc = app.Workbooks.Add()
            finally:
                app.EnableEvents = previous
        else:
            doc = app.Presentations.Add(visible)
        doc = self._wait_document_ready(doc)
        if kind == "ppt" and int(doc.Slides.Count) == 0:
            # Some PowerPoint builds return a new presentation with no slide,
            # which makes Slides.Item(1) fail for every later tool.
            doc.Slides.Add(1, 12)  # ppLayoutBlank
        record = self._remember(app, doc, vendor, kind, opened=True)
        record.live_enabled = visible
        if visible:
            self._ensure_live_window(record, foreground=True)
        result = self.save(record.id, str(path))
        if visible:
            # Save As can take a moment; re-assert foreground so the window is not
            # left behind other windows when the tool returns.
            self._ensure_live_window(record, foreground=True)
        return result

    def _open_macos(self, path: Path, kind: str, target: str, engine: str, visible: bool):
        from .macos import MacDocProxy, activate_app, list_open_docs, open_in_app
        for record in self.documents.values():
            if record.path == target:
                if engine != "auto" and record.engine != engine:
                    raise OfficeError(f"File is already open in {record.engine}. Use that engine or close it first.")
                record.managed = True
                record.live_enabled = visible
                if visible:
                    self._ensure_live_window(record, foreground=True)
                return {**self._info(record), "already_open": True}
        already_open = False
        matched_vendor = engine
        for item in list_open_docs(kind):
            full = item.get("full_name") or ""
            if (full and canonical(full) == target) or item.get("name") == path.name:
                already_open = True
                if engine != "auto" and item["engine"] != engine:
                    raise OfficeError(f"File is already open in {item['engine']}. Use that engine or close it first.")
                matched_vendor = item["engine"]
                break
        app, vendor = self._choose_app(kind, matched_vendor)
        if not already_open:
            open_in_app(app.app_name, path, visible=visible)
        elif visible:
            activate_app(app.app_name)
        doc = MacDocProxy(path, kind, vendor)
        record = self._remember(app, doc, vendor, kind, opened=True)
        record.live_enabled = visible
        if visible:
            self._ensure_live_window(record, foreground=True)
        return {**self._info(record), "already_open": already_open}

    def _create_macos(self, path: Path, kind: str, engine: str, visible: bool):
        from .macos import MacDocProxy, create_blank_ooxml, open_in_app
        app, vendor = self._choose_app(kind, engine)
        create_blank_ooxml(path, kind)
        doc = MacDocProxy(path, kind, vendor)
        record = self._remember(app, doc, vendor, kind, opened=True)
        record.live_enabled = visible
        result = self.save(record.id)
        if visible:
            open_in_app(app.app_name, path, visible=True)
            self._ensure_live_window(record, foreground=True)
            result.update(self._info(record))
        return result

    def resolve(self, document, kind=None):
        if document in self.documents:
            record = self.documents[document]
            try:
                record.doc.Name
                record.doc.Saved
            except Exception as exc:
                if getattr(exc, "hresult", None) in (-2147418111, -2147417846):
                    raise OfficeError("Native application is busy. Existing document was not reopened or modified.") from exc
                if not record.path:
                    raise OfficeError("The untitled document was closed. Supply a saved file path.")
                self._forget_objects(record.id)
                self.__dict__.get("live_windows", {}).pop(record.id, None)
                info = self.open(record.path, record.engine, visible=record.live_enabled)
                replacement = self.documents[info["document_id"]]
                record.app, record.doc = replacement.app, replacement.doc
                record.disk_hash = replacement.disk_hash
                record.opened_by_server = True
                if replacement.id != record.id:
                    self.documents.pop(replacement.id, None)
            self._info(record)
        else:
            info = self.open(document)
            record = self.documents[info["document_id"]]
        if kind and record.kind != kind:
            raise OfficeError(f"This tool requires a {kind} document, got {record.kind}.")
        if not record.managed:
            record.live_enabled = True
            self._ensure_live_window(record)
        record.managed = True
        return record

    def _before_edit(self, record):
        if record.live_enabled:
            self._ensure_live_window(record)
        if self._info(record)["read_only"]:
            raise OfficeError("Document is read-only or locked by another application; the live document was not changed.")
        if record.path and file_hash(Path(record.path)) != record.disk_hash:
            raise OfficeError("Disk file changed since attachment. Save As to a new path or resolve the conflict; "
                              "the adapter cannot infer whether a disk change came from this window or another program.")
        return self._backup(record)

    def _backup(self, record):
        if not record.path or not Path(record.path).is_file():
            return None
        folder = self.data_dir / "backups" / record.id
        folder.mkdir(parents=True, exist_ok=True)
        digest = file_hash(Path(record.path))
        previous = self.__dict__.setdefault("last_backup", {}).get(record.id)
        if previous and previous[0] == digest and Path(previous[1]).is_file():
            return previous[1]
        timestamp = dt.datetime.now(dt.timezone.utc).strftime("%Y%m%dT%H%M%S%fZ")
        destination = folder / (timestamp + "-" + Path(record.path).name)
        shutil.copy2(record.path, destination)
        self.last_backup[record.id] = (digest, str(destination))
        # Preserve the first backup plus the latest 49 versions, in this dedicated
        # document backup directory only. No recursive deletion or source deletion.
        versions = sorted(folder.iterdir())
        for old in versions[1:-49]:
            if old.is_file() and not old.is_symlink() and old.resolve().parent == folder.resolve():
                old.unlink()
        return str(destination)

    def _update_recovery_copy(self, record):
        key = hashlib.sha256(record.path.encode("utf-8")).hexdigest()[:24]
        folder = self.data_dir / "recovery" / key
        folder.mkdir(parents=True, exist_ok=True)
        destination = folder / Path(record.path).name
        temporary = folder / (uuid.uuid4().hex + ".tmp")
        try:
            shutil.copy2(record.path, temporary)
            if file_hash(temporary) != record.disk_hash:
                raise OfficeError("Original file changed while the recovery copy was being written.")
            os.replace(temporary, destination)
        finally:
            if temporary.is_file():
                temporary.unlink()
        record.recovery_copy = str(destination)
        return str(destination)

    def autosave_tick(self):
        """Called while the STA worker is idle; observes only explicitly connected documents."""
        for record in list(self.documents.values()):
            if not record.managed or record.autosave_paused:
                continue
            try:
                dirty = not bool(record.doc.Saved)
            except Exception:
                continue  # user closed the document; never reopen it just to poll
            if not dirty:
                continue
            if not record.path:
                record.autosave_state = {"state": "needs_path", "message": "Untitled document requires a save_as path."}
                continue
            try:
                self.save(record.id)
            except Exception as exc:
                record.autosave_paused = True
                record.autosave_state = {"state": "error", "message": str(exc), "inspect_before_retry": True}

    def enforce_autosave(self, result):
        # Successful public modifications cannot opt out of automatic disk saving.
        # Internal multi-step helpers may still stage edits before their final save.
        if result.get("applied") and not result.get("disk_saved") and result.get("state") != "partial_error":
            saved = self.save(result["document_id"])
            result = {**result, **saved, "autosave_enforced": True}
        return result

    def pause_after_error(self, args):
        if not args or not isinstance(args[0], str):
            return
        selected = args[0]
        for record in self.documents.values():
            if record.id == selected or (record.path and canonical(selected) == record.path):
                try:
                    if not bool(record.doc.Saved):
                        record.autosave_paused = True
                        record.autosave_state = {"state": "needs_inspection", "message": "An operation failed while live edits were pending. Inspect, then explicitly save to resume automatic saving."}
                except Exception:
                    pass

    def _result(self, record, result, save, backup=None):
        output = {**self._info(record), "applied": True, "operation": result,
                  "disk_saved": False, "backup": backup}
        if save:
            output.update(self.save(record.id))
            output["operation"] = result
            output["backup_before_edit"] = backup
        if record.live_enabled:
            focus = result.get("live_focus")
            if record.kind == "word":
                start = result.get("formatted_start", result.get("last_position", result.get("position")))
                if focus is None and isinstance(start, int):
                    focus = {"start": start, "end": result.get("formatted_end", start)}
            elif focus is None and record.kind == "excel" and "range" in result:
                focus = {"sheet": result.get("sheet", 1), "address": result["range"]}
            elif focus is None and record.kind == "ppt" and "slide" in result:
                focus = {"slide": result["slide"]}
            output["live_view"] = self.live_checkpoint(record, focus)
        return output

    def save(self, document, save_as=None):
        record = self.resolve(document)
        if self._info(record)["read_only"]:
            raise OfficeError("Document is read-only. Original file was not overwritten.")
        if save_as:
            target = self._validate_path(save_as, exists=False)
            if EXTENSIONS[target.suffix.lower()] != record.kind:
                raise OfficeError("Save As must keep the document type.")
        else:
            if not record.path:
                raise OfficeError("Untitled document requires an absolute save_as path.")
            target = Path(record.path)
            if file_hash(target) != record.disk_hash:
                raise OfficeError("Disk file changed since attachment. Live edits remain open; use Save As or resolve the conflict.")
        backup = self._backup(record)
        with self._excel_events(record) if record.kind == "excel" else _no_context():
            if save_as:
                if record.kind == "word":
                    # SaveAs is implemented by both Microsoft Word and WPS Writer.
                    record.doc.SaveAs(str(target), 12)
                elif record.kind == "excel":
                    record.doc.SaveAs(str(target), 51)
                else:
                    record.doc.SaveAs(str(target), 24)  # ppSaveAsOpenXMLPresentation
            else:
                record.doc.Save()
        if not target.is_file() or not bool(record.doc.Saved):
            raise OfficeError("Application did not confirm disk save. Live edits may still be pending; do not replay insertions.")
        with zipfile.ZipFile(target) as archive:
            required = {"word": "word/document.xml", "excel": "xl/workbook.xml", "ppt": "ppt/presentation.xml"}[record.kind]
            if required not in archive.namelist():
                raise OfficeError("Saved file does not contain the expected Office document structure.")
        record.path = canonical(target)
        record.disk_hash = file_hash(target)
        record.autosave_paused = False
        record.autosave_state = {"state": "saved", "time": dt.datetime.now(dt.timezone.utc).isoformat()}
        try:
            self._update_recovery_copy(record)
            record.autosave_state["recovery_state"] = "updated"
        except Exception as exc:
            record.autosave_state["recovery_state"] = "error"
            record.autosave_state["recovery_error"] = str(exc)
        return {**self._info(record), "disk_saved": True, "sha256": record.disk_hash,
                "size_bytes": target.stat().st_size, "backup": backup}

    def create_copy(self, document, filename=None):
        record = self.resolve(document)
        saved = self.save(record.id)
        source = Path(record.path)
        if filename:
            destination = self._validate_path(filename, exists=False)
        else:
            timestamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
            destination = source.with_name(source.stem + ".copy-" + timestamp + "-" + uuid.uuid4().hex[:6] + source.suffix)
        if EXTENSIONS[destination.suffix.lower()] != record.kind:
            raise OfficeError("Copy must keep the document type.")
        # Exclusive creation prevents races from overwriting an existing user's file.
        with source.open("rb") as incoming, destination.open("xb") as outgoing:
            shutil.copyfileobj(incoming, outgoing)
        if file_hash(destination) != record.disk_hash:
            raise OfficeError("Copy verification failed; original was not renamed or removed.")
        return {**saved, "copy_created": True, "copy_path": str(destination), "copy_sha256": file_hash(destination)}

    def close(self, document, save=True):
        record = self.resolve(document)
        result = self.save(record.id) if save else self._info(record)
        if not save and not bool(record.doc.Saved):
            raise OfficeError("Document has unsaved edits. Save it or Save As before closing; this tool does not discard user edits.")
        if record.kind == "ppt":
            record.doc.Close()
        else:
            record.doc.Close(False)
        self.documents.pop(record.id, None)
        self._forget_objects(record.id)
        self.__dict__.get("live_windows", {}).pop(record.id, None)
        return {**result, "closed": True}

    def word_read(self, document, start=1, count=100):
        r = self.resolve(document, "word")
        if start < 1 or not 1 <= count <= 500:
            raise OfficeError("start is 1-based; count must be between 1 and 500.")
        total = int(r.doc.Paragraphs.Count)
        paragraphs = []
        for i in range(start, min(total + 1, start + count)):
            span = r.doc.Paragraphs.Item(i).Range
            paragraphs.append({"index": i, "start": int(span.Start), "end": int(span.End),
                               "text": str(span.Text)[:20000]})
        return {**self._info(r), "total_paragraphs": total, "tables": int(r.doc.Tables.Count),
                "paragraphs": paragraphs, "next_start": start + count if start + count <= total else None}

    def word_insert(self, document, text, position="end", save=True):
        r = self.resolve(document, "word")
        if not text or len(text) > 200000:
            raise OfficeError("text must contain 1 to 200000 characters.")
        end = max(0, int(r.doc.Content.End) - 1)
        if position == "end":
            offset = end
        elif position == "start":
            offset = 0
        elif isinstance(position, int) and not isinstance(position, bool) and 0 <= position <= end:
            offset = position
        else:
            raise OfficeError("position must be start, end, or a valid zero-based native Word character offset.")
        backup = self._before_edit(r)
        content = text.replace("\r\n", "\r").replace("\n", "\r")
        parts = content.splitlines(keepends=True) if r.live_enabled else [content]
        if r.live_enabled and len(parts) == 1 and len(content) > 400:
            parts = [content[i:i + 400] for i in range(0, len(content), 400)]
        group = max(1, math.ceil(len(parts) / 24))
        chunks = ["".join(parts[i:i + group]) for i in range(0, len(parts), group)]
        cursor = offset
        for index, chunk in enumerate(chunks, 1):
            r.doc.Range(cursor, cursor).InsertBefore(chunk)
            start = cursor
            cursor += len(chunk.encode("utf-16-le")) // 2
            if r.live_enabled and len(chunks) > 1:
                self.live_checkpoint(r, {"start": start, "end": cursor}, pause_ms=140, label=f"Text {index}/{len(chunks)}")
        return self._result(r, {"inserted_characters": len(text), "position": offset,
                                "last_position": cursor, "visible_steps": len(chunks),
                                "live_focus": {"start": start, "end": cursor}}, save, backup)

    def word_replace(self, document, search, replacement, match_case=False, whole_word=False, replace_all=True, save=True):
        r = self.resolve(document, "word")
        if not search or len(search) > 255 or len(replacement) > 255:
            raise OfficeError("Native Word Find accepts a nonempty search and at most 255 characters per search/replacement.")
        backup = self._before_edit(r)
        span = r.doc.Content.Duplicate
        finder = span.Find
        finder.ClearFormatting()
        finder.Replacement.ClearFormatting()
        found = bool(finder.Execute(FindText=search, MatchCase=match_case, MatchWholeWord=whole_word,
                                    MatchWildcards=False, Forward=True, Wrap=0, Format=False,
                                    ReplaceWith=replacement, Replace=2 if replace_all else 1))
        return self._result(r, {"matched": found, "replace_all": replace_all}, save, backup)

    def word_format(self, document, start, end, font_name=None, font_size=None, bold=None,
                    italic=None, color=None, ascii_font_name=None, save=True):
        r = self.resolve(document, "word")
        if not (0 <= start < end <= int(r.doc.Content.End)):
            raise OfficeError("start/end must be valid native Word character offsets; read paragraphs first.")
        self._validate_font(font_size, color)
        backup = self._before_edit(r)
        self._font(r.doc.Range(start, end).Font, font_name, font_size, bold, italic, color,
                   ascii_name=ascii_font_name)
        return self._result(r, {"formatted_start": start, "formatted_end": end}, save, backup)

    def word_add_table(self, document, values, position="end", style="default",
                       font_name=None, ascii_font_name=None, font_size=None,
                       header_fill=None, save=True):
        r = self.resolve(document, "word")
        self._matrix(values)
        if style not in ("default", "grid", "three_line"):
            raise OfficeError("style must be default, grid, or three_line.")
        self._validate_font(font_size, header_fill)
        if len(values) * len(values[0]) > 2000:
            raise OfficeError("At most 2000 table cells per request.")
        end = max(0, int(r.doc.Content.End) - 1)
        offset = end if position == "end" else position
        if isinstance(offset, bool) or not isinstance(offset, int) or not 0 <= offset <= end:
            raise OfficeError("position must be end or a valid character offset.")
        backup = self._before_edit(r)
        table = r.doc.Tables.Add(r.doc.Range(offset, offset), len(values), len(values[0]))
        for row, data in enumerate(values, 1):
            for col, value in enumerate(data, 1):
                table.Cell(row, col).Range.Text = str(value if value is not None else "")
            group = max(1, math.ceil(len(values) / 24))
            if r.live_enabled and (row % group == 0 or row == len(values)):
                cell_range = table.Cell(row, 1).Range
                self.live_checkpoint(r, {"start": int(cell_range.Start), "end": int(cell_range.End)},
                                     pause_ms=120, label=f"Table row {row}/{len(values)}")
        if style != "default" or any(x is not None for x in (font_name, ascii_font_name, font_size, header_fill)):
            self._style_word_table(table, style, font_name, ascii_font_name, font_size, header_fill)
            if r.live_enabled:
                self.live_checkpoint(r, {"start": int(table.Range.Start), "end": int(table.Range.End)},
                                     pause_ms=160, label=f"Styled {style} table")
        return self._result(r, {"table_index": int(r.doc.Tables.Count), "rows": len(values),
                                "columns": len(values[0]), "style": style}, save, backup)

    @classmethod
    def _style_word_table(cls, table, style, font_name, ascii_font_name, font_size, header_fill):
        try:
            table.Rows.Alignment = 1  # wdAlignRowCenter
        except Exception:
            pass
        try:
            table.Range.ParagraphFormat.Alignment = 1  # wdAlignParagraphCenter
        except Exception:
            pass
        for attr, val in (("TopPadding", 4.5), ("BottomPadding", 4.5), ("LeftPadding", 6.0), ("RightPadding", 6.0)):
            try:
                setattr(table, attr, val)
            except Exception:
                pass
        cls._font(table.Range.Font, font_name, font_size, None, None, None, ascii_name=ascii_font_name)
        if style == "three_line":
            table.Borders.Enable = 0
            # wdBorderTop=-1 (1.5pt=12), wdBorderBottom=-3 (1.5pt=12), header bottom=0.75pt (6)
            for border_id, width in ((-1, 12), (-3, 12)):
                b = table.Borders.Item(border_id)
                b.LineStyle = 1
                b.LineWidth = width
            hdr = table.Rows.Item(1)
            hb = hdr.Borders.Item(-3)
            hb.LineStyle = 1
            hb.LineWidth = 6
            hdr.Range.Font.Bold = -1
            if header_fill:
                fill_val = rgb(header_fill)
                hdr.Shading.BackgroundPatternColor = fill_val
                if not cls._is_light(fill_val):
                    hdr.Range.Font.Color = 0xFFFFFF
        elif style == "grid":
            table.Borders.Enable = 1
            hdr = table.Rows.Item(1)
            hdr.Range.Font.Bold = -1
            fill_hex = header_fill or "#1F4E79"
            fill_val = rgb(fill_hex)
            hdr.Shading.BackgroundPatternColor = fill_val
            if not cls._is_light(fill_val):
                hdr.Range.Font.Color = 0xFFFFFF

    @staticmethod
    def _matrix(values):
        if (not isinstance(values, list) or not values or not all(isinstance(row, list) and row for row in values)
                or any(len(row) != len(values[0]) for row in values)):
            raise OfficeError("values must be a nonempty rectangular two-dimensional array.")
        if any(not (v is None or isinstance(v, (str, int, float, bool))) for row in values for v in row):
            raise OfficeError("Cell values must be strings, numbers, booleans, or null.")
        if any(isinstance(v, float) and not math.isfinite(v) for row in values for v in row):
            raise OfficeError("NaN and infinity are not valid cell values.")

    def _sheet(self, record, sheet):
        try:
            return record.doc.Worksheets.Item(sheet)
        except Exception as exc:
            raise OfficeError(f"Worksheet not found: {sheet}. Use excel_sheets to list actual names.") from exc

    def excel_read(self, document, sheet, address):
        r = self.resolve(document, "excel")
        rows, cols = a1_size(address)
        cell = self._sheet(r, sheet).Range(address)

        def matrix(value):
            return [[json_value(value)]] if rows == cols == 1 else json_value(value)

        return {**self._info(r), "sheet": str(cell.Worksheet.Name), "range": address,
                "values": matrix(cell.Value2), "formulas": matrix(cell.Formula)}

    def excel_write(self, document, sheet, address, values, formulas=False, save=True):
        r = self.resolve(document, "excel")
        rows, cols = a1_size(address)
        self._matrix(values)
        if len(values) != rows or len(values[0]) != cols:
            raise OfficeError(f"values must have exactly {rows} rows and {cols} columns.")
        cell = self._sheet(r, sheet).Range(address)
        if formulas:
            data = tuple(tuple(v for v in row) for row in values)
        else:
            # Keep user strings literal; formulas must be requested explicitly.
            data = tuple(tuple("'" + v if isinstance(v, str) else v
                               for v in row) for row in values)
        backup = self._before_edit(r)
        last_address = address
        with self._excel_events(r):
            if formulas:
                cell.Formula = data
            elif r.live_enabled and rows > 1 and cell.MergeCells is False:
                first, last = address.replace("$", "").split(":") if ":" in address else (address.replace("$", ""), address.replace("$", ""))
                left = re.match(r"[A-Za-z]+", first)[0]
                right = re.match(r"[A-Za-z]+", last)[0]
                first_row = int(re.search(r"[0-9]+", first)[0])
                group = max(1, math.ceil(rows / 24))
                for offset in range(0, rows, group):
                    count = min(group, rows - offset)
                    last_address = f"{left}{first_row + offset}:{right}{first_row + offset + count - 1}"
                    self._sheet(r, sheet).Range(last_address).Value2 = data[offset:offset + count]
                    self.live_checkpoint(r, {"sheet": str(cell.Worksheet.Name), "address": last_address},
                                         pause_ms=120, label=f"Data rows {offset + 1}-{offset + count}/{rows}")
            else:
                cell.Value2 = data
        return self._result(r, {"sheet": str(cell.Worksheet.Name), "range": address,
                                "cells_written": rows * cols, "formulas": formulas,
                                "live_focus": {"sheet": str(cell.Worksheet.Name), "address": last_address}}, save, backup)

    def excel_sheets(self, document, action="list", name=None, new_name=None, save=True):
        r = self.resolve(document, "excel")
        if action == "list":
            return {**self._info(r), "sheets": [str(r.doc.Worksheets.Item(i).Name)
                    for i in range(1, int(r.doc.Worksheets.Count) + 1)]}
        if action not in ("add", "rename"):
            raise OfficeError("action must be list, add, or rename.")
        target = name if action == "add" else new_name
        if not target or len(target) > 31 or re.search(r"[\[\]:*?/\\]", target) or target.startswith("'") or target.endswith("'"):
            raise OfficeError("Invalid worksheet name.")
        names = [str(r.doc.Worksheets.Item(i).Name).casefold() for i in range(1, int(r.doc.Worksheets.Count) + 1)]
        if target.casefold() in names:
            raise OfficeError("A worksheet with that name already exists.")
        sheet = self._sheet(r, name) if action == "rename" else None
        backup = self._before_edit(r)
        with self._excel_events(r):
            if action == "add":
                sheet = r.doc.Worksheets.Add(After=r.doc.Worksheets.Item(r.doc.Worksheets.Count))
            sheet.Name = target
        return self._result(r, {"action": action, "name": target}, save, backup)

    @staticmethod
    def _validate_font(font_size, color):
        if font_size is not None and not 1 <= font_size <= 409:
            raise OfficeError("font_size must be between 1 and 409 points.")
        if color is not None:
            rgb(color)

    @staticmethod
    def _font(font, name, size, bold, italic, color, ppt=False, ascii_name=None):
        if name is not None:
            font.Name = name
            if ascii_name is not None and not ppt:
                try:
                    font.NameFarEast = name
                except Exception:
                    pass
        if ascii_name is not None and not ppt:
            for attr in ("NameAscii", "NameOther"):
                try:
                    setattr(font, attr, ascii_name)
                except Exception:
                    pass
        if size is not None:
            font.Size = size
        if bold is not None:
            font.Bold = -1 if bold else 0
        if italic is not None:
            font.Italic = -1 if italic else 0
        if color is not None:
            if ppt:
                font.Color.RGB = rgb(color)
            else:
                font.Color = rgb(color)

    def excel_format(self, document, sheet, address, font_name=None, font_size=None, bold=None,
                     italic=None, color=None, fill=None, number_format=None, column_width=None, save=True):
        r = self.resolve(document, "excel")
        a1_size(address)
        self._validate_font(font_size, color)
        if fill is not None:
            rgb(fill)
        if column_width is not None and not 0 < column_width <= 255:
            raise OfficeError("column_width must be greater than 0 and at most 255.")
        cell = self._sheet(r, sheet).Range(address)
        backup = self._before_edit(r)
        with self._excel_events(r):
            self._font(cell.Font, font_name, font_size, bold, italic, color)
            if fill is not None:
                cell.Interior.Color = rgb(fill)
            if number_format is not None:
                cell.NumberFormat = number_format
            if column_width is not None:
                cell.ColumnWidth = column_width
        return self._result(r, {"sheet": str(cell.Worksheet.Name), "range": address}, save, backup)

    def _slide(self, record, index):
        if not 1 <= index <= int(record.doc.Slides.Count):
            raise OfficeError("slide is 1-based and must refer to an existing slide.")
        return record.doc.Slides.Item(index)

    def _shape(self, record, slide, shape):
        page = self._slide(record, slide)
        if not 1 <= shape <= int(page.Shapes.Count):
            raise OfficeError("shape is 1-based; inspect slides before editing.")
        obj = page.Shapes.Item(shape)
        if not obj.HasTextFrame:
            raise OfficeError("Selected shape has no editable text frame.")
        return obj

    def ppt_read(self, document, start=1, count=20):
        r = self.resolve(document, "ppt")
        if start < 1 or not 1 <= count <= 100:
            raise OfficeError("start is 1-based; count must be between 1 and 100.")
        slides = []
        total = int(r.doc.Slides.Count)
        for i in range(start, min(total + 1, start + count)):
            page = r.doc.Slides.Item(i)
            shapes = []
            for j in range(1, int(page.Shapes.Count) + 1):
                shape = page.Shapes.Item(j)
                has_text = bool(shape.HasTextFrame)
                shapes.append({"index": j, "id": int(shape.Id), "name": str(shape.Name),
                               "has_text_frame": has_text,
                               "text": str(shape.TextFrame.TextRange.Text)[:20000] if has_text else None})
            slides.append({"index": i, "id": int(page.SlideID), "shapes": shapes})
        return {**self._info(r), "total_slides": total, "slides": slides,
                "next_start": start + count if start + count <= total else None}

    def ppt_set_text(self, document, slide, shape, text, save=True):
        r = self.resolve(document, "ppt")
        if len(text) > 200000:
            raise OfficeError("text exceeds 200000 characters.")
        obj = self._shape(r, slide, shape)
        backup = self._before_edit(r)
        obj.TextFrame.TextRange.Text = text
        return self._result(r, {"slide": slide, "shape": shape, "characters": len(text)}, save, backup)

    def ppt_add_slide(self, document, title="", text="", position=None, save=True):
        r = self.resolve(document, "ppt")
        index = int(r.doc.Slides.Count) + 1 if position is None else position
        if not 1 <= index <= int(r.doc.Slides.Count) + 1:
            raise OfficeError("position must be a 1-based insertion position.")
        backup = self._before_edit(r)
        page = r.doc.Slides.Add(index, 12)  # ppLayoutBlank works in both engines.
        width, height = float(r.doc.PageSetup.SlideWidth), float(r.doc.PageSetup.SlideHeight)
        if title:
            shape = page.Shapes.AddTextbox(1, 36, 24, max(1, width - 72), 60)
            shape.TextFrame.TextRange.Text = title
            shape.TextFrame.TextRange.Font.Size = 28
        if text:
            shape = page.Shapes.AddTextbox(1, 36, 100, max(1, width - 72), max(1, height - 136))
            shape.TextFrame.TextRange.Text = text
            shape.TextFrame.TextRange.Font.Size = 18
        return self._result(r, {"slide": index, "slide_id": int(page.SlideID)}, save, backup)

    def ppt_add_textbox(self, document, slide, text, left=36, top=100, width=400, height=100, save=True):
        r = self.resolve(document, "ppt")
        if not all(math.isfinite(v) for v in (left, top, width, height)) or min(left, top) < 0 or min(width, height) <= 0:
            raise OfficeError("Coordinates must be finite points, with nonnegative left/top and positive width/height.")
        page = self._slide(r, slide)
        backup = self._before_edit(r)
        shape = page.Shapes.AddTextbox(1, left, top, width, height)
        shape.TextFrame.TextRange.Text = text
        return self._result(r, {"slide": slide, "shape": int(page.Shapes.Count), "shape_id": int(shape.Id)}, save, backup)

    def ppt_format(self, document, slide, shape, font_name=None, font_size=None, bold=None,
                   italic=None, color=None, save=True):
        r = self.resolve(document, "ppt")
        self._validate_font(font_size, color)
        obj = self._shape(r, slide, shape)
        backup = self._before_edit(r)
        self._font(obj.TextFrame.TextRange.Font, font_name, font_size, bold, italic, color, ppt=True)
        return self._result(r, {"slide": slide, "shape": shape}, save, backup)

    def ppt_add_box(self, document, text, left, top, width, height, fill=None, font_size=None,
                    font_color=None, rounded=True, slide=1, save=True):
        """Add a native box with text and an explicit fill.

        PowerPoint's default shape style gives white text, which is invisible on
        a light fill. When fill is supplied and font_color is omitted, the font
        color is chosen automatically for contrast, so users never get blank boxes.
        """
        r = self.resolve(document, "ppt")
        if not all(math.isfinite(v) for v in (left, top, width, height)) or min(left, top) < 0 or min(width, height) <= 0:
            raise OfficeError("Coordinates must be finite points, with nonnegative left/top and positive width/height.")
        if font_size is not None and (not math.isfinite(font_size) or font_size <= 0):
            raise OfficeError("font_size must be a positive number of points.")
        page = self._slide(r, slide)
        backup = self._before_edit(r)
        shape = page.Shapes.AddShape(5 if rounded else 1, float(left), float(top), float(width), float(height))
        if fill is not None:
            fill_value = rgb(fill)
            shape.Fill.Visible = -1
            shape.Fill.ForeColor.RGB = fill_value
            if font_color is None:
                font_color = "#262626" if self._is_light(fill_value) else "#FFFFFF"
        shape.TextFrame.TextRange.Text = text
        shape.TextFrame.TextRange.Font.Size = float(font_size) if font_size is not None else 10.0
        if font_color is not None:
            shape.TextFrame.TextRange.Font.Color.RGB = rgb(font_color)
        return self._result(r, {"slide": slide, "shape": int(page.Shapes.Count), "shape_id": int(shape.Id)},
                             save, backup)

    def ppt_add_connector(self, document, begin_x, begin_y, end_x, end_y, slide=1, color=None,
                          weight=None, arrow=True, begin_arrow=False, save=True):
        """Add a native straight connector.

        Arrowheads live on Shape.Line (LineFormat), not on Shape itself; setting
        Shape.EndArrowheadStyle fails. This convenience method sets them correctly.
        """
        r = self.resolve(document, "ppt")
        if not all(math.isfinite(v) for v in (begin_x, begin_y, end_x, end_y)):
            raise OfficeError("Connector coordinates must be finite points.")
        if weight is not None and (not math.isfinite(weight) or weight <= 0):
            raise OfficeError("weight must be a positive number of points.")
        page = self._slide(r, slide)
        backup = self._before_edit(r)
        shape = page.Shapes.AddConnector(1, float(begin_x), float(begin_y), float(end_x), float(end_y))
        line = shape.Line
        if color is not None:
            line.ForeColor.RGB = rgb(color)
        if weight is not None:
            line.Weight = float(weight)
        line.EndArrowheadStyle = 2 if arrow else 1
        line.BeginArrowheadStyle = 2 if begin_arrow else 1
        return self._result(r, {"slide": slide, "shape": int(page.Shapes.Count), "shape_id": int(shape.Id)},
                             save, backup)

    @staticmethod
    def _is_light(rgb_value):
        red = rgb_value & 0xFF
        green = (rgb_value >> 8) & 0xFF
        blue = (rgb_value >> 16) & 0xFF
        return (0.299 * red + 0.587 * green + 0.114 * blue) > 140


@contextmanager
def _no_context():
    yield
