"""Platform discovery. Importing this module does not start any Office application."""
from __future__ import annotations

import os
import shutil
import subprocess
import sys
from pathlib import Path
import uuid

PROGIDS = {
    "microsoft": {
        "word": ("Word.Application",),
        "excel": ("Excel.Application",),
        "ppt": ("PowerPoint.Application",),
    },
    "wps": {
        "word": ("KWPS.Application", "WPS.Application"),
        "excel": ("KET.Application", "ET.Application"),
        "ppt": ("KWPP.Application", "WPP.Application"),
    },
}
COLLECTIONS = {"word": "Documents", "excel": "Workbooks", "ppt": "Presentations"}
EXTENSIONS = {".docx": "word", ".xlsx": "excel", ".pptx": "ppt"}

MACOS_APPS = {
    "microsoft": {
        "word": ("/Applications/Microsoft Word.app", "~/Applications/Microsoft Word.app", "com.microsoft.Word"),
        "excel": ("/Applications/Microsoft Excel.app", "~/Applications/Microsoft Excel.app", "com.microsoft.Excel"),
        "ppt": ("/Applications/Microsoft PowerPoint.app", "~/Applications/Microsoft PowerPoint.app", "com.microsoft.Powerpoint"),
    },
    "wps": {
        "word": ("/Applications/WPS Office.app", "/Applications/wpsoffice.app", "~/Applications/WPS Office.app", "cn.wps.office", "com.kingsoft.wpsoffice"),
        "excel": ("/Applications/WPS Office.app", "/Applications/wpsoffice.app", "~/Applications/WPS Office.app", "cn.wps.office", "com.kingsoft.wpsoffice"),
        "ppt": ("/Applications/WPS Office.app", "/Applications/wpsoffice.app", "~/Applications/WPS Office.app", "cn.wps.office", "com.kingsoft.wpsoffice"),
    },
}


def canonical(path: str | Path) -> str:
    return os.path.normcase(os.path.realpath(os.path.abspath(path)))


def registration(progid: str) -> dict | None:
    if sys.platform != "win32":
        return None
    import winreg

    # WPS can be registered per user and without a LocalServer32 subkey.
    for hive, prefix in ((winreg.HKEY_CURRENT_USER, "Software\\Classes\\"),
                         (winreg.HKEY_LOCAL_MACHINE, "Software\\Classes\\")):
        for view in (winreg.KEY_WOW64_64KEY, winreg.KEY_WOW64_32KEY):
            try:
                with winreg.OpenKey(hive, prefix + progid + "\\CLSID", 0,
                                    winreg.KEY_READ | view) as key:
                    clsid = winreg.QueryValueEx(key, None)[0]
                return {"progid": progid, "clsid": clsid}
            except OSError:
                continue
    return None


def macos_app_installed(vendor: str, kind: str) -> bool:
    if sys.platform != "darwin":
        return False
    targets = MACOS_APPS.get(vendor, {}).get(kind, ())
    for target in targets:
        if target.endswith(".app"):
            if Path(target).expanduser().is_dir():
                return True
    if shutil.which("mdfind"):
        for target in targets:
            if not target.endswith(".app"):
                try:
                    res = subprocess.run(
                        ["mdfind", f"kMDItemCFBundleIdentifier == '{target}'"],
                        capture_output=True, text=True, timeout=2
                    )
                    if res.returncode == 0 and res.stdout.strip():
                        return True
                except Exception:
                    pass
    return False


def capabilities() -> dict:
    is_win = sys.platform == "win32"
    is_mac = sys.platform == "darwin"

    if is_mac:
        engines = {
            vendor: {
                kind: {
                    "registered": macos_app_installed(vendor, kind),
                    "targets": list(targets),
                }
                for kind, targets in kinds.items()
            }
            for vendor, kinds in MACOS_APPS.items()
        }
        note = (
            "Native live editing supported on macOS via AppleScript and desktop application bridge "
            "for Microsoft Office and WPS Office."
        )
    else:
        engines = {
            vendor: {
                kind: {
                    "registered": any(registration(p) for p in ids),
                    "progids": list(ids),
                }
                for kind, ids in kinds.items()
            }
            for vendor, kinds in PROGIDS.items()
        }
        note = (
            "Registration is not a successful connection test. WPS builds must expose COM automation on Windows. "
            "Native live editing requires a logged-in Windows or macOS desktop session; Linux is unsupported."
        )

    return {
        "platform": sys.platform,
        "native_supported": is_win or is_mac,
        "engines": engines,
        "formats": list(EXTENSIONS),
        "note": note,
    }


def vendor_of(app, fallback: str | None = None) -> str:
    # WPS deliberately returns 'Microsoft Word/Excel/PowerPoint' for Application.Name.
    # Executable location is a better discriminator than the advertised product name.
    try:
        location = str(app.Path).lower()
        if any(part in location for part in ("kingsoft", "wps office", "office6")):
            return "wps"
        if any(part in location for part in ("microsoft office", "office16", "office15")):
            return "microsoft"
    except Exception:
        pass
    return fallback or "unknown"


def native_window_apps(kind: str) -> list:
    """Find additional Microsoft Office instances, including a second Excel process.

    OBJID_NATIVEOM returns the existing native object model, never starts an app.
    COM objects are acquired and used only on the caller's STA thread.
    """
    import ctypes
    import pythoncom
    import win32gui
    from .com_dispatch import dispatch

    classes = {"word": "_WwG", "excel": "EXCEL7", "ppt": "paneClassDC"}
    handles = []

    def child(hwnd, _):
        if win32gui.GetClassName(hwnd) == classes[kind]:
            handles.append(hwnd)
        return True

    def top(hwnd, _):
        win32gui.EnumChildWindows(hwnd, child, None)
        return True

    win32gui.EnumWindows(top, None)
    oleacc = ctypes.OleDLL("oleacc")
    accessible = oleacc.AccessibleObjectFromWindow
    accessible.argtypes = [ctypes.c_void_p, ctypes.c_uint32, ctypes.c_void_p,
                           ctypes.POINTER(ctypes.c_void_p)]
    accessible.restype = ctypes.c_long
    iid = (ctypes.c_ubyte * 16).from_buffer_copy(uuid.UUID(str(pythoncom.IID_IDispatch)).bytes_le)
    result = []
    for hwnd in handles:
        pointer = ctypes.c_void_p()
        try:
            hr = accessible(hwnd, 0xFFFFFFF0, ctypes.byref(iid), ctypes.byref(pointer))
            if hr == 0 and pointer.value:
                try:
                    dispatch = pythoncom.ObjectFromAddress(pointer.value, pythoncom.IID_IDispatch)
                    from .com_dispatch import dispatch as wrap
                    window = wrap(dispatch)
                    result.append(window.Application)
                finally:
                    # ObjectFromAddress takes its own reference; release the reference from oleacc.
                    vtable = ctypes.cast(pointer, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
                    release = ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)(vtable[2])
                    release(pointer)
        except Exception:
            continue
    return result

