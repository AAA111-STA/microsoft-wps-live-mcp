"""macOS desktop Office and WPS Office automation via AppleScript and desktop application bridge."""
from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time
from typing import Any
import xml.etree.ElementTree as ET
import zipfile

from .discovery import MACOS_APPS


def run_osascript(script: str, timeout: float = 10.0) -> tuple[int, str, str]:
    """Execute AppleScript via osascript."""
    if sys.platform != "darwin":
        return 1, "", "Not running on macOS"
    try:
        proc = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True,
            text=True,
            timeout=timeout,
        )
        return proc.returncode, proc.stdout.strip(), proc.stderr.strip()
    except Exception as exc:
        return 1, "", str(exc)


def is_process_running(process_name: str) -> bool:
    """Check if process is running on macOS."""
    if sys.platform != "darwin":
        return False
    code, out, _ = run_osascript(f'tell application "System Events" to exists (process "{process_name}")')
    return code == 0 and out.lower() == "true"


def get_app_name(vendor: str, kind: str) -> str:
    """Return the macOS AppleScript application name."""
    if vendor == "microsoft":
        names = {"word": "Microsoft Word", "excel": "Microsoft Excel", "ppt": "Microsoft PowerPoint"}
        return names.get(kind, "Microsoft Word")
    return "WPS Office"


def open_in_app(app_name: str, path: Path, visible: bool = True) -> bool:
    """Open document in macOS application."""
    if sys.platform != "darwin":
        return False
    path_str = str(path.resolve())
    if "Microsoft" in app_name:
        cmd = f'tell application "{app_name}"\n'
        if visible:
            cmd += '    activate\n'
        cmd += f'    open POSIX file "{path_str}"\nend tell'
        code, _, _ = run_osascript(cmd)
        if code == 0:
            return True
    # Fallback to /usr/bin/open
    open_args = ["open", "-a", app_name, path_str]
    if not visible:
        open_args.append("-g")
    res = subprocess.run(open_args, capture_output=True, text=True)
    return res.returncode == 0


def activate_app(app_name: str) -> bool:
    """Bring macOS application window to front."""
    if sys.platform != "darwin":
        return False
    code, _, _ = run_osascript(f'tell application "{app_name}" to activate')
    return code == 0


def save_in_app(app_name: str, path: Path) -> bool:
    """Request active document save via AppleScript."""
    if sys.platform != "darwin":
        return False
    if "Microsoft" in app_name:
        code, _, _ = run_osascript(f'tell application "{app_name}" to save active document')
        return code == 0
    return False


def close_in_app(app_name: str, path: Path) -> bool:
    """Close document via AppleScript."""
    if sys.platform != "darwin":
        return False
    if "Microsoft" in app_name:
        code, _, _ = run_osascript(f'tell application "{app_name}" to close (first document whose full name is "{path}")')
        return code == 0
    return False


# Minimal standard OOXML templates for blank new documents
def create_blank_ooxml(path: Path, kind: str):
    """Create a completely valid, standard blank OOXML file (.docx, .xlsx, or .pptx)."""
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as z:
        if kind == "word":
            z.writestr("[Content_Types].xml", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                '</Types>'
            ))
            z.writestr("_rels/.rels", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                '</Relationships>'
            ))
            z.writestr("word/document.xml", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:body><w:p><w:r><w:t></w:t></w:r></w:p></w:body></w:document>'
            ))
        elif kind == "excel":
            z.writestr("[Content_Types].xml", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                '</Types>'
            ))
            z.writestr("_rels/.rels", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
                '</Relationships>'
            ))
            z.writestr("xl/workbook.xml", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
                'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
                '<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
            ))
            z.writestr("xl/_rels/workbook.xml.rels", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
                '</Relationships>'
            ))
            z.writestr("xl/worksheets/sheet1.xml", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
                '<sheetData></sheetData></worksheet>'
            ))
        elif kind == "ppt":
            z.writestr("[Content_Types].xml", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>'
                '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
                '</Types>'
            ))
            z.writestr("_rels/.rels", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
                '</Relationships>'
            ))
            z.writestr("ppt/presentation.xml", (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
                '<p:sldIdLst></p:sldIdLst></p:presentation>'
            ))
