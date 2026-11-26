"""Idempotent helper to register microsoft-wps-live MCP and skills across detected local AI agents."""
from __future__ import annotations

import argparse
import datetime as dt
import json
from pathlib import Path
import re
import shutil
import sys
import tomllib

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from microsoft_wps_live_mcp.clients import client_config

ROOT = Path(__file__).resolve().parents[1]
PPT_SKILL_SRC = ROOT / "skills" / "ppt-architecture-diagram"
ROOT_SKILL_SRC = ROOT / "SKILL.md"


def _backup(path: Path, stamp: str) -> Path | None:
    if not path.exists():
        return None
    bak = path.with_name(f"{path.name}.before-microsoft-wps-live-{stamp}.bak")
    shutil.copy2(path, bak)
    return bak


def _sync_json_client(path: Path, profile: str, stamp: str, *, key: str = "mcpServers", apply: bool = False) -> str:
    payload = json.loads(client_config(profile))
    entry = payload[key]["microsoft-wps-live"]
    existing = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    section = existing.get(key, {})
    if section.get("microsoft-wps-live") == entry:
        return f"unchanged ({path})"
    if not apply:
        return f"would update ({path})"
    _backup(path, stamp)
    existing.setdefault(key, {})["microsoft-wps-live"] = entry
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(existing, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return f"updated ({path})"


def _sync_codex(path: Path, stamp: str, *, apply: bool = False) -> str:
    block = client_config("codex").strip()
    text = path.read_text(encoding="utf-8") if path.exists() else ""
    parsed = tomllib.loads(text) if text.strip() else {}
    want = tomllib.loads(block)["mcp_servers"]["microsoft-wps-live"]
    if parsed.get("mcp_servers", {}).get("microsoft-wps-live") == want:
        return f"unchanged ({path})"
    if not apply:
        return f"would update ({path})"
    _backup(path, stamp)
    # Remove any prior microsoft-wps-live table and its .env subtable cleanly before appending
    cleaned = re.sub(
        r"(?m)^\[mcp_servers\.(?:\"microsoft-wps-live\"|microsoft-wps-live)(?:\.env)?\][^\n]*(?:\n(?!\[)[^\n]*)*",
        "",
        text,
    ).rstrip()
    new_text = (cleaned + "\n\n" if cleaned else "") + block + "\n"
    tomllib.loads(new_text)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(new_text, encoding="utf-8")
    return f"updated ({path})"


def _copy_if_changed(src_dir: Path, dst_dir: Path, *, apply: bool) -> str:
    changed = False
    for src_file in src_dir.rglob("*"):
        if not src_file.is_file():
            continue
        rel = src_file.relative_to(src_dir)
        dst_file = dst_dir / rel
        if not dst_file.exists() or dst_file.read_bytes() != src_file.read_bytes():
            changed = True
            if apply:
                dst_file.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(src_file, dst_file)
    return ("updated" if apply else "would update") if changed else "unchanged"


def sync_all(apply: bool = False) -> dict[str, dict[str, str]]:
    home = Path.home()
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    mcp_results: dict[str, str] = {}
    skill_results: dict[str, str] = {}

    json_clients = [
        ("antigravity", home / ".gemini" / "config" / "mcp_config.json", home / ".gemini" / "config", "mcpServers"),
        ("gemini", home / ".gemini" / "settings.json", home / ".gemini", "mcpServers"),
        ("claude-code", home / ".claude.json", home / ".claude.json", "mcpServers"),
        ("claude", home / "AppData" / "Roaming" / "Claude" / "claude_desktop_config.json",
         home / "AppData" / "Roaming" / "Claude", "mcpServers"),
        ("cursor", home / ".cursor" / "mcp.json", home / ".cursor", "mcpServers"),
        ("opencode", home / ".config" / "opencode" / "opencode.json", home / ".config" / "opencode", "mcp"),
    ]
    for profile, cfg_path, probe_path, container_key in json_clients:
        if probe_path.exists() or cfg_path.exists():
            mcp_results[profile] = _sync_json_client(cfg_path, profile, stamp, key=container_key, apply=apply)

    codex_dir = home / ".codex"
    if codex_dir.exists():
        mcp_results["codex"] = _sync_codex(codex_dir / "config.toml", stamp, apply=apply)

    skill_roots = [
        ("antigravity", home / ".gemini" / "config" / "skills", home / ".gemini" / "config"),
        ("gemini", home / ".gemini" / "skills", home / ".gemini"),
        ("claude-code", home / ".claude" / "skills", home / ".claude"),
        ("codex", home / ".codex" / "skills", home / ".codex"),
        ("cursor", home / ".cursor" / "skills", home / ".cursor"),
        ("opencode", home / ".config" / "opencode" / "skills", home / ".config" / "opencode"),
    ]
    for name, sroot, gate in skill_roots:
        if not gate.exists():
            continue
        r1 = _copy_if_changed(PPT_SKILL_SRC, sroot / "ppt-architecture-diagram", apply=apply)
        dst_root_skill = sroot / "microsoft-wps-live" / "SKILL.md"
        root_changed = not dst_root_skill.exists() or dst_root_skill.read_bytes() != ROOT_SKILL_SRC.read_bytes()
        if root_changed and apply:
            dst_root_skill.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(ROOT_SKILL_SRC, dst_root_skill)
        r2 = ("updated" if apply else "would update") if root_changed else "unchanged"
        skill_results[name] = f"ppt-architecture-diagram:{r1}, microsoft-wps-live:{r2} ({sroot})"

    return {"mcp": mcp_results, "skills": skill_results}


def main():
    parser = argparse.ArgumentParser(description="Sync microsoft-wps-live MCP & skills to installed local agent clients.")
    parser.add_argument("--apply", action="store_true", help="Write changes (creating .bak backups). Without this, dry-run only.")
    args = parser.parse_args()
    res = sync_all(apply=args.apply)
    print(json.dumps(res, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
