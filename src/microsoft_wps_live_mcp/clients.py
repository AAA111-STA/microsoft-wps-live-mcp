"""Client configuration generators. No author-specific paths or client setting mutations."""
import json
import sys

PROFILES = ("codex", "claude", "claude-code", "cursor", "gemini", "antigravity",
            "opencode", "kimi", "zcode", "workbuddy", "deepseek-harness", "trae", "qoder", "json")


def client_config(client):
    if client not in PROFILES:
        raise ValueError("Unknown client profile: " + client)
    command = sys.executable
    args = ["-m", "microsoft_wps_live_mcp"]
    env = {"PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}
    if client == "codex":
        return "[mcp_servers.microsoft-wps-live]\ncommand = " + json.dumps(command) + "\nargs = " + json.dumps(args) + \
               "\nstartup_timeout_sec = 30\ntool_timeout_sec = 120\n\n" + \
               "[mcp_servers.microsoft-wps-live.env]\nPYTHONUTF8 = \"1\"\nPYTHONIOENCODING = \"utf-8\"\n"
    if client == "opencode":
        return json.dumps({
            "mcp": {
                "microsoft-wps-live": {
                    "type": "local",
                    "command": [command, *args],
                    "environment": env,
                    "enabled": True,
                    "timeout": 120000,
                }
            }
        }, indent=2)
    if client == "deepseek-harness":
        # Harness uses Cordis plugin entries rather than a mcpServers JSON map.
        return "- id: microsoft-wps-live\n  name: '@deepseek-ai/dsh-mcp-client'\n  config:\n" + \
               "    serverName: microsoft-wps-live\n    transport: stdio\n    command: " + json.dumps(command) + \
               "\n    args: " + json.dumps(args) + "\n    env: " + json.dumps(env) + "\n"
    entry = {"command": command, "args": args, "env": env}
    if client in ("claude-code", "zcode", "workbuddy", "qoder"):
        entry["type"] = "stdio"
    if client in ("gemini", "workbuddy"):
        entry["timeout"] = 120000
    return json.dumps({"mcpServers": {"microsoft-wps-live": entry}}, indent=2)


def write_configs(directory):
    from pathlib import Path
    root = Path(directory).expanduser().resolve()
    root.mkdir(parents=True, exist_ok=True)
    files = []
    for profile in PROFILES:
        extension = ".toml" if profile == "codex" else ".yml" if profile == "deepseek-harness" else ".json"
        destination = root / (profile + extension)
        destination.write_text(client_config(profile) + "\n", encoding="utf-8")
        files.append(str(destination))
    return files
