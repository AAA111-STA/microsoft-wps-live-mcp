"""Verify published schemas, structured object-model calls, and client profiles."""
import asyncio
import json
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from mcp import Client
from microsoft_wps_live_mcp.clients import PROFILES, client_config, write_configs
from microsoft_wps_live_mcp.server import build_server


class StubWorker:
    async def call(self, method, *args, **kwargs):
        return {"method": method, "args": list(args)}


class ProtocolTests(unittest.TestCase):
    def test_native_object_model_request_round_trips_through_mcp(self):
        async def check():
            async with Client(build_server(StubWorker())) as client:
                tools = await client.list_tools()
                names = {t.name for t in tools.tools}
                self.assertIn("office_create_copy", names)
                for tool in tools.tools:
                    self.assertNotIn("save", tool.input_schema.get("properties", {}), tool.name)
                steps = [{"action": "call", "target": "document.Slides", "member": "Add", "args": [1, 12], "as": "slide"}]
                result = await client.call_tool("office_api_batch", {"document": "sample", "steps": steps})
                self.assertFalse(result.is_error)
                payload = result.structured_content or json.loads(result.content[0].text)
                self.assertEqual(payload["method"], "api_batch")
                self.assertEqual(payload["args"], ["sample", steps, True])
        asyncio.run(check())

    def test_config_profiles_emit_the_correct_container_formats(self):
        for profile in PROFILES:
            text = client_config(profile)
            if profile == "codex":
                import tomllib
                self.assertEqual(tomllib.loads(text)["mcp_servers"]["microsoft-wps-live"]["command"], sys.executable)
            elif profile == "deepseek-harness":
                self.assertIn("@deepseek-ai/dsh-mcp-client", text)
                self.assertIn("transport: stdio", text)
            elif profile == "opencode":
                entry = json.loads(text)["mcp"]["microsoft-wps-live"]
                self.assertEqual(entry["type"], "local")
                self.assertEqual(entry["command"][0], sys.executable)
            else:
                self.assertEqual(json.loads(text)["mcpServers"]["microsoft-wps-live"]["command"], sys.executable)
        with tempfile.TemporaryDirectory() as directory:
            files = write_configs(directory)
            self.assertEqual(len(files), len(PROFILES))
            self.assertTrue(all(Path(path).is_file() for path in files))


if __name__ == "__main__":
    unittest.main()
