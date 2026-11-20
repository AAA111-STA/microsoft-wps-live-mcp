"""Run from a source checkout without relying on the shell working directory."""
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from microsoft_wps_live_mcp.server import main

if __name__ == "__main__":
    main()
