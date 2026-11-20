#!/usr/bin/env bash
# macOS / Linux setup for microsoft-wps-live-mcp
set -euo pipefail
cd "$(dirname "$0")/.."

PYTHON="${1:-python3}"
"$PYTHON" -c "import sys; assert sys.version_info >= (3, 11), 'Python 3.11+ required'"

if [ ! -f .venv/bin/python ]; then
    "$PYTHON" -m venv .venv
fi

.venv/bin/python -m pip install -q .
.venv/bin/python -m microsoft_wps_live_mcp --version

mkdir -p config/clients
.venv/bin/python -m microsoft_wps_live_mcp --write-configs config/clients

echo "Installed microsoft-wps-live mcp."
echo "Client configs: config/clients/"
echo "Compatibility check:"
echo "  .venv/bin/python -m microsoft_wps_live_mcp --doctor"
