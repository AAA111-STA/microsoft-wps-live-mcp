param([string]$Python = 'python')
$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$environmentRoot = Join-Path $projectRoot '.venv'
$environmentPython = Join-Path $environmentRoot 'Scripts\python.exe'

& $Python -c "import sys; assert sys.version_info >= (3, 11), 'Python 3.11+ is required'"
if ($LASTEXITCODE -ne 0) { throw 'Install Python 3.11+ and run this script again.' }
if (-not (Test-Path -LiteralPath $environmentPython)) {
    & $Python -m venv $environmentRoot
    if ($LASTEXITCODE -ne 0) { throw 'Failed to create the local Python environment.' }
}
& $environmentPython -m pip install $projectRoot
if ($LASTEXITCODE -ne 0) { throw 'Failed to install MCP dependencies.' }

$configurationRoot = Join-Path $projectRoot 'config'
New-Item -ItemType Directory -Path $configurationRoot -Force | Out-Null
$codexConfig = & $environmentPython -X utf8 -m microsoft_wps_live_mcp --config codex
if ($LASTEXITCODE -ne 0) { throw 'Could not generate Codex configuration.' }
$jsonConfig = & $environmentPython -X utf8 -m microsoft_wps_live_mcp --config json
if ($LASTEXITCODE -ne 0) { throw 'Could not generate MCP JSON configuration.' }
$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText((Join-Path $configurationRoot 'mcp-codex.toml'), ($codexConfig -join "`n") + "`n", $utf8)
[System.IO.File]::WriteAllText((Join-Path $configurationRoot 'mcp-client.json'), ($jsonConfig -join "`n") + "`n", $utf8)
& $environmentPython -X utf8 -m microsoft_wps_live_mcp --write-configs (Join-Path $configurationRoot 'clients')
if ($LASTEXITCODE -ne 0) { throw 'Could not generate client profiles.' }
Write-Host 'Installed microsoft-wps-live mcp.'
Write-Host "Codex configuration: $(Join-Path $configurationRoot 'mcp-codex.toml')"
Write-Host "Other MCP clients:   $(Join-Path $configurationRoot 'mcp-client.json')"
Write-Host 'Add the generated entry to your client MCP configuration and start a new client session.'
Write-Host 'Compatibility check:'
Write-Host "& '$environmentPython' -m microsoft_wps_live_mcp --doctor"
