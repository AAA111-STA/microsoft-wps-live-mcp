param([string]$Python = 'python')
$ErrorActionPreference = 'Stop'
$bundleRoot = Split-Path -Parent $PSScriptRoot
$distroName = 'OfficeLiveEditor'
$runtimePath = Join-Path $bundleRoot 'wsl'
$downloadPath = Join-Path $bundleRoot 'downloads'
$configurationPath = Join-Path $bundleRoot 'config'

if ($bundleRoot -notmatch '^[A-Za-z]:\\') { throw 'Use a local Windows drive for this bundle.' }
if (-not (Get-Command $Python -ErrorAction SilentlyContinue)) { throw 'Python 3.11+ is required.' }
& wsl --status
if ($LASTEXITCODE -ne 0) { throw 'Install or enable WSL 2 before running this setup.' }

New-Item -ItemType Directory -Path $downloadPath,$configurationPath -Force | Out-Null
$ownerSid = [System.Security.Principal.WindowsIdentity]::GetCurrent().User.Value
& icacls.exe $configurationPath /inheritance:r /grant:r ('*' + $ownerSid + ':(OI)(CI)F') '*S-1-5-18:(OI)(CI)F' | Out-Null
if ($LASTEXITCODE -ne 0) { throw 'Could not protect private configuration.' }

$registered = @(Get-ChildItem -LiteralPath 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Lxss' -ErrorAction SilentlyContinue |
    ForEach-Object { Get-ItemProperty -LiteralPath $_.PSPath } |
    Where-Object { $_.DistributionName -eq $distroName })
if ($registered.Count -gt 0) {
    $registeredRoot = [IO.Path]::GetFullPath($registered[0].BasePath).TrimEnd('\')
    $expectedRoot = [IO.Path]::GetFullPath($runtimePath).TrimEnd('\')
    if ($registeredRoot -ne $expectedRoot) { throw "WSL name $distroName already belongs to a different location. No changes made to it." }
} else {
    # Pinned Canonical image; SHA256 verified against the official release list.
    $imageName = 'ubuntu-24.04.5-wsl-amd64.wsl'
    $imagePath = Join-Path $downloadPath $imageName
    $expectedHash = 'bb415d824822c4b878125729af451a5d18fb13d1cf5cbed9a7393ad64ac6039e'
    if (-not (Test-Path -LiteralPath $imagePath)) {
        Invoke-WebRequest -Uri ('https://releases.ubuntu.com/noble/' + $imageName) -OutFile $imagePath
    }
    if ((Get-FileHash -LiteralPath $imagePath -Algorithm SHA256).Hash.ToLowerInvariant() -ne $expectedHash) {
        throw 'Ubuntu image checksum mismatch. Import cancelled.'
    }
    & wsl --import $distroName $runtimePath $imagePath --version 2
    if ($LASTEXITCODE -ne 0) { throw 'Isolated WSL import failed.' }
}

$linuxBundle = '/mnt/' + $bundleRoot.Substring(0,1).ToLowerInvariant() + '/' + $bundleRoot.Substring(3).Replace('\','/')
& wsl -d $distroName -u root -- bash ($linuxBundle + '/scripts/setup-linux.sh')
if ($LASTEXITCODE -ne 0) { throw 'Linux dependencies were not installed successfully.' }
& $Python (Join-Path $PSScriptRoot 'office_live.py') init
if ($LASTEXITCODE -ne 0) { throw 'Local configuration failed.' }
& $Python (Join-Path $PSScriptRoot 'office_live.py') start
if ($LASTEXITCODE -ne 0) { throw 'Document services did not start.' }
