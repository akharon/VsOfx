# Runs npm (or node) for this project, even when Node.js is not installed on the machine.
#
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts/npm.ps1 run package:clean
#   powershell -NoProfile -ExecutionPolicy Bypass -File scripts/npm.ps1 -Node esbuild.mjs --dev
#
# 1. Uses node/npm from PATH if a version >= 22.18 is installed.
# 2. Otherwise uses a portable Node.js in .tools/ (git-ignored), downloading it once from
#    nodejs.org and checking it against the official SHA-256 list before it is used.
# Works in Windows PowerShell 5.1 and PowerShell 7.

$ErrorActionPreference = 'Stop'
$NodeVersion = 'v24.22.3'   # the portable fallback; package.json requires >= 22.18

$root = Split-Path -Parent $PSScriptRoot
$useNode = $false
$rest = @($args)
if ($rest.Count -gt 0 -and $rest[0] -eq '-Node') {
  $useNode = $true
  # Not "$rest = if (...) {...}": that unrolls a one-element array into a string, which @rest then
  # splats character by character.
  $rest = @($rest | Select-Object -Skip 1)
}

function Test-NodeOk([string]$exe) {
  try {
    $v = (& $exe --version) 2>$null
    if ($v -match '^v(\d+)\.(\d+)') { return ([int]$Matches[1] -gt 22) -or ([int]$Matches[1] -eq 22 -and [int]$Matches[2] -ge 18) }
  } catch { }
  return $false
}

$nodeDir = $null
$onPath = Get-Command node -ErrorAction SilentlyContinue
if ($onPath -and (Test-NodeOk $onPath.Source)) {
  $nodeDir = Split-Path -Parent $onPath.Source
} else {
  $arch = if ($env:PROCESSOR_ARCHITECTURE -eq 'ARM64') { 'arm64' } else { 'x64' }
  $name = "node-$NodeVersion-win-$arch"
  $tools = Join-Path $root '.tools'
  $nodeDir = Join-Path $tools $name
  if (-not (Test-Path (Join-Path $nodeDir 'node.exe')) -or -not (Test-Path (Join-Path $nodeDir 'node_modules\npm\lib\commands\ci.js'))) {
    Write-Host "Node.js >= 22.18 not found on PATH: setting up portable $name in .tools/ (one time) ..."
    New-Item -ItemType Directory -Force $tools | Out-Null
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $base = "https://nodejs.org/dist/$NodeVersion"
    $zip = Join-Path $tools "$name.zip"
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest "$base/$name.zip" -OutFile $zip -UseBasicParsing
    $sums = (Invoke-WebRequest "$base/SHASUMS256.txt" -UseBasicParsing).Content
    $expected = ($sums -split "`n" | Where-Object { $_ -match "\s$name\.zip$" } | ForEach-Object { ($_ -split '\s+')[0] })
    $actual = (Get-FileHash $zip -Algorithm SHA256).Hash.ToLower()
    if (-not $expected -or $actual -ne $expected.ToLower()) {
      Remove-Item $zip -Force
      throw "Checksum mismatch for $name.zip (expected $expected, got $actual). Download aborted."
    }
    if (Test-Path $nodeDir) { Remove-Item -Recurse -Force $nodeDir }
    Expand-Archive $zip -DestinationPath $tools -Force
    Remove-Item $zip -Force
    Write-Host "Portable Node.js ready: $nodeDir"
  }
}

# npm's child processes (scripts, npx, vsce) must find the same node
$env:PATH = "$nodeDir;$env:PATH"
Set-Location $root
$npm = Join-Path $nodeDir 'npm.cmd'

# Self-healing: dependencies missing or half-removed (e.g. an interrupted install) -> install them first.
$isInstall = (-not $useNode) -and $rest.Count -gt 0 -and @('install', 'i', 'ci') -contains $rest[0]
if (-not $isInstall -and -not (Test-Path (Join-Path $root 'node_modules\esbuild\package.json'))) {
  Write-Host 'Dependencies are missing or incomplete: running npm install ...'
  & $npm install --ignore-scripts --no-audit --no-fund
  if ($LASTEXITCODE -ne 0) {
    Write-Host ''
    Write-Host 'npm install failed. If it says ENOTEMPTY or EPERM, another program holds files in node_modules open:'
    Write-Host 'close other terminals running npm/node (or a running build/watch task) and run this again.'
    exit $LASTEXITCODE
  }
}

if ($useNode) {
  & (Join-Path $nodeDir 'node.exe') @rest
} else {
  & $npm @rest
}
exit $LASTEXITCODE
