$ErrorActionPreference = "Stop"

$workspace = Split-Path -Parent $MyInvocation.MyCommand.Path
$runtimeNode = "C:\Users\ADMIN\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe"
$runtimeNodeModules = "C:\Users\ADMIN\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules"
$workspaceNodeModules = Join-Path $workspace "node_modules"
$exampleConfig = Join-Path $workspace "config\login.config.example.json"
$activeConfig = Join-Path $workspace "config\login.config.json"

if (-not (Test-Path $runtimeNode)) {
  throw "Bundled Node.js runtime not found at $runtimeNode"
}

if (-not (Test-Path $workspaceNodeModules)) {
  New-Item -ItemType Junction -Path $workspaceNodeModules -Target $runtimeNodeModules | Out-Null
}

if (-not (Test-Path $activeConfig) -and (Test-Path $exampleConfig)) {
  Copy-Item -LiteralPath $exampleConfig -Destination $activeConfig
}

Push-Location $workspace
try {
  & $runtimeNode ".\scripts\auto-test-login.mjs" @args
} finally {
  Pop-Location
}
