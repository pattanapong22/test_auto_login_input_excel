$ErrorActionPreference = "Stop"

$workspace = Split-Path -Parent $MyInvocation.MyCommand.Path
$runtimeNodeModules = "C:\Users\ADMIN\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\node_modules"
$workspaceNodeModules = Join-Path $workspace "node_modules"
$exampleConfig = Join-Path $workspace "config\login.config.example.json"
$activeConfig = Join-Path $workspace "config\login.config.json"

if (-not (Test-Path $workspaceNodeModules)) {
  New-Item -ItemType Junction -Path $workspaceNodeModules -Target $runtimeNodeModules | Out-Null
  Write-Host "Linked node_modules -> bundled runtime"
} else {
  Write-Host "node_modules already ready"
}

if (-not (Test-Path $activeConfig) -and (Test-Path $exampleConfig)) {
  Copy-Item -LiteralPath $exampleConfig -Destination $activeConfig
  Write-Host "Created config/login.config.json from example"
} elseif (Test-Path $activeConfig) {
  Write-Host "Config already exists"
}

Write-Host "Setup complete"
