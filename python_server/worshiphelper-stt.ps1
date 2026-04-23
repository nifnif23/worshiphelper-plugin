# ============================================================================
# worshiphelper-stt.ps1
#
# One-command launcher for the WorshipHelper speech stack. Starts:
#   * STT WebSocket server on ws://127.0.0.1:8765/
#   * Embed HTTP server on http://127.0.0.1:8766/
#
# Usage (from the python_server folder):
#   .\worshiphelper-stt.ps1                  # auto-detect GPU/CPU, use small.en
#   .\worshiphelper-stt.ps1 -Model small.en  # force a specific model
#   .\worshiphelper-stt.ps1 -Device cpu      # force CPU even if GPU is available
#   .\worshiphelper-stt.ps1 -NoEmbed         # skip the semantic-search embed server
#
# Creates / uses the local .venv. First run downloads models (~500MB).
# ============================================================================
param(
    [string]$Model  = "small.en",
    [string]$Device = "auto",
    [string]$Compute = "auto",
    [switch]$NoEmbed,
    [switch]$Preload
)

$ErrorActionPreference = "Stop"

Set-Location -Path $PSScriptRoot

# -------------------------- venv ---------------------------------------------
if (-not (Test-Path .venv\Scripts\Activate.ps1)) {
    Write-Host "Creating .venv ..." -ForegroundColor Cyan
    python -m venv .venv
}
. .\.venv\Scripts\Activate.ps1

# -------------------------- deps ---------------------------------------------
$freeze = pip freeze 2>$null
if (-not ($freeze | Select-String -Quiet "^faster-whisper==")) {
    Write-Host "Installing dependencies (this can take a few minutes) ..." -ForegroundColor Cyan
    pip install --upgrade pip | Out-Null
    pip install -r requirements.txt
}

# -------------------------- start --------------------------------------------
Write-Host ""
Write-Host "Starting WorshipHelper STT server..." -ForegroundColor Green
Write-Host "  model  = $Model"
Write-Host "  device = $Device"
Write-Host "  compute= $Compute"
Write-Host ""

if (-not $NoEmbed) {
    $embedArgs = @()
    if ($Preload) { $embedArgs += "--preload" }
    Start-Process -FilePath python `
        -ArgumentList (@("embed_server.py") + $embedArgs) `
        -WindowStyle Minimized `
        -WorkingDirectory $PSScriptRoot | Out-Null
    Write-Host "Embed server starting (background) on http://127.0.0.1:8766/" -ForegroundColor DarkGray
}

python server.py --model $Model --device $Device --compute $Compute
