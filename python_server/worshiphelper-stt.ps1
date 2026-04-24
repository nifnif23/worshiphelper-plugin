# ============================================================================
# worshiphelper-stt.ps1  --  v7
#
# Launches the WorshipHelper speech stack:
#   * STT WebSocket server  ws://127.0.0.1:8765/
#   * Embed HTTP server     http://127.0.0.1:8766/
#
# RTX 3050 recommended launch (runs in CUDA float16, ~2 GB VRAM total):
#   .\worshiphelper-stt.ps1
#
# Larger GPU / workstation (4+ GB VRAM):
#   .\worshiphelper-stt.ps1 -Model large-v3
#
# Force CPU (e.g. testing without GPU):
#   .\worshiphelper-stt.ps1 -Device cpu -Model small.en
#
# Skip semantic search embed server:
#   .\worshiphelper-stt.ps1 -NoEmbed
#
# First run: installs deps and downloads the model (~1.5 GB for distil-large-v3).
# Subsequent runs: starts in ~10-15 seconds (model already cached).
# ============================================================================
param(
    # distil-large-v3: ~1.5 GB VRAM float16, near large-v3 accuracy, ~medium speed.
    # Best default for RTX 3050 (4 GB). Change to large-v3 if you have 6+ GB VRAM.
    [string]$Model   = "distil-large-v3",
    [string]$Device  = "auto",
    [string]$Compute = "float16",
    [switch]$NoEmbed,
    [string]$EmbedDevice = "auto"
)

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

# ---- venv ------------------------------------------------------------------
if (-not (Test-Path .venv\Scripts\Activate.ps1)) {
    Write-Host "Creating virtual environment..." -ForegroundColor Cyan
    python -m venv .venv
}
. .\.venv\Scripts\Activate.ps1

# ---- CUDA torch (must come before requirements.txt) ------------------------
$torchVersion = pip show torch 2>$null | Select-String "^Version"
$hasCuda = python -c "import torch; print(torch.cuda.is_available())" 2>$null
if ($hasCuda -ne "True" -and $Device -ne "cpu") {
    Write-Host "Installing CUDA-enabled PyTorch (this may take a few minutes)..." -ForegroundColor Cyan
    pip install torch --index-url https://download.pytorch.org/whl/cu118 --quiet
}

# ---- deps ------------------------------------------------------------------
$installed = pip show faster-whisper 2>$null
if (-not $installed) {
    Write-Host "Installing Python dependencies..." -ForegroundColor Cyan
    pip install --upgrade pip --quiet
    pip install -r requirements.txt --quiet
}

# ---- start embed server ----------------------------------------------------
if (-not $NoEmbed) {
    Write-Host "Starting embed server (BAAI/bge-base-en-v1.5)..." -ForegroundColor DarkGray
    Start-Process -FilePath python `
        -ArgumentList "embed_server.py", "--device", $EmbedDevice `
        -WindowStyle Minimized `
        -WorkingDirectory $PSScriptRoot | Out-Null
    Write-Host "  http://127.0.0.1:8766/  (preloading model in background)" -ForegroundColor DarkGray
}

# ---- banner ----------------------------------------------------------------
Write-Host ""
Write-Host "WorshipHelper STT server v7" -ForegroundColor Green
Write-Host "  model   = $Model"
Write-Host "  device  = $Device"
Write-Host "  compute = $Compute"
Write-Host ""
Write-Host "First launch downloads the model (~1.5 GB). Subsequent starts are fast." -ForegroundColor DarkGray
Write-Host ""

# ---- start STT server (foreground — keep window open) ----------------------
python server.py --model $Model --device $Device --compute $Compute
