<#
  setup_env.ps1 – авто-установка Python + venv + зависимости
  Запуск: ПКМ → “Run with PowerShell” ИЛИ
          pwsh -File scripts\setup_env.ps1
#>

param(
  [string]$PY_VER  = "3.12.3",
  [string]$VENV    = "venv"
)

$ErrorActionPreference = "Stop"
$BASE = Split-Path -Parent $MyInvocation.MyCommand.Definition

# ── 1. Python ────────────────────────────────────────────────────────────────
if (-not (Get-Command python -ErrorAction SilentlyContinue)) {
    Write-Host "→ Python не найден. Скачиваем $PY_VER…"
    $url = "https://www.python.org/ftp/python/$PY_VER/python-$PY_VER-amd64.exe"
    $tmp = "$env:TEMP\python-installer.exe"
    Invoke-WebRequest $url -OutFile $tmp
    Start-Process -Wait $tmp "/quiet InstallAllUsers=1 PrependPath=1"
}

# ── 2. venv ──────────────────────────────────────────────────────────────────
if (-not (Test-Path "$BASE\$VENV")) {
    python -m venv $VENV
}
& "$BASE\$VENV\Scripts\Activate.ps1"
python -m pip install --upgrade pip
pip install -r "$BASE\requirements.txt"

Write-Host "✓ Venv готов: $BASE\$VENV"

