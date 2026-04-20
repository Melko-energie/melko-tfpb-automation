# melko-tfpb-automation — commande unique "Go" (Windows / PowerShell)
#
# Usage :
#   .\go.ps1            -> install si necessaire + lance le backend Flask (:8000)
#   .\go.ps1 install    -> installe seulement
#   .\go.ps1 backend    -> lance seulement le backend
#
# Particularites :
#   - Backend Python (Flask) lance via `python -m backend.main` (port 8000)
#   - Pas de frontend separe : Flask sert directement les HTML statiques de frontend/
#   - requirements.txt et .venv a la RACINE du projet (pas dans backend/)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Cmd  = if ($args.Count -gt 0) { $args[0] } else { "dev" }

function Step($m) { Write-Host "==> $m" -ForegroundColor Cyan }
function Info($m) { Write-Host "    $m" -ForegroundColor Yellow }
function Fail($m) { Write-Host "!!  $m" -ForegroundColor Red; exit 1 }

function Check-ExitCode($what) {
    if ($LASTEXITCODE -ne 0) { Fail "$what a echoue (exit code $LASTEXITCODE)." }
}
function Require($cmd) {
    if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) { Fail "Commande manquante : $cmd" }
}

# -------- Install --------
function Install-Backend {
    Step "Backend Python (Flask) : venv + pip install"
    Require "python"
    Push-Location $Root
    try {
        if (-not (Test-Path ".venv")) {
            python -m venv .venv
            Check-ExitCode "python -m venv"
        }
        $py = Join-Path $Root ".venv\Scripts\python.exe"
        & $py -m pip install --upgrade pip
        Check-ExitCode "pip upgrade"
        & $py -m pip install --prefer-binary -r requirements.txt
        Check-ExitCode "pip install -r requirements.txt"
    } finally { Pop-Location }
}

function Enable-GitHooks {
    if (-not (Test-Path (Join-Path $Root ".git"))) { return }
    if (-not (Test-Path (Join-Path $Root ".githooks\post-merge"))) { return }
    Push-Location $Root
    try {
        $current = (git config --get core.hooksPath) 2>$null
        if ($current -ne ".githooks") {
            git config core.hooksPath .githooks
            Info "Hook git post-merge active (auto-install apres chaque git pull)."
        }
    } finally { Pop-Location }
}

function Install-All {
    Install-Backend
    Enable-GitHooks
}

function Need-Install {
    $py = Join-Path $Root ".venv\Scripts\python.exe"
    if (-not (Test-Path $py)) { return $true }
    & $py -c "import flask" 2>$null
    if ($LASTEXITCODE -ne 0) { return $true }
    return $false
}

# -------- Run --------
function Start-Backend {
    $py = Join-Path $Root ".venv\Scripts\python.exe"
    if (-not (Test-Path $py)) { Fail "venv introuvable. Lance : .\go.ps1 install" }
    Push-Location $Root
    try {
        Info "Flask demarre sur http://localhost:8000  (Ctrl+C pour arreter)"
        & $py -m backend.main
    } finally { Pop-Location }
}

# -------- Dispatch --------
switch ($Cmd) {
    "install" { Install-All; Write-Host "`nInstallation OK." -ForegroundColor Green; break }
    "backend" { if (Need-Install) { Install-All }; Start-Backend; break }
    "dev"     {
        if (Need-Install) { Install-All }
        if (-not (Test-Path (Join-Path $Root ".venv\Scripts\python.exe"))) {
            Fail "Le backend n'est pas installe correctement. Corrige les erreurs ci-dessus puis relance."
        }
        Start-Backend
        break
    }
    default { Fail "Commande inconnue : $Cmd. Utilise: install | dev | backend" }
}
