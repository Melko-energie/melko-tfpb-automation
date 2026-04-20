#!/usr/bin/env bash
# melko-tfpb-automation — commande unique "Go" (Linux / macOS / Git Bash)
#
# Usage :
#   bash go.sh            -> install si necessaire + lance le backend Flask (:8000)
#   bash go.sh install    -> installe seulement
#   bash go.sh backend    -> lance seulement le backend
#
# Particularites :
#   - Backend Python (Flask) lance via `python -m backend.main` (port 8000)
#   - Pas de frontend separe : Flask sert les HTML statiques de frontend/
#   - requirements.txt et .venv a la RACINE du projet (pas dans backend/)

set -e
ROOT="$(cd "$(dirname "$0")" && pwd)"
CMD="${1:-dev}"

step() { printf "\033[36m==> %s\033[0m\n" "$1"; }
info() { printf "\033[33m    %s\033[0m\n" "$1"; }
fail() { printf "\033[31m!!  %s\033[0m\n" "$1"; exit 1; }

require() { command -v "$1" >/dev/null 2>&1 || fail "Commande manquante : $1"; }

venv_python() {
  if   [ -x "$ROOT/.venv/Scripts/python.exe" ]; then echo "$ROOT/.venv/Scripts/python.exe"
  elif [ -x "$ROOT/.venv/bin/python" ];          then echo "$ROOT/.venv/bin/python"
  fi
}

install_backend() {
  step "Backend Python (Flask) : venv + pip install"
  require python
  ( cd "$ROOT" && [ -d ".venv" ] || python -m venv .venv )
  PY="$(venv_python)"
  "$PY" -m pip install --upgrade pip >/dev/null
  "$PY" -m pip install --prefer-binary -r "$ROOT/requirements.txt"
}

enable_git_hooks() {
  [ -d "$ROOT/.git" ] || return 0
  [ -f "$ROOT/.githooks/post-merge" ] || return 0
  local current
  current="$(cd "$ROOT" && git config --get core.hooksPath 2>/dev/null || true)"
  if [ "$current" != ".githooks" ]; then
    ( cd "$ROOT" && git config core.hooksPath .githooks )
    chmod +x "$ROOT/.githooks/post-merge" 2>/dev/null || true
    info "Hook git post-merge active (auto-install apres chaque git pull)."
  fi
}

install_all() {
  install_backend
  enable_git_hooks
}

need_install() {
  local py
  py="$(venv_python)"
  [ -z "$py" ] && return 0
  "$py" -c "import flask" 2>/dev/null || return 0
  return 1
}

start_backend() {
  PY="$(venv_python)"
  [ -z "$PY" ] && fail "venv introuvable. Lance : bash go.sh install"
  info "Flask demarre sur http://localhost:8000  (Ctrl+C pour arreter)"
  ( cd "$ROOT" && "$PY" -m backend.main )
}

case "$CMD" in
  install)  install_all ;;
  backend)  need_install && install_all; start_backend ;;
  dev|"")
    need_install && install_all
    start_backend
    ;;
  *) fail "Commande inconnue : $CMD. Utilise: install | dev | backend" ;;
esac
