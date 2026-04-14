#!/bin/bash
echo "============================================"
echo "  Melko Energie - Installation automatique"
echo "============================================"
echo ""

# Verifier Python
if ! command -v python3 &> /dev/null && ! command -v python &> /dev/null; then
    echo "[ERREUR] Python n'est pas installe."
    echo "Installez Python 3.10+ : https://www.python.org/downloads/"
    exit 1
fi

PYTHON=$(command -v python3 || command -v python)
echo "[OK] $($PYTHON --version) detecte"

# Creer le venv
echo ""
echo "[1/3] Creation de l'environnement virtuel..."
if [ -d ".venv" ]; then
    echo "     .venv existe deja, on le reutilise."
else
    $PYTHON -m venv .venv 2>/dev/null
    if [ $? -ne 0 ]; then
        echo "     venv standard a echoue, tentative avec --without-pip..."
        $PYTHON -m venv .venv --without-pip 2>/dev/null
        if [ $? -ne 0 ]; then
            echo "[ERREUR] Impossible de creer le venv."
            exit 1
        fi
    fi
    echo "     [OK] .venv cree"
fi

# Activer le venv
echo ""
echo "[2/3] Activation du venv..."
if [ -f ".venv/bin/activate" ]; then
    source .venv/bin/activate
elif [ -f ".venv/Scripts/activate" ]; then
    source .venv/Scripts/activate
else
    echo "[ERREUR] Impossible d'activer le venv."
    exit 1
fi
echo "     [OK] venv active"

# Assurer pip
python -m ensurepip --upgrade 2>/dev/null

# Installer les dependances
echo ""
echo "[3/3] Installation des dependances..."
pip install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "[ERREUR] Echec de l'installation des dependances."
    exit 1
fi

echo ""
echo "============================================"
echo "  Installation terminee avec succes !"
echo "============================================"
echo ""
echo "Pour lancer l'application :"
echo "  source .venv/bin/activate   # Linux/Mac"
echo "  source .venv/Scripts/activate  # Windows Git Bash"
echo "  python -m backend.main"
echo ""
echo "Puis ouvrir http://localhost:8000"
