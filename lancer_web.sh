#!/usr/bin/env bash
# Praxedo Processor - Lanceur Web (macOS / Linux)

set -e
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null 2>&1; then
    echo "[ERREUR] python3 n'est pas installé."
    exit 1
fi

echo "========================================"
echo "  Praxedo Processor - Version Web"
echo "========================================"
echo ""
echo "[1/2] Installation des dépendances..."
python3 -m pip install -q -r requirements_web.txt

echo ""
echo "[2/2] Démarrage de l'application..."
echo "Le navigateur va s'ouvrir automatiquement."
echo "Appuyez sur Ctrl+C pour arrêter."
echo ""

python3 run_web.py
