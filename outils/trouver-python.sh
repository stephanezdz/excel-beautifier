#!/bin/sh
# Rend le chemin du premier Python capable de faire tourner le contrôle.
# Il en faut un qui ait pandas, openpyxl et pillow.
#
# Sur le PC, c'est celui d'Hermès qui les a déjà — inutile d'installer quoi
# que ce soit. Ailleurs, un environnement local .venv fait l'affaire.

RACINE=$(git rev-parse --show-toplevel 2>/dev/null || pwd)

for CANDIDAT in \
  "$RACINE/.venv/bin/python" \
  "$HOME/.hermes/hermes-agent/venv/bin/python" \
  python3 \
  python
do
  if "$CANDIDAT" -c "import pandas, openpyxl, PIL" >/dev/null 2>&1; then
    echo "$CANDIDAT"
    exit 0
  fi
done

echo "✗ Aucun Python utilisable pour le contrôle." >&2
echo "  Il lui faut pandas, openpyxl et pillow." >&2
echo >&2
echo "  Pour en fabriquer un dans le projet :" >&2
echo "    python3 -m venv .venv && .venv/bin/pip install -r requirements.txt" >&2
exit 1
