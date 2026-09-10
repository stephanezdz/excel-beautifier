#!/bin/bash
# Une seule commande pour tout envoyer, dans le bon ordre.
#
#   ./publier.sh "ce que j'ai changé"
#
# Elle contrôle, montre ce qui part, puis enregistre et envoie.
# Si le contrôle échoue, rien n'est enregistré et rien n'est envoyé.

set -u
cd "$(git rev-parse --show-toplevel)" || exit 1

MESSAGE="${1:-}"
if [ -z "$MESSAGE" ]; then
  echo "✗ Il manque le message. Exemple :"
  echo "  ./publier.sh \"corrige l'alignement des colonnes de dates\""
  exit 1
fi

if [ -z "$(git status --porcelain)" ]; then
  echo "Rien n'a changé. Il n'y a rien à envoyer."
  exit 0
fi

echo
echo "Ce qui va partir :"
git status --short | sed 's/^/   /'

PYTHON=$(outils/trouver-python.sh) || exit 1
"$PYTHON" outils/verifier.py || exit 1

git add -A
# --no-verify : le contrôle vient d'être passé juste au-dessus, inutile de
# le rejouer. Le crochet reste actif pour un « git commit » lancé à la main.
git commit -q --no-verify -m "$MESSAGE" || exit 1
echo "Enregistré : $(git log --oneline -1)"

echo "Envoi vers GitHub..."
if git push -q origin HEAD; then
  echo "✓ Envoyé. Streamlit se met à jour tout seul dans la minute."
else
  echo "✗ L'envoi a échoué. C'est enregistré en local, rien n'est perdu."
  exit 1
fi
