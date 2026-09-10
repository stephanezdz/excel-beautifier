# À lire avant de toucher à ce dépôt

Ces règles valent pour tout le monde : Hermès, Claude Code, ou Stéphane.

## L'application, c'est `app.py`. Rien d'autre.

Un seul fichier fait tourner tout ce qui est en ligne sur
[excel-beautifier.streamlit.app](https://excel-beautifier.streamlit.app).

⚠️ **Les dossiers `backend/` et `frontend/` sont abandonnés.** Ce sont les
restes d'une première version en React et FastAPI, remplacée par Streamlit.
`app.py` ne les appelle jamais. `deploy.py`, `deploy.md` et
`DEPLOY_INSTRUCTIONS.md` visent Hugging Face, où le projet n'est pas déployé.

👉 **« Mets tout à jour » ne veut donc rien dire ici** : il y a trois versions
du même produit dans le dépôt, et deux sont mortes. Ne travaille que sur
`app.py`, `requirements.txt` et `packages.txt`.

## Une seule commande avant d'envoyer

```bash
./publier.sh "ce que j'ai changé"
```

Elle contrôle, montre ce qui part, enregistre et envoie. Si le contrôle
échoue, **rien n'est enregistré et rien n'est envoyé**.

Pour contrôler sans rien envoyer :

```bash
python3 outils/verifier.py
```

Sept contrôles, chacun correspondant à une panne réellement arrivée sur ce
projet. Le détail est en commentaire dans le fichier.

## Ce qui est interdit

⛔ **Ne modifie jamais `test_data.xlsx`.** C'est le fichier de référence des
contrôles. Si tu as besoin d'un brouillon, copie-le sous un autre nom **hors
du dépôt** (`/tmp/`, par exemple).

⛔ **Ne laisse aucun fichier à la racine** : pas de sortie de test, pas de
copie de travail, pas de capture d'écran.

⛔ **Ne « répare » jamais un contrôle pour qu'il passe.** Si le contrôle
échoue, c'est le code qui est en cause. Dis-le et arrête-toi.

⛔ **Ne touche pas à ces deux lignes d'`app.py`** sans savoir ce que tu fais,
elles corrigent des bugs qui ont coûté cher :
- `load_workbook(BytesIO(file_bytes))` — openpyxl attend un fichier, pas des
  octets bruts.
- `output.seek(0)` avant le `return` — sans lui, l'aperçu vide le flux et le
  téléchargement rend un fichier de 0 octet.

## Comment rendre compte

**Montre la sortie des commandes, ne la résume pas.**

- ❌ « J'ai vérifié, tout est cohérent. »
- ✅ Coller les sept lignes du contrôle.

Une conclusion sans sortie de commande n'a aucune valeur ici. Le 9 septembre,
une application qui plantait à chaque clic a été déclarée « 100 % cohérente,
prête pour le déploiement » après une simple relecture du code. C'est
exactement ce que ces règles existent pour empêcher.
