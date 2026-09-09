# 🚀 Guide de déploiement - Excel Beautifier

## 📦 Fichiers à déployer

Votre application est prête ! Voici tous les fichiers à push sur GitHub :

```
excel-beautifier/
├── app.py                    # ✅ Fichier principal Streamlit
├── requirements_streamlit.txt # ✅ Dépendances Python
└── README.md                 # ⏳ À ajouter (voir plus bas)
```

---

## 🌐 Étape 1: Créer un compte GitHub

1. Allez sur https://github.com
2. Cliquez sur "Sign up"
3. Créez un compte gratuit (3 min)
4. Validez votre email

---

## 📂 Étape 2: Créer un dépôt GitHub

1. Sur GitHub, cliquez sur **+** (en haut à droite) → **New repository**
2. Nom du dépôt : `excel-beautifier`
3. Description (optionnel) : "Beautify Excel files with a single click"
4. **NE PAS** cocher "Add a README file"
5. Cliquez sur **Create repository**

---

## 📤 Étape 3: Pusher votre code

### Option A: Via l'interface GitHub (Plus simple)

1. Dans votre dépôt, cliquez sur **Add file** → **Upload files**
2. Glissez-déposez :
   - `app.py`
   - `requirements_streamlit.txt`
3. En bas, dans la case "Commit changes", écrivez :
   ```
   First commit: Add app.py and requirements.txt
   ```
4. Cliquez sur **Commit changes**

### Option B: Via terminal (Si vous préférez)

```bash
cd /home/zotac/excel-beautifier
git init
git add app.py requirements_streamlit.txt
git commit -m "First commit: Add app.py and requirements.txt"
git branch -M main
git remote add origin https://github.com/VOTRE_NOM/excel-beautifier.git
git push -u origin main
```

---

## 🌟 Étape 4: Déployer sur Streamlit Cloud

1. Allez sur https://streamlit.io/cloud
2. Connectez votre compte GitHub (bouton "Sign in with GitHub")
3. Cliquez sur **New app from GitHub**
4. Sélectionnez le dépôt **excel-beautifier**
5. Cliquez sur **Deploy**

---

## ✅ Étape 5: Attendre le déploiement

Streamlit Cloud va :
- Installer les dépendances (1-2 min)
- Lancer l'application (1-2 min)

Votre lien d'accès sera : **https://excel-beautifier.streamlit.app**

---

## 🎉 Étape 6: Tester !

1. Ouvrez votre lien
2. Uploader un fichier Excel
3. Choisissez un thème
4. Cliquez sur "Embellir le fichier"
5. Téléchargez le résultat !

---

## 🛠️ Dépannage

### "Erreur: Cannot find module 'streamlit'"
→ Le déploiement échoue ? Cliquez sur "Redeploy" dans Streamlit Cloud.

### "L'application est lente"
→ C'est normal pour Streamlit Cloud gratuit. Patience !

### "Je veux changer le nom de l'URL"
→ Allez dans Settings → Change name (nécessite Streamlit Cloud Pro pour personnaliser)

---

## 📝 Bonus: Ajouter un README.md

Pour rendre votre dépôt plus professionnel, ajoutez ce fichier :

```markdown
# ✨ Excel Beautifier

Transformez vos fichiers Excel en un clin d'œil !

## 🎯 Fonctionnalités

- 📂 Upload de fichiers Excel
- 🎨 3 thèmes personnalisables
- 🔤 Personnalisation des polices et alignement
- 📊 Alternance de lignes (zebra striping)
- 📥 Téléchargement immédiat

## 🚀 Déploiement

Déployé sur [Streamlit Cloud](https://streamlit.io/cloud)

## 💡 Comment l'utiliser

1. Allez sur https://excel-beautifier.streamlit.app
2. Uploader votre fichier Excel
3. Choisissez votre thème préféré
4. Cliquez sur "Embellir le fichier"
5. Téléchargez le résultat !

## 🛠️ Technologies

- **Frontend:** Streamlit
- **Backend:** Python
- **Librairies:** pandas, openpyxl
```

---

## 📞 Besoin d'aide ?

Si vous bloquez à une étape, dites-le-moi et je vous aide ! 😊

---

## 🎯 Résumé rapide

1. ✅ Créer compte GitHub (3 min)
2. ✅ Créer dépôt (1 min)
3. ✅ Push le code (1 min)
4. ✅ Déployer sur Streamlit Cloud (1 min)
5. ✅ Tester votre lien ! 🎉

**Total : environ 5-10 minutes !** 🚀
