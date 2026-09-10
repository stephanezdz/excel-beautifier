# ✨ Excel Beautifier

Transformez vos fichiers Excel en un clin d'œil !

## 🎯 Fonctionnalités

- 📂 Upload de fichiers Excel
- 🎨 3 thèmes personnalisables
- 🔤 Personnalisation des polices et alignement
- 📊 Alternance de lignes (zebra striping)
- 📥 Téléchargement immédiat

## 🚀 Déploiement

Déployé sur [Streamlit Cloud](https://excel-beautifier.streamlit.app)

## 💡 Comment l'utiliser

1. Allez sur https://excel-beautifier.streamlit.app
2. Uploader votre fichier Excel
3. Choisissez votre thème préféré
4. Cliquez sur "Embellir le fichier"
5. Téléchargez le résultat !

## 🛠️ Technologies

- **Application :** Streamlit, un seul fichier (`app.py`)
- **Librairies :** pandas, openpyxl, pillow
- **Avant tout envoi :** `./publier.sh "ce que j'ai changé"` (voir [CONTRIBUER.md](CONTRIBUER.md))

## 📁 Fichiers

- `app.py` - Code principal
- `requirements_streamlit.txt` - Dépendances

## 🎨 Thèmes disponibles

### 📊 Professionnel
- Bleu classique (#2E74B5)
- Idéal pour le business

### 🎨 Moderne
- Bleu foncé (#1E3A5F)
- Look contemporain

### 🤍 Épuré
- Gris neutre (#F8F9FA)
- Minimaliste

## 💻 Développement local

```bash
pip install -r requirements_streamlit.txt
streamlit run app.py
```

## 📝 Licence

MIT License
