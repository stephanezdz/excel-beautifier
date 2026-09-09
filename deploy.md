# 🚀 Déploiement sur Hugging Face Spaces

## Option 1: Déploiement Rapide (Recommandé)

### Étape 1: Créer un compte Hugging Face
1. Allez sur https://huggingface.co/settings
2. Créez un compte gratuit (5 min)

### Étape 2: Créer un Space
1. Cliquez sur "New Space"
2. Nom: `excel-beautifier` (ou votre choix)
3. License: MIT
4. Language: **Python**
5. Framework: **Streamlit** (plus simple que FastAPI pour débuter)

### Étape 3: Installer les dépendances
Dans le terminal du Space, installez :
```bash
pip install pandas openpyxl xlsxwriter streamlit
```

### Étape 4: Créer app.py
Créez un fichier `app.py` avec le code ci-dessous.

### Étape 5: Lancer
```bash
streamlit run app.py
```

---

## Option 2: Déploiement avec FastAPI (Plus avancé)

### Étape 1: Créer un Space avec Docker
1. Nouveau Space -> Dockerfile
2. Pushz le code avec le Dockerfile fourni

### Dockerfile (backend)
```dockerfile
FROM python:3.11-slim

WORKDIR /app

COPY backend/requirements.txt .
RUN pip install -r requirements.txt

COPY backend/ .

CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
```

### Étape 2: Créer Dockerfile (frontend)
```dockerfile
FROM node:18-alpine

WORKDIR /app

COPY frontend/ .

RUN npm install && npm run build

CMD ["npm", "run", "preview"]
```

---

## Alternative: Render.com (Gratuit & Simple)

1. Allez sur https://render.com
2. Connectez votre GitHub
3. Nouveau projet -> Importer depuis GitHub
4. Choose "Web Service"
5. Build Command: `cd backend && pip install -r requirements.txt`
6. Start Command: `python main.py`

---

## Option la plus simple: Streamlit Cloud

Si vous convertissez le backend en Streamlit :

### 1. Installer Streamlit
```bash
pip install streamlit pandas openpyxl xlsxwriter
```

### 2. Créer app.py
```python
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from io import BytesIO

st.title("✨ Excel Beautifier")

uploaded_file = st.file_uploader("📂 Glissez votre fichier Excel ici", type=["xlsx"])

if uploaded_file:
    # Charger l'Excel
    data = pd.read_excel(uploaded_file)
    st.dataframe(data)
    
    # Options
    theme = st.selectbox("🎨 Choisir un thème", ["Professionnel", "Moderne", "Épuré"])
    header_size = st.slider("Taille police en-tête", 10, 30, 14)
    bold = st.checkbox("Police grasse", value=True)
    alignment = st.selectbox("Alignement", ["Centré", "Gauche", "Droite"], index=0)
    
    if st.button("✨ Embellir le fichier"):
        # Appliquer les styles
        wb = load_workbook(uploaded_file)
        ws = wb.active
        
        # Styles
        if theme == "Professionnel":
            header_font = Font(color="FFFFFF", name="Arial", size=header_size, bold=bold)
            fill = PatternFill(start_color="2E74B5", end_color="2E74B5", fill_type="solid")
        elif theme == "Moderne":
            header_font = Font(color="FFFFFF", name="Arial", size=header_size, bold=bold)
            fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
        else:  # Épuré
            header_font = Font(color="333333", name="Arial", size=header_size, bold=bold)
            fill = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")
        
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        if alignment == "Centré":
            cell_align = Alignment(horizontal="center", vertical="center")
        elif alignment == "Gauche":
            cell_align = Alignment(horizontal="left", vertical="center")
        else:
            cell_align = Alignment(horizontal="right", vertical="center")
        
        # Appliquer aux en-têtes (ligne 1)
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = fill
            cell.border = border
            cell.alignment = cell_align
        
        # Alternance de lignes
        for row in range(2, ws.max_row + 1):
            if row % 2 == 0:
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row=row, column=col)
                    cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                    cell.border = border
        
        # Sauvegarder
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        # Télécharger
        st.success("✅ Fichier embellit avec succès !")
        st.download_button(
            label="📥 Télécharger le fichier",
            data=output,
            file_name="beautified_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
```

### 3. Déployer sur Streamlit Cloud
1. Créez un dépôt GitHub
2. Push le fichier app.py
3. Allez sur https://streamlit.io/cloud
4. Connectez GitHub
5. Importez le dépôt
6. Cliquez sur "Deploy" !

---

## 🎯 Ma recommandation

Pour vous, je vous conseille **Streamlit Cloud** :
- ✅ 100% gratuit
- ✅ Pas de configuration Docker
- ✅ Déploiement en 5 min
- ✅ Interface intuitive
- ✅ Pas besoin de connaissances avancées

Souhaitez-vous que je crée le fichier `app.py` pour Streamlit maintenant ?

