import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from io import BytesIO

# Configuration de la page
st.set_page_config(
    page_title="✨ Excel Beautifier",
    page_icon="🎨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Titre et description
st.title("✨ Excel Beautifier")
st.markdown("""
**Transformez vos fichiers Excel en un clin d'œil !**  
Glissez-déposez votre fichier, choisissez un thème, et obtenez un résultat professionnel.
""")

# Barre latérale pour les options
with st.sidebar:
    st.header("🎨 Personnalisation")
    
    # Upload du fichier
    uploaded_file = st.file_uploader(
        "📂 Choisissez votre fichier Excel", 
        type=["xlsx"],
        help="Formats supportés : .xlsx"
    )
    
    st.divider()
    
    # Options de personnalisation (définies uniquement si fichier uploadé)
    if uploaded_file:
        # Charger et afficher le fichier avant traitement
        try:
            # Lire le fichier Excel correctement
            data = pd.read_excel(uploaded_file, engine='openpyxl')
            st.success("✅ Fichier chargé avec succès !")
            
            # Prévisualisation
            st.subheader("📊 Aperçu du fichier")
            st.dataframe(data, use_container_width=True)
            
            st.divider()
            
            # Options de personnalisation
            st.subheader("⚙️ Options")
            
            theme = st.selectbox(
                "🎨 Choisir un thème",
                ["Professionnel", "Moderne", "Épuré"],
                index=0
            )
            
            header_size = st.slider(
                "🔤 Taille police en-tête",
                10, 30, 14,
                1
            )
            
            bold = st.checkbox(
                "🔤 Police grasse",
                value=True
            )
            
            alignment = st.selectbox(
                "📝 Alignement en-têtes",
                ["Centré", "Gauche", "Droite"],
                index=0
            )
            
            st.divider()
            
            if st.button("✨ Embellir le fichier", type="primary", use_container_width=True):
                st.info("🔄 Traitement en cours...")
                process_excel(uploaded_file.getvalue(), theme, header_size, bold, alignment)
        except Exception as e:
            st.error(f"❌ Erreur lors de la lecture du fichier : {str(e)}")
    
    st.divider()
    
    st.markdown("""
### 💡 Astuces
- **Professionnel** : Bleu classique, idéal pour le business
- **Moderne** : Bleu foncé, look contemporain
- **Épuré** : Gris neutre, minimaliste
    """)

# Fonction principale de traitement
def process_excel(file_bytes, theme, header_size, bold, alignment):
    try:
        # Charger le workbook depuis les bytes
        wb = load_workbook(file_bytes, read_only=False)
        ws = wb.active
        
        # Définition des couleurs par thème
        theme_colors = {
            "Professionnel": {
                "header_font_color": "FFFFFF",
                "header_fill": "2E74B5",
                "zebra_fill": "F2F2F2"
            },
            "Moderne": {
                "header_font_color": "FFFFFF",
                "header_fill": "1E3A5F",
                "zebra_fill": "F5F7FA"
            },
            "Épuré": {
                "header_font_color": "333333",
                "header_fill": "F8F9FA",
                "zebra_fill": "FFFFFF"
            }
        }
        
        colors = theme_colors[theme]
        
        # Créer les styles
        header_font = Font(
            color=colors["header_font_color"],
            name="Arial",
            size=header_size,
            bold=bold
        )
        
        header_fill = PatternFill(
            start_color=colors["header_fill"],
            end_color=colors["header_fill"],
            fill_type="solid"
        )
        
        zebra_fill = PatternFill(
            start_color=colors["zebra_fill"],
            end_color=colors["zebra_fill"],
            fill_type="solid"
        )
        
        border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )
        
        # Alignements
        if alignment == "Centré":
            header_align = Alignment(horizontal="center", vertical="center")
        elif alignment == "Gauche":
            header_align = Alignment(horizontal="left", vertical="center")
        else:
            header_align = Alignment(horizontal="right", vertical="center")
        
        # Appliquer les styles aux en-têtes (ligne 1)
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(row=1, column=col)
            cell.font = header_font
            cell.fill = header_fill
            cell.border = border
            cell.alignment = header_align
        
        # Alternance de lignes (zebra striping)
        for row in range(2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                cell = ws.cell(row=row, column=col)
                if row % 2 == 0:
                    cell.fill = zebra_fill
                    cell.border = border
        
        # Sauvegarder le fichier
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        # Afficher le résultat
        st.success("✨ Fichier embellit avec succès !")
        st.markdown("---")
        
        # Bouton de téléchargement
        st.download_button(
            label="📥 Télécharger le fichier embellie",
            data=output,
            file_name=f"beautified_{uploaded_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        # Option pour télécharger avec le nom original
        st.download_button(
            label="📥 Télécharger avec le nom original",
            data=output,
            file_name=uploaded_file.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
        
        # Prévisualisation du fichier
        st.markdown("---")
        st.subheader("👁️ Aperçu du résultat")
        
        # Créer un DataFrame pour prévisualisation
        preview_data = pd.read_excel(output, engine='openpyxl')
        st.dataframe(preview_data, use_container_width=True)
        
    except Exception as e:
        st.error(f"❌ Erreur lors du traitement : {str(e)}")

# Footer
st.markdown("---")
st.markdown("""
**Excel Beautifier** - Transformez vos fichiers Excel en un clin d'œil ✨

Développé avec ❤️ pour simplifier votre travail
""")
