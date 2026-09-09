"""
Excel Beautifier — met en forme un classeur Excel en trois niveaux d'intensité.

La mise en forme vit entièrement dans embellir(), qui ne connaît pas Streamlit :
elle prend des octets, rend des octets. Toute l'interface est en dessous.
"""

import random
from datetime import date, datetime
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ══════════════════════════════════════════════════════════════════════
#  LES PALETTES
#
#  Trois niveaux, du plus sobre au plus affirmé. Chaque niveau porte
#  plusieurs palettes : le bouton « Autre palette » tire dans cette liste
#  et nulle part ailleurs — jamais de couleur générée au hasard.
#
#  Pour changer une couleur, il suffit de changer le code ici.
#    entete     = fond de la ligne de titres (le texte y est toujours blanc)
#    alternance = fond d'une ligne sur deux dans le corps
#    filet      = trait fin sous l'en-tête          (niveau 2 seulement)
#    bordure    = trait vertical entre les colonnes (niveau 3 seulement)
#    chiffres   = couleur du texte des colonnes de nombres (niveau 3)
# ══════════════════════════════════════════════════════════════════════

NIVEAUX = {
    "Épuré": {
        "resume": "Noir et blanc, aucune bordure. Le papier.",
        "palettes": [
            {"nom": "Encre",    "entete": "1A1A1A", "alternance": "F7F7F7"},
            {"nom": "Graphite", "entete": "2D2D2D", "alternance": "FAFAFA"},
            {"nom": "Ardoise",  "entete": "404040", "alternance": "F4F4F4"},
            {"nom": "Absolu",   "entete": "000000", "alternance": "F9F9F9"},
        ],
    },
    "Professionnel": {
        "resume": "Une seule couleur, tenue. Un filet sous l'en-tête.",
        "palettes": [
            {"nom": "Ardoise",   "entete": "334155", "alternance": "F1F5F9", "filet": "CBD5E1"},
            {"nom": "Bleu nuit", "entete": "1E3A5F", "alternance": "EEF2F7", "filet": "C3D0E0"},
            {"nom": "Forêt",     "entete": "14532D", "alternance": "EFF6F0", "filet": "BBD6C3"},
            {"nom": "Bordeaux",  "entete": "7F1D1D", "alternance": "FBF0F0", "filet": "E5C0C0"},
            {"nom": "Terre",     "entete": "78350F", "alternance": "FAF5EF", "filet": "DFCBB4"},
        ],
    },
    "Moderne": {
        "resume": "Deux couleurs. La seconde ne sert qu'aux chiffres.",
        "palettes": [
            {"nom": "Bleu / ambre",       "entete": "1D4ED8", "alternance": "EFF4FF",
             "bordure": "C7D7FE", "chiffres": "B45309"},
            {"nom": "Émeraude / or",      "entete": "047857", "alternance": "ECFDF5",
             "bordure": "A7F3D0", "chiffres": "B45309"},
            {"nom": "Prune / rose",       "entete": "6D28D9", "alternance": "F5F0FF",
             "bordure": "DDD1FB", "chiffres": "BE185D"},
            {"nom": "Sarcelle / corail",  "entete": "0F766E", "alternance": "EFFAF9",
             "bordure": "99F6E4", "chiffres": "BE123C"},
            {"nom": "Anthracite / lime",  "entete": "27272A", "alternance": "F4F4F5",
             "bordure": "D4D4D8", "chiffres": "4D7C0F"},
        ],
    },
}

TEXTE_CORPS = "1F2328"     # gris très sombre, plus doux qu'un noir pur
TEXTE_ENTETE = "FFFFFF"


# ══════════════════════════════════════════════════════════════════════
#  LA GÉOMÉTRIE — ce qui rend un tableau lisible avant toute couleur
# ══════════════════════════════════════════════════════════════════════

def _lisible(valeur):
    """Le texte tel qu'Excel l'affichera, pour mesurer la largeur d'une colonne."""
    if valeur is None:
        return ""
    if isinstance(valeur, (datetime, date)):
        return valeur.strftime("%d/%m/%Y")
    if isinstance(valeur, float) and valeur == int(valeur):
        return str(int(valeur))
    return str(valeur)


def _type_colonne(valeurs):
    """
    Le type d'une colonne se lit dans son CONTENU, jamais dans son titre.
    Une colonne « Année » remplie de nombres est une colonne de nombres.
    """
    utiles = [v for v in valeurs if v is not None and v != ""]
    if not utiles:
        return "texte"
    if all(isinstance(v, (datetime, date)) for v in utiles):
        return "date"
    if all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in utiles):
        return "nombre"
    return "texte"


def _alignement_du_type(type_colonne):
    return {"nombre": "right", "date": "center", "texte": "left"}[type_colonne]


def _largeur_colonne(titre, valeurs):
    """Ajustée au contenu le plus long, entre 10 et 50 caractères, marges comprises."""
    longueurs = [len(_lisible(titre))] + [len(_lisible(v)) for v in valeurs]
    return max(10, min(50, max(longueurs) + 4))


# ══════════════════════════════════════════════════════════════════════
#  LE MOTEUR — octets en entrée, octets en sortie, aucune dépendance
#              à Streamlit : il se teste et se rejoue tout seul.
# ══════════════════════════════════════════════════════════════════════

def embellir(file_bytes, niveau, palette, taille_entete=14):
    """
    Met en forme la première feuille du classeur et rend le fichier prêt
    à télécharger. Ne lève rien : c'est l'appelant qui gère l'échec.
    """
    # /!\ BytesIO obligatoire : openpyxl attend un fichier, pas des octets bruts.
    classeur = load_workbook(BytesIO(file_bytes))
    feuille = classeur.active

    nb_lignes = feuille.max_row
    nb_colonnes = feuille.max_column

    fond_entete = PatternFill(start_color=palette["entete"],
                              end_color=palette["entete"], fill_type="solid")
    fond_alterne = PatternFill(start_color=palette["alternance"],
                               end_color=palette["alternance"], fill_type="solid")

    filet = palette.get("filet")        # niveau 2 : trait sous l'en-tête
    separateur = palette.get("bordure")  # niveau 3 : trait entre les colonnes
    couleur_chiffres = palette.get("chiffres")

    bas_entete = Border(bottom=Side(style="thin", color=filet)) if filet else Border()
    trait_gauche = Side(style="thin", color=separateur) if separateur else None

    for colonne in range(1, nb_colonnes + 1):
        titre = feuille.cell(row=1, column=colonne).value
        corps = [feuille.cell(row=l, column=colonne).value
                 for l in range(2, nb_lignes + 1)]

        type_colonne = _type_colonne(corps)
        alignement = _alignement_du_type(type_colonne)
        lettre = feuille.cell(row=1, column=colonne).column_letter
        feuille.column_dimensions[lettre].width = _largeur_colonne(titre, corps)

        # L'en-tête suit l'alignement de sa colonne, pas sa propre nature.
        entete = feuille.cell(row=1, column=colonne)
        entete.font = Font(bold=True, size=taille_entete, color=TEXTE_ENTETE)
        entete.fill = fond_entete
        entete.alignment = Alignment(horizontal=alignement, vertical="center", indent=1)
        entete.border = bas_entete

        couleur_texte = couleur_chiffres if (couleur_chiffres and type_colonne == "nombre") \
            else TEXTE_CORPS

        for ligne in range(2, nb_lignes + 1):
            cellule = feuille.cell(row=ligne, column=colonne)
            cellule.font = Font(bold=False, size=11, color=couleur_texte)
            if type_colonne == "date":
                cellule.number_format = "DD/MM/YYYY"
            # /!\ On ne touche PAS au format des nombres : un séparateur de
            # milliers transformerait une colonne d'années en « 2 026 ».
            cellule.alignment = Alignment(horizontal=alignement, vertical="center", indent=1)
            # On repart d'aucune bordure : le fichier d'origine en portait peut-être.
            if trait_gauche and colonne > 1:
                cellule.border = Border(left=trait_gauche)
            else:
                cellule.border = Border()
            if ligne % 2 == 0:
                cellule.fill = fond_alterne
            else:
                cellule.fill = PatternFill(fill_type=None)

    # Hauteurs : l'en-tête respire plus que le corps.
    feuille.row_dimensions[1].height = round(taille_entete * 2.0)
    for ligne in range(2, nb_lignes + 1):
        feuille.row_dimensions[ligne].height = 20

    feuille.freeze_panes = "A2"
    # Le quadrillage d'Excel parasite la mise en forme : les fonds suffisent.
    feuille.sheet_view.showGridLines = False

    sortie = BytesIO()
    classeur.save(sortie)
    sortie.seek(0)
    return sortie


def palette_du_niveau(niveau, index):
    palettes = NIVEAUX[niveau]["palettes"]
    return palettes[index % len(palettes)]


# ══════════════════════════════════════════════════════════════════════
#  L'INTERFACE
# ══════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="✨ Excel Beautifier",
    page_icon="🎨",
    layout="wide",
    initial_sidebar_state="expanded",
)

if "palette_index" not in st.session_state:
    st.session_state.palette_index = {niveau: 0 for niveau in NIVEAUX}
if "resultat" not in st.session_state:
    st.session_state.resultat = None


def _pastilles(palette):
    """Les couleurs de la palette, en petites pastilles, pour juger à l'œil."""
    cases = []
    for cle, legende in (("entete", "en-tête"), ("alternance", "alternance"),
                         ("filet", "filet"), ("bordure", "bordure"),
                         ("chiffres", "chiffres")):
        if palette.get(cle):
            cases.append(
                f'<div style="text-align:center">'
                f'<div style="width:38px;height:38px;border-radius:8px;'
                f'background:#{palette[cle]};border:1px solid rgba(128,128,128,.35)"></div>'
                f'<div style="font-size:10px;opacity:.6;margin-top:3px">{legende}</div>'
                f"</div>"
            )
    return f'<div style="display:flex;gap:10px;margin:6px 0 2px">{"".join(cases)}</div>'


st.title("✨ Excel Beautifier")
st.markdown(
    "**Transformez vos fichiers Excel en un clin d'œil !**  \n"
    "Glissez-déposez votre fichier, choisissez un niveau, et obtenez un résultat professionnel."
)

with st.sidebar:
    st.header("🎨 Personnalisation")

    fichier = st.file_uploader(
        "📂 Choisissez votre fichier Excel",
        type=["xlsx"],
        help="Formats supportés : .xlsx",
    )

    st.divider()

    if fichier:
        st.subheader("⚙️ Mise en forme")

        niveau = st.radio(
            "Niveau",
            list(NIVEAUX.keys()),
            index=1,
            help="Du plus sobre au plus affirmé.",
        )
        st.caption(NIVEAUX[niveau]["resume"])

        palettes = NIVEAUX[niveau]["palettes"]
        index = st.session_state.palette_index[niveau]
        palette = palette_du_niveau(niveau, index)

        st.markdown(f"**Palette :** {palette['nom']}  \n"
                    f"<span style='font-size:11px;opacity:.6'>"
                    f"{index + 1} sur {len(palettes)}</span>", unsafe_allow_html=True)
        st.markdown(_pastilles(palette), unsafe_allow_html=True)

        if st.button("🎲 Autre palette", use_container_width=True,
                     disabled=len(palettes) < 2):
            autres = [i for i in range(len(palettes)) if i != index]
            st.session_state.palette_index[niveau] = random.choice(autres)
            st.rerun()

        taille_entete = st.slider("🔤 Taille de l'en-tête", 10, 30, 14, 1)

        st.divider()

        if st.button("✨ Embellir le fichier", type="primary", use_container_width=True):
            try:
                st.session_state.resultat = embellir(
                    fichier.getvalue(), niveau, palette, taille_entete
                ).getvalue()
                st.session_state.resultat_nom = fichier.name
                st.session_state.resultat_legende = f"{niveau} · {palette['nom']}"
            except Exception as erreur:
                st.session_state.resultat = None
                st.error(f"❌ Erreur lors du traitement : {erreur}")

# ─── La zone principale : le fichier, puis le résultat ───────────────

if not fichier:
    st.info("👈 Commencez par déposer un fichier Excel dans le panneau de gauche.")
else:
    try:
        apercu = pd.read_excel(fichier, engine="openpyxl")
        st.subheader("📊 Le fichier d'origine")
        st.dataframe(apercu, use_container_width=True)
    except Exception as erreur:
        st.error(f"❌ Erreur lors de la lecture du fichier : {erreur}")

    if st.session_state.resultat:
        st.divider()
        st.success(f"✨ Fichier embelli — {st.session_state.resultat_legende}")

        gauche, droite = st.columns(2)
        with gauche:
            st.download_button(
                "📥 Télécharger le fichier embelli",
                data=st.session_state.resultat,
                file_name=f"beautified_{st.session_state.resultat_nom}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with droite:
            st.download_button(
                "📥 Télécharger avec le nom original",
                data=st.session_state.resultat,
                file_name=st.session_state.resultat_nom,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

        st.subheader("👁️ Aperçu du résultat")
        st.caption("Les couleurs et les largeurs n'apparaissent que dans Excel : "
                   "cet aperçu montre les données, pas la mise en forme.")
        st.dataframe(
            pd.read_excel(BytesIO(st.session_state.resultat), engine="openpyxl"),
            use_container_width=True,
        )

st.markdown("---")
st.markdown(
    "**Excel Beautifier** - Transformez vos fichiers Excel en un clin d'œil ✨\n\n"
    "Développé avec ❤️ pour simplifier votre travail"
)
