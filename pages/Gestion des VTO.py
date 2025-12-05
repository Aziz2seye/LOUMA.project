import streamlit as st
import pandas as pd
import os
import sys
from pathlib import Path
from PIL import Image

# ⚡ Ajouter le dossier parent au path pour utils.py
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from utils import load_vto

# ====================
# Configuration page
# ====================
st.set_page_config(page_title="LOUMA - Gestion des VTO", layout="wide", initial_sidebar_state="expanded")

# ====================
# CSS personnalisé Orange Sonatel
# ====================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;700&display=swap');

    /* Cache le header Streamlit */
    header[data-testid="stHeader"] { display: none; }

    /* Marges principales */
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 2rem !important;
    }

    /* Fond principal */
    .main {
        font-family: 'Poppins', sans-serif;
        background: linear-gradient(135deg, #fff5f0 0%, #ffffff 50%, #f0f8ff 100%);
    }

    /* Sidebar avec couleurs VERT/BLEU Sonatel */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #009CA6 0%, #00A0B0 100%) !important;
        color: white !important;
        box-shadow: 4px 0 15px rgba(0, 156, 166, 0.2);
    }

    section[data-testid="stSidebar"] * {
        color: white !important;
    }

    /* Conteneurs de tableau */
    .dataframe-container {
        box-shadow: 0 8px 20px rgba(255, 121, 0, 0.15);
        border-radius: 15px;
        padding: 1rem;
        background: white;
        margin-bottom: 1.5rem;
        border: 2px solid #FFE5CC;
    }

    /* Titres de section */
    .section-title {
        background: linear-gradient(135deg, #FF7900 0%, #FF5000 100%);
        color: white;
        padding: 0.7rem 1.2rem;
        border-radius: 10px;
        font-weight: 600;
        font-size: 1.1rem;
        margin-bottom: 0.8rem;
        box-shadow: 0 4px 12px rgba(255, 121, 0, 0.25);
        text-align: center;
        max-width: 400px;
        margin-left: auto;
        margin-right: auto;
    }

    /* Tableau personnalisé */
    .custom-table {
        width: 100%;
        border-collapse: collapse;
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
    }

    .custom-table th {
        background: linear-gradient(135deg, #009CA6 0%, #00B8C5 100%);
        color: white;
        font-weight: 600;
        padding: 10px;
        text-align: left;
        font-size: 0.95rem;
    }

    .custom-table td {
        padding: 10px;
        border-bottom: 1px solid #f0f0f0;
        font-size: 0.9rem;
    }

    .custom-table tr:hover {
        background: rgba(255, 121, 0, 0.1) !important;
        transform: scale(1.01);
        transition: all 0.2s ease;
    }

    .custom-table tr:hover td {
        color: #333 !important;
        font-weight: 600;
    }

    /* Boutons Streamlit avec style Orange */
    .stButton > button {
        background: linear-gradient(135deg, #FF7900 0%, #FF5000 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.6rem 2rem;
        font-weight: 600;
        font-size: 1rem;
        box-shadow: 0 4px 12px rgba(255, 121, 0, 0.3);
        transition: all 0.3s ease;
        width: 100%;
    }

    .stButton > button:hover {
        background: linear-gradient(135deg, #FF5000 0%, #FF3000 100%);
        box-shadow: 0 6px 18px rgba(255, 121, 0, 0.5);
        transform: translateY(-2px);
    }

    /* Boutons de formulaire */
    .stForm button[type="submit"] {
        background: linear-gradient(135deg, #FF7900 0%, #FF5000 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.6rem 2rem;
        font-weight: 600;
        font-size: 1rem;
        box-shadow: 0 4px 12px rgba(255, 121, 0, 0.3);
        transition: all 0.3s ease;
        width: 100%;
    }

    .stForm button[type="submit"]:hover {
        background: linear-gradient(135deg, #FF5000 0%, #FF3000 100%);
        box-shadow: 0 6px 18px rgba(255, 121, 0, 0.5);
        transform: translateY(-2px);
    }

    /* Champs de formulaire */
    .stTextInput > div > div > input,
    .stSelectbox > div > div > select {
        border: 2px solid #FFE5CC;
        border-radius: 10px;
        padding: 0.6rem;
        font-size: 1rem;
        transition: all 0.3s ease;
    }

    .stTextInput > div > div > input:focus,
    .stSelectbox > div > div > select:focus {
        border-color: #FF7900;
        box-shadow: 0 0 0 3px rgba(255, 121, 0, 0.1);
    }

    /* Messages de succès */
    .stSuccess {
        background: linear-gradient(135deg, #00D4AA 0%, #00B890 100%);
        color: white;
        border-radius: 10px;
        padding: 1rem;
        border: none;
    }

    /* Messages info */
    .stInfo {
        background: linear-gradient(135deg, #009CA6 0%, #00B8C5 100%);
        color: white;
        border-radius: 10px;
        padding: 1rem;
        border: none;
    }

    /* Cards pour formulaires */
    .form-card {
        background: white;
        border-radius: 12px;
        padding: 1.2rem;
        box-shadow: 0 4px 15px rgba(255, 121, 0, 0.1);
        border: 2px solid #FFE5CC;
        margin-bottom: 1.5rem;
    }

    /* Animations */
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(20px); }
        to { opacity: 1; transform: translateY(0); }
    }

    .dataframe-container, .form-card {
        animation: fadeIn 0.5s ease-out;
    }
</style>
""", unsafe_allow_html=True)

# ====================
# Charger le logo
# ====================
logo_path = Path(__file__).parent.parent / "assets" / "logo sonatel.png"
try:
    logo = Image.open(logo_path)
except FileNotFoundError:
    logo = None
    try:
        logo_path_alt = Path("assets") / "logo sonatel.png"
        logo = Image.open(logo_path_alt)
    except:
        pass

# ====================
# Header avec logo et titre
# ====================
col_logo, col_title = st.columns([1, 3])

with col_logo:
    if logo:
        st.image(logo, width=280)
    else:
        st.warning("Logo non trouvé")

with col_title:
    st.markdown("""
    <div style="
        background: linear-gradient(135deg, #FF7900 0%, #FF5000 100%);
        padding: 2rem;
        border-radius: 20px;
        box-shadow: 0 8px 25px rgba(255, 121, 0, 0.4);
        display: flex;
        flex-direction: column;
        justify-content: center;
        border: 3px solid rgba(255, 255, 255, 0.2);
        height: 100%;
    ">
        <h1 style="
            color: white;
            font-size: 2.5rem;
            font-weight: 700;
            margin: 0;
            text-shadow: 3px 3px 10px rgba(0, 0, 0, 0.3);
        ">
            🧍 Gestion des VTO
        </h1>
        <p style="
            color: rgba(255, 255, 255, 0.95);
            font-size: 1.2rem;
            margin: 0.8rem 0 0 0;
            font-weight: 400;
        ">
            Plateforme de gestion commerciale - Orange Sénégal
        </p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ====================
# Fichier Excel et fonctions
# ====================
DATA_PATH = "vto_list.xlsx"

def save_vto(df):
    df.to_excel(DATA_PATH, index=False)

vto_df = load_vto()

# ====================
# Liste des VTO avec style Orange/Cyan
# ====================
st.markdown('<div class="section-title">Liste actuelle des VTO</div>', unsafe_allow_html=True)

if vto_df.empty:
    st.info("📭 Aucun VTO enregistré pour le moment")
else:
    def render_vto_table(df):
        html = '<div class="dataframe-container"><table class="custom-table">'
        # En-tête
        html += '<thead><tr><th>DRV</th><th>Prénom</th><th>Nom</th><th>PVT</th><th>Login</th><th>KABBU</th></tr></thead><tbody>'
        # Lignes avec alternance Orange/Cyan
        colors = ["#FF7900", "#009CA6"]
        for i, row in df.iterrows():
            color = colors[i % 2]
            html += f'<tr style="background:{color}; color:white; font-weight:500;">'
            html += f'<td>{row.get("DRV", "N/A")}</td>'
            html += f'<td>{row.get("PRENOM_VENDEUR", "N/A")}</td>'
            html += f'<td>{row.get("NOM_VENDEUR", "N/A")}</td>'
            html += f'<td>{row.get("PVT", "N/A")}</td>'
            html += f'<td>{row.get("LOGIN", "N/A")}</td>'
            html += f'<td>{row.get("KABBU", "N/A")}</td>'
            html += '</tr>'
        html += '</tbody></table></div>'
        st.markdown(html, unsafe_allow_html=True)

    render_vto_table(vto_df)

# ====================
# Ajouter un VTO
# ====================
st.markdown('<div class="section-title">Ajouter un nouveau VTO</div>', unsafe_allow_html=True)
st.markdown('<p style="text-align:center; color:#666; font-size:0.95rem; margin-bottom:1rem;">Remplissez les informations du nouveau VTO</p>', unsafe_allow_html=True)

with st.form("form_ajout_vto"):
    col1, col2, col3 = st.columns(3)
    with col1:
        drv = st.text_input("DRV", placeholder="Ex: DRV_DAKAR")
        prenom = st.text_input("Prénom", placeholder="Ex: Moussa")
    with col2:
        nom = st.text_input("Nom", placeholder="Ex: Sow")
        pvt = st.text_input("PVT", placeholder="Ex: PVT_001")
    with col3:
        login = st.text_input("Login", placeholder="Ex: msow")
        kabbu = st.text_input("KABBU", placeholder="Ex: KABBU_123")

    submit_ajout = st.form_submit_button("➕ Ajouter le VTO")
    if submit_ajout:
        if all([drv, prenom, nom, pvt, login, kabbu]):
            new_vto = pd.DataFrame([[drv, prenom, nom, pvt, login, kabbu]],
                                   columns=["DRV", "PRENOM_VENDEUR", "NOM_VENDEUR", "PVT", "LOGIN", "KABBU"])
            vto_df = pd.concat([vto_df, new_vto], ignore_index=True)
            save_vto(vto_df)
            st.success("✅ VTO ajouté avec succès !")
            st.rerun()
        else:
            st.error("⚠️ Veuillez remplir tous les champs")

# ====================
# Modifier un VTO
# ====================
st.markdown('<div class="section-title">Modifier un VTO existant</div>', unsafe_allow_html=True)

if not vto_df.empty:
    # Créer une liste de noms complets pour la sélection
    vto_names = [f"{row['PRENOM_VENDEUR']} {row['NOM_VENDEUR']} ({row['LOGIN']})" for _, row in vto_df.iterrows()]
    selected_vto_name = st.selectbox("Choisir un VTO à modifier :", vto_names)

    # Récupérer l'index du VTO sélectionné
    selected_index = vto_names.index(selected_vto_name)
    vto_to_edit = vto_df.iloc[selected_index]

    with st.form("form_modif_vto"):
        col1, col2, col3 = st.columns(3)
        with col1:
            new_drv = st.text_input("DRV", value=str(vto_to_edit["DRV"]))
            new_prenom = st.text_input("Prénom", value=str(vto_to_edit["PRENOM_VENDEUR"]))
        with col2:
            new_nom = st.text_input("Nom", value=str(vto_to_edit["NOM_VENDEUR"]))
            new_pvt = st.text_input("PVT", value=str(vto_to_edit["PVT"]))
        with col3:
            new_login = st.text_input("Login", value=str(vto_to_edit["LOGIN"]))
            new_kabbu = st.text_input("KABBU", value=str(vto_to_edit["KABBU"]))

        submit_modif = st.form_submit_button("💾 Enregistrer les modifications")
        if submit_modif:
            vto_df.loc[selected_index, "DRV"] = new_drv
            vto_df.loc[selected_index, "PRENOM_VENDEUR"] = new_prenom
            vto_df.loc[selected_index, "NOM_VENDEUR"] = new_nom
            vto_df.loc[selected_index, "PVT"] = new_pvt
            vto_df.loc[selected_index, "LOGIN"] = new_login
            vto_df.loc[selected_index, "KABBU"] = new_kabbu
            save_vto(vto_df)
            st.success("✏️ VTO modifié avec succès !")
            st.rerun()
else:
    st.info("📭 Aucun VTO disponible pour modification")

# ====================
# Supprimer un VTO
# ====================
st.markdown('<div class="section-title">Supprimer un VTO</div>', unsafe_allow_html=True)

if not vto_df.empty:
    with st.form("form_suppr_vto"):
        # Créer une liste de noms complets pour la sélection
        vto_names_delete = [f"{row['PRENOM_VENDEUR']} {row['NOM_VENDEUR']} ({row['LOGIN']})" for _, row in vto_df.iterrows()]
        selected_vto_delete = st.selectbox("Choisir un VTO à supprimer :", vto_names_delete)

        submit_suppr = st.form_submit_button("🗑️ Supprimer définitivement")
        if submit_suppr:
            selected_index_delete = vto_names_delete.index(selected_vto_delete)
            vto_df = vto_df.drop(vto_df.index[selected_index_delete]).reset_index(drop=True)
            save_vto(vto_df)
            st.success("❌ VTO supprimé avec succès !")
            st.rerun()
else:
    st.info("📭 Aucun VTO à supprimer")

# ====================
# Footer
# ====================
st.markdown("""
<div style="
    margin-top: 2rem;
    padding: 1rem;
    background: linear-gradient(135deg, #009CA6 0%, #00B8C5 100%);
    border-radius: 12px;
    text-align: center;
    color: white;
    box-shadow: 0 4px 15px rgba(0, 156, 166, 0.3);
">
    <p style="margin: 0; font-size: 0.9rem; font-weight: 500;">
        Propulsé par Orange Sénégal - Sonatel SA | 2025
    </p>
</div>
""", unsafe_allow_html=True)