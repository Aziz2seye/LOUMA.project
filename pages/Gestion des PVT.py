import streamlit as st
import pandas as pd
import os
import sys
from pathlib import Path
from PIL import Image

# ⚡ Ajouter le dossier parent au path pour utils.py
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))
from utils import load_pvt

# ====================
# Configuration page
# ====================
st.set_page_config(page_title="LOUMA - Gestion des PVT", layout="wide", initial_sidebar_state="expanded")

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

    /* Sidebar avec couleurs VERT/BLEU Sonatel comme l'accueil */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #009CA6 0%, #00A0B0 100%) !important;
        color: white !important;
        box-shadow: 4px 0 15px rgba(0, 156, 166, 0.2);
    }

    section[data-testid="stSidebar"] * {
        color: white !important;
    }

    /* Logo wrapper */
    .logo-wrapper {
        background: white;
        border-radius: 20px;
        padding: 20px;
        box-shadow: 0 8px 25px rgba(255, 121, 0, 0.3);
        border: 3px solid #FF7900;
        display: flex;
        align-items: center;
        justify-content: center;
        margin-bottom: 1rem;
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
        font-size: 1rem;
    }

    .custom-table td {
        padding: 10px;
        border-bottom: 1px solid #f0f0f0;
        font-size: 0.95rem;
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
            🧍 Gestion des PVT
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
# Fichier Excel et fonctions avec stockage persistant
# ====================
DATA_PATH = "pvt_list.xlsx"

async def save_pvt(df):
    """Sauvegarde les PVT dans le fichier Excel ET dans le stockage persistant"""
    # Sauvegarder dans le fichier Excel local
    df.to_excel(DATA_PATH, index=False)

    # Sauvegarder aussi dans le stockage persistant de Streamlit
    try:
        import json
        pvt_data = df.to_dict('records')
        await st.storage.set('pvt_data', json.dumps(pvt_data))
    except Exception as e:
        st.warning(f"Attention: Les données ne sont pas sauvegardées de façon permanente. Erreur: {e}")

async def load_pvt_from_storage():
    """Charge les PVT depuis le stockage persistant"""
    try:
        import json
        stored_data = await st.storage.get('pvt_data')
        if stored_data and stored_data.get('value'):
            pvt_list = json.loads(stored_data['value'])
            return pd.DataFrame(pvt_list)
    except Exception as e:
        print(f"Erreur chargement stockage: {e}")
    return None

# Charger les données
import asyncio

# Essayer d'abord depuis le stockage persistant
try:
    pvt_df_stored = asyncio.run(load_pvt_from_storage())
    if pvt_df_stored is not None and not pvt_df_stored.empty:
        pvt_df = pvt_df_stored
    else:
        pvt_df = load_pvt()
except Exception as e:
    pvt_df = load_pvt()

# ====================
# Liste des PVT avec style Orange/Cyan
# ====================
st.markdown('<div class="section-title">Liste actuelle des PVT</div>', unsafe_allow_html=True)

def render_pvt_table(df):
    html = '<div class="dataframe-container"><table class="custom-table">'
    # En-tête
    html += '<thead><tr><th>PVT</th><th>CONTACT</th></tr></thead><tbody>'
    # Lignes avec alternance Orange/Cyan
    colors = ["#FF7900", "#009CA6"]
    for i, row in df.iterrows():
        color = colors[i % 2]
        html += f'<tr style="background:{color}; color:white; font-weight:500;">'
        html += f'<td>{row["PVT"]}</td>'
        html += f'<td>{row["CONTACT"]}</td>'
        html += '</tr>'
    html += '</tbody></table></div>'
    st.markdown(html, unsafe_allow_html=True)

render_pvt_table(pvt_df)

# ====================
# Ajouter un PVT
# ====================
st.markdown('<div class="section-title">Ajouter un nouveau PVT</div>', unsafe_allow_html=True)
with st.form("form_ajout"):
    col1, col2 = st.columns(2)
    with col1:
        nom = st.text_input("Nom du PVT", placeholder="Ex: Ahmed Diop")
    with col2:
        contact = st.text_input("Numéro de contact", placeholder="Ex: +221 77 123 45 67")

    submit = st.form_submit_button("Ajouter le PVT")
    if submit and contact:
        new_pvt = pd.DataFrame([[nom, contact]], columns=["PVT", "CONTACT"])
        pvt_df = pd.concat([pvt_df, new_pvt], ignore_index=True)
        asyncio.run(save_pvt(pvt_df))
        st.success("✅ PVT ajouté avec succès !")
        st.rerun()

# ====================
# Modifier un PVT
# ====================
st.markdown('<div class="section-title">Modifier un PVT existant</div>', unsafe_allow_html=True)
if not pvt_df.empty:
    nom_to_edit = st.selectbox("Choisir un PVT à modifier :", pvt_df["PVT"].unique())
    pvt_to_edit = pvt_df[pvt_df["PVT"] == nom_to_edit].iloc[0]

    with st.form("form_modif"):
        col1, col2 = st.columns(2)
        with col1:
            new_nom = st.text_input("Nouveau nom", pvt_to_edit["PVT"])
        with col2:
            new_contact = st.text_input("Nouveau contact", pvt_to_edit["CONTACT"])

        submit_modif = st.form_submit_button("Enregistrer les modifications")
        if submit_modif:
            pvt_df.loc[pvt_df["PVT"] == nom_to_edit, ["PVT", "CONTACT"]] = [new_nom, new_contact]
            asyncio.run(save_pvt(pvt_df))
            st.success("✏️ PVT modifié avec succès !")
            st.rerun()
else:
    st.info("Aucun PVT disponible pour modification")

# ====================
# Supprimer un PVT
# ====================
st.markdown('<div class="section-title">Supprimer un PVT</div>', unsafe_allow_html=True)
if not pvt_df.empty:
    with st.form("form_suppr"):
        pvt_to_delete = st.selectbox("Choisir un PVT à supprimer :", pvt_df["PVT"])
        submit_suppr = st.form_submit_button("Supprimer définitivement")
        if submit_suppr:
            pvt_df = pvt_df[pvt_df["PVT"] != pvt_to_delete]
            asyncio.run(save_pvt(pvt_df))
            st.success("❌ PVT supprimé avec succès !")
            st.rerun()
else:
    st.info("Aucun PVT à supprimer")

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