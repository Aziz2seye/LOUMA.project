import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import tempfile
import sys
from pathlib import Path
from PIL import Image
import plotly.express as px
import plotly.graph_objects as go
import os
import plotly.io as pio
from datetime import datetime

# Configuration pour l'export PNG
pio.kaleido.scope.default_format = "png"
pio.kaleido.scope.default_width = 1200
pio.kaleido.scope.default_height = 600
pio.kaleido.scope.default_scale = 2

# Ajouter le dossier pages au path
current_dir = Path(__file__).parent
if str(current_dir) not in sys.path:
    sys.path.insert(0, str(current_dir))

from db_manager import ReportingDatabase
# Ajouter le répertoire parent au path Python
current_dir = Path(__file__).parent
parent_dir = current_dir.parent
sys.path.insert(0, str(parent_dir))

from utils import load_vto

# ====================
# FONCTION D'EXPORT DES GRAPHIQUES EN PNG
# ====================
def download_plotly_as_png(fig, filename):
    """Convertit un graphique Plotly en PNG téléchargeable"""
    import io
    buffer = io.BytesIO()
    fig.write_image(buffer, format='png')
    buffer.seek(0)
    return buffer

# ====================
# Configuration page
# ====================
st.set_page_config(page_title="LOUMA - Reporting", layout="wide", initial_sidebar_state="expanded")

# Initialiser la base de données
if 'db_manager' not in st.session_state:
    st.session_state.db_manager = ReportingDatabase()
db = st.session_state.db_manager

# ====================
# CSS personnalisé Orange Sonatel
# ====================
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;700&display=swap');
    header[data-testid="stHeader"] { display: none; }
    .block-container { padding-top: 2rem !important; padding-bottom: 2rem !important; }
    .main { font-family: 'Poppins', sans-serif; background: linear-gradient(135deg, #fff5f0 0%, #ffffff 50%, #f0f8ff 100%); }
    section[data-testid="stSidebar"] { background: linear-gradient(180deg, #FF7900 0%, #FF5000 100%) !important; }
    section[data-testid="stSidebar"] * { color: white !important; }
    .stDataFrame { border-radius: 15px; overflow: hidden; box-shadow: 0 8px 20px rgba(255, 121, 0, 0.15); border: 2px solid #FFE5CC; }
    .section-title { background: linear-gradient(135deg, #FF7900 0%, #FF5000 100%); color: white; padding: 0.7rem 1.2rem; border-radius: 10px; font-weight: 600; font-size: 1.1rem; margin-bottom: 0.8rem; box-shadow: 0 4px 12px rgba(255, 121, 0, 0.25); text-align: center; max-width: 500px; margin-left: auto; margin-right: auto; }
    .stButton > button { background: linear-gradient(135deg, #FF7900 0%, #FF5000 100%); color: white; border: none; border-radius: 10px; padding: 0.8rem 2rem; font-weight: 600; font-size: 1.1rem; box-shadow: 0 4px 12px rgba(255, 121, 0, 0.3); transition: all 0.3s ease; width: 100%; }
    .stButton > button:hover { background: linear-gradient(135deg, #FF5000 0%, #FF3000 100%); box-shadow: 0 6px 18px rgba(255, 121, 0, 0.5); transform: translateY(-2px); }
    .stDownloadButton > button { background: linear-gradient(135deg, #00D4AA 0%, #00B890 100%); color: white; border: none; border-radius: 10px; padding: 0.8rem 2rem; font-weight: 600; font-size: 1.1rem; box-shadow: 0 4px 12px rgba(0, 212, 170, 0.3); transition: all 0.3s ease; width: 100%; }
    .stDownloadButton > button:hover { background: linear-gradient(135deg, #00B890 0%, #009A7A 100%); box-shadow: 0 6px 18px rgba(0, 212, 170, 0.5); transform: translateY(-2px); }
    .metric-card { background: white; border-radius: 12px; padding: 1.5rem; box-shadow: 0 4px 15px rgba(255, 121, 0, 0.15); border: 2px solid #FFE5CC; text-align: center; }
    .info-card { background: linear-gradient(135deg, #009CA6 0%, #00B8C5 100%); color: white; padding: 1.5rem; border-radius: 12px; box-shadow: 0 4px 15px rgba(0, 156, 166, 0.3); margin-bottom: 1.5rem; }
</style>
""", unsafe_allow_html=True)

# ====================
# Charger le logo
# ====================
logo = None
logo_paths = [
    parent_dir / "assets" / "logo sonatel.png",
    Path("assets") / "logo sonatel.png",
    Path("../assets/logo sonatel.png"),
]

for logo_path in logo_paths:
    try:
        if logo_path.exists():
            logo = Image.open(logo_path)
            break
    except:
        continue

# ====================
# Header avec logo et titre
# ====================
col_logo, col_title = st.columns([1, 3])

with col_logo:
    if logo:
        st.image(logo, width=280)

with col_title:
    st.markdown("""
    <div style="background: linear-gradient(135deg, #FF7900 0%, #FF5000 100%); padding: 2rem; border-radius: 20px; box-shadow: 0 8px 25px rgba(255, 121, 0, 0.4); display: flex; flex-direction: column; justify-content: center; border: 3px solid rgba(255, 255, 255, 0.2); height: 100%;">
        <h1 style="color: white; font-size: 2.5rem; font-weight: 700; margin: 0; text-shadow: 3px 3px 10px rgba(0, 0, 0, 0.3);">📈 Générateur de Reporting</h1>
        <p style="color: rgba(255, 255, 255, 0.95); font-size: 1.2rem; margin: 0.8rem 0 0 0; font-weight: 400;">Reporting Journalier & Hebdomadaire - Orange Sénégal</p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# Mapping DRV
DRV_MAPPING = {
    "DV-DRV2_DIRECTION REGIONALE DES VENTES DAKAR 2": "DR2",
    "DV-DRVS_DIRECTION REGIONALE DES VENTES SUD": "DRS",
    "DV-DRVSE_DIRECTION REGIONALE DES VENTES SUD-EST": "DRSE",
    "DV-DRVN_DIRECTION REGIONALE DES VENTES NORD": "DRN",
    "DV-DRVC_DIRECTION REGIONALE DES VENTES CENTRE": "DRC",
    "DV-DRVE_DIRECTION REGIONALE DES VENTES EST": "DRE"
}

# Bouton retour
if st.session_state.get("reporting_type"):
    if st.button("↩ Retour au menu principal"):
        st.session_state.reporting_type = None
        st.rerun()

# Menu principal
if not st.session_state.get("reporting_type"):
    st.markdown('<div class="section-title">Choisissez un type de reporting</div>', unsafe_allow_html=True)
    st.markdown('<div class="info-card"><h3 style="margin: 0 0 0.5rem 0;">ℹ Information</h3><p style="margin: 0;">Sélectionnez le type de rapport : journalier pour les performances quotidiennes ou hebdomadaire pour un résumé de la semaine.</p></div>', unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("🕐 Reporting Journalier", use_container_width=True):
            st.session_state.reporting_type = "journalier"
            st.rerun()
    with col2:
        if st.button("📅 Reporting Hebdomadaire", use_container_width=True):
            st.session_state.reporting_type = "hebdomadaire"
            st.rerun()
    with col3:
        if st.button("📊 Historique & Stats", use_container_width=True):
            st.session_state.reporting_type = "historique"
            st.rerun()

# REPORTING JOURNALIER
if st.session_state.get("reporting_type") == "journalier":
    st.markdown('<div class="section-title">🕐 Reporting Journalier</div>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("📁 Importer le fichier Excel brut (Journalier)", type=["xlsx", "csv"])

    if uploaded_file:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, encoding='utf-8', sep='|')
        else:
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            selected_sheet = st.selectbox("🗂 Choisir la feuille à exploiter :", options=sheet_names)
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

        vto_df = load_vto()
        logins_concernes = vto_df["LOGIN"].astype(str).str.lower().tolist()
        details = ["En Cours-Identification", "Identifie", "Identifie Photo"]

        column_mapping = {
            'MSISDN': 'TOTAL_SIM',
            'ACCUEIL_VENDEUR': 'PVT',
            'LOGIN_VENDEUR': 'LOGIN',
            'AGENCE_VENDEUR': 'DR'
        }

        missing_columns = [col for col in column_mapping.keys() if col not in df.columns]
        if missing_columns:
            st.error(f"❌ Colonnes manquantes : {', '.join(missing_columns)}")
            st.stop()

        df = df.rename(columns=column_mapping)
        df['LOGIN'] = df['LOGIN'].astype(str).str.lower()
        df['DR'] = df['DR'].astype(str).str.strip().str.upper()
        df['NOM_VENDEUR'] = df['NOM_VENDEUR'].astype(str).str.strip().str.upper()
        df['PRENOM_VENDEUR'] = df['PRENOM_VENDEUR'].astype(str).str.strip().str.upper()

        df_filtre = df[df['LOGIN'].isin(logins_concernes) & df['ETAT_IDENTIFICATION'].astype(str).isin(details)]
        df_filtre["DR"] = df_filtre["DR"].replace(DRV_MAPPING)

        st.success(f"✅ Fichier filtré avec succès ! {df_filtre.shape[0]} ventes journalières")

        # Afficher un message simple
        st.info("📊 Données chargées et prêtes pour l'analyse")

        # Afficher les données brutes
        st.dataframe(df_filtre.head(20), use_container_width=True)

# REPORTING HEBDOMADAIRE
elif st.session_state.get("reporting_type") == "hebdomadaire":
    st.markdown('<div class="section-title">📅 Reporting Hebdomadaire</div>', unsafe_allow_html=True)
    st.info("Section hebdomadaire - En cours de développement")

# HISTORIQUE
elif st.session_state.get("reporting_type") == "historique":
    st.markdown('<div class="section-title">📊 Historique & Statistiques</div>', unsafe_allow_html=True)
    st.info("Section historique - En cours de développement")