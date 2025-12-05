import streamlit as st
import pandas as pd
import sqlite3
from contextlib import contextmanager
import sys
from pathlib import Path
from PIL import Image

# ⚡ Ajouter le dossier parent au path pour importer utils.py
current_dir = Path(_file_).parent
parent_dir = current_dir.parent
sys.path.insert(0, str(parent_dir))

# =========================
# Configuration Streamlit
# =========================
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

    /* Sidebar avec couleurs VERT/BLEU Sonatel comme l'accueil */
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #009CA6 0%, #00A0B0 100%) !important;
        color: white !important;
        box-shadow: 4px 0 15px rgba(0, 156, 166, 0.2);
    }

    section[data-testid="stSidebar"] * {
        color: white !important;
    }

    /* Style pour les dataframes */
    .stDataFrame {
        border-radius: 15px;
        overflow: hidden;
        box-shadow: 0 8px 20px rgba(255, 121, 0, 0.15);
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

    /* Messages d'erreur */
    .stError {
        background: linear-gradient(135deg, #FF3B30 0%, #D32F2F 100%);
        color: white;
        border-radius: 10px;
        padding: 1rem;
        border: none;
    }

    /* Messages d'info */
    .stInfo {
        background: linear-gradient(135deg, #009CA6 0%, #00B8C5 100%);
        color: white;
        border-radius: 10px;
        padding: 1rem;
        border: none;
    }

    /* Messages warning */
    .stWarning {
        background: linear-gradient(135deg, #FFB84D 0%, #FF9500 100%);
        color: white;
        border-radius: 10px;
        padding: 1rem;
        border: none;
    }

    /* Animations */
    @keyframes fadeIn {
        from { opacity: 0; transform: translateY(20px); }
        to { opacity: 1; transform: translateY(0); }
    }

    .dataframe-container, .stDataFrame {
        animation: fadeIn 0.5s ease-out;
    }
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

# =========================
# Gestion de connexion SQLite avec Context Manager
# =========================
DB_PATH = parent_dir / "louma_vto.db" if parent_dir.exists() else "louma_vto.db"

@contextmanager
def get_db_connection():
    """Context manager pour gérer proprement les connexions"""
    conn = sqlite3.connect(DB_PATH, timeout=10.0)
    try:
        yield conn
    finally:
        conn.close()

# Initialiser la base de données
def init_db():
    with get_db_connection() as conn:
        c = conn.cursor()
        c.execute("""
        CREATE TABLE IF NOT EXISTS vto (
            DRV TEXT,
            PVT TEXT,
            login TEXT PRIMARY KEY,
            prenom TEXT,
            nom TEXT,
            KABBU TEXT
        )
        """)
        conn.commit()

init_db()

# =========================
# Fonctions CRUD
# =========================
def get_all_vto():
    with get_db_connection() as conn:
        c = conn.cursor()
        c.execute("SELECT * FROM vto")
        rows = c.fetchall()
        if not rows:
            return pd.DataFrame(columns=["DRV", "PVT", "LOGIN", "PRENOM", "NOM", "KABBU"])
        return pd.DataFrame(rows, columns=["DRV", "PVT", "LOGIN", "PRENOM", "NOM", "KABBU"])

def add_vto(DRV, PVT, login, prenom, nom, KABBU):
    try:
        with get_db_connection() as conn:
            c = conn.cursor()
            c.execute(
                "INSERT INTO vto (DRV, PVT, login, prenom, nom, KABBU) VALUES (?, ?, ?, ?, ?, ?)",
                (DRV, PVT, login, prenom, nom, KABBU)
            )
            conn.commit()
        return True
    except sqlite3.IntegrityError:
        return False

def update_vto(old_login, DRV, PVT, new_login, new_prenom, new_nom, KABBU):
    try:
        with get_db_connection() as conn:
            c = conn.cursor()
            c.execute("""
                UPDATE vto SET DRV=?, PVT=?, login=?, prenom=?, nom=?, KABBU=? WHERE login=?
            """, (DRV, PVT, new_login, new_prenom, new_nom, KABBU, old_login))
            conn.commit()
            return c.rowcount > 0
    except sqlite3.OperationalError as e:
        st.error(f"Erreur de base de données : {e}")
        return False

def delete_vto(login):
    try:
        with get_db_connection() as conn:
            c = conn.cursor()
            c.execute("DELETE FROM vto WHERE login=?", (login,))
            conn.commit()
            return c.rowcount > 0
    except sqlite3.OperationalError as e:
        st.error(f"Erreur de base de données : {e}")
        return False

# =========================
# Affichage de la table
# =========================
st.markdown('<div class="section-title">Liste actuelle des VTO</div>', unsafe_allow_html=True)
vto_df = get_all_vto()

if not vto_df.empty:
    st.dataframe(
        vto_df[["DRV", "PVT", "LOGIN", "PRENOM", "NOM", "KABBU"]],
        use_container_width=True,
        hide_index=True
    )
else:
    st.info("📭 Aucun VTO enregistré pour le moment")

st.markdown("<br>", unsafe_allow_html=True)

# =========================
# Ajouter un VTO
# =========================
st.markdown('<div class="section-title">Ajouter un nouveau VTO</div>', unsafe_allow_html=True)
with st.form("form_ajout"):
    st.markdown("##### Remplissez les informations du nouveau VTO")

    col1, col2, col3 = st.columns(3)
    with col1:
        DRV = st.text_input("DRV", placeholder="Ex: Dakar")
        prenom = st.text_input("Prénom", placeholder="Ex: Amadou")
    with col2:
        PVT = st.text_input("PVT", placeholder="Ex: PVT001")
        nom = st.text_input("Nom", placeholder="Ex: Diop")
    with col3:
        login = st.text_input("Login", placeholder="Ex: adiop")
        KABBU = st.text_input("KABBU", placeholder="Ex: KB123")

    submit = st.form_submit_button("➕ Ajouter le VTO")
    if submit and login:
        success = add_vto(DRV, PVT, login, prenom, nom, KABBU)
        if success:
            st.success("✅ VTO ajouté avec succès !")
            st.rerun()
        else:
            st.error("❌ Ce login existe déjà !")

st.markdown("<br>", unsafe_allow_html=True)

# =========================
# Modifier un VTO
# =========================
st.markdown('<div class="section-title">Modifier un VTO existant</div>', unsafe_allow_html=True)
if not vto_df.empty:
    login_to_edit = st.selectbox("Choisir un login à modifier :", vto_df["LOGIN"].unique())
    vto_to_edit = vto_df[vto_df["LOGIN"] == login_to_edit].iloc[0]

    with st.form("form_modif"):
        st.markdown("##### Modifiez les informations")

        col1, col2, col3 = st.columns(3)
        with col1:
            DRV = st.text_input("DRV", vto_to_edit["DRV"])
            new_prenom = st.text_input("Nouveau prénom", vto_to_edit["PRENOM"])
        with col2:
            PVT = st.text_input("PVT", vto_to_edit["PVT"])
            new_nom = st.text_input("Nouveau nom", vto_to_edit["NOM"])
        with col3:
            new_login = st.text_input("Nouveau login", vto_to_edit["LOGIN"])
            KABBU = st.text_input("KABBU", vto_to_edit["KABBU"])

        submit_modif = st.form_submit_button("✏ Enregistrer les modifications")
        if submit_modif:
            ok = update_vto(login_to_edit, DRV, PVT, new_login, new_prenom, new_nom, KABBU)
            if ok:
                st.success("✏ VTO modifié avec succès !")
                st.rerun()
            else:
                st.error("❌ Erreur lors de la modification.")
else:
    st.info("📭 Aucun VTO disponible pour modification")

st.markdown("<br>", unsafe_allow_html=True)

# =========================
# Supprimer un VTO
# =========================
st.markdown('<div class="section-title">Supprimer un VTO</div>', unsafe_allow_html=True)
if not vto_df.empty:
    with st.form("form_suppr"):
        login_to_delete = st.selectbox("Choisir un login à supprimer :", vto_df["LOGIN"])
        st.warning("⚠ Cette action est irréversible !")

        submit_suppr = st.form_submit_button("🗑 Supprimer définitivement")
        if submit_suppr:
            ok = delete_vto(login_to_delete)
            if ok:
                st.success("✅ VTO supprimé avec succès !")
                st.rerun()
            else:
                st.error("❌ Erreur lors de la suppression.")
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