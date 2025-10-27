import json
import streamlit as st
import gspread
import pandas as pd
from google.oauth2.service_account import Credentials
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import tempfile
from io import BytesIO
import os

st.set_page_config(page_title="🧍 Gestion des VTO", layout="wide")
st.title("🧍 Gestion des VTO")

# =========================
# 🔐 Auth: secrets → fallback fichier local (dev)
# =========================
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

service_info = None
if "gcp_service_account" in st.secrets:         # Streamlit Cloud OU local avec secrets.toml
    service_info = dict(st.secrets["gcp_service_account"])
else:                                            # Dev local sans secrets.toml → fichier .json
    with open("service_account.json", "r", encoding="utf-8") as f:
        service_info = json.load(f)

credentials = Credentials.from_service_account_info(service_info, scopes=SCOPES)
gc = gspread.authorize(credentials)

# =========================
# 📄 Cible : fichier + onglet
# =========================
# Ouvre par URL (plus robuste)
SHEET_URL = st.text_input(
    "URL du Google Sheet (partagé avec le service account):",
    value="https://docs.google.com/spreadsheets/d/1rwz5S71DXwD2r1bonqG1Sax1ebxl2dINi-RHypQMkSw/export?format=csv&gid=0"
)
WORKSHEET_NAME = st.text_input("Nom de l'onglet", value="LOUMA VTO")

@st.cache_data(ttl=30)
def get_worksheet(sheet_url: str, ws_name: str):
    sh = gc.open_by_url(sheet_url)
    ws = sh.worksheet(ws_name)
    return ws

def ensure_headers(ws):
    # S’assure que la 1re ligne = entêtes LOGIN / PRENOM / NOM
    values = ws.get_all_values()
    if not values:
        ws.update("A1:C1", [["LOGIN", "PRENOM", "NOM"]])
    else:
        headers = values[0]
        expected = ["LOGIN", "PRENOM", "NOM"]
        if [h.strip().upper() for h in headers[:3]] != expected:
            ws.update("A1:C1", [expected])

def load_vto_from_sheet(ws) -> pd.DataFrame:
    data = ws.get_all_records()
    if not data:
        return pd.DataFrame(columns=["LOGIN", "PRENOM", "NOM"])
    df = pd.DataFrame(data)
    # normalise les colonnes au cas où
    for col in ["LOGIN", "PRENOM", "NOM"]:
        if col not in df.columns:
            df[col] = ""
    return df[["LOGIN", "PRENOM", "NOM"]]

def add_vto(ws, login, prenom, nom):
    ws.append_row([login, prenom, nom])

def update_vto(ws, old_login, new_login, new_prenom, new_nom):
    records = ws.get_all_records()
    for i, row in enumerate(records, start=2):  # +1 pour entêtes → données commencent ligne 2
        if str(row.get("LOGIN", "")).strip() == str(old_login).strip():
            ws.update(f"A{i}:C{i}", [[new_login, new_prenom, new_nom]])
            return True
    return False

def delete_vto(ws, login):
    records = ws.get_all_records()
    for i, row in enumerate(records, start=2):
        if str(row.get("LOGIN", "")).strip() == str(login).strip():
            ws.delete_rows(i)
            return True
    return False

# =========================
# 🚀 UI
# =========================
if SHEET_URL and WORKSHEET_NAME:
    try:
        sheet = get_worksheet(SHEET_URL, WORKSHEET_NAME)
        ensure_headers(sheet)
        vto_df = load_vto_from_sheet(sheet)

        st.subheader("📋 Liste actuelle des VTO")
        st.dataframe(vto_df, use_container_width=True)

        st.subheader("➕ Ajouter un VTO")
        with st.form("form_ajout"):
            login = st.text_input("Login")
            prenom = st.text_input("Prénom")
            nom = st.text_input("Nom")
            submit = st.form_submit_button("Ajouter")
            if submit and login:
                add_vto(sheet, login, prenom, nom)
                st.success("✅ VTO ajouté avec succès !")
                st.rerun()

        st.subheader("✏️ Modifier un VTO")
        if not vto_df.empty:
            login_to_edit = st.selectbox("Choisir un login à modifier :", vto_df["LOGIN"].unique())
            vto_to_edit = vto_df[vto_df["LOGIN"] == login_to_edit].iloc[0]
            with st.form("form_modif"):
                new_prenom = st.text_input("Nouveau prénom", vto_to_edit["PRENOM_VENDEUR"])
                new_nom    = st.text_input("Nouveau nom",    vto_to_edit["NOM_VENDEUR"])
                new_login  = st.text_input("Nouveau login",  vto_to_edit["LOGIN"])
                submit_modif = st.form_submit_button("Modifier")
                if submit_modif:
                    ok = update_vto(sheet, login_to_edit, new_login, new_prenom, new_nom)
                    if ok:
                        st.success("✏️ VTO modifié avec succès !")
                        st.rerun()
                    else:
                        st.error("Login introuvable.")

        st.subheader("🗑️ Supprimer un VTO")
        if not vto_df.empty:
            login_to_delete = st.selectbox("Choisir un login à supprimer :", vto_df["LOGIN"])
            if st.button("Supprimer"):
                ok = delete_vto(sheet, login_to_delete)
                if ok:
                    st.success("❌ VTO supprimé !")
                    st.rerun()
                else:
                    st.error("Login introuvable.")

    except Exception as e:
        st.error(f"Erreur d’accès au Google Sheet : {e}")
