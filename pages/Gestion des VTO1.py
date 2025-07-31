import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import tempfile
from io import BytesIO
import os
# 📁 Structure du projet :
# - main.py (point d'entrée)
# - pages/
#   - 1_Gestion_VTO.py
#   - 2_Reporting.py
#   - 3_Reporting_Mensuel.py
#   - data/
#     - vto_list.csv

# ====================
# main.py
# ====================
import streamlit as st

st.set_page_config(page_title="🧍 Gestion des VTO", layout="wide")

st.title("🧍 Gestion des VTO")

DATA_PATH = r"C:\Users\hp\Downloads\Dossier LOUMA\vto_list.xlsx"

# Charger la liste existante
def load_vto():
    if os.path.exists(DATA_PATH):
        return pd.read_excel(DATA_PATH)
    else:
        return pd.DataFrame(columns=["LOGIN", "PRENOM", "NOM"])

# Sauvegarder la liste
def save_vto(df):
    df.to_excel(DATA_PATH, index=False)

vto_df = load_vto()

st.subheader("📋 Liste actuelle des VTO")
st.dataframe(vto_df)

st.subheader("➕ Ajouter un VTO")
with st.form("form_ajout"):
    login = st.text_input("Login")
    prenom = st.text_input("Prénom")
    nom = st.text_input("Nom")
    submit = st.form_submit_button("Ajouter")
    if submit and login:
        new_vto = pd.DataFrame([[login, prenom, nom]], columns=["LOGIN", "PRENOM", "NOM"])
        vto_df = pd.concat([vto_df, new_vto], ignore_index=True)
        save_vto(vto_df)
        st.success("✅ VTO ajouté avec succès !")
        st.experimental_rerun()

st.subheader("✏️ Modifier un VTO")
if not vto_df.empty:
    login_to_edit = st.selectbox("Choisir un login à modifier :", vto_df["LOGIN"].unique())
    vto_to_edit = vto_df[vto_df["LOGIN"] == login_to_edit].iloc[0]

    with st.form("form_modif"):
        new_prenom = st.text_input("Nouveau prénom", vto_to_edit["PRENOM_VENDEUR"])
        new_nom = st.text_input("Nouveau nom", vto_to_edit["NOM_VENDEUR"])
        new_login = st.text_input("Nouveau login", vto_to_edit["LOGIN"])
        submit_modif = st.form_submit_button("Modifier")

        if submit_modif:
            vto_df.loc[vto_df["LOGIN"] == login_to_edit, ["LOGIN", "PRENOM", "NOM"]] = [new_login, new_prenom, new_nom]
            save_vto(vto_df)
            st.success("✏️ VTO modifié avec succès !")
            st.experimental_rerun()

st.subheader("🗑️ Supprimer un VTO")
login_to_delete = st.selectbox("Choisir un login à supprimer :", vto_df["LOGIN"])
if st.button("Supprimer"):
    vto_df = vto_df[vto_df["LOGIN"] != login_to_delete]
    save_vto(vto_df)
    st.success("❌ VTO supprimé !")
    st.experimental_rerun()


