import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import tempfile
from io import BytesIO
import os
from utils import load_vto  
from utils import load_pvt
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

st.set_page_config(page_title="📑 Gestion des PVT", layout="wide")

st.title("🧍 Gestion des PVT")

DATA_PATH1 = "pvt_list.xlsx"

# Charger la liste existante
#def load_vto():
    #if os.path.exists(DATA_PATH):
        #return pd.read_excel(DATA_PATH)
    #else:
        #return pd.DataFrame(columns=["LOGIN", "PRENOM", "NOM"])

# Sauvegarder la liste
def save_vto(df):
    df.to_excel(DATA_PATH1, index=False)

vto_df = load_pvt()

st.subheader("📋 Liste actuelle des PVT")
st.dataframe(vto_df)

st.subheader("➕ Ajouter un PVT")
with st.form("form_ajout"):
    nom = st.text_input("Nom PVT")
    contact = st.text_input("Contact")
    submit = st.form_submit_button("Ajouter")
    if submit and contact:
        new_pvt = pd.DataFrame([[nom, contact]], columns=["PVT", "CONTACT"])
        vto_df = pd.concat([vto_df, new_pvt], ignore_index=True)
        save_vto(vto_df)
        st.success("✅ PVT ajouté avec succès !")
        st.experimental_rerun()

st.subheader("✏️ Modifier un PVT")
if not vto_df.empty:
    nom_to_edit = st.selectbox("Choisir un pvt à modifier :", vto_df["PVT"].unique())
    pvt_to_edit = vto_df[vto_df["PVT"] == nom_to_edit].iloc[0]

    with st.form("form_modif"):
        new_nom = st.text_input("Nouveau nom PVT", pvt_to_edit["PVT"])
        new_contact = st.text_input("Nouveau contact", pvt_to_edit["CONTACT"])
        
        submit_modif = st.form_submit_button("Modifier")

        if submit_modif:
            vto_df.loc[vto_df["PVT"] == nom_to_edit, ["PVT", "CONTACT"]] = [new_nom, new_contact]
            save_vto(vto_df)
            st.success("✏️ PVT modifié avec succès !")
            st.experimental_rerun()

st.subheader("🗑️ Supprimer un PVT")
pvt_to_delete = st.selectbox("Choisir un pvt à supprimer :", vto_df["PVT"])
if st.button("Supprimer"):
    vto_df = vto_df[vto_df["PVT"] != pvt_to_delete]
    save_vto(vto_df)
    st.success("❌ PVT supprimé !")
    st.experimental_rerun()


