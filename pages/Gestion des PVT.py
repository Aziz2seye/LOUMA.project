import streamlit as st
import pandas as pd
import os
from utils import load_pvt

st.set_page_config(page_title="📑 Gestion des PVT", layout="wide")
st.title("🧍 Gestion des PVT")

DATA_PATH1 = "pvt_list.xlsx"

# Sauvegarder la liste
def save_pvt(df):
    df.to_excel(DATA_PATH1, index=False)

pvt_df = load_pvt()

st.subheader("📋 Liste actuelle des PVT")
st.dataframe(pvt_df)

st.subheader("➕ Ajouter un PVT")
with st.form("form_ajout"):
    nom = st.text_input("Nom PVT")
    contact = st.text_input("Contact")
    submit = st.form_submit_button("Ajouter")

    if submit and nom and contact:
        new_pvt = pd.DataFrame([[nom, contact]], columns=["PVT", "CONTACT"])
        pvt_df = pd.concat([pvt_df, new_pvt], ignore_index=True)
        save_pvt(pvt_df)
        st.success("✅ PVT ajouté avec succès !")
        st.rerun()  # ✅ CORRECTION ICI

st.subheader("✏️ Modifier un PVT")
if not pvt_df.empty:
    nom_to_edit = st.selectbox("Choisir un pvt à modifier :", pvt_df["PVT"].unique())
    pvt_to_edit = pvt_df[pvt_df["PVT"] == nom_to_edit].iloc[0]

    with st.form("form_modif"):
        new_nom = st.text_input("Nouveau nom PVT", pvt_to_edit["PVT"])
        new_contact = st.text_input("Nouveau contact", pvt_to_edit["CONTACT"])

        submit_modif = st.form_submit_button("Modifier")

        if submit_modif:
            pvt_df.loc[pvt_df["PVT"] == nom_to_edit, ["PVT", "CONTACT"]] = [new_nom, new_contact]
            save_pvt(pvt_df)
            st.success("✏️ PVT modifié avec succès !")
            st.rerun()  # ✅ CORRECTION ICI

st.subheader("🗑️ Supprimer un PVT")
pvt_to_delete = st.selectbox("Choisir un pvt à supprimer :", pvt_df["PVT"])
if st.button("Supprimer"):
    pvt_df = pvt_df[pvt_df["PVT"] != pvt_to_delete]
    save_pvt(pvt_df)
    st.success("❌ PVT supprimé !")
    st.rerun()  # ✅ CORRECTION ICI