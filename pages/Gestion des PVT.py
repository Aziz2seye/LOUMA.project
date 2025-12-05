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
    drv = st.selectbox("Direction Régionale (DRV)", [
        "DR2",
        "DR SUD",
        "SUD EST",
        "DR NORD",
        "DR CENTRE",
        "DR EST"
    ])
    nom = st.text_input("Nom PVT")
    contact = st.text_input("Contact")
    submit = st.form_submit_button("Ajouter")

    if submit and nom and contact:
        new_pvt = pd.DataFrame([[drv, nom, contact]], columns=["DRV", "PVT", "CONTACT"])
        pvt_df = pd.concat([pvt_df, new_pvt], ignore_index=True)
        save_pvt(pvt_df)
        st.success("✅ PVT ajouté avec succès !")
        st.rerun()

st.subheader("✏️ Modifier un PVT")
if not pvt_df.empty:
    nom_to_edit = st.selectbox("Choisir un PVT à modifier :", pvt_df["PVT"].unique())
    pvt_to_edit = pvt_df[pvt_df["PVT"] == nom_to_edit].iloc[0]

    with st.form("form_modif"):
        new_drv = st.selectbox("Nouvelle Direction Régionale", [
            "DR2",
            "DR SUD",
            "SUD EST",
            "DR NORD",
            "DR CENTRE",
            "DR EST"
        ], index=["DR2", "DR SUD", "SUD EST", "DR NORD", "DR CENTRE", "DR EST"].index(pvt_to_edit["DRV"]) if pvt_to_edit["DRV"] in ["DR2", "DR SUD", "SUD EST", "DR NORD", "DR CENTRE", "DR EST"] else 0)
        new_nom = st.text_input("Nouveau nom PVT", pvt_to_edit["PVT"])
        new_contact = st.text_input("Nouveau contact", pvt_to_edit["CONTACT"])

        submit_modif = st.form_submit_button("Modifier")

        if submit_modif:
            pvt_df.loc[pvt_df["PVT"] == nom_to_edit, ["DRV", "PVT", "CONTACT"]] = [new_drv, new_nom, new_contact]
            save_pvt(pvt_df)
            st.success("✏️ PVT modifié avec succès !")
            st.rerun()

st.subheader("🗑️ Supprimer un PVT")
if not pvt_df.empty:
    pvt_to_delete = st.selectbox("Choisir un PVT à supprimer :", pvt_df["PVT"])
    if st.button("Supprimer"):
        pvt_df = pvt_df[pvt_df["PVT"] != pvt_to_delete]
        save_pvt(pvt_df)
        st.success("❌ PVT supprimé !")
        st.rerun()