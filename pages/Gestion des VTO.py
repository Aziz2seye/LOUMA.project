import streamlit as st
import pandas as pd
import os
from utils import load_vto, load_pvt

st.set_page_config(page_title="🧍 Gestion des VTO", layout="wide")
st.title("🧍 Gestion des VTO")

DATA_PATH = "vto_list.xlsx"

# Sauvegarder la liste
def save_vto(df):
    df.to_excel(DATA_PATH, index=False, sheet_name="vto")

vto_df = load_vto()
pvt_df = load_pvt()  # Charger la liste des PVT

st.subheader("📋 Liste actuelle des VTO")
st.dataframe(vto_df)

st.subheader("➕ Ajouter un VTO")
with st.form("form_ajout"):
    # Sélection de la Direction Régionale
    drv = st.selectbox("Direction Régionale (DR)", [
        "DR2",
        "DR SUD",
        "SUD EST",
        "DR NORD",
        "DR CENTRE",
        "DR EST"
    ])

    # Filtrer les PVT selon la DR sélectionnée
    pvt_filtres = pvt_df[pvt_df["DRV"] == drv]["PVT"].tolist()

    if len(pvt_filtres) > 0:
        pvt = st.selectbox("Point de Vente (PVT)", pvt_filtres)
    else:
        st.warning(f"⚠️ Aucun PVT trouvé pour {drv}. Ajoutez d'abord des PVT dans 'Gestion des PVT'.")
        pvt = None

    # Informations du VTO
    login = st.text_input("Login")
    prenom = st.text_input("Prénom")
    nom = st.text_input("Nom")
    kabbu = st.text_input("KABBU (numéro de téléphone)")

    submit = st.form_submit_button("Ajouter")

    if submit and login and pvt:
        new_vto = pd.DataFrame([[drv, pvt, login, prenom, nom, kabbu]],
                                columns=["DRV", "PVT", "LOGIN", "PRENOM", "NOM", "KABBU"])
        vto_df = pd.concat([vto_df, new_vto], ignore_index=True)
        save_vto(vto_df)
        st.success("✅ VTO ajouté avec succès !")
        st.rerun()

st.subheader("✏️ Modifier un VTO")
if not vto_df.empty:
    login_to_edit = st.selectbox("Choisir un login à modifier :", vto_df["LOGIN"].unique())
    vto_to_edit = vto_df[vto_df["LOGIN"] == login_to_edit].iloc[0]

    with st.form("form_modif"):
        # Sélection DR
        new_drv = st.selectbox("Direction Régionale", [
            "DR2", "DR SUD", "SUD EST", "DR NORD", "DR CENTRE", "DR EST"
        ], index=["DR2", "DR SUD", "SUD EST", "DR NORD", "DR CENTRE", "DR EST"].index(vto_to_edit["DRV"]) if vto_to_edit["DRV"] in ["DR2", "DR SUD", "SUD EST", "DR NORD", "DR CENTRE", "DR EST"] else 0)

        # Filtrer PVT selon DR
        pvt_filtres_modif = pvt_df[pvt_df["DRV"] == new_drv]["PVT"].tolist()

        if len(pvt_filtres_modif) > 0:
            current_pvt_index = pvt_filtres_modif.index(vto_to_edit["PVT"]) if vto_to_edit["PVT"] in pvt_filtres_modif else 0
            new_pvt = st.selectbox("Point de Vente", pvt_filtres_modif, index=current_pvt_index)
        else:
            st.warning(f"⚠️ Aucun PVT pour {new_drv}")
            new_pvt = vto_to_edit["PVT"]

        new_login = st.text_input("Nouveau login", vto_to_edit["LOGIN"])
        new_prenom = st.text_input("Nouveau prénom", vto_to_edit["PRENOM"])
        new_nom = st.text_input("Nouveau nom", vto_to_edit["NOM"])
        new_kabbu = st.text_input("Nouveau KABBU", vto_to_edit.get("KABBU", ""))

        submit_modif = st.form_submit_button("Modifier")

        if submit_modif:
            vto_df.loc[vto_df["LOGIN"] == login_to_edit,
                       ["DRV", "PVT", "LOGIN", "PRENOM", "NOM", "KABBU"]] = [
                new_drv, new_pvt, new_login, new_prenom, new_nom, new_kabbu
            ]
            save_vto(vto_df)
            st.success("✏️ VTO modifié avec succès !")
            st.rerun()

st.subheader("🗑️ Supprimer un VTO")
if not vto_df.empty:
    login_to_delete = st.selectbox("Choisir un login à supprimer :", vto_df["LOGIN"])
    if st.button("Supprimer"):
        vto_df = vto_df[vto_df["LOGIN"] != login_to_delete]
        save_vto(vto_df)
        st.success("❌ VTO supprimé !")
        st.rerun()