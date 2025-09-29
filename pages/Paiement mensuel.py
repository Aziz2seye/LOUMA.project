import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import tempfile
from utils import load_vto 
import streamlit as st
import pandas as pd
from io import BytesIO

import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import tempfile
from utils import load_vto 
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="📊 Paiement Mensuel Global", layout="wide")
st.title("📊 Génération Paiement Mensuel (SIM + OM)")

# === Upload des fichiers
file_sim = st.file_uploader("📥 Importer le fichier SIM", type=["xlsx", "csv"])
file_om = st.file_uploader("📥 Importer le fichier OM", type=["xlsx", "csv"])

if file_sim and file_om:

    # === Charger SIM
    if file_sim.name.endswith(".csv"):
        df_sim = pd.read_csv(file_sim, sep=";", encoding="utf-8")
    else:
        xls = pd.ExcelFile(file_sim)
        sheet_names = xls.sheet_names
        selected_sheet = st.selectbox("🗂️ Feuille SIM :", sheet_names)
        df_sim = pd.read_excel(file_sim, sheet_name=selected_sheet)

    # === Charger OM
    if file_om.name.endswith(".csv"):
        df_om = pd.read_csv(file_om, sep=";", encoding="utf-8")
    else:
        xls = pd.ExcelFile(file_om)
        sheet_names = xls.sheet_names
        selected_sheet = st.selectbox("🗂️ Feuille OM :", sheet_names)
        df_om = pd.read_excel(file_om, sheet_name=selected_sheet)

    # Harmonisation colonnes OM
    df_om = df_om.rename(columns={
        "NUMERO": "KABBU",
        "REALISATIONS AOUT": "REALISATION_OM"
    })

    # === Fusionner sur KABBU et vendeur
    df_final = pd.merge(
        df_sim,
        df_om[["LOGIN", "KABBU", "PRENOM_VENDEUR", "NOM_VENDEUR", "REALISATION_OM"]],
        on=["LOGIN", "KABBU", "PRENOM_VENDEUR", "NOM_VENDEUR"],
        how="outer"
    )

    # === Calcul des paiements
    # SIM
    df_final["OBJECTIF_SIM"] = 240
    df_final["TAUX_ATTEINTE_SIM"] = (df_final["REALISATION"] / df_final["OBJECTIF_SIM"]).apply(lambda x: f"{round(x*100)}%")
    df_final["SI_100_SIM"] = 100000
    df_final["PAIEMENT_SIM"] = df_final["REALISATION"].apply(lambda x: 100000 if x >= 240 else round((x/240)*100000))

    # OM
    df_final["OBJECTIF_OM"] = 120
    df_final["TAUX_ATTEINTE_OM"] = (df_final["REALISATION_OM"] / df_final["OBJECTIF_OM"]).apply(lambda x: f"{round(x*100)}%")
    df_final["SI_100_OM"] = 50000
    df_final["PAIEMENT_OM"] = df_final["REALISATION_OM"].apply(lambda x: 50000 if x >= 120 else round((x/120)*50000))

    # Chauffeur
    df_final["PAIEMENT_CHAUFFEUR"] = 150000
    df_final["TOTAL_PAIEMENT"] = df_final["PAIEMENT_SIM"].fillna(0) + df_final["PAIEMENT_OM"].fillna(0) + df_final["PAIEMENT_CHAUFFEUR"]

    # Réorganisation des colonnes
    cols_final = [
        "DRV", "PVT", "PRENOM_VENDEUR", "NOM_VENDEUR", "KABBU",
        "REALISATION", "OBJECTIF_SIM", "TAUX_ATTEINTE_SIM", "SI_100_SIM", "PAIEMENT_SIM",
        "REALISATION_OM", "OBJECTIF_OM", "TAUX_ATTEINTE_OM", "SI_100_OM", "PAIEMENT_OM",
        "PAIEMENT_CHAUFFEUR", "TOTAL_PAIEMENT"
    ]
    df_final = df_final[cols_final]

    st.success("✅ Fusion et calculs terminés !")
    st.dataframe(df_final)

    # === Résumé par PVT
    df_par_pvt = df_final.groupby(["DRV", "PVT"]).agg({
        "TOTAL_PAIEMENT": "sum"
    }).reset_index()

    df_par_pvt["GAIN PVT (5%)"] = df_par_pvt["TOTAL_PAIEMENT"] * 0.05
    df_par_pvt["TOTAL GENERAL"] = df_par_pvt["TOTAL_PAIEMENT"] + df_par_pvt["GAIN PVT (5%)"]

    # Ajouter ND PARTENAIRE (ex : numéro de téléphone du PVT)
    df_par_pvt["ND PARTENAIRE"] = ""  # 👉 tu pourras l'alimenter depuis ton fichier VTO

    # Réorganisation colonnes
    df_par_pvt = df_par_pvt[["DRV", "PVT", "ND PARTENAIRE", "TOTAL_PAIEMENT", "GAIN PVT (5%)", "TOTAL GENERAL"]]

    st.subheader("📊 Résumé par PVT")
    st.dataframe(df_par_pvt)

    # === Export Excel
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_final.to_excel(writer, sheet_name="Détails Paiement", index=False)
        df_par_pvt.to_excel(writer, sheet_name="Paiement par PVT", index=False)
    buffer.seek(0)

    st.download_button(
        label="📥 Télécharger le fichier Paiement Mensuel",
        data=buffer,
        file_name="paiement_mensuel_global.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
