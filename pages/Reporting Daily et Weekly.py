import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import tempfile
from utils import load_vto  
st.title("📈 Générateur de Reporting")

# 🔁 Bouton pour revenir à la sélection
if st.session_state.get("reporting_type"):
    if st.button("↩️ Retour au menu principal"):
        st.session_state.reporting_type = None

if not st.session_state.get("reporting_type"):
    
    st.markdown("""**Choisissez un type de reporting :** """)

    col1, col2 = st.columns(2)
    with col1:
        if st.button("🕐 Reporting Journalier"):
            st.session_state.reporting_type = "journalier"
    with col2:
        if st.button("📅 Reporting Hebdomadaire"):
            st.session_state.reporting_type = "hebdomadaire"
    

# 🚀 Bloc principal : Reporting Journalier
if st.session_state.get("reporting_type") == "journalier":
    uploaded_file = st.file_uploader("📁 Importer le fichier Excel brut (Journalier)", type=["xlsx", "csv"])
    
    if uploaded_file: 
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, encoding='utf-8', sep='|')
        else:
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            selected_sheet = st.selectbox("🗂️ Choisir la feuille à exploiter :", options=sheet_names)
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

        # ✅ Charger logins depuis fichier VTO
        vto_df = load_vto()
        logins_concernes = vto_df["LOGIN"].astype(str).str.lower().tolist()
        details = ["En Cours-Identification", "Identifie", "Identifie Photo"]

        # ✅ Nettoyage des colonnes
        df = df.rename(columns={
            'MSISDN': 'TOTAL_SIM',
            'ACCUEIL_VENDEUR': 'PVT',
            'LOGIN_VENDEUR': 'LOGIN',
            'AGENCE_VENDEUR': 'DRV'
        })

        df['LOGIN'] = df['LOGIN'].astype(str).str.lower()
        df['DRV'] = df['DRV'].astype(str).str.strip().str.upper()
        df['NOM_VENDEUR'] = df['NOM_VENDEUR'].astype(str).str.strip().str.upper()
        df['PRENOM_VENDEUR'] = df['PRENOM_VENDEUR'].astype(str).str.strip().str.upper()

        # 🔍 Filtrage
        df_filtre = df[df['LOGIN'].isin(logins_concernes) & df['ETAT_IDENTIFICATION'].astype(str).isin(details)]

        st.success("✅ Fichier filtré avec succès !")
        st.write("📊 Ventes LOUMA journalier :", df_filtre.shape[0], "lignes")
        st.dataframe(df_filtre)

        # 📊 Résumé par VTO
        df_summary = df_filtre.groupby(['DRV', 'PVT', 'PRENOM_VENDEUR', 'NOM_VENDEUR', 'LOGIN']).agg({
            'TOTAL_SIM': 'count'
        }).reset_index().sort_values(['DRV', 'PVT'])

        # Remplacer DRV
        df_summary["DRV"] = df_summary["DRV"].replace({ 
            "DV-DRV2_DIRECTION REGIONALE DES VENTES DAKAR 2": "DR2",
            "DV-DRVS_DIRECTION REGIONALE DES VENTES SUD": "DR SUD",
            "DV-DRVSE_DIRECTION REGIONALE DES VENTES SUD-EST": "SUD EST",
            "DV-DRVN_DIRECTION REGIONALE DES VENTES NORD": "DR NORD",
            "DV-DRVC_DIRECTION REGIONALE DES VENTES CENTRE": "DR CENTRE",
            "DV-DRVE_DIRECTION REGIONALE DES VENTES EST": "DR EST"
        })

        df_summary_display = df_summary.copy()
        df_summary_display['DRV'] = df_summary_display['DRV'].mask(df_summary_display['DRV'].duplicated())
        df_summary_display['PVT'] = df_summary_display['PVT'].mask(df_summary_display['PVT'].duplicated())

        # 📊 Ventes par PVT
        df_summary2 = df_filtre.groupby(['DRV', 'PVT']).agg({'TOTAL_SIM': 'count'}).reset_index()
        df_summary2["OBJECTIF"] = 40
        df_summary2["TR"] = (df_summary2['TOTAL_SIM'] / df_summary2['OBJECTIF']).apply(lambda x: f"{round(x*100)}%")

        total_sim_sum = df_summary2['TOTAL_SIM'].sum()
        objectif_sum = df_summary2['OBJECTIF'].sum()
        tr_mean = df_summary2['TR'].apply(lambda x: float(x.strip('%'))).mean()

        df_summary2.loc['Total'] = ['', '', total_sim_sum, objectif_sum, f'{tr_mean:.1f}%']

        df_summary2["DRV"] = df_summary2["DRV"].replace({ 
            "DV-DRV2_DIRECTION REGIONALE DES VENTES DAKAR 2": "DR2",
            "DV-DRVS_DIRECTION REGIONALE DES VENTES SUD": "DRS",
            "DV-DRVSE_DIRECTION REGIONALE DES VENTES SUD-EST": "DRSE",
            "DV-DRVN_DIRECTION REGIONALE DES VENTES NORD": "DRN",
            "DV-DRVC_DIRECTION REGIONALE DES VENTES CENTRE": "DRC",
            "DV-DRVE_DIRECTION REGIONALE DES VENTES EST": "DRE"
        })

        

        # 🧾 Export Excel
        temp_file = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        with pd.ExcelWriter(temp_file.name, engine='openpyxl') as writer:
            df_summary_display.to_excel(writer, sheet_name='Résumé Ventes', index=False)
            df_summary2.to_excel(writer, sheet_name='Ventes Par PVT', index=False)

        wb = load_workbook(temp_file.name)
        wb.save(temp_file.name)

        final_buffer = BytesIO()
        wb.save(final_buffer)
        final_buffer.seek(0)

        st.success("✅ Fichier généré avec succès !")
        st.download_button(
            label="📥 Télécharger le fichier Excel",
            data=final_buffer,
            file_name="Daily Reporting.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

# 🚀 Bloc principal : Reporting Hebdomadaire
if st.session_state.get("reporting_type") == "hebdomadaire":
    uploaded_file = st.file_uploader("📁 Importer le fichier Excel brut (hebdomadaire)", type=["xlsx", "csv"])
    
    if uploaded_file: 
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, encoding='utf-8', sep=';')
        else:
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            selected_sheet = st.selectbox("🗂️ Choisir la feuille à exploiter :", options=sheet_names)
            df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

        # ✅ Charger logins depuis fichier VTO
        vto_df = load_vto()
        logins_concernes = vto_df["LOGIN"].astype(str).str.lower().tolist()
        details = ["En Cours-Identification", "Identifie", "Identifie Photo"]

        # ✅ Nettoyage des colonnes
        df = df.rename(columns={
            'MSISDN': 'TOTAL_SIM',
            'ACCUEIL_VENDEUR': 'PVT',
            'LOGIN_VENDEUR': 'LOGIN',
            'AGENCE_VENDEUR': 'DRV'
        })

        df['LOGIN'] = df['LOGIN'].astype(str).str.lower()
        df['DRV'] = df['DRV'].astype(str).str.strip().str.upper()
        df['NOM_VENDEUR'] = df['NOM_VENDEUR'].astype(str).str.strip().str.upper()
        df['PRENOM_VENDEUR'] = df['PRENOM_VENDEUR'].astype(str).str.strip().str.upper()

        # 🔍 Filtrage
        df_filtre = df[df['LOGIN'].isin(logins_concernes) & df['ETAT_IDENTIFICATION'].astype(str).isin(details)]

        st.success("✅ Fichier filtré avec succès !")
        st.write("📊 Ventes LOUMA journalier :", df_filtre.shape[0], "lignes")
        st.dataframe(df_filtre)

        # 📊 Résumé par VTO
        df_summary = df_filtre.groupby(['DRV', 'PVT', 'PRENOM_VENDEUR', 'NOM_VENDEUR', 'LOGIN']).agg({
            'TOTAL_SIM': 'count'
        }).reset_index().sort_values(['DRV', 'PVT'])

        # Remplacer DRV
        df_summary["DRV"] = df_summary["DRV"].replace({ 
            "DV-DRV2_DIRECTION REGIONALE DES VENTES DAKAR 2": "DR2",
            "DV-DRVS_DIRECTION REGIONALE DES VENTES SUD": "DRS",
            "DV-DRVSE_DIRECTION REGIONALE DES VENTES SUD-EST": "DRSE",
            "DV-DRVN_DIRECTION REGIONALE DES VENTES NORD": "DRN",
            "DV-DRVC_DIRECTION REGIONALE DES VENTES CENTRE": "DRC",
            "DV-DRVE_DIRECTION REGIONALE DES VENTES EST": "DRE"
        })

        df_summary_display = df_summary.copy()
        df_summary_display['DRV'] = df_summary_display['DRV'].mask(df_summary_display['DRV'].duplicated())
        df_summary_display['PVT'] = df_summary_display['PVT'].mask(df_summary_display['PVT'].duplicated())

        # 📊 Ventes par PVT
        df_summary2 = df_filtre.groupby(['DRV', 'PVT']).agg({'TOTAL_SIM': 'count'}).reset_index()
        df_summary2["OBJECTIF"] = 240
        df_summary2["TR"] = (df_summary2['TOTAL_SIM'] / df_summary2['OBJECTIF']).apply(lambda x: f"{round(x*100)}%")

        total_sim_sum = df_summary2['TOTAL_SIM'].sum()
        objectif_sum = df_summary2['OBJECTIF'].sum()
        tr_mean = df_summary2['TR'].apply(lambda x: float(x.strip('%'))).mean()

        df_summary2.loc['Total'] = ['', '', total_sim_sum, objectif_sum, f'{tr_mean:.1f}%']

        df_summary2["DRV"] = df_summary2["DRV"].replace({ 
            "DV-DRV2_DIRECTION REGIONALE DES VENTES DAKAR 2": "DR2",
            "DV-DRVS_DIRECTION REGIONALE DES VENTES SUD": "DRS",
            "DV-DRVSE_DIRECTION REGIONALE DES VENTES SUD-EST": "DRSE",
            "DV-DRVN_DIRECTION REGIONALE DES VENTES NORD": "DRN",
            "DV-DRVC_DIRECTION REGIONALE DES VENTES CENTRE": "DRC",
            "DV-DRVE_DIRECTION REGIONALE DES VENTES EST": "DRE"
        })

        df_summary_display = df_summary_display.rename(columns={
            'TOTAL_SIM':'REALISATIONS'
        })
        df_summary2 = df_summary2.rename(columns={
            'TOTAL_SIM':'REALISATIONS'
        })

        # 🧾 Export Excel
        temp_file = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
        with pd.ExcelWriter(temp_file.name, engine='openpyxl') as writer:
            df_summary_display.to_excel(writer, sheet_name='Résumé Ventes', index=False)
            df_summary2.to_excel(writer, sheet_name='Ventes Par PVT', index=False)

        wb = load_workbook(temp_file.name)
        wb.save(temp_file.name)

        final_buffer = BytesIO()
        wb.save(final_buffer)
        final_buffer.seek(0)

        st.success("✅ Fichier généré avec succès !")
        st.download_button(
            label="📥 Télécharger le fichier Excel",
            data=final_buffer,
            file_name="Weekly Reporting.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


