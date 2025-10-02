import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import tempfile
from utils import load_vto  
st.title("📈 Reporting mensuel")

# 🔁 Bouton pour revenir à la sélection
if st.session_state.get("reporting_type"):
    if st.button("↩️ Retour au menu principal"):
        st.session_state.reporting_type = None

if not st.session_state.get("reporting_type"):
    
    #st.markdown("""**Choisissez un type de reporting :** """)

    col1, col2= st.columns(2)
    with col1:
        if st.button("📊 Générer le reporting du mois"):
            st.session_state.reporting_type = "reporting mensuel"
    with col2:
        if st.button("💰 Générer le paiement mensuel"):
            st.session_state.reporting_type = "paiement mensuel"

# 🚀 Bloc principal : Reporting mensuel
if st.session_state.get("reporting_type") == "reporting mensuel":
    uploaded_file = st.file_uploader("📁 Importer le fichier Excel brut (mensuel)", type=["xlsx", "csv"])
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
                'MSISDN': 'REALISATION',
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
            st.write("📊 Ventes LOUMA mensuels :", df_filtre.shape[0], "lignes")
            st.dataframe(df_filtre)

            # 📊 Résumé par VTO
            df_summary = df_filtre.groupby(['DRV', 'PVT', 'PRENOM_VENDEUR', 'NOM_VENDEUR', 'LOGIN']).agg({
                'REALISATION': 'count'
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
            df_summary2 = df_filtre.groupby(['DRV', 'PVT']).agg({'REALISATION': 'count'}).reset_index()
            df_summary2["OBJECTIF"] = 960
            df_summary2["TR"] = (df_summary2['REALISATION'] / df_summary2['OBJECTIF']).apply(lambda x: f"{round(x*100)}%")

            total_sim_sum = df_summary2['REALISATION'].sum()
            objectif_sum = df_summary2['OBJECTIF'].sum()
            tr_mean = df_summary2['TR'].apply(lambda x: float(x.strip('%'))).mean()

            df_summary2.loc['Total'] = ['', '', total_sim_sum, objectif_sum, f'{tr_mean:.1f}%']

            df_summary2["DRV"] = df_summary2["DRV"].replace({ 
                "DV-DRV2_DIRECTION REGIONALE DES VENTES DAKAR 2": "DR2",
                "DV-DRVS_DIRECTION REGIONALE DES VENTES SUD": "DR SUD",
                "DV-DRVSE_DIRECTION REGIONALE DES VENTES SUD-EST": "SUD EST",
                "DV-DRVN_DIRECTION REGIONALE DES VENTES NORD": "DR NORD",
                "DV-DRVC_DIRECTION REGIONALE DES VENTES CENTRE": "DR CENTRE",
                "DV-DRVE_DIRECTION REGIONALE DES VENTES EST": "DR EST"
            })

            # 🧾 Export Excel
            temp_file = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
            with pd.ExcelWriter(temp_file.name, engine='openpyxl') as writer:
                df_summary_display.to_excel(writer, sheet_name='Ventes par VTO', index=False)
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
                file_name="Monthly Reporting.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if st.session_state.get("reporting_type") == "paiement mensuel":

        uploaded_file = st.file_uploader("📁 Importer le fichier Excel brut (mensuel)", type=["xlsx", "csv"])
    
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
            logins_concernes = vto_df["LOGIN"].astype(str).tolist()
            details = ["En Cours-Identification", "Identifie", "Identifie Photo"]

            # ✅ Nettoyage des colonnes
            df = df.rename(columns={
                'MSISDN': 'REALISATION',
                'ACCUEIL_VENDEUR': 'PVT',
                'LOGIN_VENDEUR': 'LOGIN',
                'AGENCE_VENDEUR': 'DRV'
              })
            
            df['LOGIN'] = df['LOGIN'].astype(str)
            df['DRV'] = df['DRV'].astype(str).str.strip().str.upper()
            df['NOM_VENDEUR'] = df['NOM_VENDEUR'].astype(str).str.strip().str.upper()
            df['PRENOM_VENDEUR'] = df['PRENOM_VENDEUR'].astype(str).str.strip().str.upper()

            # 🔍 Filtrage
            df_filtre = df[df['LOGIN'].isin(logins_concernes) & df['ETAT_IDENTIFICATION'].astype(str).isin(details)]

            st.success("✅ Fichier filtré avec succès !")
            st.write("📊 Ventes LOUMA mensuels :", df_filtre.shape[0], "lignes")
            st.dataframe(df_filtre)

            # Remplacer DRV
            df_filtre["DRV"] = df_filtre["DRV"].replace({ 
                "DV-DRV2_DIRECTION REGIONALE DES VENTES DAKAR 2": "DR2",
                "DV-DRVS_DIRECTION REGIONALE DES VENTES SUD": "DR SUD",
                "DV-DRVSE_DIRECTION REGIONALE DES VENTES SUD-EST": "SUD EST",
                "DV-DRVN_DIRECTION REGIONALE DES VENTES NORD": "DR NORD",
                "DV-DRVC_DIRECTION REGIONALE DES VENTES CENTRE": "DR CENTRE",
                "DV-DRVE_DIRECTION REGIONALE DES VENTES EST": "DR EST"})
            
            #définir les colonnes pour les paiements
            df_filtre = df_filtre.groupby(['DRV', 'PVT', 'PRENOM_VENDEUR', 'NOM_VENDEUR', 'LOGIN']).agg({
            'REALISATION': 'count'}).reset_index().sort_values(['DRV', 'PVT'])
            df_filtre['OBJECTIF'] = 240
            df_filtre["TAUX D'ATTEINTE"] = (df_filtre['REALISATION'] / df_filtre['OBJECTIF']).apply(lambda x: f"{round(x*100)}%")
            df_filtre['SI 100% ATTEINT'] = 75000
            df_filtre['PAIEMENT'] = df_filtre['REALISATION'].apply(lambda x: 75000 if x >= 240 else round((x/240)*75000))
            df_filtre['PAIEMENT CHAUFFEUR'] = 100000
            df_filtre['PAIEMENT CHAUFFEUR'] = df_filtre['PAIEMENT CHAUFFEUR'].mask(df_filtre['PVT'].duplicated())
            df_filtre['TOTAL SIM+CHAUFFEUR'] = None

            # Fusionner pour ajouter la colonne KABBU
            df_filtre = df_filtre.merge(vto_df[["LOGIN", "KABBU"]], how="left")

            # Supprimer la colonne LOGIN si tu veux uniquement garder PVT et KABBU
            #df_filtre.drop(columns=["LOGIN"], inplace=True)

            # 👉 Ajouter les lignes de total après chaque DRV
            df_with_totals = pd.DataFrame(columns=df_filtre.columns)

            for drv, group in df_filtre.groupby('DRV'):
                df_with_totals = pd.concat([df_with_totals, group], ignore_index=True)

                total_paiement = group['PAIEMENT'].sum()
                total_general = group['PAIEMENT'].sum() + group['PAIEMENT CHAUFFEUR'].sum()
                row_total = {
                    'DRV': f"{drv}",
                    'PVT': "TOTAL PVT",
                    'PAIEMENT': total_paiement ,
                    'TOTAL SIM+CHAUFFEUR': total_general
                        }
                df_with_totals = pd.concat([df_with_totals, pd.DataFrame([row_total])], ignore_index=True)



            # === Générer tableau Paiement par PVT ===

            # 1. Grouper par DRV et PVT pour obtenir le total des paiements
            df_par_pvt = df_filtre.groupby(['DRV', 'PVT']).agg({'PAIEMENT': 'sum'}).reset_index()
            df_par_pvt = df_par_pvt.rename(columns={'PAIEMENT': 'MONTANT'})

            df_par_pvt['MONTANT'] = df_par_pvt['MONTANT'] + 100000

            # 2. Ajouter GAIN PVT (5%) et TOTAL GENERAL
            df_par_pvt['GAIN PVT (5%)'] = df_par_pvt['MONTANT'] * 0.05
            df_par_pvt['TOTAL GENERAL'] = df_par_pvt['MONTANT'] + df_par_pvt['GAIN PVT (5%)']




            # Affichage du tableau simplifié
            #cols_affichage = ['DRV', 'PVT', 'PRENOM_VENDEUR', 'NOM_VENDEUR', 'TOTAL_SIM']
            cols_affichage = ['DRV', 'PVT', 'PRENOM_VENDEUR', 'NOM_VENDEUR','KABBU','REALISATION', 'OBJECTIF', "TAUX D'ATTEINTE", 'SI 100% ATTEINT', 'PAIEMENT', 'PAIEMENT CHAUFFEUR', 'TOTAL SIM+CHAUFFEUR']
            st.dataframe(df_with_totals[cols_affichage])

            # Export Excel
            buffer_paiement = BytesIO()
            with pd.ExcelWriter(buffer_paiement, engine='openpyxl') as writer:
                df_with_totals[cols_affichage].to_excel(writer, sheet_name='DETAILS PAIEMENT JUIN VTO', index=False)
                #df_filtre[cols_affichage].to_excel(writer, sheet_name='PAIEMENT PAR PVT', index=False)
                df_par_pvt.to_excel(writer, sheet_name='PAIEMENT PAR PVT', index=False)
            buffer_paiement.seek(0)

            st.download_button(
                label="📥 Télécharger le fichier de Paiement Mensuel",
                data=buffer_paiement,
                file_name="paiement_mensuel_vto.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )



