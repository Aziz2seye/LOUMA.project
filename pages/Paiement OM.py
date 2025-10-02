import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
import tempfile
from utils import load_vto  
st.title("PAIEMENT OM")



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
            df = df.rename(columns={"INSCRIPTIONS": "REALISATION"})
            
            df['LOGIN'] = df['LOGIN'].astype(str)
            
            df['NOM_VENDEUR'] = df['NOM_VENDEUR'].astype(str).str.strip().str.upper()
            df['PRENOM_VENDEUR'] = df['PRENOM_VENDEUR'].astype(str).str.strip().str.upper()

            # 🔍 Filtrage
            df_filtre = df[df['LOGIN'].isin(logins_concernes)]

            st.success("✅ Fichier filtré avec succès !")
            st.write("📊 Ventes LOUMA mensuels :", df_filtre.shape[0], "lignes")
            st.dataframe(df_filtre)

            
            
            df_filtre['OBJECTIF'] = 120
            df_filtre["TAUX D'ATTEINTE"] = (df_filtre['REALISATION'] / df_filtre['OBJECTIF']).apply(lambda x: f"{round(x*100)}%")
            df_filtre['SI 100% ATTEINT'] = 25000
            df_filtre['PAIEMENT'] = df_filtre['REALISATION'].apply(lambda x: 50000 if x >= 120 else round((x/120)*50000))
            df_filtre['PAIEMENT CHAUFFEUR'] = 150000
            
            df_filtre['TOTAL SIM+CHAUFFEUR'] = None

            
            

            



            
            # Export Excel
            buffer_paiement = BytesIO()
            with pd.ExcelWriter(buffer_paiement, engine='openpyxl') as writer:
                df_filtre.to_excel(writer, sheet_name='DETAILS PAIEMENT JUIN VTO', index=False)
                #df_filtre[cols_affichage].to_excel(writer, sheet_name='PAIEMENT PAR PVT', index=False)
                #df_par_pvt.to_excel(writer, sheet_name='PAIEMENT PAR PVT', index=False)
            buffer_paiement.seek(0)

            st.download_button(
                label="📥 Télécharger le fichier de Paiement Mensuel",
                data=buffer_paiement,
                file_name="paiement_mensuel_vto.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
