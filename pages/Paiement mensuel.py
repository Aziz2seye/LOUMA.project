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
from utils import load_vto

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


    #Harmonisation colonnes SIM
    # ✅ Charger logins depuis fichier VTO
    vto_df = load_vto()
    logins_concernes = vto_df["LOGIN"].astype(str).tolist()
    details = ["En Cours-Identification", "Identifie", "Identifie Photo"]
    df = df_sim.copy()
    # ✅ Nettoyage des colonnes
    df = df.rename(columns={
    'MSISDN': 'REALISATION_SIM',
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
            'REALISATION_SIM': 'count'}).reset_index().sort_values(['DRV', 'PVT'])
    df_filtre['OBJECTIF SIM'] = 240
    df_filtre["TAUX D'ATTEINTE SIM"] = (df_filtre['REALISATION_SIM'] / df_filtre['OBJECTIF SIM']).apply(lambda x: f"{round(x*100)}%")
    df_filtre['SI 100% ATTEINT SIM'] = 75000
    df_filtre['PAIEMENT_SIM'] = df_filtre['REALISATION_SIM'].apply(lambda x: 75000 if x >= 240 else round((x/240)*75000))
            #df_filtre['PAIEMENT CHAUFFEUR'] = 100000
            #df_filtre['PAIEMENT CHAUFFEUR'] = df_filtre['PAIEMENT CHAUFFEUR'].mask(df_filtre['DRV'].duplicated())
            #df_filtre['TOTAL SIM+CHAUFFEUR'] = None

    # Fusionner pour ajouter la colonne KABBU
    df_filtre = df_filtre.merge(vto_df[["LOGIN", "KABBU"]], how="left")

            

    # 👉 Ajouter les lignes de total après chaque DRV
    df_with_totals = pd.DataFrame(columns=df_filtre.columns)

    for drv, group in df_filtre.groupby('DRV'):
        df_with_totals = pd.concat([df_with_totals, group], ignore_index=True)

        total_paiement = group['PAIEMENT_SIM'].sum()
        total_general = group['PAIEMENT_SIM'].sum()
        row_total = {
                    'DRV': f"{drv}",
                    'PVT': "TOTAL PVT",
                    'PAIEMENT_SIM': total_paiement ,
                    
                        }
        df_with_totals = pd.concat([df_with_totals, pd.DataFrame([row_total])], ignore_index=True)
    
    st.dataframe(df_with_totals)  

    # Affichage du tableau simplifié
    cols_affichage = ['DRV', 'PVT', 'PRENOM_VENDEUR', 'NOM_VENDEUR','KABBU','REALISATION', 'OBJECTIF', "TAUX D'ATTEINTE", 'SI 100% ATTEINT', 'PAIEMENT', 'PAIEMENT CHAUFFEUR', 'TOTAL SIM+CHAUFFEUR']



    # Harmonisation colonnes OM
    # ✅ Charger logins depuis fichier VTO
    vto_df = load_vto()
    logins_concernes = vto_df["LOGIN"].astype(str).tolist()
    details = ["En Cours-Identification", "Identifie", "Identifie Photo"]

    # ✅ Nettoyage des colonnes
    
    df_om['LOGIN'] = df_om['LOGIN'].astype(str)
            
    df_om['NOM_VENDEUR'] = df_om['NOM_VENDEUR'].astype(str).str.strip().str.upper()
    df_om['PRENOM_VENDEUR'] = df_om['PRENOM_VENDEUR'].astype(str).str.strip().str.upper()

    # 🔍 Filtrage
    df_filtre_om = df_om[df_om['LOGIN'].isin(logins_concernes)]

    st.success("✅ Fichier filtré avec succès !")
    st.write("📊 Ventes LOUMA mensuels :", df_filtre_om.shape[0], "lignes")
    st.dataframe(df_filtre_om)

            
    #Définition des colonnes pour les paiements      
    df_filtre_om['OBJECTIF OM'] = 120
    df_filtre_om["TAUX D'ATTEINTE OM"] = (df_filtre_om['REALISATION_OM'] / df_filtre_om['OBJECTIF OM']).apply(lambda x: f"{round(x*100)}%")
    df_filtre_om['SI 100% ATTEINT OM'] = 25000
    df_filtre_om['PAIEMENT_OM'] = df_filtre_om['REALISATION_OM'].apply(lambda x: 25000 if x >= 120 else round((x/120)*25000))
    #df_filtre['PAIEMENT CHAUFFEUR'] = 150000
           
    # Fusionner pour ajouter la colonne KABBU
    df_filtre_om = df_filtre_om.merge(vto_df[["LOGIN", "DRV", "PVT"]], how="left")
    st.dataframe(df_filtre_om)


    # 👉 Ajouter les lignes de total après chaque DRV
    df_with_totals_om = pd.DataFrame(columns=df_filtre_om.columns)

    for drv, group in df_filtre_om.groupby('DRV'):
                df_with_totals_om = pd.concat([df_with_totals_om, group], ignore_index=True)

                total_paiement_om = group['PAIEMENT_OM'].sum()
                total_general = group['PAIEMENT_OM'].sum()
                row_total = {
                    'DRV': f"{drv}",
                    'PVT': "TOTAL PVT",
                    'PAIEMENT_OM': total_paiement_om ,
                    
                        }
                df_with_totals_om = pd.concat([df_with_totals_om, pd.DataFrame([row_total])], ignore_index=True)


    st.dataframe(df_with_totals_om)







    # === Fusionner sur KABBU et vendeur
    df_final = pd.merge(
        df_with_totals,
        df_with_totals_om[["LOGIN", "PRENOM_VENDEUR", "NOM_VENDEUR", "REALISATION_OM", "OBJECTIF OM","TAUX D'ATTEINTE OM", "SI 100% ATTEINT OM", "PAIEMENT_OM"]],
        on=["LOGIN", "PRENOM_VENDEUR", "NOM_VENDEUR"],
        how="outer"
    )
 
    st.dataframe(df_final)

    # 👉 Ajouter les lignes de total après chaque DRV
    df_final_with_totals = pd.DataFrame(columns=df_final.columns)

    for drv, group in df_final.groupby('DRV'):
                df_final_with_totals = pd.concat([df_final_with_totals, group], ignore_index=True)

                total_paiement_om = group['PAIEMENT_OM'].sum()
                total_paiement_sim = group['PAIEMENT_SIM'].sum()
                #total_general = group['PAIEMENT_OM'].sum()
                row_total = {
                    'DRV': f"{drv}",
                    'PVT': "TOTAL PVT",
                    'PAIEMENT_OM': total_paiement_om ,
                    'PAIEMENT_SIM': total_paiement_sim 
                    
                        }
                df_final_with_totals = pd.concat([df_final_with_totals, pd.DataFrame([row_total])], ignore_index=True)


    st.dataframe(df_final_with_totals)

    #--------------------------------
    df_test = pd.merge(
        df_filtre,
        df_filtre_om[["LOGIN", "PRENOM_VENDEUR", "NOM_VENDEUR", "REALISATION_OM", "OBJECTIF OM","TAUX D'ATTEINTE OM", "SI 100% ATTEINT OM", "PAIEMENT_OM"]],
        on=["LOGIN", "PRENOM_VENDEUR", "NOM_VENDEUR"],
        how="outer"
    )
    df_test["PAIEMENT CHAUFFEUR"] = None
    df_test["PAIEMENT SIM + OM + CHAUFFEUR"] = None

    # 👉 Ajouter les lignes de total après chaque DRV
    df_test_with_totals = pd.DataFrame(columns=df_test.columns)

    for drv, group_drv in df_test.groupby('DRV'):
        for pvt, group_pvt in group_drv.groupby('PVT'):
                    df_test_with_totals = pd.concat([df_test_with_totals, group_pvt], ignore_index=True)

                    total_paiement_om = group_pvt['PAIEMENT_OM'].sum()
                    total_paiement_sim = group_pvt['PAIEMENT_SIM'].sum()
                    chauffeur = 100000
                    total_pvt = chauffeur + total_paiement_om + total_paiement_sim

                    #total_general = group['PAIEMENT_OM'].sum()
                    row_total = {
                        
                        'PVT': "TOTAL PVT",
                        'PAIEMENT_OM': total_paiement_om ,
                        'PAIEMENT_SIM': total_paiement_sim,
                        'PAIEMENT CHAUFFEUR' : chauffeur,
                        'PAIEMENT SIM + OM + CHAUFFEUR' : total_pvt
                        
                            }
                    df_test_with_totals = pd.concat([df_test_with_totals, pd.DataFrame([row_total])], ignore_index=True)

        total_paiement_om_drv = group_drv['PAIEMENT_OM'].sum()
        total_paiement_sim_drv = group_drv['PAIEMENT_SIM'].sum()
        chauffeur_drv = 200000
        total = chauffeur_drv + total_paiement_om_drv + total_paiement_sim_drv
        #total_general = group['PAIEMENT_OM'].sum()
        row_total_drv = {
                        'DRV': f"{drv}",
                        'PVT': "TOTAL",
                        'PAIEMENT_OM': total_paiement_om_drv ,
                        'PAIEMENT_SIM': total_paiement_sim_drv,
                        'PAIEMENT CHAUFFEUR' : chauffeur_drv,
                        'PAIEMENT SIM + OM + CHAUFFEUR' : total
                            
                            }
        df_test_with_totals = pd.concat([df_test_with_totals, pd.DataFrame([row_total_drv])], ignore_index=True)

    st.dataframe(df_test_with_totals)




    # === Calcul des paiements
    
    #---
    df_test["MONTANT"] = df_test["PAIEMENT_SIM"] + df_test["PAIEMENT_OM"]

    # === Résumé par PVT
    df_par_pvt = df_test.groupby(["DRV", "PVT"]).agg({ 'MONTANT':'sum' }).reset_index()
    df_par_pvt["MONTANT"] = df_par_pvt["MONTANT"] + 100000
    df_par_pvt["GAIN PVT (5%)"] = df_par_pvt["MONTANT"] * 0.05
    df_par_pvt["TOTAL GENERAL"] = df_par_pvt["MONTANT"] + df_par_pvt["GAIN PVT (5%)"]

    # Ajouter ND PARTENAIRE (ex : numéro de téléphone du PVT)
    df_par_pvt["ND PARTENAIRE"] = ""  # 👉 tu pourras l'alimenter depuis ton fichier VTO

    # Réorganisation colonnes
    df_par_pvt = df_par_pvt[["DRV", "PVT", "ND PARTENAIRE", "MONTANT", "GAIN PVT (5%)", "TOTAL GENERAL"]]

    st.subheader("📊 Résumé par PVT")
    st.dataframe(df_par_pvt)

    # === Export Excel
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df_test_with_totals.to_excel(writer, sheet_name="Détails Paiement", index=False)
        df_par_pvt.to_excel(writer, sheet_name="Paiement par PVT", index=False)
    buffer.seek(0)

    st.download_button(
        label="📥 Télécharger le fichier Paiement Mensuel",
        data=buffer,
        file_name="paiement_mensuel_global.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
