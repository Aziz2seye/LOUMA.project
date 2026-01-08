import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from pathlib import Path
import os

# Import des fonctions locales (assurez-vous que le chemin est correct)
try:
    from utils import load_vto, load_pvt
except ImportError:
    st.error("Le module utils est introuvable. Vérifiez l'organisation de vos fichiers.")

# Configuration
st.set_page_config(page_title="LOUMA - Paiement Mensuel", layout="wide")

st.title("💰 Paiement Mensuel Global - Maquette Loumas")
st.markdown("---")

# Upload des fichiers
col1, col2 = st.columns(2)
with col1:
    file_sim = st.file_uploader("📥 Fichier SIM", type=["xlsx", "csv"])
with col2:
    file_om = st.file_uploader("📥 Fichier OM", type=["xlsx", "csv"])

if file_sim and file_om:
    # --- CHARGEMENT ET NETTOYAGE VTO ---
    vto_df = load_vto()
    vto_df['LOGIN'] = vto_df['LOGIN'].astype(str).str.strip().str.lower()
    logins_concernes = vto_df["LOGIN"].tolist()
    details_id = ["En Cours-Identification", "Identifie", "Identifie Photo"]

    # --- TRAITEMENT SIM ---
    df_sim_raw = pd.read_csv(file_sim, sep=";") if file_sim.name.endswith(".csv") else pd.read_excel(file_sim)
    df_sim_raw['LOGIN_VENDEUR'] = df_sim_raw['LOGIN_VENDEUR'].astype(str).str.strip().str.lower()

    # Filtrage et Renommage
    df_sim = df_sim_raw[df_sim_raw['LOGIN_VENDEUR'].isin(logins_concernes) & df_sim_raw['ETAT_IDENTIFICATION'].isin(details_id)].copy()
    df_sim = df_sim.rename(columns={'MSISDN': 'R_SIM', 'ACCUEIL_VENDEUR': 'PVT', 'LOGIN_VENDEUR': 'LOGIN', 'AGENCE_VENDEUR': 'DRV'})

    # Agrégation SIM
    df_sim = df_sim.groupby(['DRV', 'PVT', 'NOM_VENDEUR', 'PRENOM_VENDEUR', 'LOGIN']).agg({'R_SIM': 'count'}).reset_index()
    df_sim = df_sim.sort_values('R_SIM', ascending=False).drop_duplicates(subset=['LOGIN'], keep='first')

    # --- TRAITEMENT OM ---
    df_om_raw = pd.read_csv(file_om, sep=";") if file_om.name.endswith(".csv") else pd.read_excel(file_om)
    df_om_raw['LOGIN'] = df_om_raw['LOGIN'].astype(str).str.strip().str.lower()
    df_om = df_om_raw[df_om_raw['LOGIN'].isin(logins_concernes)].copy()
    df_om = df_om.sort_values('REALISATION_OM', ascending=False).drop_duplicates(subset=['LOGIN'], keep='first')

    # --- FUSION (MERGE) ---
    df_final = pd.merge(df_sim, df_om[['LOGIN', 'REALISATION_OM']], on='LOGIN', how='outer')
    df_final = df_final.merge(vto_df[['LOGIN', 'KABBU', 'DRV', 'PVT']], on='LOGIN', how='left', suffixes=('', '_vto'))

    # Nettoyage après fusion
    df_final['DRV'] = df_final['DRV'].fillna(df_final['DRV_vto'])
    df_final['PVT'] = df_final['PVT'].fillna(df_final['PVT_vto'])
    df_final = df_final.drop(columns=['DRV_vto', 'PVT_vto'])

    # Normalisation DRV
    mapping_drv = {
        "DV-DRV2_DIRECTION REGIONALE DES VENTES DAKAR 2": "DR2",
        "DV-DRVS_DIRECTION REGIONALE DES VENTES SUD": "DR SUD",
        "DV-DRVSE_DIRECTION REGIONALE DES VENTES SUD-EST": "SUD EST",
        "DV-DRVN_DIRECTION REGIONALE DES VENTES NORD": "DR NORD",
        "DV-DRVC_DIRECTION REGIONALE DES VENTES CENTRE": "DR CENTRE",
        "DV-DRVE_DIRECTION REGIONALE DES VENTES EST": "DR EST"
    }
    df_final["DRV"] = df_final["DRV"].replace(mapping_drv)

    # --- CALCULS DES COLONNES MAQUETTE ---
    df_final['O_SIM'] = 240
    df_final['RO_SIM'] = (df_final['R_SIM'].fillna(0) / 240).apply(lambda x: f"{round(x*100)}%")
    df_final['GainMax_SIM'] = 75000
    df_final['Gain_SIM'] = df_final['R_SIM'].apply(lambda x: 75000 if x >= 240 else round((x/240)*75000) if pd.notnull(x) else 0)

    df_final['O_OM'] = 120
    df_final['RO_OM'] = (df_final['REALISATION_OM'].fillna(0) / 120).apply(lambda x: f"{round(x*100)}%")
    df_final['GainMax_OM'] = 25000
    df_final['Gain_OM'] = df_final['REALISATION_OM'].apply(lambda x: 25000 if x >= 120 else round((x/120)*25000) if pd.notnull(x) else 0)

    # --- CONSTRUCTION DU TABLEAU AVEC TOTAL PVT & DR ---
    cols_order = ['DRV', 'PVT', 'PRENOM_VENDEUR', 'NOM_VENDEUR', 'LOGIN', 'KABBU', 'R_SIM', 'O_SIM', 'RO_SIM', 'GainMax_SIM', 'Gain_SIM', 'REALISATION_OM', 'O_OM', 'RO_OM', 'GainMax_OM', 'Gain_OM', 'Gain Chauffeur', 'Total']

    output_rows = []
    for drv, group_drv in df_final.groupby('DRV'):
        for pvt, group_pvt in group_drv.groupby('PVT'):
            for _, row in group_pvt.iterrows():
                output_rows.append(row.to_dict())

            # Ligne TOTAL PVT
            total_pvt_sim = group_pvt['Gain_SIM'].sum()
            total_pvt_om = group_pvt['Gain_OM'].sum()
            output_rows.append({
                'PVT': 'TOTAL PVT',
                'Gain_SIM': total_pvt_sim,
                'Gain_OM': total_pvt_om,
                'Gain Chauffeur': 100000,
                'Total': total_pvt_sim + total_pvt_om + 100000
            })

        # Ligne TOTAL DR
        total_dr_sim = group_drv['Gain_SIM'].sum()
        total_dr_om = group_drv['Gain_OM'].sum()
        output_rows.append({
            'DRV': f'TOTAL {drv}',
            'Gain_SIM': total_dr_sim,
            'Gain_OM': total_dr_om,
            'Gain Chauffeur': 200000, # Selon votre logique
            'Total': total_dr_sim + total_dr_om + 200000
        })

    df_details_export = pd.DataFrame(output_rows)

    # --- FEUILLE RÉSUMÉ (INCHANGÉE) ---
    df_resume = df_final.groupby(["DRV", "PVT"]).agg({'Gain_SIM': 'sum', 'Gain_OM': 'sum'}).reset_index()
    df_resume["MONTANT"] = df_resume["Gain_SIM"] + df_resume["Gain_OM"] + 100000
    df_resume["GAIN PVT (5%)"] = df_resume["MONTANT"] * 0.05
    df_resume["TOTAL GENERAL"] = df_resume["MONTANT"] + df_resume["GAIN PVT (5%)"]

    pvt_info = load_pvt()
    df_resume = df_resume.merge(pvt_info[["PVT", "CONTACT"]], on="PVT", how="left")
    df_resume = df_resume[["DRV", "PVT", "CONTACT", "MONTANT", "GAIN PVT (5%)", "TOTAL GENERAL"]]

    # --- EXPORT ET FORMATAGE EXCEL ---
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_details_export.to_excel(writer, sheet_name='Détails Paiement', index=False)
        df_resume.to_excel(writer, sheet_name='Résumé PVT', index=False)

    # Recharger pour le style
    buffer.seek(0)
    wb = load_workbook(buffer)
    ws1 = wb['Détails Paiement']

    # Application des styles (Maquette Louma)
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    pvt_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    dr_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")

    for cell in ws1[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill

    for row in ws1.iter_rows(min_row=2):
        if row[1].value == 'TOTAL PVT':
            for cell in row: cell.fill = pvt_fill
        elif str(row[0].value).startswith('TOTAL DR'):
            for cell in row: cell.fill = dr_fill

    final_buffer = BytesIO()
    wb.save(final_buffer)

    st.success("✅ Fichier adapté à la maquette Loumas prêt !")
    st.download_button("📥 Télécharger le fichier pour Manager", data=final_buffer.getvalue(), file_name="paiement_ventes_loumas_automatise.xlsx")