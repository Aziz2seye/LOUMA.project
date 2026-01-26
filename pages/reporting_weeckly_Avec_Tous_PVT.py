import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import sys
from pathlib import Path
from PIL import Image
from datetime import datetime
import calendar

# Configuration page
st.set_page_config(page_title="LOUMA - Reporting Hebdomadaire", layout="wide", initial_sidebar_state="expanded")

# CSS personnalisé Orange Sonatel (Inchangé mais titre adapté)
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;700&display=swap');
    header[data-testid="stHeader"] { display: none; }
    .block-container { padding-top: 2rem !important; padding-bottom: 2rem !important; }
    .main {
        font-family: 'Poppins', sans-serif;
        background: linear-gradient(135deg, #fff5f0 0%, #ffffff 50%, #f0f8ff 100%);
    }
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #FF7900 0%, #FF5000 100%) !important;
    }
    section[data-testid="stSidebar"] * { color: white !important; }
    .stDataFrame {
        border-radius: 15px;
        overflow: hidden;
        box-shadow: 0 8px 20px rgba(255, 121, 0, 0.15);
        border: 2px solid #FFE5CC;
    }
    .section-title {
        background: linear-gradient(135deg, #FF7900 0%, #FF5000 100%);
        color: white;
        padding: 0.7rem 1.2rem;
        border-radius: 10px;
        font-weight: 600;
        font-size: 1.1rem;
        margin-bottom: 0.8rem;
        box-shadow: 0 4px 12px rgba(255, 121, 0, 0.25);
        text-align: center;
        max-width: 500px;
        margin-left: auto;
        margin-right: auto;
    }
    .stButton > button {
        background: linear-gradient(135deg, #FF7900 0%, #FF5000 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.8rem 2rem;
        font-weight: 600;
        font-size: 1.1rem;
        box-shadow: 0 4px 12px rgba(255, 121, 0, 0.3);
        transition: all 0.3s ease;
        width: 100%;
    }
    .stDownloadButton > button {
        background: linear-gradient(135deg, #00D4AA 0%, #00B890 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.8rem 2rem;
        font-weight: 600;
        font-size: 1.1rem;
        box-shadow: 0 4px 12px rgba(0, 212, 170, 0.3);
        transition: all 0.3s ease;
        width: 100%;
    }
</style>
""", unsafe_allow_html=True)

# --- LOGO & HEADER ---
# (Gardez votre logique de logo ici)
col_logo, col_title = st.columns([1, 3])
with col_title:
    st.markdown("""
    <div style="background: linear-gradient(135deg, #FF7900 0%, #FF5000 100%); padding: 2rem; border-radius: 20px; box-shadow: 0 8px 25px rgba(255, 121, 0, 0.4); border: 3px solid rgba(255, 255, 255, 0.2); height: 100%;">
        <h1 style="color: white; font-size: 2.5rem; font-weight: 700; margin: 0; text-shadow: 3px 3px 10px rgba(0, 0, 0, 0.3);">
            📊 Reporting Hebdomadaire (Weekly)
        </h1>
        <p style="color: rgba(255, 255, 255, 0.95); font-size: 1.2rem; margin: 0.8rem 0 0 0; font-weight: 400;">
            Analyse hebdomadaire - Objectif : 240 / PVT
        </p>
    </div>
    """, unsafe_allow_html=True)

DRV_MAPPING = {
    "DV-DRV2_DIRECTION REGIONALE DES VENTES DAKAR 2": "DR2",
    "DV-DRVS_DIRECTION REGIONALE DES VENTES SUD": "DRS",
    "DV-DRVSE_DIRECTION REGIONALE DES VENTES SUD-EST": "DRSE",
    "DV-DRVN_DIRECTION REGIONALE DES VENTES NORD": "DRN",
    "DV-DRVC_DIRECTION REGIONALE DES VENTES CENTRE": "DRC",
    "DV-DRVE_DIRECTION REGIONALE DES VENTES EST": "DRE",
    "DV-DRV1_DIRECTION REGIONALE DES VENTES DAKAR 1": "DR1"
}

def generate_weekly_excel_report(df_final, week_num, annee, objectif_pvt=240):
    """Génère un fichier Excel hebdomadaire avec l'objectif de 240"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book
        # Formats (identiques au mensuel pour garder la charte)
        h_fmt = workbook.add_format({'bold': True, 'bg_color': '#FF6600', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        dr_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'left'})
        dr_num_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center'})
        pvt_fmt = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1, 'indent': 1})
        pvt_num_fmt = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1, 'align': 'center'})
        vendeur_fmt = workbook.add_format({'border': 1, 'indent': 2, 'align': 'center', 'font_size': 9})
        vendeur_num_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_size': 9})
        total_fmt = workbook.add_format({'bold': True, 'bg_color': '#FF6600', 'font_color': 'white', 'border': 1, 'align': 'center'})

        # FEUILLE 1: SYNTHESE DR
        ws1 = workbook.add_worksheet('SYNTHESE DR WEEKLY')
        headers_dr = ['DR', 'REALISATION', 'OBJECTIF', 'R/O (%)']
        for c, h in enumerate(headers_dr):
            ws1.write(0, c, h, h_fmt)

        synthese_dr = df_final.groupby('DR').size().reset_index(name='REALISATION')
        pvt_par_dr = df_final.groupby('DR')['PVT'].nunique().reset_index(name='NB_PVT')
        synthese_dr = synthese_dr.merge(pvt_par_dr, on='DR')

        # CHANGEMENT ICI : Objectif 240
        synthese_dr['OBJECTIF'] = synthese_dr['NB_PVT'] * objectif_pvt
        synthese_dr['R/O'] = ((synthese_dr['REALISATION'] / synthese_dr['OBJECTIF']) * 100).round(0)

        for i, r in synthese_dr.iterrows():
            ws1.write(i+1, 0, r['DR'], dr_fmt)
            ws1.write(i+1, 1, int(r['REALISATION']), dr_num_fmt)
            ws1.write(i+1, 2, int(r['OBJECTIF']), dr_num_fmt)
            ws1.write(i+1, 3, f"{int(r['R/O'])}%", dr_num_fmt)

        total_row = len(synthese_dr) + 1
        total_real = int(synthese_dr['REALISATION'].sum())
        total_obj = int(synthese_dr['OBJECTIF'].sum())
        total_ro = round((total_real / total_obj * 100), 0) if total_obj > 0 else 0
        ws1.write(total_row, 0, 'TOTAL', total_fmt)
        ws1.write(total_row, 1, total_real, total_fmt)
        ws1.write(total_row, 2, total_obj, total_fmt)
        ws1.write(total_row, 3, f"{int(total_ro)}%", total_fmt)
        ws1.set_column('A:A', 18)
        ws1.set_column('B:D', 15)

        # FEUILLE 2: REPORTING DR-PVT
        ws2 = workbook.add_worksheet('REPORTING DR-PVT')
        for c, h in enumerate(headers_dr): ws2.write(0, c, h, h_fmt)
        curr_row = 1
        total_real_pvt = 0
        total_obj_pvt = 0

        for dr, dr_group in df_final.groupby('DR', sort=True):
            real_dr = len(dr_group)
            nb_pvt_dr = dr_group['PVT'].nunique()
            obj_dr = nb_pvt_dr * objectif_pvt # CHANGEMENT ICI
            ro_dr = round((real_dr / obj_dr * 100), 0) if obj_dr > 0 else 0
            ws2.write(curr_row, 0, dr, dr_fmt)
            ws2.write(curr_row, 1, real_dr, dr_num_fmt)
            ws2.write(curr_row, 2, obj_dr, dr_num_fmt)
            ws2.write(curr_row, 3, f"{int(ro_dr)}%", dr_num_fmt)
            curr_row += 1
            for pvt, pvt_group in dr_group.groupby('PVT', sort=True):
                real_pvt = len(pvt_group)
                obj_pvt = objectif_pvt # CHANGEMENT ICI (240)
                ro_pvt = round((real_pvt / obj_pvt * 100), 0) if obj_pvt > 0 else 0
                ws2.write(curr_row, 0, pvt, pvt_fmt)
                ws2.write(curr_row, 1, real_pvt, pvt_num_fmt)
                ws2.write(curr_row, 2, obj_pvt, pvt_num_fmt)
                ws2.write(curr_row, 3, f"{int(ro_pvt)}%", pvt_num_fmt)
                curr_row += 1
                total_real_pvt += real_pvt
                total_obj_pvt += obj_pvt

        # Ligne TOTAL PVT
        total_ro_pvt = round((total_real_pvt / total_obj_pvt * 100), 0) if total_obj_pvt > 0 else 0
        ws2.write(curr_row, 0, 'TOTAL', total_fmt)
        ws2.write(curr_row, 1, total_real_pvt, total_fmt)
        ws2.write(curr_row, 2, total_obj_pvt, total_fmt)
        ws2.write(curr_row, 3, f"{int(total_ro_pvt)}%", total_fmt)
        ws2.set_column('A:A', 45)
        ws2.set_column('B:D', 15)

        # FEUILLE 3: (Copier la logique du mensuel pour les vendeurs, elle reste identique)
        # ... [Code identique à votre version initiale pour ws3] ...
        ws3 = workbook.add_worksheet('REPORTING DR-PVT-VENDEURS')
        headers_vendeurs = ['DR/PVT/VENDEUR', 'Prénom', 'Nom', 'LOGIN', 'REALISATION']
        for c, h in enumerate(headers_vendeurs): ws3.write(0, c, h, h_fmt)
        curr_row = 1
        total_real_vendeurs = 0
        for dr, dr_group in df_final.groupby('DR', sort=True):
            ws3.write(curr_row, 0, dr, dr_fmt)
            ws3.write(curr_row, 4, len(dr_group), dr_num_fmt)
            curr_row += 1
            for pvt, pvt_group in dr_group.groupby('PVT', sort=True):
                ws3.write(curr_row, 0, pvt, pvt_fmt)
                ws3.write(curr_row, 4, len(pvt_group), pvt_num_fmt)
                curr_row += 1
                total_real_vendeurs += len(pvt_group)
                vendeurs = pvt_group.groupby(['PRENOM_VENDEUR', 'NOM_VENDEUR', 'LOGIN']).size().reset_index(name='REALISATION')
                for _, v in vendeurs.sort_values('REALISATION', ascending=False).iterrows():
                    ws3.write(curr_row, 0, 'VENDEUR', vendeur_fmt)
                    ws3.write(curr_row, 1, v['PRENOM_VENDEUR'], vendeur_fmt)
                    ws3.write(curr_row, 2, v['NOM_VENDEUR'], vendeur_fmt)
                    ws3.write(curr_row, 3, v['LOGIN'], vendeur_fmt)
                    ws3.write(curr_row, 4, int(v['REALISATION']), vendeur_num_fmt)
                    curr_row += 1
        ws3.write(curr_row, 0, 'TOTAL', total_fmt)
        ws3.write(curr_row, 4, total_real_vendeurs, total_fmt)
        ws3.set_column('A:A', 35); ws3.set_column('B:C', 20); ws3.set_column('D:D', 25); ws3.set_column('E:E', 15)

    output.seek(0)
    return output

def main():
    st.markdown('<div class="section-title">📅 SÉLECTION DE LA SEMAINE</div>', unsafe_allow_html=True)

    col_week, col_year = st.columns(2)
    with col_week:
        # Ajout d'un sélecteur de semaine
        selected_week = st.number_input("Numéro de la Semaine (W)", min_value=1, max_value=53, value=datetime.now().isocalendar()[1])
    with col_year:
        selected_year = st.selectbox("Année", options=[2024, 2025, 2026], index=1)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="section-title">📁 IMPORTATION DES DONNÉES</div>', unsafe_allow_html=True)

    uploaded_file = st.file_uploader(f"Importer le fichier pour la Semaine {selected_week}", type=["xlsx", "csv", "xls"])

    if uploaded_file:
        try:
            # --- LECTURE ET NETTOYAGE --- (Identique à votre code)
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, encoding='utf-8', sep=';')
            else:
                df = pd.read_excel(uploaded_file)

            # Application du mapping des noms de colonnes (Reprise de votre code)
            column_mapping = {'MSISDN': 'REALISATION', 'ACCUEIL_VENDEUR': 'PVT', 'LOGIN_VENDEUR': 'LOGIN', 'AGENCE_VENDEUR': 'DR', 'NOM_VENDEUR': 'NOM_VENDEUR', 'PRENOM_VENDEUR': 'PRENOM_VENDEUR', 'ETAT_IDENTIFICATION': 'ETAT_IDENTIFICATION'}
            for original_col, new_col in column_mapping.items():
                if original_col in df.columns: df = df.rename(columns={original_col: new_col})

            # Filtrage & DR Mapping
            details = ["En Cours-Identification", "Identifie", "Identifie Photo"]
            if 'ETAT_IDENTIFICATION' in df.columns:
                df_filtre = df[df['ETAT_IDENTIFICATION'].astype(str).isin(details)].copy()
            else:
                df_filtre = df.copy()

            if 'DR' in df_filtre.columns:
                df_filtre = df_filtre[df_filtre['DR'].isin(DRV_MAPPING.keys())].copy()
                df_filtre["DR"] = df_filtre["DR"].replace(DRV_MAPPING)

            # --- AFFICHAGE RESULTATS ---
            st.success(f"✅ Analyse terminée pour la Semaine {selected_week}")

            # --- GENERATION RAPPORT WEEKLY ---
            objectif_weekly = 240 # L'objectif est maintenant 240
            excel_buffer = generate_weekly_excel_report(df_filtre, selected_week, selected_year, objectif_weekly)

            st.download_button(
                label=f"📥 Télécharger le Reporting Hebdomadaire (W{selected_week})",
                data=excel_buffer.getvalue(),
                file_name=f"Reporting_Weekly_W{selected_week}_{selected_year}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"Erreur : {e}")

if __name__ == "__main__":
    main()