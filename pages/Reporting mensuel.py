import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from PIL import Image
from pathlib import Path

from utils import load_vto, load_pvt

# Configuration
st.set_page_config(page_title="LOUMA - Paiement Mensuel", layout="wide")

# CSS minimaliste
st.markdown("""
<style>
    .main { font-family: 'Arial', sans-serif; }
    .stButton > button {
        background: #FF7900;
        color: white;
        border-radius: 8px;
        padding: 0.5rem 1.5rem;
        font-weight: 600;
        width: 100%;
    }
    .stButton > button:hover {
        background: #FF5000;
    }
</style>
""", unsafe_allow_html=True)

# Logo et titre
logo_path = Path(__file__).parent.parent / "assets" / "logo sonatel.png"
if logo_path.exists():
    st.image(str(logo_path), width=200)

st.title("💰 Paiement Mensuel Global - SIM + OM")
st.markdown("---")

# Upload des fichiers
col1, col2 = st.columns(2)
with col1:
    file_sim = st.file_uploader("📥 Fichier SIM", type=["xlsx", "csv"])
with col2:
    file_om = st.file_uploader("📥 Fichier OM", type=["xlsx", "csv"])

if file_sim and file_om:

    # === TRAITEMENT SIM ===
    if file_sim.name.endswith(".csv"):
        df_sim = pd.read_csv(file_sim, sep=";", encoding="utf-8")
    else:
        df_sim = pd.read_excel(file_sim)

    # === TRAITEMENT OM ===
    if file_om.name.endswith(".csv"):
        df_om = pd.read_csv(file_om, sep=";", encoding="utf-8")
    else:
        df_om = pd.read_excel(file_om)

    # === CHARGEMENT VTO ===
    vto_df = load_vto()
    vto_df['LOGIN'] = vto_df['LOGIN'].astype(str).str.strip().str.lower()
    logins_concernes = vto_df["LOGIN"].tolist()
    details = ["En Cours-Identification", "Identifie", "Identifie Photo"]

    # === TRAITEMENT SIM ===
    df_sim['LOGIN_VENDEUR'] = df_sim['LOGIN_VENDEUR'].astype(str).str.strip().str.lower()
    df = df_sim.rename(columns={
        'MSISDN': 'REALISATION_SIM',
        'ACCUEIL_VENDEUR': 'PVT',
        'LOGIN_VENDEUR': 'LOGIN',
        'AGENCE_VENDEUR': 'DRV'
    })

    df['LOGIN'] = df['LOGIN'].astype(str)
    df['DRV'] = df['DRV'].astype(str).str.strip().str.upper()
    df['NOM_VENDEUR'] = df['NOM_VENDEUR'].astype(str).str.strip().str.upper()
    df['PRENOM_VENDEUR'] = df['PRENOM_VENDEUR'].astype(str).str.strip().str.upper()

    df_filtre = df[df['LOGIN'].isin(logins_concernes) & df['ETAT_IDENTIFICATION'].isin(details)]

    df_filtre["DRV"] = df_filtre["DRV"].replace({
        "DV-DRV2_DIRECTION REGIONALE DES VENTES DAKAR 2": "DR2",
        "DV-DRVS_DIRECTION REGIONALE DES VENTES SUD": "DR SUD",
        "DV-DRVSE_DIRECTION REGIONALE DES VENTES SUD-EST": "SUD EST",
        "DV-DRVN_DIRECTION REGIONALE DES VENTES NORD": "DR NORD",
        "DV-DRVC_DIRECTION REGIONALE DES VENTES CENTRE": "DR CENTRE",
        "DV-DRVE_DIRECTION REGIONALE DES VENTES EST": "DR EST"
    })

    df_filtre = df_filtre.groupby(['DRV', 'PVT', 'PRENOM_VENDEUR', 'NOM_VENDEUR', 'LOGIN']).agg({
        'REALISATION_SIM': 'count'
    }).reset_index()

    # DIAGNOSTIC : Afficher les doublons AVANT traitement
    duplicates_check = df_filtre[df_filtre.duplicated(subset=['LOGIN'], keep=False)].sort_values(['LOGIN', 'REALISATION_SIM'], ascending=[True, False])

    if not duplicates_check.empty:
        st.warning(f"⚠️ {duplicates_check['LOGIN'].nunique()} LOGIN dupliqués détectés dans SIM")

    # DÉDUPLICATION
    nb_avant_dedup = len(df_filtre)
    df_filtre = df_filtre.sort_values('REALISATION_SIM', ascending=False)
    df_filtre = df_filtre.drop_duplicates(subset=['LOGIN'], keep='first')
    df_filtre = df_filtre.sort_values(['DRV', 'PVT'])
    nb_apres_dedup = len(df_filtre)

    df_filtre['OBJECTIF SIM'] = 240
    df_filtre["TAUX_SIM"] = (df_filtre['REALISATION_SIM'] / 240 * 100).apply(lambda x: f"{round(x)}%")
    df_filtre['GAIN_MAX_SIM'] = 75000
    df_filtre['GAIN_SIM'] = df_filtre['REALISATION_SIM'].apply(lambda x: 75000 if x >= 240 else round((x/240)*75000))
    df_filtre = df_filtre.merge(vto_df[["LOGIN", "KABBU"]], how="left")

    # === TRAITEMENT OM ===
    df_om['LOGIN'] = df_om['LOGIN'].astype(str).str.strip().str.lower()
    df_om['NOM_VENDEUR'] = df_om['NOM_VENDEUR'].astype(str).str.strip().str.upper()
    df_om['PRENOM_VENDEUR'] = df_om['PRENOM_VENDEUR'].astype(str).str.strip().str.upper()

    df_filtre_om = df_om[df_om['LOGIN'].isin(logins_concernes)].fillna(0)

    # DÉDUPLICATION OM
    duplicates_om = df_filtre_om[df_filtre_om.duplicated(subset=['LOGIN'], keep=False)].sort_values(['LOGIN', 'REALISATION_OM'], ascending=[True, False])

    if not duplicates_om.empty:
        st.warning(f"⚠️ {duplicates_om['LOGIN'].nunique()} LOGIN dupliqués détectés dans OM")

    nb_avant_dedup_om = len(df_filtre_om)
    df_filtre_om = df_filtre_om.sort_values('REALISATION_OM', ascending=False)
    df_filtre_om = df_filtre_om.drop_duplicates(subset=['LOGIN'], keep='first')
    nb_apres_dedup_om = len(df_filtre_om)

    df_filtre_om['OBJECTIF OM'] = 120
    df_filtre_om["TAUX_OM"] = (df_filtre_om['REALISATION_OM'] / 120 * 100).fillna(0).apply(lambda x: f"{round(x)}%")
    df_filtre_om['GAIN_MAX_OM'] = 25000
    df_filtre_om['GAIN_OM'] = df_filtre_om['REALISATION_OM'].apply(lambda x: 25000 if x >= 120 else round((x/120)*25000))
    df_filtre_om = df_filtre_om.merge(vto_df[["LOGIN", "DRV", "PVT"]], how="left")

    # === FUSION SIM + OM ===
    df_test = pd.merge(
        df_filtre,
        df_filtre_om[["LOGIN", "REALISATION_OM", "OBJECTIF OM", "TAUX_OM", "GAIN_MAX_OM", "GAIN_OM"]],
        on=["LOGIN"],
        how="outer"
    )

    # VÉRIFICATION FINALE
    duplicates_final = df_test[df_test.duplicated(subset=['LOGIN'], keep=False)].sort_values('LOGIN')

    if not duplicates_final.empty:
        st.error(f"❌ ATTENTION: {duplicates_final['LOGIN'].nunique()} doublons détectés après le merge!")
        with st.expander("Voir les doublons après merge"):
            st.dataframe(duplicates_final[['LOGIN', 'PVT', 'DRV', 'PRENOM_VENDEUR', 'NOM_VENDEUR', 'REALISATION_SIM', 'REALISATION_OM']])
        df_test = df_test.drop_duplicates(subset=['LOGIN'], keep='first')
        st.warning(f"⚠️ Déduplication forcée appliquée - {len(df_test)} lignes restantes")
    else:
        st.success(f"✅ Aucun doublon après fusion - {len(df_test)} VTO uniques")

    # Remplir les NaN
    df_test['GAIN_SIM'] = df_test['GAIN_SIM'].fillna(0)
    df_test['GAIN_OM'] = df_test['GAIN_OM'].fillna(0)
    df_test['REALISATION_SIM'] = df_test['REALISATION_SIM'].fillna(0)
    df_test['REALISATION_OM'] = df_test['REALISATION_OM'].fillna(0)

    # === RÉORGANISATION DES COLONNES SELON FORMAT MANAGER ===
    df_final = pd.DataFrame({
        'DR': df_test['DRV'],
        'PVT': df_test['PVT'],
        'PRENOM_VTO': df_test['PRENOM_VENDEUR'],
        'NOM_VTO': df_test['NOM_VENDEUR'],
        'LOGIN': df_test['LOGIN'],
        'Numéro Kabbu': df_test['KABBU'],
        'R': df_test['REALISATION_SIM'].astype(int),
        'O': df_test['OBJECTIF SIM'].fillna(240).astype(int),
        'R/O': df_test['TAUX_SIM'],
        'Gain Max': df_test['GAIN_MAX_SIM'].fillna(75000).astype(int),
        'Gain': df_test['GAIN_SIM'].astype(int),
        'R.1': df_test['REALISATION_OM'].astype(int),
        'O.1': df_test['OBJECTIF OM'].fillna(120).astype(int),
        'R/O.1': df_test['TAUX_OM'],
        'Gain Max.1': df_test['GAIN_MAX_OM'].fillna(25000).astype(int),
        'Gain.1': df_test['GAIN_OM'].astype(int),
        'Gain Chauffeur': None,
        'Gain SIM + OM + Chauffeur': None
    })

    # === CRÉATION DES TOTAUX PAR PVT ET DRV ===
    df_with_totals = pd.DataFrame(columns=df_final.columns)

    for drv, group_drv in df_final.groupby('DR'):
        for pvt, group_pvt in group_drv.groupby('PVT'):
            df_with_totals = pd.concat([df_with_totals, group_pvt], ignore_index=True)

            row_total_pvt = {
                'DR': drv,
                'PVT': "TOTAL PVT",
                'PRENOM_VTO': '',
                'NOM_VTO': '',
                'LOGIN': '',
                'Numéro Kabbu': '',
                'R': group_pvt['R'].sum(),
                'O': group_pvt['O'].sum(),
                'R/O': f'{group_pvt["R/O"].apply(lambda x: float(str(x).strip("%")) if pd.notnull(x) else 0).mean():.0f}%',
                'Gain Max': group_pvt['Gain Max'].sum(),
                'Gain': group_pvt['Gain'].sum(),
                'R.1': group_pvt['R.1'].sum(),
                'O.1': group_pvt['O.1'].sum(),
                'R/O.1': f'{group_pvt["R/O.1"].apply(lambda x: float(str(x).replace("%", "")) if pd.notnull(x) else 0).mean():.0f}%',
                'Gain Max.1': group_pvt['Gain Max.1'].sum(),
                'Gain.1': group_pvt['Gain.1'].sum(),
                'Gain Chauffeur': 100000,
                'Gain SIM + OM + Chauffeur': group_pvt['Gain'].sum() + group_pvt['Gain.1'].sum() + 100000
            }
            df_with_totals = pd.concat([df_with_totals, pd.DataFrame([row_total_pvt])], ignore_index=True)

        row_total_drv = {
            'DR': f"TOTAL {drv}",
            'PVT': '',
            'PRENOM_VTO': '',
            'NOM_VTO': '',
            'LOGIN': '',
            'Numéro Kabbu': '',
            'R': '',
            'O': '',
            'R/O': '',
            'Gain Max': '',
            'Gain': group_drv['Gain'].sum(),
            'R.1': '',
            'O.1': '',
            'R/O.1': '',
            'Gain Max.1': '',
            'Gain.1': group_drv['Gain.1'].sum(),
            'Gain Chauffeur': 200000,
            'Gain SIM + OM + Chauffeur': group_drv['Gain'].sum() + group_drv['Gain.1'].sum() + 200000
        }
        df_with_totals = pd.concat([df_with_totals, pd.DataFrame([row_total_drv])], ignore_index=True)

    # === TOTAL GLOBAL ===
    total_global_sim = df_final['Gain'].sum()
    total_global_om = df_final['Gain.1'].sum()
    nb_dr = df_final['DR'].nunique()
    total_chauffeur_global = nb_dr * 200000

    row_total_global = {
        'DR': "TOTAL GÉNÉRAL",
        'PVT': '',
        'PRENOM_VTO': '',
        'NOM_VTO': '',
        'LOGIN': '',
        'Numéro Kabbu': '',
        'R': '',
        'O': '',
        'R/O': '',
        'Gain Max': '',
        'Gain': total_global_sim,
        'R.1': '',
        'O.1': '',
        'R/O.1': '',
        'Gain Max.1': '',
        'Gain.1': total_global_om,
        'Gain Chauffeur': total_chauffeur_global,
        'Gain SIM + OM + Chauffeur': total_global_sim + total_global_om + total_chauffeur_global
    }
    df_with_totals = pd.concat([df_with_totals, pd.DataFrame([row_total_global])], ignore_index=True)

    # === CALCUL RÉSUMÉ PVT ===
    df_final["MONTANT"] = df_final["Gain"] + df_final["Gain.1"]
    df_par_pvt = df_final.groupby(["DR", "PVT"]).agg({'MONTANT': 'sum'}).reset_index()
    df_par_pvt["MONTANT"] = df_par_pvt["MONTANT"] + 100000
    df_par_pvt["GAIN PVT (5%)"] = df_par_pvt["MONTANT"] * 0.05
    df_par_pvt["TOTAL GENERAL"] = df_par_pvt["MONTANT"] + df_par_pvt["GAIN PVT (5%)"]

    pvt_df = load_pvt()
    df_par_pvt = df_par_pvt.merge(pvt_df[["PVT", "CONTACT"]], on="PVT", how="left")
    df_par_pvt = df_par_pvt[["DR", "PVT", "CONTACT", "MONTANT", "GAIN PVT (5%)", "TOTAL GENERAL"]]

    df_par_pvt_display = df_par_pvt.copy()
    df_par_pvt_display.loc[len(df_par_pvt_display)] = [
        'TOTAL', '', '',
        df_par_pvt['MONTANT'].sum(),
        df_par_pvt['GAIN PVT (5%)'].sum(),
        df_par_pvt['TOTAL GENERAL'].sum()
    ]

    # === MESSAGE DE SUCCÈS ===
    st.success(f"✅ Traitement terminé : {df_final['LOGIN'].nunique()} VTO dans {df_final['PVT'].nunique()} PVT")

    # === EXPORT EXCEL ===
    try:
        buffer_output = BytesIO()

        with pd.ExcelWriter(buffer_output, engine='openpyxl') as writer:
            df_with_totals.to_excel(writer, sheet_name='Détails Paiement', index=False)
            df_par_pvt_display.to_excel(writer, sheet_name='Résumé PVT', index=False)

        buffer_output.seek(0)
        wb = load_workbook(buffer_output)

        # Style commun
        header_fill_gray = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        header_fill_blue = PatternFill(start_color="00B0F0", end_color="00B0F0", fill_type="solid")
        header_fill_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        header_font = Font(bold=True, size=11, color="000000")
        header_font_white = Font(bold=True, size=11, color="FFFFFF")
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # Formater "Détails Paiement" avec en-tête sur 2 lignes
        ws1 = wb['Détails Paiement']

        # Ligne 1: Groupes principaux
        ws1.insert_rows(1)

        # FUSIONNER LES COLONNES A à F sur les lignes 1 et 2
        ws1.merge_cells('A1:A2')
        ws1['A1'] = 'DR'
        ws1['A1'].fill = header_fill_gray
        ws1['A1'].font = header_font
        ws1['A1'].alignment = Alignment(horizontal='center', vertical='center')
        ws1['A1'].border = thin_border

        ws1.merge_cells('B1:B2')
        ws1['B1'] = 'PVT'
        ws1['B1'].fill = header_fill_gray
        ws1['B1'].font = header_font
        ws1['B1'].alignment = Alignment(horizontal='center', vertical='center')
        ws1['B1'].border = thin_border

        ws1.merge_cells('C1:C2')
        ws1['C1'] = 'PRENOM_VTO'
        ws1['C1'].fill = header_fill_gray
        ws1['C1'].font = header_font
        ws1['C1'].alignment = Alignment(horizontal='center', vertical='center')
        ws1['C1'].border = thin_border

        ws1.merge_cells('D1:D2')
        ws1['D1'] = 'NOM_VTO'
        ws1['D1'].fill = header_fill_gray
        ws1['D1'].font = header_font
        ws1['D1'].alignment = Alignment(horizontal='center', vertical='center')
        ws1['D1'].border = thin_border

        ws1.merge_cells('E1:E2')
        ws1['E1'] = 'LOGIN'
        ws1['E1'].fill = header_fill_gray
        ws1['E1'].font = header_font
        ws1['E1'].alignment = Alignment(horizontal='center', vertical='center')
        ws1['E1'].border = thin_border

        ws1.merge_cells('F1:F2')
        ws1['F1'] = 'Numéro Kabbu'
        ws1['F1'].fill = header_fill_gray
        ws1['F1'].font = header_font
        ws1['F1'].alignment = Alignment(horizontal='center', vertical='center')
        ws1['F1'].border = thin_border

        ws1.merge_cells('G1:K1')
        ws1['G1'] = 'SIM'
        ws1['G1'].fill = header_fill_blue
        ws1['G1'].font = header_font
        ws1['G1'].alignment = Alignment(horizontal='center', vertical='center')
        ws1['G1'].border = thin_border

        ws1.merge_cells('L1:P1')
        ws1['L1'] = 'OM'
        ws1['L1'].fill = header_fill_yellow
        ws1['L1'].font = header_font
        ws1['L1'].alignment = Alignment(horizontal='center', vertical='center')
        ws1['L1'].border = thin_border

        ws1.merge_cells('Q1:Q2')
        ws1['Q1'] = 'Gain Chauffeur'
        ws1['Q1'].fill = header_fill_gray
        ws1['Q1'].font = header_font
        ws1['Q1'].alignment = Alignment(horizontal='center', vertical='center')
        ws1['Q1'].border = thin_border

        ws1.merge_cells('R1:R2')
        ws1['R1'] = 'Gain SIM + OM + Chauffeur'
        ws1['R1'].fill = header_fill_gray
        ws1['R1'].font = header_font
        ws1['R1'].alignment = Alignment(horizontal='center', vertical='center')
        ws1['R1'].border = thin_border

        # Ligne 2: Headers détaillés pour SIM et OM uniquement (G à P)
        for col_idx, cell in enumerate(ws1[2], start=1):
            if 7 <= col_idx <= 16:  # Colonnes G à P (SIM et OM)
                cell.fill = header_fill_gray
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = thin_border

        # Dictionnaire pour stocker les plages de lignes par DR et par PVT
        dr_ranges = {}
        pvt_ranges = {}
        current_dr = None
        current_pvt = None
        dr_start_row = None
        pvt_start_row = None

        # ✅ Première passe : identifier les plages - SANS inclure TOTAL PVT dans la fusion
        for row_idx in range(3, ws1.max_row + 1):
            pvt_val = ws1.cell(row_idx, 2).value
            dr_val = ws1.cell(row_idx, 1).value

            # Gestion des DR
            if dr_val and not str(dr_val).startswith('TOTAL') and dr_val != 'TOTAL GÉNÉRAL':
                if current_dr != dr_val:
                    if current_dr and dr_start_row:
                        total_dr_row = row_idx - 1
                        while total_dr_row >= dr_start_row and ws1.cell(total_dr_row, 2).value == 'TOTAL PVT':
                            total_dr_row -= 1
                        if ws1.cell(total_dr_row + 1, 2).value == 'TOTAL PVT':
                            total_dr_row += 1
                        dr_ranges[current_dr] = (dr_start_row, total_dr_row)
                    current_dr = dr_val
                    dr_start_row = row_idx
            elif str(dr_val).startswith('TOTAL ') and dr_val != 'TOTAL GÉNÉRAL':
                if current_dr and dr_start_row:
                    dr_ranges[current_dr] = (dr_start_row, row_idx - 1)
                    current_dr = None
                    dr_start_row = None

            # ✅ Gestion des PVT - INCLURE TOTAL PVT dans la plage de fusion
            if pvt_val and not str(dr_val).startswith('TOTAL') and dr_val != 'TOTAL GÉNÉRAL':
                if pvt_val != 'TOTAL PVT':
                    if current_pvt != pvt_val:
                        if current_pvt and pvt_start_row:
                            # Chercher la ligne TOTAL PVT correspondante
                            total_pvt_row = None
                            for search_idx in range(pvt_start_row + 1, ws1.max_row + 1):
                                if ws1.cell(search_idx, 2).value == 'TOTAL PVT':
                                    total_pvt_row = search_idx
                                    break
                                elif ws1.cell(search_idx, 2).value and ws1.cell(search_idx, 2).value != current_pvt:
                                    # On a atteint un nouveau PVT sans trouver TOTAL PVT
                                    total_pvt_row = search_idx - 1
                                    break

                            if total_pvt_row:
                                pvt_ranges[current_pvt] = (pvt_start_row, total_pvt_row)

                        current_pvt = pvt_val
                        pvt_start_row = row_idx

        # ✅ Fermer le dernier PVT s'il existe
        if current_pvt and pvt_start_row:
            total_pvt_row = None
            for search_idx in range(pvt_start_row + 1, ws1.max_row + 1):
                if ws1.cell(search_idx, 2).value == 'TOTAL PVT':
                    total_pvt_row = search_idx
                    break
                elif str(ws1.cell(search_idx, 1).value).startswith('TOTAL '):
                    total_pvt_row = search_idx - 1
                    break

            if total_pvt_row:
                pvt_ranges[current_pvt] = (pvt_start_row, total_pvt_row)

        # Fusionner les cellules DR
        for dr_name, (start, end) in dr_ranges.items():
            if start < end:
                ws1.merge_cells(start_row=start, start_column=1, end_row=end, end_column=1)
                ws1.cell(start, 1).alignment = Alignment(horizontal='center', vertical='center')
                ws1.cell(start, 1).value = dr_name

        # ✅ Fusionner les cellules Gain Chauffeur et Gain Total pour TOUS les PVT (incluant TOTAL PVT)
        for pvt_name, (start, end) in pvt_ranges.items():
            # Fusionner Gain Chauffeur (colonne Q = 17)
            if start <= end:
                ws1.merge_cells(start_row=start, start_column=17, end_row=end, end_column=17)
                ws1.cell(start, 17).alignment = Alignment(horizontal='center', vertical='center')
                ws1.cell(start, 17).value = 100000

                # Fusionner Gain SIM + OM + Chauffeur (colonne R = 18)
                ws1.merge_cells(start_row=start, start_column=18, end_row=end, end_column=18)
                ws1.cell(start, 18).alignment = Alignment(horizontal='center', vertical='center')
                # Calculer le total pour ce PVT (exclure TOTAL PVT du calcul)
                total_sim = 0
                total_om = 0
                for r in range(start, end + 1):
                    if ws1.cell(r, 2).value != 'TOTAL PVT':
                        total_sim += ws1.cell(r, 11).value or 0
                        total_om += ws1.cell(r, 16).value or 0
                ws1.cell(start, 18).value = total_sim + total_om + 100000

        # Appliquer les styles
        for row_idx in range(3, ws1.max_row + 1):
            for col_idx in range(1, ws1.max_column + 1):
                cell = ws1.cell(row_idx, col_idx)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')

            first_cell_val = ws1.cell(row_idx, 2).value
            dr_val = ws1.cell(row_idx, 1).value

            if first_cell_val == 'TOTAL PVT':
                for col in range(2, 17):
                    ws1.cell(row_idx, col).fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                    ws1.cell(row_idx, col).font = Font(bold=True)

            elif str(dr_val).startswith('TOTAL ') and dr_val != 'TOTAL GÉNÉRAL':
                for col in range(1, ws1.max_column + 1):
                    ws1.cell(row_idx, col).fill = PatternFill(start_color="FFE5CC", end_color="FFE5CC", fill_type="solid")
                    ws1.cell(row_idx, col).font = Font(bold=True)

            elif dr_val == 'TOTAL GÉNÉRAL':
                for col in range(1, ws1.max_column + 1):
                    ws1.cell(row_idx, col).fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
                    ws1.cell(row_idx, col).font = Font(bold=True, size=12, color="000000")

        # AJUSTER LA LARGEUR DES COLONNES
        column_widths = {
            'A': 15,   # DR
            'B': 25,   # PVT
            'C': 18,   # PRENOM_VTO
            'D': 18,   # NOM_VTO
            'E': 15,   # LOGIN
            'F': 18,   # Numéro Kabbu
            'G': 8,    # R (SIM)
            'H': 8,    # O (SIM)
            'I': 8,    # R/O (SIM)
            'J': 12,   # Gain Max (SIM)
            'K': 12,   # Gain (SIM)
            'L': 8,    # R (OM)
            'M': 8,    # O (OM)
            'N': 8,    # R/O (OM)
            'O': 12,   # Gain Max (OM)
            'P': 12,   # Gain (OM)
            'Q': 18,   # Gain Chauffeur
            'R': 28    # Gain SIM + OM + Chauffeur
        }

        for col, width in column_widths.items():
            ws1.column_dimensions[col].width = width

        # FIXER L'EN-TÊTE (lignes 1 et 2) pour la feuille "Détails Paiement"
        ws1.freeze_panes = 'A3'

        # Formater "Résumé PVT"
        ws2 = wb['Résumé PVT']
        for cell in ws2[1]:
            cell.fill = header_fill_gray
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border

        for row_idx in range(2, ws2.max_row + 1):
            for col_idx in range(1, ws2.max_column + 1):
                cell = ws2.cell(row_idx, col_idx)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')

                if row_idx == ws2.max_row:
                    cell.fill = PatternFill(start_color="FFE5CC", end_color="FFE5CC", fill_type="solid")
                    cell.font = Font(bold=True)

        # AJUSTER LA LARGEUR DES COLONNES pour Résumé PVT
        ws2.column_dimensions['A'].width = 15   # DR
        ws2.column_dimensions['B'].width = 25   # PVT
        ws2.column_dimensions['C'].width = 18   # CONTACT
        ws2.column_dimensions['D'].width = 15   # MONTANT
        ws2.column_dimensions['E'].width = 18   # GAIN PVT (5%)
        ws2.column_dimensions['F'].width = 18   # TOTAL GENERAL

        # FIXER L'EN-TÊTE (ligne 1) pour la feuille "Résumé PVT"
        ws2.freeze_panes = 'A2'

        # Sauvegarder
        final_buffer = BytesIO()
        wb.save(final_buffer)
        final_buffer.seek(0)

        st.download_button(
            label="📥 Télécharger le Rapport Excel (Format Manager)",
            data=final_buffer,
            file_name="paiement_mensuel_format_manager.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Erreur : {str(e)}")
        import traceback
        st.code(traceback.format_exc())