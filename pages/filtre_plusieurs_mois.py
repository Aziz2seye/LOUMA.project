# reporting_periode_multimois_sonatel.py
import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import sys
from pathlib import Path
from PIL import Image
from datetime import datetime
import calendar

# Ajouter le répertoire parent au path Python
current_dir = Path(__file__).parent
parent_dir = current_dir.parent
sys.path.insert(0, str(parent_dir))

# ====================
# Configuration page
# ====================
st.set_page_config(page_title="LOUMA - Reporting Multi-Mois", layout="wide")

# ====================
# COULEURS SONATEL (format aRGB pour openpyxl)
# ====================
SONATEL_COLORS = {
    "orange_primary": "FFFF7900",      # Orange principal Sonatel (format aRGB)
    "orange_dark": "FFFF5000",         # Orange foncé
    "orange_light": "FFFFA500",        # Orange clair
    "blue_dark": "FF003366",           # Bleu foncé
    "blue_light": "FF0066CC",          # Bleu clair
    "white": "FFFFFFFF",               # Blanc
    "gray_light": "FFF5F5F5",          # Gris clair
    "gray_dark": "FF333333",           # Gris foncé
    "success": "FF00B894",             # Vert succès
    "warning": "FFFDCB6E",             # Jaune avertissement
    "error": "FFE17055",               # Rouge erreur
    "light_orange": "FFFFE5CC"         # Orange clair
}

# ====================
# CSS avec couleurs Sonatel (format hex normal)
# ====================
st.markdown(f"""
<style>
    .section-title {{
        background: linear-gradient(135deg, #FF5000 0%, #FF7900 100%);
        color: white;
        padding: 0.8rem 1.5rem;
        border-radius: 10px;
        font-weight: 600;
        font-size: 1.2rem;
        margin-bottom: 1.2rem;
        text-align: center;
        box-shadow: 0 4px 12px rgba(255, 121, 0, 0.25);
    }}

    .metric-card {{
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 4px 15px rgba(255, 121, 0, 0.15);
        border: 2px solid #FFE5CC;
        text-align: center;
        transition: transform 0.3s ease;
    }}

    .metric-card:hover {{
        transform: translateY(-5px);
        box-shadow: 0 6px 20px rgba(255, 121, 0, 0.25);
    }}

    .metric-value {{
        font-size: 2.2rem;
        font-weight: 700;
        color: #FF7900;
        margin-bottom: 0.5rem;
    }}

    .metric-label {{
        font-size: 1rem;
        color: #333333;
        font-weight: 500;
    }}

    .stButton > button {{
        background: linear-gradient(135deg, #FF5000 0%, #FF7900 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.8rem 2rem;
        font-weight: 600;
        font-size: 1.1rem;
        box-shadow: 0 4px 12px rgba(255, 121, 0, 0.3);
        transition: all 0.3s ease;
        width: 100%;
    }}

    .stButton > button:hover {{
        background: linear-gradient(135deg, #FF7900 0%, #FF3000 100%);
        box-shadow: 0 6px 18px rgba(255, 121, 0, 0.5);
        transform: translateY(-2px);
    }}

    .file-box {{
        background: white;
        border-radius: 8px;
        padding: 1rem;
        margin-bottom: 1rem;
        border-left: 4px solid #FF7900;
        box-shadow: 0 2px 8px rgba(255, 121, 0, 0.1);
    }}

    .file-success {{
        border-left-color: #00B894;
    }}

    .file-error {{
        border-left-color: #E17055;
    }}

    .stSelectbox > div > div {{
        border-color: #FF7900 !important;
    }}

    .stSelectbox > div > div:hover {{
        border-color: #FF5000 !important;
    }}
</style>
""", unsafe_allow_html=True)

# ====================
# Charger le logo
# ====================
logo = None
logo_paths = [
    parent_dir / "assets" / "logo sonatel.png",
    Path("assets") / "logo sonatel.png",
    Path("../assets/logo sonatel.png"),
]

for logo_path in logo_paths:
    try:
        if logo_path.exists():
            logo = Image.open(logo_path)
            break
    except:
        continue

# ====================
# Header avec design Sonatel
# ====================
col_logo, col_title = st.columns([1, 3])

with col_logo:
    if logo:
        st.image(logo, width=220)

with col_title:
    st.markdown(f"""
    <div style="
        background: linear-gradient(135deg, #FF5000 0%, #FF7900 100%);
        padding: 2rem;
        border-radius: 20px;
        box-shadow: 0 8px 25px rgba(255, 121, 0, 0.4);
        display: flex;
        flex-direction: column;
        justify-content: center;
        border: 3px solid rgba(255, 255, 255, 0.2);
        height: 100%;
    ">
        <h1 style="
            color: white;
            font-size: 2.5rem;
            font-weight: 700;
            margin: 0;
            text-shadow: 3px 3px 10px rgba(0, 0, 0, 0.3);
        ">
            📊 Reporting Multi-Mois
        </h1>
        <p style="
            color: rgba(255, 255, 255, 0.95);
            font-size: 1.2rem;
            margin: 0.8rem 0 0 0;
            font-weight: 400;
        ">
            Analyse consolidée sur plusieurs mois - Orange Sénégal
        </p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# 🎯 Mapping DRV
DRV_MAPPING = {
    "DV-DRV2_DIRECTION REGIONALE DES VENTES DAKAR 2": "DR2",
    "DV-DRVS_DIRECTION REGIONALE DES VENTES SUD": "DRS",
    "DV-DRVSE_DIRECTION REGIONALE DES VENTES SUD-EST": "DRSE",
    "DV-DRVN_DIRECTION REGIONALE DES VENTES NORD": "DRN",
    "DV-DRVC_DIRECTION REGIONALE DES VENTES CENTRE": "DRC",
    "DV-DRVE_DIRECTION REGIONALE DES VENTES EST": "DRE"
}

# ====================
# FONCTIONS
# ====================
def load_vto():
    """Fonction pour charger les données VTO"""
    try:
        from utils import load_vto as real_load_vto
        return real_load_vto()
    except:
        data = {
            'LOGIN': ['vto001', 'vto002', 'vto003', 'vto004', 'vto005', 'vto006', 'vto007', 'vto008'],
            'NOM': ['DIOUF', 'NDIAYE', 'SARR', 'FALL', 'DIOP', 'GUEYE', 'MBOW', 'SY'],
            'PRENOM': ['Mamadou', 'Fatou', 'Moussa', 'Aïssatou', 'Cheikh', 'Alioune', 'Mariama', 'Oumar'],
            'DR': ['DR2', 'DRS', 'DRSE', 'DRN', 'DRC', 'DRE', 'DR2', 'DRS'],
            'PVT': ['PVT DAKAR CENTRE', 'PVT SUD', 'PVT SUD-EST', 'PVT NORD', 'PVT CENTRE', 'PVT EST', 'PVT DAKAR PLATEAU', 'PVT SUD 2']
        }
        return pd.DataFrame(data)

def process_single_month(df, vto_df, details, mois_nom):
    """Traite les données d'un mois"""
    column_mapping = {
        'MSISDN': 'REALISATION',
        'ACCUEIL_VENDEUR': 'PVT',
        'LOGIN_VENDEUR': 'LOGIN',
        'AGENCE_VENDEUR': 'DR',
        'NOM_VENDEUR': 'NOM_VENDEUR',
        'PRENOM_VENDEUR': 'PRENOM_VENDEUR',
        'ETAT_IDENTIFICATION': 'ETAT_IDENTIFICATION'
    }

    for original_col, new_col in column_mapping.items():
        if original_col in df.columns:
            df = df.rename(columns={original_col: new_col})

    if 'LOGIN' in df.columns:
        df['LOGIN'] = df['LOGIN'].astype(str).str.lower().str.strip()

    if 'DR' in df.columns:
        df['DR'] = df['DR'].astype(str).str.strip().str.upper()

    if 'NOM_VENDEUR' in df.columns:
        df['NOM_VENDEUR'] = df['NOM_VENDEUR'].astype(str).str.strip().str.upper()

    if 'PRENOM_VENDEUR' in df.columns:
        df['PRENOM_VENDEUR'] = df['PRENOM_VENDEUR'].astype(str).str.strip().str.upper()

    logins_concernes = vto_df["LOGIN"].astype(str).str.lower().str.strip().tolist()

    df_filtre_login = df[df['LOGIN'].isin(logins_concernes)] if 'LOGIN' in df.columns else df

    if 'ETAT_IDENTIFICATION' in df_filtre_login.columns:
        df_filtre = df_filtre_login[df_filtre_login['ETAT_IDENTIFICATION'].astype(str).isin(details)]
    else:
        df_filtre = df_filtre_login

    if 'DR' in df_filtre.columns:
        df_filtre["DR"] = df_filtre["DR"].replace(DRV_MAPPING)
        df_filtre["DR"] = df_filtre["DR"].apply(lambda x: x if x in DRV_MAPPING.values() else 'AUTRE')

    df_filtre['MOIS_NOM'] = mois_nom

    return df_filtre

def process_multi_month_data(df_list, mois_noms, vto_df, details, objectif_pvt=960):
    """Traite les données pour plusieurs mois"""
    all_dfs = []

    for i, (df, mois_nom) in enumerate(zip(df_list, mois_noms)):
        if df is not None and not df.empty:
            df_filtre = process_single_month(df, vto_df, details, mois_nom)
            if not df_filtre.empty:
                df_filtre['MOIS_INDEX'] = i + 1
                all_dfs.append(df_filtre)

    if not all_dfs:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    df_combined = pd.concat(all_dfs, ignore_index=True)
    nb_mois = len(df_list)

    # Résumé par PVT avec groupement unique
    if all(col in df_combined.columns for col in ['DR', 'PVT']):
        df_pvt_summary = df_combined.groupby(['DR', 'PVT'], as_index=False).size()
        df_pvt_summary.columns = ['DR', 'PVT', 'REALISATION']
        df_pvt_summary['OBJECTIF'] = objectif_pvt * nb_mois
        df_pvt_summary['R/O'] = (df_pvt_summary['REALISATION'] / df_pvt_summary['OBJECTIF'] * 100).round(0).astype(int)
        df_pvt_summary = df_pvt_summary.sort_values(['DR', 'PVT'])
    else:
        df_pvt_summary = pd.DataFrame(columns=['DR', 'PVT', 'REALISATION', 'OBJECTIF', 'R/O'])

    # Résumé par DR
    if 'DR' in df_combined.columns:
        df_dr_summary = df_combined.groupby('DR', as_index=False).size()
        df_dr_summary.columns = ['DR', 'REALISATION']

        if not df_pvt_summary.empty:
            pvt_count_per_dr = df_pvt_summary.groupby('DR')['PVT'].nunique()
            df_dr_summary['OBJECTIF'] = df_dr_summary['DR'].map(pvt_count_per_dr) * objectif_pvt * nb_mois
        else:
            df_dr_summary['OBJECTIF'] = 0

        df_dr_summary['R/O'] = (df_dr_summary['REALISATION'] / df_dr_summary['OBJECTIF'] * 100).round(0).astype(int)
        df_dr_summary = df_dr_summary.sort_values('DR')
    else:
        df_dr_summary = pd.DataFrame(columns=['DR', 'REALISATION', 'OBJECTIF', 'R/O'])

    # Détails par VTO
    required_vto_cols = ['DR', 'PVT', 'PRENOM_VENDEUR', 'NOM_VENDEUR', 'LOGIN']
    if all(col in df_combined.columns for col in required_vto_cols):
        df_reporting = df_combined.groupby(['DR', 'PVT', 'PRENOM_VENDEUR', 'NOM_VENDEUR', 'LOGIN'], as_index=False).size()
        df_reporting.columns = ['DR', 'PVT', 'PRENOM_VENDEUR', 'NOM_VENDEUR', 'LOGIN', 'REALISATION']
        df_reporting = df_reporting.sort_values(['DR', 'PVT', 'REALISATION'], ascending=[True, True, False])
    else:
        df_reporting = pd.DataFrame(columns=['DR', 'PVT', 'PRENOM_VENDEUR', 'NOM_VENDEUR', 'LOGIN', 'REALISATION'])

    return df_combined, df_pvt_summary, df_dr_summary, df_reporting, nb_mois

def generate_multi_month_excel_report(df_pvt_summary, df_dr_summary, df_reporting, periode_nom, annee, nb_mois):
    """Génère le fichier Excel avec formatage Sonatel"""
    buffer_output = BytesIO()

    # Totaux pour PVT
    total_realisation_pvt = int(df_pvt_summary['REALISATION'].sum()) if not df_pvt_summary.empty else 0
    total_objectif_pvt = int(df_pvt_summary['OBJECTIF'].sum()) if not df_pvt_summary.empty else 0
    total_ro_pvt = round((total_realisation_pvt / total_objectif_pvt * 100), 0) if total_objectif_pvt > 0 else 0

    # Totaux pour DR
    total_realisation_dr = int(df_dr_summary['REALISATION'].sum()) if not df_dr_summary.empty else 0
    total_objectif_dr = int(df_dr_summary['OBJECTIF'].sum()) if not df_dr_summary.empty else 0
    total_ro_dr = round((total_realisation_dr / total_objectif_dr * 100), 0) if total_objectif_dr > 0 else 0

    # Préparer les DataFrames pour export
    df_dr_summary_export = df_dr_summary.copy().reset_index(drop=True)
    new_row_dr = {'DR': 'TOTAL', 'REALISATION': total_realisation_dr, 'OBJECTIF': total_objectif_dr, 'R/O': int(total_ro_dr)}
    df_dr_summary_export = pd.concat([df_dr_summary_export, pd.DataFrame([new_row_dr])], ignore_index=True)

    df_pvt_summary_export = df_pvt_summary.copy().reset_index(drop=True)
    new_row_pvt = {'DR': '', 'PVT': 'TOTAL', 'REALISATION': total_realisation_pvt, 'OBJECTIF': total_objectif_pvt, 'R/O': int(total_ro_pvt)}
    df_pvt_summary_export = pd.concat([df_pvt_summary_export, pd.DataFrame([new_row_pvt])], ignore_index=True)

    total_realisation_vto = int(df_reporting['REALISATION'].sum()) if not df_reporting.empty else 0
    df_reporting_export = df_reporting.copy().reset_index(drop=True)
    new_row_vto = {'DR': '', 'PVT': '', 'PRENOM_VENDEUR': '', 'NOM_VENDEUR': '', 'LOGIN': 'TOTAL', 'REALISATION': total_realisation_vto}
    df_reporting_export = pd.concat([df_reporting_export, pd.DataFrame([new_row_vto])], ignore_index=True)

    # Écrire dans Excel
    with pd.ExcelWriter(buffer_output, engine='openpyxl') as writer:
        # Informations
        info_data = {
            'Paramètre': ['Période', 'Année', 'Nombre de mois', 'Objectif mensuel par PVT', 'Objectif total par PVT'],
            'Valeur': [periode_nom, str(annee), str(nb_mois), '960', str(960 * nb_mois)]
        }
        df_info = pd.DataFrame(info_data)
        df_info.to_excel(writer, sheet_name='Informations', index=False)

        # Feuilles principales
        df_dr_summary_export.to_excel(writer, sheet_name='Résumé DR', index=False)
        df_pvt_summary_export.to_excel(writer, sheet_name='Résumé PVT', index=False)
        df_reporting_export.to_excel(writer, sheet_name='Détails VTO', index=False)

    buffer_output.seek(0)
    wb = load_workbook(buffer_output)

    # STYLES SONATEL (format aRGB corrigé)
    sonatel_orange_fill = PatternFill(start_color=SONATEL_COLORS['orange_primary'], end_color=SONATEL_COLORS['orange_primary'], fill_type="solid")
    sonatel_orange_light_fill = PatternFill(start_color=SONATEL_COLORS['light_orange'], end_color=SONATEL_COLORS['light_orange'], fill_type="solid")
    sonatel_blue_fill = PatternFill(start_color=SONATEL_COLORS['blue_light'], end_color=SONATEL_COLORS['blue_light'], fill_type="solid")
    sonatel_gray_fill = PatternFill(start_color=SONATEL_COLORS['gray_light'], end_color=SONATEL_COLORS['gray_light'], fill_type="solid")

    # Couleurs pour R/O (format aRGB)
    vert_fill = PatternFill(start_color="FFC6EFCE", end_color="FFC6EFCE", fill_type="solid")
    vert_font = Font(color="FF006100", bold=True)

    jaune_fill = PatternFill(start_color="FFFFEB9C", end_color="FFFFEB9C", fill_type="solid")
    jaune_font = Font(color="FF9C6500", bold=True)

    rouge_fill = PatternFill(start_color="FFFFC7CE", end_color="FFFFC7CE", fill_type="solid")
    rouge_font = Font(color="FF9C0006", bold=True)

    header_font = Font(color="FFFFFFFF", bold=True, size=11)
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))

    total_fill = PatternFill(start_color=SONATEL_COLORS['light_orange'], end_color=SONATEL_COLORS['light_orange'], fill_type="solid")
    total_font = Font(bold=True, color=SONATEL_COLORS['orange_primary'], size=11)

    # ============================================
    # FEUILLE INFORMATIONS
    # ============================================
    ws_info = wb['Informations']
    ws_info.sheet_view.showGridLines = False

    # Style du titre
    ws_info.merge_cells('A1:B1')
    ws_info['A1'] = f"REPORTING {periode_nom.upper()} {annee}"
    ws_info['A1'].font = Font(color=SONATEL_COLORS['orange_primary'], bold=True, size=14)
    ws_info['A1'].alignment = Alignment(horizontal='center', vertical='center')
    ws_info['A1'].fill = sonatel_orange_light_fill

    # Style des en-têtes
    for cell in ws_info[2]:
        cell.fill = sonatel_orange_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border

    # Style des données
    for row in ws_info.iter_rows(min_row=3, max_row=ws_info.max_row, min_col=1, max_col=2):
        for cell in row:
            cell.border = border
            cell.alignment = Alignment(horizontal='left', vertical='center')

    # Ajuster largeurs
    ws_info.column_dimensions['A'].width = 25
    ws_info.column_dimensions['B'].width = 30

    # ============================================
    # FEUILLE RÉSUMÉ DR
    # ============================================
    ws_dr = wb['Résumé DR']
    ws_dr.sheet_view.showGridLines = False

    # En-têtes avec couleur Sonatel
    for cell in ws_dr[1]:
        cell.fill = sonatel_orange_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border

    # Renommer les colonnes
    ws_dr['A1'].value = 'DIRECTION RÉGIONALE'
    ws_dr['B1'].value = 'RÉALISATION'
    ws_dr['C1'].value = 'OBJECTIF'
    ws_dr['D1'].value = 'R/O (%)'

    # Formater les données
    for row_idx in range(2, ws_dr.max_row + 1):
        for col_idx in range(1, 5):
            cell = ws_dr.cell(row=row_idx, column=col_idx)
            cell.border = border
            cell.alignment = Alignment(horizontal='center', vertical='center')

            # Style pour R/O avec couleurs
            if col_idx == 4:  # Colonne R/O
                if cell.value is not None and row_idx < ws_dr.max_row:
                    try:
                        ro_value = float(cell.value)
                        cell.value = f"{int(ro_value)}%"

                        if ro_value >= 100:
                            cell.fill = vert_fill
                            cell.font = vert_font
                        elif 80 <= ro_value < 100:
                            cell.fill = jaune_fill
                            cell.font = jaune_font
                        elif ro_value < 80:
                            cell.fill = rouge_fill
                            cell.font = rouge_font
                    except:
                        pass

            # Style pour la ligne TOTAL
            if row_idx == ws_dr.max_row:
                cell.fill = total_fill
                cell.font = total_font
                if col_idx == 4 and cell.value:
                    try:
                        ro_value = float(str(cell.value).replace('%', ''))
                        cell.value = f"{int(ro_value)}%"
                    except:
                        pass

    # Ajuster largeurs
    ws_dr.column_dimensions['A'].width = 25
    ws_dr.column_dimensions['B'].width = 15
    ws_dr.column_dimensions['C'].width = 15
    ws_dr.column_dimensions['D'].width = 12

    # ============================================
    # FEUILLE RÉSUMÉ PVT (AVEC FUSION DES CELLULES DR)
    # ============================================
    ws_pvt = wb['Résumé PVT']
    ws_pvt.sheet_view.showGridLines = False

    # En-têtes avec couleur Sonatel
    for cell in ws_pvt[1]:
        cell.fill = sonatel_orange_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border

    # Renommer les colonnes
    ws_pvt['A1'].value = 'DR'
    ws_pvt['B1'].value = 'POINT DE VENTE'
    ws_pvt['C1'].value = 'RÉALISATION'
    ws_pvt['D1'].value = 'OBJECTIF'
    ws_pvt['E1'].value = 'R/O (%)'

    # Identifier les plages à fusionner pour les DR
    dr_ranges = {}
    current_dr = None
    start_row = 2

    for row_idx in range(2, ws_pvt.max_row):
        dr_value = ws_pvt.cell(row=row_idx, column=1).value

        if dr_value != current_dr:
            if current_dr is not None and start_row < row_idx:
                dr_ranges[current_dr] = (start_row, row_idx - 1)
            current_dr = dr_value
            start_row = row_idx

    # Dernière plage
    if current_dr is not None and start_row < ws_pvt.max_row:
        dr_ranges[current_dr] = (start_row, ws_pvt.max_row - 1)

    # Fusionner les cellules DR et centrer verticalement
    for dr, (start, end) in dr_ranges.items():
        if end > start:  # Fusionner seulement si plusieurs lignes
            ws_pvt.merge_cells(f'A{start}:A{end}')
            merged_cell = ws_pvt.cell(row=start, column=1)
            merged_cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            merged_cell.font = Font(bold=True, color=SONATEL_COLORS['blue_dark'])

    # Formater toutes les cellules
    for row_idx in range(2, ws_pvt.max_row + 1):
        for col_idx in range(1, 6):
            cell = ws_pvt.cell(row=row_idx, column=col_idx)
            cell.border = border

            # Alignement
            if col_idx in [1, 2]:  # DR et PVT
                cell.alignment = Alignment(horizontal='left', vertical='center')
            else:  # Chiffres
                cell.alignment = Alignment(horizontal='center', vertical='center')

            # Style pour R/O
            if col_idx == 5 and row_idx < ws_pvt.max_row:  # R/O sauf TOTAL
                if cell.value is not None:
                    try:
                        ro_value = float(cell.value)
                        cell.value = f"{int(ro_value)}%"

                        if ro_value >= 100:
                            cell.fill = vert_fill
                            cell.font = vert_font
                        elif 80 <= ro_value < 100:
                            cell.fill = jaune_fill
                            cell.font = jaune_font
                        elif ro_value < 80:
                            cell.fill = rouge_fill
                            cell.font = rouge_font
                    except:
                        pass

            # Style pour la ligne TOTAL
            if row_idx == ws_pvt.max_row:
                cell.fill = total_fill
                cell.font = total_font
                if col_idx == 1:
                    ws_pvt.merge_cells(f'A{row_idx}:B{row_idx}')
                    cell.value = 'TOTAL'
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                elif col_idx == 5 and cell.value:
                    try:
                        ro_value = float(str(cell.value).replace('%', ''))
                        cell.value = f"{int(ro_value)}%"
                    except:
                        pass

    # Ajuster largeurs
    ws_pvt.column_dimensions['A'].width = 8
    ws_pvt.column_dimensions['B'].width = 40
    ws_pvt.column_dimensions['C'].width = 15
    ws_pvt.column_dimensions['D'].width = 15
    ws_pvt.column_dimensions['E'].width = 12

    # ============================================
    # FEUILLE DÉTAILS VTO (AVEC FUSION DES CELLULES)
    # ============================================
    ws_vto = wb['Détails VTO']
    ws_vto.sheet_view.showGridLines = False

    # En-têtes avec couleur Sonatel
    for cell in ws_vto[1]:
        cell.fill = sonatel_orange_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border

    # Renommer les colonnes
    ws_vto['A1'].value = 'DR'
    ws_vto['B1'].value = 'POINT DE VENTE'
    ws_vto['C1'].value = 'PRÉNOM'
    ws_vto['D1'].value = 'NOM'
    ws_vto['E1'].value = 'LOGIN'
    ws_vto['F1'].value = 'RÉALISATION'

    # Fusionner les cellules DR et PVT
    dr_ranges_vto = {}
    pvt_ranges_vto = {}
    current_dr = None
    current_pvt = None
    dr_start = 2
    pvt_start = 2

    for row_idx in range(2, ws_vto.max_row):
        dr_value = ws_vto.cell(row=row_idx, column=1).value
        pvt_value = ws_vto.cell(row=row_idx, column=2).value

        # DR fusion
        if dr_value != current_dr:
            if current_dr is not None and dr_start < row_idx:
                dr_ranges_vto[current_dr] = (dr_start, row_idx - 1)
            current_dr = dr_value
            dr_start = row_idx

        # PVT fusion
        if pvt_value != current_pvt:
            if current_pvt is not None and pvt_start < row_idx:
                pvt_ranges_vto[current_pvt] = (pvt_start, row_idx - 1)
            current_pvt = pvt_value
            pvt_start = row_idx

    # Dernières plages
    if current_dr is not None and dr_start < ws_vto.max_row:
        dr_ranges_vto[current_dr] = (dr_start, ws_vto.max_row - 1)

    if current_pvt is not None and pvt_start < ws_vto.max_row:
        pvt_ranges_vto[current_pvt] = (pvt_start, ws_vto.max_row - 1)

    # Fusionner DR
    for dr, (start, end) in dr_ranges_vto.items():
        if end > start:
            ws_vto.merge_cells(f'A{start}:A{end}')
            merged_cell = ws_vto.cell(row=start, column=1)
            merged_cell.alignment = Alignment(horizontal='center', vertical='center')
            merged_cell.font = Font(bold=True, color=SONATEL_COLORS['blue_dark'])

    # Fusionner PVT
    for pvt, (start, end) in pvt_ranges_vto.items():
        if end > start:
            ws_vto.merge_cells(f'B{start}:B{end}')
            merged_cell = ws_vto.cell(row=start, column=2)
            merged_cell.alignment = Alignment(horizontal='left', vertical='center')
            merged_cell.font = Font(bold=True)

    # Formater toutes les cellules
    for row_idx in range(2, ws_vto.max_row + 1):
        for col_idx in range(1, 7):
            cell = ws_vto.cell(row=row_idx, column=col_idx)
            cell.border = border

            # Alignement
            if col_idx in [1, 2, 3, 4, 5]:  # Textes
                cell.alignment = Alignment(horizontal='left', vertical='center')
            else:  # Chiffres
                cell.alignment = Alignment(horizontal='center', vertical='center')

            # Style pour la ligne TOTAL
            if row_idx == ws_vto.max_row:
                cell.fill = total_fill
                cell.font = total_font
                if col_idx == 1:
                    ws_vto.merge_cells(f'A{row_idx}:E{row_idx}')
                    cell.value = 'TOTAL'
                    cell.alignment = Alignment(horizontal='center', vertical='center')

    # Ajuster largeurs
    ws_vto.column_dimensions['A'].width = 8
    ws_vto.column_dimensions['B'].width = 35
    ws_vto.column_dimensions['C'].width = 15
    ws_vto.column_dimensions['D'].width = 15
    ws_vto.column_dimensions['E'].width = 15
    ws_vto.column_dimensions['F'].width = 15

    final_buffer = BytesIO()
    wb.save(final_buffer)
    final_buffer.seek(0)

    return final_buffer

def load_data_file(uploaded_file):
    """Charge un fichier"""
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file, encoding='utf-8', sep=';')
        elif uploaded_file.name.endswith('.xls'):
            df = pd.read_excel(uploaded_file, engine='xlrd')
        else:
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            best_sheet = sheet_names[0]

            for sheet in sheet_names:
                if 'data' in sheet.lower() or 'feuille' in sheet.lower():
                    best_sheet = sheet
                    break

            df = pd.read_excel(uploaded_file, sheet_name=best_sheet)

        df.columns = df.columns.str.strip()

        return df, True, f"✅ {len(df):,} lignes"

    except Exception as e:
        return pd.DataFrame(), False, f"❌ Erreur: {str(e)}"

# ====================
# INTERFACE PRINCIPALE
# ====================
def main():
    st.markdown('<div class="section-title">📅 CONFIGURATION DE LA PÉRIODE</div>', unsafe_allow_html=True)

    current_date = datetime.now()
    current_year = current_date.year

    col1, col2 = st.columns(2)

    with col1:
        nb_mois = st.selectbox(
            "Nombre de mois à analyser",
            options=list(range(2, 13)),
            format_func=lambda x: f"{x} mois",
            index=2  # 4 mois par défaut
        )

        st.markdown(f"""
        <div class="file-box">
            <div style="color: #FF7900; font-weight: 600;">
                Période sélectionnée : {nb_mois} mois
            </div>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        annee = st.selectbox(
            "Année de référence",
            options=list(range(current_year - 2, current_year + 1)),
            index=2
        )

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown('<div class="section-title">📁 IMPORTATION DES FICHIERS</div>', unsafe_allow_html=True)

    st.write(f"**Veuillez importer {nb_mois} fichiers Excel/CSV (un par mois) :**")

    mois_options = [
        'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin',
        'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre'
    ]

    uploaded_files = []

    # Interface dynamique pour chaque mois
    cols = st.columns(3)
    col_index = 0

    for i in range(nb_mois):
        with cols[col_index]:
            mois_selectionne = st.selectbox(
                f"Mois {i+1}",
                options=mois_options,
                key=f"mois_{i}",
                index=i % 12
            )

            uploaded_file = st.file_uploader(
                f"Fichier {mois_selectionne}",
                type=["xlsx", "csv", "xls"],
                key=f"file_{i}"
            )

            if uploaded_file:
                uploaded_files.append((mois_selectionne, uploaded_file))
                st.markdown(f"""
                <div class="file-box file-success">
                    <div style="color: #00B894;">
                        ✅ {mois_selectionne} chargé
                    </div>
                </div>
                """, unsafe_allow_html=True)

        col_index = (col_index + 1) % 3

    st.markdown("<br>", unsafe_allow_html=True)

    # Bouton pour traiter les données
    if st.button("🚀 TRAITER LES DONNÉES ET GÉNÉRER LE RAPPORT", use_container_width=True):
        if len(uploaded_files) == nb_mois:
            with st.spinner(f"Traitement des données pour {nb_mois} mois..."):
                try:
                    # Charger les données
                    vto_df = load_vto()
                    details = ["En Cours-Identification", "Identifie", "Identifie Photo"]
                    objectif_mensuel = 960

                    df_list = []
                    mois_noms = []
                    fichiers_valides = True

                    # Charger chaque fichier
                    for mois_nom, file in uploaded_files:
                        df, success, message = load_data_file(file)
                        if success and not df.empty:
                            df_list.append(df)
                            mois_noms.append(mois_nom)
                            st.markdown(f"""
                            <div class="file-box">
                                <div style="color: #FF7900;">
                                    📊 {mois_nom} : {len(df):,} lignes traitées
                                </div>
                            </div>
                            """, unsafe_allow_html=True)
                        else:
                            st.error(f"Erreur avec le fichier {mois_nom}: {message}")
                            fichiers_valides = False
                            break

                    if fichiers_valides and len(df_list) == nb_mois:
                        # Traiter les données
                        df_combined, df_pvt_summary, df_dr_summary, df_reporting, nb_mois_traites = process_multi_month_data(
                            df_list, mois_noms, vto_df, details, objectif_mensuel
                        )

                        if not df_combined.empty:
                            # Afficher les métriques
                            st.markdown('<div class="section-title">📊 RÉSUMÉ DE LA PÉRIODE</div>', unsafe_allow_html=True)

                            col1, col2, col3, col4 = st.columns(4)

                            with col1:
                                total_ventes = df_combined.shape[0]
                                st.markdown(f"""
                                <div class="metric-card">
                                    <div class="metric-value">{total_ventes:,}</div>
                                    <div class="metric-label">Ventes totales</div>
                                </div>
                                """, unsafe_allow_html=True)

                            with col2:
                                vto_actifs = df_combined['LOGIN'].nunique() if 'LOGIN' in df_combined.columns else 0
                                st.markdown(f"""
                                <div class="metric-card">
                                    <div class="metric-value">{vto_actifs}</div>
                                    <div class="metric-label">VTO Actifs</div>
                                </div>
                                """, unsafe_allow_html=True)

                            with col3:
                                pvt_concernes = df_combined['PVT'].nunique() if 'PVT' in df_combined.columns else 0
                                st.markdown(f"""
                                <div class="metric-card">
                                    <div class="metric-value">{pvt_concernes}</div>
                                    <div class="metric-label">PVT Actifs</div>
                                </div>
                                """, unsafe_allow_html=True)

                            with col4:
                                total_objectif = df_pvt_summary['OBJECTIF'].sum() if not df_pvt_summary.empty else 1
                                taux_realisation = round((total_ventes / total_objectif * 100), 1) if total_objectif > 0 else 0
                                st.markdown(f"""
                                <div class="metric-card">
                                    <div class="metric-value">{taux_realisation}%</div>
                                    <div class="metric-label">Taux R/O Global</div>
                                </div>
                                """, unsafe_allow_html=True)

                            # Générer et télécharger le fichier
                            st.markdown("<br>", unsafe_allow_html=True)
                            st.markdown('<div class="section-title">📥 TÉLÉCHARGEMENT DU RAPPORT</div>', unsafe_allow_html=True)

                            periode_nom = f"{nb_mois}_mois_{'_'.join([m[:3] for m in mois_noms])}"

                            excel_buffer = generate_multi_month_excel_report(
                                df_pvt_summary, df_dr_summary, df_reporting, periode_nom, annee, nb_mois
                            )

                            file_name = f"Reporting_{nb_mois}_mois_{annee}.xlsx"

                            st.download_button(
                                label=f"📥 Télécharger le Reporting ({nb_mois} mois)",
                                data=excel_buffer,
                                file_name=file_name,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )

                            st.success(f"✅ Rapport prêt pour {nb_mois} mois ({', '.join(mois_noms)}) !")

                            # Aperçu des données
                            st.markdown("<br>", unsafe_allow_html=True)
                            with st.expander("👁️ APERÇU DES DONNÉES TRAITÉES", expanded=False):
                                tab1, tab2, tab3 = st.tabs(["Résumé DR", "Résumé PVT (avec fusion DR)", "Détails VTO"])

                                with tab1:
                                    if not df_dr_summary.empty:
                                        st.dataframe(df_dr_summary, use_container_width=True)

                                with tab2:
                                    if not df_pvt_summary.empty:
                                        st.write("**Note :** Dans le fichier Excel, les cellules DR sont fusionnées et centrées")
                                        st.dataframe(df_pvt_summary, use_container_width=True)

                                with tab3:
                                    if not df_reporting.empty:
                                        st.dataframe(df_reporting.head(15), use_container_width=True)

                        else:
                            st.warning("⚠️ Aucune donnée trouvée après traitement.")
                    else:
                        st.error(f"❌ Tous les fichiers doivent être valides. {len(df_list)}/{nb_mois} fichiers valides.")

                except Exception as e:
                    st.error(f"❌ Erreur lors du traitement : {str(e)}")
                    import traceback
                    with st.expander("Détails de l'erreur"):
                        st.code(traceback.format_exc())
        else:
            st.warning(f"⚠️ Veuillez importer {nb_mois} fichiers (vous en avez {len(uploaded_files)})")

# ====================
# EXÉCUTION
# ====================
if __name__ == "__main__":
    main()