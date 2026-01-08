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
    st.write("🔍 **Vérification des doublons dans les données SIM:**")
    duplicates_check = df_filtre[df_filtre.duplicated(subset=['LOGIN'], keep=False)].sort_values(['LOGIN', 'REALISATION_SIM'], ascending=[True, False])

    if not duplicates_check.empty:
        st.warning(f"⚠️ {duplicates_check['LOGIN'].nunique()} LOGIN dupliqués détectés (erreur dans les données sources)")
        with st.expander("Voir les doublons détectés"):
            st.dataframe(duplicates_check[['LOGIN', 'PVT', 'DRV', 'PRENOM_VENDEUR', 'NOM_VENDEUR', 'REALISATION_SIM']])
            st.info("💡 Solution : On garde uniquement la ligne avec la plus grande réalisation SIM pour chaque LOGIN")

    # DÉDUPLICATION : Garder uniquement la ligne avec le MAXIMUM de réalisation pour chaque LOGIN
    nb_avant_dedup = len(df_filtre)
    df_filtre = df_filtre.sort_values('REALISATION_SIM', ascending=False)
    df_filtre = df_filtre.drop_duplicates(subset=['LOGIN'], keep='first')
    df_filtre = df_filtre.sort_values(['DRV', 'PVT'])
    nb_apres_dedup = len(df_filtre)

    if nb_avant_dedup > nb_apres_dedup:
        st.success(f"✅ Déduplication SIM: {nb_avant_dedup - nb_apres_dedup} doublons supprimés ({nb_avant_dedup} → {nb_apres_dedup} lignes)")
    else:
        st.success(f"✅ Aucun doublon dans les données SIM ({nb_apres_dedup} lignes uniques)")

    df_filtre['OBJECTIF SIM'] = 240
    df_filtre["TAUX D'ATTEINTE SIM"] = (df_filtre['REALISATION_SIM'] / 240 * 100).apply(lambda x: f"{round(x)}%")
    df_filtre['SI 100% ATTEINT SIM'] = 75000
    df_filtre['PAIEMENT_SIM'] = df_filtre['REALISATION_SIM'].apply(lambda x: 75000 if x >= 240 else round((x/240)*75000))
    df_filtre = df_filtre.merge(vto_df[["LOGIN", "KABBU"]], how="left")

    # === TRAITEMENT OM ===
    df_om['LOGIN'] = df_om['LOGIN'].astype(str).str.strip().str.lower()
    df_om['NOM_VENDEUR'] = df_om['NOM_VENDEUR'].astype(str).str.strip().str.upper()
    df_om['PRENOM_VENDEUR'] = df_om['PRENOM_VENDEUR'].astype(str).str.strip().str.upper()

    df_filtre_om = df_om[df_om['LOGIN'].isin(logins_concernes)].fillna(0)

    # DÉDUPLICATION OM : Vérifier et supprimer les doublons
    st.write("🔍 **Vérification des doublons dans les données OM:**")
    duplicates_om = df_filtre_om[df_filtre_om.duplicated(subset=['LOGIN'], keep=False)].sort_values(['LOGIN', 'REALISATION_OM'], ascending=[True, False])

    if not duplicates_om.empty:
        st.warning(f"⚠️ {duplicates_om['LOGIN'].nunique()} LOGIN OM dupliqués détectés")
        with st.expander("Voir les doublons OM"):
            st.dataframe(duplicates_om[['LOGIN', 'PRENOM_VENDEUR', 'NOM_VENDEUR', 'REALISATION_OM']])

    nb_avant_dedup_om = len(df_filtre_om)
    df_filtre_om = df_filtre_om.sort_values('REALISATION_OM', ascending=False)
    df_filtre_om = df_filtre_om.drop_duplicates(subset=['LOGIN'], keep='first')
    nb_apres_dedup_om = len(df_filtre_om)

    if nb_avant_dedup_om > nb_apres_dedup_om:
        st.success(f"✅ Déduplication OM: {nb_avant_dedup_om - nb_apres_dedup_om} doublons supprimés ({nb_avant_dedup_om} → {nb_apres_dedup_om} lignes)")
    else:
        st.success(f"✅ Aucun doublon dans les données OM ({nb_apres_dedup_om} lignes uniques)")

    df_filtre_om['OBJECTIF OM'] = 120
    df_filtre_om["TAUX D'ATTEINTE OM"] = (df_filtre_om['REALISATION_OM'] / 120 * 100).fillna(0).apply(lambda x: f"{round(x)}%")
    df_filtre_om['SI 100% ATTEINT OM'] = 25000
    df_filtre_om['PAIEMENT_OM'] = df_filtre_om['REALISATION_OM'].apply(lambda x: 25000 if x >= 120 else round((x/120)*25000))
    df_filtre_om = df_filtre_om.merge(vto_df[["LOGIN", "DRV", "PVT"]], how="left")

    # === FUSION SIM + OM ===
    st.write("🔄 **Fusion des données SIM et OM...**")

    df_test = pd.merge(
        df_filtre,
        df_filtre_om[["LOGIN", "REALISATION_OM", "OBJECTIF OM", "TAUX D'ATTEINTE OM", "SI 100% ATTEINT OM", "PAIEMENT_OM"]],
        on=["LOGIN"],
        how="outer"
    )

    # VÉRIFICATION FINALE : S'assurer qu'il n'y a aucun doublon après le merge
    duplicates_final = df_test[df_test.duplicated(subset=['LOGIN'], keep=False)].sort_values('LOGIN')

    if not duplicates_final.empty:
        st.error(f"❌ ATTENTION: {duplicates_final['LOGIN'].nunique()} doublons détectés après le merge!")
        with st.expander("Voir les doublons après merge"):
            st.dataframe(duplicates_final[['LOGIN', 'PVT', 'DRV', 'PRENOM_VENDEUR', 'NOM_VENDEUR', 'REALISATION_SIM', 'REALISATION_OM']])

        # Forcer la déduplication finale
        df_test = df_test.drop_duplicates(subset=['LOGIN'], keep='first')
        st.warning(f"⚠️ Déduplication forcée appliquée - {len(df_test)} lignes restantes")
    else:
        st.success(f"✅ Aucun doublon après fusion - {len(df_test)} VTO uniques")

    df_test["PAIEMENT CHAUFFEUR"] = None
    df_test["PAIEMENT SIM + OM + CHAUFFEUR"] = None
    df_test['PAIEMENT_SIM'] = df_test['PAIEMENT_SIM'].fillna(0)
    df_test['PAIEMENT_OM'] = df_test['PAIEMENT_OM'].fillna(0)
    df_test['REALISATION_SIM'] = df_test['REALISATION_SIM'].fillna(0)
    df_test['REALISATION_OM'] = df_test['REALISATION_OM'].fillna(0)

    # === CRÉATION DES TOTAUX PAR PVT ET DRV ===
    df_test_with_totals = pd.DataFrame(columns=df_test.columns)

    for drv, group_drv in df_test.groupby('DRV'):
        for pvt, group_pvt in group_drv.groupby('PVT'):
            df_test_with_totals = pd.concat([df_test_with_totals, group_pvt], ignore_index=True)

            # Total PVT
            row_total = {
                'PVT': "TOTAL PVT",
                'REALISATION_SIM': group_pvt['REALISATION_SIM'].sum(),
                'OBJECTIF SIM': group_pvt['OBJECTIF SIM'].sum(),
                "TAUX D'ATTEINTE SIM": f'{group_pvt["TAUX D\'ATTEINTE SIM"].apply(lambda x: float(x.strip("%"))).mean():.1f}%',
                'SI 100% ATTEINT SIM': group_pvt['SI 100% ATTEINT SIM'].sum(),
                'REALISATION_OM': group_pvt['REALISATION_OM'].sum(),
                'OBJECTIF OM': group_pvt['OBJECTIF OM'].sum(),
                "TAUX D'ATTEINTE OM": f'{group_pvt["TAUX D\'ATTEINTE OM"].apply(lambda x: float(str(x).replace("%", "")) if pd.notnull(x) else 0).mean():.1f}%',
                'SI 100% ATTEINT OM': group_pvt['SI 100% ATTEINT OM'].sum(),
                'PAIEMENT_OM': group_pvt['PAIEMENT_OM'].sum(),
                'PAIEMENT_SIM': group_pvt['PAIEMENT_SIM'].sum(),
                'PAIEMENT CHAUFFEUR': 100000,
                'PAIEMENT SIM + OM + CHAUFFEUR': group_pvt['PAIEMENT_SIM'].sum() + 100000 + group_pvt['PAIEMENT_OM'].sum()
            }
            df_test_with_totals = pd.concat([df_test_with_totals, pd.DataFrame([row_total])], ignore_index=True)

        # Total DRV
        row_total_drv = {
            'DRV': f"{drv}",
            'PVT': "TOTAL",
            'PAIEMENT_OM': group_drv['PAIEMENT_OM'].sum(),
            'PAIEMENT_SIM': group_drv['PAIEMENT_SIM'].sum(),
            'PAIEMENT CHAUFFEUR': 200000,
            'PAIEMENT SIM + OM + CHAUFFEUR': group_drv['PAIEMENT_SIM'].sum() + 200000 + group_drv['PAIEMENT_OM'].sum()
        }
        df_test_with_totals = pd.concat([df_test_with_totals, pd.DataFrame([row_total_drv])], ignore_index=True)

    # === TOTAL GLOBAL DE TOUS LES DR ===
    total_global_sim = df_test['PAIEMENT_SIM'].sum()
    total_global_om = df_test['PAIEMENT_OM'].sum()
    # Compter le nombre de DR pour les chauffeurs (200k par DR)
    nb_dr = df_test['DRV'].nunique()
    total_chauffeur_global = nb_dr * 200000

    row_total_global = {
        'DRV': "TOTAL GLOBAL",
        'PVT': "",
        'PAIEMENT_OM': total_global_om,
        'PAIEMENT_SIM': total_global_sim,
        'PAIEMENT CHAUFFEUR': total_chauffeur_global,
        'PAIEMENT SIM + OM + CHAUFFEUR': total_global_sim + total_global_om + total_chauffeur_global
    }
    df_test_with_totals = pd.concat([df_test_with_totals, pd.DataFrame([row_total_global])], ignore_index=True)

    # === CALCUL RÉSUMÉ PVT ===
    df_test["MONTANT"] = df_test["PAIEMENT_SIM"] + df_test["PAIEMENT_OM"]
    df_par_pvt = df_test.groupby(["DRV", "PVT"]).agg({'MONTANT': 'sum'}).reset_index()
    df_par_pvt["MONTANT"] = df_par_pvt["MONTANT"] + 100000
    df_par_pvt["GAIN PVT (5%)"] = df_par_pvt["MONTANT"] * 0.05
    df_par_pvt["TOTAL GENERAL"] = df_par_pvt["MONTANT"] + df_par_pvt["GAIN PVT (5%)"]

    pvt_df = load_pvt()
    df_par_pvt = df_par_pvt.merge(pvt_df[["PVT", "CONTACT"]], on="PVT", how="left")
    df_par_pvt = df_par_pvt[["DRV", "PVT", "CONTACT", "MONTANT", "GAIN PVT (5%)", "TOTAL GENERAL"]]

    df_par_pvt_display = df_par_pvt.copy()
    df_par_pvt_display.loc[len(df_par_pvt_display)] = [
        'TOTAL', '', '',
        df_par_pvt['MONTANT'].sum(),
        df_par_pvt['GAIN PVT (5%)'].sum(),
        df_par_pvt['TOTAL GENERAL'].sum()
    ]

    # === MESSAGE DE SUCCÈS ===
    st.success(f"✅ Traitement terminé : {df_test['LOGIN'].nunique()} VTO dans {df_test['PVT'].nunique()} PVT")

    # === EXPORT EXCEL ===
    try:
        buffer_output = BytesIO()

        with pd.ExcelWriter(buffer_output, engine='openpyxl') as writer:
            df_test_with_totals.to_excel(writer, sheet_name='Détails Paiement', index=False)
            df_par_pvt_display.to_excel(writer, sheet_name='Résumé PVT', index=False)

        buffer_output.seek(0)
        wb = load_workbook(buffer_output)

        # Style commun
        header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
        header_font = Font(bold=True, size=11)
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # Formater "Détails Paiement"
        ws1 = wb['Détails Paiement']
        for cell in ws1[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border

        for row_idx in range(2, ws1.max_row + 1):
            for col_idx in range(1, ws1.max_column + 1):
                cell = ws1.cell(row_idx, col_idx)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='center', vertical='center')

                if cell.value == 'TOTAL PVT':
                    for col in range(1, ws1.max_column + 1):
                        ws1.cell(row_idx, col).fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                        ws1.cell(row_idx, col).font = Font(bold=True)
                elif cell.value == 'TOTAL':
                    for col in range(1, ws1.max_column + 1):
                        ws1.cell(row_idx, col).fill = PatternFill(start_color="FFE5CC", end_color="FFE5CC", fill_type="solid")
                        ws1.cell(row_idx, col).font = Font(bold=True)
                elif cell.value == 'TOTAL GLOBAL':
                    for col in range(1, ws1.max_column + 1):
                        ws1.cell(row_idx, col).fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
                        ws1.cell(row_idx, col).font = Font(bold=True, size=12, color="FFFFFF")
                        ws1.cell(row_idx, col).alignment = Alignment(horizontal='center', vertical='center')

        # Formater "Résumé PVT"
        ws2 = wb['Résumé PVT']
        for cell in ws2[1]:
            cell.fill = header_fill
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

        # Sauvegarder
        final_buffer = BytesIO()
        wb.save(final_buffer)
        final_buffer.seek(0)

        st.download_button(
            label="📥 Télécharger le Rapport Excel",
            data=final_buffer,
            file_name="paiement_mensuel_global.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Erreur : {str(e)}")
        import traceback
        st.code(traceback.format_exc())