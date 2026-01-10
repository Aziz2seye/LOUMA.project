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
st.set_page_config(page_title="LOUMA - Reporting Mensuel", layout="wide", initial_sidebar_state="expanded")

# CSS personnalisé Orange Sonatel
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
    .stButton > button:hover {
        background: linear-gradient(135deg, #FF5000 0%, #FF3000 100%);
        box-shadow: 0 6px 18px rgba(255, 121, 0, 0.5);
        transform: translateY(-2px);
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
    .metric-card {
        background: white;
        border-radius: 12px;
        padding: 1.5rem;
        box-shadow: 0 4px 15px rgba(255, 121, 0, 0.15);
        border: 2px solid #FFE5CC;
        text-align: center;
    }
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        color: #FF7900;
    }
    .metric-label {
        font-size: 1rem;
        color: #666;
        margin-top: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# Charger le logo
logo = None
current_dir = Path(__file__).parent
parent_dir = current_dir.parent
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

# Header avec logo et titre
col_logo, col_title = st.columns([1, 3])

with col_logo:
    if logo:
        st.image(logo, width=280)

with col_title:
    st.markdown("""
    <div style="
        background: linear-gradient(135deg, #FF7900 0%, #FF5000 100%);
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
            📈 Reporting Mensuel
        </h1>
        <p style="
            color: rgba(255, 255, 255, 0.95);
            font-size: 1.2rem;
            margin: 0.8rem 0 0 0;
            font-weight: 400;
        ">
            Analyse des performances commerciales - Orange Sénégal
        </p>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# Mapping DRV unique
DRV_MAPPING = {
    "DV-DRV2_DIRECTION REGIONALE DES VENTES DAKAR 2": "DR2",
    "DV-DRVS_DIRECTION REGIONALE DES VENTES SUD": "DRS",
    "DV-DRVSE_DIRECTION REGIONALE DES VENTES SUD-EST": "DRSE",
    "DV-DRVN_DIRECTION REGIONALE DES VENTES NORD": "DRN",
    "DV-DRVC_DIRECTION REGIONALE DES VENTES CENTRE": "DRC",
    "DV-DRVE_DIRECTION REGIONALE DES VENTES EST": "DRE",
    "DV-DRV1_DIRECTION REGIONALE DES VENTES DAKAR 1": "DR1"
}

def generate_monthly_excel_report(df_final, mois_nom, annee, objectif_pvt=960):
    """Génère un fichier Excel avec mise en forme style NFC"""

    output = BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # FORMATS
        h_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#FF6600', 'font_color': 'white',
            'border': 1, 'align': 'center', 'valign': 'vcenter'
        })
        dr_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'left'
        })
        dr_num_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center'
        })
        pvt_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#F2F2F2', 'border': 1, 'indent': 1
        })
        pvt_num_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#F2F2F2', 'border': 1, 'align': 'center'
        })
        vendeur_fmt = workbook.add_format({
            'border': 1, 'indent': 2, 'align': 'center', 'font_size': 9
        })
        vendeur_num_fmt = workbook.add_format({
            'border': 1, 'align': 'center', 'font_size': 9
        })
        total_fmt = workbook.add_format({
            'bold': True, 'bg_color': '#FF6600', 'font_color': 'white',
            'border': 1, 'align': 'center'
        })

        # FEUILLE 1: SYNTHESE DR
        ws1 = workbook.add_worksheet('SYNTHESE DR')
        headers_dr = ['DR', 'REALISATION', 'OBJECTIF', 'R/O (%)']
        for c, h in enumerate(headers_dr):
            ws1.write(0, c, h, h_fmt)

        synthese_dr = df_final.groupby('DR').size().reset_index(name='REALISATION')
        pvt_par_dr = df_final.groupby('DR')['PVT'].nunique().reset_index(name='NB_PVT')
        synthese_dr = synthese_dr.merge(pvt_par_dr, on='DR')
        synthese_dr['OBJECTIF'] = synthese_dr['NB_PVT'] * objectif_pvt
        synthese_dr['R/O'] = ((synthese_dr['REALISATION'] / synthese_dr['OBJECTIF']) * 100).round(0)

        for i, r in synthese_dr.iterrows():
            ws1.write(i+1, 0, r['DR'], dr_fmt)
            ws1.write(i+1, 1, int(r['REALISATION']), dr_num_fmt)
            ws1.write(i+1, 2, int(r['OBJECTIF']), dr_num_fmt)
            ws1.write(i+1, 3, f"{int(r['R/O'])}%", dr_num_fmt)

        # Ligne TOTAL
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
        for c, h in enumerate(headers_dr):
            ws2.write(0, c, h, h_fmt)

        curr_row = 1
        total_real_pvt = 0
        total_obj_pvt = 0

        for dr, dr_group in df_final.groupby('DR', sort=True):
            if len(dr_group) == 0:
                continue

            # Ligne DR
            real_dr = len(dr_group)
            nb_pvt_dr = dr_group['PVT'].nunique()
            obj_dr = nb_pvt_dr * objectif_pvt
            ro_dr = round((real_dr / obj_dr * 100), 0) if obj_dr > 0 else 0

            ws2.write(curr_row, 0, dr, dr_fmt)
            ws2.write(curr_row, 1, real_dr, dr_num_fmt)
            ws2.write(curr_row, 2, obj_dr, dr_num_fmt)
            ws2.write(curr_row, 3, f"{int(ro_dr)}%", dr_num_fmt)
            curr_row += 1

            # Lignes PVT
            for pvt, pvt_group in dr_group.groupby('PVT', sort=True):
                if len(pvt_group) == 0:
                    continue

                real_pvt = len(pvt_group)
                obj_pvt = objectif_pvt
                ro_pvt = round((real_pvt / obj_pvt * 100), 0) if obj_pvt > 0 else 0

                ws2.write(curr_row, 0, pvt, pvt_fmt)
                ws2.write(curr_row, 1, real_pvt, pvt_num_fmt)
                ws2.write(curr_row, 2, obj_pvt, pvt_num_fmt)
                ws2.write(curr_row, 3, f"{int(ro_pvt)}%", pvt_num_fmt)
                curr_row += 1

                total_real_pvt += real_pvt
                total_obj_pvt += obj_pvt

        # Ligne TOTAL
        total_ro_pvt = round((total_real_pvt / total_obj_pvt * 100), 0) if total_obj_pvt > 0 else 0
        ws2.write(curr_row, 0, 'TOTAL', total_fmt)
        ws2.write(curr_row, 1, total_real_pvt, total_fmt)
        ws2.write(curr_row, 2, total_obj_pvt, total_fmt)
        ws2.write(curr_row, 3, f"{int(total_ro_pvt)}%", total_fmt)

        ws2.set_column('A:A', 45)
        ws2.set_column('B:D', 15)

        # FEUILLE 3: REPORTING DR-PVT-VENDEURS
        ws3 = workbook.add_worksheet('REPORTING DR-PVT-VENDEURS')
        headers_vendeurs = ['DR/PVT/VENDEUR', 'Prénom', 'Nom', 'LOGIN', 'REALISATION']
        for c, h in enumerate(headers_vendeurs):
            ws3.write(0, c, h, h_fmt)

        curr_row = 1
        total_real_vendeurs = 0

        for dr, dr_group in df_final.groupby('DR', sort=True):
            if len(dr_group) == 0:
                continue

            # Ligne DR
            real_dr = len(dr_group)
            ws3.write(curr_row, 0, dr, dr_fmt)
            ws3.write(curr_row, 4, real_dr, dr_num_fmt)
            curr_row += 1

            for pvt, pvt_group in dr_group.groupby('PVT', sort=True):
                if len(pvt_group) == 0:
                    continue

                # Ligne PVT
                real_pvt = len(pvt_group)
                ws3.write(curr_row, 0, pvt, pvt_fmt)
                ws3.write(curr_row, 4, real_pvt, pvt_num_fmt)
                curr_row += 1

                total_real_vendeurs += real_pvt

                # Lignes VENDEURS
                vendeurs = pvt_group.groupby(['PRENOM_VENDEUR', 'NOM_VENDEUR', 'LOGIN']).size().reset_index(name='REALISATION')
                vendeurs = vendeurs.sort_values('REALISATION', ascending=False)

                for _, v in vendeurs.iterrows():
                    ws3.write(curr_row, 0, 'VENDEUR', vendeur_fmt)
                    ws3.write(curr_row, 1, v['PRENOM_VENDEUR'], vendeur_fmt)
                    ws3.write(curr_row, 2, v['NOM_VENDEUR'], vendeur_fmt)
                    ws3.write(curr_row, 3, v['LOGIN'], vendeur_fmt)
                    ws3.write(curr_row, 4, int(v['REALISATION']), vendeur_num_fmt)
                    curr_row += 1

        # Ligne TOTAL
        ws3.write(curr_row, 0, 'TOTAL', total_fmt)
        ws3.write(curr_row, 4, total_real_vendeurs, total_fmt)

        ws3.set_column('A:A', 35)
        ws3.set_column('B:C', 20)
        ws3.set_column('D:D', 25)
        ws3.set_column('E:E', 15)

    output.seek(0)
    return output

# INTERFACE PRINCIPALE
def main():
    st.markdown('<div class="section-title">📅 SÉLECTION DE LA PÉRIODE</div>', unsafe_allow_html=True)

    current_date = datetime.now()
    current_year = current_date.year
    current_month = current_date.month

    col_month, col_year = st.columns(2)

    with col_month:
        selected_month = st.selectbox(
            "Mois",
            options=list(range(1, 13)),
            format_func=lambda x: calendar.month_name[x],
            index=current_month - 1
        )
        mois_nom = calendar.month_name[selected_month]

    with col_year:
        selected_year = st.selectbox(
            "Année",
            options=list(range(current_year - 2, current_year + 1)),
            index=2
        )

    st.markdown("<br>", unsafe_allow_html=True)

    st.markdown('<div class="section-title">📁 IMPORTATION DES DONNÉES</div>', unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        f"Importer le fichier Excel/CSV pour {mois_nom} {selected_year}",
        type=["xlsx", "csv", "xls"],
        help="Le fichier doit contenir: MSISDN, ACCUEIL_VENDEUR (PVT), LOGIN_VENDEUR, AGENCE_VENDEUR (DR), NOM_VENDEUR, PRENOM_VENDEUR, ETAT_IDENTIFICATION"
    )

    if uploaded_file:
        st.markdown("<br>", unsafe_allow_html=True)

        try:
            # Lecture du fichier
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, encoding='utf-8', sep=';')
            elif uploaded_file.name.endswith('.xls'):
                df = pd.read_excel(uploaded_file, engine='xlrd')
            else:
                xls = pd.ExcelFile(uploaded_file)
                sheet_names = xls.sheet_names
                if len(sheet_names) == 1:
                    selected_sheet = sheet_names[0]
                else:
                    selected_sheet = st.selectbox(
                        "🗂 Choisir la feuille à exploiter :",
                        options=sheet_names
                    )
                df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

            st.success(f"✅ Fichier chargé avec succès ! {len(df)} lignes trouvées")

            # Renommer les colonnes
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

            # Nettoyage des données
            if 'LOGIN' in df.columns:
                df['LOGIN'] = df['LOGIN'].astype(str).str.lower().str.strip()
            if 'DR' in df.columns:
                df['DR'] = df['DR'].astype(str).str.strip().str.upper()
            if 'PVT' in df.columns:
                df['PVT'] = df['PVT'].astype(str).str.strip().str.upper()
            if 'NOM_VENDEUR' in df.columns:
                df['NOM_VENDEUR'] = df['NOM_VENDEUR'].astype(str).str.strip().str.upper()
            if 'PRENOM_VENDEUR' in df.columns:
                df['PRENOM_VENDEUR'] = df['PRENOM_VENDEUR'].astype(str).str.strip().str.upper()

            # Filtrage selon l'état d'identification
            details = ["En Cours-Identification", "Identifie", "Identifie Photo"]

            total_avant = len(df)
            st.info(f"📊 **Total AVANT filtrage** : {total_avant:,} lignes")

            if 'ETAT_IDENTIFICATION' in df.columns:
                df_filtre = df[df['ETAT_IDENTIFICATION'].astype(str).isin(details)].copy()
            else:
                df_filtre = df.copy()

            # Mapper les DR
            if 'DR' in df_filtre.columns:
                df_filtre = df_filtre[df_filtre['DR'].isin(DRV_MAPPING.keys())].copy()
                df_filtre["DR"] = df_filtre["DR"].replace(DRV_MAPPING)

            total_apres = len(df_filtre)
            difference = total_avant - total_apres

            col_stat1, col_stat2, col_stat3 = st.columns(3)
            with col_stat1:
                st.metric("Total AVANT", f"{total_avant:,}")
            with col_stat2:
                st.metric("Total APRÈS", f"{total_apres:,}")
            with col_stat3:
                st.metric("Différence", f"{difference:,}",
                         delta=f"{(difference/total_avant*100):.1f}%" if total_avant > 0 else "0%",
                         delta_color="inverse")

            if not df_filtre.empty:
                pvt_count = df_filtre['PVT'].nunique() if 'PVT' in df_filtre.columns else 0
                vendeurs_count = df_filtre['LOGIN'].nunique() if 'LOGIN' in df_filtre.columns else 0

                st.success(f"✅ {total_apres:,} ventes analysées | {pvt_count} PVT | {vendeurs_count} vendeurs")

                st.markdown("<br>", unsafe_allow_html=True)
                st.markdown('<div class="section-title">📥 TÉLÉCHARGEMENT DU RAPPORT</div>', unsafe_allow_html=True)

                try:
                    objectif_mensuel = 960
                    excel_buffer = generate_monthly_excel_report(
                        df_filtre, mois_nom, selected_year, objectif_mensuel
                    )

                    file_name = f"Reporting_Mensuel_{mois_nom}_{selected_year}.xlsx"

                    st.download_button(
                        label=f"📥 Télécharger le Reporting Mensuel complet",
                        data=excel_buffer.getvalue(),
                        file_name=file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="download_monthly_report",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"❌ Erreur lors de la génération du fichier Excel : {str(e)}")
                    import traceback
                    st.code(traceback.format_exc())
            else:
                st.warning("⚠️ Aucune donnée filtrée disponible. Vérifiez votre fichier source.")

        except Exception as e:
            st.error(f"❌ Erreur lors du traitement des données : {str(e)}")
            import traceback
            st.code(traceback.format_exc())

if __name__ == "__main__":
    main()