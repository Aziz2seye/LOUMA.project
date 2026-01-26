import streamlit as st
import pandas as pd
from io import BytesIO
import xlsxwriter
from datetime import datetime

# Configuration page
st.set_page_config(page_title="LOUMA - Reporting Daily", layout="wide", initial_sidebar_state="expanded")

# --- CSS PERSONNALISÉ (Identique au Weekly) ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;700&display=swap');
    header[data-testid="stHeader"] { display: none; }
    .main { font-family: 'Poppins', sans-serif; background: linear-gradient(135deg, #fff5f0 0%, #ffffff 50%, #f0f8ff 100%); }
    section[data-testid="stSidebar"] { background: linear-gradient(180deg, #FF7900 0%, #FF5000 100%) !important; }
    section[data-testid="stSidebar"] * { color: white !important; }
    .section-title {
        background: linear-gradient(135deg, #FF7900 0%, #FF5000 100%);
        color: white; padding: 0.7rem; border-radius: 10px;
        font-weight: 600; text-align: center; max-width: 500px; margin: auto;
    }
    .stDownloadButton > button {
        background: linear-gradient(135deg, #00D4AA 0%, #00B890 100%);
        color: white; border-radius: 10px; width: 100%; font-weight: 600;
    }
</style>
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

def generate_daily_excel_report(df_final, date_str, objectif_pvt=40):
    """Génère un fichier Excel Daily structuré comme le Weekly avec calculs de Totaux"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        workbook = writer.book

        # --- FORMATS (Centrage et Styles) ---
        h_fmt = workbook.add_format({'bold': True, 'bg_color': '#FF6600', 'font_color': 'white', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
        dr_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'left'})
        dr_num_fmt = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1, 'align': 'center'})
        pvt_fmt = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1, 'indent': 1})
        pvt_num_fmt = workbook.add_format({'bold': True, 'bg_color': '#F2F2F2', 'border': 1, 'align': 'center'})
        vendeur_fmt = workbook.add_format({'border': 1, 'align': 'center', 'font_size': 9})
        total_fmt = workbook.add_format({'bold': True, 'bg_color': '#FF6600', 'font_color': 'white', 'border': 1, 'align': 'center'})

        # --- FEUILLE 1: SYNTHESE DR DAILY ---
        ws1 = workbook.add_worksheet('SYNTHESE DR DAILY')
        headers = ['DR', 'REALISATION', 'OBJECTIF', 'R/O (%)']
        for c, h in enumerate(headers): ws1.write(0, c, h, h_fmt)

        res_dr = df_final.groupby('DR').size().reset_index(name='REALISATION')
        pvt_dr = df_final.groupby('DR')['PVT'].nunique().reset_index(name='NB_PVT')
        res_dr = res_dr.merge(pvt_dr, on='DR')
        res_dr['OBJECTIF'] = res_dr['NB_PVT'] * objectif_pvt
        res_dr['RO'] = (res_dr['REALISATION'] / res_dr['OBJECTIF'] * 100).round(0)

        row = 1
        for i, r in res_dr.iterrows():
            ws1.write(row, 0, r['DR'], dr_fmt)
            ws1.write(row, 1, int(r['REALISATION']), dr_num_fmt)
            ws1.write(row, 2, int(r['OBJECTIF']), dr_num_fmt)
            ws1.write(row, 3, f"{int(r['RO'])}%", dr_num_fmt)
            row += 1

        # Ligne Total Feuille 1
        t_real = res_dr['REALISATION'].sum()
        t_obj = res_dr['OBJECTIF'].sum()
        t_ro = round((t_real / t_obj * 100), 0) if t_obj > 0 else 0
        ws1.write(row, 0, 'TOTAL', total_fmt)
        ws1.write(row, 1, t_real, total_fmt)
        ws1.write(row, 2, t_obj, total_fmt)
        ws1.write(row, 3, f"{int(t_ro)}%", total_fmt)

        # --- FEUILLE 2: REPORTING DR-PVT ---
        ws2 = workbook.add_worksheet('REPORTING DR-PVT')
        for c, h in enumerate(headers): ws2.write(0, c, h, h_fmt)

        curr = 1
        for dr, dr_group in df_final.groupby('DR'):
            nb_pvt_dr = dr_group['PVT'].nunique()
            ws2.write(curr, 0, dr, dr_fmt)
            ws2.write(curr, 1, len(dr_group), dr_num_fmt)
            ws2.write(curr, 2, nb_pvt_dr * objectif_pvt, dr_num_fmt)
            curr += 1
            for pvt, pvt_group in dr_group.groupby('PVT'):
                ws2.write(curr, 0, pvt, pvt_fmt)
                ws2.write(curr, 1, len(pvt_group), pvt_num_fmt)
                ws2.write(curr, 2, objectif_pvt, pvt_num_fmt)
                ws2.write(curr, 3, f"{int(len(pvt_group)/objectif_pvt*100)}%", pvt_num_fmt)
                curr += 1

        ws2.write(curr, 0, 'TOTAL', total_fmt)
        ws2.write(curr, 1, t_real, total_fmt)
        ws2.write(curr, 2, t_obj, total_fmt)
        ws2.write(curr, 3, f"{int(t_ro)}%", total_fmt)

        # --- FEUILLE 3: REPORTING DR-PVT-VENDEURS ---
        ws3 = workbook.add_worksheet('REPORTING DR-PVT-VENDEURS')
        headers_v = ['DR/PVT/VENDEUR', 'Prénom', 'Nom', 'LOGIN', 'REALISATION']
        for c, h in enumerate(headers_v): ws3.write(0, c, h, h_fmt)

        curr_v = 1
        for dr, dr_group in df_final.groupby('DR'):
            ws3.write(curr_v, 0, dr, dr_fmt)
            ws3.write(curr_v, 4, len(dr_group), dr_num_fmt)
            curr_v += 1
            for pvt, pvt_group in dr_group.groupby('PVT'):
                ws3.write(curr_v, 0, pvt, pvt_fmt)
                ws3.write(curr_v, 4, len(pvt_group), pvt_num_fmt)
                curr_v += 1
                vendeurs = pvt_group.groupby(['PRENOM_VENDEUR', 'NOM_VENDEUR', 'LOGIN']).size().reset_index(name='REALISATION')
                for _, v in vendeurs.sort_values('REALISATION', ascending=False).iterrows():
                    ws3.write(curr_v, 0, 'VENDEUR', vendeur_fmt)
                    ws3.write(curr_v, 1, v['PRENOM_VENDEUR'], vendeur_fmt)
                    ws3.write(curr_v, 2, v['NOM_VENDEUR'], vendeur_fmt)
                    ws3.write(curr_v, 3, v['LOGIN'], vendeur_fmt)
                    ws3.write(curr_v, 4, int(v['REALISATION']), vendeur_fmt)
                    curr_v += 1

        ws3.write(curr_v, 0, 'TOTAL', total_fmt)
        ws3.write(curr_v, 4, t_real, total_fmt)

        # Ajustement colonnes
        ws1.set_column('A:D', 20)
        ws2.set_column('A:A', 45); ws2.set_column('B:D', 15)
        ws3.set_column('A:A', 30); ws3.set_column('B:D', 20); ws3.set_column('E:E', 15)

    output.seek(0)
    return output

def main():
    st.markdown('<div class="section-title">📅 REPORTING DAILY - LOUMA</div>', unsafe_allow_html=True)
    selected_date = st.date_input("Choisir la date", datetime.now())

    uploaded_file = st.file_uploader("📁 Importer le fichier CSV ou Excel", type=["xlsx", "csv"])

    if uploaded_file:
        try:
            # Détection séparateur pour CSV
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, sep='|', encoding='utf-8')
                if df.shape[1] <= 1:
                    uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, sep=';', encoding='utf-8')
            else:
                df = pd.read_excel(uploaded_file)

            # Nettoyage et Renommage
            df.columns = [c.strip() for c in df.columns]
            column_mapping = {
                'MSISDN': 'REALISATION', 'ACCUEIL_VENDEUR': 'PVT',
                'LOGIN_VENDEUR': 'LOGIN', 'AGENCE_VENDEUR': 'DR',
                'NOM_VENDEUR': 'NOM_VENDEUR', 'PRENOM_VENDEUR': 'PRENOM_VENDEUR',
                'ETAT_IDENTIFICATION': 'ETAT'
            }
            df = df.rename(columns=column_mapping)

            # Filtrage
            details = ["En Cours-Identification", "Identifie", "Identifie Photo"]
            df_filtre = df[df['ETAT'].astype(str).isin(details)].copy()
            df_filtre = df_filtre[df_filtre['DR'].isin(DRV_MAPPING.keys())].copy()
            df_filtre["DR"] = df_filtre["DR"].replace(DRV_MAPPING)

            st.success(f"✅ {len(df_filtre)} lignes validées pour le {selected_date.strftime('%d/%m/%Y')}")

            # Export Excel
            excel_data = generate_daily_excel_report(df_filtre, selected_date.strftime("%d/%m/%Y"), 40)

            st.download_button(
                label=f"📥 Télécharger le Reporting Daily du {selected_date.strftime('%d/%m/%Y')}",
                data=excel_data,
                file_name=f"Reporting_Daily_{selected_date.strftime('%Y_%m_%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"Erreur : {e}")

if __name__ == "__main__":
    main()