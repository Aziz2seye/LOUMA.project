import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import tempfile
# 📁 Structure du projet :
# - main.py (point d'entrée)
# - pages/
#   - 1_Gestion_VTO.py
#   - 2_Reporting.py
#   - 3_Reporting_Mensuel.py
#   - data/
#     - vto_list.csv

# ====================
# main.py
# ====================
import streamlit as st

st.set_page_config(page_title="LOUMA - Accueil", layout="wide")

st.title("📦 Automatisation des Reportings et paiements des commissions LOUMAS - Accueil")

st.write("Bienvenue dans l'application.")

st.write("Avec les animations LOUMAS, Renforçons notre présence" \
" commerciale dans les Loumas en y établissant un point de rencontre dédié aux clients. ")

st.write("Utilisez le menu sur la gauche pour naviguer entre les pages :")

st.markdown("""
- 🧍 **Gestion des VTO** : Ajouter, modifier ou supprimer les VTO
- 📑 **Gestion des PVT** : Ajouter, modifier ou supprimer les PVT            
- 📊 **Reporting Daily & Weekly** : Choisir entre reporting journalier ou hebdomadaire 
- 💰 **Reporting Mensuel** : Accéder aux fonctionnalités de reporting du mois
- 💰 **Paiement des commissions** : Payer automatiquement les VTO & PVT
""")

