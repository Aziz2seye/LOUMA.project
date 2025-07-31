import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

# 🔑 Charger les credentials
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("sustained-truck-467510-u9-108d76acfef8.json", scope)
client = gspread.authorize(creds)

# 📄 Ouvrir la feuille (par son nom)
SHEET_NAME = "Louma_VTO"
sheet = client.open(SHEET_NAME).sheet1

# 🧾 Charger la feuille comme DataFrame
def load_vto_from_sheet():
    data = sheet.get_all_records()
    return pd.DataFrame(data)

# ➕ Ajouter un VTO
def add_vto_to_sheet(login, prenom, nom):
    sheet.append_row([login, prenom, nom])
