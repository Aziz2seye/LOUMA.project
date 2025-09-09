import pandas as pd


DATA_PATH = "vto_list.xlsx"
def load_vto():
    try:
        return pd.read_excel(DATA_PATH, sheet_name="vto")
    except FileNotFoundError:
        return pd.DataFrame(columns=["LOGIN", "PRENOM", "NOM"])
    

def load_pvt():
    try:
        return pd.read_excel(DATA_PATH, sheet_name="pvt")
    except FileNotFoundError:
        return pd.DataFrame(columns=["DRV", "PVT", "CONTACT"])



