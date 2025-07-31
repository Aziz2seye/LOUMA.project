import pandas as pd


DATA_PATH = "vto_list.xlsx"
def load_vto():
    try:
        return pd.read_excel(DATA_PATH)
    except FileNotFoundError:
        return pd.DataFrame(columns=["LOGIN", "PRENOM", "NOM"])
