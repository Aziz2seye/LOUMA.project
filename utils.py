import pandas as pd
import os

# Chemins des fichiers
PVT_DATA_PATH = "pvt_list.xlsx"
VTO_DATA_PATH = "vto_list.xlsx"

def load_pvt():
    """Charge la liste des PVT depuis le fichier Excel"""
    if os.path.exists(PVT_DATA_PATH):
        try:
            return pd.read_excel(PVT_DATA_PATH)
        except Exception as e:
            print(f"Erreur lors du chargement des PVT: {e}")
            return pd.DataFrame(columns=["PVT", "CONTACT"])
    else:
        return pd.DataFrame(columns=["PVT", "CONTACT"])

def load_vto():
    """Charge la liste des VTO depuis le fichier Excel"""
    if os.path.exists(VTO_DATA_PATH):
        try:
            # Essayer de lire le fichier sans spécifier de sheet_name
            df = pd.read_excel(VTO_DATA_PATH)

            # Vérifier si les colonnes nécessaires existent
            required_columns = ["DRV", "PRENOM_VENDEUR", "NOM_VENDEUR", "PVT", "LOGIN", "KABBU"]

            # Si les colonnes n'existent pas toutes, créer un DataFrame vide avec les bonnes colonnes
            if not all(col in df.columns for col in required_columns):
                return pd.DataFrame(columns=required_columns)

            return df
        except Exception as e:
            print(f"Erreur lors du chargement des VTO: {e}")
            return pd.DataFrame(columns=["DRV", "PRENOM_VENDEUR", "NOM_VENDEUR", "PVT", "LOGIN", "KABBU"])
    else:
        # Si le fichier n'existe pas, retourner un DataFrame vide avec les colonnes appropriées
        return pd.DataFrame(columns=["DRV", "PRENOM_VENDEUR", "NOM_VENDEUR", "PVT", "LOGIN", "KABBU"])

def load_vto2():
    """Charge la liste des VTO2 depuis le fichier Excel"""
    VTO2_DATA_PATH = "vto2_list.xlsx"
    if os.path.exists(VTO2_DATA_PATH):
        try:
            df = pd.read_excel(VTO2_DATA_PATH)
            required_columns = ["DRV", "PRENOM_VENDEUR", "NOM_VENDEUR", "PVT", "LOGIN", "KABBU"]

            if not all(col in df.columns for col in required_columns):
                return pd.DataFrame(columns=required_columns)

            return df
        except Exception as e:
            print(f"Erreur lors du chargement des VTO2: {e}")
            return pd.DataFrame(columns=["DRV", "PRENOM_VENDEUR", "NOM_VENDEUR", "PVT", "LOGIN", "KABBU"])
    else:
        return pd.DataFrame(columns=["DRV", "PRENOM_VENDEUR", "NOM_VENDEUR", "PVT", "LOGIN", "KABBU"])