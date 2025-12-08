import sqlite3
import pandas as pd
from datetime import datetime
import os

class ReportingDatabase:
    def __init__(self, db_path="louma_reporting.db"):
        """Initialise la connexion à la base de données"""
        self.db_path = db_path
        self.init_database()

    def init_database(self):
        """Crée les tables si elles n'existent pas"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()

        # Table pour les PVT
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS pvt (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                nom TEXT UNIQUE NOT NULL,
                contact TEXT NOT NULL,
                date_creation TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                date_modification TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        # Table pour les reporting daily
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS reporting_daily (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date DATE NOT NULL,
                pvt_nom TEXT NOT NULL,
                zone TEXT NOT NULL,
                commune TEXT,
                site TEXT,
                nb_1g INTEGER DEFAULT 0,
                nb_2g INTEGER DEFAULT 0,
                nb_3g INTEGER DEFAULT 0,
                nb_4g INTEGER DEFAULT 0,
                nb_5g INTEGER DEFAULT 0,
                nb_total INTEGER DEFAULT 0,
                date_creation TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (pvt_nom) REFERENCES pvt(nom)
            )
        ''')

        # Table pour les reporting weekly
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS reporting_weekly (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                semaine TEXT NOT NULL,
                pvt_nom TEXT NOT NULL,
                zone TEXT NOT NULL,
                commune TEXT,
                site TEXT,
                nb_1g INTEGER DEFAULT 0,
                nb_2g INTEGER DEFAULT 0,
                nb_3g INTEGER DEFAULT 0,
                nb_4g INTEGER DEFAULT 0,
                nb_5g INTEGER DEFAULT 0,
                nb_total INTEGER DEFAULT 0,
                date_creation TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (pvt_nom) REFERENCES pvt(nom)
            )
        ''')

        conn.commit()
        conn.close()

    # ============ MÉTHODES POUR LES PVT ============

    def get_all_pvt(self):
        """Récupère tous les PVT"""
        conn = sqlite3.connect(self.db_path)
        df = pd.read_sql_query("SELECT * FROM pvt ORDER BY nom", conn)
        conn.close()
        return df

    def save_pvt(self, nom, contact):
        """Sauvegarde un nouveau PVT"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute(
                "INSERT INTO pvt (nom, contact) VALUES (?, ?)",
                (nom, contact)
            )
            conn.commit()
            conn.close()
            return True, f"PVT '{nom}' ajouté avec succès!"
        except sqlite3.IntegrityError:
            return False, f"Le PVT '{nom}' existe déjà!"
        except Exception as e:
            return False, f"Erreur: {str(e)}"

    def update_pvt(self, old_nom, new_nom, new_contact):
        """Met à jour un PVT existant"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute(
                """UPDATE pvt
                   SET nom = ?, contact = ?, date_modification = CURRENT_TIMESTAMP
                   WHERE nom = ?""",
                (new_nom, new_contact, old_nom)
            )
            conn.commit()
            conn.close()
            return True, f"PVT mis à jour avec succès!"
        except sqlite3.IntegrityError:
            return False, f"Le nom '{new_nom}' est déjà utilisé!"
        except Exception as e:
            return False, f"Erreur: {str(e)}"

    def delete_pvt(self, nom):
        """Supprime un PVT"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM pvt WHERE nom = ?", (nom,))
            conn.commit()
            conn.close()
            return True, f"PVT '{nom}' supprimé avec succès!"
        except Exception as e:
            return False, f"Erreur: {str(e)}"

    # ============ MÉTHODES POUR REPORTING DAILY ============

    def save_reporting_daily(self, data_dict):
        """Sauvegarde un reporting daily"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute(
                """INSERT INTO reporting_daily
                   (date, pvt_nom, zone, commune, site, nb_1g, nb_2g, nb_3g, nb_4g, nb_5g, nb_total)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (data_dict['date'], data_dict['pvt_nom'], data_dict['zone'],
                 data_dict['commune'], data_dict['site'], data_dict['nb_1g'],
                 data_dict['nb_2g'], data_dict['nb_3g'], data_dict['nb_4g'],
                 data_dict['nb_5g'], data_dict['nb_total'])
            )
            conn.commit()
            conn.close()
            return True, "Reporting daily sauvegardé!"
        except Exception as e:
            return False, f"Erreur: {str(e)}"

    def get_reporting_daily(self, start_date=None, end_date=None, pvt_nom=None):
        """Récupère les reporting daily avec filtres optionnels"""
        conn = sqlite3.connect(self.db_path)
        query = "SELECT * FROM reporting_daily WHERE 1=1"
        params = []

        if start_date:
            query += " AND date >= ?"
            params.append(start_date)
        if end_date:
            query += " AND date <= ?"
            params.append(end_date)
        if pvt_nom:
            query += " AND pvt_nom = ?"
            params.append(pvt_nom)

        query += " ORDER BY date DESC"

        df = pd.read_sql_query(query, conn, params=params)
        conn.close()
        return df

    # ============ MÉTHODES POUR REPORTING WEEKLY ============

    def save_reporting_weekly(self, data_dict):
        """Sauvegarde un reporting weekly"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute(
                """INSERT INTO reporting_weekly
                   (semaine, pvt_nom, zone, commune, site, nb_1g, nb_2g, nb_3g, nb_4g, nb_5g, nb_total)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                (data_dict['semaine'], data_dict['pvt_nom'], data_dict['zone'],
                 data_dict['commune'], data_dict['site'], data_dict['nb_1g'],
                 data_dict['nb_2g'], data_dict['nb_3g'], data_dict['nb_4g'],
                 data_dict['nb_5g'], data_dict['nb_total'])
            )
            conn.commit()
            conn.close()
            return True, "Reporting weekly sauvegardé!"
        except Exception as e:
            return False, f"Erreur: {str(e)}"

    def get_reporting_weekly(self, semaine=None, pvt_nom=None):
        """Récupère les reporting weekly avec filtres optionnels"""
        conn = sqlite3.connect(self.db_path)
        query = "SELECT * FROM reporting_weekly WHERE 1=1"
        params = []

        if semaine:
            query += " AND semaine = ?"
            params.append(semaine)
        if pvt_nom:
            query += " AND pvt_nom = ?"
            params.append(pvt_nom)

        query += " ORDER BY semaine DESC"

        df = pd.read_sql_query(query, conn, params=params)
        conn.close()
        return df