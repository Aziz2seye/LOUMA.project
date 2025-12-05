"""
Gestionnaire de base de données pour l'application de reporting
Fichier: db_manager.py
"""

import sqlite3
import pandas as pd
from datetime import datetime
from pathlib import Path
import json

class ReportingDatabase:
    """
    Gestionnaire de base de données pour sauvegarder les reportings
    avant l'export Excel
    """

    def _init_(self, db_path="reporting_data.db"):
        """Initialise la connexion à la base de données"""
        self.db_path = db_path
        self.conn = None
        self.create_tables()

    def get_connection(self):
        """Crée ou retourne la connexion à la base de données"""
        if self.conn is None:
            self.conn = sqlite3.connect(self.db_path, check_same_thread=False)
        return self.conn

    def create_tables(self):
        """Crée les tables nécessaires si elles n'existent pas"""
        conn = self.get_connection()
        cursor = conn.cursor()

        # Table pour les reportings journaliers - Résumé PVT
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS reporting_journalier_pvt (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date_reporting DATE NOT NULL,
                date_creation TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                drv TEXT NOT NULL,
                pvt TEXT NOT NULL,
                total_sim INTEGER NOT NULL,
                objectif INTEGER NOT NULL,
                taux_realisation TEXT NOT NULL,
                UNIQUE(date_reporting, drv, pvt)
            )
        ''')

        # Table pour les reportings journaliers - Détails VTO
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS reporting_journalier_vto (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date_reporting DATE NOT NULL,
                date_creation TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                drv TEXT NOT NULL,
                pvt TEXT NOT NULL,
                prenom_vendeur TEXT NOT NULL,
                nom_vendeur TEXT NOT NULL,
                login TEXT NOT NULL,
                total_sim INTEGER NOT NULL,
                UNIQUE(date_reporting, login)
            )
        ''')

        # Table pour les reportings hebdomadaires - Résumé PVT
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS reporting_hebdo_pvt (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                semaine TEXT NOT NULL,
                date_debut DATE NOT NULL,
                date_fin DATE NOT NULL,
                date_creation TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                drv TEXT NOT NULL,
                pvt TEXT NOT NULL,
                total_sim INTEGER NOT NULL,
                objectif INTEGER NOT NULL,
                taux_realisation TEXT NOT NULL,
                UNIQUE(semaine, drv, pvt)
            )
        ''')

        # Table pour les reportings hebdomadaires - Détails VTO
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS reporting_hebdo_vto (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                semaine TEXT NOT NULL,
                date_debut DATE NOT NULL,
                date_fin DATE NOT NULL,
                date_creation TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                drv TEXT NOT NULL,
                pvt TEXT NOT NULL,
                prenom_vendeur TEXT NOT NULL,
                nom_vendeur TEXT NOT NULL,
                login TEXT NOT NULL,
                total_sim INTEGER NOT NULL,
                UNIQUE(semaine, login)
            )
        ''')

        # Table pour l'historique des exports
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS historique_exports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                type_reporting TEXT NOT NULL,
                date_reporting TEXT NOT NULL,
                date_export TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                nombre_pvt INTEGER,
                nombre_vto INTEGER,
                total_ventes INTEGER,
                fichier_nom TEXT,
                statut TEXT DEFAULT 'SUCCESS'
            )
        ''')

        conn.commit()

    def save_daily_report(self, date_reporting, df_pvt_summary, df_reporting):
        """
        Sauvegarde un reporting journalier dans la base de données

        Args:
            date_reporting: Date du reporting (format 'YYYY-MM-DD')
            df_pvt_summary: DataFrame avec le résumé par PVT
            df_reporting: DataFrame avec les détails par VTO
        """
        conn = self.get_connection()
        cursor = conn.cursor()

        try:
            # Supprimer les données existantes pour cette date
            cursor.execute('DELETE FROM reporting_journalier_pvt WHERE date_reporting = ?', (date_reporting,))
            cursor.execute('DELETE FROM reporting_journalier_vto WHERE date_reporting = ?', (date_reporting,))

            # Insérer le résumé PVT (sans la ligne TOTAL)
            df_pvt_clean = df_pvt_summary[df_pvt_summary['PVT'] != 'TOTAL'].copy()
            for _, row in df_pvt_clean.iterrows():
                cursor.execute('''
                    INSERT INTO reporting_journalier_pvt
                    (date_reporting, drv, pvt, total_sim, objectif, taux_realisation)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (date_reporting, row['DRV'], row['PVT'],
                      int(row['TOTAL_SIM']), int(row['OBJECTIF']), row['TR']))

            # Insérer les détails VTO (sans la ligne TOTAL)
            df_vto_clean = df_reporting[df_reporting['LOGIN'] != 'TOTAL'].copy()
            for _, row in df_vto_clean.iterrows():
                cursor.execute('''
                    INSERT INTO reporting_journalier_vto
                    (date_reporting, drv, pvt, prenom_vendeur, nom_vendeur, login, total_sim)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (date_reporting, row['DRV'], row['PVT'],
                      row['PRENOM_VENDEUR'], row['NOM_VENDEUR'],
                      row['LOGIN'], int(row['TOTAL_SIM'])))

            # Enregistrer l'historique
            total_ventes = int(df_vto_clean['TOTAL_SIM'].sum())
            cursor.execute('''
                INSERT INTO historique_exports
                (type_reporting, date_reporting, nombre_pvt, nombre_vto, total_ventes, fichier_nom)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', ('JOURNALIER', date_reporting, len(df_pvt_clean),
                  len(df_vto_clean), total_ventes, f'Daily_Reporting_{date_reporting}.xlsx'))

            conn.commit()
            return True, "Données sauvegardées avec succès"

        except Exception as e:
            conn.rollback()
            return False, f"Erreur lors de la sauvegarde : {str(e)}"

    def save_weekly_report(self, semaine, date_debut, date_fin, df_pvt_summary, df_reporting):
        """
        Sauvegarde un reporting hebdomadaire dans la base de données

        Args:
            semaine: Numéro de la semaine (ex: 'S47-2025')
            date_debut: Date de début de la semaine
            date_fin: Date de fin de la semaine
            df_pvt_summary: DataFrame avec le résumé par PVT
            df_reporting: DataFrame avec les détails par VTO
        """
        conn = self.get_connection()
        cursor = conn.cursor()

        try:
            # Supprimer les données existantes pour cette semaine
            cursor.execute('DELETE FROM reporting_hebdo_pvt WHERE semaine = ?', (semaine,))
            cursor.execute('DELETE FROM reporting_hebdo_vto WHERE semaine = ?', (semaine,))

            # Insérer le résumé PVT
            df_pvt_clean = df_pvt_summary[df_pvt_summary['PVT'] != 'TOTAL'].copy()
            for _, row in df_pvt_clean.iterrows():
                cursor.execute('''
                    INSERT INTO reporting_hebdo_pvt
                    (semaine, date_debut, date_fin, drv, pvt, total_sim, objectif, taux_realisation)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', (semaine, date_debut, date_fin, row['DRV'], row['PVT'],
                      int(row['TOTAL_SIM']), int(row['OBJECTIF']), row['TR']))

            # Insérer les détails VTO
            df_vto_clean = df_reporting[df_reporting['LOGIN'] != 'TOTAL'].copy()
            for _, row in df_vto_clean.iterrows():
                cursor.execute('''
                    INSERT INTO reporting_hebdo_vto
                    (semaine, date_debut, date_fin, drv, pvt, prenom_vendeur, nom_vendeur, login, total_sim)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                ''', (semaine, date_debut, date_fin, row['DRV'], row['PVT'],
                      row['PRENOM_VENDEUR'], row['NOM_VENDEUR'],
                      row['LOGIN'], int(row['TOTAL_SIM'])))

            # Enregistrer l'historique
            total_ventes = int(df_vto_clean['TOTAL_SIM'].sum())
            cursor.execute('''
                INSERT INTO historique_exports
                (type_reporting, date_reporting, nombre_pvt, nombre_vto, total_ventes, fichier_nom)
                VALUES (?, ?, ?, ?, ?, ?)
            ''', ('HEBDOMADAIRE', semaine, len(df_pvt_clean),
                  len(df_vto_clean), total_ventes, f'Weekly_Reporting_{semaine}.xlsx'))

            conn.commit()
            return True, "Données sauvegardées avec succès"

        except Exception as e:
            conn.rollback()
            return False, f"Erreur lors de la sauvegarde : {str(e)}"

    def get_daily_report(self, date_reporting):
        """Récupère un reporting journalier depuis la base de données"""
        conn = self.get_connection()

        df_pvt = pd.read_sql_query('''
            SELECT drv, pvt, total_sim, objectif, taux_realisation as TR
            FROM reporting_journalier_pvt
            WHERE date_reporting = ?
            ORDER BY drv, pvt
        ''', conn, params=(date_reporting,))

        df_vto = pd.read_sql_query('''
            SELECT drv, pvt, prenom_vendeur, nom_vendeur, login, total_sim
            FROM reporting_journalier_vto
            WHERE date_reporting = ?
            ORDER BY drv, pvt, total_sim DESC
        ''', conn, params=(date_reporting,))

        return df_pvt, df_vto

    def get_weekly_report(self, semaine):
        """Récupère un reporting hebdomadaire depuis la base de données"""
        conn = self.get_connection()

        df_pvt = pd.read_sql_query('''
            SELECT drv, pvt, total_sim, objectif, taux_realisation as TR
            FROM reporting_hebdo_pvt
            WHERE semaine = ?
            ORDER BY drv, pvt
        ''', conn, params=(semaine,))

        df_vto = pd.read_sql_query('''
            SELECT drv, pvt, prenom_vendeur, nom_vendeur, login, total_sim
            FROM reporting_hebdo_vto
            WHERE semaine = ?
            ORDER BY drv, pvt, total_sim DESC
        ''', conn, params=(semaine,))

        return df_pvt, df_vto

    def get_export_history(self, limit=50):
        """Récupère l'historique des exports"""
        conn = self.get_connection()

        df_history = pd.read_sql_query('''
            SELECT
                type_reporting,
                date_reporting,
                date_export,
                nombre_pvt,
                nombre_vto,
                total_ventes,
                fichier_nom,
                statut
            FROM historique_exports
            ORDER BY date_export DESC
            LIMIT ?
        ''', conn, params=(limit,))

        return df_history

    def get_available_dates(self, report_type='daily'):
        """Récupère les dates disponibles pour un type de reporting"""
        conn = self.get_connection()
        cursor = conn.cursor()

        if report_type == 'daily':
            cursor.execute('''
                SELECT DISTINCT date_reporting
                FROM reporting_journalier_pvt
                ORDER BY date_reporting DESC
            ''')
        else:
            cursor.execute('''
                SELECT DISTINCT semaine, date_debut, date_fin
                FROM reporting_hebdo_pvt
                ORDER BY date_debut DESC
            ''')

        return cursor.fetchall()

    def get_statistics(self):
        """Récupère des statistiques globales"""
        conn = self.get_connection()

        stats = {}

        # Stats journalières
        cursor = conn.cursor()
        cursor.execute('SELECT COUNT(DISTINCT date_reporting) FROM reporting_journalier_pvt')
        stats['nb_jours'] = cursor.fetchone()[0]

        cursor.execute('SELECT SUM(total_sim) FROM reporting_journalier_pvt')
        result = cursor.fetchone()[0]
        stats['total_ventes_jour'] = result if result else 0

        # Stats hebdomadaires
        cursor.execute('SELECT COUNT(DISTINCT semaine) FROM reporting_hebdo_pvt')
        stats['nb_semaines'] = cursor.fetchone()[0]

        cursor.execute('SELECT SUM(total_sim) FROM reporting_hebdo_pvt')
        result = cursor.fetchone()[0]
        stats['total_ventes_hebdo'] = result if result else 0

        # Meilleurs VTO (tous temps)
        cursor.execute('''
            SELECT prenom_vendeur || ' ' || nom_vendeur as nom_complet,
                   SUM(total_sim) as total
            FROM reporting_journalier_vto
            GROUP BY login
            ORDER BY total DESC
            LIMIT 10
        ''')
        stats['top_vto'] = cursor.fetchall()

        return stats

    def close(self):
        """Ferme la connexion à la base de données"""
        if self.conn:
            self.conn.close()
            self.conn = None