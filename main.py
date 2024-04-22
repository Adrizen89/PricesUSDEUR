import sys
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QApplication, QFileDialog, QMessageBox, QInputDialog, QProgressBar
from PyQt5.QtCore import QTimer
import configparser
import os
import pandas as pd
from openpyxl import load_workbook
import subprocess
from functions_scrapping import ZLME_scrap, EURX_scrap
import locale
from datetime import datetime, timedelta, date
from dateutil.easter import easter
import datetime


locale.setlocale(locale.LC_ALL, 'fr_FR.UTF-8')
now = datetime.datetime.now().date()
yesterday = now - timedelta(days=1)

# Ajd 'vendredi'
day_of_week = now.strftime("%A")

# Hier '01/06/2023
date_yesterday = yesterday.strftime("%d/%m/%Y")

# Hier 'jeudi'
yesterday_day_of_week = yesterday.strftime("%A")

# Hier 'jeudi 01 juin 2023
yesterday_holiday = yesterday.strftime("%A %d %B")


# Lire le chemin du fichier à partir du fichier config.ini
script_dir = os.path.dirname(os.path.abspath(__file__))
config_path = os.path.join(script_dir, 'config.ini')
config = configparser.ConfigParser()
if os.path.exists(config_path):
    config.read(config_path)
    default_path = config.get('SETTINGS', 'path', fallback="")
else:
    default_path = ""

class MyApp(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        
        self.initUI()

        self.config = configparser.ConfigParser()
        self.config.read('config.ini')
        auto_start = self.config.getboolean('SETTINGS', 'auto_start', fallback=False)
        self.param1_checkbox.setChecked(auto_start)

        if auto_start:
            self.run_script()
            QtCore.QTimer.singleShot(120000, self.stop_script)

        
    def initUI(self):
        # Création de l'interface
        layout = QtWidgets.QVBoxLayout()
        
        
        # Chemin d'accès
        self.path_edit = QtWidgets.QLineEdit(default_path)
        self.path_edit.setReadOnly(True)
        layout.addWidget(self.path_edit)
        
        # Boutons Modifier et Ouvrir
        button_layout = QtWidgets.QHBoxLayout()
        self.modify_button = QtWidgets.QPushButton('Modifier')
        self.open_button = QtWidgets.QPushButton('Ouvrir')
        button_layout.addWidget(self.modify_button)
        button_layout.addWidget(self.open_button)
        layout.addLayout(button_layout)

        self.start_date_edit = QtWidgets.QDateEdit(datetime.datetime.now().date())
        self.end_date_edit = QtWidgets.QDateEdit(datetime.datetime.now().date())

        # Champs choix des dates
        self.start_date_edit.setCalendarPopup(True)
        self.end_date_edit.setCalendarPopup(True)

        layout.addWidget(QtWidgets.QLabel("Date de début :"))
        layout.addWidget(self.start_date_edit)
        layout.addWidget(QtWidgets.QLabel("Date de fin :"))
        layout.addWidget(self.end_date_edit)

        
        # Logger
        self.logger = QtWidgets.QTextEdit()
        self.logger.setReadOnly(True)
        layout.addWidget(self.logger)
        
        # Bouton Lancer
        self.run_button = QtWidgets.QPushButton('Lancer')
        layout.addWidget(self.run_button)
        
        # Connecter les boutons aux fonctions
        self.modify_button.clicked.connect(self.modify_path)
        self.open_button.clicked.connect(self.open_file)
        self.run_button.clicked.connect(self.run_script)
        
        self.setLayout(layout)

        # Création de la section Paramètres
        self.settings_group = QtWidgets.QGroupBox("Paramètres")
        settings_layout = QtWidgets.QVBoxLayout()
        # Créez l'objet QCheckBox avant de l'ajouter au layout
        self.use_date = QtWidgets.QCheckBox("Utiliser la date du jour")
        settings_layout.addWidget(self.use_date)
        # Ajout de différents widgets pour les paramètres
        self.param1_checkbox = QtWidgets.QCheckBox("Lancer le script automatiquement au démarrage de l'application.")
        self.param1_checkbox.stateChanged.connect(self.saveSettings)
        # Ajout des widgets au layout des paramètres
        settings_layout.addWidget(self.param1_checkbox)
        # Définition du layout des paramètres comme layout du QGroupBox
        self.settings_group.setLayout(settings_layout)
        # Ajout du QGroupBox au layout principal
        layout.addWidget(self.settings_group)
        
        # Paramètres de la fenêtre
        self.setWindowTitle('Prices USD/EUR')
        self.show()

        self.update_run_button_status(day_of_week)

    def update_run_button_status(self, day):
        if day in ['samedi', 'dimanche']:
            self.run_button.setEnabled(False)
        else:
            self.run_button.setEnabled(True)

    def modify_path(self):
        # Fonction pour modifier le chemin d'accès
        file_dialog = QFileDialog()
        path = file_dialog.getOpenFileName(self, 'Sélectionner un fichier Excel', '', 'Excel Files (*.xlsx *.xls)')[0]
        if path:
            self.path_edit.setText(path)
            # Sauvegarder le chemin dans config.ini
            config['SETTINGS'] = {'path': path}
            with open('config.ini', 'w') as configfile:
                config.write(configfile)
            self.log('Chemin modifié.')

    def open_file(self):
        # Fonction pour ouvrir le fichier
        try:
            subprocess.run(["start", default_path], shell=True, check=True)
            self.log('Fichier ouvert.')
        except subprocess.CalledProcessError as e:
            self.log('Fichier non trouvé.')
    
    def saveSettings(self):
        self.config.read('config.ini')
    
        # Mettez à jour seulement la clé spécifique
        if not self.config.has_section('SETTINGS'):
            self.config.add_section('SETTINGS')
        self.config.set('SETTINGS', 'auto_start', str(self.param1_checkbox.isChecked()))
        
        # Écrivez le fichier de configuration mis à jour
        with open('config.ini', 'w') as configfile:
            self.config.write(configfile)

    def run_script(self):

        # Fonction pour lancer le script
        self.log('Script lancé.')
        # Vérifiez l'état de la checkbox pour déterminer les dates à utiliser
        if self.use_date.isChecked():
            start_date = end_date = datetime.date.today()
        else:
            start_date = self.start_date_edit.date().toPyDate()
            end_date = self.end_date_edit.date().toPyDate()
        
        # #  # # # # ZLME # # # # # # 

        try:
            result_ZLME = ZLME_scrap(start_date, end_date)
            
            # Vérifiez si une erreur s'est produite en vérifiant si le résultat est une chaîne
            if isinstance(result_ZLME, str):
                self.log(f"Erreur lors du scrapping ZLME : {result_ZLME}")
                return
            
            # Log chaque entrée dans result_ZLME
            for date_ZLME, data_ZLME in result_ZLME:
                self.log(f"ZLME = {date_ZLME} : {data_ZLME}")
        except Exception as e:
            self.log(f"Erreur lors du scrapping ZLME : {e}")
            return
        
        # # # # # # EURX # # # # # # #

        try:
            result_EURX = EURX_scrap(start_date, end_date)
            
            # Vérifiez si une erreur s'est produite en vérifiant si le résultat est une chaîne
            if isinstance(result_EURX, str):
                self.log(f"Erreur lors du scrapping EURX : {result_EURX}")
                return
            
            # Log chaque entrée dans result_EURX
            for date_EURX, data_EURX in result_EURX:
                self.log(f"EURX = {date_EURX} : {data_EURX}")
        except Exception as e:
            self.log(f"Erreur lors du scrapping EURX : {e}")
            return

        # Logger les données dans l'interface PyQt
        # self.log(f"Date: {date_ZLME}, Valeur: {data_ZLME}")
        # self.log(f"Date: {date_EURX}, Valeur: {data_EURX}")

        # Chemin vers le fichier Excel
        excel_path = self.path_edit.text()
        # "L:/ECHANGE/SAP/Metals_Prices/pricesUSDEUR.xlsx"
        
        if not os.path.exists(excel_path):
            self.log(f"Fichier '{excel_path}' non trouvé.")
            return

        try:
            # Charger le fichier Excel existant
            book = load_workbook(excel_path)
        except Exception as e:
            self.log(f'Erreur lors du chargement du fichier Excel : {e}')

        try:    
            # Obtenir la feuille de calcul 'ZLME'
            sheet = book['ZLME']

            # Vérifiez si result_ZLME est une liste de tuples
            if isinstance(result_ZLME, list) and all(isinstance(i, tuple) for i in result_ZLME):
                for date, value in result_ZLME:
                    # Trouver la dernière ligne et ajouter les nouvelles données
                    new_row = ['ZLME', date, value, 'USD', 'EUR']
                    # Trouver la dernière ligne disponible et ajouter la nouvelle ligne
                    sheet.append(new_row)
            else:
                self.log(f"Erreur lors du scrapping ZLME : {result_ZLME}")
                return
        except Exception as e:
            self.log(f'Erreur lors l\'écriture dans la feuille ZLME : {e}')

        try:    
            # Obtenir la feuille de calcul 'EURX'
            sheet = book['EURX']

            # Vérifiez si result_EURX est une liste de tuples
            if isinstance(result_EURX, list) and all(isinstance(i, tuple) for i in result_EURX):
                for date, value in result_EURX:
                    # Trouver la dernière ligne et ajouter les nouvelles données
                    new_row = ['EURX', date, value, 'USD', 'EUR']
                    # Trouver la dernière ligne disponible et ajouter la nouvelle ligne
                    sheet.append(new_row)
            else:
                self.log(f"Erreur lors du scrapping ZLME : {result_EURX}")
                return
        except Exception as e:
            self.log(f'Erreur lors l\'écriture dans la feuille ZLME : {e}')

        try:
            # Sauvegarder le fichier
            book.save(excel_path)
            self.log("Les données ont été ajoutées avec succès.")
        except Exception as e:
            self.log(f"Erreur lors de la sauvegarde du fichier Excel : {e}")

    def log(self, message):
        # Fonction pour log les messages
        self.logger.append(message)

    def stop_script(self):
        self.log("Le script s'arrête...")
        QApplication.instance().quit()

if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    ex = MyApp()
    sys.exit(app.exec_())
