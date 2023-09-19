import sys
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QTimer
import configparser
import os
import pandas as pd
from openpyxl import load_workbook
import subprocess
from functions_scrapping import ZLME_scrap, EURX_scrap

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

        self.startup_timer = QTimer()
        self.startup_timer.singleShot(0, self.run_script)

        self.shutdown_timer = QTimer()
        self.shutdown_timer.singleShot(60000, self.stop_script)
        
        self.initUI()
        
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
        
        # Logger
        self.logger = QtWidgets.QTextEdit()
        layout.addWidget(self.logger)
        
        # Bouton Lancer
        self.run_button = QtWidgets.QPushButton('Lancer')
        layout.addWidget(self.run_button)
        
        # Connecter les boutons aux fonctions
        self.modify_button.clicked.connect(self.modify_path)
        self.open_button.clicked.connect(self.open_file)
        self.run_button.clicked.connect(self.run_script)
        
        self.setLayout(layout)
        
        # Paramètres de la fenêtre
        self.setWindowTitle('Prices USD/EUR')
        self.show()
        
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
            

    def run_script(self):
        # Fonction pour lancer le script
        self.log('Script lancé.')
        try:
            result_ZLME = ZLME_scrap()
            if len(result_ZLME) == 2:
                date_ZLME, data_ZLME = result_ZLME
            else:
                self.log(f"Erreur lors du scrapping ZLME : {result_ZLME}")
                return
        except Exception as e:
            self.log(f"Erreur lors du scrapping ZLME : {e}")
            return

        try:
            result_EURX = EURX_scrap()
            if len(result_EURX) == 2:
                date_EURX, data_EURX= result_EURX
            else:
                self.log(f"Erreur lors du scrapping EURX : {result_EURX}")

        except Exception as e:
            self.log(f'Erreur lors du scrapping EURX : {e}')
            return

        # Logger les données dans l'interface PyQt
        self.log(f"Date: {date_ZLME}, Valeur: {data_ZLME}")
        self.log(f"Date: {date_EURX}, Valeur: {data_EURX}")

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

            # Trouver la dernière ligne et ajouter les nouvelles données
            new_row = ['ZLME', date_ZLME, data_ZLME, 'USD', 'EUR']
            
            # Trouver la dernière ligne disponible et ajouter la nouvelle ligne
            sheet.append(new_row)
        except Exception as e:
            self.log(f'Erreur lors l\'écriture dans la feuille ZLME : {e}')

        try:
            # Obtenir la feuille de calcul 'EURX'
            sheet_eurx = book['EURX']

            # Ajouter les données dans la feuille 'EURX'
            new_row_eurx = ['EURX', date_EURX, data_EURX, 'USD', 'EUR']
            sheet_eurx.append(new_row_eurx)
        except Exception as e:
            self.log(f'Erreur de l\'écriture dans la feuille EURX : {e}')

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
