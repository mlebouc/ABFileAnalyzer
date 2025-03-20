from PySide6.QtWidgets import QApplication, QMainWindow, QWidget, QDialog, QVBoxLayout, QLabel, QPushButton, QHBoxLayout, QLineEdit, QCheckBox, QProgressBar, QFileDialog, QMessageBox, QListWidget, QSizePolicy
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QPainter, QPen, QColor, QIcon
from PySide6.QtGui import QPixmap, QFont

from functools import partial  # Importer partial pour passer des arguments aux fonctions
import sys
import os
import pandas as pd
import pyabf
import openpyxl  # Assurez-vous que openpyxl est importé ici
from openpyxl import Workbook
import numpy as np
import re
import pyabf.tools.memtest
import matplotlib
matplotlib.use('QtAgg')
import matplotlib.pyplot as plt
import matplotlib.widgets
from matplotlib.widgets import Button
from matplotlib.backend_bases import MouseEvent
import csv
import os
import pyabf
import pandas as pd
import openpyxl
import re
from scipy import signal
from scipy.signal import find_peaks, butter, lfilter
import pyabf.filter

class SweepNavigator:
    def __init__(self, abf, channel, sweepX_start_courant, sweepX_end_courant, sweepC_at_start_column, peaks_dict, index_start_courant):
        """
        Initialise la fenêtre interactive et les boutons pour naviguer entre les sweeps.
        """
        self.abf = abf
        self.channel = channel
        self.sweep_index = abf.sweepCount - 1  # Index du sweep affiché
        self.sweepX_start_courant = sweepX_start_courant
        self.index_start_courant = index_start_courant
        self.sweepX_end_courant = sweepX_end_courant
        self.sweepC_at_start_column = sweepC_at_start_column
        self.column_name_y = f"{self.sweepC_at_start_column:.2f}pA"
        self.peaks_dict = peaks_dict
        self.peak_times, self.peak_values = self.peaks_dict.get(self.sweep_index, ([], []))
        # Création de la figure et de l'axe
        self.fig, self.ax = plt.subplots(figsize=(8, 6))
        self.fig.subplots_adjust(bottom=0.2)  # Laisser de la place pour les boutons
        self.scatter = None
        # Initialisation du tracé du premier sweep
        self.abf.setSweep(sweepNumber=self.sweep_index, channel=self.channel)
        self.time = self.abf.sweepX
        self.signal = self.abf.sweepY
        self.mask = (self.time >= self.sweepX_start_courant - 0.5) & (self.time <= self.sweepX_end_courant + 0.5)

        self.selected_time = self.time[self.mask]
        self.line, = self.ax.plot(self.time[self.mask], self.signal[self.mask], label=f"Sweep {self.column_name_y}", color=(98/255, 114/255, 164/255))
        self.scatter = self.ax.scatter(self.peaks_dict[self.column_name_y]["Time (s)"], self.peaks_dict[self.column_name_y]["Max Peak Value"], color=(189/255, 147/255, 249/255), label='Spikes')

        self.ax.set_xlabel("Time (s)")
        self.ax.set_ylabel("Amplitude")
        self.ax.set_title(f"Sweep {self.sweep_index} / {self.abf.sweepCount - 1}")
        self.ax.legend()

        # Bouton "Précédent"
        self.axprev = self.fig.add_axes([0.3, 0.05, 0.15, 0.075])
        self.btn_prev = Button(self.axprev, "Précédent")
        self.btn_prev.on_clicked(self.prev_sweep)

        # Bouton "Suivant"
        self.axnext = self.fig.add_axes([0.55, 0.05, 0.15, 0.075])
        self.btn_next = Button(self.axnext, "Suivant")
        self.btn_next.on_clicked(self.next_sweep)

        plt.show(block=False)  # Affichage non bloquant
        self.fig.canvas.mpl_connect('button_press_event', self.on_click)

    def on_click(self, event=None):
        """Active l'interactivité uniquement pour la figure cliquée"""
        for fig_num in plt.get_fignums():  # Liste toutes les figures ouvertes
            fig = plt.figure(fig_num)
            if fig is not self.fig:  # Désactive toutes sauf la figure active
                fig.canvas.mpl_disconnect(fig.canvas.manager.key_press_handler_id)
                for ax in fig.get_axes():  # Désactive les boutons des autres figures
                    for artist in ax.artists:
                        if isinstance(artist, Button):
                            artist.disconnect_events()

        # Réactive les boutons uniquement pour la figure active
        self.btn_prev.on_clicked(self.prev_sweep)
        self.btn_next.on_clicked(self.next_sweep)
        
    def update_plot(self):
        """Met à jour le tracé du sweep courant sans recréer toute la figure."""
                
        self.abf.setSweep(sweepNumber=self.sweep_index, channel=self.channel)
        self.signal = self.abf.sweepY
        if self.scatter is not None:
            try:
                self.scatter.remove()
            except ValueError:
                pass
        
        print(f"Scatter après suppression : {self.scatter}")
        sweepC_at_start_column = self.abf.sweepC[self.index_start_courant]
        self.sweepC_at_start_column = sweepC_at_start_column
        self.column_name_y = f"{self.sweepC_at_start_column:.2f}pA"

        self.line.set_ydata(self.signal[self.mask])  # Mise à jour des données Y uniquement
        self.line.set_label(f"Sweep {self.column_name_y}")  # Met à jour le label
        if self.column_name_y in self.peaks_dict:
            self.scatter = self.ax.scatter(self.peaks_dict[self.column_name_y]["Time (s)"], self.peaks_dict[self.column_name_y]["Max Peak Value"], color=(189/255, 147/255, 249/255), label='Spikes')
        self.ax.set_title(f"Sweep {self.sweep_index} / {self.abf.sweepCount - 1}")
        self.ax.legend()  # Mise à jour de la légende
        self.fig.canvas.draw_idle()  # Rafraîchissement de la figure

    def next_sweep(self, event):
        """Affiche le sweep suivant."""
        if self.sweep_index < self.abf.sweepCount - 1:
            self.sweep_index += 1
            print(f"🔹 Sweep suivant : {self.sweep_index}")
            self.update_plot()

    def prev_sweep(self, event):
        """Affiche le sweep précédent."""
        if self.sweep_index > 0:
            self.sweep_index -= 1
            print(f"🔹 Sweep précédent : {self.sweep_index}")
            self.update_plot()

# 🔹 Fonctions indépendantes (utilitaires)
def get_total_abf_files_in_folders(root_folder, target_folders):
    """Calcule le nombre total de fichiers ABF dans les sous-dossiers."""
    total_files = 0
    for root, dirs, files in os.walk(root_folder):
        for folder_name in dirs:
            if folder_name in target_folders:
                folder_path = os.path.join(root, folder_name)
                total_files += len([f for f in os.listdir(folder_path) if f.endswith('.abf')])
    return total_files

def filter_signal(data, sampling_frequency=20000, cutoff_frequency=2000, order=4):
    """Filtre un signal avec un filtre passe-bas Butterworth."""
    nyquist_frequency = 0.5 * sampling_frequency
    normalized_cutoff = cutoff_frequency / nyquist_frequency
    b, a = butter(order, normalized_cutoff, btype='low', analog=False)
    return lfilter(b, a, data)

# Fonction pour détecter les pics
def detect_peaks_and_calculate_frequency(time, signal, height_threshold=-30):
    """
    Détecte les minima directement dans un signal (sans inversion) et calcule les fréquences instantanées.
    """
    from scipy.signal import find_peaks

    # Détecter les indices des minima significatifs (valeurs plus petites que le threshold)
    peak_indices, _ = find_peaks(-signal, height=-height_threshold)

    # Récupérer les temps et les vraies valeurs des minima
    peak_times = time[peak_indices]  # Temps des minima
    peak_values = signal[peak_indices]  # Valeurs originales des minima

    # Calculer les intervalles inter-spikes (ISI)
    isi = peak_times[1:] - peak_times[:-1] if len(peak_times) > 1 else []

    # Calculer les fréquences instantanées
    frequencies = 1 / isi if len(isi) > 0 else []

    return {
        "Time (s)": peak_times,
        "Peak Values": peak_values,  # Ajout des vraies valeurs des minima
        "ISI": isi,
        "Frequency (Hz)": frequencies
    }

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("ABFile Analyzer")
        self.setGeometry(80, 80, 600, 150)

        # Interface utilisateur
        self.setup_ui()

    def setup_ui(self):
        """Crée l'interface graphique."""
        # Définir le style (fond bleu, texte blanc)
        self.setStyleSheet("""
            QMainWindow {
            background-color: rgb(33, 37, 43);
            }
            QWidget {
                color: rgb(221, 221, 221);
	            font: 10pt "Segoe UI";
            }
            #leftMenuBg {
                background-color: rgb(33, 37, 43);
            }
            #mainContent {
                background-color: rgb(40, 44, 52);
            }
            QLabel#mainTitle {
                font-size: 36px;
                font-weight: bold;
                color: rgb(189, 147, 249);
            }
            QPushButton {
                background-color: rgb(98, 114, 164);
                border: none;
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-size: 12pt;
            }
            QPushButton:hover {
                background-color: rgb(189, 147, 249);
            }
            QLineEdit {
                background-color: rgb(33, 37, 43);
                border: 2px solid rgb(44, 49, 58);
                border-radius: 5px;
                padding: 5px;
                color: rgb(255, 255, 255);
            }
            QLineEdit:hover {
                border: 2px solid rgb(98, 114, 164);
            }
            QDialog {
                background-color: rgb(40, 44, 52);
        """)
        
        main_widget = QWidget()
        self.setCentralWidget(main_widget)

        # Layout principal
        main_layout = QHBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # Barre latérale gauche
        left_menu = QWidget()
        left_menu.setObjectName("leftMenuBg")
        left_layout = QVBoxLayout(left_menu)
        left_layout.setContentsMargins(10, 10, 10, 10)

        # Titre et description dans la barre latérale
        description_label = QLabel(
    "<p align='center' style='color: rgb(189, 147, 249); font-size: 12pt; font-weight: bold;'>"
    "ABFile Analyzer"
    "</p>"
    "<p align='center' style='color: rgb(221, 221, 221); font-size: 10pt; font-weight: normal; margin-top: 15px; margin-bottom: 15px;'>"
    "An Python-base application for analyzing ABF files"
    "</p>"
    "</p>"
    "<p align='center' style='color: rgb(189, 147, 249); font-size: 12pt; font-weight: bold;'>"
    "MIT License"
    "</p>"
    "<p align='center' style='color: rgb(221, 221, 221); font-size: 10pt; font-weight: normal; margin-top: 15px; margin-bottom: 15px;'>"
    "Created by Margaux Lebouc"
    
)
        description_label.setWordWrap(True)  # Activer le retour à la ligne si nécessaire

        
        # Titre et image
        title = QLabel("ABFile Analyzer")
        title.setAlignment(Qt.AlignCenter)
        #pixmap = QPixmap(r"C:\Users\marga\Documents\CodeTestPythonetPyside\neuron-png")  # Remplace par le chemin de ton image
        #pixmap = pixmap.scaled(200, 200, Qt.KeepAspectRatio, Qt.SmoothTransformation)  # Redimensionne en conservant le ratio

        image_label = QLabel()
        #image_label.setPixmap(pixmap)
        
        image_label.setAlignment(Qt.AlignCenter)

        
        left_layout.addSpacing(10)
        left_layout.addWidget(description_label)
        left_layout.addStretch()
        

        bottom_box = QWidget()
        bottom_box_layout = QVBoxLayout()
        bottom_box.setLayout(bottom_box_layout)

        icon_path = r"C:\Users\marga\Documents\CodeTestPythonetPyside\livre"
        notes_button = QPushButton("    User Guide")
        notes_button.clicked.connect(self.notes)
        notes_button.setIcon(QIcon(icon_path))
        notes_button.setIconSize(QSize(20, 20)) 
        notes_button.setStyleSheet("""
            QPushButton {
                background-color: rgb(98, 114, 164);
                border: none;
                color: white;
                font-size: 10pt;
                font-weight: bold;
                text-align: center;
                padding-left: 0px;  /* Aligner le texte à gauche */
                height: 30px;  /* Hauteur du bouton */
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: rgb(189, 147, 249);  /* Changement de couleur au survol */
            }
        """)
        # Ajouter le bouton dans le bottom_box
        bottom_box_layout.setContentsMargins(0, 0, 0, 0)  # Enlève les marges
        bottom_box_layout.addWidget(notes_button)
        # Ajouter bottom_box à la barre latérale gauche
        left_layout.addStretch()  # Place le bouton tout en bas
        left_layout.addWidget(bottom_box)


        # Partie droite principale
        right_content = QWidget()
        right_content.setObjectName("mainContent")
        right_layout = QVBoxLayout(right_content)
        right_layout.setContentsMargins(40, 40, 40, 40)

        # Titre principal
        main_title = QLabel("ABFile Analyzer")
        main_title.setObjectName("mainTitle")
        main_title.setAlignment(Qt.AlignCenter)

        # Barre de sélection de dossier
        browse_layout = QHBoxLayout()
        self.browse_input = QLineEdit()
        self.browse_input.setPlaceholderText("Select a folder...")
        browse_button = QPushButton("Browse")
        browse_button.clicked.connect(self.choose_directory)

        browse_layout.addWidget(self.browse_input)
        browse_layout.addWidget(browse_button)

        # Checkbox et boutons
        self.manual_verification_checkbox = QCheckBox("Display the graphs")
        process_button = QPushButton("Process Files")
        process_button.clicked.connect(self.process_files)
        

        # Barre de progression
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)

        # Ajouter les widgets au layout de droite
        
        right_layout.addWidget(image_label)
        right_layout.addWidget(main_title)
        right_layout.addSpacing(20)
        right_layout.addLayout(browse_layout)
        right_layout.addWidget(self.manual_verification_checkbox)
        right_layout.addWidget(process_button)
        right_layout.addWidget(self.progress_bar)
        right_layout.addStretch()

        # Ajouter la barre latérale gauche et le contenu principal
        main_layout.addWidget(left_menu, 2)  # La barre de gauche prend 2 parts
        main_layout.addWidget(right_content, 5)  # La partie droite prend 5 parts

    def notes(self):
        """Affiche une boîte de dialogue expliquant le fonctionnement de l'application."""
        msg_box = QMessageBox(self)
        msg_box.setWindowTitle("Instructions - ABFile Analyzer")

        # 🔹 Appliquer le fond noir + texte clair
        msg_box.setStyleSheet("""
            QMessageBox {
                background-color: rgb(40, 44, 52); 
                color: rgb(221, 221, 221);
            }
            QLabel {
                color: rgb(221, 221, 221);
                font-size: 10pt;
            }
            QPushButton {
                background-color: rgb(98, 114, 164);
                color: white;
                padding: 8px;
                border-radius: 5px;
                font-size: 10pt;
            }
            QPushButton:hover {
                background-color: rgb(189, 147, 249);
            }
        """)

        msg_box.setText(
            """<h3 style='color:rgb(189, 147, 249);'>ABFile Analyzer - Guide d'utilisation</h3>
            
            <p>Cet outil permet d'analyser des fichiers ABF en effectuant diverses analyses électrophysiologiques :</p>
            
            <ul>
            <li><b>AMPA/NMDA</b> : Mesure des courants AMPA et NMDA</li>
            <li><b>Stimulation</b> : Détection des courants liés à une stimulation</li>
            <li><b>Rhéobase</b> : Calcul de la rhéobase</li>
            <li><b>Cell-attached</b> : Détection des potentiels d'action</li>
            <li><b>Capacitance</b> : Mesure de la capacitance et de la résistance</li>
            <li><b>Potentiel de repos (Em)</b> : Détection des spikes spontanés et mesure du potentiel de repos</li>
            </ul>

            <p><b>Utilisation :</b></p>
            <ol>
            <li>Sélectionnez un dossier contenant des fichiers ABF via le bouton <b>"Browse"</b>.</li>
            <li>Cochez <b>"Display the graphs"</b> si vous souhaitez voir les courbes après analyse.</li>
            <li>Cliquez sur <b>"Process Files"</b> pour lancer l'analyse.</li>
            <li>Une fois terminé, un fichier Excel sera généré avec les résultats.</li>
            </ol>

            <p style='color:rgb(98, 114, 164);'><b>Note :</b> Pour les fichiers Cell attached, vous pouvez modifier les seuils ou exclure certains sweeps après analyse.</p>
            """
        )

        msg_box.exec()


    def choose_directory(self):
        directory = QFileDialog.getExistingDirectory(self, "Select Folder")
        if directory:
            self.browse_input.setText(directory)

    def show_error_message(self, title, message):
        """
        Affiche un message d'erreur en utilisant PySide6.

        Parameters:
            parent (QWidget): La fenêtre parent (ex: self dans une QMainWindow ou QDialog).
            title (str): Le titre de la boîte de dialogue.
            message (str): Le message d'erreur à afficher.
        """
        QMessageBox.critical(self, title, message)

    def update_progress(progress_var, value):
        value = max(0, min(value, 100))
        progress_var.set(value)
        window.update_idletasks()
        if value < 100:
            window.update()  
  
    def create_continue_window(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Information")
        
        dialog.resize(200, 100)
        dialog.setStyleSheet("""
            QDialog {
            background-color: rgb(40, 44, 52);
            }
            QPushButton {
                background-color: rgb(98, 114, 164);
                border: none;
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-size: 12pt;
            }
            QPushButton:hover {
                background-color: rgb(189, 147, 249);
            }                    
        """)
        layout = QVBoxLayout(dialog)
        label = QLabel("Cliquer pour continuer")
        layout.addWidget(label)
        button_close = QPushButton("Fermer")
        button_close.clicked.connect(lambda: [plt.close('all'), dialog.accept()])
        layout.addWidget(button_close)
        dialog.setLayout(layout)
        dialog.show()

        return dialog

    def recalculate_peaks_and_update_graphs(self, avg_values_dict, signal_data, removed_sweeps, threshold_modified_sweeps, details_folder_path):
        """
        Recalcule les pics, met à jour les statistiques, réaffiche les graphes et sauvegarde les moyennes dans un fichier Excel.

        Parameters:
            avg_values_dict (dict): Contient les thresholds pour chaque sweep.
            signal_data (dict): Contient les données des signaux (filename, sweep_number, time, signal).
            modified_sweeps (list): Liste des sweeps modifiés à mettre à jour.
            output_folder (str): Dossier pour sauvegarder les résultats Excel.
        """
        # Identifier les fichiers affectés
        modified_files = set()
        if threshold_modified_sweeps is None:
            threshold_modified_sweeps = []

        for sweep_key in removed_sweeps + threshold_modified_sweeps:
            filename, _ = sweep_key.rsplit("_Sweep_", 1)
            modified_files.add(filename)

        print("Fichiers modifiés :", modified_files)

        # Fermer les graphes pour les fichiers modifiés
        for fig_num in plt.get_fignums():
            fig = plt.figure(fig_num)
            if fig.canvas.manager.get_window_title() in [f"Figure_{filename}" for filename in modified_files]:
                print(f"Fermeture de la figure : {fig.canvas.manager.get_window_title()}")
                plt.close(fig)

        # Réafficher les graphes et recalculer les stats pour les fichiers affectés
        for filename in modified_files:
            sweeps = [
                int(sweep_key.rsplit("_Sweep_", 1)[1])
                for sweep_key in avg_values_dict.keys()
                if sweep_key.startswith(filename) # Exclure uniquement les sweeps supprimés
            ]
            sweeps.sort()

            # Initialiser des listes pour les moyennes globales par fichier
            all_avg_freqs = []
            all_cv_isis = []

            # Fixer la taille des sous-graphiques
            subplot_width = 5  # Largeur fixe de chaque sous-graphe
            subplot_height = 3  # Hauteur fixe de chaque sous-graphe

            # Calculer le nombre de sous-graphiques et organiser en 2 colonnes
            n_subplots = len(sweeps)
            nrows = (n_subplots + 1) // 2  # Nombre de lignes nécessaires pour 2 colonnes

            # Calculer la taille totale de la figure
            fig_width = subplot_width * 2  # 2 colonnes
            fig_height = subplot_height * nrows  # Taille adaptée au nombre de lignes

            # Créer la figure et les sous-graphiques
            fig, axes = plt.subplots(nrows=nrows, ncols=2, figsize=(fig_width, fig_height))

            # Si axes est un tableau NumPy, on le transforme en liste plate
            if isinstance(axes, np.ndarray):
                axes = axes.flatten().tolist()  
            else:
                axes = [axes]  # Si un seul subplot, on le met dans une liste

            fig.canvas.manager.set_window_title(f"Figure_{filename}")

            # Recalculer et tracer chaque sweep
            for i, sweep_number in enumerate(sweeps):
                if i < len(axes):
                    sweep_key = f"{filename}_Sweep_{sweep_number}"
                    if sweep_key in avg_values_dict and (filename, sweep_number) in signal_data:
                        time = signal_data[(filename, sweep_number)]["time"]
                        signal = signal_data[(filename, sweep_number)]["signal"]
                        threshold = avg_values_dict[sweep_key]["Threshold"]

                        # Recalculer les pics avec le nouveau threshold
                        peak_data = detect_peaks_and_calculate_frequency(time, signal, height_threshold=threshold)

                        # Calculer les ISI et fréquences
                        if len(peak_data["Time (s)"]) > 1:
                            isi = np.diff(peak_data["Time (s)"])
                            frequencies = 1 / isi
                            avg_freq = np.mean(frequencies)
                            isi_cv = np.std(isi) / np.mean(isi)
                        else:
                            avg_freq = 0
                            isi_cv = 0

                        # Mettre à jour avg_values_dict
                        avg_values_dict[sweep_key] = {
                            "Avg Freq Inst": avg_freq,
                            "Avg CV": isi_cv,
                            "Threshold": threshold
                        }

                        # Ajouter aux moyennes globales
                        all_avg_freqs.append(avg_freq)
                        all_cv_isis.append(isi_cv)

                        # Tracer le signal et les pics
                        axes[i].plot(time, signal, label="Filtered Signal", color=(98/255, 114/255, 164/255))
                        axes[i].scatter(
                            peak_data["Time (s)"], peak_data["Peak Values"],
                            color=(189/255, 147/255, 249/255), label="Spikes", zorder=5
                        )
                        axes[i].set_title(f"{filename} - Sweep {sweep_number} (Threshold={threshold})")
                        axes[i].set_xlabel("Time (s)")
                        axes[i].set_ylabel("Signal")
                        

            plt.tight_layout()
            plt.show(block=False)

            graph_filename = os.path.join(details_folder_path, f'{os.path.splitext(filename)[0]}.png')
            plt.savefig(graph_filename)
            print(f"Graphes mis à jour pour {filename} avec les thresholds actuels.")

            # Calculer les moyennes globales pour le fichier
            #global_avg_freq = np.mean(all_avg_freqs) if all_avg_freqs else 0
            #global_cv_isi = np.mean(all_cv_isis) if all_cv_isis else 0

            # Sauvegarder les moyennes dans un fichier Excel
            #output_excel_path = os.path.join(folder_path, f"{filename}_results.xlsx")
            #results_df = pd.DataFrame.from_dict(avg_values_dict, orient='index').reset_index()
            #results_df.columns = ["Sweep", "Avg Freq Inst", "Avg CV", "Threshold"]
            #results_df.to_excel(output_excel_path, index=False)

            #print(f"Résultats sauvegardés pour {filename} dans {output_excel_path}.")

    def create_continue_window_V2(self, avg_values_dict, signal_data, details_folder_path):
        """
        Crée une fenêtre interactive pour permettre à l'utilisateur de continuer ou de modifier les thresholds.

        Parameters:
            avg_values_dict (dict): Contient les informations sur les thresholds et autres données calculées.
            signal_data (dict): Contient les données des signaux pour chaque fichier et sweep.
            sweep_names (list): Liste des noms des sweeps au format "filename_Sweep_sweep_number".
        """
        dialog = QDialog(self)
        dialog.setWindowTitle("Cliquer pour continuer")
        dialog.resize(300, 200)
        dialog.setStyleSheet("""
            QDialog {                
            background-color: rgb(40, 44, 52);
            }
            QPushButton {
                background-color: rgb(98, 114, 164);
                border: none;
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-size: 12pt;
            }
            QPushButton:hover {
                background-color: rgb(189, 147, 249);
            }
        """)
        # Layout principal vertical
        layout = QVBoxLayout(dialog)

        # Label d'information
        label = QLabel("Cliquer pour continuer")
        layout.addWidget(label)

        # Bouton pour fermer les graphes et continuer
        
        def close_dialog():
            plt.close('all')  # Ferme tous les graphiques ouverts
            dialog.accept()   # Ferme la fenêtre

        button_close = QPushButton("Fermer")
        button_close.clicked.connect(close_dialog)
        layout.addWidget(button_close)

        # Bouton pour modifier les thresholds
        button_modify = QPushButton("Modifier Threshold")
        button_modify.clicked.connect(lambda: self.modify_threshold(avg_values_dict, signal_data, details_folder_path))
        layout.addWidget(button_modify)

        # Bouton pour supprimer les sweeps
        button_remove = QPushButton("Supprimer sweeps")
        button_remove.clicked.connect(lambda: self.select_sweep(avg_values_dict, signal_data, details_folder_path))
        layout.addWidget(button_remove)

        dialog.setLayout(layout)
        dialog.show()

        # Afficher la boîte de dialogue
        return dialog

    def close_plots():
        plt.close('all')
        
    def create_choice_window(self):
        # Boîte de dialogue modale
        dialog = QDialog(self)
        dialog.setWindowTitle("Select Current Type")
        dialog.resize(300, 150)
        dialog.setStyleSheet("""
            QDialog {
            background-color: rgb(40, 44, 52);
            }
            QPushButton {
                background-color: rgb(98, 114, 164);
                border: none;
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-size: 12pt;
            }
            QPushButton:hover {
                background-color: rgb(189, 147, 249);
            }                    
        """)

        # Layout principal vertical
        layout = QVBoxLayout(dialog)

        # Label d'instruction
        label = QLabel("Are the currents EPSC or IPSC?")
        label.setAlignment(Qt.AlignCenter)
        layout.addWidget(label)

        # Layout pour les boutons
        button_layout = QHBoxLayout()

        # Stocker le choix de l'utilisateur
        choice = {"value": None}

        # Bouton EPSC
        button_epsc = QPushButton("EPSC")
        button_epsc.clicked.connect(lambda: [self.set_choice(choice, "EPSC"), dialog.accept()])
        button_layout.addWidget(button_epsc)

        # Bouton IPSC
        button_ipsc = QPushButton("IPSC")
        button_ipsc.clicked.connect(lambda: [self.set_choice(choice, "IPSC"), dialog.accept()])
        button_layout.addWidget(button_ipsc)

        layout.addLayout(button_layout)

        # Afficher la boîte de dialogue
        dialog.exec()

        return choice["value"]

    def modify_threshold(self, avg_values_dict, signal_data, details_folder_path, parent=None):
        """
        Ouvre une fenêtre interactive pour modifier le threshold
        pour un ou plusieurs sweeps combinés (filename + sweep).
        """
        threshold_modified_sweeps = [] 
        # Création de la boîte de dialogue
        dialog = QDialog(parent)
        dialog.setWindowTitle("Modifier le Threshold")
        dialog.resize(500, 400)
        dialog.setStyleSheet("""
            QDialog {                
            background-color: rgb(40, 44, 52);
            }
            QPushButton {
                background-color: rgb(98, 114, 164);
                border: none;
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-size: 12pt;
            }
            QPushButton:hover {
                background-color: rgb(189, 147, 249);
            }
            QLineEdit {
                background-color: rgb(33, 37, 43);
                border: 2px solid rgb(44, 49, 58);
                border-radius: 5px;
                padding: 5px;
                color: rgb(255, 255, 255);
            }
            QLineEdit:hover {
                border: 2px solid rgb(98, 114, 164);
            }
            QListWidget {
                background-color: rgb(33, 37, 43);
                color: white;
            }   
            QListWidget::item:hover {
                background-color: rgb(189, 147, 249);
            }                  
        """)

        # Layout principal
        layout = QVBoxLayout(dialog)

        # Label d'instruction
        label_instruction = QLabel(
        "<p style='color: rgb(189, 147, 249); font-size: 12pt; font-weight: bold;'>"
        "Sélectionner un ou plusieurs sweeps")
        layout.addWidget(label_instruction)

        # Liste des sweeps avec sélection multiple
        sweep_list = QListWidget()
        sweep_list.setSelectionMode(QListWidget.MultiSelection)
        for sweep_key in avg_values_dict.keys():
            sweep_list.addItem(sweep_key)
        layout.addWidget(sweep_list)

        # Label pour l'entrée du threshold
        label_threshold = QLabel(
        "<p style='color: rgb(189, 147, 249); '>"
        "Saisir le nouveau threshold: ")
        layout.addWidget(label_threshold)

        # Champ d'entrée du threshold
        threshold_entry = QLineEdit()
        layout.addWidget(threshold_entry)

        # Fonction pour appliquer le threshold
        def apply_threshold():
            selected_items = [item.text() for item in sweep_list.selectedItems()]
            new_threshold = threshold_entry.text()

            if not selected_items:
                QMessageBox.critical(dialog, "Erreur", "Veuillez sélectionner au moins un sweep.")
                return

            try:
                new_threshold = float(new_threshold)
            except ValueError:
                QMessageBox.critical(dialog, "Erreur", "Veuillez entrer une valeur numérique pour le threshold.")
                return

            # Appliquer le nouveau threshold aux sweeps sélectionnés
            for sweep_key in selected_items:
                if sweep_key in avg_values_dict:
                    avg_values_dict[sweep_key]["Threshold"] = new_threshold
                    threshold_modified_sweeps.append(sweep_key)
                    print(f"Threshold modifié pour {sweep_key} : {new_threshold}")
                else:
                    print(f"Attention : {sweep_key} n'a pas été trouvé.")

            # Recalculer les pics et mettre à jour les graphes
            self.recalculate_peaks_and_update_graphs(avg_values_dict, signal_data, [], threshold_modified_sweeps, details_folder_path)
            dialog.accept()  # Fermer la boîte de dialogue

        # Bouton pour appliquer les changements
        button_apply = QPushButton("Appliquer")
        button_apply.clicked.connect(apply_threshold)
        layout.addWidget(button_apply)

        dialog.exec()  # Afficher la boîte de dialogue
        return threshold_modified_sweeps
    
    def select_sweep(self, avg_values_dict, signal_data, details_folder_path, parent=None):
        removed_sweeps = []

        dialog = QDialog(self)
        dialog.setWindowTitle("Sélectionner/Désélectionner les Sweeps")
        dialog.resize(500, 400)
        dialog.setStyleSheet("""
            QDialog {                
            background-color: rgb(40, 44, 52);
            }
            QPushButton {
                background-color: rgb(98, 114, 164);
                border: none;
                color: white;
                padding: 10px;
                border-radius: 5px;
                font-size: 12pt;
            }
            QPushButton:hover {
                background-color: rgb(189, 147, 249);
            }
            QLineEdit {
                background-color: rgb(33, 37, 43);
                border: 2px solid rgb(44, 49, 58);
                border-radius: 5px;
                padding: 5px;
                color: rgb(255, 255, 255);
            }
            QLineEdit:hover {
                border: 2px solid rgb(98, 114, 164);
            }
            QListWidget {
                background-color: rgb(33, 37, 43);
                color: white;
            }   
            QListWidget::item:hover {
                background-color: rgb(189, 147, 249);
            }                  
        """)


        layout = QVBoxLayout(dialog)

        label = QLabel("Select the sweeps to exclude from the analysis: ")
        layout.addWidget(label)

        list_widget = QListWidget()
        list_widget.setSelectionMode(QListWidget.MultiSelection)
        for sweep_key in avg_values_dict.keys():
            list_widget.addItem(sweep_key)
        layout.addWidget(list_widget)

        def apply_removesweep():
            selected_items = [item.text() for item in list_widget.selectedItems()]
            if not selected_items:
                QMessageBox.critical(dialog, "Erreur", "Veuillez sélectionner au moins un sweep.")
                return
            removed_sweeps.extend(selected_items)
            print("✅ Sweeps supprimés :", removed_sweeps)  # Vérification console

            # Supprimer ces sweeps de avg_values_dict pour ne plus les utiliser
            for sweep in removed_sweeps:
                if sweep in avg_values_dict:
                    del avg_values_dict[sweep]


            dialog.accept()  # Fermer la boîte de dialogue

        button_apply = QPushButton("Appliquer")
        button_apply.clicked.connect(apply_removesweep)
        layout.addWidget(button_apply)

        dialog.exec()
        
        # 🔄 Après suppression, recalculer les pics et mettre à jour les graphes
        print("🔄 Recalcul et mise à jour des graphes après suppression des sweeps...")
        self.recalculate_peaks_and_update_graphs(avg_values_dict, signal_data, removed_sweeps, None, details_folder_path)

    def set_choice(self, choice, value):
        """
        Définit la valeur choisie par l'utilisateur.

        Parameters:
            choice (dict): Dictionnaire mutable pour stocker le choix.
            value (str): Valeur à définir (EPSC ou IPSC).
        """
        choice["value"] = value
 
    def Stim_files(self, directory_path, progress_var, verification_input):
        
        total_files = get_total_abf_files_in_folders(directory_path, ["Stim"])
        current_value = 0
        choice = None

        for root, dirs, files in os.walk(directory_path):
            for folder_name in dirs:
                if folder_name == 'Stim':
                    if choice is None:  # Appeler `create_choice_window` une seule fois
                        choice = self.create_choice_window()
                        if not choice:
                            print("Operation cancelled by user.")
                            return
        
                    measure_positive_values = (choice == 'IPSC')

                    folder_path = os.path.join(root, folder_name)
                    # Liste pour stocker les résultats de chaque fichier
                    all_data_list = []
                    all_extreme_values = []
                    
                    # Parcourir tous les fichiers ABF dans le répertoire
                    for filename in os.listdir(folder_path):
                        
                        if filename.endswith(".abf"):
                            # Construire le chemin complet vers le fichier ABF
                            abf_file_path = os.path.join(folder_path, filename)

                            # Charger le fichier ABF
                            abf = pyabf.ABF(abf_file_path)
                            
                            
                            # Set sweep to the last one for channel 2
                            abf.setSweep(sweepNumber=abf.sweepCount - 1, channel=2)

                            # Calculate start and end voltages
                            start_voltage = np.argmax(abf.sweepY > 1)
                            end_voltage_all = np.nonzero(abf.sweepY > 1)[0]
                            
                            last_stim_indices = []
                            last_stim_times = []
                            
                            for i in range(1, len(end_voltage_all)):
                                if end_voltage_all[i] > end_voltage_all[i-1] + 1:
                                    last_stim_indices.append(end_voltage_all[i-1])
                            
                            # Handle the case where there is only one continuous stimulation event
                            if len(last_stim_indices) == 0 and len(end_voltage_all) > 0:
                                last_stim_indices.append(end_voltage_all[-1])  
                                    
                            # Find corresponding times for filtered indices
                            for idx in last_stim_indices:
                                if idx < len(abf.sweepX):
                                    last_stim_times.append(abf.sweepX[idx])
                                
                            # Filter to keep only the first index for each stimulation event
                            filtered_indices = [end_voltage_all[0]] if end_voltage_all.size > 0 else []
                            for i in range(1, len(end_voltage_all)):
                                if end_voltage_all[i] > end_voltage_all[i-1] + 1:
                                    filtered_indices.append(end_voltage_all[i])

                            
                            # Initialize list to store filtered times
                            filtered_times = []

                            # Find corresponding times for filtered indices
                            for idx in filtered_indices:
                                if idx < len(abf.sweepX):
                                    filtered_times.append(abf.sweepX[idx])

                            

                            # Adjust indices to include 100 points before and after
                            index_start_voltage = max(0, start_voltage - 100)
                            end_voltage = end_voltage_all[-1]
                            index_end_voltage = end_voltage + 600
                            
                            sweep_data = {"Time (s)": abf.sweepX[index_start_voltage:index_end_voltage]}
                            time_end_voltage = abf.sweepX[index_end_voltage]
                            
                            for sweep_number in range(abf.sweepCount):
                                baseline_start = abf.sweepX[0]
                                baseline_end = abf.sweepX[start_voltage - 100]

                                abf.setSweep(sweepNumber=sweep_number, baseline=[baseline_start, baseline_end], channel=0)
                                column_name_y = f"Sweep_{sweep_number}_data"
                                sweep_data[column_name_y] = abf.sweepY[index_start_voltage:index_end_voltage]

                            df = pd.DataFrame(sweep_data)
                            # Compute the average trace across all sweeps
                            df["Average_Trace"] = df.drop(columns=["Time (s)"]).mean(axis=1)
                            
                            # Define constants
                            time_after_stim = 0.0015  # Time in seconds after stim time to consider for finding most negative value

                            
                            mean_values_dict = {"Filename": os.path.basename(abf_file_path)}
                            file_extreme_values = []

                            for idx, time_stim in enumerate(filtered_times):
                                # Find the index in df where time is time_stim + time_after_stim
                                idx_after_stim = df[df["Time (s)"] >= (time_stim + time_after_stim)].index.min()

                                if idx_after_stim is None:
                                    continue

                                # Store the index as "indices_stim" in your result
                                values_dict = {"Filename": os.path.basename(abf_file_path), "stim": idx + 1}
                                
                                range_start = max(0, idx_after_stim - 10)
                                range_end = min(len(df), idx_after_stim + 150)

                                if measure_positive_values:
                                    extreme_value = df["Average_Trace"][range_start:range_end].max()
                                    extreme_time = df["Time (s)"][df["Average_Trace"][range_start:range_end].idxmax()]
                                    
                                else:
                                    extreme_value = df["Average_Trace"][range_start:range_end].min()
                                    extreme_time = df["Time (s)"][df["Average_Trace"][range_start:range_end].idxmin()]

                                
                                # Store the single extreme value
                                mean_values_dict[f"Valeur moyenne stim {idx + 1}"] = extreme_value
                                values_dict["Average_Trace_extreme_value"] = extreme_value
                                file_extreme_values.append((extreme_value, extreme_time))

                            first_peak_amplitude = file_extreme_values[0][0] if file_extreme_values else None
                            first_peak_time = file_extreme_values[0][1] if file_extreme_values else None
                            # Now you can calculate Firstpeak_10, Firstpeak_90, and rise time as needed
                            if first_peak_amplitude is not None:
                                Firstpeak_10 = first_peak_amplitude * 0.10
                                Firstpeak_90 = first_peak_amplitude * 0.90                                
                                # Find the indices for Firstpeak_10 and Firstpeak_90 before the peak
                                firstpeak_10_before_index = None
                                firstpeak_90_before_index = None
                                firstpeak_10_after_index = None
                                firstpeak_90_after_index = None
                                
                                first_stim_time = last_stim_times[0]
                                if len(filtered_times) > 1:
                                    end_firstpeak = filtered_times[1] - 0.010
                                else: 
                                    end_firstpeak = last_stim_times[0] + 0.050
                                range_start_time = first_stim_time
                                range_mid_time = first_peak_time
                                range_end_time = end_firstpeak                              
                                
                            
                                # Filter the DataFrame to include only the rows within this time range
                                filtered_range = df[(df["Time (s)"] >= range_start_time) & (df["Time (s)"] <= range_mid_time)]
                                filtered_range2 = df[(df["Time (s)"] >= range_mid_time) & (df["Time (s)"] <= range_end_time)]
                                # Now you can perform further analysis on `filtered_range`
                                # For example, finding the closest value to Firstpeak_10 in this range
                                firstpeak_10_before_index = (filtered_range["Average_Trace"] - Firstpeak_10).abs().idxmin()

                                # Similarly, you can find other indices or perform other calculations within this range
                                firstpeak_90_before_index = (filtered_range["Average_Trace"] - Firstpeak_90).abs().idxmin()
                                firstpeak_10_after_index = (filtered_range2["Average_Trace"] - Firstpeak_10).abs().idxmin()
                                firstpeak_90_after_index = (filtered_range2["Average_Trace"] - Firstpeak_90).abs().idxmin()
                                
                                # Example: Calculate rise time using these indices
                                if firstpeak_10_before_index is not None and firstpeak_90_before_index is not None:
                                    rise_time = df["Time (s)"].iloc[firstpeak_90_before_index] - df["Time (s)"].iloc[firstpeak_10_before_index]
                                else:
                                    rise_time = None


                                if firstpeak_10_after_index is not None and firstpeak_90_after_index is not None:
                                    decay_time =  df["Time (s)"].iloc[firstpeak_10_after_index] - df["Time (s)"].iloc[firstpeak_90_after_index]
                                else:
                                    decay_time = None

                                # Add the results to values_dict or mean_values_dict
                                mean_values_dict["Rise Time (s)"] = rise_time
                                mean_values_dict["Decay Time (s)"] = decay_time


                            all_data_list.append(mean_values_dict)
                            all_extreme_values.append(file_extreme_values) 
                            
                            # Create a DataFrame from mean_values_dict
                            result_df = pd.DataFrame(all_data_list)
                            
                            graphes_folder = os.path.join(folder_path, "Graphes")
                            os.makedirs(graphes_folder, exist_ok=True)
                            
                            fig, ax = plt.subplots(figsize=(8, 6))
                            # Plot the Average_Trace column
                            ax.plot(df["Time (s)"], df["Average_Trace"], label='Average Trace', color=(98/255, 114/255, 164/255))
                            # Plot the extreme values on the Average_Trace
                            for i, (ext_value, ext_time) in enumerate(file_extreme_values):
                                ax.scatter(ext_time, ext_value, color=(189/255, 147/255, 249/255), label=f'Extreme Value stim {i+1}' if i == 0 else "")

                            ax.set_xlabel('Time (s)')
                            ax.set_ylabel('Filtered Data')
                            ax.set_title(f'{filename} - with Currents Max amplitude')
                            ax.legend()
                            
                            # Save the plot as a PNG file in the Graphes folder
                            graph_filename = os.path.join(graphes_folder, f'{os.path.splitext(filename)[0]}.png')
                            plt.savefig(graph_filename)
                            
                            if verification_input == 'Y':
                                plt.show(block=False)

                            # Display or further process result_df as needed
                            print("Analysis Results:")
                            print(result_df)
                            
                            # Save the DataFrame to an Excel file with the ABF file name followed by "_values"
                            folder_name = os.path.basename(directory_path) 
                            excel_output_path = os.path.join(folder_path, f"stim_values_{folder_name}.xlsx")
                            
                            result_df.to_excel(excel_output_path, index=False)
                                                    
                            # Create ExcelWriter and save the DataFrame
                            with pd.ExcelWriter(excel_output_path, engine='openpyxl') as writer:
                                result_df.to_excel(writer, index=False)

                                # Access the default sheet
                                sheet = writer.sheets['Sheet1']

                                # Adjust column widths based on content in the first row
                                for column in sheet.columns:
                                    max_length = 0
                                    column = [cell for cell in column]
                                    for cell in column:
                                        try:  # Necessary to avoid error on empty cells
                                            if len(str(cell.value)) > max_length:
                                                max_length = len(cell.value)
                                        except:
                                            pass
                                    adjusted_width = (max_length + 2)
                                    sheet.column_dimensions[column[0].column_letter].width = adjusted_width
                                    
                            current_value += 1
                            progress_percentage = (current_value / total_files) * 100
                            self.progress_bar.setValue(progress_percentage)
                            QApplication.processEvents()

                            
                    if verification_input == 'Y':                            
                        # Créer une fenêtre contextuelle pour attendre l'action de l'utilisateur
                        continue_window = self.create_continue_window()

                        # Attendre que l'utilisateur clique sur la fenêtre contextuelle
                        continue_window.exec()
        
                            
    def AMPA_NMDA_files(self, directory_path, progress_var, verification_input):
            
        total_files = get_total_abf_files_in_folders(directory_path, ["AMPA_NMDA"])
        current_value = 0
        for root, dirs, files in os.walk(directory_path):
            for folder_name in dirs:
                if folder_name == 'AMPA_NMDA':
                    folder_path = os.path.join(root, folder_name)
                    # List to store the results of each file
                    all_data_list = []
                    all_extreme_values = []
                    
                    for filename in os.listdir(folder_path):
                        if filename.endswith(".abf"):
                            abf_file_path = os.path.join(folder_path, filename)
                            abf = pyabf.ABF(abf_file_path)
                            
                            # Set sweep to the last one for channel 2
                            abf.setSweep(sweepNumber=abf.sweepCount - 1, channel=2)
                            start_voltage = np.argmax(abf.sweepY > 1)
                            end_voltage_all = np.nonzero(abf.sweepY > 1)[0]
                            
                            filtered_indices = [end_voltage_all[0]] if end_voltage_all.size > 0 else []
                            for i in range(1, len(end_voltage_all)):
                                if end_voltage_all[i] > end_voltage_all[i-1] + 1:
                                    filtered_indices.append(end_voltage_all[i])

                            filtered_times = []
                            for idx in filtered_indices:
                                if idx < len(abf.sweepX):
                                    filtered_times.append(abf.sweepX[idx])

                            index_start_voltage = max(0, start_voltage - 100)
                            end_voltage = end_voltage_all[-1]
                            index_end_voltage = end_voltage + 2000

                            sweep_data = {"Time (s)": abf.sweepX[index_start_voltage:index_end_voltage]}
                            for sweep_number in range(abf.sweepCount):
                                baseline_start_idx = max(0, start_voltage - 200)
                                baseline_end_idx = baseline_start_idx + 200

                                baseline_start = abf.sweepX[baseline_start_idx]
                                baseline_end = abf.sweepX[baseline_end_idx]

                                abf.setSweep(sweepNumber=sweep_number, baseline=[baseline_start, baseline_end], channel=0)
                                column_name_y = f"Sweep_{sweep_number}_data"
                                sweep_data[column_name_y] = abf.sweepY[index_start_voltage:index_end_voltage]

                            df = pd.DataFrame(sweep_data)
                            print(df)
                            selected_columns = df.columns[(df.columns.get_loc('Time (s)') + 1):]
                            time_after_stim = 0.002  # Time in seconds after stim time to consider for finding most negative value
                            time_after_stim50ms = 0.050
                            mean_values_dict = {"Filename": os.path.basename(abf_file_path)}
                            file_extreme_values = []

                            for idx, time_stim in enumerate(filtered_times):
                                idx_stim = df[df["Time (s)"] == time_stim].index[0]
                                idx_after_stim = df[df["Time (s)"] >= (time_stim + time_after_stim)].index.min()
                                if idx_after_stim is None:
                                    continue
                                idx_after_stim_50ms = df[df["Time (s)"] >= (time_stim + time_after_stim50ms)].index.min()
                                print(idx_after_stim_50ms)
                                if idx_after_stim_50ms is None:
                                    continue
                                
                                values_dict = {"Filename": os.path.basename(abf_file_path), "stim": idx + 1}
                                extreme_values = []
                                extreme_times = []
                                
                                for sweep_number, col in enumerate(selected_columns):
                                    if sweep_number == 0:  # First sweep: most negative value
                                        range_start = max(0, idx_after_stim - 10)
                                        range_end = min(len(df), idx_after_stim + 150)
                                        extreme_value = df[col][range_start:range_end].min()
                                        extreme_time = df["Time (s)"][df[col][range_start:range_end].idxmin()]
                                        values_dict[f"Valeur AMPA stim {idx + 1}"] = extreme_value
                                    elif sweep_number == 1:  # Second sweep: most positive value 50ms after stim                                    
                                        extreme_value = df.at[idx_after_stim_50ms, col]
                                        extreme_time = df.at[idx_after_stim_50ms, "Time (s)"]
                                        values_dict[f"Valeur NMDA stim {idx + 1}"] = extreme_value

                                    extreme_values.append(extreme_value)
                                    extreme_times.append(extreme_time)
                                    
                            
                            file_extreme_values.append((extreme_values, extreme_times))
                            all_data_list.append(values_dict)
                            all_extreme_values.append(file_extreme_values)
                            
                            result_df = pd.DataFrame(all_data_list)
                            
                            if verification_input == 'Y':
                                fig, ax = plt.subplots(figsize=(8, 6))
                                for column in selected_columns:
                                    ax.plot(df["Time (s)"], df[column], label=f'{column}')

                                for i, (ext_values, ext_times) in enumerate(file_extreme_values):
                                    ampa_value = ext_values[0]  # AMPA value
                                    ampa_time = ext_times[0]    # Time corresponding to AMPA value
                                    nmda_value = ext_values[1]  # NMDA value
                                    nmda_time = filtered_times[i] + 0.050  # Time for NMDA value is 50 ms after stim

                                    # Tracer la valeur AMPA
                                    ax.scatter(ampa_time, ampa_value, color='red', label=f'AMPA Value stim {i+1}' if i == 0 else "")
                                    

                                    # Tracer la valeur NMDA
                                    ax.scatter(nmda_time, nmda_value, color='green', label=f'NMDA Value stim {i+1}' if i == 0 else "")
                                    

                                    # Tracer la ligne de temps de stimulation
                                    ax.axvline(x=filtered_times[i], color='blue', linestyle='--', label='Time stim' if i == 0 else "")

                                ax.set_xlabel('Time (s)')
                                ax.set_ylabel('Filtered Data')
                                ax.set_title(f'{filename} - Filtered Data with AMPA and NMDA Values')
                                ax.legend()
                                plt.show(block=False)

                            print("Analysis Results:")
                            print(result_df)
                            
                            excel_output_path = os.path.join(folder_path, "stim_values.xlsx")
                            
                            with pd.ExcelWriter(excel_output_path, engine='openpyxl') as writer:
                                result_df.to_excel(writer, index=False)
                                sheet = writer.sheets['Sheet1']
                                for column in sheet.columns:
                                    max_length = 0
                                    column = [cell for cell in column]
                                    for cell in column:
                                        try:
                                            if len(str(cell.value)) > max_length:
                                                max_length = len(cell.value)
                                        except:
                                            pass
                                    adjusted_width = (max_length + 2)
                                    sheet.column_dimensions[column[0].column_letter].width = adjusted_width
                                    
                            current_value += 1
                            progress_percentage = (current_value / total_files) * 100
                            self.progress_bar.setValue(progress_percentage)
                            QApplication.processEvents()
                            
                    if verification_input == 'Y':                            
                        continue_window = self.create_continue_window()
                        continue_window.exec()
                       
                                                                    
    def Rheo_files(self, directory_path, progress_var, verification_input):
        total_files = get_total_abf_files_in_folders(directory_path, ["Rheobase"]) 
        if not total_files:
            total_files = get_total_abf_files_in_folders(directory_path, ["Rheo"])
        current_value = 0
        for root, dirs, files in os.walk(directory_path):
            for folder_name in dirs:
                if folder_name == 'Rheobase' or folder_name == 'Rheo':
                    folder_path = os.path.join(root, folder_name)
            
                    paramsDict = {}  
                    paramsList = []      
                    # Parcourir tous les fichiers ABF dans le répertoire
                    for filename in os.listdir(folder_path):
                        
                        if filename.endswith(".abf"):
                            # Construire le chemin complet vers le fichier ABF
                            abf_file_path = os.path.join(folder_path, filename)

                            # Charger le fichier ABF
                            abf = pyabf.ABF(abf_file_path)
                            channel_to_analyze = None
                            for channel in [0, 1]: 
                                # Trouver l'indice du dernier sweep
                                abf.setSweep(sweepNumber=abf.sweepCount - 1  , channel=channel)                                
                                end_courant_all = np.nonzero(abf.sweepC > 0)[0]

                                if end_courant_all.size > 0:
                                    print(f"🔹 Données trouvées sur channel {channel}")
                                    channel_to_analyze = channel
                                    break  # Sortir de la boucle si un channel contient des données
                           
                            if channel_to_analyze is None:
                                print(f"❌ Aucune donnée détectée sur {filename}, skipping...")
                                continue  

                            abf.setSweep(sweepNumber=abf.sweepCount - 1, channel=channel_to_analyze)
                            sweep_data = {"Time (s)": abf.sweepX}  
                            start_courant = np.argmax(abf.sweepC > 0)
                            sweepX_start_courant = abf.sweepX[start_courant]                             
                            index_start_courant = np.where(abf.sweepX == sweepX_start_courant)[0][0]
                            
                            end_courant = end_courant_all[-1]                            
                            sweepX_end_courant = abf.sweepX[end_courant]

                            for sweep_number in range(abf.sweepCount):
                                abf.setSweep(sweepNumber=sweep_number, channel=channel_to_analyze)
                                sweepC_at_start_column = abf.sweepC[index_start_courant]
                                column_name_y = f"{sweepC_at_start_column:.2f}pA"
                                sweep_data[column_name_y] = abf.sweepY

                            df = pd.DataFrame(sweep_data)                       
                            print(df)        
                                       
                            # Supprimer les deux premières lignes de chaque colonne
                            df = df.iloc[0:, :]

                            # Remove rows with NaN values
                            df = df.dropna()
                            
                            # Remplacer les ',' par '.' dans toutes les colonnes
                            df = df.replace(',', '.', regex=True)

                            # Convertir toutes les valeurs en numériques
                            df = df.apply(pd.to_numeric, errors='coerce')

                            # Sélectionner les données temporelles entre 0 et sweepX_start_courant
                            filtered_data = df[(df['Time (s)'] >= 0) & (df['Time (s)'] <= sweepX_start_courant)]

                            # Calculer la médiane de tous les points du DataFrame (ignorant la colonne "Time (s)")
                            median_values = filtered_data.drop("Time (s)", axis=1).median()
                            # Calculer la moyenne des médianes
                            mean_of_median_values = median_values.mean()
                            # Ajouter la valeur moyenne au dictionnaire
                            # Ajouter la valeur moyenne à la liste
                            paramsDict[filename] = {"Vrest (mV)": mean_of_median_values}
                            print(paramsList)

                            filtered_df = df[(df['Time (s)'] >= sweepX_start_courant) & (df['Time (s)'] <= sweepX_end_courant)]
                            df_graphe = df.copy()
                            print (df_graphe)
                            # Identifier les colonnes (à partir de la colonne 2) où la valeur est supérieure à -20
                            columns_to_check = filtered_df.columns[1:][filtered_df.iloc[:, 1:].gt(-20).any()]
                            paramsRheo = {}
                            # Trouver la première colonne qui satisfait la condition
                            for col in columns_to_check:
                                if any(filtered_df[col] > -20):
                                    # Sauvegarder la première colonne qui remplit la condition dans le dictionnaire
                                    #numeric_part = re.search(r'\d+', col).group()
                                    #paramsDict[filename]["Rhéobase (pA)"] = int(numeric_part)  # Convertir en entier
                                    paramsDict[filename]["Rhéobase"] = col
                                    if verification_input == 'Y':
                                        # Créer une figure et un axe
                                        fig, ax = plt.subplots()                            
                                        # Tracer les données de col_to_plot en rouge
                                        col_index = df_graphe.columns.get_loc(col)
                                        # Calculer la différence entre sweepX_end_courant et sweepX_start_courant
                                        difference = sweepX_end_courant - sweepX_start_courant

                                        # Créer un masque pour col_to_plot dans df_graphe basé sur la différence calculée
                                        mask_col_to_plot = (df_graphe["Time (s)"] >= (sweepX_start_courant - 2 * difference)) & (df_graphe["Time (s)"] <= (sweepX_end_courant + 2 * difference))

                                        #mask_col_to_plot = (df_graphe["Time (s)"] >= sweepX_start_courant - 0.2) & (df_graphe["Time (s)"] <= sweepX_end_courant + 0.2)
                                        selected_time_col_to_plot = df_graphe["Time (s)"][mask_col_to_plot]
                                        selected_data_col_to_plot = df_graphe[col][mask_col_to_plot]
                                        ax.plot(selected_time_col_to_plot, selected_data_col_to_plot, label=f'{col}', color=(98/255, 114/255, 164/255))
                                        # Ajouter une ligne horizontale pour le potentiel de repos
                                        ax.axhline(y=paramsDict[filename]["Vrest (mV)"], color=(189/255, 147/255, 249/255), linestyle='--', label='Resting potential')
                                        # Vérifier si la colonne précédente existe dans df_graphe
                                        if col_index > 0:
                                            previous_col = df_graphe.columns[col_index - 1]
                                            if previous_col in df_graphe.columns:
                                                mask_prev_col = (df_graphe["Time (s)"] >= (sweepX_start_courant - 2 * difference)) & (df_graphe["Time (s)"] <= (sweepX_end_courant + 2 * difference))
                                                selected_time_prev_col = df_graphe["Time (s)"][mask_prev_col]
                                                selected_data_prev_col = df_graphe[previous_col][mask_prev_col]
                                                ax.plot(selected_time_prev_col, selected_data_prev_col, label=f'{previous_col}', color='gray')
                                        
                                        # Ajouter une légende
                                        ax.legend()

                                        # Ajouter des titres et des étiquettes d'axe
                                        ax.set_title(f"{filename}: La Rhéobase est de {col}")
                                        ax.set_xlabel('Time (s)')
                                        ax.set_ylabel('Current (pA)')

                                        # Afficher le graphe
                                        plt.show(block=False)
                                        
                                    # Récupérer les valeurs de la colonne qui satisfont la condition
                                    relevant_values = filtered_df.loc[filtered_df[col] > -20, col].tolist()
                                    relevant_times = filtered_df.loc[filtered_df[col] > -20, 'Time (s)'].tolist()
            
                                    # Parcourir les valeurs pour trouver la première valeur > -20 suivie d'une valeur plus petite
                                    for i in range(len(relevant_values) - 1):
                                        if relevant_values[i] > -20 and relevant_values[i + 1] < relevant_values[i]:
                                            # Calculer la différence de temps entre la valeur maximale et sweepX_start_courant
                                            max_peak_value_diff = (relevant_times[i] - sweepX_start_courant) * 1000
                                            
                                            # Ajouter les données de paramètres à paramsList
                                            paramsDict[filename]["Spike latency (ms)"] = max_peak_value_diff
                        
                                            break  # Sortir de la boucle interne une fois que la valeur est trouvée
                                        
                                    break
                                    
                            
                            # Afficher le dictionnaire avec les valeurs des paramètres
                            print("Paramètres sauvegardés:", paramsDict)
                            
                            # Convert the paramsDict dictionary to a DataFrame
                            df_Rheo = pd.DataFrame.from_dict(paramsDict, orient='index').reset_index()
                            df_Rheo.columns = ['File', 'Vrest (mV)', 'Rhéobase (pA)', 'Spike latency (ms)']                     
                            
                            # Save the DataFrame to an Excel file with the ABF file name followed by "_values"
                            excel_output_path = os.path.join(folder_path, "rheo_values.xlsx")
                            df_Rheo.to_excel(excel_output_path, index=False)
                            print (paramsDict)
                            
                            # Create ExcelWriter and save the DataFrame
                            with pd.ExcelWriter(excel_output_path, engine='openpyxl') as writer:
                                df_Rheo.to_excel(writer, index=False)

                                # Access the default sheet
                                sheet = writer.sheets['Sheet1']

                                # Adjust column widths based on content in the first row
                                for column in sheet.columns:
                                    max_length = 0
                                    column = [cell for cell in column]
                                    for cell in column:
                                        try:  # Necessary to avoid error on empty cells
                                            if len(str(cell.value)) > max_length:
                                                max_length = len(cell.value)
                                        except:
                                            pass
                                    adjusted_width = (max_length + 2)
                                    sheet.column_dimensions[column[0].column_letter].width = adjusted_width
                                    
                            current_value += 1
                            progress_percentage = (current_value / total_files) * 100
                            self.progress_bar.setValue(progress_percentage)
                            QApplication.processEvents()
                            
                            
                            
                    if verification_input == 'Y':                            
                        # Créer une fenêtre contextuelle pour attendre l'action de l'utilisateur
                        continue_window = self.create_continue_window()

                        # Attendre que l'utilisateur clique sur la fenêtre contextuelle
                        continue_window.exec()

                            
    def Cellattached_files(self, directory_path, progress_var, verification_input):
        
        total_files = get_total_abf_files_in_folders(directory_path, ["Cell attached"])
        current_value = 0

        avg_values_dict = {} 
        signal_data = {}
        abf_files = []
        folder_path = False

        for root, dirs, files in os.walk(directory_path):
            for folder_name in dirs:
                if folder_name == 'Cell attached':

                    print("📂 Début de l'analyse Cell attached...")  # 🔥 Vérifie si cette fonction est exécutée

                    folder_found = True
                    folder_path = os.path.join(root, folder_name)
                    
                    paramNames = ["Freq. inst (Hz)", "CV "]
                    abf_files = []  
                    sweep_names = []  
                    figures_dict = {}
                    for filename in os.listdir(folder_path):
                        
                        # Vérifier si c'est un dossier et si le nom du dossier est 'CCIV'
                        if filename.endswith(".abf"):      
                            print(f"🔄 Traitement du fichier : {filename}")  # 🔍 DEBUG
                            details_folder_path = os.path.join(folder_path, 'Graphes')
                            os.makedirs(details_folder_path, exist_ok=True)                         
                            # Construire le chemin complet vers le fichier ABF
                            abf_file_path = os.path.join(folder_path, filename)
                            abf_files.append(filename)
                            # Charger le fichier ABF
                            abf = pyabf.ABF(abf_file_path)      
                            # Initialiser des listes pour collecter les données de tous les sweeps
                            all_avg_freqs = []
                            all_cv_isis = []
                        
                            # Initialiser un dictionnaire pour stocker les données de chaque sweep
                            sweep_data = {"Time (s)": abf.sweepX}
                            all_max_peak_values_dict = {}
                            
                            channel_to_analyze = None
                            for channel in [0, 1]: 
                                abf.setSweep(sweepNumber=abf.sweepCount - 1, channel=channel) 
                                signal = abf.sweepY

                                if np.any(np.abs(signal) > 1):
                                    print(f"🔹 Données trouvées sur channel {channel}")
                                    channel_to_analyze = channel
                                    break

                            if channel_to_analyze is None:
                                print(f"❌ Aucune donnée détectée sur {filename}, skipping...")
                                continue  

                            abf.setSweep(sweepNumber=abf.sweepCount - 1, channel=channel_to_analyze)
                            # Initialiser une liste pour stocker les valeurs médianes des deux derniers sweeps
                            for sweep_number in range(abf.sweepCount):
                                                                    
                                baseline_start = abf.sweepX[0]
                                baseline_end = baseline_start + 0.200

                                abf.setSweep(sweepNumber=sweep_number, baseline=[baseline_start, baseline_end])
                                time = abf.sweepX
                                signal = abf.sweepY

                                # Appliquer le filtre au signal
                                filtered_signal = filter_signal(signal)

                                # Stocker les données dans signal_data
                                signal_data[(filename, sweep_number)] = {
                                    "time": time,
                                    "signal": filtered_signal
                                }

                                peak_data = detect_peaks_and_calculate_frequency(time, filtered_signal, height_threshold=-30)
                                column_name_y = f"Sweep_{sweep_number}_data"
                                all_max_peak_values_dict[column_name_y] = {
                                    "Time (s)": peak_data["Time (s)"],
                                    "Max Peak Value": peak_data["Peak Values"]
                                }
                                
                                # Calculer les ISI et fréquences avec les pics restants
                                if len(peak_data["Time (s)"]) > 1:
                                    isi = np.diff(peak_data["Time (s)"])
                                    frequencies = 1 / isi
                                    avg_freq = np.mean(frequencies)
                                    isi_cv = np.std(isi) / np.mean(isi)
                                else:
                                    avg_freq = 0
                                    isi_cv = 0

                                # Ajouter les moyennes du sweep aux listes globales
                                all_avg_freqs.append(avg_freq)
                                all_cv_isis.append(isi_cv)

                                # Ajouter les données filtrées au DataFrame
                                sweep_data[column_name_y] = filtered_signal

                                # Create a unique sweep name and add to sweep_names
                                sweep_name = f"{filename}_Sweep_{sweep_number}"
                                sweep_names.append(sweep_name)

                                # Add sweep data to avg_values_dict
                                avg_values_dict[sweep_name] = {
                                    "Avg Freq Inst": avg_freq,
                                    "Avg CV": isi_cv,
                                    "Threshold": -30  # Default threshold
                                }
                            print(avg_values_dict)

                            # Créer un DataFrame final pour les données de tous les sweeps
                            df = pd.DataFrame(sweep_data)
                            print(df)            

                            # Fixer la taille des sous-graphiques
                            subplot_width = 5  # Largeur fixe de chaque sous-graphe
                            subplot_height = 3  # Hauteur fixe de chaque sous-graphe

                            # Calculer le nombre de lignes nécessaires pour une grille de 2 colonnes
                            n_subplots = abf.sweepCount
                            nrows = (n_subplots + 1) // 2  # Nombre de lignes nécessaires (division entière)

                            # Calculer la taille totale de la figure
                            fig_width = subplot_width * 2  # 2 colonnes
                            fig_height = subplot_height * nrows  # Taille adaptée au nombre de lignes

                            # Créer la figure et les sous-graphiques
                            fig, axes = plt.subplots(nrows=nrows, ncols=2, figsize=(fig_width, fig_height))
                            
                            if isinstance(axes, np.ndarray):
                                axes = axes.flatten().tolist()  # Aplatir et convertir en liste
                            else:
                                axes = [axes]  # Si un seul subplot, le transformer en liste
                            fig.canvas.manager.set_window_title(f"Figure_{filename}")    
                            figures_dict[filename] = fig

                            for i, (column_name, max_peak_values) in enumerate(all_max_peak_values_dict.items()):

                                if i < len(axes):  # Vérifier que i ne dépasse pas la taille de 'axes'
                                    
                                    axes[i].plot(df["Time (s)"], df[column_name], label=f'{column_name} Filtered', color=(98/255, 114/255, 164/255))
                                    
                                    # Tracer les points de "Max Peak Value" pour chaque colonne
                                    axes[i].scatter(max_peak_values["Time (s)"], max_peak_values["Max Peak Value"], color=(189/255, 147/255, 249/255), zorder=5)


                                    # Ajouter des labels et un titre pour chaque sous-graphique
                                    axes[i].set_xlabel('Time (s)')
                                    axes[i].set_ylabel('Filtered Data')
                                    axes[i].set_title(f'{filename} - {column_name}')
                                    #axes[i].legend()


                            # Désactiver les axes inutilisés si le nombre de graphes est impair
                            for j in range(len(all_max_peak_values_dict), len(axes)):
                                axes[j].axis('off')

                            fig.subplots_adjust(top=0.95, bottom=0.05, left=0.1, right=0.9, hspace=0.6, wspace=0.4)

                            graph_filename = os.path.join(details_folder_path, f'{os.path.splitext(filename)[0]}.png')
                            plt.savefig(graph_filename)
                            
                            # Option interactive si vérification demandée
                            if verification_input == 'Y':
                                plt.tight_layout()
                                plt.show(block=False)

                    if folder_found and avg_values_dict:            
                        # Créer une fenêtre contextuelle pour attendre l'action de l'utilisateur
                        continue_window_V2 = self.create_continue_window_V2(avg_values_dict, signal_data, details_folder_path)

                        # Attendre que l'utilisateur clique sur la fenêtre contextuelle
                        continue_window_V2.exec()


                        print (avg_values_dict)
                    
                    if folder_found:
                        global_avg_list = []

                        for filename in abf_files:
                            # Filtrer les sweeps correspondant à ce fichier
                            sweeps = [
                                sweep_key for sweep_key in avg_values_dict.keys()
                                if sweep_key.startswith(filename)
                            ]

                            # Moyennes globales pour ce fichier
                            all_avg_freqs = [avg_values_dict[sweep]["Avg Freq Inst"] for sweep in sweeps]
                            all_cv_isis = [avg_values_dict[sweep]["Avg CV"] for sweep in sweeps]

                            global_avg_freq = np.mean(all_avg_freqs) if all_avg_freqs else 0
                            global_cv_isi = np.mean(all_cv_isis) if all_cv_isis else 0

                            # Ajouter à une liste pour la sauvegarde sous format Excel
                            global_avg_list.append({
                                "Filename": filename,
                                "Global Avg Freq Inst": global_avg_freq,
                                "Global Avg CV": global_cv_isi
                            })

                        # Sauvegarder les moyennes globales dans un fichier Excel
                        df_global_avg = pd.DataFrame(global_avg_list)
                        output_excel_path = os.path.join(folder_path, "Global_Averages.xlsx")
                        df_global_avg.to_excel(output_excel_path, index=False)

                        print(f"Les moyennes globales ont été sauvegardées dans {output_excel_path}")


                        current_value += 1
                        progress_percentage = (current_value / total_files) * 100
                        self.progress_bar.setValue(progress_percentage)
                        QApplication.processEvents()

    def CCIV_files(self, directory_path, progress_var, verification_input):
        total_files = get_total_abf_files_in_folders(directory_path, ["CCIV"])
        current_value = 0

        for root, dirs, files in os.walk(directory_path):
            for folder_name in dirs:
                if folder_name == 'CCIV':
                    folder_path = os.path.join(root, folder_name)
                    
                    openpyxl.Workbook().save(os.path.join(folder_path, "combined_CCIV.xlsx"))
                    combined_excel_file_path = os.path.join(folder_path, "combined_CCIV.xlsx")
                    for filename in os.listdir(folder_path):
                        if filename.endswith(".abf"):
                            # Construire le chemin complet vers le fichier ABF
                            abf_file_path = os.path.join(folder_path, filename)
                                                    
                            details_folder_path = os.path.join(folder_path, 'Graphes')
                            os.makedirs(details_folder_path, exist_ok=True)
                            abf = pyabf.ABF(abf_file_path)

                            channel_to_analyze = None
                            for channel in [0, 1]: 
                                abf.setSweep(sweepNumber=abf.sweepCount - 1, channel=channel)
                                end_courant_all = np.nonzero(abf.sweepC > 0)[0]
                                
                                if end_courant_all.size > 0:  
                                    print(f"🔹 Données trouvées sur channel {channel}")
                                    channel_to_analyze = channel
                                    break
                            
                            if channel_to_analyze is None:
                                print(f"❌ Aucune donnée détectée sur {filename}, skipping...")
                                continue  
    

                            abf.setSweep(sweepNumber=abf.sweepCount - 1, channel=channel_to_analyze)                
                            sweep_data = {"Time (s)": abf.sweepX}
                            start_courant = np.argmax(abf.sweepC > 0)
                            sweepX_start_courant = abf.sweepX[start_courant]
                            index_start_courant = np.where(abf.sweepX == sweepX_start_courant)[0][0]                                                                           
                    
                            end_courant = end_courant_all[-1]
                            sweepX_end_courant = abf.sweepX[end_courant]

                            for sweep_number in range(abf.sweepCount):  
                                abf.setSweep(sweepNumber=sweep_number, channel=channel_to_analyze)                              
                                sweepC_at_start_column = abf.sweepC[index_start_courant]
                                column_name_y = f"{sweepC_at_start_column:.2f}pA"
                                sweep_data[column_name_y] = abf.sweepY

                            df = pd.DataFrame(sweep_data)
                                                                                
                            print (df)
                            data_list = []

                            df = df.iloc[1:, :]
                            df = df.dropna()
                            df = df.replace(',', '.', regex=True)
                            df = df.apply(pd.to_numeric, errors='coerce')

                            selected_columns = df.columns[(df.columns.get_loc('Time (s)') + 1):(df.columns.get_loc('0.00pA') + 1)]

                            print(selected_columns)
                            sweepX_end_courant_modified = sweepX_end_courant - 0.07805

                            values_at_2_5_sec = df.loc[df['Time (s)'] == sweepX_end_courant_modified, selected_columns]

                            data_dict = {
                                "Courant (pA)": selected_columns,
                                "Potentiel (mV)": values_at_2_5_sec.iloc[0].tolist()
                            }

                            data_list.append(data_dict)
                            data_df = pd.DataFrame.from_dict(data_dict)

                            row_0pA = data_df[data_df['Courant (pA)'] == '0.00pA']
                            valeur_potentiel = row_0pA['Potentiel (mV)'].values[0]

                            data_df['Différence (mV)'] = data_df['Potentiel (mV)'] - valeur_potentiel
                        
                            intermediate_results = {}

                            filtered_df = df[(df['Time (s)'] >= sweepX_start_courant) & (df['Time (s)'] <= sweepX_end_courant)]
                            numeric_part = re.search(r'\d+.*\d', os.path.splitext(filename)[0]).group()

                            intermediate_results[numeric_part] = {'filtered_data': filtered_df}

                            filtered_df = filtered_df.copy()
                            df_graphe = df.copy()
                            columns_freq = filtered_df.columns[filtered_df.columns.get_loc('0.00pA'):]
                            print(columns_freq)
                            last_values_dict = {"Courant (pA)": [], "Freq (Hz)": [], "Freq Inst. (Hz)": []}
                            peaks_dict = {}
                            for column in columns_freq:
                                signal = filtered_df[column].values
                                time = filtered_df["Time (s)"].values

                                peak_indices, _ = find_peaks(signal, height=0)

                                # Récupérer les temps et les vraies valeurs des minima
                                peak_times = time[peak_indices]  # Temps des minima
                                peak_values = signal[peak_indices]  # Valeurs originales des minima
                                
                                peaks_dict[column] = {"Time (s)": peak_times, "Max Peak Value": peak_values}
                                num_events = len(peak_indices)

                                # Calculer les intervalles inter-spikes (ISI)
                                isi = peak_times[1:] - peak_times[:-1] if len(peak_times) > 1 else []

                                # Calculer les fréquences instantanées
                                frequencies = 1 / isi if len(isi) > 0 else []
                                mean_freq_inst = np.mean(frequencies) if len(frequencies) > 0 else 0

                                # Ajouter les données calculées au dictionnaire
                                last_values_dict["Courant (pA)"].append(column)
                                last_values_dict["Freq (Hz)"].append(num_events)
                                last_values_dict["Freq Inst. (Hz)"].append(mean_freq_inst)
                            print(last_values_dict)
                            print(f"✅ Peaks Dict après analyse : {peaks_dict}")
                            # Convertir le dictionnaire en DataFrame
                            freq_data_df = pd.DataFrame(last_values_dict)
                            filtered_df = filtered_df.copy()

                            # Fusion des deux DataFrames (Potentiels et Fréquences)
                            combined_df = pd.concat([data_df, freq_data_df], axis=1)

                            numeric_part = re.search(r'\d+.*\d', os.path.splitext(filename)[0]).group()
                            # Sauvegarde dans la feuille correspondante
                            with pd.ExcelWriter(combined_excel_file_path, engine='openpyxl', mode='a') as writer:
                                combined_df.to_excel(writer, sheet_name=numeric_part, index=False, startrow=0)

                            # Ajuster la largeur des colonnes
                            book = openpyxl.load_workbook(combined_excel_file_path)
                            sheet = book[numeric_part]

                            for column in sheet.columns:
                                max_length = 0
                                column = [cell for cell in column]
                                for cell in column:
                                    try:
                                        if len(str(cell.value)) > max_length:
                                            max_length = len(cell.value)
                                    except:
                                        pass
                                adjusted_width = (max_length + 2)
                                sheet.column_dimensions[column[0].column_letter].width = adjusted_width

                            book.save(combined_excel_file_path)



                            # Créer un graphe du DataFrame filtré avec les valeurs de "Max Peak Value"
                            #fig, ax = plt.subplots(figsize=(8, 6))

                            # Tracer les données filtrées pour la dernière colonne dans le graphe
                            #last_column = df_graphe.columns[-1]  # Obtenir le nom de la dernière colonne
                            # Sélectionner les valeurs de la dernière colonne dans la plage spécifiée
                            #mask = (df_graphe["Time (s)"] >= sweepX_start_courant - 0.5) & (df_graphe["Time (s)"] <= sweepX_end_courant + 0.5)
                            #selected_time = df_graphe["Time (s)"][mask]
                            #selected_data = df_graphe[last_column][mask]

                            # Tracer les données sélectionnées
                            #ax.plot(selected_time, selected_data, label=f'{last_column}', color=(98/255, 114/255, 164/255))

                            # Tracer les points de "Max Peak Value" pour la dernière colonne
                            #ax.scatter(peak_times, peak_values, color=(189/255, 147/255, 249/255), label='Spikes')

                            # Ajouter des labels et un titre pour le graphe
                            #ax.set_xlabel('Time (s)')
                            #ax.set_ylabel('Filtered Data')
                            #ax.set_title(f'{filename} - {last_column} - Filtered Data with Max Peak Values')
                            #ax.legend()
                            
                            SweepNavigator(abf, channel_to_analyze, sweepX_start_courant, sweepX_end_courant, sweepC_at_start_column, peaks_dict, index_start_courant)

                            graph_filename = os.path.join(details_folder_path, f'{os.path.splitext(filename)[0]}.png')
                            plt.savefig(graph_filename)

                            if verification_input == 'Y': 
                                # Afficher le graphe
                                plt.show(block=False)


                            current_value += 1
                            progress_percentage = (current_value / total_files) * 100
                            self.progress_bar.setValue(progress_percentage)
                            QApplication.processEvents()

                    if verification_input == 'Y':                
                        # Créer une fenêtre contextuelle pour attendre l'action de l'utilisateur
                        continue_window = self.create_continue_window()

                        # Attendre que l'utilisateur clique sur la fenêtre contextuelle
                        continue_window.exec()

                        
    def Capa_files(self, directory_path, progress_var, verification_input):
        total_files = get_total_abf_files_in_folders(directory_path, ["Capa"])
        current_value = 0

        for root, dirs, files in os.walk(directory_path):
            for folder_name in dirs:
                if folder_name == 'Capa':
                    folder_path = os.path.join(root, folder_name)
                    paramsDict = {}

                    for filename in os.listdir(folder_path):     
                        if filename.endswith(".abf"):
                            abf_file_path = os.path.join(folder_path, filename)
                            abfFiles = [f for f in os.listdir(folder_path) if f.endswith(".abf")]
                            # Charger le fichier ABF
                            abf = pyabf.ABF(abf_file_path)   
                                                                    
                            paramNames = ["Rm (MOhm)", "CmStep (pF)", "Ra (MOhm)"]
                            for channel in [0, 1]   :                 
                            # Trouver l'indice du dernier sweep
                                abf.setSweep(sweepNumber=abf.sweepCount - 1  , channel=channel) 
                                # Identifier les indices où abf.sweepC < 0
                                negative_indices = np.nonzero(abf.sweepC < 0)[0]

                                if negative_indices.size > 0:
                                    print(f"🔹 Données trouvées sur channel {channel}")
                                    break  # Sortir de la boucle si un channel contient des données
                            else:
                                print(f"❌ Aucune donnée détectée sur {filename}, skipping...")
                                continue  # Passer au fichier suivant si aucun channel ne fonctionne

                            # Trouver le début et la fin de chaque créneau
                            creaneaux = []
                            start_idx = negative_indices[0]

                            for i in range(1, len(negative_indices)):
                                if negative_indices[i] != negative_indices[i - 1] + 1:
                                    end_idx = negative_indices[i - 1]
                                    creaneaux.append((start_idx, end_idx))
                                    start_idx = negative_indices[i]

                            # Ajouter le dernier créneau
                            creaneaux.append((start_idx, negative_indices[-1]))

                            # Utiliser seulement le dernier créneau
                            start_voltage, end_voltage = creaneaux[-1]

                            sweepX_start_voltage = abf.sweepX[start_voltage]
                            index_start_voltage = np.where(abf.sweepX == sweepX_start_voltage)[0][0]

                            sweepX_end_voltage = abf.sweepX[end_voltage]
                            index_end_voltage = np.where(abf.sweepX == sweepX_end_voltage)[0][0]

                            sweep_data = {"Time (s)": abf.sweepX[index_start_voltage:index_end_voltage]}

                            for sweep_number in range(abf.sweepCount):
                                baseline_start = abf.sweepX[0]
                                baseline_end = baseline_start + 0.200

                                abf.setSweep(sweepNumber=sweep_number, baseline=[baseline_start, baseline_end])
                                column_name_y = f"Sweep_{sweep_number}_data"
                                sweep_data[column_name_y] = abf.sweepY[index_start_voltage:index_end_voltage]

                            df = pd.DataFrame(sweep_data)

                            # Trouver l'indice correspondant à 10 ms avant la fin
                            temps_10ms_before_end = sweepX_end_voltage - 0.010
                            index_10ms_before_end_df = (df['Time (s)'] - (sweepX_end_voltage - 0.010)).abs().idxmin()
                            
                            # Créer un dictionnaire pour stocker les moyennes de chaque colonne "Sweep"
                            mean_sweep = {}
                            # Créer un dictionnaire pour stocker les données modifiées de chaque colonne "Sweep"
                            modified_data = {'Time (s)': df['Time (s)']}  # Ajouter la colonne "Time (s)" au dictionnaire

                            # Calculer la moyenne pour chaque colonne "Sweep"
                            for column in df.columns[1:]:  # Ignorer la colonne "Time (s)"
                                mean_sweep[column] = df[column].loc[index_10ms_before_end_df:].mean()
                                modified_data[f'{column}_modified'] = df[column] - mean_sweep[column]

                            overall_mean = sum(mean_sweep.values()) / len(mean_sweep)  
                            print (overall_mean)  

                            # Calculer la résistance membranaire Rm
                            Rm = (-5 / overall_mean) * 1000
                            
                            print("La valeur de Rm est :", Rm)

                            # Créer un nouveau DataFrame à partir du dictionnaire de données modifiées
                            df_modified = pd.DataFrame(modified_data)
                            
                            # Afficher le nouveau DataFrame
                            df_average_Capa = df_modified.iloc[:, 1:].mean(axis=1)
                            df_average_Capa = pd.DataFrame({'Time (s)': df_modified['Time (s)'], 'Average': df_average_Capa})
                            print (df_average_Capa)
                                                
                            # Trouver la valeur de temps correspondant à 30 millisecondes
                            temps_30ms = df_average_Capa['Time (s)'].iloc[0] + 0.030
                            # Trouver l'indice de la valeur la plus proche de temps_30ms
                            index_temps_30ms = (np.abs(df_average_Capa['Time (s)'] - temps_30ms)).idxmin()

                            # Sélectionner uniquement les données jusqu'à 30 millisecondes
                            df_30ms = df_average_Capa[df_average_Capa['Time (s)'] <= temps_30ms]
                            print (df_30ms)
                            # Calculer l'aire sous la courbe de la colonne "Average" jusqu'à 30 millisecondes
                            area_average_30ms = np.trapz(df_30ms['Average'], df_30ms['Time (s)'])

                            Capa = (area_average_30ms*1000)/-5
                            # Imprimer l'aire sous la courbe jusqu'à 30 millisecondes
                            print("La capa est :", Capa)
                            Peak = df_average_Capa['Average'].min()
                            index_Peak = df_average_Capa['Average'].idxmin()
                            
                            # Sélectionner les lignes entre l'index de la valeur négative et temps 30 ms
                            df_subset = df_average_Capa.loc[index_Peak:index_temps_30ms]
                            print (df_subset)
                            area_subset = np.trapz(df_subset['Average'], df_subset['Time (s)']) 
                            print (area_subset)
                            Tau = area_subset*1000 / Peak
                            Rs = Tau/Capa*1000
                            paramsDict[filename] = {'Rm (MOhm)': Rm, 'Cm (pF)': Capa, 'Ra (MOhm)': Rs}

                        print (paramsDict)
                        df_Capa = pd.DataFrame.from_dict(paramsDict, orient='index').reset_index()
                        df_Capa.columns = ['File'] + paramNames
                        excel_output_path = os.path.join(folder_path, "param_values.xlsx")
                        df_Capa.to_excel(excel_output_path, index=False)

                        # Create ExcelWriter and save the DataFrame
                        with pd.ExcelWriter(excel_output_path, engine='openpyxl') as writer:
                            df_Capa.to_excel(writer, index=False)

                            # Access the default sheet
                            sheet = writer.sheets['Sheet1']

                            # Adjust column widths based on content in the first row
                            for column in sheet.columns:
                                max_length = 0
                                column = [cell for cell in column]
                                for cell in column:
                                    try:  # Necessary to avoid error on empty cells
                                        if len(str(cell.value)) > max_length:
                                            max_length = len(cell.value)
                                    except:
                                        pass
                                adjusted_width = (max_length + 2)
                                sheet.column_dimensions[column[0].column_letter].width = adjusted_width

                            current_value += 1
                            progress_percentage = (current_value / total_files) * 100
                            self.progress_bar.setValue(progress_percentage)
                            QApplication.processEvents()

    def Em_files(self, directory_path, progress_var, verification_input):
            total_files = get_total_abf_files_in_folders(directory_path, ["Em"])
            current_value = 0 
            avg_values_dict = {}
            
            for root, dirs, files in os.walk(directory_path):
                for folder_name in dirs:
                    if folder_name == 'Em':
                        avg_values_dict = {}  
                        folder_path = os.path.join(root, folder_name)  
                        global_avg_list = []  
                        
                        for filename in os.listdir(folder_path):    
                            if filename.endswith(".abf"):                           
                                
                                abf_file_path = os.path.join(folder_path, filename)
                                abf = pyabf.ABF(abf_file_path)     
                                details_folder_path = os.path.join(folder_path, 'Graphes')
                                os.makedirs(details_folder_path, exist_ok=True)

                                for channel in [0, 1]:  
                                    abf.setSweep(sweepNumber=abf.sweepCount - 1, channel=channel)
                                    signal = abf.sweepY

                                    if np.any(np.abs(signal) > 1):
                                        break

                                else:
                                    continue
                                                            
                                sweep_data = {"Time (s)": abf.sweepX}
                                all_avg_freqs = []
                                all_cv_isis = []
                                all_Em = []

                                for sweep_number in range(abf.sweepCount):
                                    
                                    time = abf.sweepX
                                    signal = abf.sweepY
                                    
                                    column_name_y = f"Sweep {sweep_number:.2f}"
                                    sweep_data[column_name_y] = abf.sweepY
                                    
                                    peak_indices, _ = find_peaks(signal, height=0)
                                    peak_times = time[peak_indices]  # Temps des minima
                                    peak_values = signal[peak_indices]  # Valeurs originales des minima

                                    if len(peak_times) > 1:
                                        isi = np.diff(peak_times)    
                                        frequencies = 1 / isi
                                        avg_freq = np.mean(frequencies)
                                        isi_cv = np.std(isi) / np.mean(isi)
                                    else:
                                        avg_freq = 0
                                        isi_cv = 0

                                    sweep_name = f"{filename}_Sweep_{sweep_number}"
                                    vrest = np.median(signal)

                                    avg_values_dict[sweep_name] = {
                                        "Avg Freq Inst": avg_freq,
                                        "Avg CV": isi_cv,
                                        "Vrest": vrest
                                    }

                                    all_avg_freqs.append(avg_freq)
                                    all_cv_isis.append(isi_cv)
                                    all_Em.append(vrest)

                                # Créer un DataFrame pandas
                                df = pd.DataFrame(sweep_data)

                                global_avg_freq = np.mean(all_avg_freqs) if all_avg_freqs else 0
                                global_cv_isi = np.mean(all_cv_isis) if all_cv_isis else 0
                                global_avg_Em = np.mean(all_Em) if all_Em else 0
                                                    
                                # Ajouter à une liste pour la sauvegarde sous format Excel
                                global_avg_list.append({
                                    "Filename": filename,
                                    "Global Avg Freq Inst": global_avg_freq,
                                    "Global Avg CV": global_cv_isi,
                                    "Global Avg Em": global_avg_Em
                                })

                                print(global_avg_list)

                                # Créer un graphe du DataFrame filtré avec les valeurs de "Max Peak Value"
                                fig, ax = plt.subplots(figsize=(8, 6))

                                # Tracer les données filtrées pour la dernière colonne dans le graphe
                                last_column = df.columns[-1]  # Obtenir le nom de la dernière colonne

                                # Tracer les données sélectionnées
                                ax.plot(time, df.iloc[:, -1], color=(98/255, 114/255, 164/255))

                                # Tracer les points de "Max Peak Value" pour la dernière colonne
                                ax.scatter(peak_times, peak_values, color=(189/255, 147/255, 249/255), label='Detected spikes')
                                global_avg_Em = next((item["Global Avg Em"] for item in global_avg_list if item["Filename"] == filename), None)

                                if global_avg_Em is not None:
                                    ax.axhline(y=global_avg_Em, color=(189/255, 147/255, 249/255), linestyle='--', label='Resting potential')
                                
                                # Ajouter des labels et un titre pour le graphe
                                ax.set_xlabel('Time (s)')
                                ax.set_ylabel('Filtered Data')
                                ax.set_title(f'{filename} - {last_column}')
                                ax.legend()

                                graph_filename = os.path.join(details_folder_path, f'{os.path.splitext(filename)[0]}.png')
                                plt.savefig(graph_filename)

                                if verification_input == 'Y': 
                                    # Afficher le graphe
                                    plt.show(block=False)

                        # Sauvegarder les moyennes globales dans un fichier Excel
                        df_global_avg = pd.DataFrame(global_avg_list)
                        output_excel_path = os.path.join(folder_path, "Em_file.xlsx")
                        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
                            df_global_avg.to_excel(writer, index=False, sheet_name='Sheet1')

                            # Ajustement des colonnes
                            worksheet = writer.sheets['Sheet1']
                            for col in worksheet.columns:
                                max_length = max(len(str(cell.value)) for cell in col if cell.value)
                                col_letter = col[0].column_letter
                                worksheet.column_dimensions[col_letter].width = max_length + 2

                                                                    
                        current_value += 1
                        progress_percentage = (current_value / total_files) * 100
                        self.progress_bar.setValue(progress_percentage)
                        QApplication.processEvents()
                        
                        if verification_input == 'Y':                            
                            # Créer une fenêtre contextuelle pour attendre l'action de l'utilisateur
                            continue_window = self.create_continue_window()

                            # Attendre que l'utilisateur clique sur la fenêtre contextuelle
                            continue_window.exec()


    # Fonction principale
    def process_files(self):
        directory_path = self.browse_input.text()

        if not directory_path:
            QMessageBox.critical(self, "Error", "Please select a folder.")
            return

        # Vérifier si l'utilisateur a activé la vérification manuelle
        verification_input = 'Y' if self.manual_verification_checkbox.isChecked() else 'N'

        # Réinitialiser la barre de progression
        self.progress_bar.setValue(0)

        # Liste des fonctions à exécuter
        analysis_functions = [
             self.AMPA_NMDA_files,
            self.Stim_files,
            self.Rheo_files,
            self.Cellattached_files,
            self.Capa_files,
            self.Em_files,
            self.CCIV_files
        ]

        # Nombre total d'étapes pour la progression
        total_steps = len(analysis_functions)
        for step, analysis_function in enumerate(analysis_functions, start=1):
            analysis_function(directory_path, self.progress_bar, verification_input)
            self.progress_bar.setValue(int((step / total_steps) * 100))
            QApplication.processEvents()

        # Afficher un message d'information à la fin
        msg_box = QMessageBox(self)
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setWindowTitle("Information")
        msg_box.setText("File processing completed successfully.")
        msg_box.setStyleSheet("""
            QMessageBox {
                background-color: rgb(40, 44, 52);
            }
        """)
        msg_box.exec()
        
# Lancement de l'application
if __name__ == "__main__":
    app = QApplication([])

    # Créer et afficher la fenêtre principale
    window = MainWindow()
    window.show()

    app.exec()







     

