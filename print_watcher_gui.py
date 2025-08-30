import os
import time
import json
from zipfile import ZipFile
import win32print
from PySide6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QTextEdit, QLabel
from PySide6.QtCore import QThread, Signal, Qt, QDir
import sys
import shutil
import re  # Ajout pour utiliser les expressions régulières

# Configuration
WATCH_FOLDER = r"C:\Temp\print_jobs"  # Destination finale
DOWNLOADS_FOLDER = r"C:\Users\aissi\Downloads"  # Remplace <TonNom> par ton nom d'utilisateur Windows
OUTPUT_FOLDER = r"C:\Temp\print_output"
CHECK_INTERVAL = 5
LOG_FILE = r"C:\Temp\logs\output.log"
ERROR_LOG_FILE = r"C:\Temp\logs\error.log"
PRINTERS = {
    "facture_A4": "EPSON L3250 Series",
    "ticket_thermique": "XP-80C"
}

# Crée les dossiers s'ils n'existent pas
for path in [WATCH_FOLDER, OUTPUT_FOLDER, os.path.dirname(LOG_FILE)]:
    if not os.path.exists(path):
        os.makedirs(path)

class WorkerThread(QThread):
    log_signal = Signal(str)
    error_signal = Signal(str)

    def __init__(self):
        super().__init__()
        self.running = False

    def run(self):
        self.running = True
        processed_files = set()
        self.log_signal.emit(f"Surveillance démarrée. Vérification du dossier : {WATCH_FOLDER}")
        while self.running:
            try:
                if not os.path.exists(WATCH_FOLDER):
                    self.error_signal.emit(f"Dossier {WATCH_FOLDER} introuvable.")
                    time.sleep(CHECK_INTERVAL)
                    continue
                # Vérifie les nouveaux fichiers téléchargés dans Downloads
                for file in os.listdir(DOWNLOADS_FOLDER):
                    if file.endswith(".zip") and re.match(r'^order_\d+\.zip$', file):  # Filtre pour order_<nombre>.zip
                        source_path = os.path.join(DOWNLOADS_FOLDER, file)
                        dest_path = os.path.join(WATCH_FOLDER, file)
                        if not os.path.exists(dest_path):  # Évite les doublons
                            shutil.move(source_path, dest_path)
                            processed_files.add(dest_path)
                            self.log_signal.emit(f"Fichier déplacé : {file} vers {WATCH_FOLDER}")
                            self.process_zip(dest_path)
                            processed_files.remove(dest_path)
                # Vérifie les fichiers déjà dans WATCH_FOLDER
                for file in os.listdir(WATCH_FOLDER):
                    if file.endswith(".zip") and re.match(r'^order_\d+\.zip$', file):  # Filtre pour order_<nombre>.zip
                        file_path = os.path.join(WATCH_FOLDER, file)
                        if file_path not in processed_files:
                            self.log_signal.emit(f"Nouveau fichier ZIP détecté : {file}")
                            self.process_zip(file_path)
                            processed_files.add(file_path)
                            os.remove(file_path)
                            processed_files.remove(file_path)
            except Exception as e:
                self.error_signal.emit(f"Erreur dans la surveillance : {e}")
            time.sleep(CHECK_INTERVAL)

    def process_zip(self, zip_path):
        try:
            with ZipFile(zip_path, 'r') as zip_file:
                order_id = os.path.basename(zip_path).split("_")[1].split(".")[0]
                for file_name in zip_file.namelist():
                    if file_name.startswith("facture_") and file_name.endswith(".pdf"):
                        self.log_signal.emit(f"Traitement de {file_name}")
                        with zip_file.open(file_name) as facture_file:
                            file_content = facture_file.read()
                            printer_name = PRINTERS.get("facture_A4")
                            if printer_name and self.is_printer_available(printer_name):
                                self.print_to_printer(file_content, printer_name)
                            else:
                                self.error_signal.emit(f"Imprimante {printer_name} non disponible ou non configurée.")
                    elif file_name.startswith("ticket_") and file_name.endswith(".json"):
                        with zip_file.open(file_name) as ticket_file:
                            ticket_data = json.load(ticket_file)
                            ticket_text = f"Ticket #{order_id}\n{'='*20}\n" + \
                                          "\n".join([f"{k}: {v}" for k, v in ticket_data.items()]) + \
                                          f"\n{'='*20}"
                            printer_name = PRINTERS.get("ticket_thermique")
                            if printer_name and self.is_printer_available(printer_name):
                                self.print_to_printer(ticket_text.encode('utf-8'), printer_name)
                            else:
                                self.error_signal.emit(f"Imprimante {printer_name} non disponible ou non configurée.")
        except Exception as e:
            self.error_signal.emit(f"Erreur lors du traitement du ZIP : {e}")

    def print_to_printer(self, file_content, printer_name):
        try:
            hPrinter = win32print.OpenPrinter(printer_name)
            hJob = win32print.StartDocPrinter(hPrinter, 1, ("Impression", None, "RAW"))
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, file_content)
            win32print.EndPagePrinter(hPrinter)
            win32print.EndDocPrinter(hPrinter)
            win32print.ClosePrinter(hPrinter)
            self.log_signal.emit(f"Impression envoyée avec succès à {printer_name}")
        except Exception as e:
            self.error_signal.emit(f"Erreur lors de l'impression sur {printer_name} : {e}")

    def is_printer_available(self, printer_name):
        try:
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
            return any(printer[2] == printer_name for printer in printers)
        except Exception:
            return False

    def stop(self):
        self.running = False

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Print Watcher")
        self.setGeometry(100, 100, 600, 400)

        # Layout and widgets
        layout = QVBoxLayout()
        self.status_label = QLabel("Statut : Arrêté")
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.start_button = QPushButton("Démarrer la surveillance")
        self.stop_button = QPushButton("Arrêter la surveillance")
        self.stop_button.setEnabled(False)

        layout.addWidget(self.status_label)
        layout.addWidget(self.log_text)
        layout.addWidget(self.start_button)
        layout.addWidget(self.stop_button)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # Worker thread
        self.worker = WorkerThread()
        self.worker.log_signal.connect(self.append_log)
        self.worker.error_signal.connect(self.append_error)

        # Connect buttons
        self.start_button.clicked.connect(self.start_watching)
        self.stop_button.clicked.connect(self.stop_watching)

    def append_log(self, message):
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        log_message = f"[{timestamp}] [INFO] {message}"
        self.log_text.append(log_message)
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(log_message + "\n")

    def append_error(self, message):
        timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
        error_message = f"[{timestamp}] [ERROR] {message}"
        self.log_text.append(error_message)
        with open(ERROR_LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(error_message + "\n")

    def start_watching(self):
        if not self.worker.isRunning():
            self.worker.start()
            self.status_label.setText("Statut : En cours")
            self.start_button.setEnabled(False)
            self.stop_button.setEnabled(True)
        else:
            self.append_log("La surveillance est déjà en cours.")

    def stop_watching(self):
        if self.worker.isRunning():
            self.worker.stop()
            self.worker.wait()  # Attend que le thread se termine
            self.status_label.setText("Statut : Arrêté")
            self.start_button.setEnabled(True)
            self.stop_button.setEnabled(False)
        else:
            self.append_log("La surveillance est déjà arrêtée.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())