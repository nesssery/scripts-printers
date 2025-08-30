import os
import time
import json
from zipfile import ZipFile
import win32print
from PySide6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QTextEdit, QLabel
from PySide6.QtCore import QThread, Signal
import sys
import shutil
import re
from pdf2image import convert_from_bytes
from PIL import Image
import io

# Configuration
WATCH_FOLDER = r"C:\Temp\print_jobs"
DOWNLOADS_FOLDER = r"C:\Users\aissi\Downloads"
OUTPUT_FOLDER = r"C:\Temp\print_output"
CHECK_INTERVAL = 5
LOG_FILE = r"C:\Temp\logs\output.log"
ERROR_LOG_FILE = r"C:\Temp\logs\error.log"
PRINTERS = {
    "facture_A4": "EPSON L3250 Series",
    "ticket_thermique": "Xprinter XP-80C"
}

# Créer les dossiers s'ils n'existent pas
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
        self.log_signal.emit("Surveillance démarrée.")
        while self.running:
            try:
                if os.path.exists(DOWNLOADS_FOLDER):
                    for file in os.listdir(DOWNLOADS_FOLDER):
                        if file.endswith(".zip") and re.match(r'^order_\d+\.zip$', file):
                            source_path = os.path.join(DOWNLOADS_FOLDER, file)
                            dest_path = os.path.join(WATCH_FOLDER, file)
                            if not os.path.exists(dest_path):
                                shutil.move(source_path, dest_path)
                                self.log_signal.emit(f"Fichier déplacé : {file}")
                                self.process_zip(dest_path)
                for file in os.listdir(WATCH_FOLDER):
                    if file.endswith(".zip") and re.match(r'^order_\d+\.zip$', file):
                        file_path = os.path.join(WATCH_FOLDER, file)
                        if file_path not in processed_files:
                            self.log_signal.emit(f"Nouveau ZIP : {file}")
                            self.process_zip(file_path)
                            processed_files.add(file_path)
                            os.remove(file_path)
                            processed_files.remove(file_path)
            except Exception as e:
                self.error_signal.emit(f"Erreur : {e}")
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
                            self.print_pdf(file_content, printer_name, order_id)
                    elif file_name.startswith("ticket_") and file_name.endswith(".json"):
                        self.log_signal.emit(f"Traitement de {file_name}")
                        with zip_file.open(file_name) as ticket_file:
                            ticket_data = json.load(ticket_file)
                            ticket_text = f"Ticket #{order_id}\n{'='*20}\n" + \
                                          "\n".join([f"{k}: {v}" for k, v in ticket_data.items()]) + \
                                          f"\n{'='*20}"
                            printer_name = PRINTERS.get("ticket_thermique")
                            self.print_text(ticket_text.encode('utf-8'), printer_name)
        except Exception as e:
            self.error_signal.emit(f"Erreur ZIP : {e}")

    def print_pdf(self, file_content, printer_name, order_id):
        try:
            images = convert_from_bytes(file_content)
            if not images:
                self.error_signal.emit("Erreur : Aucun contenu dans le PDF.")
                return
            temp_image_path = os.path.join(OUTPUT_FOLDER, f"facture_{order_id}.png")
            images[0].save(temp_image_path, "PNG")
            hPrinter = win32print.OpenPrinter(printer_name)
            hJob = win32print.StartDocPrinter(hPrinter, 1, ("Facture", None, "RAW"))
            win32print.StartPagePrinter(hPrinter)
            with open(temp_image_path, 'rb') as img_file:
                win32print.WritePrinter(hPrinter, img_file.read())
            win32print.EndPagePrinter(hPrinter)
            win32print.EndDocPrinter(hPrinter)
            win32print.ClosePrinter(hPrinter)
            self.log_signal.emit(f"PDF (image) imprimé sur {printer_name}")
            os.remove(temp_image_path)
        except Exception as e:
            self.error_signal.emit(f"Erreur d'impression PDF : {e}")

    def print_text(self, file_content, printer_name):
        try:
            hPrinter = win32print.OpenPrinter(printer_name)
            hJob = win32print.StartDocPrinter(hPrinter, 1, ("Ticket", None, "RAW"))
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, file_content)
            win32print.EndPagePrinter(hPrinter)
            win32print.EndDocPrinter(hPrinter)
            win32print.ClosePrinter(hPrinter)
            self.log_signal.emit(f"Ticket imprimé sur {printer_name}")
        except Exception as e:
            self.error_signal.emit(f"Erreur d'impression ticket : {e}")

    def stop(self):
        self.running = False

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Print Watcher")
        self.setGeometry(100, 100, 600, 400)
        layout = QVBoxLayout()
        self.status_label = QLabel("Statut : Arrêté")
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.start_button = QPushButton("Démarrer")
        self.stop_button = QPushButton("Arrêter")
        self.stop_button.setEnabled(False)
        layout.addWidget(self.status_label)
        layout.addWidget(self.log_text)
        layout.addWidget(self.start_button)
        layout.addWidget(self.stop_button)
        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
        self.worker = WorkerThread()
        self.worker.log_signal.connect(self.append_log)
        self.worker.error_signal.connect(self.append_error)
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
            self.append_log("Surveillance déjà en cours.")

    def stop_watching(self):
        if self.worker.isRunning():
            self.worker.stop()
            self.worker.wait()
            self.status_label.setText("Statut : Arrêté")
            self.start_button.setEnabled(True)
            self.stop_button.setEnabled(False)
        else:
            self.append_log("Surveillance déjà arrêtée.")

if __name__ == "__main__":
    try:
        import win32print
        import pdf2image
        import PIL
    except ImportError as e:
        print(f"Erreur : Module manquant. Installez avec 'pip install {e.name}'.")
        sys.exit(1)
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())