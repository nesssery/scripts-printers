import os
import time
import json
from zipfile import ZipFile
import win32print
from PySide6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QWidget, QPushButton, QTextEdit, QLabel
from PySide6.QtCore import QThread, Signal
import sys

# Configuration
WATCH_FOLDER = r"C:\Temp\print_jobs"
OUTPUT_FOLDER = r"C:\Temp\print_output"
CHECK_INTERVAL = 5
LOG_FILE = r"C:\Temp\logs\output.log"
ERROR_LOG_FILE = r"C:\Temp\logs\error.log"
PRINTERS = {
    "facture_A4": "EPSON L3250 Series",
    "ticket_thermique": "XP-80C"
}

# Create directories if they don't exist
if not os.path.exists(WATCH_FOLDER):
    os.makedirs(WATCH_FOLDER)
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)
if not os.path.exists(os.path.dirname(LOG_FILE)):
    os.makedirs(os.path.dirname(LOG_FILE))

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
                for file in os.listdir(WATCH_FOLDER):
                    file_path = os.path.join(WATCH_FOLDER, file)
                    if file.endswith(".zip") and file_path not in processed_files:
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
                        self.log_signal.emit(f"____________________ {file_name}")
                        with zip_file.open(file_name) as facture_file:
                            file_content = facture_file.read()
                            printer_name = PRINTERS.get("facture_A4")
                            if printer_name:
                                self.print_to_printer(file_content, printer_name)
                            else:
                                self.error_signal.emit("Imprimante pour facture non configurée.")
                    elif file_name.startswith("ticket_") and file_name.endswith(".json"):
                        with zip_file.open(file_name) as ticket_file:
                            ticket_data = json.load(ticket_file)
                            ticket_text = json.dumps(ticket_data, ensure_ascii=False).encode('utf-8')
                            printer_name = PRINTERS.get("ticket_thermique")
                            if printer_name:
                                self.print_to_printer(ticket_text, printer_name)
                            else:
                                self.error_signal.emit("Imprimante pour ticket non configurée.")
        except Exception as e:
            self.error_signal.emit(f"Erreur lors du traitement du ZIP : {e}")

    def print_to_printer(self, file_content, printer_name):
        try:
            hPrinter = win32print.OpenPrinter(printer_name)
            hJob = win32print.StartDocPrinter(hPrinter, 1, ("Impression de facture", None, "RAW"))
            win32print.StartPagePrinter(hPrinter)
            win32print.WritePrinter(hPrinter, file_content)
            win32print.EndPagePrinter(hPrinter)
            win32print.EndDocPrinter(hPrinter)
            win32print.ClosePrinter(hPrinter)
            self.log_signal.emit(f"Impression envoyée à {printer_name}")
        except Exception as e:
            self.error_signal.emit(f"Erreur lors de l'impression sur {printer_name} : {e}")

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
        self.log_text.append(f"[INFO] {message}")
        with open(LOG_FILE, 'a') as f:
            f.write(f"{time.ctime()}: {message}\n")

    def append_error(self, message):
        self.log_text.append(f"[ERROR] {message}")
        with open(ERROR_LOG_FILE, 'a') as f:
            f.write(f"{time.ctime()}: {message}\n")

    def start_watching(self):
        self.worker.start()
        self.status_label.setText("Statut : En cours")
        self.start_button.setEnabled(False)
        self.stop_button.setEnabled(True)

    def stop_watching(self):
        self.worker.stop()
        self.status_label.setText("Statut : Arrêté")
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())