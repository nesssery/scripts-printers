import os
import time
import json
from zipfile import ZipFile
import win32print

# Configuration
WATCH_FOLDER = "/print_jobs"  # Dossier surveillé (monté depuis l'hôte)
OUTPUT_FOLDER = "/print_output"  # Dossier de sortie (monté depuis l'hôte)
CHECK_INTERVAL = 5  # Intervalle de vérification en secondes
LOG_FILE = "/app/logs/output.log"
ERROR_LOG_FILE = "/app/logs/error.log"
PRINTERS = {
    "facture_A4": "EPSON L3250 Series",  # Imprimante pour les factures PDF
    "ticket_thermique": "XP-80C"         # Imprimante pour les tickets (JSON converti)
}

# Crée les dossiers s'ils n'existent pas
if not os.path.exists(WATCH_FOLDER):
    os.makedirs(WATCH_FOLDER)
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

def log(message):
    with open(LOG_FILE, 'a') as f:
        f.write(f"{time.ctime()}: {message}\n")

def log_error(message):
    with open(ERROR_LOG_FILE, 'a') as f:
        f.write(f"{time.ctime()}: {message}\n")

def print_to_printer(file_content, printer_name):
    try:
        hPrinter = win32print.OpenPrinter(printer_name)
        hJob = win32print.StartDocPrinter(hPrinter, 1, ("Impression de facture", None, "RAW"))
        win32print.StartPagePrinter(hPrinter)
        win32print.WritePrinter(hPrinter, file_content)
        win32print.EndPagePrinter(hPrinter)
        win32print.EndDocPrinter(hPrinter)
        win32print.ClosePrinter(hPrinter)
        log(f"Impression envoyée à {printer_name}")
    except Exception as e:
        log_error(f"Erreur lors de l'impression sur {printer_name} : {e}")

def process_zip(zip_path):
    try:
        with ZipFile(zip_path, 'r') as zip_file:
            order_id = os.path.basename(zip_path).split("_")[1].split(".")[0]
            for file_name in zip_file.namelist():
                if file_name.startswith("facture_") and file_name.endswith(".pdf"):
                    log(f"____________________ {file_name}")
                    with zip_file.open(file_name) as facture_file:
                        file_content = facture_file.read()
                        printer_name = PRINTERS.get("facture_A4")
                        if printer_name:
                            print_to_printer(file_content, printer_name)
                        else:
                            log_error("Imprimante pour facture non configurée.")
                elif file_name.startswith("ticket_") and file_name.endswith(".json"):
                    with zip_file.open(file_name) as ticket_file:
                        ticket_data = json.load(ticket_file)
                        ticket_text = json.dumps(ticket_data, ensure_ascii=False).encode('utf-8')
                        printer_name = PRINTERS.get("ticket_thermique")
                        if printer_name:
                            print_to_printer(ticket_text, printer_name)
                        else:
                            log_error("Imprimante pour ticket non configurée.")
    except Exception as e:
        log_error(f"Erreur lors du traitement du ZIP : {e}")

def main():
    log(f"Surveillance démarrée. Vérification du dossier : {WATCH_FOLDER}")
    processed_files = set()

    while True:
        try:
            for file in os.listdir(WATCH_FOLDER):
                file_path = os.path.join(WATCH_FOLDER, file)
                if file.endswith(".zip") and file_path not in processed_files:
                    log(f"Nouveau fichier ZIP détecté : {file}")
                    process_zip(file_path)
                    processed_files.add(file_path)
                    os.remove(file_path)
                    processed_files.remove(file_path)
        except Exception as e:
            log_error(f"Erreur dans la surveillance : {e}")
        time.sleep(CHECK_INTERVAL)

if __name__ == "__main__":
    main()