import os
import time
import json
from zipfile import ZipFile
import win32print
from io import BytesIO

# Configuration
WATCH_FOLDER = "C:\\Temp\\print_jobs"  # Dossier surveillé mis à jour pour Windows
OUTPUT_FOLDER = "C:\\Temp\\print_output"  # Dossier pour logs ou fichiers temporaires
CHECK_INTERVAL = 5  # Intervalle de vérification en secondes
PRINTERS = {
    "facture_A4": "EPSON L3250 Series",  # Imprimante pour les factures PDF
    "ticket_thermique": "XP-80C"         # Imprimante pour les tickets (JSON converti)
}

# Crée les dossiers s'ils n'existent pas
if not os.path.exists(WATCH_FOLDER):
    os.makedirs(WATCH_FOLDER)
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

def print_to_printer(file_content, printer_name):
    try:
        hPrinter = win32print.OpenPrinter(printer_name)
        hJob = win32print.StartDocPrinter(hPrinter, 1, ("Impression de facture", None, "RAW"))
        win32print.StartPagePrinter(hPrinter)
        win32print.WritePrinter(hPrinter, file_content)
        win32print.EndPagePrinter(hPrinter)
        win32print.EndDocPrinter(hPrinter)
        win32print.ClosePrinter(hPrinter)
        print(f"Impression envoyée à {printer_name}")
    except Exception as e:
        print(f"Erreur lors de l'impression sur {printer_name} : {e}")

def process_zip(zip_path):
    try:
        with ZipFile(zip_path, 'r') as zip_file:
            order_id = os.path.basename(zip_path).split("_")[1].split(".")[0]  # Extrait l'ID du nom du fichier
            for file_name in zip_file.namelist():
                if file_name.startswith("facture_") and file_name.endswith(".pdf"):
                    with zip_file.open(file_name) as facture_file:
                        file_content = facture_file.read()
                        printer_name = PRINTERS.get("facture_A4")
                        if printer_name:
                            print_to_printer(file_content, printer_name)
                        else:
                            print("Imprimante pour facture non configurée.")
                elif file_name.startswith("ticket_") and file_name.endswith(".json"):
                    with zip_file.open(file_name) as ticket_file:
                        ticket_data = json.load(ticket_file)
                        # Convertir les données JSON en texte simple pour l'impression thermique
                        ticket_text = json.dumps(ticket_data, ensure_ascii=False).encode('utf-8')
                        printer_name = PRINTERS.get("ticket_thermique")
                        if printer_name:
                            print_to_printer(ticket_text, printer_name)
                        else:
                            print("Imprimante pour ticket non configurée.")
    except Exception as e:
        print(f"Erreur lors du traitement du ZIP : {e}")

def main():
    print(f"Surveillance démarrée. Vérification du dossier : {WATCH_FOLDER}")
    processed_files = set()

    while True:
        try:
            # Vérifier les nouveaux fichiers dans le dossier WATCH_FOLDER
            for file in os.listdir(WATCH_FOLDER):
                file_path = os.path.join(WATCH_FOLDER, file)
                if file.endswith(".zip") and file_path not in processed_files:
                    print(f"Nouveau fichier ZIP détecté : {file}")
                    process_zip(file_path)
                    processed_files.add(file_path)
                    # Supprimer le fichier après traitement pour éviter les doublons
                    os.remove(file_path)
                    processed_files.remove(file_path)
        except Exception as e:
            print(f"Erreur dans la surveillance : {e}")
        time.sleep(CHECK_INTERVAL)

if __name__ == "__main__":
    main()