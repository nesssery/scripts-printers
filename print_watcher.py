import os
import time
import json
from zipfile import ZipFile
from io import BytesIO
import requests

# Configuration
WATCH_FOLDER = "/tmp/print_jobs"  # Dossier surveillé mis à jour
OUTPUT_FOLDER = "/tmp/print_output"  # Dossier pour simuler l'impression
CHECK_INTERVAL = 5  # Intervalle de vérification en secondes

# Crée les dossiers s'ils n'existent pas
if not os.path.exists(WATCH_FOLDER):
    os.makedirs(WATCH_FOLDER)
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

def simulate_print(file_content, file_type, order_id):
    # Détermine l'extension correcte en fonction du type
    extension = '.pdf' if file_type == 'facture_A4' else '.json'
    # Simule l'impression en enregistrant un fichier avec l'extension appropriée
    output_filename = f"{file_type.split('_')[0].lower()}_{order_id}{extension}"
    output_path = os.path.join(OUTPUT_FOLDER, output_filename)
    with open(output_path, 'wb') as target:
        target.write(file_content)
    print(f"Simulé l'impression de {file_type} : {output_path}")

def process_zip(zip_path):
    try:
        with ZipFile(zip_path, 'r') as zip_file:
            order_id = os.path.basename(zip_path).split("_")[1].split(".")[0]  # Extrait l'ID du nom du fichier
            for file_name in zip_file.namelist():
                if file_name.startswith("facture_") and file_name.endswith(".pdf"):
                    print("____________________", file_name)
                    with zip_file.open(file_name) as facture_file:
                        simulate_print(facture_file.read(), "facture_A4", order_id)
                elif file_name.startswith("ticket_") and file_name.endswith(".json"):
                    with zip_file.open(file_name) as ticket_file:
                        ticket_data = json.load(ticket_file)
                        simulate_print(json.dumps(ticket_data).encode(), "ticket_thermique", order_id)
    except Exception as e:
        print(f"Erreur lors du traitement du ZIP : {e}")

def main():
    print(f"Surveillance démarrée. Vérification du dossier : {WATCH_FOLDER}")
    processed_files = set()

    while True:
        try:
            # Vérifier les nouveaux fichiers dans le dossier /tmp/print_jobs
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