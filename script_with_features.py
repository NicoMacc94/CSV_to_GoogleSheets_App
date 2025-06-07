#!/usr/bin/env python3

import sys
import os
import gspread
from gspread.exceptions import SpreadsheetNotFound, WorksheetNotFound
import tkinter as tk
from tkinter import filedialog, messagebox
import json
from difflib import SequenceMatcher
import csv

def similar(a, b):
    """Calculate string similarity ratio between two strings."""
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()

def normalize_name(name):
    """Normalize a name by removing extra spaces and converting to lowercase."""
    return ' '.join(name.lower().split())

def are_names_similar(name1, name2, threshold=0.8):
    """Check if two names are similar enough to be considered the same person."""
    # Normalize both names
    name1 = normalize_name(name1)
    name2 = normalize_name(name2)
    
    # Direct match
    if name1 == name2:
        return True
    
    # Split names into parts
    parts1 = set(name1.split())
    parts2 = set(name2.split())
    
    # Check if all parts of one name are in the other
    if parts1.issubset(parts2) or parts2.issubset(parts1):
        return True
    
    # Check similarity ratio
    return similar(name1, name2) > threshold

def load_last_directory():
    """Load the last used directory from a config file."""
    config_file = 'last_directory.json'
    if os.path.exists(config_file):
        try:
            with open(config_file, 'r') as f:
                return json.load(f).get('last_directory', '')
        except:
            return ''
    return ''

def save_last_directory(directory):
    """Save the last used directory to a config file."""
    config_file = 'last_directory.json'
    try:
        with open(config_file, 'w') as f:
            json.dump({'last_directory': directory}, f)
    except:
        pass

def select_files():
    """Open file dialog and return selected files."""
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Get the last used directory
    initial_dir = load_last_directory()
    
    # Open file dialog
    files = filedialog.askopenfilenames(
        initialdir=initial_dir,
        title="Seleziona i file CSV",
        filetypes=(("CSV files", "*.csv"), ("All files", "*.*"))
    )
    
    # Save the directory if files were selected
    if files:
        save_last_directory(os.path.dirname(files[0]))
    
    return files

def extract_name_from_filename(filename):
    """Extract the name from a filename by removing common suffixes and numbers."""
    # Remove file extension
    name = os.path.splitext(os.path.basename(filename))[0]
    
    # Remove common suffixes like "PROVA", "SETT", and numbers
    words = name.split()
    filtered_words = []
    for word in words:
        if not word.isdigit() and word.upper() not in ['PROVA', 'SETT']:
            filtered_words.append(word)
    
    return ' '.join(filtered_words)

def find_matching_patient(patient_name, existing_patients):
    """Find a matching patient in the existing list using name similarity."""
    for existing_patient in existing_patients:
        if are_names_similar(patient_name, existing_patient):
            return existing_patient
    return None

def process_csv_file(csv_file, ws, patients, parameters):
    """Process a CSV file and update the Google Sheet."""
    try:
        with open(csv_file, 'r') as f:
            reader = csv.reader(f)
            data = list(reader)
            
        if not data:
            print(f"File {csv_file} è vuoto")
            return False
            
        # Estrai il nome del paziente dal nome del file
        patient_name = extract_name_from_filename(csv_file)
        
        # Trova il paziente corrispondente nel foglio
        matching_patient = find_matching_patient(patient_name, patients)
        
        if not matching_patient:
            # Chiedi all'utente se vuole aggiungere un nuovo paziente
            root = tk.Tk()
            root.withdraw()
            if messagebox.askyesno("Nuovo Paziente", 
                                 f"Nessun paziente corrispondente trovato per '{patient_name}'.\nVuoi aggiungerlo come nuovo paziente?"):
                # Aggiungi il nuovo paziente
                patients.append(patient_name)
                # Trova la prossima colonna disponibile
                col = 2
                while col-1 < len(ws.row_values(2)) and ws.row_values(2)[col-1].strip():
                    col += 3
                # Scrivi il nome del nuovo paziente
                ws.update_cell(2, col, patient_name)
                matching_patient = patient_name
            else:
                print(f"File {csv_file} ignorato")
                return False
        
        # Trova l'indice del paziente
        patient_index = patients.index(matching_patient)
        
        # Calcola la colonna di partenza per questo paziente
        start_col = 2 + patient_index * 3  # B=2, E=5, H=8, ...
        
        # Aggiorna i dati nel foglio
        for i, param in enumerate(parameters):
            if i < len(data):
                row_values = data[i]
                if len(row_values) >= 3:  # Assicurati che ci siano almeno 3 valori
                    # Aggiorna i valori nella riga corrispondente al parametro
                    ws.update(f"{start_col}:{start_col+2}", [[row_values[0], row_values[1], row_values[2]]])
        
        return True
        
    except Exception as e:
        print(f"Errore durante l'elaborazione del file {csv_file}: {e}")
        return False

def main():
    # 1) Autenticazione
    try:
        gc = gspread.service_account(filename='service_account.json')
    except Exception as e:
        print(f"Errore autenticazione: {e}")
        sys.exit(1)

    # 2) Apertura file
    try:
        sh = gc.open("TABELLA PER RANDOMIZZAZIONE DEI PAZIENTI")
    except SpreadsheetNotFound:
        print("File Google Sheets non trovato")
        sys.exit(1)

    # 3) Apertura foglio
    try:
        ws = sh.worksheet("ANALISI DATI 1")
    except WorksheetNotFound:
        print("Foglio 'ANALISI DATI 1' non trovato")
        sys.exit(1)

    # 4) Legge nomi pazienti in riga 2, da colonna B (indice 2), ogni 3 colonne
    all_row2 = ws.row_values(2)
    patients = []
    col = 2
    while col-1 < len(all_row2) and all_row2[col-1].strip():
        patients.append(all_row2[col-1].strip())
        col += 3

    if not patients:
        print("Nessun paziente trovato in riga 2")
        sys.exit(1)

    # 5) Legge i "Parameter" da colonna A, riga 4 in giù fino all'ultimo non vuoto
    all_colA = ws.col_values(1)[3:]  # celle dalla riga 4 in poi
    parameters = [p.strip() for p in all_colA if p.strip()]

    if not parameters:
        print("Nessun Parameter trovato in colonna A da riga 4")
        sys.exit(1)

    # 6) Seleziona i file CSV da processare
    csv_files = select_files()
    if not csv_files:
        print("Nessun file selezionato")
        sys.exit(0)

    # 7) Processa ogni file CSV
    success_count = 0
    for csv_file in csv_files:
        if process_csv_file(csv_file, ws, patients, parameters):
            success_count += 1

    print(f"Elaborazione completata. {success_count} file processati con successo.")

if __name__ == "__main__":
    main() 