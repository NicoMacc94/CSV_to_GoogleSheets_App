#!/usr/bin/env python3

import sys
import os
import gspread
from gspread.exceptions import SpreadsheetNotFound, WorksheetNotFound
import tkinter as tk
from tkinter import filedialog, messagebox
import json
from difflib import SequenceMatcher
import re
import PyPDF2

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
        title="Seleziona i file PDF",
        filetypes=(("PDF files", "*.pdf"), ("All files", "*.*"))
    )
    
    # Save the directory if files were selected
    if files:
        save_last_directory(os.path.dirname(files[0]))
    
    return files

def extract_name_and_prova(filename):
    """Extract the name and prova number from a filename."""
    # Remove file extension
    name = os.path.splitext(os.path.basename(filename))[0]
    
    # Extract prova number
    prova_match = re.search(r'PROVA\s+(\d+)', name, re.IGNORECASE)
    prova_number = int(prova_match.group(1)) if prova_match else None
    
    # Remove common suffixes and numbers
    words = name.split()
    filtered_words = []
    for word in words:
        if not word.isdigit() and word.upper() not in ['PROVA', 'SETT', 'GRUPPO']:
            filtered_words.append(word)
    
    return ' '.join(filtered_words), prova_number

def find_matching_patient(patient_name, existing_patients):
    """Find a matching patient in the existing list using name similarity."""
    for existing_patient in existing_patients:
        if are_names_similar(patient_name, existing_patient):
            return existing_patient
    return None

def extract_data_from_pdf(pdf_file):
    """Extract data from PDF file."""
    try:
        with open(pdf_file, 'rb') as file:
            reader = PyPDF2.PdfReader(file)
            text = ""
            for page in reader.pages:
                text += page.extract_text()
            
            # TODO: Implement the actual data extraction logic here
            # This is a placeholder - you need to implement the specific data extraction
            # based on your PDF format and requirements
            
            return []  # Return the extracted data in the required format
            
    except Exception as e:
        print(f"Errore durante l'estrazione dei dati dal PDF: {e}")
        return None

def process_pdf_file(pdf_file, ws, patients):
    """Process a PDF file and update the Google Sheet."""
    try:
        # Estrai il nome del paziente e il numero della prova dal nome del file
        patient_name, prova_number = extract_name_and_prova(pdf_file)
        
        if prova_number not in [4, 5]:
            print(f"File {pdf_file} ignorato: non è una PROVA 4 o 5")
            return False
            
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
                print(f"File {pdf_file} ignorato")
                return False
        
        # Trova l'indice del paziente
        patient_index = patients.index(matching_patient)
        
        # Calcola la colonna di partenza per questo paziente
        start_col = 2 + patient_index * 3  # B=2, E=5, H=8, ...
        
        # Determina la colonna per PROVA 4 o 5
        if prova_number == 4:
            target_col = 6  # Colonna F
        else:  # prova_number == 5
            target_col = 7  # Colonna G
        
        # Scrivi "PROVA X" nella riga 3
        ws.update_cell(3, target_col, f"PROVA {prova_number}")
        
        # Estrai i dati dal PDF
        data = extract_data_from_pdf(pdf_file)
        if data is None:
            return False
            
        # Aggiorna i dati nel foglio
        for i, row_data in enumerate(data):
            if len(row_data) >= 3:  # Assicurati che ci siano almeno 3 valori
                # Aggiorna i valori nella riga corrispondente
                ws.update(f"{target_col}:{target_col+2}", [[row_data[0], row_data[1], row_data[2]]])
        
        # Se è PROVA 5, aggiungi la nota in A21
        if prova_number == 5:
            ws.update_cell(21, 1, "PROVA 5")
        
        return True
        
    except Exception as e:
        print(f"Errore durante l'elaborazione del file {pdf_file}: {e}")
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
        ws = sh.worksheet("RIEPILOGO HRV")
    except WorksheetNotFound:
        print("Foglio 'RIEPILOGO HRV' non trovato")
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

    # 5) Seleziona i file PDF da processare
    pdf_files = select_files()
    if not pdf_files:
        print("Nessun file selezionato")
        sys.exit(0)

    # 6) Processa ogni file PDF
    success_count = 0
    for pdf_file in pdf_files:
        if process_pdf_file(pdf_file, ws, patients):
            success_count += 1

    print(f"Elaborazione completata. {success_count} file processati con successo.")

if __name__ == "__main__":
    main() 