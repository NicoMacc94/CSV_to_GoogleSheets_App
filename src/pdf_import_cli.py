#!/usr/bin/env python3
import os
import sys
import argparse

# Aggiungi la directory corrente al path per importare pdf_to_sheets
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from src import pdf_to_sheets

def main():
    """Versione a riga di comando dell'importazione PDF."""
    parser = argparse.ArgumentParser(description='Importa dati HRV da PDF a Google Sheets')
    parser.add_argument('pdf_path', nargs='?', help='Percorso del file PDF da elaborare')
    parser.add_argument('--interactive', '-i', action='store_true', help='Modalità interattiva')
    
    args = parser.parse_args()
    
    if args.interactive or not args.pdf_path:
        # Modalità interattiva
        print("=== IMPORTAZIONE DATI HRV DA PDF ===")
        print("Questa versione a riga di comando consente di importare")
        print("dati HRV da un PDF a un foglio Google Sheets esistente.")
        print("")
        
        # Chiedi il percorso del file PDF
        pdf_path = input("Inserisci il percorso completo del file PDF: ")
        if not pdf_path or not os.path.exists(pdf_path):
            print("Errore: File non trovato o percorso non valido.")
            return 1
    else:
        # Modalità diretta con percorso specificato da linea di comando
        pdf_path = args.pdf_path
        if not os.path.exists(pdf_path):
            print(f"Errore: File {pdf_path} non trovato.")
            return 1
    
    # Estrai informazioni dal nome del file
    basename = os.path.basename(pdf_path)
    name_parts = os.path.splitext(basename)[0].split()
    if len(name_parts) >= 2:
        nome = name_parts[0]
        cognome = name_parts[1]
        print(f"Nome e cognome rilevati: {nome} {cognome}")
    else:
        print("ATTENZIONE: Impossibile estrarre nome e cognome dal nome del file.")
    
    # Elabora il file PDF
    print(f"Elaborazione del file: {pdf_path}")
    try:
        success = pdf_to_sheets.process_pdf(pdf_path)
        
        if success:
            print("✅ Dati HRV importati con successo!")
            return 0
        else:
            print("❌ Impossibile importare i dati.")
            print("Verifica che esista un foglio Google Sheets per questa persona.")
            return 1
    
    except Exception as e:
        print(f"Errore durante l'elaborazione: {str(e)}")
        import traceback
        traceback.print_exc()
        return 1

if __name__ == "__main__":
    sys.exit(main()) 