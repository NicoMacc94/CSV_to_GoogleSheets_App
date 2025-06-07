#!/usr/bin/env python3

# script.py
import sys
import gspread
from gspread.exceptions import SpreadsheetNotFound, WorksheetNotFound

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

    # 6) Per ogni parameter genera una tabella a partire da colonna L (indice 12),
    #    lasciando 1 colonna vuota tra una tabella e l'altra (step = 4)
    start_col_base = 12  # L
    header_row = 18
    data_start_row = header_row + 1
    source_data_start_row = 4

    for i, param in enumerate(parameters):
        base = start_col_base + i * 4

        # 6.b) Scrive nome trattamento in (header_row, base)
        ws.update_cell(header_row, base, param)

        # 6.c) Scrive PROVA 1/2/3 in (header_row, base+1..+3)
        ws.update(header_row, base + 1, [['PROVA 1', 'PROVA 2', 'PROVA 3']])

        # 6.d) Per ogni paziente scrive nome e i 3 valori
        for j, paz in enumerate(patients):
            r = data_start_row + j

            # scrive nome paziente
            ws.update_cell(r, base, paz)

            # calcola colonna di partenza nella tabella di origine per questo paziente
            orig_base = 2 + j * 3  # B=2, E=5, H=8, ...

            # legge i 3 valori dalla riga source_data_start_row + i, colonne orig_base..orig_base+2
            source_row = source_data_start_row + i
            values = ws.row_values(source_row)[orig_base-1:orig_base+2]

            # scrive i 3 valori in (r, base+1..base+3)
            for k, v in enumerate(values):
                ws.update_cell(r, base + 1 + k, v)

        # la colonna base+4 rimane vuota (margine), poi riparte la tabella successiva

    print("Tutte le tabelle generate con successo.")

if __name__ == "__main__":
    main()
