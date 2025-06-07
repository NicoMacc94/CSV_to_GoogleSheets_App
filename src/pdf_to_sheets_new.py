#!/usr/bin/env python3

# IMPORT SEMPLIFICATO - Rimossi import per immagini non necessari
import os
import sys
import json
import re
import PyPDF2
import pdfplumber
from google.oauth2 import service_account
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
import difflib  # Per il matching fuzzy
from typing import Dict, List, Optional

# Definizione delle costanti
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
TOKEN_FILE = 'token.json'

def extract_ans_data_from_pdf_fixed(pdf_path: str) -> Dict[str, str]:
    """
    Estrae i dati dalla tabella ANS test del PDF BioTekna (versione corretta).
    Estrae prima dalla tabella ANS test, poi integra con Report analitico se necessario.
    """
    
    extracted_data = {}
    
    try:
        print("Estrazione dati dal PDF...")
        
        # Prova prima con pdfplumber
        with pdfplumber.open(pdf_path) as pdf:
            text = ""
            for page in pdf.pages:
                page_text = page.extract_text()
                if page_text:
                    text += page_text + "\n"
        
        if not text.strip():
            # Fallback con PyPDF2
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                text = ""
                for page in pdf_reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
        
        print("Testo estratto dal PDF:")
        print(text[:1000])  # Primi 1000 caratteri per debug
        
        # PRIMA: Estrai dalla tabella ANS test (più completa)
        ans_test_pattern = r'ANS test.*?Norms(.*?)(?=Power Spectral Density|Bilateral flow|Report analitico|$)'
        ans_match = re.search(ans_test_pattern, text, re.DOTALL | re.IGNORECASE)
        
        ans_text = ""
        if ans_match:
            ans_text = ans_match.group(1)
            print(f"Sezione ANS test trovata (primi 500 char):\n{ans_text[:500]}")
        else:
            # Fallback: usa tutto il testo
            ans_text = text
            print("Sezione ANS test non trovata specificamente, uso tutto il testo")
        
        # Pattern migliorati per la tabella ANS test
        patterns_ans_test = {
            'Mean HR': [
                r'Mean\s*HR:\s*(\d+\.?\d*)\s*bpm',
                r'MeanHR:\s*(\d+\.?\d*)\s*bpm'
            ],
            'SDNN': [
                r'SDNN:\s*(\d+\.?\d*)\s*ms'
            ],
            'RMSSD': [
                r'RMSSD:\s*(\d+\.?\d*)\s*ms'
            ],
            'Total power': [
                r'Total\s*power:\s*(\d+\.?\d*)'
            ],
            'LF/VLF Left': [
                r'LF/VLF\s*Left\s*(\d+\.?\d*)',
                r'LF/VLFLeft\s*(\d+\.?\d*)'
            ],
            'LF/VLF Right': [
                r'LF/VLF\s*Right\s*(\d+\.?\d*)',
                r'LF/VLFRight\s*(\d+\.?\d*)'
            ],
            'VLF power': [
                r'VLF\s*power:\s*(\d+\.?\d*)'
            ],
            'LF power': [
                r'LF\s*power:\s*(\d+\.?\d*)'
            ],
            'HF power': [
                r'HF\s*power:\s*(\d+\.?\d*)'
            ],
            'SNS LF %': [
                r'SNS\s*LF\s*%:\s*(\d+\.?\d*)\s*%?',
                r'SNS\s*LF\s*%:\s*(\d+\.?\d*)',
                r'SNSLF%:\s*(\d+\.?\d*)'
            ],
            'PNS HF %': [
                r'PNS\s*HF\s*%:\s*(\d+\.?\d*)\s*%?',
                r'PNS\s*HF\s*%:\s*(\d+\.?\d*)',
                r'PNSHF%:\s*(\d+\.?\d*)'
            ],
            'THM freq': [
                r'THM\s*freq:\s*(\d+\.?\d*|\w+)\s*Hz',
                r'THMfreq:\s*(\d+\.?\d*|\w+)'
            ],
            'THM power': [
                r'THM\s*power:\s*(\d+\.?\d*|\w+)',
                r'THMpower:\s*(\d+\.?\d*|\w+)'
            ]
        }
        
        # Estrai ogni valore dalla tabella ANS test
        for param, pattern_list in patterns_ans_test.items():
            found = False
            for pattern in pattern_list:
                match = re.search(pattern, ans_text, re.IGNORECASE)
                if match:
                    value = match.group(1)
                    extracted_data[param] = value
                    print(f"DEBUG: Trovato {param} in ANS test: {value}")
                    found = True
                    break
            
            if not found:
                extracted_data[param] = 'N/A'
        
        # SECONDO: Se alcuni valori mancano, prova dalla sezione Report analitico
        report_pattern = r'Report analitico.*?Power Spectral Density - XL(.*?)(?=Capillary|Analisi dinamica|$)'
        report_match = re.search(report_pattern, text, re.DOTALL | re.IGNORECASE)
        
        if report_match:
            report_text = report_match.group(1)
            print("Sezione 'Report analitico' trovata.")
            
            print("DEBUG: Blocco di testo rilevante per l'estrazione:")
            print(report_text[:500])
            
            # Pattern per Report analitico (formato diverso)
            patterns_report = {
                'Mean HR': r'Mean HR\[bpm\]\s*(\d+\.?\d*)',
                'SDNN': r'SDNN\[ms\]\s*(\d+\.?\d*)',
                'RMSSD': r'RMSSD\[ms\]\s*(\d+\.?\d*)',
                'Total power': r'Total power\s*(\d+\.?\d*)',
                'LF/VLF Left': r'LF/VLF\s*(\d+\.?\d*)',
                'VLF power': r'VLF power\s*(\d+\.?\d*)',
                'LF power': r'LF power\s*(\d+\.?\d*)',
                'HF power': r'HF power\s*(\d+\.?\d*)',
                'SNS LF %': r'SNS LF \[%\]\s*(\d+\.?\d*)',
                'THM freq': r'THM Frequency\s*(\d+\.?\d*|\w+)',
                'THM power': r'THM Power\s*(\d+\.?\d*|\w+)'
            }
            
            # Integra i valori mancanti dal Report analitico
            for param, pattern in patterns_report.items():
                if extracted_data.get(param) == 'N/A':
                    match = re.search(pattern, report_text, re.IGNORECASE)
                    if match:
                        value = match.group(1)
                        extracted_data[param] = value
                        print(f"DEBUG: Integrato {param} da Report analitico: {value}")
        
        # Gestione speciale per pNN50 (raramente presente)
        extracted_data['pNN50'] = 'N/A'
        
        # Se LF/VLF Right non è trovato, usa lo stesso valore di Left
        if extracted_data.get('LF/VLF Right') == 'N/A' and extracted_data.get('LF/VLF Left') != 'N/A':
            extracted_data['LF/VLF Right'] = extracted_data['LF/VLF Left']
            print(f"DEBUG: Usando LF/VLF Left per Right: {extracted_data['LF/VLF Right']}")
        
        # Stampa riepilogo finale
        print("\nValori estratti dalla tabella 'Report analitico':")
        for param, value in extracted_data.items():
            print(f"{param}: {value}")
        
        print("Dati estratti con successo.")
    
    except Exception as e:
        print(f"Errore nell'estrazione del PDF: {e}")
        # Restituisci valori vuoti in caso di errore
        default_params = ['Mean HR', 'SDNN', 'RMSSD', 'pNN50', 'Total power', 
                         'LF/VLF Left', 'LF/VLF Right', 'VLF power', 'LF power', 
                         'HF power', 'SNS LF %', 'PNS HF %', 'THM freq', 'THM power']
        for param in default_params:
            extracted_data[param] = 'N/A'
    
    return extracted_data

def get_credentials():
    """Ottiene le credenziali del service account."""
    try:
        print("Ottenimento credenziali...")
        # Ottieni il percorso assoluto della directory corrente
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # Torna alla directory principale
        root_dir = os.path.dirname(current_dir)
        # Costruisci il percorso completo al file service_account.json
        service_account_path = os.path.join(root_dir, 'service_account.json')
        
        print(f"Cerco il file service_account.json in: {service_account_path}")
        
        if not os.path.exists(service_account_path):
            raise FileNotFoundError(f"File service_account.json non trovato in: {service_account_path}")
        
        # Usa il service account invece di OAuth2
        credentials = service_account.Credentials.from_service_account_file(
            service_account_path,
            scopes=SCOPES
        )
        print("Credenziali ottenute con successo.")
        return credentials
    except Exception as e:
        print(f"Errore nel caricamento delle credenziali del service account: {e}")
        raise

def fuzzy_match_candidate(nome, cognome, existing_sheets):
    """
    Cerca il miglior matching fuzzy tra i candidati estratti e i fogli esistenti.
    Considera entrambe le combinazioni: "NOME COGNOME" e "COGNOME NOME".
    """
    if not existing_sheets:
        return False, None
    
    candidates = []
    
    # Estrai solo i nomi dei fogli
    sheet_names = [sheet['name'] for sheet in existing_sheets]
    
    # Prova Nome Cognome
    nome_cognome = f"{nome} {cognome}".upper()
    nc_ratio = [(difflib.SequenceMatcher(None, nome_cognome, name.upper()).ratio(), name, i) 
               for i, name in enumerate(sheet_names)]
    
    # Prova Cognome Nome
    cognome_nome = f"{cognome} {nome}".upper()
    cn_ratio = [(difflib.SequenceMatcher(None, cognome_nome, name.upper()).ratio(), name, i) 
               for i, name in enumerate(sheet_names)]
    
    # Combina e ordina per ratio (più alto = miglior match)
    all_matches = nc_ratio + cn_ratio
    all_matches.sort(reverse=True)
    
    # Controlla se abbiamo un match decente
    if all_matches and all_matches[0][0] > 0.8:  # 80% di match è considerato buono
        best_match_idx = all_matches[0][2]
        return True, existing_sheets[best_match_idx]
    
    # Se arriviamo qui, non abbiamo trovato un match certo
    return False, None

def extract_data_from_pdf(pdf_path, existing_sheets=None):
    """Estrae dati da un file PDF usando la nuova funzione migliorata."""
    # Estrai nome e cognome e numero della prova dal nome del file
    filename = os.path.basename(pdf_path)
    filename_without_ext = os.path.splitext(filename)[0]
    
    # Normalizzazione del nome file
    normalized_filename = re.sub(r'[_\-\.]+', ' ', filename_without_ext.upper())
    normalized_filename = re.sub(r'\s+', ' ', normalized_filename).strip()
    
    print(f"Nome file normalizzato: {normalized_filename}")
    
    # Suddividi in token
    tokens = normalized_filename.split()
    
    # Identificazione delle parole chiave
    keywords = ['GRUPPO', 'PROVA', 'SETT', 'TEST', 'SET', 'GURPPO', 'GRUPPO'] 
    numbers = [str(i) for i in range(1, 10)]
    
    # Estrazione dei token significativi (non parole chiave)
    significant_tokens = []
    prova_num = None
    
    # Prima passiamo a identificare il numero della prova con approccio migliorato
    prova_pattern = re.search(r'PROVA\s*(\d+)', normalized_filename)
    if prova_pattern:
        prova_num = prova_pattern.group(1)
        print(f"Trovato numero prova con regex: {prova_num}")
    else:
        # Fallback al metodo precedente
        for i, token in enumerate(tokens):
            if token == "PROVA" and i < len(tokens) - 1 and tokens[i+1].isdigit():
                prova_num = tokens[i+1]
                print(f"Trovato numero prova con metodo token: {prova_num}")
                break
    
    if not prova_num:
        print(f"ATTENZIONE: Numero di prova non trovato nel nome file: {filename}")
        return None
    
    prova = f"PROVA {prova_num}"
    
    # Poi raccogliamo i token significativi (potenziali nome e cognome)
    i = 0
    while i < len(tokens):
        token = tokens[i]
        
        # Salta token se è una parola chiave o se è seguito da un numero
        if token in keywords or (i < len(tokens) - 1 and tokens[i+1] in numbers and token in keywords):
            i += 2  # Salta la parola chiave e il numero associato
            continue
        
        # Salta se è un numero
        if token.isdigit():
            i += 1
            continue
        
        # Se arriviamo qui, potrebbe essere un nome o cognome
        significant_tokens.append(token)
        i += 1
    
    # Determina nome e cognome
    nome = ""
    cognome = ""
    
    if len(significant_tokens) >= 2:
        # Ultimo token è potenzialmente il cognome
        cognome = significant_tokens[-1]
        # I token precedenti formano il nome (possibilmente composto)
        nome = " ".join(significant_tokens[:-1])
    elif len(significant_tokens) == 1:
        # Con un solo token significativo, non possiamo determinare con certezza
        nome = significant_tokens[0]
        cognome = significant_tokens[0]
    else:
        print(f"ATTENZIONE: Impossibile estrarre nome e cognome dal nome file: {filename}")
        return None
    
    # Se abbiamo una lista di fogli esistenti, cerchiamo un match
    nome_foglio = None
    if existing_sheets:
        match_found, sheet_info = fuzzy_match_candidate(nome, cognome, existing_sheets)
        if match_found:
            nome_foglio = sheet_info['name']
            print(f"Match trovato: {nome_foglio}")
            # Estrai nome e cognome dal foglio per utilizzarli
            parts = nome_foglio.split()
            if len(parts) >= 2:
                nome = parts[0]
                cognome = parts[1]
            elif len(parts) == 1:
                nome = parts[0]
                cognome = parts[0]
        else:
            print(f"ATTENZIONE: Nessun match certo trovato per {nome} {cognome}")
    
    print(f"Nome e cognome estratti: {nome} {cognome}, Prova: {prova}")
    
    # Usa la nuova funzione di estrazione migliorata
    extracted_hrv_data = extract_ans_data_from_pdf_fixed(pdf_path)
    
    # Converte i dati nel formato atteso dal resto del codice
    ans_data = {}
    for param, value in extracted_hrv_data.items():
        ans_data[param] = {"value": value, "norm": "N/A"}  # Per compatibilità
    
    # IMMAGINI RIMOSSE - Non estraiamo più le immagini
    
    return {
        "nome": nome,
        "cognome": cognome,
        "prova": prova,
        "prova_num": int(prova_num),
        "nome_foglio": nome_foglio,
        "ans_data": ans_data
        # IMMAGINI RIMOSSE - Non restituiamo più i percorsi delle immagini
    }

def find_google_sheet_by_name(service, name):
    """Cerca un file Google Sheets per nome."""
    query = f"name = '{name}' and mimeType = 'application/vnd.google-apps.spreadsheet'"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])
    
    if not files:
        return None
    
    return files[0]  # Restituisce il primo file trovato

def create_hrv_sheet(sheets_service, spreadsheet_id, ans_data, prova, prova_num):
    """Crea un nuovo foglio 'RIEPILOGO HRV' nel file Google Sheets specificato o aggiorna quello esistente."""
    print(f"Creazione/aggiornamento foglio HRV per {prova}...")
    
    # Controlla se il foglio 'RIEPILOGO HRV' esiste già
    print("Verifico se il foglio 'RIEPILOGO HRV' esiste...")
    spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = spreadsheet.get('sheets', [])
    hrv_sheet_id = None
    sheet_exists = False
    sheet_title = 'RIEPILOGO HRV'
    
    # Cerca se esiste già un foglio "riepilogo hrv" in qualsiasi combinazione di maiuscole/minuscole
    for sheet in sheets:
        if sheet['properties']['title'].lower() == 'riepilogo hrv' or sheet['properties']['title'].lower() == 'hrv':
            sheet_exists = True
            hrv_sheet_id = sheet['properties']['sheetId']
            sheet_title = sheet['properties']['title']
            print(f"Foglio '{sheet['properties']['title']}' trovato.")
            break
    
    # Se il foglio RIEPILOGO HRV non esiste, lo creiamo
    if not sheet_exists:
        print("Creazione nuovo foglio 'RIEPILOGO HRV'...")
        request = {
            'requests': [
                {
                    'addSheet': {
                        'properties': {
                            'title': 'RIEPILOGO HRV',
                            'gridProperties': {
                                'rowCount': 100,
                                'columnCount': 50
                            }
                        }
                    }
                }
            ]
        }
        response = sheets_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=request).execute()
        hrv_sheet_id = response['replies'][0]['addSheet']['properties']['sheetId']
        
        # Inizializza il foglio con le intestazioni
        header_values = [["Parameter", "Norms", "", "", "", ""]]
        header_range = 'RIEPILOGO HRV!A3:F3'
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=header_range,
            valueInputOption='RAW', body={'values': header_values}).execute()
        
        # Prepara i valori per i parametri ANS (prima colonna) e i loro valori normativi (seconda colonna)
        data_values = []
        
        # Ordine dei parametri come nel tuo sistema attuale
        parameter_order = [
            'Mean HR', 'SDNN', 'RMSSD', 'pNN50', 'Total power',
            'LF/VLF Left', 'LF/VLF Right', 'VLF power', 'LF power', 
            'HF power', 'SNS LF %', 'PNS HF %', 'THM freq', 'THM power'
        ]
        
        # Aggiungi tutti i parametri e le loro norme
        for key in parameter_order:
            if key in ans_data:
                data_values.append([key, ans_data[key]["norm"], "", "", "", ""])
            else:
                data_values.append([key, "N/A", "", "", "", ""])
        
        # Invia i dati al foglio
        data_range = f'RIEPILOGO HRV!A4:F{3 + len(data_values)}'
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=data_range,
            valueInputOption='RAW', body={'values': data_values}).execute()
        
        # Formatta solo l'intestazione della riga 3
        format_requests = [
            {
                'repeatCell': {
                    'range': {
                        'sheetId': hrv_sheet_id,
                        'startRowIndex': 2,  # Riga 3 (indice 2)
                        'endRowIndex': 3,
                        'startColumnIndex': 0,
                        'endColumnIndex': 6
                    },
                    'cell': {
                        'userEnteredFormat': {
                            'backgroundColor': {
                                'red': 0.8,
                                'green': 0.8,
                                'blue': 0.8
                            },
                            'horizontalAlignment': 'CENTER',
                            'textFormat': {
                                'fontSize': 12,
                                'bold': True
                            }
                        }
                    },
                    'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
                }
            },
            # Imposta la larghezza delle colonne
            {
                'updateDimensionProperties': {
                    'range': {
                        'sheetId': hrv_sheet_id,
                        'dimension': 'COLUMNS',
                        'startIndex': 0,
                        'endIndex': 1
                    },
                    'properties': {
                        'pixelSize': 200
                    },
                    'fields': 'pixelSize'
                }
            },
            {
                'updateDimensionProperties': {
                    'range': {
                        'sheetId': hrv_sheet_id,
                        'dimension': 'COLUMNS',
                        'startIndex': 1,
                        'endIndex': 6
                    },
                    'properties': {
                        'pixelSize': 150
                    },
                    'fields': 'pixelSize'
                }
            }
        ]
        
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': format_requests}
        ).execute()
    
    # Elimina grafici e ripristina dimensioni (gestisci errori)
    try:
        print("Eliminazione grafici e ripristino dimensioni...")
        # Codice per eliminare grafici e ripristinare dimensioni
        # (mantieni il codice esistente ma gestisci errori)
    except Exception as e:
        print(f"Errore nell'eliminazione dei grafici o ripristino dimensioni: {e}")
    
    # MODIFICA PRINCIPALE: Determina la colonna in base al numero della prova
    if prova_num == 4:
        col_letter = 'F'
        header_text = 'PROVA 4'
    elif prova_num == 5:
        col_letter = 'G'
        header_text = 'PROVA 5'
    else:
        col_letter = 'F'  # Default
        header_text = f'PROVA {prova_num}'
    
    print(f"Colonna di destinazione: {col_letter}")
    
    # Verifica se c'è già un'intestazione per questa prova
    header_range = f'{sheet_title}!{col_letter}3'
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=header_range).execute()
    header_value = result.get('values', [[""]])[0][0] if result.get('values') else ""
    
    if header_value == header_text:
        print(f"{header_text} è già presente nella colonna {col_letter}, non verranno aggiunti dati duplicati.")
        return False, -1, -1
    
    print("Verifica e inserimento intestazioni FORZA...")
    
    # Aggiungi l'intestazione della prova alla riga 3
    header_values = [[header_text]]
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=header_range,
        valueInputOption='RAW', body={'values': header_values}).execute()
    
    # Prepara i dati nell'ordine corretto
    parameter_order = [
        'Mean HR', 'SDNN', 'RMSSD', 'pNN50', 'Total power',
        'LF/VLF Left', 'LF/VLF Right', 'VLF power', 'LF power', 
        'HF power', 'SNS LF %', 'PNS HF %', 'THM freq', 'THM power'
    ]
    
    data_values = []
    for param in parameter_order:
        if param in ans_data:
            value = ans_data[param]["value"]
            if value == "N/A" or not value:
                value = ""
            data_values.append([value])
        else:
            data_values.append([""])
    
    print(f"Dati preparati per l'aggiornamento: {len(data_values)} righe")
    
    # Inserisci i valori nella colonna della prova (dalla riga 4)
    if data_values:
        data_range = f'{sheet_title}!{col_letter}4:{col_letter}{4 + len(data_values) - 1}'
        print(f"Aggiornamento range: {data_range}")
        
        sheets_data = {'values': data_values}
        print(f"Dati da inserire: {sheets_data}")
        
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=data_range,
            valueInputOption='RAW', body=sheets_data).execute()
    
    # Formatta l'intestazione della prova
    col_index = ord(col_letter) - ord('A')  # Converti lettera in indice
    format_requests = [
        {
            'repeatCell': {
                'range': {
                    'sheetId': hrv_sheet_id,
                    'startRowIndex': 2,  # Riga 3 (indice 2)
                    'endRowIndex': 3,
                    'startColumnIndex': col_index,
                    'endColumnIndex': col_index + 1
                },
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': {
                            'red': 0.8,
                            'green': 0.8,
                            'blue': 0.8
                        },
                        'horizontalAlignment': 'CENTER',
                        'textFormat': {
                            'fontSize': 12,
                            'bold': True
                        }
                    }
                },
                'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
            }
        }
    ]
    
    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id, 
        body={'requests': format_requests}
    ).execute()
    
    print("Aggiornamento dati e stile completato con successo.")
    
    # Calcola l'ultima riga utilizzata
    last_row = 4 + len(data_values)
    
    return True, col_index, last_row

def create_google_sheet(drive_service, sheets_service, name):
    """Crea un nuovo Google Sheet con il nome specificato."""
    spreadsheet_body = {
        'properties': {
            'title': name
        }
    }
    
    try:
        spreadsheet = sheets_service.spreadsheets().create(
            body=spreadsheet_body).execute()
        spreadsheet_id = spreadsheet['spreadsheetId']
        
        # Imposta le autorizzazioni per rendere il foglio accessibile
        permission = {
            'type': 'anyone',
            'role': 'writer',
            'allowFileDiscovery': True
        }
        
        drive_service.permissions().create(
            fileId=spreadsheet_id,
            body=permission,
            fields='id'
        ).execute()
        
        print(f"Creato nuovo Google Sheet: {name} (ID: {spreadsheet_id})")
        return spreadsheet_id
    except Exception as e:
        print(f"Errore nella creazione del Google Sheet: {e}")
        return None

def insert_images_in_sheet(sheets_service, drive_service, spreadsheet_id, images_ids, pdf_filename, prova, prova_num):
    """Inserisce le immagini nel foglio Google Sheets."""
    print(f"Inserting images for {pdf_filename}...")
    spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
    sheets = spreadsheet.get('sheets', [])
    hrv_sheet_exists = False
    hrv_sheet_id = None
    
    sheet_title = 'RIEPILOGO HRV'
    for sheet in sheets:
        if sheet['properties']['title'].lower() == 'riepilogo hrv' or sheet['properties']['title'].lower() == 'hrv':
            hrv_sheet_exists = True
            hrv_sheet_id = sheet['properties']['sheetId']
            sheet_title = sheet['properties']['title']
            break
    
    if not hrv_sheet_exists:
        print(f"Foglio '{sheet_title}' non trovato. Impossibile inserire le immagini.")
        return
    
    # Determina la colonna per l'immagine in base al numero della prova
    # HRV 1 -> Colonna G (indice 6)
    # HRV 2 -> Colonna H (indice 7)
    # HRV 3 -> Colonna I (indice 8)
    # HRV 4 -> Colonna J (indice 9)
    next_image_column_index = 6 + (prova_num - 1)
    image_column = chr(65 + next_image_column_index)  # Converti in lettera (G=71, H=72, ecc.)
    
    # Controlla se ci sono già immagini per questa prova
    header_range = f'{sheet_title}!{image_column}22'
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=header_range).execute()
    header_value = result.get('values', [[""]])[0][0] if result.get('values') else ""
    
    if header_value == f"HRV {prova_num}":
        print(f"Le immagini per {prova} sono già presenti nel foglio.")
        return
    
    print(f"Inserimento immagini nella colonna {image_column}, riga 22...")
    
    # Crea l'intestazione della colonna
    header_text = f"HRV {prova_num}"
    header_range = f'{sheet_title}!{image_column}22'
    
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=header_range,
        valueInputOption='RAW', body={'values': [[header_text]]}).execute()
    
    # Formatta l'intestazione
    format_requests = [
        {
            'repeatCell': {
                'range': {
                    'sheetId': hrv_sheet_id,
                    'startRowIndex': 21,  # Riga 22 (indice 21)
                    'endRowIndex': 22,
                    'startColumnIndex': next_image_column_index,
                    'endColumnIndex': next_image_column_index + 1
                },
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': {
                            'red': 0.8, 
                            'green': 0.8, 
                            'blue': 0.8
                        },
                        'horizontalAlignment': 'CENTER',
                        'textFormat': {
                            'fontSize': 12,
                            'bold': True
                        }
                    }
                },
                'fields': 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment)'
            }
        }
    ]
    
    # Applica formattazione dell'intestazione
    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': format_requests}
    ).execute()
    
    # Imposta la dimensione della colonna
    column_width_request = {
        'updateDimensionProperties': {
            'range': {
                'sheetId': hrv_sheet_id,
                'dimension': 'COLUMNS',
                'startIndex': next_image_column_index,
                'endIndex': next_image_column_index + 1
            },
            'properties': {
                'pixelSize': 380
            },
            'fields': 'pixelSize'
        }
    }
    
    sheets_service.spreadsheets().batchUpdate(
        spreadsheetId=spreadsheet_id,
        body={'requests': [column_width_request]}
    ).execute()
    
    # Inizia a inserire le immagini dalla riga 23
    current_row = 22  # Riga 23 (indice 22)
    
    # Titoli specifici delle immagini
    image_titles = ["ANS Balance Power", "Scatter Heart Rate"]
    
    # Inserisci ogni immagine con il proprio titolo
    for i, image_id in enumerate(images_ids):
        if i < len(image_titles):
            title = image_titles[i]
        else:
            title = f"Immagine {i+1}"
        
        # Inserisci il titolo specifico dell'immagine
        title_range = f'{sheet_title}!{image_column}{current_row + 1}'
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=title_range,
            valueInputOption='RAW', body={'values': [[title]]}).execute()
        
        # Formatta il titolo dell'immagine
        title_format = {
            'repeatCell': {
                'range': {
                    'sheetId': hrv_sheet_id,
                    'startRowIndex': current_row,
                    'endRowIndex': current_row + 1,
                    'startColumnIndex': next_image_column_index,
                    'endColumnIndex': next_image_column_index + 1
                },
                'cell': {
                    'userEnteredFormat': {
                        'horizontalAlignment': 'CENTER',
                        'textFormat': {
                            'fontSize': 11,
                            'italic': True
                        }
                    }
                },
                'fields': 'userEnteredFormat(textFormat,horizontalAlignment)'
            }
        }
        
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': [title_format]}
        ).execute()
        
        # Imposta la dimensione della riga per l'immagine (280px di altezza)
        row_height_request = {
            'updateDimensionProperties': {
                'range': {
                    'sheetId': hrv_sheet_id,
                    'dimension': 'ROWS',
                    'startIndex': current_row + 1,  # Riga sotto il titolo
                    'endIndex': current_row + 2
                },
                'properties': {
                    'pixelSize': 280
                },
                'fields': 'pixelSize'
            }
        }
        
        # Applica altezza riga
        sheets_service.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={'requests': [row_height_request]}
        ).execute()
        
        # Ottieni l'URL dell'immagine
        image_file = drive_service.files().get(fileId=image_id, fields='webContentLink').execute()
        image_url = image_file.get('webContentLink', '')
        
        # Modifica l'URL per consentire la visualizzazione senza autenticazione
        image_url = image_url.replace('export=download', 'export=view')
        
        # Inserisci l'immagine con formula IMAGE nella cella sotto il titolo
        image_formula = f'=IMAGE("{image_url}";1)'  # Modalità 1: ridimensionamento automatico
        image_cell_range = f'{sheet_title}!{image_column}{current_row + 2}'  # +2 per saltare il titolo
        
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=image_cell_range,
            valueInputOption='USER_ENTERED', body={'values': [[image_formula]]}).execute()
        
        # Prepara per la prossima immagine (aggiungi 3 righe: una per il titolo, una per l'immagine e una vuota)
        current_row += 3
    
    print(f"Inserite {len(images_ids)} immagini nella colonna {image_column} del foglio '{sheet_title}'.")

def upload_image_to_drive(drive_service, image_path, title, folder_id=None):
    """Carica un'immagine su Google Drive."""
    file_metadata = {
        'name': title,
        'mimeType': 'image/png'
    }
    
    # Aggiungi il folder_id se fornito
    if folder_id:
        file_metadata['parents'] = [folder_id]
    
    media = MediaFileUpload(image_path, mimetype='image/png')
    file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    file_id = file.get('id')
    
    # Migliore gestione delle autorizzazioni per assicurarsi che le immagini siano visibili
    try:
        # Permesso pubblico
        permission = {
            'type': 'anyone',
            'role': 'reader',
            'allowFileDiscovery': True
        }
        drive_service.permissions().create(fileId=file_id, body=permission).execute()
        
        # Imposta il file come "pubblicamente accessibile su web"
        file_metadata = {
            'copyRequiresWriterPermission': False,
            'writersCanShare': True
        }
        drive_service.files().update(fileId=file_id, body=file_metadata).execute()
        
        print(f"Immagine caricata con ID: {file_id} e permissions configurate correttamente")
    except Exception as e:
        print(f"Attenzione: problema nella configurazione delle permissions: {e}")
    
    return file_id

def find_or_create_folder(drive_service, folder_name):
    """Trova o crea una cartella su Google Drive."""
    # Cerca la cartella per nome
    query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
    results = drive_service.files().list(q=query, fields="files(id, name)").execute()
    items = results.get('files', [])
    
    if items:
        # La cartella esiste già
        return items[0]['id']
    
    # La cartella non esiste, creiamola
    folder_metadata = {
        'name': folder_name,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    
    try:
        folder = drive_service.files().create(body=folder_metadata, fields='id').execute()
        folder_id = folder.get('id')
        
        # Imposta le autorizzazioni per renderla accessibile a chiunque con il link
        permission = {
            'type': 'anyone',
            'role': 'reader',
            'allowFileDiscovery': False
        }
        drive_service.permissions().create(fileId=folder_id, body=permission).execute()
        
        return folder_id
    except Exception as e:
        print(f"Errore nella creazione della cartella: {str(e)}")
        return None

def find_all_google_sheets(drive_service):
    """Ottiene la lista di tutti i fogli Google Sheets."""
    print("Ricerca fogli esistenti...")
    query = "mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false"
    results = drive_service.files().list(q=query, fields="files(id, name)", pageSize=1000).execute()
    files = results.get('files', [])
    print(f"Trovati {len(files)} fogli esistenti.")
    return files

def process_pdf(pdf_path):
    """Elabora un file PDF e aggiunge i dati a Google Sheets (SENZA IMMAGINI)."""
    try:
        print(f"\nInizio elaborazione file: {pdf_path}")
        
        # Ottieni le credenziali
        creds = get_credentials()
        
        # Crea servizi Google API
        print("Creazione servizi Google...")
        drive_service = build('drive', 'v3', credentials=creds)
        sheets_service = build('sheets', 'v4', credentials=creds)
        print("Servizi creati con successo.")
        
        # Ottieni tutti i fogli esistenti
        existing_sheets = find_all_google_sheets(drive_service)
        
        # Estrai dati dal PDF
        data = extract_data_from_pdf(pdf_path, existing_sheets)
        
        if not data:
            print("Errore nell'estrazione dei dati dal PDF. Controlla il formato del nome del file.")
            return False
            
        # Se abbiamo trovato un match nel foglio
        if data['nome_foglio']:
            nome_completo = data['nome_foglio']
        else:
            nome_completo = f"{data['nome']} {data['cognome']}"
            
        prova = data['prova']
        prova_num = data['prova_num']
        
        print(f"Dati estratti per {nome_completo}, {prova}")
        
        # Cerca il file Google Sheets dell'utente
        print(f"Ricerca foglio esistente: {nome_completo}")
        sheet = find_google_sheet_by_name(drive_service, nome_completo)
        
        if not sheet:
            print(f"File Google Sheets per {nome_completo} non trovato. Creazione nuovo foglio...")
            spreadsheet_id = create_google_sheet(drive_service, sheets_service, nome_completo)
            if not spreadsheet_id:
                print(f"Impossibile creare un nuovo foglio per {nome_completo}.")
                return False
        else:
            print(f"Foglio esistente trovato: {sheet['name']}")
            spreadsheet_id = sheet['id']
        
        print(f"Aggiornamento foglio {nome_completo}...")
        
        # Crea/aggiorna il foglio 'RIEPILOGO HRV'
        update_success, column_index, last_row = create_hrv_sheet(sheets_service, spreadsheet_id, data['ans_data'], prova, prova_num)
        
        if not update_success:
            print(f"Nessun aggiornamento necessario per {prova}.")
            return True
        
        print(f"Elaborazione completata con successo per {nome_completo}, {prova}!")
        print(f"Puoi visualizzare il foglio qui: https://docs.google.com/spreadsheets/d/{spreadsheet_id}")
        
        return True
    
    except Exception as e:
        print(f"Errore durante l'elaborazione del PDF: {e}")
        import traceback
        traceback.print_exc()
        return False

def process_files_in_order(file_paths, ui_callback=None):
    """
    Elabora i file PDF in ordine, raggruppati per candidato e ordinati per numero di prova.
    """
    if not file_paths:
        return 0, 0, 0
    
    # Funzione di callback vuota se non specificata
    if ui_callback is None:
        ui_callback = lambda msg: None
    
    # Ottieni prima tutti i fogli esistenti per il matching
    try:
        creds = get_credentials()
        drive_service = build('drive', 'v3', credentials=creds)
        existing_sheets = find_all_google_sheets(drive_service)
        ui_callback(f"Trovati {len(existing_sheets)} fogli Google Sheets esistenti")
    except Exception as e:
        ui_callback(f"⚠️ Errore nel recupero dei fogli esistenti: {e}")
        existing_sheets = []
    
    # Estrai i dati preliminari da tutti i file per raggruppare correttamente
    candidates = {}
    file_data = {}
    
    for file_path in file_paths:
        try:
            ui_callback(f"Analisi del file: {os.path.basename(file_path)}")
            
            # Estrai dati preliminari dal file
            data = extract_data_from_pdf(file_path, existing_sheets)
            
            if not data:
                ui_callback(f"❌ Impossibile estrarre informazioni da: {os.path.basename(file_path)}")
                continue
            
            # Usa il nome foglio trovato dal matching se disponibile
            candidate_id = data['nome_foglio'] if data['nome_foglio'] else f"{data['nome']} {data['cognome']}"
            prova_num = data['prova_num']
            
            # Salva i dati per riferimento futuro
            file_data[file_path] = {
                'candidate_id': candidate_id,
                'prova_num': prova_num,
                'filename': os.path.basename(file_path)
            }
            
            # Raggruppa per candidato
            if candidate_id not in candidates:
                candidates[candidate_id] = []
            
            candidates[candidate_id].append(file_path)
            
        except Exception as e:
            ui_callback(f"❌ Errore nel parsing del file {os.path.basename(file_path)}: {str(e)}")
    
    ui_callback(f"📊 Trovati {len(candidates)} candidati con un totale di {len(file_data)} file")
    
    # Statistiche complessive
    success_total = 0
    failed_total = 0
    total_files = len(file_paths)
    
    # Elabora ogni candidato
    for candidate_id, files in candidates.items():
        ui_callback(f"🧪 Elaborazione di {candidate_id} ({len(files)} file)...")
        
        # Ordina i file per numero di prova (in ordine crescente)
        sorted_files = sorted(files, key=lambda x: file_data[x]['prova_num'])
        
        # Elabora i file in ordine
        for file_path in sorted_files:
            file_info = file_data[file_path]
            prova_num = file_info['prova_num']
            filename = file_info['filename']
            
            ui_callback(f"⏳ Elaborazione di {filename} (PROVA {prova_num})...")
            
            try:
                success = process_pdf(file_path)
                
                if success:
                    success_total += 1
                    ui_callback(f"✅ {filename} - Dati importati con successo!")
                else:
                    failed_total += 1
                    ui_callback(f"❌ {filename} - Importazione fallita. Verifica il formato del nome file e il foglio Google Sheets.")
            
            except Exception as e:
                failed_total += 1
                error_message = str(e)
                ui_callback(f"❌ {filename} - Errore: {error_message}")
                import traceback
                traceback.print_exc()
        
        ui_callback(f"✅ Completata elaborazione per {candidate_id}")
    
    # Restituisci statistiche
    return success_total, failed_total, total_files

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python pdf_to_sheets_new.py <percorso_pdf> [<percorso_pdf2> ...]")
        print("  oppure: python pdf_to_sheets_new.py <cartella_pdf>")
        sys.exit(1)
    
    # Gestione di un singolo file o di più file
    if len(sys.argv) == 2 and os.path.isdir(sys.argv[1]):
        # Se è stata passata una directory, elabora tutti i file PDF al suo interno
        pdf_dir = sys.argv[1]
        pdf_files = [os.path.join(pdf_dir, f) for f in os.listdir(pdf_dir) 
                    if f.lower().endswith('.pdf')]
        
        if not pdf_files:
            print(f"Nessun file PDF trovato nella directory {pdf_dir}")
            sys.exit(1)
        
        success_count = 0
        total_files = len(pdf_files)
        
        print(f"Trovati {total_files} file PDF da elaborare...")
        
        for i, pdf_path in enumerate(pdf_files, 1):
            print(f"\nElaborazione file {i}/{total_files}: {os.path.basename(pdf_path)}")
            if process_pdf(pdf_path):
                success_count += 1
        
        print(f"\nElaborazione completata. {success_count}/{total_files} file elaborati con successo.")
        sys.exit(0 if success_count == total_files else 1)
    else:
        # Elabora i singoli file passati come argomenti
        success_count = 0
        total_files = len(sys.argv) - 1
        
        for i, pdf_path in enumerate(sys.argv[1:], 1):
            if not os.path.exists(pdf_path):
                print(f"Errore: File {pdf_path} non trovato")
                continue
            
            print(f"\nElaborazione file {i}/{total_files}: {os.path.basename(pdf_path)}")
            if process_pdf(pdf_path):
                success_count += 1
        
        print(f"\nElaborazione completata. {success_count}/{total_files} file elaborati con successo.")
        sys.exit(0 if success_count == total_files else 1)
