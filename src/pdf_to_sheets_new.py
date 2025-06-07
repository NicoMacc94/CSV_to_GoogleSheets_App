#!/usr/bin/env python3

import re
import PyPDF2
import pdfplumber
from typing import Dict, List, Optional

def extract_ans_data_from_pdf_fixed(pdf_path: str) -> Dict[str, str]:
    """
    Estrae i dati dalla tabella ANS test del PDF BioTekna (versione corretta).
    Estrae prima dalla tabella ANS test, poi integra con Report analitico se necessario.
    """
    
    extracted_data = {}
    
    try:
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
        
        print("DEBUG: Inizio estrazione dati ANS - versione corretta")
        
        # PRIMA: Estrai dalla tabella ANS test (più completa)
        ans_test_pattern = r'ANS test.*?Norms(.*?)(?=Power Spectral Density|Bilateral flow|Report analitico|$)'
        ans_match = re.search(ans_test_pattern, text, re.DOTALL | re.IGNORECASE)
        
        ans_text = ""
        if ans_match:
            ans_text = ans_match.group(1)
            print(f"DEBUG: Sezione ANS test trovata:\n{ans_text[:500]}")
        else:
            # Fallback: usa tutto il testo
            ans_text = text
            print("DEBUG: Sezione ANS test non trovata, uso tutto il testo")
        
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
            print(f"DEBUG: Sezione Report analitico trovata per integrazione")
            
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
        print("\nDEBUG: Riepilogo valori estratti:")
        for param, value in extracted_data.items():
            print(f"  {param}: {value}")
    
    except Exception as e:
        print(f"Errore nell'estrazione del PDF: {e}")
        # Restituisci valori vuoti in caso di errore
        default_params = ['Mean HR', 'SDNN', 'RMSSD', 'pNN50', 'Total power', 
                         'LF/VLF Left', 'LF/VLF Right', 'VLF power', 'LF power', 
                         'HF power', 'SNS LF %', 'PNS HF %', 'THM freq', 'THM power']
        for param in default_params:
            extracted_data[param] = 'N/A'
    
    return extracted_data

def prepare_data_for_sheets_with_headers(extracted_data: Dict[str, str], test_number: int) -> tuple:
    """
    Prepara i dati estratti per l'inserimento in Google Sheets con gli header corretti.
    
    Returns:
        tuple: (header_data, values_data, column_letter)
    """
    
    # Determina la colonna e l'header
    if test_number == 4:
        column_letter = 'F'
        header = 'PROVA 4'
    elif test_number == 5:
        column_letter = 'G'
        header = 'PROVA 5'
    else:
        column_letter = 'F'  # Default
        header = f'PROVA {test_number}'
    
    # Ordine corretto per Google Sheets
    ordered_parameters = [
        'Mean HR',
        'SDNN', 
        'RMSSD',
        'pNN50',
        'Total power',
        'LF/VLF Left',
        'LF/VLF Right',
        'VLF power',
        'LF power', 
        'HF power',
        'SNS LF %',
        'PNS HF %',
        'THM freq',
        'THM power'
    ]
    
    # Prepara i dati
    values_data = []
    for param in ordered_parameters:
        value = extracted_data.get(param, 'N/A')
        values_data.append([value])
        print(f"DEBUG: Preparando {param}: {value}")
    
    return header, values_data, column_letter

def update_google_sheet_with_headers(worksheet, extracted_data: Dict[str, str], test_number: int):
    """
    Aggiorna Google Sheets con header e dati corretti.
    
    Args:
        worksheet: Oggetto worksheet di gspread
        extracted_data: Dati estratti dal PDF
        test_number: Numero del test (4 o 5)
    """
    
    try:
        # Prepara i dati
        header, values_data, column_letter = prepare_data_for_sheets_with_headers(extracted_data, test_number)
        
        print(f"DEBUG: Aggiornamento colonna {column_letter} con header '{header}'")
        
        # 1. Aggiorna l'header nella riga 3
        header_range = f"{column_letter}3"
        worksheet.update(header_range, [[header]])
        print(f"DEBUG: Header aggiornato in {header_range}")
        
        # 2. Aggiorna i dati dalla riga 4 in poi
        start_row = 4
        end_row = start_row + len(values_data) - 1
        data_range = f"{column_letter}{start_row}:{column_letter}{end_row}"
        
        worksheet.update(data_range, values_data)
        print(f"DEBUG: Dati aggiornati in {data_range}")
        
        # 3. Applica la formattazione (se necessario)
        # Qui puoi aggiungere codice per applicare lo stesso stile delle altre colonne
        
        print("Aggiornamento completato con successo!")
        return True
        
    except Exception as e:
        print(f"Errore nell'aggiornamento del foglio: {e}")
        return False

# Funzione di test per verificare l'estrazione
def test_extraction_complete(pdf_path: str):
    """Testa l'estrazione completa dei dati dal PDF."""
    print(f"Testing extraction from: {pdf_path}")
    
    # Estrai il numero del test dal nome del file
    filename = pdf_path.split('/')[-1] if '/' in pdf_path else pdf_path
    test_match = re.search(r'PROVA\s+(\d+)', filename.upper())
    test_number = int(test_match.group(1)) if test_match else 4
    
    print(f"Numero test rilevato: {test_number}")
    
    # Estrai i dati
    extracted_data = extract_ans_data_from_pdf_fixed(pdf_path)
    
    # Prepara per Google Sheets
    header, values_data, column_letter = prepare_data_for_sheets_with_headers(extracted_data, test_number)
    
    print(f"\nRisultato finale:")
    print(f"Header: {header} (colonna {column_letter})")
    print(f"Dati da inserire:")
    for i, value in enumerate(values_data):
        param_names = ['Mean HR', 'SDNN', 'RMSSD', 'pNN50', 'Total power', 
                      'LF/VLF Left', 'LF/VLF Right', 'VLF power', 'LF power', 
                      'HF power', 'SNS LF %', 'PNS HF %', 'THM freq', 'THM power']
        print(f"  Riga {i+4}: {param_names[i]} = {value[0]}")

if __name__ == "__main__":
    # Test con il tuo PDF
    pdf_path = "path/to/your/pdf/file.pdf"  # Sostituisci con il percorso reale
    test_extraction_complete(pdf_path)