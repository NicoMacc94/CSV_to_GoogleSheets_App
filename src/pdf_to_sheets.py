#!/usr/bin/env python3
import os
import sys
import json
import re
import PyPDF2
import fitz  # PyMuPDF
from PIL import Image
import io
import tempfile
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
import difflib  # Per il matching fuzzy

# Definizione delle costanti
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
TOKEN_FILE = 'token.json'

def get_credentials():
    """Ottiene le credenziali del service account."""
    try:
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
        return credentials
    except Exception as e:
        print(f"Errore nel caricamento delle credenziali del service account: {e}")
        raise

def fuzzy_match_candidate(nome, cognome, existing_sheets):
    """
    Cerca il miglior matching fuzzy tra i candidati estratti e i fogli esistenti.
    Considera entrambe le combinazioni: "NOME COGNOME" e "COGNOME NOME".
    
    Args:
        nome: Il nome estratto dal filename
        cognome: Il cognome estratto dal filename
        existing_sheets: Lista di dizionari {id, name} dei fogli esistenti
    
    Returns:
        Tuple (match_found, sheet_info) con il miglior match o None
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
    """Estrae dati da un file PDF."""
    # Estrai nome e cognome e numero della prova dal nome del file
    filename = os.path.basename(pdf_path)
    filename_without_ext = os.path.splitext(filename)[0]
    
    # Normalizzazione del nome file
    # Converti in maiuscolo, rimuovi caratteri speciali e doppi spazi
    normalized_filename = re.sub(r'[_\-\.]+', ' ', filename_without_ext.upper())  # Aggiungiamo . tra i caratteri da rimuovere
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
        
        # Salta token se è una parola chiave o se è seguito da un numero (es: GRUPPO 1, PROVA 2)
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
        # Potrebbe essere o il nome o il cognome
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
            # Qui potremmo chiedere conferma all'utente
    
    print(f"Nome e cognome estratti: {nome} {cognome}, Prova: {prova}")
    
    # Estrazione testo da PDF
    reader = PyPDF2.PdfReader(pdf_path)
    full_text = ""
    for page_num in range(len(reader.pages)):
        page = reader.pages[page_num]
        text = page.extract_text()
        full_text += text
    
    # Funzione helper per cercare valori nel testo
    def find_value_after(text, key, end_marker="\n", offset=0):
        """Cerca un valore dopo una chiave nel testo."""
        start_idx = text.find(key)
        if start_idx == -1:
            return "N/A"
        
        start_idx += len(key) + offset
        end_idx = text.find(end_marker, start_idx)
        if end_idx == -1:
            end_idx = len(text)
        
        value = text[start_idx:end_idx].strip()
        return value if value else "N/A"
    
    # Estrazione dei dati della tabella ANS test
    ans_data = {}
    
    # Sezione ANS test
    ans_start = full_text.find("ANS test")
    if ans_start == -1:
        print("ATTENZIONE: Sezione 'ANS test' non trovata nel PDF!")
        return {
            "nome": nome,
            "cognome": cognome,
            "prova": prova,
            "prova_num": int(prova_num),
            "nome_foglio": nome_foglio,
            "ans_data": {},
            "ans_balance_image": None,
            "scatter_hr_image": None
        }
    
    # Estrai i valori direttamente usando la struttura del testo
    # Mean HR
    mean_hr = find_value_after(full_text, "Mean HR:", "bpm").strip()
    
    # Approccio migliorato per estrarre la norma di Mean HR
    mean_hr_norm = "N/A"
    mean_hr_idx = full_text.find("Mean HR:")
    if mean_hr_idx != -1:
        # Cerca apertura parentesi quadra dopo "Mean HR:"
        open_bracket_idx = full_text.find("[", mean_hr_idx)
        if open_bracket_idx != -1 and open_bracket_idx < mean_hr_idx + 50:  # Cerca entro 50 caratteri
            # Cerca chiusura parentesi quadra
            close_bracket_idx = full_text.find("]", open_bracket_idx)
            if close_bracket_idx != -1:
                # Estrai il valore tra parentesi quadre
                norm_text = full_text[open_bracket_idx:close_bracket_idx+1].strip()
                if norm_text:
                    mean_hr_norm = norm_text
    
    # SDNN
    sdnn = find_value_after(full_text, "SDNN:", "ms").strip()
    sdnn_norm = find_value_after(full_text, "min", "ms", 0).strip()
    if sdnn_norm and sdnn_norm.isdigit():
        sdnn_norm = f"min {sdnn_norm}"
    else:
        sdnn_norm = "N/A"
    
    # RMSSD
    rmssd = find_value_after(full_text, "RMSSD:", "ms").strip()
    rmssd_norm_idx = full_text.find("min", full_text.find("RMSSD:"))
    if rmssd_norm_idx != -1:
        rmssd_norm_end = full_text.find("ms", rmssd_norm_idx)
        if rmssd_norm_end != -1:
            rmssd_norm_val = full_text[rmssd_norm_idx + 3:rmssd_norm_end].strip()
            if rmssd_norm_val and rmssd_norm_val.isdigit():
                rmssd_norm = f"min {rmssd_norm_val}"
            else:
                rmssd_norm = "N/A"
        else:
            rmssd_norm = "N/A"
    else:
        rmssd_norm = "N/A"
    
    # Total power
    total_power = find_value_after(full_text, "Total power:", "\n").strip()
    if total_power == "N/A" or not total_power:
        # Prova con un approccio alternativo
        tp_start = full_text.find("Total power:")
        if tp_start != -1:
            # Cerca il primo numero dopo "Total power:"
            for i in range(tp_start + 12, tp_start + 50):
                if i < len(full_text) and (full_text[i].isdigit() or full_text[i] == '.'):
                    # Abbiamo trovato l'inizio del numero
                    j = i
                    while j < len(full_text) and (full_text[j].isdigit() or full_text[j] == '.'):
                        j += 1
                    total_power = full_text[i:j]
                    break
    
    total_power_norm_idx = full_text.find("min", full_text.find("Total power:"))
    total_power_norm = "N/A"
    if total_power_norm_idx != -1:
        norm_value = full_text[total_power_norm_idx + 3:total_power_norm_idx + 10].strip()
        if norm_value and any(c.isdigit() for c in norm_value):
            # Estrai solo la parte numerica
            digits_only = ''.join(c for c in norm_value if c.isdigit() or c == '.')
            if digits_only:
                total_power_norm = f"min {digits_only}"
    
    # LF/VLF Left e Right - miglioramento dell'estrazione
    # LF/VLF Left
    lfvlf_left = "N/A"
    lfvlf_left_segments = ["LF/VLF Left", "LF / VLF Left", "LF/ VLF Left", "LF /VLF Left"]
    
    for segment in lfvlf_left_segments:
        segment_idx = full_text.find(segment)
        if segment_idx != -1:
            # Cerca il primo numero dopo questo segmento
            for i in range(segment_idx + len(segment), segment_idx + len(segment) + 50):
                if i < len(full_text) and (full_text[i].isdigit() or full_text[i] == '.'):
                    # Abbiamo trovato l'inizio del numero
                    j = i
                    while j < len(full_text) and (full_text[j].isdigit() or full_text[j] == '.'):
                        j += 1
                    lfvlf_left = full_text[i:j]
                    break
            if lfvlf_left != "N/A":
                break
    
    # Estrai la norma per LF/VLF Left
    lfvlf_left_norm = "N/A"
    for segment in lfvlf_left_segments:
        segment_idx = full_text.find(segment)
        if segment_idx != -1:
            min_idx = full_text.find("min", segment_idx, segment_idx + 100)
            if min_idx != -1:
                # Cerca il primo numero dopo "min"
                for i in range(min_idx + 3, min_idx + 20):
                    if i < len(full_text) and (full_text[i].isdigit() or full_text[i] == '.'):
                        # Abbiamo trovato l'inizio del numero
                        j = i
                        while j < len(full_text) and (full_text[j].isdigit() or full_text[j] == '.'):
                            j += 1
                        lfvlf_left_norm = f"min {full_text[i:j]}"
                        break
                if lfvlf_left_norm != "N/A":
                    break
    
    # LF/VLF Right
    lfvlf_right = "N/A"
    lfvlf_right_segments = ["LF/VLF Right", "LF / VLF Right", "LF/ VLF Right", "LF /VLF Right"]
    
    for segment in lfvlf_right_segments:
        segment_idx = full_text.find(segment)
        if segment_idx != -1:
            # Cerca il primo numero dopo questo segmento
            for i in range(segment_idx + len(segment), segment_idx + len(segment) + 50):
                if i < len(full_text) and (full_text[i].isdigit() or full_text[i] == '.'):
                    # Abbiamo trovato l'inizio del numero
                    j = i
                    while j < len(full_text) and (full_text[j].isdigit() or full_text[j] == '.'):
                        j += 1
                    lfvlf_right = full_text[i:j]
                    break
            if lfvlf_right != "N/A":
                break
    
    # Estrai la norma per LF/VLF Right
    lfvlf_right_norm = "N/A"
    for segment in lfvlf_right_segments:
        segment_idx = full_text.find(segment)
        if segment_idx != -1:
            min_idx = full_text.find("min", segment_idx, segment_idx + 100)
            if min_idx != -1:
                # Cerca il primo numero dopo "min"
                for i in range(min_idx + 3, min_idx + 20):
                    if i < len(full_text) and (full_text[i].isdigit() or full_text[i] == '.'):
                        # Abbiamo trovato l'inizio del numero
                        j = i
                        while j < len(full_text) and (full_text[j].isdigit() or full_text[j] == '.'):
                            j += 1
                        lfvlf_right_norm = f"min {full_text[i:j]}"
                        break
                if lfvlf_right_norm != "N/A":
                    break
    
    # Crea il dizionario dei dati ANS
    ans_data = {
        "Mean HR": {"value": mean_hr, "norm": mean_hr_norm},
        "SDNN": {"value": sdnn, "norm": sdnn_norm},
        "RMSSD": {"value": rmssd, "norm": rmssd_norm},
        "Total power": {"value": total_power, "norm": total_power_norm},
        "LF/VLF Left": {"value": lfvlf_left, "norm": lfvlf_left_norm},
        "LF/VLF Right": {"value": lfvlf_right, "norm": lfvlf_right_norm}
    }
    
    # VLF power - miglioramento estrazione valore
    vlf_power = find_value_after(full_text, "VLF power:", "\n").strip()
    if vlf_power == "N/A" or not vlf_power:
        # Prova un approccio alternativo
        vlf_start = full_text.find("VLF power:")
        if vlf_start != -1:
            # Cerca il primo numero dopo "VLF power:"
            for i in range(vlf_start + 10, vlf_start + 50):
                if i < len(full_text) and (full_text[i].isdigit() or full_text[i] == '.'):
                    # Abbiamo trovato l'inizio del numero
                    j = i
                    while j < len(full_text) and (full_text[j].isdigit() or full_text[j] == '.'):
                        j += 1
                    vlf_power = full_text[i:j]
                    break
    
    vlf_norm_idx = full_text.find("max", full_text.find("VLF power:"))
    vlf_norm = "N/A"
    if vlf_norm_idx != -1:
        norm_value = full_text[vlf_norm_idx + 3:vlf_norm_idx + 10].strip()
        if norm_value and any(c.isdigit() for c in norm_value):
            digits_only = ''.join(c for c in norm_value if c.isdigit() or c == '.')
            if digits_only:
                vlf_norm = f"max {digits_only}"
    ans_data["VLF power"] = {"value": vlf_power, "norm": vlf_norm}
    
    # LF power - miglioramento estrazione valore
    lf_power = find_value_after(full_text, "LF power:", "\n").strip()
    if lf_power == "N/A" or not lf_power:
        # Prova un approccio alternativo
        lf_start = full_text.find("LF power:")
        if lf_start != -1:
            # Cerca il primo numero dopo "LF power:"
            for i in range(lf_start + 9, lf_start + 50):
                if i < len(full_text) and (full_text[i].isdigit() or full_text[i] == '.'):
                    # Abbiamo trovato l'inizio del numero
                    j = i
                    while j < len(full_text) and (full_text[j].isdigit() or full_text[j] == '.'):
                        j += 1
                    lf_power = full_text[i:j]
                    break
    
    lf_norm_idx = full_text.find("min", full_text.find("LF power:"))
    lf_norm = "N/A"
    if lf_norm_idx != -1:
        norm_value = full_text[lf_norm_idx + 3:lf_norm_idx + 10].strip()
        if norm_value and any(c.isdigit() for c in norm_value):
            digits_only = ''.join(c for c in norm_value if c.isdigit() or c == '.')
            if digits_only:
                lf_norm = f"min {digits_only}"
    ans_data["LF power"] = {"value": lf_power, "norm": lf_norm}
    
    # HF power - miglioramento estrazione valore
    hf_power = find_value_after(full_text, "HF power:", "\n").strip()
    if hf_power == "N/A" or not hf_power:
        # Prova un approccio alternativo
        hf_start = full_text.find("HF power:")
        if hf_start != -1:
            # Cerca il primo numero dopo "HF power:"
            for i in range(hf_start + 9, hf_start + 50):
                if i < len(full_text) and (full_text[i].isdigit() or full_text[i] == '.'):
                    # Abbiamo trovato l'inizio del numero
                    j = i
                    while j < len(full_text) and (full_text[j].isdigit() or full_text[j] == '.'):
                        j += 1
                    hf_power = full_text[i:j]
                    break
    
    hf_norm_idx = full_text.find("min", full_text.find("HF power:"))
    hf_norm = "N/A"
    if hf_norm_idx != -1:
        norm_value = full_text[hf_norm_idx + 3:hf_norm_idx + 10].strip()
        if norm_value and any(c.isdigit() for c in norm_value):
            digits_only = ''.join(c for c in norm_value if c.isdigit() or c == '.')
            if digits_only:
                hf_norm = f"min {digits_only}"
    ans_data["HF power"] = {"value": hf_power, "norm": hf_norm}
    
    # SNS LF % - miglioramento estrazione
    sns_lf = find_value_after(full_text, "SNS LF %:", "%").strip()
    if sns_lf == "N/A" or not sns_lf:
        sns_start = full_text.find("SNS LF %:")
        if sns_start != -1:
            # Cerca il primo numero dopo "SNS LF %:"
            for i in range(sns_start + 9, sns_start + 50):
                if i < len(full_text) and (full_text[i].isdigit() or full_text[i] == '.'):
                    # Abbiamo trovato l'inizio del numero
                    j = i
                    while j < len(full_text) and (full_text[j].isdigit() or full_text[j] == '.'):
                        j += 1
                    sns_lf = full_text[i:j]
                    break
    
    sns_norm_idx = full_text.find("min", full_text.find("SNS LF %:"))
    sns_norm = "N/A"
    if sns_norm_idx != -1:
        norm_value = full_text[sns_norm_idx + 3:sns_norm_idx + 10].strip()
        if norm_value and any(c.isdigit() for c in norm_value):
            digits_only = ''.join(c for c in norm_value if c.isdigit() or c == '%')
            if digits_only and "%" in digits_only:
                sns_norm = f"min {digits_only.replace('%', '')}%"
            elif digits_only:
                sns_norm = f"min {digits_only}%"
    ans_data["SNS LF %"] = {"value": sns_lf, "norm": sns_norm}
    
    # PNS HF % - miglioramento estrazione
    pns_hf = find_value_after(full_text, "PNS HF %:", "%").strip()
    if pns_hf == "N/A" or not pns_hf:
        pns_start = full_text.find("PNS HF %:")
        if pns_start != -1:
            # Cerca il primo numero dopo "PNS HF %:"
            for i in range(pns_start + 9, pns_start + 50):
                if i < len(full_text) and (full_text[i].isdigit() or full_text[i] == '.'):
                    # Abbiamo trovato l'inizio del numero
                    j = i
                    while j < len(full_text) and (full_text[j].isdigit() or full_text[j] == '.'):
                        j += 1
                    pns_hf = full_text[i:j]
                    break
    ans_data["PNS HF %"] = {"value": pns_hf, "norm": "N/A"}  # Non sembra avere norme
    
    # Aggiungi THM freq e THM power
    thm_freq = find_value_after(full_text, "THM freq:", "Hz").strip()
    thm_freq_norm_idx = full_text.find("[", full_text.find("THM freq:"))
    thm_freq_norm = "N/A"
    if thm_freq_norm_idx != -1:
        thm_freq_norm_end = full_text.find("]", thm_freq_norm_idx)
        if thm_freq_norm_end != -1:
            thm_freq_norm = full_text[thm_freq_norm_idx:thm_freq_norm_end+1].strip()
    ans_data["THM freq"] = {"value": thm_freq, "norm": thm_freq_norm}
    
    thm_power = find_value_after(full_text, "THM power:", "ms").strip()
    thm_power_norm_idx = full_text.find("min", full_text.find("THM power:"))
    thm_power_norm = "N/A"
    if thm_power_norm_idx != -1:
        norm_value = full_text[thm_power_norm_idx + 3:thm_power_norm_idx + 10].strip()
        if norm_value and any(c.isdigit() for c in norm_value):
            digits_only = ''.join(c for c in norm_value if c.isdigit() or c == '.')
            if digits_only:
                thm_power_norm = f"min {digits_only}"
    ans_data["THM power"] = {"value": thm_power, "norm": thm_power_norm}
    
    # Estrazione immagini usando PyMuPDF
    doc = fitz.open(pdf_path)
    images = []
    
    # Cerchiamo le immagini specifiche
    for page_num in range(len(doc)):
        page = doc[page_num]
        image_list = page.get_images(full=True)
        
        for img_index, img in enumerate(image_list):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image = Image.open(io.BytesIO(image_bytes))
            
            # Salva l'immagine temporaneamente
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".png")
            image.save(temp_file.name)
            
            images.append({
                "path": temp_file.name,
                "page": page_num + 1,
                "index": img_index + 1
            })
    
    # Seleziona esattamente le immagini corrette in base agli indici confermati
    # Indice 4 per ANS Balance Power e indice 6 per Scatter - Heart rate
    
    # ANS Balance Power - indice 4 nella pagina 1
    ans_balance_image = next((img for img in images if img["page"] == 1 and img["index"] == 4), None)
    
    # Scatter - Heart rate - indice 6 nella pagina 1
    scatter_hr_image = next((img for img in images if img["page"] == 1 and img["index"] == 6), None)
    
    # Log per debug
    print(f"Immagini trovate nella pagina 1: {sum(1 for img in images if img['page'] == 1)}")
    print(f"ANS Balance Power (indice 4): {'Trovato' if ans_balance_image else 'Non trovato'}")
    print(f"Scatter - Heart rate (indice 6): {'Trovato' if scatter_hr_image else 'Non trovato'}")
    
    # Se non troviamo le immagini agli indici esatti, proviamo un fallback con il metodo precedente
    if not ans_balance_image:
        print("Fallback per ANS Balance Power: ricerca tra le immagini 3-5")
        ans_balance_image = next((img for img in images if img["page"] == 1 and img["index"] >= 3 and img["index"] <= 5), None)
        
    if not scatter_hr_image:
        print("Fallback per Scatter - Heart rate: ricerca tra le immagini 5-7")
        scatter_hr_image = next((img for img in images if img["page"] == 1 and img["index"] >= 5 and img["index"] <= 7), None)
    
    return {
        "nome": nome,
        "cognome": cognome,
        "prova": prova,
        "prova_num": int(prova_num),
        "nome_foglio": nome_foglio,
        "ans_data": ans_data,
        "ans_balance_image": ans_balance_image["path"] if ans_balance_image else None,
        "scatter_hr_image": scatter_hr_image["path"] if scatter_hr_image else None
    }

def find_google_sheet_by_name(service, name):
    """Cerca un file Google Sheets per nome."""
    query = f"name = '{name}' and mimeType = 'application/vnd.google-apps.spreadsheet'"
    results = service.files().list(q=query, fields="files(id, name)").execute()
    files = results.get('files', [])
    
    if not files:
        return None
    
    return files[0]  # Restituisce il primo file trovato

def process_strength_data(sheets_service, spreadsheet_id):
    """
    Raccoglie i dati di forza dai fogli "SETT X PROVA Y" e li aggiunge nel foglio "RIEPILOGO HRV".
    Questa funzione deve essere chiamata solo dopo la creazione di un nuovo foglio RIEPILOGO HRV.
    """
    print("Elaborazione dati di forza dai fogli SETT X PROVA Y...")
    
    try:
        # Ottieni la lista dei fogli nel documento
        spreadsheet = sheets_service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = spreadsheet.get('sheets', [])
        
        # Filtra i fogli che seguono il pattern "SETT X PROVA Y"
        sett_prova_sheets = []
        pattern = re.compile(r"SETT\s+(\d+)\s+PROVA\s+(\d+)", re.IGNORECASE)
        
        for sheet in sheets:
            sheet_title = sheet['properties']['title']
            match = pattern.match(sheet_title)
            if match:
                sett_num = int(match.group(1))
                prova_num = int(match.group(2))
                sett_prova_sheets.append({
                    'title': sheet_title,
                    'sett': sett_num,
                    'prova': prova_num
                })
        
        if not sett_prova_sheets:
            print("Nessun foglio 'SETT X PROVA Y' trovato.")
            return
        
        print(f"Trovati {len(sett_prova_sheets)} fogli di tipo 'SETT X PROVA Y'.")
        
        # Raccogliamo tutti i dati prima di scriverli
        strength_data = []
        
        for sheet_info in sett_prova_sheets:
            sheet_title = sheet_info['title']
            prova_num = sheet_info['prova']
            
            print(f"Leggendo dati da {sheet_title}, cella H7...")
            
            # Leggi il valore dalla cella H7
            range_name = f"'{sheet_title}'!H7"
            result = sheets_service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id, range=range_name).execute()
            values = result.get('values', [])
            
            if not values or not values[0]:
                print(f"Nessun valore trovato in {sheet_title}!H7")
                continue
            
            strength_value = values[0][0]
            print(f"Valore trovato: {strength_value}")
            
            strength_data.append({
                'prova': prova_num,
                'value': strength_value
            })
        
        if not strength_data:
            print("Nessun dato di forza trovato.")
            return
        
        # Trova il foglio da usare (RIEPILOGO HRV o Hrv per compatibilità)
        sheet_title = 'RIEPILOGO HRV'
        for sheet in sheets:
            if sheet['properties']['title'].lower() == 'riepilogo hrv' or sheet['properties']['title'].lower() == 'hrv':
                sheet_title = sheet['properties']['title']
                break
        
        # Verifica se ci sono già dati nel foglio nelle righe riservate
        range_check = f"{sheet_title}!A17:D25"
        result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=range_check).execute()
        existing_values = result.get('values', [])
        
        # Estendi esistenti per evitare errori di indice
        while len(existing_values) < 9:  # Garantisci almeno 9 righe (da 17 a 25)
            existing_values.append([])
        
        # Scrivi i dati nel foglio
        updates = []
        
        for data in strength_data:
            prova_num = data['prova']
            value = data['value']
            
            # Calcola la riga in cui scrivere (17 + prova_num - 1)
            row_index = 16 + prova_num  # 17 è la prima riga
            label = f"FORZA {prova_num}"
            
            # Verifica se la riga è già occupata
            if row_index - 17 < len(existing_values) and existing_values[row_index - 17]:
                existing_row = existing_values[row_index - 17]
                if len(existing_row) > 0 and existing_row[0]:
                    print(f"Riga {row_index+1} già occupata con {existing_row[0]}, salto...")
                    continue
            
            # Determina la colonna in cui scrivere (B o D)
            col_index = 3 if prova_num == 4 else 1  # D = 3, B = 1
            col_letter = 'D' if prova_num == 4 else 'B'
            
            # Aggiungi l'aggiornamento per l'etichetta (colonna A)
            updates.append({
                'range': f'{sheet_title}!A{row_index+1}',
                'values': [[label]]
            })
            
            # Aggiungi l'aggiornamento per il valore (colonna B o D)
            updates.append({
                'range': f'{sheet_title}!{col_letter}{row_index+1}',
                'values': [[value]]
            })
        
        # Esegui gli aggiornamenti in batch
        if updates:
            body = {
                'valueInputOption': 'USER_ENTERED',
                'data': updates
            }
            result = sheets_service.spreadsheets().values().batchUpdate(
                spreadsheetId=spreadsheet_id, body=body).execute()
            
            print(f"Dati di forza aggiornati con successo: {len(updates)//2} valori.")
        else:
            print("Nessun dato da aggiornare.")
            
    except Exception as e:
        print(f"Errore durante l'elaborazione dei dati di forza: {str(e)}")
        traceback.print_exc()

def create_hrv_sheet(sheets_service, spreadsheet_id, ans_data, prova, prova_num):
    """Crea un nuovo foglio 'RIEPILOGO HRV' nel file Google Sheets specificato o aggiorna quello esistente."""
    # Controlla se il foglio 'RIEPILOGO HRV' esiste già
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
            print(f"Trovato foglio '{sheet['properties']['title']}' esistente")
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
        # Riga 3: Intestazioni (A: Parameter, B: Norms, ecc.)
        header_values = [["Parameter", "Norms", "", "", "", ""]]
        header_range = 'RIEPILOGO HRV!A3:F3'
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=header_range,
            valueInputOption='RAW', body={'values': header_values}).execute()
        
        # Prepara i valori per i parametri ANS (prima colonna) e i loro valori normativi (seconda colonna)
        data_values = []
        
        # Aggiungi tutti i parametri e le loro norme
        for key, data in ans_data.items():
            data_values.append([key, data["norm"], "", "", "", ""])  # Parametro, Norms, valori vuoti
        
        # Invia i dati al foglio
        data_range = f'RIEPILOGO HRV!A4:F{3 + len(data_values)}'
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=data_range,
            valueInputOption='RAW', body={'values': data_values}).execute()
        
        # Formatta solo l'intestazione della riga 3
        format_requests = [
            # Formatta l'intestazione della riga 3
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
    
    # Determina la colonna di destinazione in base al numero della prova
    # Prova 1 -> Colonna C (indice 2)
    # Prova 2 -> Colonna D (indice 3)
    # Prova 3 -> Colonna E (indice 4)
    # Prova 4 -> Colonna F (indice 5)
    # ecc.
    col_index = 1 + prova_num  # 0-based + offset di 2 (A e B sono già occupate)
    col_letter = chr(65 + col_index)  # Converti in lettera (A=65, B=66, C=67, ecc.)
    
    # Verifica se c'è già un'intestazione per questa prova
    header_range = f'{sheet_title}!{col_letter}3'
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id, range=header_range).execute()
    header_value = result.get('values', [[""]])[0][0]
    
    if header_value == prova:
        print(f"{prova} è già presente nella colonna {col_letter}, non verranno aggiunti dati duplicati.")
        return False, col_index, -1
    
    # Aggiungi l'intestazione della prova alla riga 3
    header_values = [[prova]]
    header_range = f'{sheet_title}!{col_letter}3'
    sheets_service.spreadsheets().values().update(
        spreadsheetId=spreadsheet_id, range=header_range,
        valueInputOption='RAW', body={'values': header_values}).execute()
    
    # Se il foglio esiste già, leggi i parametri esistenti
    if sheet_exists:
        # Leggi i parametri esistenti nella prima colonna (a partire dalla riga 4)
        params_range = f'{sheet_title}!A4:A100'
        params_result = sheets_service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=params_range).execute()
        param_rows = params_result.get('values', [])
        existing_params = [row[0] if row else "" for row in param_rows]
        
        # Prepara i valori per i parametri ANS
        data_values = []
        processed_params = []
        
        # Aggiungi i valori per i parametri esistenti
        for i, param in enumerate(existing_params):
            if param and param in ans_data:
                # Validazione dei dati
                value = ans_data[param]["value"]
                if value == "N/A" or not value:
                    print(f"WARNING: Valore mancante o invalido per {param} in {prova}")
                    value = ""
                
                data_values.append([value])
                processed_params.append(param)
            else:
                if param:  # Se c'è un parametro ma non abbiamo il valore, mettiamo una cella vuota
                    data_values.append([""])
    else:
        # Se è un foglio nuovo, usa i parametri dall'ans_data direttamente
        data_values = []
        for key, data in ans_data.items():
            # Validazione dei dati
            value = data["value"]
            if value == "N/A" or not value:
                print(f"WARNING: Valore mancante o invalido per {key} in {prova}")
                value = ""
            
            data_values.append([value])
    
    # Inserisci i valori nella colonna della prova
    if data_values:
        data_range = f'{sheet_title}!{col_letter}4:{col_letter}{4 + len(data_values) - 1}'
        sheets_service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id, range=data_range,
            valueInputOption='RAW', body={'values': data_values}).execute()
    
    # Formatta l'intestazione della prova
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
    
    # Elabora i dati di forza dai fogli SETT X PROVA Y
    if not sheet_exists:
        process_strength_data(sheets_service, spreadsheet_id)
    
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
    header_value = result.get('values', [[""]])[0][0]
    
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
    query = "mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false"
    results = drive_service.files().list(q=query, fields="files(id, name)", pageSize=1000).execute()
    files = results.get('files', [])
    return files

def process_pdf(pdf_path):
    """Elabora un file PDF e aggiunge i dati a Google Sheets."""
    try:
        # Ottieni le credenziali
        print("Autenticazione con Google...")
        creds = get_credentials()
        
        # Crea servizi Google API
        drive_service = build('drive', 'v3', credentials=creds)
        sheets_service = build('sheets', 'v4', credentials=creds)
        
        # Ottieni tutti i fogli esistenti
        existing_sheets = find_all_google_sheets(drive_service)
        
        # Estrai dati dal PDF
        print(f"Estrazione dati da {pdf_path}...")
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
        print(f"Ricerca del foglio Google Sheets per {nome_completo}...")
        sheet = find_google_sheet_by_name(drive_service, nome_completo)
        
        if not sheet:
            print(f"File Google Sheets per {nome_completo} non trovato. Creazione nuovo foglio...")
            spreadsheet_id = create_google_sheet(drive_service, sheets_service, nome_completo)
            if not spreadsheet_id:
                print(f"Impossibile creare un nuovo foglio per {nome_completo}.")
                return False
        else:
            print(f"File Google Sheets trovato: {sheet['name']} (ID: {sheet['id']})")
            spreadsheet_id = sheet['id']
        
        # Crea/aggiorna il foglio 'RIEPILOGO HRV'
        print(f"Gestione foglio 'RIEPILOGO HRV' per {prova}...")
        update_success, column_index, last_row = create_hrv_sheet(sheets_service, spreadsheet_id, data['ans_data'], prova, prova_num)
        
        if not update_success:
            print(f"Nessun aggiornamento necessario per {prova}.")
            
            # Pulisci i file temporanei
            if data['ans_balance_image'] and os.path.exists(data['ans_balance_image']):
                os.unlink(data['ans_balance_image'])
            
            if data['scatter_hr_image'] and os.path.exists(data['scatter_hr_image']):
                os.unlink(data['scatter_hr_image'])
                
            return True
        
        # Gestisci il caricamento delle immagini
        image_ids = []
        
        # Crea una cartella per le immagini se non esiste
        folder_name = f"{nome_completo} - Immagini HRV"
        folder_id = find_or_create_folder(drive_service, folder_name)
        
        if not folder_id:
            print(f"Impossibile creare la cartella per le immagini su Google Drive")
        else:
            # Carica l'immagine ANS Balance Power se disponibile
            if data['ans_balance_image']:
                print("Caricamento immagine ANS Balance Power...")
                img_title = f"{nome_completo} - ANS Balance Power - {prova}"
                img_id = upload_image_to_drive(drive_service, data['ans_balance_image'], img_title, folder_id)
                if img_id:
                    image_ids.append(img_id)
            
            # Carica l'immagine Scatter HR se disponibile
            if data['scatter_hr_image']:
                print("Caricamento immagine Scatter - Heart rate...")
                img_title = f"{nome_completo} - Scatter HR - {prova}"
                img_id = upload_image_to_drive(drive_service, data['scatter_hr_image'], img_title, folder_id)
                if img_id:
                    image_ids.append(img_id)
        
        # Inserisci le immagini nel foglio se ne abbiamo almeno una
        if image_ids:
            print("Inserimento immagini nel foglio...")
            insert_images_in_sheet(
                sheets_service, 
                drive_service,
                spreadsheet_id,
                image_ids,
                pdf_path,
                prova,
                prova_num
            )
        
        print(f"Elaborazione completata con successo per {nome_completo}, {prova}!")
        print(f"Puoi visualizzare il foglio qui: https://docs.google.com/spreadsheets/d/{spreadsheet_id}")
        
        # Pulisci i file temporanei
        if data['ans_balance_image'] and os.path.exists(data['ans_balance_image']):
            os.unlink(data['ans_balance_image'])
        
        if data['scatter_hr_image'] and os.path.exists(data['scatter_hr_image']):
            os.unlink(data['scatter_hr_image'])
        
        return True
    
    except Exception as e:
        print(f"Errore durante l'elaborazione del PDF: {e}")
        import traceback
        traceback.print_exc()
        return False

def process_files_in_order(file_paths, ui_callback=None):
    """
    Elabora i file PDF in ordine, raggruppati per candidato e ordinati per numero di prova.
    
    Args:
        file_paths: Lista di percorsi dei file PDF da elaborare
        ui_callback: Funzione di callback per aggiornare l'interfaccia utente
    
    Returns:
        Tuple (successo_totale, fallimento_totale, totale_file)
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
                traceback.print_exc()
        
        ui_callback(f"✅ Completata elaborazione per {candidate_id}")
    
    # Restituisci statistiche
    return success_total, failed_total, total_files

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python pdf_to_sheets.py <percorso_pdf> [<percorso_pdf2> ...]")
        print("  oppure: python pdf_to_sheets.py <cartella_pdf>")
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