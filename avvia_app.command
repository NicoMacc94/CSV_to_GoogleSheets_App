#!/bin/bash

# Script di avvio robusto per CSV_to_GoogleSheets
# Gestisce automaticamente i file di configurazione

# Vai alla directory dell'applicazione
cd "$(dirname "$0")"

# Visualizza messaggio di benvenuto
echo "=== CSV to Google Sheets ==="
echo "Preparazione avvio applicazione..."

# Verifica se esiste la cartella config
if [ ! -d "config" ]; then
  echo "Creazione cartella config..."
  mkdir -p config
fi

# Verifica e copia le credenziali del service account
if [ -f "config/service_account.json" ]; then
  echo "File service_account.json trovato in config/"
  cp -f config/service_account.json .
elif [ -f "service_account.json" ]; then
  echo "File service_account.json trovato nella cartella principale"
  # Crea una copia di backup in config/
  cp -f service_account.json config/
else
  echo "ATTENZIONE: File service_account.json non trovato!"
  echo "L'applicazione potrebbe non funzionare correttamente."
  echo "Per favore, assicurati di aver scaricato il file service_account.json"
  echo "e di averlo posizionato nella cartella 'config' o nella cartella principale."
fi

# Rimuovi eventuali token scaduti se l'utente lo desidera
if [ -f "token.json" ]; then
  echo "Trovato file token.json esistente."
  read -p "Vuoi ripristinare il token di autenticazione? (s/n): " risposta
  if [ "$risposta" = "s" ]; then
    echo "Rimozione token..."
    rm token.json
  fi
fi

# Menu di selezione
echo ""
echo "Scegli l'operazione da eseguire:"
echo "1. Importa file CSV"
echo "2. Importa file PDF (dati HRV)"
echo ""
read -p "Inserisci il numero dell'operazione (1-2): " operazione

case $operazione in
  1)
    # Avvia l'applicazione CSV
    echo "Avvio importazione CSV..."
    ./CSV_to_GoogleSheets
    ;;
  2)
    # Avvia l'applicazione di importazione PDF
    echo "Avvio importazione PDF (HRV)..."
    # Verifica che le librerie necessarie siano presenti
    REQUIRED_PACKAGES=("PyPDF2" "pymupdf" "pillow" "google-api-python-client" "google-auth-httplib2" "google-auth-oauthlib" "PyQt5")
    MISSING_PACKAGES=()

    # Controlla i pacchetti mancanti
    for pkg in "${REQUIRED_PACKAGES[@]}"; do
      if ! python3 -c "import $pkg" &>/dev/null; then
        MISSING_PACKAGES+=("$pkg")
      fi
    done

    # Se ci sono pacchetti mancanti, prova a installarli
    if [ ${#MISSING_PACKAGES[@]} -gt 0 ]; then
      echo "Installazione dei pacchetti mancanti..."
      python3 -m pip install --target=./lib "${MISSING_PACKAGES[@]}"
    fi
    
    # Imposta il PYTHONPATH per includere le librerie locali
    export PYTHONPATH="$PYTHONPATH:$PWD:$PWD/lib"
    
    # Avvia l'applicazione con la nuova interfaccia PyQt
    python3 src/pdf_import_pyqt.py
    ;;
  *)
    echo "Operazione non valida. Uscita."
    ;;
esac

# Messaggio di completamento
echo ""
echo "Applicazione terminata."
echo "Se hai riscontrato problemi, assicurati che il file 'service_account.json'"
echo "sia presente nella cartella 'config' o nella cartella principale."
echo ""
read -p "Premi INVIO per chiudere questa finestra..." 