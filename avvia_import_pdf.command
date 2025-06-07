#!/bin/bash

# Ottieni il percorso assoluto della directory dello script
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"

# Vai alla directory dello script
cd "$DIR"

# Attiva l'ambiente virtuale se esiste
if [ -d "venv" ]; then
    source venv/bin/activate
fi

# Esegui lo script Python
python3 src/pdf_import_pyqt.py

echo "Applicazione terminata." 