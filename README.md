# Google Sheets Data Processor

Questo script processa i dati dal foglio Google Sheets "TABELLA PER RANDOMIZZAZIONE DEI PAZIENTI" e genera tabelle organizzate nel foglio "ANALISI DATI 1".

## Setup

1. Installa le dipendenze:

```bash
pip install -r requirements.txt
```

2. Crea un Service Account su Google Cloud Console:

    - Vai su [Google Cloud Console](https://console.cloud.google.com/)
    - Crea un nuovo progetto
    - Abilita Google Sheets API
    - Crea un Service Account
    - Scarica il file JSON delle credenziali
    - Rinomina il file in `service_account.json`
    - Metti il file nella stessa cartella dello script

3. Condividi il foglio Google Sheets con l'email del Service Account (la trovi nel file JSON scaricato)

## Uso

Esegui lo script:

```bash
python script.py
```

Lo script:

1. Legge i nomi dei pazienti dalla riga 2
2. Legge i parametri dalla colonna A
3. Genera tabelle organizzate a partire dalla colonna L
