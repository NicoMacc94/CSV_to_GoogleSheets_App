===============================
AGGIORNAMENTO: IMPORTAZIONE DATI HRV
===============================

È stata aggiunta una nuova funzionalità all'applicazione CSV_to_GoogleSheets che consente di importare i dati HRV da file PDF e aggiungerli ai fogli Google Sheets esistenti.

COME FUNZIONA:
-------------
1. L'applicazione estrae nome e cognome dal nome del file PDF
2. Cerca un file Google Sheets esistente con quel nome e cognome
3. Se il file esiste, crea un nuovo foglio chiamato "hrv" all'interno del file
4. Estrae i dati HRV e le immagini dal PDF e li inserisce nel foglio
5. Se il file non esiste, mostra un messaggio di errore

COME USARE LA FUNZIONALITÀ:
--------------------------
1. Avvia l'applicazione PDF_to_GoogleSheets tramite "python3 pdf_import_ui.py"
2. Seleziona il file PDF usando il pulsante "Sfoglia"
3. Clicca su "Elabora PDF" per avviare l'importazione
4. Segui le istruzioni e i messaggi visualizzati

REQUISITI PER I FILE PDF:
-----------------------
1. Il nome del file PDF deve contenere il nome e cognome dell'utente (es. "ALESSANDRO PIZZI GURPPO 1 PROVA 1.pdf")
2. Il file PDF deve essere un report HRV standard con i dati ANS
3. Deve esistere già un file Google Sheets con lo stesso nome e cognome

RISOLUZIONE PROBLEMI:
-------------------
1. Se appare il messaggio "Il soggetto non ha un foglio", devi prima creare un foglio Google Sheets per quell'utente usando la funzionalità CSV dell'applicazione principale.
2. Se i grafici non vengono visualizzati correttamente, assicurati che l'account Google abbia i permessi per accedere alle immagini.
3. Se ricevi errori di autenticazione, elimina il file "token.json" e riavvia l'applicazione per effettuare nuovamente l'accesso.

Per maggiori informazioni o assistenza, contatta il supporto tecnico. 