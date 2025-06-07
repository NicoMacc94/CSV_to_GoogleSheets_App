#!/usr/bin/env python3
import os
import sys
import re
import traceback
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QPushButton, QFileDialog, 
                            QListWidget, QTextEdit, QMessageBox, QFrame,
                            QSplitter, QListWidgetItem, QProgressDialog,
                            QProgressBar)
from PyQt5.QtCore import Qt, QSize, QObject, pyqtSignal, QTimer, QThread
from PyQt5.QtGui import QFont, QIcon

# Aggiungi la directory corrente al path per importare pdf_to_sheets
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from src import pdf_to_sheets_new as pdf_to_sheets

class PDFProcessWorker(QObject):
    """Classe worker per elaborare i PDF in un thread separato."""
    fileStarted = pyqtSignal(str, int, int)
    fileProgress = pyqtSignal(str, bool, str)
    fileCompleted = pyqtSignal(str, bool, str)
    allCompleted = pyqtSignal(int, int, int)
    logMessage = pyqtSignal(str)
    error = pyqtSignal(str, str)
    
    def __init__(self, file_paths):
        super().__init__()
        self.file_paths = file_paths
        self.stopped = False
    
    def process(self):
        """Elabora tutti i file PDF in modo ordinato."""
        def ui_callback(message):
            self.logMessage.emit(message)
        
        try:
            successful, failed, total = pdf_to_sheets.process_files_in_order(
                self.file_paths, ui_callback)
            self.allCompleted.emit(successful, failed, total)
        except Exception as e:
            self.error.emit("Error", f"Errore durante l'elaborazione: {str(e)}")
            self.allCompleted.emit(0, len(self.file_paths), len(self.file_paths))

class PDFImportApp(QMainWindow):
    def __init__(self):
        super().__init__()
        
        # Configurazione della finestra principale
        self.setWindowTitle("CSV to Google Sheets - Importazione PDF HRV")
        self.setMinimumSize(900, 650)
        
        # Widget centrale
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Layout principale
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(15, 15, 15, 15)
        main_layout.setSpacing(10)
        
        # Titolo principale
        header_frame = QFrame()
        header_frame.setObjectName("header_frame")
        header_frame.setFrameShape(QFrame.StyledPanel)
        header_frame.setFrameShadow(QFrame.Raised)
        
        header_layout = QVBoxLayout(header_frame)
        
        title_label = QLabel("Importazione PDF - BioTekna HRV")
        title_label.setObjectName("title_label")
        title_font = QFont()
        title_font.setPointSize(16)
        title_font.setBold(True)
        title_label.setFont(title_font)
        title_label.setAlignment(Qt.AlignCenter)
        header_layout.addWidget(title_label)
        
        subtitle_label = QLabel("Importa i dati HRV dai report PDF di BioTekna e aggiungili ai fogli Google Sheets")
        subtitle_label.setObjectName("subtitle_label")
        subtitle_font = QFont()
        subtitle_font.setPointSize(12)
        subtitle_label.setFont(subtitle_font)
        subtitle_label.setAlignment(Qt.AlignCenter)
        header_layout.addWidget(subtitle_label)
        
        main_layout.addWidget(header_frame)
        
        # Riquadro informativo
        info_frame = QFrame()
        info_frame.setObjectName("info_frame")
        info_frame.setFrameShape(QFrame.StyledPanel)
        info_frame.setFrameShadow(QFrame.Raised)
        
        info_layout = QVBoxLayout(info_frame)
        
        info_title = QLabel("Istruzioni:")
        info_title.setObjectName("info_title")
        info_title_font = QFont()
        info_title_font.setBold(True)
        info_title.setFont(info_title_font)
        info_layout.addWidget(info_title)
        
        info_text = QLabel("• Seleziona uno o più file PDF di report HRV BioTekna\n"
                         "• I nomi dei file devono avere il formato: 'NOME COGNOME GRUPPO X SETT Y PROVA Z.pdf'\n"
                         "• Dove Z può essere 4 o 5\n"
                         "• Per ogni candidato, le prove verranno elaborate in ordine crescente\n"
                         "• Per ogni persona verrà gestito un foglio 'RIEPILOGO HRV' all'interno di un Google Sheet\n"
                         "• Se il Google Sheet non esiste, verrà creato automaticamente\n"
                         "• I dati verranno inseriti nelle colonne F (PROVA 4) o G (PROVA 5)")
        info_text.setObjectName("info_text")
        info_layout.addWidget(info_text)
        
        main_layout.addWidget(info_frame)
        
        # Area centrale
        content_frame = QFrame()
        content_frame.setObjectName("content_frame")
        content_frame.setFrameShape(QFrame.StyledPanel)
        content_frame.setFrameShadow(QFrame.Raised)
        
        content_layout = QVBoxLayout(content_frame)
        
        # Area selezione file
        file_select_layout = QHBoxLayout()
        
        file_label = QLabel("File selezionati:")
        file_label.setObjectName("file_label")
        file_select_layout.addWidget(file_label)
        
        file_select_layout.addStretch()
        
        self.browse_button = QPushButton("Seleziona file PDF...")
        self.browse_button.setObjectName("browse_button")
        self.browse_button.setMinimumWidth(150)
        self.browse_button.clicked.connect(self.browse_files)
        file_select_layout.addWidget(self.browse_button)
        
        content_layout.addLayout(file_select_layout)
        
        # Lista file
        self.files_list = QListWidget()
        self.files_list.setObjectName("files_list")
        self.files_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.files_list.setAlternatingRowColors(True)
        self.files_list.setMinimumHeight(150)
        content_layout.addWidget(self.files_list)
        
        # Pulsanti di gestione file
        file_buttons_layout = QHBoxLayout()
        
        self.remove_button = QPushButton("Rimuovi selezionati")
        self.remove_button.setObjectName("remove_button")
        self.remove_button.clicked.connect(self.remove_selected_files)
        file_buttons_layout.addWidget(self.remove_button)
        
        self.clear_button = QPushButton("Rimuovi tutti")
        self.clear_button.setObjectName("clear_button")
        self.clear_button.clicked.connect(self.clear_files)
        file_buttons_layout.addWidget(self.clear_button)
        
        file_buttons_layout.addStretch()
        
        content_layout.addLayout(file_buttons_layout)
        
        # Barra di progresso complessiva
        progress_layout = QHBoxLayout()
        
        progress_label = QLabel("Progresso:")
        progress_label.setObjectName("progress_label")
        progress_layout.addWidget(progress_label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progress_bar")
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        self.progress_bar.setValue(0)
        progress_layout.addWidget(self.progress_bar)
        
        content_layout.addLayout(progress_layout)
        
        # Separazione tra file e log
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        content_layout.addWidget(separator)
        
        # Area log
        log_layout = QVBoxLayout()
        
        log_header = QHBoxLayout()
        
        log_label = QLabel("Log elaborazione:")
        log_label.setObjectName("log_label")
        log_header.addWidget(log_label)
        
        log_header.addStretch()
        
        self.clear_log_button = QPushButton("Pulisci log")
        self.clear_log_button.setObjectName("clear_log_button")
        self.clear_log_button.clicked.connect(self.clear_log)
        log_header.addWidget(self.clear_log_button)
        
        log_layout.addLayout(log_header)
        
        self.log_text = QTextEdit()
        self.log_text.setObjectName("log_text")
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        content_layout.addLayout(log_layout)
        
        main_layout.addWidget(content_frame)
        
        # Pulsanti in basso
        buttons_layout = QHBoxLayout()
        
        self.status_label = QLabel("Pronto per l'importazione...")
        self.status_label.setObjectName("status_label")
        buttons_layout.addWidget(self.status_label)
        
        buttons_layout.addStretch()
        
        self.process_button = QPushButton("Avvia elaborazione")
        self.process_button.setObjectName("process_button")
        self.process_button.setMinimumWidth(150)
        self.process_button.clicked.connect(self.process_files)
        buttons_layout.addWidget(self.process_button)
        
        self.close_button = QPushButton("Chiudi")
        self.close_button.setObjectName("close_button")
        self.close_button.clicked.connect(self.close)
        buttons_layout.addWidget(self.close_button)
        
        main_layout.addLayout(buttons_layout)
        
        # Thread e worker per elaborazione PDF
        self.thread = None
        self.worker = None
        
        # Struttura dati per organizzare i file
        self.files_by_candidate = {}
        
        # Inizializzazione
        self.log("Applicazione avviata. Seleziona file PDF per iniziare.")
        self.update_buttons_state()
    
    def browse_files(self):
        """Apre la finestra di dialogo per selezionare i file PDF."""
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Seleziona file PDF",
            "",
            "File PDF (*.pdf);;Tutti i file (*.*)"
        )
        
        if not files:
            return
            
        # Aggiungi file alla lista
        added_count = 0
        for file_path in files:
            if not file_path.lower().endswith('.pdf'):
                continue
            
            # Controlla se il file è già nella lista
            for i in range(self.files_list.count()):
                if self.files_list.item(i).data(Qt.UserRole) == file_path:
                    continue
            
            # Estrai informazioni dal nome del file
            filename = os.path.basename(file_path)
            
            # Verifica se il file contiene "PROVA 4" o "PROVA 5"
            if not ("PROVA 4" in filename.upper() or "PROVA 5" in filename.upper()):
                QMessageBox.warning(
                    self,
                    "File non valido",
                    f"Il file {filename} non contiene 'PROVA 4' o 'PROVA 5' nel nome.\n"
                    "Il nome del file deve seguire il formato: 'NOME COGNOME GRUPPO X SETT Y PROVA Z.pdf'"
                )
                continue
            
            # Aggiungi il file alla lista
            item = QListWidgetItem(filename)
            item.setData(Qt.UserRole, file_path)
            self.files_list.addItem(item)
            added_count += 1
        
        if added_count > 0:
            self.log(f"Aggiunti {added_count} nuovi file alla lista.")
        self.update_buttons_state()
    
    def remove_selected_files(self):
        """Rimuove i file selezionati dalla lista."""
        selected_items = self.files_list.selectedItems()
        if not selected_items:
            return
        
        for item in selected_items:
            self.files_list.takeItem(self.files_list.row(item))
        
        self.log(f"Rimossi {len(selected_items)} file dalla lista.")
        self.update_buttons_state()
    
    def clear_files(self):
        """Rimuove tutti i file dalla lista."""
        self.files_list.clear()
        self.log("Lista file svuotata.")
        self.update_buttons_state()
    
    def clear_log(self):
        """Pulisce l'area di log."""
        self.log_text.clear()
        self.log("Log pulito.")
    
    def update_buttons_state(self):
        """Aggiorna lo stato dei pulsanti in base alla lista file."""
        has_files = self.files_list.count() > 0
        self.process_button.setEnabled(has_files)
        self.remove_button.setEnabled(has_files)
        self.clear_button.setEnabled(has_files)
    
    def process_files(self):
        """Avvia l'elaborazione dei file PDF."""
        if self.files_list.count() == 0:
            return
        
        # Raccogli i percorsi dei file
        file_paths = []
        for i in range(self.files_list.count()):
            item = self.files_list.item(i)
            file_paths.append(item.data(Qt.UserRole))
        
        # Crea e avvia il thread di elaborazione
        self.thread = QThread()
        self.worker = PDFProcessWorker(file_paths)
        self.worker.moveToThread(self.thread)
        
        # Connessioni dei segnali
        self.thread.started.connect(self.worker.process)
        self.worker.logMessage.connect(self.log)
        self.worker.allCompleted.connect(self.on_all_completed)
        self.worker.error.connect(self.on_file_error)
        self.worker.allCompleted.connect(self.thread.quit)
        self.thread.finished.connect(self.thread.deleteLater)
        
        # Disabilita i pulsanti durante l'elaborazione
        self.browse_button.setEnabled(False)
        self.process_button.setEnabled(False)
        self.remove_button.setEnabled(False)
        self.clear_button.setEnabled(False)
        
        # Aggiorna la barra di progresso
        self.progress_bar.setValue(0)
        self.status_label.setText("Elaborazione in corso...")
        
        # Avvia il thread
        self.thread.start()
    
    def on_all_completed(self, successful, failed, total):
        """Gestisce il completamento dell'elaborazione."""
        # Riabilita i pulsanti
        self.browse_button.setEnabled(True)
        self.process_button.setEnabled(True)
        self.remove_button.setEnabled(True)
        self.clear_button.setEnabled(True)
        
        # Aggiorna la barra di progresso
        self.progress_bar.setValue(100)
        
        # Mostra il messaggio di completamento
        if failed == 0:
            self.status_label.setText("Elaborazione completata con successo!")
            QMessageBox.information(
                self,
                "Elaborazione completata",
                f"Tutti i {total} file sono stati elaborati con successo."
            )
        else:
            self.status_label.setText(f"Elaborazione completata con {failed} errori.")
            QMessageBox.warning(
                self,
                "Elaborazione completata",
                f"Elaborazione completata con {failed} errori su {total} file.\n"
                "Controlla il log per i dettagli."
            )
    
    def on_file_error(self, title, message):
        """Gestisce gli errori durante l'elaborazione."""
        QMessageBox.critical(self, title, message)
    
    def log(self, message):
        """Aggiunge un messaggio al log."""
        self.log_text.append(message)
        # Scorri automaticamente in fondo
        self.log_text.verticalScrollBar().setValue(
            self.log_text.verticalScrollBar().maximum()
        )
    
    def closeEvent(self, event):
        """Gestisce la chiusura della finestra."""
        if self.thread and self.thread.isRunning():
            reply = QMessageBox.question(
                self,
                "Conferma chiusura",
                "L'elaborazione è ancora in corso. Vuoi davvero chiudere?",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.thread.quit()
                self.thread.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()

def main():
    app = QApplication(sys.argv)
    
    # Imposta lo stile
    app.setStyle('Fusion')
    
    # Crea e mostra la finestra principale
    window = PDFImportApp()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main() 