#!/usr/bin/env python3
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

# Aggiungi la directory corrente al path per importare pdf_to_sheets
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from src import pdf_to_sheets

class PDFImportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Importa Dati HRV da PDF")
        self.root.geometry("700x600")
        self.root.resizable(True, True)
        
        # Configura lo stile dell'applicazione
        self.setup_styles()
        
        # Crea l'interfaccia utente
        self.create_ui()
    
    def setup_styles(self):
        """Configura gli stili per l'interfaccia."""
        style = ttk.Style()
        
        # Configurazione per diversi sistemi operativi
        if os.name == "nt":  # Windows
            style.configure("Title.TLabel", font=("Segoe UI", 18, "bold"), padding=10)
            style.configure("Subtitle.TLabel", font=("Segoe UI", 12), padding=5)
            style.configure("Heading.TLabel", font=("Segoe UI", 11, "bold"), padding=5)
            style.configure("Accent.TButton", font=("Segoe UI", 10))
        else:  # macOS e Linux
            style.configure("Title.TLabel", font=("Helvetica", 18, "bold"), padding=10)
            style.configure("Subtitle.TLabel", font=("Helvetica", 12), padding=5)
            style.configure("Heading.TLabel", font=("Helvetica", 11, "bold"), padding=5)
            style.configure("Accent.TButton", font=("Helvetica", 12))
    
    def create_ui(self):
        # Frame principale con padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Titolo e sottotitolo
        title_label = ttk.Label(main_frame, text="Importazione Dati HRV da PDF", style="Title.TLabel")
        title_label.pack(pady=(0, 10))
        
        subtitle_label = ttk.Label(
            main_frame, 
            text="Importa i dati da un report PDF di BioTekna e aggiungili a un foglio Google Sheets.",
            style="Subtitle.TLabel",
            wraplength=650
        )
        subtitle_label.pack(pady=(0, 20))
        
        # Cornice informativa
        info_frame = ttk.LabelFrame(main_frame, text="Importante", padding="10")
        info_frame.pack(fill=tk.X, pady=10)
        
        info_text = ttk.Label(
            info_frame, 
            text="• Il nome del file PDF deve avere il formato: 'NOME COGNOME GRUPPO X PROVA Y.pdf'\n"
                 "• Per ogni persona verrà creato un foglio Google Sheet (se non esiste) con il foglio 'Hrv'\n"
                 "• Ogni prova (PROVA 1, PROVA 2, ecc.) verrà inserita in una colonna separata",
            wraplength=650
        )
        info_text.pack(pady=5)
        
        # Frame per selezione file
        file_frame = ttk.LabelFrame(main_frame, text="Selezione File", padding="10")
        file_frame.pack(fill=tk.X, pady=10)
        
        self.file_paths = []
        
        file_label = ttk.Label(file_frame, text="Seleziona uno o più file PDF con i dati HRV:")
        file_label.pack(anchor="w", pady=(0, 5))
        
        # Lista dei file selezionati
        self.files_listbox = tk.Listbox(file_frame, height=6, selectmode=tk.EXTENDED)
        self.files_listbox.pack(fill=tk.X, expand=True, pady=5)
        
        # Scrollbar per la lista
        scrollbar = ttk.Scrollbar(self.files_listbox, orient="vertical", command=self.files_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.files_listbox.config(yscrollcommand=scrollbar.set)
        
        # Pulsanti per la gestione dei file
        file_buttons_frame = ttk.Frame(file_frame)
        file_buttons_frame.pack(fill=tk.X, pady=5)
        
        browse_button = ttk.Button(file_buttons_frame, text="Aggiungi File", command=self.browse_files)
        browse_button.pack(side=tk.LEFT, padx=5)
        
        remove_button = ttk.Button(file_buttons_frame, text="Rimuovi Selezionati", command=self.remove_selected)
        remove_button.pack(side=tk.LEFT, padx=5)
        
        clear_button = ttk.Button(file_buttons_frame, text="Rimuovi Tutti", command=self.clear_files)
        clear_button.pack(side=tk.LEFT, padx=5)
        
        # Separatore
        separator = ttk.Separator(main_frame, orient="horizontal")
        separator.pack(fill=tk.X, pady=20)
        
        # Pulsanti azione
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, pady=10)
        
        process_button = ttk.Button(
            buttons_frame, 
            text="Elabora PDF", 
            command=self.process_pdfs,
            style="Accent.TButton"
        )
        process_button.pack(side=tk.RIGHT, padx=5)
        
        cancel_button = ttk.Button(
            buttons_frame, 
            text="Annulla", 
            command=self.root.destroy
        )
        cancel_button.pack(side=tk.RIGHT, padx=5)
        
        # Area log
        log_label = ttk.Label(main_frame, text="Log:", style="Heading.TLabel")
        log_label.pack(anchor="w", pady=(20, 5))
        
        log_frame = ttk.Frame(main_frame, padding=2)
        log_frame.pack(fill=tk.BOTH, expand=True)
        
        self.log_text = tk.Text(log_frame, height=10, width=70, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        # Scrollbar per il log
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # Messaggio iniziale nel log
        self.log("Applicazione avviata. Seleziona uno o più file PDF per iniziare.")
    
    def browse_files(self):
        filenames = filedialog.askopenfilenames(
            title="Seleziona file PDF",
            filetypes=(("PDF files", "*.pdf"), ("All files", "*.*"))
        )
        
        if not filenames:
            return
            
        for filename in filenames:
            if filename not in self.file_paths:
                self.file_paths.append(filename)
                
                # Estrai nome e cognome dal nome del file
                basename = os.path.basename(filename)
                name_parts = os.path.splitext(basename)[0].split()
                
                # Cerca la parola "PROVA" nel nome del file
                prova_index = -1
                for i, part in enumerate(name_parts):
                    if part == "PROVA" and i < len(name_parts) - 1:
                        prova_index = i
                        break
                
                display_name = basename
                if prova_index != -1 and len(name_parts) >= 2:
                    nome = name_parts[0]
                    cognome = name_parts[1]
                    numero_prova = name_parts[prova_index + 1]
                    display_name = f"{nome} {cognome} - PROVA {numero_prova}"
                
                self.files_listbox.insert(tk.END, display_name)
        
        self.log(f"Aggiunti {len(filenames)} file alla lista.")
    
    def remove_selected(self):
        selected = self.files_listbox.curselection()
        if not selected:
            return
            
        # Rimuovi gli elementi selezionati in ordine inverso
        for i in sorted(selected, reverse=True):
            del self.file_paths[i]
            self.files_listbox.delete(i)
        
        self.log(f"Rimossi {len(selected)} file dalla lista.")
    
    def clear_files(self):
        count = self.files_listbox.size()
        if count > 0:
            self.files_listbox.delete(0, tk.END)
            self.file_paths = []
            self.log(f"Rimossi tutti i {count} file dalla lista.")
    
    def process_pdfs(self):
        if not self.file_paths:
            messagebox.showerror("Errore", "Seleziona almeno un file PDF")
            return
        
        self.log("\n" + "=" * 50)
        self.log(f"Inizio elaborazione di {len(self.file_paths)} file PDF...")
        self.log("=" * 50)
        
        # Disabilita i pulsanti durante l'elaborazione
        for widget in self.root.winfo_children():
            self.disable_all_buttons(widget)
        
        # Aggiorna l'interfaccia
        self.root.update()
        
        successful = 0
        failed = 0
        
        try:
            # Reindirizza stdout a una funzione di callback
            original_stdout = sys.stdout
            sys.stdout = self.StdoutRedirector(self.log)
            
            # Elabora i PDF uno alla volta
            for i, pdf_path in enumerate(self.file_paths):
                basename = os.path.basename(pdf_path)
                self.log(f"\n[{i+1}/{len(self.file_paths)}] Elaborazione {basename}...")
                
                try:
                    # Elabora il PDF
                    success = pdf_to_sheets.process_pdf(pdf_path)
                    
                    if success:
                        self.log(f"✅ {basename} - Dati importati con successo!")
                        successful += 1
                    else:
                        self.log(f"❌ {basename} - Importazione fallita. Controlla il log per dettagli.")
                        failed += 1
                
                except Exception as e:
                    self.log(f"❌ {basename} - Errore: {str(e)}")
                    failed += 1
            
            # Ripristina stdout
            sys.stdout = original_stdout
            
            # Log riassuntivo
            self.log("\n" + "=" * 50)
            if failed == 0:
                self.log(f"✅ Elaborazione completata con successo! {successful}/{len(self.file_paths)} file importati.")
            else:
                self.log(f"⚠️ Elaborazione completata con {failed} errori.")
                self.log(f"File importati con successo: {successful}/{len(self.file_paths)}")
                self.log(f"File non importati: {failed}/{len(self.file_paths)}")
            self.log("=" * 50)
            
            # Messaggio di completamento
            if failed == 0:
                messagebox.showinfo(
                    "Importazione completata", 
                    f"Importazione completata con successo!\n\n"
                    f"File importati: {successful}/{len(self.file_paths)}"
                )
            else:
                messagebox.showwarning(
                    "Importazione con errori", 
                    f"Importazione completata, ma con errori.\n\n"
                    f"File importati con successo: {successful}/{len(self.file_paths)}\n"
                    f"File non importati: {failed}/{len(self.file_paths)}\n\n"
                    f"Controlla il log per maggiori dettagli."
                )
        
        except Exception as e:
            sys.stdout = original_stdout
            self.log(f"Errore generale: {str(e)}")
            messagebox.showerror("Errore", f"Si è verificato un errore durante l'elaborazione: {str(e)}")
        
        finally:
            # Riabilita i pulsanti
            for widget in self.root.winfo_children():
                self.enable_all_buttons(widget)
    
    def disable_all_buttons(self, widget):
        if isinstance(widget, ttk.Button):
            widget.state(['disabled'])
        for child in widget.winfo_children():
            self.disable_all_buttons(child)
    
    def enable_all_buttons(self, widget):
        if isinstance(widget, ttk.Button):
            widget.state(['!disabled'])
        for child in widget.winfo_children():
            self.enable_all_buttons(child)
    
    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    class StdoutRedirector:
        def __init__(self, callback):
            self.callback = callback
        
        def write(self, message):
            if message.strip():
                self.callback(message.strip())
        
        def flush(self):
            pass

def main():
    root = tk.Tk()
    app = PDFImportApp(root)
    root.mainloop()

if __name__ == "__main__":
    main() 