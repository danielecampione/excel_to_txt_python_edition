"""
gui.py
------
Interfaccia grafica con supporto drag-and-drop e textarea per colonne.
Dipende da logic.py per l'elaborazione dei file.
"""

import os
import tkinter as tk
from logic import process_excel, save_output

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False


class App:
    BG           = "#f8fafc" 
    ACCENT       = "#3b82f6" 
    ACCENT_LIGHT = "#1e40af" 
    FG           = "#0f172a" 
    FG_DIM       = "#475569" 
    SUCCESS      = "#16a34a" 
    ERROR        = "#dc2626" 

    def __init__(self, root: tk.Tk):
        self.root = root
        self.current_file = None 
        
        root.title("Excel → TXT")
        root.resizable(False, False)
        root.configure(bg=self.BG)
        self._build_ui()

        if DND_AVAILABLE:
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind("<<Drop>>", self._on_drop)
        else:
            self._set_status(
                "⚠ tkinterdnd2 non trovato.\nInstalla con: pip install tkinterdnd2",
                self.ERROR,
            )

    # ── costruzione UI ───────────────────────────────────────────────────────

    def _build_ui(self):
        # Titolo
        tk.Label(
            self.root, text="Excel → TXT",
            font=("Segoe UI", 18, "bold"),
            bg=self.BG, fg=self.ACCENT_LIGHT,
        ).pack(padx=24, pady=(15, 5))

        # Drag & Drop Frame
        self.drop_frame = tk.Frame(
            self.root,
            bg=self.BG,
            highlightbackground=self.ACCENT,
            highlightthickness=2,
            width=420, height=110,
        )
        self.drop_frame.pack(padx=24, pady=(0, 10))
        self.drop_frame.pack_propagate(False)

        tk.Label(
            self.drop_frame, text="📂",
            font=("Segoe UI", 28),
            bg=self.BG,
        ).pack(pady=(10, 0))

        tk.Label(
            self.drop_frame,
            text="Trascina qui il file Excel (.xlsx)",
            font=("Segoe UI", 10),
            bg=self.BG, fg=self.FG_DIM,
        ).pack()

        # Sezione Textarea per le colonne
        tk.Label(
            self.root,
            text="Colonne da estrarre (una per riga, in ordine):",
            font=("Segoe UI", 9, "bold"),
            bg=self.BG, fg=self.FG_DIM
        ).pack(padx=24, anchor="w")

        self.columns_text = tk.Text(
            self.root,
            height=4, width=45,
            font=("Segoe UI", 10),
            bg="white", fg=self.FG,
            bd=1, relief="solid"
        )
        self.columns_text.pack(padx=24, pady=(2, 10))
        
        # Popoliamo la textarea con i valori di default
        default_cols = "Numero\nDescrizione\nBreve descrizione"
        self.columns_text.insert(tk.END, default_cols)

        # Bottone Elabora
        self.process_btn = tk.Button(
            self.root,
            text="Elabora",
            font=("Segoe UI", 10, "bold"),
            bg=self.ACCENT, fg="white",
            activebackground=self.ACCENT_LIGHT, activeforeground="white",
            state=tk.DISABLED, 
            command=self._on_process,
            relief=tk.FLAT,
            padx=20, pady=4
        )
        self.process_btn.pack(pady=(5, 5))

        # Status Label
        self.status_var = tk.StringVar(value="In attesa di un file…")
        self.status_lbl = tk.Label(
            self.root,
            textvariable=self.status_var,
            font=("Segoe UI", 9),
            bg=self.BG, fg=self.FG_DIM,
            wraplength=400, justify="center",
        )
        self.status_lbl.pack(padx=24, pady=(0, 10))

    # ── eventi ───────────────────────────────────────────────────────────────

    def _on_drop(self, event):
        paths = self.root.tk.splitlist(event.data.strip())
        if not paths:
            return

        file_path = paths[0]
        if not file_path.lower().endswith((".xlsx", ".xls")):
            self._set_status("❌ File non supportato.\nTrascina un file .xlsx", self.ERROR)
            self.current_file = None
            self.process_btn.config(state=tk.DISABLED)
            return

        self.current_file = file_path
        filename = os.path.basename(file_path)
        self._set_status(f"📄 Selezionato: {filename}\nPronto per l'elaborazione.", self.FG_DIM)
        self.process_btn.config(state=tk.NORMAL)

    def _on_process(self):
        if not self.current_file:
            return

        # Leggiamo il contenuto della textarea
        raw_text = self.columns_text.get("1.0", tk.END)
        # Dividiamo per riga e rimuoviamo spazi/righe vuote
        columns_to_extract = [line.strip() for line in raw_text.split("\n") if line.strip()]

        if not columns_to_extract:
            self._set_status("❌ Errore: Inserisci almeno il nome di una colonna.", self.ERROR)
            return

        self._set_status("⏳ Elaborazione in corso…", self.FG_DIM)
        self.process_btn.config(state=tk.DISABLED) 
        self.root.update()

        try:
            # Passiamo le colonne indicate alla logica
            content = process_excel(self.current_file, columns_to_extract)
            out = save_output(self.current_file, content)
            self._set_status(f"✅ File salvato:\n{out}", self.SUCCESS)
        except Exception as exc:
            self._set_status(f"❌ Errore:\n{exc}", self.ERROR)
        finally:
            self.current_file = None
            # Rimettiamo il bottone in stato disabilitato finché non si inserisce un nuovo file
            self.process_btn.config(state=tk.DISABLED)

    def _set_status(self, msg: str, color: str = None):
        self.status_var.set(msg)
        if color:
            self.status_lbl.configure(fg=color)


# ── factory per il root window ────────────────────────────────────────────────

def build_root() -> tk.Tk:
    if DND_AVAILABLE:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()

    # Finestra ancora un po' più alta per far spazio alla textarea
    w, h = 470, 440
    root.update_idletasks()
    x = (root.winfo_screenwidth() - w) // 2
    y = (root.winfo_screenheight() - h) // 2
    root.geometry(f"{w}x{h}+{x}+{y}")
    return root