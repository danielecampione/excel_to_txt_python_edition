"""Interfaccia Tk completa con drag-and-drop, selezione file e join opzionale."""
from __future__ import annotations

import os
import threading
import tkinter as tk
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText

from logic import consume_warnings, process_excel, save_output

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False


class App:
    BG = "#f8fafc"; PANEL = "#f1f5f9"; ACCENT = "#3b82f6"; ACCENT_DARK = "#1e40af"
    FG = "#475569"; DARK = "#0f172a"; SUCCESS = "#16a34a"; ERROR = "#dc2626"
    WARNING = "#d97706"; BORDER = "#cbd5e1"

    def __init__(self, root: tk.Tk):
        self.root = root
        self.current_file: str | None = None
        self.secondary_file: str | None = None
        root.title("Excel -> TXT")
        root.configure(bg=self.BG)
        root.minsize(690, 700)
        self._build_ui()

    def _build_ui(self):
        outer = tk.Frame(self.root, bg=self.BG)
        outer.pack(fill="both", expand=True)
        canvas = tk.Canvas(outer, bg=self.BG, highlightthickness=0)
        scrollbar = tk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        self.content = tk.Frame(canvas, bg=self.BG, padx=24, pady=18)
        window = canvas.create_window((0, 0), window=self.content, anchor="n")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True); scrollbar.pack(side="right", fill="y")
        self.content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: (canvas.itemconfigure(window, width=min(e.width, 658)), canvas.coords(window, e.width / 2, 0)))
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))

        tk.Label(self.content, text="Excel -> TXT", font=("Segoe UI", 22, "bold"), bg=self.BG, fg=self.ACCENT_DARK).pack(pady=(0, 12))
        self._title("Documento principale")
        self.main_drop = self._drop_panel("Trascina qui il file Excel principale (.xlsx / .xls)", "oppure clicca per sfogliare", True, 118)
        self.columns = self._area("Colonne principali da estrarre (una per riga, nell'ordine desiderato):", "Numero\nDescrizione\nBreve descrizione", 5)

        join = tk.LabelFrame(self.content, text="Join con documento secondario", font=("Segoe UI", 12, "bold"), bg=self.BG, fg=self.DARK, bd=1, relief="solid", padx=10, pady=10)
        join.pack(fill="x", pady=(12, 12))
        tk.Label(join, text="Opzionale: senza documento secondario l'app continua a funzionare come prima.", font=("Segoe UI", 10), bg=self.PANEL, fg=self.FG, anchor="w").pack(fill="x")
        self.secondary_drop = self._drop_panel("Trascina qui il documento secondario per la join", "ad esempio: sys_user.xlsx", False, 100, join)
        key_box = tk.LabelFrame(
            join, text="1. Chiavi di collegamento", font=("Segoe UI", 10, "bold"),
            bg=self.PANEL, fg=self.ACCENT_DARK, bd=1, relief="solid", padx=8, pady=6
        )
        key_box.pack(fill="x", pady=(10, 8))
        self.main_keys = self._area(
            "Chiavi del DOCUMENTO PRINCIPALE (una per riga):",
            "Richiesto per\nReparto\nUbicazione", 4, key_box
        )
        self.secondary_keys = self._area(
            "Chiavi del DOCUMENTO SECONDARIO (una per riga):",
            "Nome utente\nReparto\nUbicazione", 4, key_box
        )
        self.join_output = self._area(
            "2. Colonne del documento secondario da aggiungere all'output:",
            "Nome utente\nNumero dipendente\nTitolo\nReparto", 5, join
        )
        self.fallback = tk.BooleanVar(value=True)
        tk.Checkbutton(join, text="Abilita join con fallback sull'ultima chiave", variable=self.fallback, bg=self.PANEL, fg=self.FG, activebackground=self.PANEL, font=("Segoe UI", 10)).pack(anchor="w")
        tk.Label(join, text="Le chiavi sono accoppiate riga per riga. Il fallback ignora soltanto l'ultima chiave\ne solo in assenza di ambiguita.", justify="left", bg=self.PANEL, fg=self.FG, font=("Segoe UI", 9)).pack(anchor="w")
        tk.Button(join, text="Rimuovi documento secondario", command=self._clear_secondary, font=("Segoe UI", 9)).pack(anchor="w", pady=(7, 0))

        self.process_btn = tk.Button(self.content, text="Elabora", command=self._on_process, state="disabled", bg=self.ACCENT, fg="white", activebackground=self.ACCENT_DARK, activeforeground="white", relief="flat", font=("Segoe UI", 10, "bold"), padx=24, pady=5)
        self.process_btn.pack(pady=(0, 9))
        status_box = tk.LabelFrame(self.content, text="Stato e warning", font=("Segoe UI", 10, "bold"), bg=self.BG, fg=self.DARK, bd=1, relief="solid")
        status_box.pack(fill="x")
        self.status = tk.Text(status_box, height=7, wrap="word", font=("Segoe UI", 10), bg="white", bd=0, padx=8, pady=7, state="disabled")
        self.status.pack(fill="both", expand=True)
        self._set_status("In attesa del documento principale...", self.FG)
        if not DND_AVAILABLE:
            self._set_status("tkinterdnd2 non trovato: il clic per sfogliare resta disponibile. Installa con: pip install tkinterdnd2", self.WARNING)

    def _title(self, text):
        tk.Label(self.content, text=text, font=("Segoe UI", 13, "bold"), bg=self.BG, fg=self.DARK).pack(fill="x", pady=(0, 6))

    def _area(self, label, value, height, parent=None):
        parent = parent or self.content
        frame = tk.Frame(parent, bg=self.PANEL if parent is not self.content else self.BG)
        frame.pack(fill="x", pady=(8, 0))
        tk.Label(frame, text=label, anchor="w", bg=frame["bg"], fg=self.FG, font=("Segoe UI", 10)).pack(fill="x")
        area = ScrolledText(frame, height=height, wrap="none", font=("Segoe UI", 10), bd=1, relief="solid")
        area.pack(fill="x", pady=(3, 0)); area.insert("1.0", value)
        return area

    def _drop_panel(self, first, second, main, height, parent=None):
        parent = parent or self.content
        bg = self.PANEL if parent is not self.content else self.BG
        frame = tk.Frame(parent, bg=bg, height=height, highlightbackground=self.ACCENT, highlightthickness=2, cursor="hand2")
        frame.pack(fill="x", pady=(0 if main else 8, 0)); frame.pack_propagate(False)
        labels = [tk.Label(frame, text="FILE", bg=bg, fg=self.ACCENT_DARK, font=("Segoe UI", 11, "bold")), tk.Label(frame, text=first, bg=bg, fg=self.FG, font=("Segoe UI", 10)), tk.Label(frame, text=second, bg=bg, fg="#94a3b8", font=("Segoe UI", 9))]
        for label in labels: label.pack(expand=True if label is labels[0] else False)
        for widget in [frame, *labels]: widget.bind("<Button-1>", lambda e, m=main: self._choose(m))
        frame._labels = labels[1:]  # type: ignore[attr-defined]
        if DND_AVAILABLE:
            frame.drop_target_register(DND_FILES)
            frame.dnd_bind("<<Drop>>", lambda e, m=main: self._on_drop(e, m))
        return frame

    @staticmethod
    def _lines(area):
        return [line.strip() for line in area.get("1.0", "end").splitlines() if line.strip()]

    @staticmethod
    def _valid(path):
        return path.lower().endswith((".xlsx", ".xls"))

    def _choose(self, main):
        path = filedialog.askopenfilename(title="Seleziona file Excel", filetypes=[("File Excel", "*.xlsx *.xls")])
        if path: self._handle_file(path, main)

    def _on_drop(self, event, main):
        paths = self.root.tk.splitlist(event.data.strip())
        if paths: self._handle_file(paths[0], main)

    def _handle_file(self, path, main):
        if not self._valid(path):
            if main: self.current_file = None; self.process_btn.config(state="disabled")
            else: self.secondary_file = None
            self._set_status("File non supportato. Usa .xlsx o .xls", self.ERROR); return
        panel = self.main_drop if main else self.secondary_drop
        panel._labels[0].config(text=f"Selezionato: {os.path.basename(path)}")
        panel._labels[1].config(text="Documento principale pronto" if main else "Join attiva")
        if main:
            self.current_file = path; self.process_btn.config(state="normal"); self._set_status("Documento principale pronto.", self.FG)
        else:
            self.secondary_file = path; self._set_status("Documento secondario pronto: la join verra applicata.", self.FG)

    def _clear_secondary(self):
        self.secondary_file = None
        self.secondary_drop._labels[0].config(text="Trascina qui il documento secondario per la join")
        self.secondary_drop._labels[1].config(text="ad esempio: sys_user.xlsx")
        self._set_status("Documento secondario rimosso. Elaborazione standard attiva.", self.FG)

    def _on_process(self):
        if not self.current_file: return
        columns = self._lines(self.columns)
        if not columns: self._set_status("Inserisci almeno una colonna principale.", self.ERROR); return
        args = (self.current_file, columns, self.secondary_file, self._lines(self.main_keys), self._lines(self.secondary_keys), self._lines(self.join_output), self.fallback.get())
        self.process_btn.config(state="disabled"); self._set_status("Elaborazione in corso...", self.FG)
        threading.Thread(target=self._worker, args=args, daemon=True).start()

    def _worker(self, *args):
        try:
            content = process_excel(*args)
            out = save_output(args[0], content)
            warning = consume_warnings()
            self.root.after(0, lambda: self._finish(out, warning))
        except Exception as exc:
            self.root.after(0, lambda e=exc: self._fail(e))

    def _finish(self, out, warning):
        if warning: self._set_status(f"File salvato correttamente:\n{out}\n\n{warning}", self.WARNING)
        else: self._set_status(f"File salvato:\n{out}", self.SUCCESS)
        self.process_btn.config(state="normal")

    def _fail(self, exc):
        self._set_status(f"Errore:\n{exc}", self.ERROR); self.process_btn.config(state="normal")

    def _set_status(self, message, color):
        self.status.config(state="normal", fg=color); self.status.delete("1.0", "end"); self.status.insert("1.0", message); self.status.config(state="disabled"); self.status.see("1.0")


def build_root() -> tk.Tk:
    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    width, height = 720, min(940, root.winfo_screenheight() - 60)
    root.geometry(f"{width}x{height}+{max(0, (root.winfo_screenwidth()-width)//2)}+{max(0, (root.winfo_screenheight()-height)//2)}")
    return root
