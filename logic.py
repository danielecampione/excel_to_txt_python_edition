"""
logic.py
--------
Logica di elaborazione: lettura Excel e scrittura file di testo.
"""

import os
import pandas as pd


def get_first_non_bkp_sheet(xls: pd.ExcelFile) -> str:
    for name in xls.sheet_names:
        if not name.lower().endswith("bkp"):
            return name
    raise ValueError("Tutti i fogli hanno il suffisso 'bkp': nessun foglio utilizzabile.")


def process_excel(path: str, columns_to_extract: list) -> str:
    path = path.strip()
    xls = pd.ExcelFile(path)
    sheet_name = get_first_non_bkp_sheet(xls)

    df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
    # Pulizia nomi colonne
    df.columns = [c.strip() for c in df.columns]

    # Verifica che le colonne richieste dalla textarea esistano nel foglio
    missing = set(columns_to_extract) - set(df.columns)
    if missing:
        raise ValueError(f"Colonne mancanti nel foglio '{sheet_name}': {missing}")

    blocks = []
    for _, row in df.iterrows():
        def cell(col):
            v = row[col]
            return str(v).strip() if pd.notna(v) else ""

        # Estraiamo dinamicamente i valori per ogni colonna richiesta
        parts = [cell(col) for col in columns_to_extract]
        
        # Uniamo le parti del record con il doppio a capo e aggiungiamo un a capo alla fine
        block = "\n\n".join(parts) + "\n"
        blocks.append(block)

    # Uniamo i record con il separatore
    return "\n---\n\n".join(blocks) + "\n---\n"


def save_output(excel_path: str, content: str) -> str:
    base = os.path.splitext(excel_path)[0]
    out_path = base + "_output.txt"

    try:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(content)
    except OSError:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        out_path = os.path.join(script_dir, os.path.basename(base) + "_output.txt")
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(content)

    return out_path