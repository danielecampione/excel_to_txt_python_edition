# Excel to TXT - Python Edition v0.03b

Conversione del primo foglio Excel in TXT, con selezione dinamica delle colonne e join opzionale con un documento secondario.

## Funzioni principali
- file `.xlsx` e `.xls`;
- trascinamento oppure selezione tramite finestra;
- colonne principali configurabili;
- join su chiavi accoppiate;
- fallback opzionale senza l'ultima chiave, solo se univoco;
- warning per duplicati, fallback e ambiguita;
- elaborazione in background per non bloccare l'interfaccia;
- output UTF-8 accanto al documento principale.

## Avvio
```bash
pip install -r requirements.txt
python main.py
```

Autore: Daniele Campione  
Licenza: GNU General Public License

## Correzione v0.03b
- area chiavi del documento secondario resa esplicita, separata e a larghezza intera;
- etichette numerate e finestra iniziale piu alta.
