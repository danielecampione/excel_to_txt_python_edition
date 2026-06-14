"""
main.py
-------
Entry point: avvia la GUI.

Uso:
    python main.py

Dipendenze:
    pip install tkinterdnd2 openpyxl pandas
"""

from gui import App, build_root


def main():
    root = build_root()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
