# -*- coding: utf-8 -*-

import sys
import openpyxl  # ważne dla pakowania .xlsx przez PyInstaller
from pathlib import Path

from core.utils import base_dir, znajdz_plik_operations, znajdz_plik_loyalty, wybierz_sciezke_wyjsciowa
from core.io_loyalty import wczytaj_loyalty
from core.io_operations import wczytaj_operations
from core.compare import porownaj
from core.report import zapisz_do_excela

from ui_gui import run_gui


def porownaj_punkty_z_kartami():
    root = base_dir()
    try:
        p_ops = znajdz_plik_operations(root)
        p_loy = znajdz_plik_loyalty(root)
    except Exception as e:
        print("❌ Błąd wyszukiwania plików:", e)
        print("W tym samym folderze umieść:")
        print(" • Operations: .xls/.xlsx ze słowem 'operation/operations' (nagłówki w 3. wierszu)")
        print(" • Loyalty:    .xls/.xlsx ze słowem 'loyalty/loyaltyexport' (nagłówki od 13. wiersza)")
        return

    print(f"🔎 Operations: {p_ops.name}")
    print(f"🔎 Loyalty:    {p_loy.name}")

    lojal_df = wczytaj_loyalty(str(p_loy))
    ops_df   = wczytaj_operations(str(p_ops))
    wyniki   = porownaj(lojal_df, ops_df, tolerancja=0.10)

    output = wybierz_sciezke_wyjsciowa(root)
    zapisz_do_excela(wyniki, output)
    print("\n✅ Gotowe. Otwórz plik:", output.name)


if __name__ == "__main__":
    # GUI jako domyślne; tryb konsolowy uruchomisz przez --cli
    if "--cli" in sys.argv:
        porownaj_punkty_z_kartami()
    else:
        run_gui()
