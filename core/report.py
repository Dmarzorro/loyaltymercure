# -*- coding: utf-8 -*-

import re
from typing import Dict
from pathlib import Path
import pandas as pd

from .config import STATUS_ALLOWED


def _colnum_to_excel(n: int) -> str:
    s=""; n+=1
    while n:
        n, r = divmod(n-1,26)
        s = chr(65+r) + s
    return s

def safe_sheet_name(name: str, used: set) -> str:
    PREFER = {
        "00_PODSUMOWANIE": "00_PODSUM",
        "01_ZGODNE_≤0,10": "01_ZGODNE≤0.10",
        "02_NIEZGODNE_>0,10": "02_NIEZGODNE>0.10",
        "03_KARTA_OK_INNE_NAZWISKA": "03_INNE_NAZWISKA",
        "04_RÓŻNA_LICZBA_POZYCJI": "04_ROZNA_LICZBA",
        "05_BRAK_KARTY_W_OPERATIONS": "05_BRAK_W_OPER",
        "06_KARTY_W_OPERATIONS_BRAK_W_LOYALTY": "06_TYLKO_OPER",
        "07_FREQ": "07_FREQ",
        "99_PRZEGLAD_TRANSAKCJI": "99_PRZEGLAD",
        "CFG": "CFG"
    }
    s = PREFER.get(name, name)
    s = re.sub(r'[\[\]\:\*\?\/\\]', '_', s)[:31]
    base, i = s, 2
    while s in used:
        suf = f"~{i}"
        s = (base[:31-len(suf)] + suf)
        i += 1
    used.add(s)
    return s

def _apply_sheet_formatting(wb, ws, df: pd.DataFrame):
    fmt_header = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})
    fmt_wrap   = wb.add_format({"text_wrap": True})
    widths = {
        "PMID":22,"Nazwiska_Loyalty":30,"Nazwiska_Operations":30,
        "Kwoty_Loyalty":30,"Kwoty_Operations":30,"Różnice_Δ":22,
        "Daty_Loyalty":30,"Daty_Operations":30,
        "Wiersze":10,"Wiersze_z_punktami":20,
        "Status_Auto":18,"Status_Manual":18,"Status_Final":18,
        "Uwaga":60,"Info":24,"Kategoria":12,"Priorytet":10,"Δ":10,
        "Kwota_Loyalty":16,"Kwota_Operations":18,
        "Data_Loyalty":16,"Data_Operations":16,
    }
    for j, col in enumerate(df.columns):
        ws.write(0, j, col, fmt_header)
        ws.set_column(j, j, widths.get(col, 24), fmt_wrap)
    if not df.empty:
        ws.autofilter(0, 0, len(df), len(df.columns)-1)
    ws.freeze_panes(1, 0)

def zapisz_do_excela(wyniki: Dict[str, pd.DataFrame], plik: Path):
    import xlsxwriter

    with pd.ExcelWriter(plik, engine="xlsxwriter") as writer:
        wb = writer.book
        used = set()

        # 00_PODSUMOWANIE
        pod = wyniki["00_PODSUMOWANIE"].copy()
        s0 = safe_sheet_name("00_PODSUMOWANIE", used)
        pod.to_excel(writer, sheet_name=s0, index=False)
        ws0 = writer.sheets[s0]
        _apply_sheet_formatting(wb, ws0, pod)
        fmt_title = wb.add_format({"bold": True, "font_size": 14})
        fmt_wrap  = wb.add_format({"text_wrap": True})
        ws0.write(2 + len(pod), 0, "Legenda:", fmt_title)
        ws0.write(3 + len(pod), 0,
                  "• Klucz porównania: PMID (Operations) vs PMID wyprowadzony z numeru karty Loyalty.\n"
                  "• Zgodność: Δ ≤ 0,10; Niezgodność: Δ > 0,10.\n"
                  "• „INNE_NAZWISKA” – zgodność kwot, ale różne nazwiska.\n"
                  "• PRZEGLĄД: Status_Auto (algorytm), Status_Manual (lista), Status_Final (kolor i kategoria).",
                  fmt_wrap)

        # CFG – słownik statusów
        cfg_name = safe_sheet_name("CFG", used)
        df_cfg = pd.DataFrame({
            "STATUS": STATUS_ALLOWED,
            "KATEGORIA": ["OK","PROBLEM","PROBLEM","PROBLEM","PROBLEM","PROBLEM"],
            "PRIORYTET": [3,2,1,1,1,1],
        })
        df_cfg.to_excel(writer, sheet_name=cfg_name, index=False)
        ws_cfg = writer.sheets[cfg_name]
        _apply_sheet_formatting(wb, ws_cfg, df_cfg)
        try:
            ws_cfg.hide()
        except Exception:
            pass

        # Pozostałe arkusze
        for name, df in wyniki.items():
            if name in {"00_PODSUMOWANIE"}:
                continue
            out = df.copy()
            if out.empty and len(out.columns) == 0:
                out = pd.DataFrame({"Info":["(brak wpisów)"]})
            for c in out.columns:
                out[c] = out[c].astype(str)

            sname = safe_sheet_name(name, used)
            out.to_excel(writer, sheet_name=sname, index=False)
            ws = writer.sheets[sname]
            _apply_sheet_formatting(wb, ws, out)

            # 99_* – data validation + formuły + CF wg Status_Final
            if sname.startswith("99_") and not out.empty:
                has_manual = "Status_Manual" in out.columns
                has_final  = "Status_Final"  in out.columns
                has_auto   = "Status_Auto"   in out.columns
                has_kat    = "Kategoria"     in out.columns
                has_prio   = "Priorytet"     in out.columns

                if has_manual and has_final and has_auto:
                    r1, rN = 2, len(out) + 1
                    c_manual = out.columns.get_loc("Status_Manual")
                    c_final  = out.columns.get_loc("Status_Final")
                    c_auto   = out.columns.get_loc("Status_Auto")
                    L_manual = _colnum_to_excel(c_manual)
                    L_final  = _colnum_to_excel(c_final)
                    L_auto   = _colnum_to_excel(c_auto)

                    ws.data_validation(r1-1, c_manual, rN-1, c_manual, {
                        "validate": "list",
                        "source": f"={cfg_name}!$A$2:$A${1+len(STATUS_ALLOWED)}"
                    })
                    for rr in range(r1, rN+1):
                        ws.write_formula(rr-1, c_final,
                                         f'=IF(LEN(${L_manual}{rr})>0, ${L_manual}{rr}, ${L_auto}{rr})')
                    if has_kat:
                        c_kat = out.columns.get_loc("Kategoria")
                        for rr in range(r1, rN+1):
                            ws.write_formula(rr-1, c_kat,
                                f'=IFERROR(VLOOKUP(${L_final}{rr}, {cfg_name}!$A$2:$C${1+len(STATUS_ALLOWED)}, 2, FALSE), "INNE")')
                    if has_prio:
                        c_pr = out.columns.get_loc("Priorytet")
                        for rr in range(r1, rN+1):
                            ws.write_formula(rr-1, c_pr,
                                f'=IFERROR(VLOOKUP(${L_final}{rr}, {cfg_name}!$A$2:$C${1+len(STATUS_ALLOWED)}, 3, FALSE), 9)')

                    fmt_green = wb.add_format({"bg_color": "#C6E0B4"})
                    fmt_yel   = wb.add_format({"bg_color": "#FFF2CC"})
                    fmt_red   = wb.add_format({"bg_color": "#F8CBAD"})
                    first_row, last_row = 1, len(out)
                    first_col, last_col = 0, len(out.columns)-1
                    ws.conditional_format(first_row, first_col, last_row, last_col, {
                        "type": "formula", "criteria": f'=${L_final}2="ZGODNE"', "format": fmt_green
                    })
                    ws.conditional_format(first_row, first_col, last_row, last_col, {
                        "type": "formula", "criteria": f'=${L_final}2="INNE_NAZWISKA"', "format": fmt_yel
                    })
                    ws.conditional_format(first_row, first_col, last_row, last_col, {
                        "type": "formula",
                        "criteria": (
                            f'=OR(${L_final}2="ROZNICA_KWOT",'
                            f'${L_final}2="ROZNA_LICZBA_TRANSAKCJI",'
                            f'${L_final}2="BRAK_W_OPERATIONS",'
                            f'${L_final}2="BRAK_W_LOYALTY")'
                        ),
                        "format": fmt_red
                    })

    print(f"✅ Raport zapisany: {plik.name}")
