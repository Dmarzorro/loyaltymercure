# -*- coding: utf-8 -*-
"""
Loyalty vs Operations — porównanie kwot i raport XLSX (wersja uproszczona)

Funkcje:
- Automatyczne znalezienie plików wejściowych w folderze programu:
  • Operations: .csv/.xls/.xlsx z frazą 'operation/operations'
  • Loyalty:    .xls/.xlsx z frazą 'loyalty/loyaltyexport' (nagłówek w 13. wierszu)
- Tolerancja: 0,10 (Δ ≤ 0,10 = zgodne; Δ > 0,10 = niezgodne)
- Arkusze wynikowe:
  00_PODSUMOWANIE, 01_ZGODNE_≤0,10, 02_NIEZGODNE_>0,10,
  03_KARTA_OK_INNE_NAZWISKA, 04_RÓŻNA_LICZBA_POZYCJI,
  05_BRAK_KARTY_W_OPERATIONS, 06_KARTY_W_OPERATIONS_BRAK_W_LOYALTY,
  07_FREQ, 99_PRZEGLAD_TRANSAKCJI
- 99_PRZEGLAD_TRANSAKCJI: sortowanie problemów (❌, ⚠️) nad OK (✓), kolorowanie całych wierszy
- Nazwy arkuszy skracane do ≤31 znaków
- Nazwa pliku wynikowego: cyklicznie 01.xlsx..31.xlsx (nadpisuje najstarszy)
"""

import sys
from pathlib import Path
import re
from typing import List, Dict, Set, Tuple
import pandas as pd

# =======================
# Stałe / Kolumny
# =======================

COLS_L = {
    "card": "Loyalty Card Number",
    "guest": "Guest Name",
    "rev": "Total Revenue (Net of VAT)"
}
COLS_O = {
    "card": "Card no.",
    "holder": "Cardholder (stamped)",
    "rev_hotel": "Revenue hotel currency",
    "points": "Rewards Points",
    "media": "Earn Media",
}

# =======================
# Ścieżki / Wyszukiwanie
# =======================

def base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent

def _find_latest(folder: Path, exts: Tuple[str, ...], keywords: Tuple[str, ...]) -> Path:
    items = []
    for ext in exts:
        for p in folder.glob(f"*{ext}"):
            name = p.name.lower()
            if any(k in name for k in keywords):
                items.append(p)
    if not items:
        raise FileNotFoundError(f"Brak pliku dla wzorca {keywords} i rozszerzeń {exts}.")
    items.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return items[0]

def znajdz_plik_operations(folder: Path) -> Path:
    return _find_latest(folder, (".xlsx", ".xls", ".csv"), ("operation", "operations"))

def znajdz_plik_loyalty(folder: Path) -> Path:
    return _find_latest(folder, (".xlsx", ".xls"), ("loyalty", "loyaltyexport"))

# =======================
# Normalizacja / helpers
# =======================

def wyodrebnij_nazwisko(pelne_imie: str) -> str:
    if pd.isnull(pelne_imie) or not str(pelne_imie).strip():
        return ""
    return str(pelne_imie).strip().split()[-1].upper()

def normalizuj_numer_karty(x) -> str:
    s = "" if x is None else str(x)
    s = re.sub(r"\s+", "", s).strip()
    try:
        return str(int(float(s))).upper()
    except Exception:
        return s.upper()

def przecinek_na_kropke(x) -> str:
    return ("" if x is None else str(x)).replace(",", ".")

def fmt_set(s: Set[str]) -> str:
    return ", ".join(sorted(s)) if s else "—"

def fmt_list(a: List[float]) -> str:
    return ", ".join(f"{v:.2f}" for v in a) if a else "—"

def fmt_deltas(a: List[float], b: List[float]) -> str:
    if not a or not b or len(a) != len(b): return "—"
    return ", ".join(f"Δ={abs(x - y):.2f}" for x, y in zip(a, b))

# =======================
# Czytanie Operations (CSV/Excel)
# =======================

def _wykryj_csv_header_sep(path: str) -> Tuple[int, str, str]:
    encs = ["utf-8-sig", "cp1250", "latin-1"]
    keys = ["card no", "revenue hotel currency", "cardholder (stamped)"]
    for enc in encs:
        with open(path, "r", encoding=enc, errors="ignore") as f:
            lines = [next(f, "") for _ in range(120)]
        lines = [l for l in lines if l]
        if not lines: continue
        hdr = None
        for i, l in enumerate(lines):
            low = l.lower()
            if any(k in low for k in keys):
                hdr = i; break
        if hdr is None:
            for i, l in enumerate(lines):
                if max(l.count(";"), l.count(","), l.count("\t")) >= 5:
                    hdr = i; break
        if hdr is None: continue
        sep = max([";", "\t", ","], key=lambda c: lines[hdr].count(c))
        return hdr, enc, sep
    return 0, "utf-8-sig", ";"

def _normalize_ops(df: pd.DataFrame) -> pd.DataFrame:
    c = COLS_O
    # filtr HOTEL LINK
    if c["media"] in df.columns:
        df[c["media"]] = df[c["media"]].astype(str).str.strip().str.upper()
        df = df[df[c["media"]] != "HOTEL LINK"].copy()
    # normy
    df["karta_norm"] = df[c["card"]].astype(str).apply(normalizuj_numer_karty)
    df["nazwisko"] = df[c["holder"]].astype(str).apply(wyodrebnij_nazwisko)
    df["ops_kwota_raw"] = df[c["rev_hotel"]].astype(str).apply(przecinek_na_kropke)
    df["ops_kwota"] = pd.to_numeric(df["ops_kwota_raw"], errors="coerce")
    if c["points"] in df.columns:
        df["ops_punkty_raw"] = df[c["points"]].astype(str).apply(przecinek_na_kropke)
        df["ops_punkty"] = pd.to_numeric(df["ops_punkty_raw"], errors="coerce")
    else:
        df["ops_punkty"] = df["ops_kwota"].where(df["ops_kwota"].notna(), 0.0)
    return df

def wczytaj_operations(path: str) -> pd.DataFrame:
    c = COLS_O
    if str(path).lower().endswith(".csv"):
        hdr, enc, sep = _wykryj_csv_header_sep(path)
        df = pd.read_csv(path, sep=sep, dtype=str, encoding=enc,
                         skiprows=hdr, header=0, on_bad_lines="skip", engine="python")
    else:
        engine = "xlrd" if str(path).lower().endswith(".xls") else None
        df = pd.read_excel(path, dtype=str, header=0, engine=engine)
    df.columns = [x.strip() for x in df.columns]
    missing = [c[k] for k in ("card","holder","rev_hotel") if c[k] not in df.columns]
    if missing:
        raise ValueError(f"Brak wymaganych kolumn w Operations: {missing}. Znalezione: {list(df.columns)}")
    return _normalize_ops(df)

# =======================
# Czytanie Loyalty
# =======================

def wczytaj_loyalty(path: str) -> pd.DataFrame:
    engine = "xlrd" if str(path).lower().endswith(".xls") else None
    df = pd.read_excel(path, dtype=str, header=12, engine=engine)
    df.columns = [x.strip() for x in df.columns]
    missing = [COLS_L[k] for k in ("card","guest","rev") if COLS_L[k] not in df.columns]
    if missing:
        raise ValueError(f"W Loyalty brakuje kolumn: {missing}. Znalezione: {list(df.columns)}")
    df = df[[COLS_L["card"], COLS_L["guest"], COLS_L["rev"]]].copy()
    df["karta_norm"] = df[COLS_L["card"]].astype(str).apply(normalizuj_numer_karty)
    df["gosc_nazwa_raw"] = df[COLS_L["guest"]].astype(str).str.strip().str.upper()
    df["gosc_nazwisko"] = df["gosc_nazwa_raw"].apply(wyodrebnij_nazwisko)
    df["loyal_kwota_raw"] = df[COLS_L["rev"]].astype(str).apply(przecinek_na_kropke)
    df["loyal_kwota"] = pd.to_numeric(df["loyal_kwota_raw"], errors="coerce")
    return df

# =======================
# Logika porównania
# =======================

def porownaj(lojal_df: pd.DataFrame, ops_df: pd.DataFrame, tolerancja: float = 0.10) -> Dict[str, pd.DataFrame]:
    # grupy
    loj_grp = lojal_df.groupby("karta_norm").agg(
        loj_kwoty=("loyal_kwota", lambda x: sorted([v for v in x if pd.notna(v)])),
        loj_nazwiska=("gosc_nazwisko", lambda x: set(s for s in x if s))
    ).reset_index()
    ops_grp = ops_df.groupby("karta_norm").agg(
        ops_kwoty=("ops_kwota", lambda x: sorted([v for v in x if pd.notna(v)])),
        ops_nazwiska=("nazwisko", lambda x: set(s for s in x if s))
    ).reset_index()
    loj_map = {r["karta_norm"]: {"kw": r["loj_kwoty"], "naz": r["loj_nazwiska"]} for _, r in loj_grp.iterrows()}
    ops_map = {r["karta_norm"]: {"kw": r["ops_kwoty"], "naz": r["ops_nazwiska"]} for _, r in ops_grp.iterrows()}

    wszystkie_ops_nazwiska: Set[str] = set(ops_df["nazwisko"].dropna().astype(str).tolist())
    wszystkie_karty = sorted(set(loj_map) | set(ops_map))

    # sekcje
    zgodne, niezgodne, inne_naz, roznaliczb, brak_w_ops, ops_brak_w_loyal = [], [], [], [], [], []
    freq_rows, przeglad_rows = [], []

    # porównanie
    for karta in wszystkie_karty:
        L = loj_map.get(karta); O = ops_map.get(karta)
        if L is None and O is not None:
            przeglad_rows.append({"Karta": karta, "Kwota_Loyalty": "—", "Kwota_Operations": fmt_list(O["kw"]),
                                  "Δ": "—", "Nazwiska_Loyalty": "—", "Nazwiska_Operations": fmt_set(O["naz"]),
                                  "Status": "❌ BRAK KARTY W LOYALTY", "Uwaga": "Brak transakcji w Loyalty."})
            ops_brak_w_loyal.append({"Karta": karta, "Nazwiska_Operations": fmt_set(O["naz"]),
                                     "Kwoty_Operations": fmt_list(O["kw"])})
            continue
        if L is not None and O is None:
            przeglad_rows.append({"Karta": karta, "Kwota_Loyalty": fmt_list(L["kw"]), "Kwota_Operations": "—",
                                  "Δ": "—", "Nazwiska_Loyalty": fmt_set(L["naz"]), "Nazwiska_Operations": "—",
                                  "Status": "❌ BRAK KARTY W OPERATIONS", "Uwaga": "Brak transakcji w Operations."})
            brak_w_ops.append({"Karta": karta, "Nazwiska_Loyalty": fmt_set(L["naz"]),
                               "Kwoty_Loyalty": fmt_list(L["kw"])})
            continue

        # karta w obu
        loj_kwoty, ops_kwoty = L["kw"], O["kw"]
        loj_naz, ops_naz = L["naz"], O["naz"]
        globalnie_brak_naz = not (loj_naz & wszystkie_ops_nazwiska)

        if len(loj_kwoty) != len(ops_kwoty):
            roznaliczb.append({"Karta": karta, "Nazwiska_Loyalty": fmt_set(loj_naz),
                               "Nazwiska_Operations": fmt_set(ops_naz),
                               "Kwoty_Loyalty": fmt_list(loj_kwoty), "Kwoty_Operations": fmt_list(ops_kwoty)})
            przeglad_rows.append({"Karta": karta, "Kwota_Loyalty": fmt_list(loj_kwoty),
                                  "Kwota_Operations": fmt_list(ops_kwoty), "Δ": "—",
                                  "Nazwiska_Loyalty": fmt_set(loj_naz), "Nazwiska_Operations": fmt_set(ops_naz),
                                  "Status": "❌ RÓŻNA LICZBA TRANSAKCJI",
                                  "Uwaga": "Nazwisko z Loyalty nie występuje w Operations (globalnie)." if globalnie_brak_naz else "—"})
            continue

        roznice = [abs(lv - ov) for lv, ov in zip(loj_kwoty, ops_kwoty)]
        wszystkie_ok = all(d <= tolerancja for d in roznice)

        for lv, ov in zip(loj_kwoty, ops_kwoty):
            d = abs(lv - ov)
            if d <= tolerancja:
                if (loj_naz & ops_naz):
                    status, uwaga = "✓ ZGODNE", ("Nazwisko z Loyalty nie występuje w Operations (globalnie)." if globalnie_brak_naz else "—")
                else:
                    status = "⚠️ INNE NAZWISKA"
                    uwaga = ("Nazwisko z Loyalty nie występuje w Operations (globalnie)." if globalnie_brak_naz
                             else f"Różne nazwiska: Loyalty={fmt_set(loj_naz)} vs Operations={fmt_set(ops_naz)}")
            else:
                status, uwaga = "❌ RÓŻNICA KWOT", ("Nazwisko z Loyalty nie występuje w Operations (globalnie)." if globalnie_brak_naz else "—")

            przeglad_rows.append({
                "Karta": karta,
                "Kwota_Loyalty": f"{lv:.2f}",
                "Kwota_Operations": f"{ov:.2f}",
                "Δ": f"{d:.2f}",
                "Nazwiska_Loyalty": fmt_set(loj_naz),
                "Nazwiska_Operations": fmt_set(ops_naz),
                "Status": status,
                "Uwaga": uwaga
            })

        if wszystkie_ok:
            target = zgodne if (loj_naz & ops_naz) else inne_naz
            target.append({
                "Karta": karta,
                "Nazwiska_Loyalty": fmt_set(loj_naz),
                "Nazwiska_Operations": fmt_set(ops_naz),
                "Kwoty_Loyalty": fmt_list(loj_kwoty),
                "Kwoty_Operations": fmt_list(ops_kwoty),
                "Różnice_Δ": fmt_deltas(loj_kwoty, ops_kwoty)
            })
        else:
            niezgodne.append({
                "Karta": karta,
                "Nazwiska_Loyalty": fmt_set(loj_naz),
                "Nazwiska_Operations": fmt_set(ops_naz),
                "Kwoty_Loyalty": fmt_list(loj_kwoty),
                "Kwoty_Operations": fmt_list(ops_kwoty),
                "Różnice_Δ": ", ".join(f"Δ={d:.2f}" for d in roznice)
            })

    # FREQ (częstotliwość nazwisk)
    ops_tmp = ops_df.copy()
    ops_tmp["ma_punkty"] = ops_tmp["ops_punkty"].fillna(0) > 0
    freq = ops_tmp.groupby("nazwisko").agg(Wiersze=("nazwisko","size"), Wiersze_z_punktami=("ma_punkty","sum")).reset_index()
    for _, r in freq.iterrows():
        nazw, rows, zpkt = r["nazwisko"] or "—", int(r["Wiersze"]), int(r["Wiersze_z_punktami"])
        if rows <= 2: continue
        if rows == 3 and zpkt == 2: status, uw = "OK", "3 wpisy, punkty za 2 — dozwolone."
        elif zpkt >= rows:          status, uw = "OSTRZEŻENIE", "Punkty za wszystkie — możliwe duplikaty."
        else:                       status, uw = "INFO", "Inny przypadek — do weryfikacji."
        freq_rows.append({"Nazwisko": nazw, "Wiersze": rows, "Wiersze_z_punktami": zpkt, "Status": status, "Uwagi": uw})

    # PRZEGLĄD — sort po ważności
    df_przeglad = pd.DataFrame(przeglad_rows)
    if not df_przeglad.empty:
        ikona = df_przeglad["Status"].astype(str).str[0]
        kat_map = {"❌": ("PROBLEM", 1), "⚠": ("PROBLEM", 2), "✓": ("OK", 3)}
        df_przeglad["Kategoria"] = ikona.map(lambda x: kat_map.get(x, ("INNE", 9))[0])
        df_przeglad["Priorytet"] = ikona.map(lambda x: kat_map.get(x, ("INNE", 9))[1])
        order = ["Kategoria","Priorytet","Status","Karta","Nazwiska_Loyalty","Nazwiska_Operations","Kwota_Loyalty","Kwota_Operations","Δ","Uwaga"]
        for c in order:
            if c not in df_przeglad.columns: df_przeglad[c] = ""
        df_przeglad = df_przeglad[order].sort_values(["Kategoria","Priorytet","Karta"], ascending=[True,True,True], kind="mergesort")
    else:
        df_przeglad = pd.DataFrame(columns=["Kategoria","Priorytet","Status","Karta","Nazwiska_Loyalty","Nazwiska_Operations","Kwota_Loyalty","Kwota_Operations","Δ","Uwaga"])

    wyniki = {
        "00_PODSUMOWANIE": pd.DataFrame([
            {"Sekcja":"01_ZGODNE_≤0,10","Wierszy":len(zgodne)},
            {"Sekcja":"02_NIEZGODNE_>0,10","Wierszy":len(niezgodne)},
            {"Sekcja":"03_KARTA_OK_INNE_NAZWISKA","Wierszy":len(inne_naz)},
            {"Sekcja":"04_RÓŻNA_LICZBA_POZYCJI","Wierszy":len(roznaliczb)},
            {"Sekcja":"05_BRAK_KARTY_W_OPERATIONS","Wierszy":len(brak_w_ops)},
            {"Sekcja":"06_KARTY_W_OPERATIONS_BRAK_W_LOYALTY","Wierszy":len(ops_brak_w_loyal)},
            {"Sekcja":"07_FREQ","Wierszy":len(freq_rows)},
            {"Sekcja":"99_PRZEGLAD_TRANSAKCJI","Wierszy":len(df_przeglad)},
        ]),
        "01_ZGODNE_≤0,10": pd.DataFrame(zgodne),
        "02_NIEZGODNE_>0,10": pd.DataFrame(niezgodne),
        "03_KARTA_OK_INNE_NAZWISKA": pd.DataFrame(inne_naz),
        "04_RÓŻNA_LICZBA_POZYCJI": pd.DataFrame(roznaliczb),
        "05_BRAK_KARTY_W_OPERATIONS": pd.DataFrame(brak_w_ops),
        "06_KARTY_W_OPERATIONS_BRAK_W_LOYALTY": pd.DataFrame(ops_brak_w_loyal),
        "07_FREQ": pd.DataFrame(freq_rows),
        "99_PRZEGLAD_TRANSAKCJI": df_przeglad,
    }
    return wyniki

# =======================
# Excel: arkusze / formaty
# =======================

def safe_sheet_name(name: str, used: set) -> str:
    PREFER = {
        "00_PODSUMOWANIE": "00_PODSUM",
        "01_ZGODNE_≤0,10": "01_ZGODNOŚĆ_≤0.10",
        "02_NIEZGODNE_>0,10": "02_NIEZGODNOŚĆ_>0.10",
        "03_KARTA_OK_INNE_NAZWISKA": "03_INNE_NAZWISKA",
        "04_RÓŻNA_LICZBA_POZYCJI": "04_ROZNA_LICZBA",
        "05_BRAK_KARTY_W_OPERATIONS": "05_BRAK_W_OPER",
        "06_KARTY_W_OPERATIONS_BRAK_W_LOYALTY": "06_TYLKO_OPER",
        "07_FREQ": "07_FREQ",
        "99_PRZEGLAD_TRANSAKCJI": "99_CAŁOŚĆ",
    }
    s = PREFER.get(name, name)
    s = re.sub(r'[\[\]\:\*\?\/\\]', '_', s)[:31]
    base, i = s, 2
    while s in used:
        suf = f"~{i}"; s = (base[:31-len(suf)] + suf); i += 1
    used.add(s); return s

def wybierz_sciezke_wyjsciowa(folder: Path, limit: int = 31, ext: str = ".xlsx") -> Path:
    candidates = [folder / f"{i:02d}{ext}" for i in range(1, limit+1)]
    for p in candidates:
        if not p.exists():
            return p
    return min(candidates, key=lambda x: x.stat().st_mtime)

def _colnum_to_excel(n: int) -> str:
    s=""; n+=1
    while n: n, r = divmod(n-1,26); s = chr(65+r)+s
    return s

def _apply_sheet_formatting(wb, ws, df: pd.DataFrame):
    fmt_header = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})
    fmt_wrap   = wb.add_format({"text_wrap": True})
    for j, col in enumerate(df.columns):
        ws.write(0, j, col, fmt_header)
        ws.set_column(j, j, {
            "Karta":22,"Nazwiska_Loyalty":30,"Nazwiska_Operations":30,
            "Kwoty_Loyalty":30,"Kwoty_Operations":30,"Różnice_Δ":22,
            "Wiersze":10,"Wiersze_z_punktami":20,"Status":22,"Uwaga":40,
            "Info":24,"Kategoria":12,"Priorytet":10,"Δ":10,
            "Kwota_Loyalty":16,"Kwota_Operations":18,
        }.get(col, 24), fmt_wrap)
    if not df.empty:
        ws.autofilter(0, 0, len(df), len(df.columns)-1)
    ws.freeze_panes(1, 0)

def _conditional_color_review(wb, ws, df: pd.DataFrame):
    if "Status" not in df.columns or df.empty: return
    fmt_red   = wb.add_format({"bg_color": "#F8CBAD"})
    fmt_yel   = wb.add_format({"bg_color": "#FFF2CC"})
    fmt_green = wb.add_format({"bg_color": "#C6E0B4"})
    n_rows, n_cols = len(df), len(df.columns)
    status_idx = df.columns.get_loc("Status")
    colL = _colnum_to_excel(status_idx)
    fr, lr, fc, lc = 1, n_rows, 0, n_cols-1
    ws.conditional_format(fr, fc, lr, lc, {"type":"formula","criteria":f'=LEFT(${colL}2,1)="❌"',"format":fmt_red})
    ws.conditional_format(fr, fc, lr, lc, {"type":"formula","criteria":f'=LEFT(${colL}2,1)="⚠"',"format":fmt_yel})
    ws.conditional_format(fr, fc, lr, lc, {"type":"formula","criteria":f'=LEFT(${colL}2,1)="✓"',"format":fmt_green})

def zapisz_do_excela(wyniki: Dict[str, pd.DataFrame], plik: Path):
    with pd.ExcelWriter(plik, engine="xlsxwriter") as writer:
        wb = writer.book
        used = set()

        # 00_PODSUMOWANIE + legenda
        pod = wyniki["00_PODSUMOWANIE"].copy()
        s0 = safe_sheet_name("00_PODSUMOWANIE", used)
        pod.to_excel(writer, sheet_name=s0, index=False)
        ws0 = writer.sheets[s0]
        _apply_sheet_formatting(wb, ws0, pod)
        fmt_title = wb.add_format({"bold": True, "font_size": 14})
        fmt_wrap  = wb.add_format({"text_wrap": True})
        ws0.write(2 + len(pod), 0, "Legenda:", fmt_title)
        ws0.write(3 + len(pod), 0,
                  "• Zgodność: Δ ≤ 0,10\n"
                  "• Niezgodność: Δ > 0,10\n"
                  "• „KARTA_OK_INNE_NAZWISKA” – karta zgodna, ale nazwiska różne.\n"
                  "• PRZEGLĄD: ✓ zgodne, ⚠️ inne nazwiska, ❌ różnice/braki kart.", fmt_wrap)

        # reszta arkuszy
        for name, df in wyniki.items():
            if name == "00_PODSUMOWANIE": continue
            out = df.copy()
            if out.empty and len(out.columns) == 0:
                out = pd.DataFrame({"Info":["(brak wpisów)"]})
            for c in out.columns: out[c] = out[c].astype(str)
            sname = safe_sheet_name(name, used)
            out.to_excel(writer, sheet_name=sname, index=False)
            ws = writer.sheets[sname]
            _apply_sheet_formatting(wb, ws, out)
            if sname in {"99_PRZEGLAD","99_CAŁOŚĆ","99_PRZEGLAD_TRANSAKCJI"}:
                _conditional_color_review(wb, ws, out)
    print(f"✅ Raport zapisany: {plik.name}")

# =======================
# Główna funkcja
# =======================

def wybierz_sciezke_wyjsciowa(folder: Path, limit: int = 31, ext: str = ".xlsx") -> Path:
    files = [folder / f"{i:02d}{ext}" for i in range(1, limit+1)]
    free = next((p for p in files if not p.exists()), None)
    return free or min(files, key=lambda x: x.stat().st_mtime)

def porownaj_punkty_z_kartami():
    root = base_dir()
    try:
        p_ops = znajdz_plik_operations(root)
        p_loy = znajdz_plik_loyalty(root)
    except Exception as e:
        print("❌ Błąd wyszukiwania plików:", e)
        print("W tym samym folderze umieść:")
        print(" • Operations: .csv/.xls/.xlsx z frazą 'operation/operations'")
        print(" • Loyalty:    .xls/.xlsx z frazą 'loyalty/loyaltyexport'")
        return

    print(f"🔎 Operations: {p_ops.name}")
    print(f"🔎 Loyalty:    {p_loy.name}")

    lojal_df = wczytaj_loyalty(str(p_loy))
    ops_df   = wczytaj_operations(str(p_ops))

    wyniki = porownaj(lojal_df, ops_df, tolerancja=0.10)

    output = wybierz_sciezke_wyjsciowa(root)
    if output.exists():
        print(f"ℹ️  Uwaga: {output.name} zostanie nadpisany (najstarszy w cyklu 01..31).")
    zapisz_do_excela(wyniki, output)
    print("\n✅ Gotowe. Otwórz plik:", output.name)

if __name__ == "__main__":
    porownaj_punkty_z_kartami()
