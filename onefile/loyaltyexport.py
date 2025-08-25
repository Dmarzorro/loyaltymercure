# -*- coding: utf-8 -*-

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
    "rev": "Total Revenue (Net of VAT)",
    "dep": "Departure",                     # data wyjazdu w Loyalty (opcjonalnie)
}
COLS_O = {
    "card": "Card no.",
    "holder": "Cardholder (stamped)",
    "rev_hotel": "Revenue hotel currency",
    "points": "Rewards Points",
    "media": "Earn Media",
    "dep": "Check-out date",                # data wyjazdu w Operations (opcjonalnie)
}

STATUS_ALLOWED = [
    "ZGODNE",
    "INNE_NAZWISKA",
    "ROZNICA_KWOT",
    "ROZNA_LICZBA_TRANSAKCJI",
    "BRAK_W_OPERATIONS",
    "BRAK_W_LOYALTY",
]

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
        raise FileNotFoundError(f"Brak pliku dla słów {keywords} i rozszerzeń {exts}.")
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
    """Czyści spacje; jeśli czysto numeryczne – zamienia na int-string; w innym wypadku uppercase."""
    s = "" if x is None else str(x)
    s = re.sub(r"\s+", "", s).strip()
    try:
        return str(int(float(s))).upper()
    except Exception:
        return s.upper()

def wyciagnij_pmid(z_karty: str) -> str:
    """
    Z numeru karty wycina PMID jako 8 znaków przed ostatnim znakiem.
    Przykład: 30810324975248MC → 4975248M (s[-9:-1]).
    Jeśli długość < 9, zwraca cały ciąg (bez zmian).
    """
    if z_karty is None:
        return ""
    s = str(z_karty).strip().upper()
    if len(s) >= 9:
        return s[-9:-1]
    return s

def przecinek_na_kropke(x) -> str:
    return ("" if x is None else str(x)).replace(",", ".")

def fmt_set(s: Set[str]) -> str:
    return ", ".join(sorted(s)) if s else "—"

def fmt_list(a: List[float]) -> str:
    return ", ".join(f"{v:.2f}" for v in a) if a else "—"

def fmt_list_s(a: List[str]) -> str:
    return ", ".join(a) if a else "—"

def fmt_deltas(a: List[float], b: List[float]) -> str:
    if not a or not b or len(a) != len(b): return "—"
    return ", ".join(f"Δ={abs(x - y):.2f}" for x, y in zip(a, b))

def parse_date_any(x):
    """Próbuje sparsować datę z różnych formatów (teksty, numery EXCEL). Zwraca pandas.Timestamp lub NaT."""
    if x is None or (isinstance(x, float) and pd.isna(x)) or (isinstance(x, str) and not x.strip()):
        return pd.NaT
    # EXCEL numeryczny
    try:
        if isinstance(x, (int, float)) or (isinstance(x, str) and re.fullmatch(r"\d+(\.\d+)?", x.strip())):
            val = float(x)
            return pd.to_datetime(val, origin="1899-12-30", unit="D", errors="coerce")
    except Exception:
        pass
    # Tekst — pozwól pandasowi zgadnąć, dayfirst pomaga przy "1. Aug."
    return pd.to_datetime(str(x).strip().replace("  ", " "),
                          dayfirst=True, infer_datetime_format=True, errors="coerce")

def fmt_date(dt) -> str:
    return "—" if pd.isna(dt) else pd.Timestamp(dt).strftime("%Y-%m-%d")

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
    if c["media"] in df.columns:
        df[c["media"]] = df[c["media"]].astype(str).str.strip().str.upper()
        df = df[df[c["media"]] != "HOTEL LINK"].copy()

    df["karta_norm"] = df[c["card"]].astype(str).apply(normalizuj_numer_karty)
    df["pmid"]       = df["karta_norm"].apply(wyciagnij_pmid)   # <<< klucz porównania
    df["nazwisko"]   = df[c["holder"]].astype(str).apply(wyodrebnij_nazwisko)

    df["ops_kwota_raw"] = df[c["rev_hotel"]].astype(str).apply(przecinek_na_kropke)
    df["ops_kwota"]     = pd.to_numeric(df["ops_kwota_raw"], errors="coerce")

    # data wyjazdu
    if c["dep"] in df.columns:
        df["ops_data"] = df[c["dep"]].apply(parse_date_any)
    else:
        df["ops_data"] = pd.NaT
    df["ops_data_str"] = df["ops_data"].apply(fmt_date)

    # punkty (tylko do FREQ)
    if c["points"] in df.columns:
        df["ops_punkty_raw"] = df[c["points"]].astype(str).apply(przecinek_na_kropke)
        df["ops_punkty"]     = pd.to_numeric(df["ops_punkty_raw"], errors="coerce")
    else:
        df["ops_punkty"]     = df["ops_kwota"].where(df["ops_kwota"].notna(), 0.0)
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
    c = COLS_L
    engine = "xlrd" if str(path).lower().endswith(".xls") else None
    df = pd.read_excel(path, dtype=str, header=12, engine=engine)
    df.columns = [x.strip() for x in df.columns]
    missing = [c[k] for k in ("card","guest","rev") if c[k] not in df.columns]
    if missing:
        raise ValueError(f"W Loyalty brakuje kolumn: {missing}. Znalezione: {list(df.columns)}")
    df = df[[c["card"], c["guest"], c["rev"]] + ([c["dep"]] if c["dep"] in df.columns else [])].copy()

    df["karta_norm"]     = df[c["card"]].astype(str).apply(normalizuj_numer_karty)
    df["pmid"]           = df["karta_norm"].apply(wyciagnij_pmid)  # <<< klucz porównania
    df["gosc_nazwa_raw"] = df[c["guest"]].astype(str).str.strip().str.upper()
    df["gosc_nazwisko"]  = df["gosc_nazwa_raw"].apply(wyodrebnij_nazwisko)
    df["loyal_kwota_raw"]= df[c["rev"]].astype(str).apply(przecinek_na_kropke)
    df["loyal_kwota"]    = pd.to_numeric(df["loyal_kwota_raw"], errors="coerce")
    if c["dep"] in df.columns:
        df["loyal_data"] = df[c["dep"]].apply(parse_date_any)
    else:
        df["loyal_data"] = pd.NaT
    df["loyal_data_str"] = df["loyal_data"].apply(fmt_date)
    return df

# =======================
# Logika porównania (po PMID)
# =======================

def porownaj(lojal_df: pd.DataFrame, ops_df: pd.DataFrame, tolerancja: float = 0.10) -> Dict[str, pd.DataFrame]:
    # przygotuj pary (kwota, data_str) dla zachowania powiązań
    lojal_df = lojal_df.copy()
    ops_df = ops_df.copy()
    lojal_df["pair_loyal"] = lojal_df.apply(
        lambda r: (r["loyal_kwota"], r["loyal_data_str"]) if pd.notna(r["loyal_kwota"]) else None, axis=1
    )
    ops_df["pair_ops"] = ops_df.apply(
        lambda r: (r["ops_kwota"], r["ops_data_str"]) if pd.notna(r["ops_kwota"]) else None, axis=1
    )

    # grupy – TERAZ po PMID (zamiast pełnego numeru karty)
    loj_grp = lojal_df.groupby("pmid").agg(
        loj_pary=("pair_loyal", lambda x: sorted([p for p in x if p is not None], key=lambda t: t[0])),
        loj_nazwiska=("gosc_nazwisko", lambda x: set(s for s in x if s))
    ).reset_index()
    ops_grp = ops_df.groupby("pmid").agg(
        ops_pary=("pair_ops",  lambda x: sorted([p for p in x if p is not None], key=lambda t: t[0])),
        ops_nazwiska=("nazwisko", lambda x: set(s for s in x if s))
    ).reset_index()

    # mapy
    loj_map = {}
    for _, r in loj_grp.iterrows():
        kw, dt = zip(*r["loj_pary"]) if r["loj_pary"] else ([], [])
        loj_map[r["pmid"]] = {"kw": list(kw), "daty": list(dt), "naz": r["loj_nazwiska"]}

    ops_map = {}
    for _, r in ops_grp.iterrows():
        kw, dt = zip(*r["ops_pary"]) if r["ops_pary"] else ([], [])
        ops_map[r["pmid"]] = {"kw": list(kw), "daty": list(dt), "naz": r["ops_nazwiska"]}

    wszystkie_ops_nazwiska: Set[str] = set(ops_df["nazwisko"].dropna().astype(str).tolist())
    wszystkie_pmid = sorted(set(loj_map) | set(ops_map))

    # sekcje
    zgodne, niezgodne, inne_naz, roznaliczb, brak_w_ops, ops_brak_w_loyal = [], [], [], [], [], []
    freq_rows, przeglad_rows = [], []

    # porównanie po PMID
    for pmid in wszystkie_pmid:
        L = loj_map.get(pmid); O = ops_map.get(pmid)

        if L is None and O is not None:
            przeglad_rows.append({
                "PMID": pmid,
                "Kwota_Loyalty": "—", "Kwota_Operations": fmt_list(O["kw"]), "Δ": "—",
                "Data_Loyalty": "—", "Data_Operations": fmt_list_s(O["daty"]),
                "Nazwiska_Loyalty": "—", "Nazwiska_Operations": fmt_set(O["naz"]),
                "Status_Auto": "BRAK_W_LOYALTY", "Uwaga": "Brak transakcji w Loyalty."
            })
            ops_brak_w_loyal.append({
                "PMID": pmid,
                "Nazwiska_Operations": fmt_set(O["naz"]),
                "Kwoty_Operations": fmt_list(O["kw"]),
                "Daty_Operations": fmt_list_s(O["daty"])
            })
            continue

        if L is not None and O is None:
            przeglad_rows.append({
                "PMID": pmid,
                "Kwota_Loyalty": fmt_list(L["kw"]), "Kwota_Operations": "—", "Δ": "—",
                "Data_Loyalty": fmt_list_s(L["daty"]), "Data_Operations": "—",
                "Nazwiska_Loyalty": fmt_set(L["naz"]), "Nazwiska_Operations": "—",
                "Status_Auto": "BRAK_W_OPERATIONS", "Uwaga": "Brak transakcji w Operations."
            })
            brak_w_ops.append({
                "PMID": pmid,
                "Nazwiska_Loyalty": fmt_set(L["naz"]),
                "Kwoty_Loyalty": fmt_list(L["kw"]),
                "Daty_Loyalty": fmt_list_s(L["daty"])
            })
            continue

        # pmid w obu
        loj_kwoty, ops_kwoty = L["kw"], O["kw"]
        loj_daty,  ops_daty  = L["daty"], O["daty"]
        loj_naz, ops_naz = L["naz"], O["naz"]
        globalnie_brak_naz = not (loj_naz & wszystkie_ops_nazwiska)

        if len(loj_kwoty) != len(ops_kwoty):
            roznaliczb.append({
                "PMID": pmid,
                "Nazwiska_Loyalty": fmt_set(loj_naz), "Nazwiska_Operations": fmt_set(ops_naz),
                "Kwoty_Loyalty": fmt_list(loj_kwoty),  "Kwoty_Operations": fmt_list(ops_kwoty),
                "Daty_Loyalty": fmt_list_s(loj_daty),  "Daty_Operations": fmt_list_s(ops_daty),
            })
            przeglad_rows.append({
                "PMID": pmid,
                "Kwota_Loyalty": fmt_list(loj_kwoty), "Kwota_Operations": fmt_list(ops_kwoty), "Δ": "—",
                "Data_Loyalty": fmt_list_s(loj_daty),  "Data_Operations": fmt_list_s(ops_daty),
                "Nazwiska_Loyalty": fmt_set(loj_naz), "Nazwiska_Operations": fmt_set(ops_naz),
                "Status_Auto": "ROZNA_LICZBA_TRANSAKCJI",
                "Uwaga": "Nazwisko z Loyalty nie występuje w Operations (globalnie)." if globalnie_brak_naz else "—"
            })
            continue

        roznice = [abs(lv - ov) for lv, ov in zip(loj_kwoty, ops_kwoty)]
        wszystkie_ok = all(d <= tolerancja for d in roznice)

        for lv, ov, dl, do in zip(loj_kwoty, ops_kwoty, loj_daty, ops_daty):
            d = abs(lv - ov)
            if d <= tolerancja:
                if (loj_naz & ops_naz):
                    status = "ZGODNE"
                    uwaga = "Nazwisko z Loyalty nie występuje w Operations (globalnie)." if globalnie_brak_naz else "—"
                else:
                    status = "INNE_NAZWISKA"
                    uwaga = "Nazwisko z Loyalty nie występuje w Operations (globalnie)." if globalnie_brak_naz \
                        else f"Różne nazwiska: Loyalty={fmt_set(loj_naz)} vs Operations={fmt_set(ops_naz)}"
            else:
                status = "ROZNICA_KWOT"
                uwaga = "Nazwisko z Loyalty nie występuje w Operations (globalnie)." if globalnie_brak_naz else "—"

            przeglad_rows.append({
                "PMID": pmid,
                "Kwota_Loyalty": f"{lv:.2f}", "Kwota_Operations": f"{ov:.2f}", "Δ": f"{d:.2f}",
                "Data_Loyalty": dl,           "Data_Operations": do,
                "Nazwiska_Loyalty": fmt_set(loj_naz), "Nazwiska_Operations": fmt_set(ops_naz),
                "Status_Auto": status, "Uwaga": uwaga
            })

        # klasyfikacje listowe
        if wszystkie_ok:
            target = zgodne if (loj_naz & ops_naz) else inne_naz
            target.append({
                "PMID": pmid,
                "Nazwiska_Loyalty": fmt_set(loj_naz), "Nazwiska_Operations": fmt_set(ops_naz),
                "Kwoty_Loyalty": fmt_list(loj_kwoty),  "Kwoty_Operations": fmt_list(ops_kwoty),
                "Daty_Loyalty": fmt_list_s(loj_daty),  "Daty_Operations": fmt_list_s(ops_daty),
                "Różnice_Δ": fmt_deltas(loj_kwoty, ops_kwoty)
            })
        else:
            niezgodne.append({
                "PMID": pmid,
                "Nazwiska_Loyalty": fmt_set(loj_naz), "Nazwiska_Operations": fmt_set(ops_naz),
                "Kwoty_Loyalty": fmt_list(loj_kwoty),  "Kwoty_Operations": fmt_list(ops_kwoty),
                "Daty_Loyalty": fmt_list_s(loj_daty),  "Daty_Operations": fmt_list_s(ops_daty),
                "Różnice_Δ": ", ".join(f"Δ={d:.2f}" for d in roznice)
            })

    # FREQ
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

    # PRZEGLĄD — sort + kolumny do ręcznej pracy
    df_przeglad = pd.DataFrame(przeglad_rows)
    if not df_przeglad.empty:
        def _kat(s): return "OK" if s=="ZGODNE" else "PROBLEM"
        def _prio(s):
            if s in ("ROZNICA_KWOT","ROZNA_LICZBA_TRANSAKCJI","BRAK_W_OPERATIONS","BRAK_W_LOYALTY"): return 1
            if s=="INNE_NAZWISKA": return 2
            return 3
        df_przeglad["Kategoria"] = df_przeglad["Status_Auto"].map(_kat)
        df_przeglad["Priorytet"] = df_przeglad["Status_Auto"].map(_prio)
        df_przeglad["Status_Manual"] = ""
        df_przeglad["Status_Final"]  = df_przeglad["Status_Auto"]  # nadpisywane formułą w Excelu

        order = ["Kategoria","Priorytet","Status_Auto","Status_Manual","Status_Final",
                 "PMID","Nazwiska_Loyalty","Nazwiska_Operations",
                 "Kwota_Loyalty","Kwota_Operations","Δ",
                 "Data_Loyalty","Data_Operations",
                 "Uwaga"]
        df_przeglad = df_przeglad[order].sort_values(["Kategoria","Priorytet","PMID"], ascending=[True,True,True], kind="mergesort")
    else:
        df_przeglad = pd.DataFrame(columns=["Kategoria","Priorytet","Status_Auto","Status_Manual","Status_Final",
                                            "PMID","Nazwiska_Loyalty","Nazwiska_Operations",
                                            "Kwota_Loyalty","Kwota_Operations","Δ",
                                            "Data_Loyalty","Data_Operations",
                                            "Uwaga"])

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
        "99_PRZEGLAD_TRANSAKCJI": "99_PRZEGLAD",
        "CFG": "CFG"
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
    widths = {
        "PMID":22,"Nazwiska_Loyalty":30,"Nazwiska_Operations":30,
        "Kwoty_Loyalty":30,"Kwoty_Operations":30,"Różnice_Δ":22,
        "Daty_Loyalty":30,"Daty_Operations":30,
        "Wiersze":10,"Wiersze_z_punktami":20,
        "Status_Auto":18,"Status_Manual":18,"Status_Final":18,
        "Uwaga":40,"Info":24,"Kategoria":12,"Priorytet":10,"Δ":10,
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
                  "• Zgodność: Δ ≤ 0,10\n"
                  "• Niezgodność: Δ > 0,10\n"
                  "• „KARTA_OK_INNE_NAZWISKA” – zgodność kwot, ale różne nazwiska.\n"
                  "• PRZEGLĄD: Status_Auto (algorytm), Status_Manual (lista), Status_Final (kolor i kategoria).\n"
                  "• Kluczem porównania jest PMID (8 znaków przed ostatnim znakiem numeru karty).",
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
        try: ws_cfg.hide()
        except Exception: pass

        # Pozostałe arkusze
        for name, df in wyniki.items():
            if name in {"00_PODSUMOWANIE"}: continue
            out = df.copy()
            if out.empty and len(out.columns) == 0:
                out = pd.DataFrame({"Info":["(brak wpisów)"]})
            for c in out.columns: out[c] = out[c].astype(str)

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

# =======================
# Główna funkcja
# =======================

def porownaj_punkty_z_kartami():
    root = base_dir()
    try:
        p_ops = znajdz_plik_operations(root)
        p_loy = znajdz_plik_loyalty(root)
    except Exception as e:
        print("❌ Błąd wyszukiwania plików:", e)
        print("W tym samym folderze umieść:")
        print(" • Operations: .csv/.xls/.xlsx ze słowem 'operation/operations'")
        print(" • Loyalty:    .xls/.xlsx ze słowem 'loyalty/loyaltyexport'")
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
