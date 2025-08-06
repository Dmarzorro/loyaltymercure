# -*- coding: utf-8 -*-
"""
Loyalty vs Operations — porównanie kwot i raport XLSX (wersja bez wrażliwości na nazwy)

Zmiany:
1) W raportach „CSV” -> „Operations”.
2) Pliki wejściowe znajdowane automatycznie po słowach kluczowych w nazwie:
   - Operations: *operation* / *operations* (plik .csv)
   - Loyalty: *loyalty* / *loyaltyexport* (plik .xls / .xlsx)
3) Plik wyjściowy: wynik_porownania_YYYYMMDD.xlsx (data uruchomienia).
4) FREQ złączone do jednego arkusza.
5) Tolerancja główna: 0,10 (Δ ≤ 0,10 = zgodne; Δ > 0,10 = niezgodne).
6) Arkusz 99_PRZEGLAD_TRANSAKCJI z ikonami statusów i uwagami.
7) Nazewnictwo i komunikaty po polsku.
"""

import sys
from pathlib import Path
import re
from typing import List, Dict, Set, Tuple
from datetime import datetime

import pandas as pd


# =======================
# Ścieżki – praca jako .exe
# =======================

def base_dir() -> Path:
    """Folder bazowy: dla .exe = katalog pliku wykonywalnego, dla .py = katalog skryptu."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


# =======================
# Wyszukiwanie plików wejściowych po wzorcu nazw
# =======================

def znajdz_plik_operations(folder: Path) -> Path:
    """
    Szuka najnowszego (po modyfikacji) pliku .csv zawierającego 'operation'/'operations' (case-insensitive).
    Przykład akceptowany: 'REPORT_OPERATION2050803-2020803.csv'
    """
    kandydaci = []
    for p in folder.glob("*.csv"):
        name = p.name.lower()
        if "operation" in name:  # pokrywa 'operation' i 'operations'
            kandydaci.append(p)
    if not kandydaci:
        raise FileNotFoundError(
            "Nie znaleziono pliku Operations (.csv) zawierającego w nazwie 'operation'/'operations'."
        )
    # wybierz najnowszy po mtime
    kandydaci.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return kandydaci[0]


def znajdz_plik_loyalty(folder: Path) -> Path:
    """
    Szuka najnowszego (po modyfikacji) pliku .xls/.xlsx zawierającego 'loyalty'/'loyaltyexport' (case-insensitive).
    Przykłady: 'H3417_LoyaltyExport_202508052.xls', 'loyalty.xlsx'
    """
    kandydaci = []
    for p in list(folder.glob("*.xls")) + list(folder.glob("*.xlsx")):
        name = p.name.lower()
        if ("loyalty" in name) or ("loyaltyexport" in name):
            kandydaci.append(p)
    if not kandydaci:
        raise FileNotFoundError(
            "Nie znaleziono pliku Loyalty (.xls/.xlsx) zawierającego w nazwie 'loyalty'/'loyaltyexport'."
        )
    kandydaci.sort(key=lambda x: x.stat().st_mtime, reverse=True)
    return kandydaci[0]


# =======================
# Pomocnicze / Normalizacja
# =======================

def wyodrebnij_nazwisko(pelne_imie: str) -> str:
    """Zwraca nazwisko (ostatnie słowo) UPPER. Pusty string dla NaN/pustych."""
    if pd.isnull(pelne_imie) or not str(pelne_imie).strip():
        return ""
    czesci = str(pelne_imie).strip().split()
    return czesci[-1].upper() if czesci else ""

def normalizuj_numer_karty(x) -> str:
    """
    Normalizacja numeru karty:
    - usuwa spacje,
    - chroni przed '3.08103E+17',
    - zwraca string w UPPER.
    """
    s = "" if x is None else str(x)
    s = re.sub(r"\s+", "", s).strip()
    try:
        return str(int(float(s))).upper()
    except Exception:
        return s.upper()

def przecinek_na_kropke(x) -> str:
    return ("" if x is None else str(x)).replace(",", ".")

def fmt_set(s: Set[str]) -> str:
    lst = sorted(list(s)) if s else []
    return ", ".join(lst) if lst else "—"

def fmt_list(a: List[float]) -> str:
    if not a:
        return "—"
    return ", ".join(f"{v:.2f}" for v in a)

def fmt_deltas(a: List[float], b: List[float]) -> str:
    if not a or not b or len(a) != len(b):
        return "—"
    return ", ".join(f"Δ={abs(x - y):.2f}" for x, y in zip(a, b))


# =======================
# Wykrywanie nagłówka/sep/encoding dla Operations (CSV)
# =======================

def wykryj_naglowek_i_separator_csv(sciezka: str) -> Tuple[int, str, str]:
    """
    Zwraca (indeks_wiersza_naglowka, encoding, separator).
    Szuka wiersza z nagłówkiem po słowach kluczowych lub liczbie separatorów.
    """
    kandydatury_enc = ["utf-8-sig", "cp1250", "latin-1"]
    slowa_kluczowe = ["card no", "revenue hotel currency", "cardholder (stamped)"]

    for enc in kandydatury_enc:
        with open(sciezka, "r", encoding=enc, errors="ignore") as f:
            linie = []
            for _ in range(120):
                try:
                    linie.append(next(f))
                except StopIteration:
                    break
        if not linie:
            continue

        nag_idx = None
        # 1) słowa kluczowe
        for i, w in enumerate(linie):
            low = w.lower()
            if any(k in low for k in slowa_kluczowe):
                nag_idx = i
                break
        # 2) „bogata w separatory” linia
        if nag_idx is None:
            for i, w in enumerate(linie):
                if w.count(";") >= 5 or w.count(",") >= 5 or w.count("\t") >= 5:
                    nag_idx = i
                    break
        if nag_idx is None:
            continue

        header_line = linie[nag_idx]
        sep = max([";", "\t", ","], key=lambda c: header_line.count(c))
        return nag_idx, enc, sep

    # Fallback
    return 0, "utf-8-sig", ";"


# =======================
# Wejście: Operations (CSV)
# =======================

def wczytaj_operations_csv(sciezka: str) -> pd.DataFrame:
    """
    Czyta plik Operations (.csv) z auto-wykryciem nagłówka/separatora/encoding.
    Wymagane kolumny:
      - Card no.
      - Cardholder (stamped)
      - Revenue hotel currency
    """
    nag_idx, enc, sep = wykryj_naglowek_i_separator_csv(sciezka)

    df = pd.read_csv(
        sciezka,
        sep=sep,
        dtype=str,
        encoding=enc,
        skiprows=nag_idx,
        header=0,
        on_bad_lines="skip",
        engine="python"
    )
    df.columns = [c.strip() for c in df.columns]

    kol_karta = "Card no."
    kol_posiadacz = "Cardholder (stamped)"
    kol_przychod_hotel = "Revenue hotel currency"
    kol_punkty = "Rewards Points"    # używane tylko do reguły częstotliwości; nie sumujemy w raporcie
    kol_media = "Earn Media"

    wymagane = [kol_karta, kol_posiadacz, kol_przychod_hotel]
    for c in wymagane:
        if c not in df.columns:
            raise ValueError(
                f"Brak wymaganej kolumny '{c}' w Operations. "
                f"Znalezione kolumny: {list(df.columns)}. "
                f"(wykryto sep='{sep}', encoding='{enc}', wiersz_naglowka={nag_idx+1})"
            )

    # filtr HOTEL LINK (jeśli jest)
    if kol_media in df.columns:
        df[kol_media] = df[kol_media].astype(str).str.strip().str.upper()
        df = df[df[kol_media] != "HOTEL LINK"].copy()

    # normalizacja
    df["karta_norm"] = df[kol_karta].astype(str).apply(normalizuj_numer_karty)
    df["nazwisko"] = df[kol_posiadacz].astype(str).apply(wyodrebnij_nazwisko)
    df["ops_kwota_raw"] = df[kol_przychod_hotel].astype(str).apply(przecinek_na_kropke)
    df["ops_kwota"] = pd.to_numeric(df["ops_kwota_raw"], errors="coerce")

    # punkty tylko do klasyfikacji FREQ (jeśli brak — heurystyka po kwocie)
    if kol_punkty in df.columns:
        df["ops_punkty_raw"] = df[kol_punkty].astype(str).apply(przecinek_na_kropke)
        df["ops_punkty"] = pd.to_numeric(df["ops_punkty_raw"], errors="coerce")
    else:
        df["ops_punkty"] = df["ops_kwota"].where(df["ops_kwota"].notna(), 0.0)

    return df


# =======================
# Wejście: Loyalty (XLS/XLSX)
# =======================

def wczytaj_loyalty(sciezka: str) -> pd.DataFrame:
    """
    Czyta Loyalty export. Na zrzucie nagłówki były w 13. wierszu => header=12.
    Dla .xls używamy xlrd; dla .xlsx — wbudowany engine.
    """
    try:
        if sciezka.lower().endswith(".xls"):
            df = pd.read_excel(sciezka, dtype=str, engine="xlrd", header=12)
        else:
            df = pd.read_excel(sciezka, dtype=str, header=12)
    except Exception:
        df = pd.read_excel(sciezka, dtype=str, header=12)

    df.columns = [c.strip() for c in df.columns]

    oczekiwane = {
        "Loyalty Card Number": None,
        "Guest Name": None,
        "Total Revenue (Net of VAT)": None
    }
    for col in df.columns:
        k = col.strip()
        if k in oczekiwane:
            oczekiwane[k] = col

    brakujace = [k for k, v in oczekiwane.items() if v is None]
    if brakujace:
        raise ValueError(
            f"W Loyalty brakuje kolumn: {', '.join(brakujace)}. "
            f"Znalezione kolumny: {list(df.columns)}"
        )

    col_card = oczekiwane["Loyalty Card Number"]
    col_guest = oczekiwane["Guest Name"]
    col_rev  = oczekiwane["Total Revenue (Net of VAT)"]

    df = df[[col_card, col_guest, col_rev]].copy()

    # normalizacja
    df["karta_norm"] = df[col_card].astype(str).apply(normalizuj_numer_karty)
    df["gosc_nazwa_raw"] = df[col_guest].astype(str).str.strip().str.upper()
    df["gosc_nazwisko"] = df["gosc_nazwa_raw"].apply(wyodrebnij_nazwisko)

    df["loyal_kwota_raw"] = df[col_rev].astype(str).apply(przecinek_na_kropke)
    df["loyal_kwota"] = pd.to_numeric(df["loyal_kwota_raw"], errors="coerce")

    return df


# =======================
# Logika porównania
# =======================

def porownaj(
    lojal_df: pd.DataFrame,
    ops_df: pd.DataFrame,
    tolerancja: float = 0.10
) -> Dict[str, pd.DataFrame]:
    """
    Zwraca słownik nazwa_arkusza -> DataFrame do zapisania w Excelu.
    Δ ≤ tolerancja => zgodne; Δ > tolerancja => niezgodne.
    """
    # grupowanie po karcie
    loj_grp = lojal_df.groupby("karta_norm").agg(
        loj_kwoty=("loyal_kwota", lambda x: sorted([v for v in x if pd.notna(v)])),
        goscie_set=("gosc_nazwa_raw", lambda x: set(s for s in x if s)),
        loj_nazwiska=("gosc_nazwisko", lambda x: set(s for s in x if s))
    ).reset_index()

    ops_grp = ops_df.groupby("karta_norm").agg(
        ops_kwoty=("ops_kwota", lambda x: sorted([v for v in x if pd.notna(v)])),
        ops_nazwiska=("nazwisko", lambda x: set(s for s in x if s))
    ).reset_index()

    loj_map: Dict[str, Dict] = {
        r["karta_norm"]: {
            "loj_kwoty": r["loj_kwoty"],
            "goscie": r["goscie_set"],
            "loj_nazwiska": r["loj_nazwiska"]
        } for _, r in loj_grp.iterrows()
    }
    ops_map: Dict[str, Dict] = {
        r["karta_norm"]: {
            "ops_kwoty": r["ops_kwoty"],
            "ops_nazwiska": r["ops_nazwiska"]
        } for _, r in ops_grp.iterrows()
    }

    wszystkie_ops_nazwiska: Set[str] = set(ops_df["nazwisko"].dropna().astype(str).tolist())
    wszystkie_karty = sorted(set(loj_map.keys()) | set(ops_map.keys()))

    # pojemniki
    zgodne, niezgodne, nazw_rozne, roznaliczb, brak_w_ops, ops_brak_w_loyal = [], [], [], [], [], []
    freq_rows = []
    przeglad_rows = []

    # porównania karta-po-karcie + PRZEGLĄD
    for karta in wszystkie_karty:
        l = loj_map.get(karta)
        o = ops_map.get(karta)

        if (l is None) and (o is not None):
            # karta tylko w Operations
            przeglad_rows.append({
                "Karta": karta,
                "Kwota_Loyalty": "—",
                "Kwota_Operations": fmt_list(o["ops_kwoty"]),
                "Δ": "—",
                "Nazwiska_Loyalty": "—",
                "Nazwiska_Operations": fmt_set(o["ops_nazwiska"]),
                "Status": "❌ BRAK KARTY W LOYALTY",
                "Uwaga": "Brak transakcji w Loyalty."
            })
            ops_brak_w_loyal.append({
                "Karta": karta,
                "Nazwiska_Operations": fmt_set(o["ops_nazwiska"]),
                "Kwoty_Operations": fmt_list(o["ops_kwoty"])
            })
            continue

        if (l is not None) and (o is None):
            # karta tylko w Loyalty
            przeglad_rows.append({
                "Karta": karta,
                "Kwota_Loyalty": fmt_list(l["loj_kwoty"]),
                "Kwota_Operations": "—",
                "Δ": "—",
                "Nazwiska_Loyalty": fmt_set(l["loj_nazwiska"]),
                "Nazwiska_Operations": "—",
                "Status": "❌ BRAK KARTY W OPERATIONS",
                "Uwaga": "Brak transakcji w Operations."
            })
            brak_w_ops.append({
                "Karta": karta,
                "Nazwiska_Loyalty": fmt_set(l["loj_nazwiska"]),
                "Kwoty_Loyalty": fmt_list(l["loj_kwoty"])
            })
            continue

        # --- karta w obu ---
        loj_kwoty = l["loj_kwoty"]
        ops_kwoty = o["ops_kwoty"]
        loj_naz = l["loj_nazwiska"]
        ops_naz = o["ops_nazwiska"]

        # 1) Globalnie: czy nazwisko z Loyalty w ogóle występuje w Operations (gdziekolwiek)
        globalnie_brak_nazwiska = not (loj_naz & wszystkie_ops_nazwiska)

        if len(loj_kwoty) != len(ops_kwoty):
            # różna liczba transakcji
            roznaliczb.append({
                "Karta": karta,
                "Nazwiska_Loyalty": fmt_set(loj_naz),
                "Nazwiska_Operations": fmt_set(ops_naz),
                "Kwoty_Loyalty": fmt_list(loj_kwoty),
                "Kwoty_Operations": fmt_list(ops_kwoty)
            })
            przeglad_rows.append({
                "Karta": karta,
                "Kwota_Loyalty": fmt_list(loj_kwoty),
                "Kwota_Operations": fmt_list(ops_kwoty),
                "Δ": "—",
                "Nazwiska_Loyalty": fmt_set(loj_naz),
                "Nazwiska_Operations": fmt_set(ops_naz),
                "Status": "❌ RÓŻNA LICZBA TRANSAKCJI",
                "Uwaga": "Nazwisko z Loyalty nie występuje w Operations (globalnie)." if globalnie_brak_nazwiska else "—"
            })
            continue

        # 2) Porównanie par kwot
        roznice = [abs(lv - ov) for lv, ov in zip(loj_kwoty, ops_kwoty)]
        wszystkie_ok = all(d <= tolerancja for d in roznice)

        for lv, ov in zip(loj_kwoty, ops_kwoty):
            d = abs(lv - ov)
            if d <= tolerancja:
                if (loj_naz & ops_naz):
                    status = "✓ ZGODNE"
                    uwaga = "—"
                    if globalnie_brak_nazwiska:
                        uwaga = "Nazwisko z Loyalty nie występuje w Operations (globalnie)."
                else:
                    status = "⚠️ INNE NAZWISKA"
                    # jeśli w ogóle w Operations takiego nazwiska nie ma — dopisz to
                    if globalnie_brak_nazwiska:
                        uwaga = "Nazwisko z Loyalty nie występuje w Operations (globalnie)."
                    else:
                        uwaga = f"Różne nazwiska: Loyalty={fmt_set(loj_naz)} vs Operations={fmt_set(ops_naz)}"
            else:
                status = "❌ RÓŻNICA KWOT"
                uwaga = "Nazwisko z Loyalty nie występuje w Operations (globalnie)." if globalnie_brak_nazwiska else "—"

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

    # --- Reguła częstotliwości nazwisk (złączenie w jeden arkusz)
    ops_kopia = ops_df.copy()
    ops_kopia["ma_punkty"] = ops_df["ops_punkty"].fillna(0) > 0
    freq = ops_kopia.groupby("nazwisko").agg(
        Wiersze=("nazwisko", "size"),
        Wiersze_z_punktami=("ma_punkty", "sum")
    ).reset_index()

    for _, r in freq.iterrows():
        nazw = r["nazwisko"] or "—"
        rows = int(r["Wiersze"])
        z_pkt = int(r["Wiersze_z_punktami"])
        if rows <= 2:
            continue
        if rows == 3 and z_pkt == 2:
            status = "OK"
            uw = "3 wpisy, punkty za 2 — dozwolone."
        elif z_pkt >= rows:
            status = "OSTRZEŻENIE"
            uw = "Punkty za wszystkie — możliwe duplikaty."
        else:
            status = "INFO"
            uw = "Inny przypadek — do weryfikacji."
        freq_rows.append({
            "Nazwisko": nazw,
            "Wiersze": rows,
            "Wiersze_z_punktami": z_pkt,
            "Status": status,
            "Uwagi": uw
        })
    df_przeglad = pd.DataFrame(przeglad_rows)

    if not df_przeglad.empty:
        # Kategoria/Priorytet po ikonie Statusu:
        # ❌ = PROBLEM (1) > ⚠️ = PROBLEM (2) > ✓ = OK (3)
        ikona = df_przeglad["Status"].astype(str).str[0]
        kat_map = {"❌": ("PROBLEM", 1), "⚠": ("PROBLEM", 2), "✓": ("OK", 3)}
        df_przeglad["Kategoria"] = ikona.map(lambda x: kat_map.get(x, ("INNE", 9))[0])
        df_przeglad["Priorytet"] = ikona.map(lambda x: kat_map.get(x, ("INNE", 9))[1])

        # Удобный порядок колонок
        kol_order = [
            "Kategoria", "Priorytet", "Status", "Karta",
            "Nazwiska_Loyalty", "Nazwiska_Operations",
            "Kwota_Loyalty", "Kwota_Operations", "Δ", "Uwaga"
        ]
        for col in kol_order:
            if col not in df_przeglad.columns:
                df_przeglad[col] = ""

        df_przeglad = df_przeglad[kol_order]

        # Сортировка: Сначала PROBLEM (❌, ⚠️), потом OK (✓)
        df_przeglad.sort_values(
            by=["Kategoria", "Priorytet", "Karta"],
            ascending=[True, True, True],
            inplace=True,
            kind="mergesort"  # стабильная, чтобы не «перемалывать» пары
        )
    else:
        df_przeglad = pd.DataFrame(columns=[
            "Kategoria","Priorytet","Status","Karta",
            "Nazwiska_Loyalty","Nazwiska_Operations",
            "Kwota_Loyalty","Kwota_Operations","Δ","Uwaga"
        ])

    # DataFrames do zapisania
    wyniki = {
        "00_PODSUMOWANIE": pd.DataFrame([
            {"Sekcja": "01_ZGODNE_≤0,10", "Wierszy": len(zgodne)},
            {"Sekcja": "02_NIEZGODNE_>0,10", "Wierszy": len(niezgodne)},
            {"Sekcja": "03_KARTA_OK_INNE_NAZWISKA", "Wierszy": len(nazw_rozne)},
            {"Sekcja": "04_RÓŻNA_LICZBA_POZYCJI", "Wierszy": len(roznaliczb)},
            {"Sekcja": "05_BRAK_KARTY_W_OPERATIONS", "Wierszy": len(brak_w_ops)},
            {"Sekcja": "06_KARTY_W_OPERATIONS_BRAK_W_LOYALTY", "Wierszy": len(ops_brak_w_loyal)},
            {"Sekcja": "07_FREQ", "Wierszy": len(freq_rows)},
            {"Sekcja": "99_PRZEGLAD_TRANSAKCJI", "Wierszy": len(przeglad_rows)},
        ]),
        "01_ZGODNE_≤0,10": pd.DataFrame(zgodne),
        "02_NIEZGODNE_>0,10": pd.DataFrame(niezgodne),
        "03_KARTA_OK_INNE_NAZWISKA": pd.DataFrame(nazw_rozne),
        "04_RÓŻNA_LICZBA_POZYCJI": pd.DataFrame(roznaliczb),
        "05_BRAK_KARTY_W_OPERATIONS": pd.DataFrame(brak_w_ops),
        "06_KARTY_W_OPERATIONS_BRAK_W_LOYALTY": pd.DataFrame(ops_brak_w_loyal),
        "07_FREQ": pd.DataFrame(freq_rows),
        "99_PRZEGLAD_TRANSAKCJI": df_przeglad,
    }
    return wyniki


# =======================
# Zapis do Excela
# =======================

def safe_sheet_name(name: str, used: set) -> str:
    """
    Zwraca bezpieczną nazwę arkusza (<=31 znaków, bez niedozwolonych znaków).
    Gwarantuje unikalność w obrębie skoroszytu (dokleja ~2, ~3...).
    """
    # Preferowane skróty dla naszych długich nazw
    PREFERED = {
        "00_PODSUMOWANIE": "00_PODSUM",
        "01_ZGODNE_≤0,10": "01_ZGODNOŚĆ_≤0.10",
        "02_NIEZGODNE_>0,10": "02_NIEZGODNOŚĆ_>0.10",
        "03_KARTA_OK_INNE_NAZWISKA": "03_INNE_NAZWISKA",
        "04_RÓŻNA_LICZBA_POZYCJI": "04_ROZNA_LICZBA_TRANSAKCJI",
        "05_BRAK_KARTY_W_OPERATIONS": "05_BRAK_W_OPERATIONS",
        "06_KARTY_W_OPERATIONS_BRAK_W_LOYALTY": "06_TYLKO_OPERATIONS",
        "07_FREQ": "07_FREQ",
        "99_PRZEGLAD_TRANSAKCJI": "99_CAŁOŚĆ",
    }

    # 1) Podstaw preferowany skrót (jeśli mamy)
    s = PREFERED.get(name, name)

    # 2) Usuń niedozwolone znaki: []:*?/\
    s = re.sub(r'[\[\]\:\*\?\/\\]', '_', s)

    # 3) Obetnij do 31 znaków (Excel limit)
    MAXLEN = 31
    s = s[:MAXLEN]

    # 4) Zapewnij unikalność
    base = s
    i = 2
    while s in used:
        suffix = f"~{i}"
        s = (base[:MAXLEN-len(suffix)] + suffix)
        i += 1

    used.add(s)
    return s

def wybierz_sciezke_wyjsciowa(folder: Path, limit: int = 31, ext: str = ".xlsx") -> Path:
    """
    Zwraca ścieżkę do pliku wyjściowego w schemacie 01..31.xlsx.
    - Najpierw wybiera pierwszy wolny numer (01..31).
    - Jeśli wszystkie istnieją, wybiera NAJSTARSZY z nich do nadpisania.
    """
    kandydaci = [folder / f"{i:02d}{ext}" for i in range(1, limit + 1)]
    # 1) pierwszy wolny numer
    for p in kandydaci:
        if not p.exists():
            return p
    # 2) jeśli wszystkie zajęte — najstarszy do nadpisania
    najstarszy = min(kandydaci, key=lambda x: x.stat().st_mtime)
    return najstarszy

def zapisz_do_excela(wyniki: Dict[str, pd.DataFrame], plik: Path):
    def colnum_to_excel(n: int) -> str:
        """0->A, 1->B ..."""
        s = ""
        n += 1
        while n:
            n, r = divmod(n - 1, 26)
            s = chr(65 + r) + s
        return s

    with pd.ExcelWriter(plik, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_header = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})
        fmt_wrap   = wb.add_format({"text_wrap": True})
        fmt_title  = wb.add_format({"bold": True, "font_size": 14})

        # Форматы для условной заливки строк в обзоре
        fmt_row_red   = wb.add_format({"bg_color": "#F8CBAD"})
        fmt_row_yellow= wb.add_format({"bg_color": "#FFF2CC"})
        fmt_row_green = wb.add_format({"bg_color": "#C6E0B4"})

        used_names = set()

        # --- 00_PODSUMOWANIE ---
        podsum_src_name = "00_PODSUMOWANIE"
        podsum_sheet = safe_sheet_name(podsum_src_name, used_names)
        df_podsum = wyniki[podsum_src_name].copy()

        df_podsum.to_excel(writer, sheet_name=podsum_sheet, index=False)
        ws0 = writer.sheets[podsum_sheet]
        for col_idx, col_name in enumerate(df_podsum.columns):
            ws0.write(0, col_idx, col_name, fmt_header)
        ws0.set_column(0, 0, 36)
        ws0.set_column(1, 1, 12)
        ws0.write(2 + len(df_podsum), 0, "Legenda:", fmt_title)
        ws0.write(3 + len(df_podsum), 0,
                  "• Zgodność: Δ ≤ 0,10\n"
                  "• Niezgodność: Δ > 0,10\n"
                  "• „KARTA_OK_INNE_NAZWISKA” – karta zgodna, ale nazwiska różne.\n"
                  "• PRZEGLĄD: ✓ zgodne, ⚠️ inne nazwiska, ❌ różnice/braki kart.",
                  fmt_wrap)

        # --- Остальные листы ---
        for src_name, df in wyniki.items():
            if src_name == podsum_src_name:
                continue

            sheet_name = safe_sheet_name(src_name, used_names)

            df = df.copy()
            if df.empty:
                # запишем заглушку, чтобы лист существовал
                if len(df.columns) == 0:
                    df = pd.DataFrame({"Info": ["(brak wpisów)"]})

            # всё в строковый для аккуратного отображения
            for col in df.columns:
                df[col] = df[col].astype(str)

            df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]

            # шапка
            for col_idx, col_name in enumerate(df.columns):
                ws.write(0, col_idx, col_name, fmt_header)

            # ширины
            szer = {
                "Karta": 22,
                "Nazwiska_Loyalty": 30,
                "Nazwiska_Operations": 30,
                "Kwoty_Loyalty": 30,
                "Kwoty_Operations": 30,
                "Różnice_Δ": 22,
                "Wiersze": 10,
                "Wiersze_z_punktami": 20,
                "Status": 22,
                "Uwaga": 40,
                "Info": 24,
                "Kategoria": 12,
                "Priorytet": 10,
                "Δ": 10,
                "Kwota_Loyalty": 16,
                "Kwota_Operations": 18,
            }
            for i, col in enumerate(df.columns):
                ws.set_column(i, i, szer.get(col, 24), fmt_wrap)

            # фильтр + заморозка
            if not df.empty:
                ws.autofilter(0, 0, len(df), len(df.columns) - 1)
            ws.freeze_panes(1, 0)

            # --- Условное форматирование для листа обзора (99_*) ---
            # Подсвечиваем ВСЮ строку по статусу: ❌ (красный), ⚠ (жёлтый), ✓ (зелёный)
            if sheet_name in {"99_PRZEGLAD", "99_CAŁOŚĆ", "99_PRZEGLAD_TRANSAKCJI"} and not df.empty:
                if "Status" in df.columns:
                    n_rows = len(df)
                    n_cols = len(df.columns)
                    status_idx = df.columns.get_loc("Status")
                    status_col_letter = colnum_to_excel(status_idx)

                    # Диапазон всех данных (кроме шапки): строки 2..(n_rows+1), колонки 1..n_cols
                    first_row, last_row = 1, n_rows
                    first_col, last_col = 0, n_cols - 1

                    # Красный: ❌
                    ws.conditional_format(first_row, first_col, last_row, last_col, {
                        "type": "formula",
                        # формула для строки 2, Excel сам сдвинет для остальных
                        "criteria": f'=LEFT(${status_col_letter}2,1)="❌"',
                        "format": fmt_row_red
                    })
                    # Жёлтый: ⚠
                    ws.conditional_format(first_row, first_col, last_row, last_col, {
                        "type": "formula",
                        "criteria": f'=LEFT(${status_col_letter}2,1)="⚠"',
                        "format": fmt_row_yellow
                    })
                    # Зелёный: ✓
                    ws.conditional_format(first_row, first_col, last_row, last_col, {
                        "type": "formula",
                        "criteria": f'=LEFT(${status_col_letter}2,1)="✓"',
                        "format": fmt_row_green
                    })

    print(f"✅ Raport zapisany: {plik.name}")

# =======================
# Główna funkcja
# =======================

def porownaj_punkty_z_kartami():
    root = base_dir()
    # znajdź pliki wejściowe automatycznie
    try:
        operations_path = znajdz_plik_operations(root)
        loyalty_path = znajdz_plik_loyalty(root)
    except Exception as e:
        print("❌ Błąd wyszukiwania plików wejściowych:", e)
        print("Umieść w folderze aplikacji:")
        print("  • plik .csv z frazą 'operation/operations' (Operations),")
        print("  • plik .xls/.xlsx z frazą 'loyalty/loyaltyexport' (Loyalty).")
        return

    print(f"🔎 Operations: {operations_path.name}")
    print(f"🔎 Loyalty:    {loyalty_path.name}")

    # wczytaj źródła
    lojal_df = wczytaj_loyalty(str(loyalty_path))
    ops_df = wczytaj_operations_csv(str(operations_path))

    # porównanie (tolerancja 0,10)
    wyniki = porownaj(lojal_df, ops_df, tolerancja=0.10)

    # nazwa pliku wyjściowego — data uruchomienia (Warszawa/UTC+1/2 nieistotne dla daty)
    output_path = wybierz_sciezke_wyjsciowa(root)
    if output_path.exists():
        print(f"ℹ️  Uwaga: plik {output_path.name} zostanie nadpisany (najstarszy z cyklu 01..31).")
    zapisz_do_excela(wyniki, plik=output_path)

    print("\n✅ Gotowe. Otwórz plik:", output_path.name)


if __name__ == "__main__":
    porownaj_punkty_z_kartami()
