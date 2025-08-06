# -*- coding: utf-8 -*-
"""
Loyalty vs Operations — porównanie kwot i raport do XLSX (gotowe pod .exe)

Jak zbudować .exe:
    pyinstaller --onefile --name LoyaltyComparator main.py
Opcjonalnie bez konsoli:
    pyinstaller --onefile --noconsole --name LoyaltyComparator main.py
"""

import sys
from pathlib import Path
import re
from typing import List, Dict, Set

import pandas as pd

# =======================
# Ścieżki – praca jako .exe
# =======================

def base_dir() -> Path:
    """Folder bazowy: dla .exe = katalog pliku wykonywalnego, dla .py = katalog skryptu."""
    if getattr(sys, "frozen", False):  # PyInstaller ustawia atrybut 'frozen'
        return Path(sys.executable).parent
    return Path(__file__).parent

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
# Wykrywanie nagłówka/sep/encoding dla CSV
# =======================

def wykryj_naglowek_i_separator_csv(sciezka: str):
    """
    Zwraca (indeks_wiersza_naglowka, encoding, separator).
    Szuka wiersza z nagłówkiem po słowach kluczowych lub liczbie separatorów.
    """
    kandydatury_enc = ["utf-8-sig", "cp1250", "latin-1"]
    slowa_kluczowe = ["card no", "revenue hotel currency", "cardholder (stamped)"]

    for enc in kandydatury_enc:
        with open(sciezka, "r", encoding=enc, errors="ignore") as f:
            linie = []
            for _ in range(100):
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
# Wejście: CSV (operacje)
# =======================

def wczytaj_operacje_csv(sciezka: str) -> pd.DataFrame:
    """
    Czyta operations.csv z auto-wykryciem nagłówka/separ./encoding.
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
    kol_punkty = "Rewards Points"    # używane tylko do reguły częstotliwości; nie raportujemy sumy

    wymagane = [kol_karta, kol_posiadacz, kol_przychod_hotel]
    for c in wymagane:
        if c not in df.columns:
            raise ValueError(
                f"Brak wymaganej kolumny '{c}' w CSV. "
                f"Znalezione kolumny: {list(df.columns)}. "
                f"(wykryto sep='{sep}', encoding='{enc}', wiersz_naglowka={nag_idx+1})"
            )

    # filtr HOTEL LINK (jeśli jest)
    kol_media = "Earn Media"
    if kol_media in df.columns:
        df[kol_media] = df[kol_media].astype(str).str.strip().str.upper()
        df = df[df[kol_media] != "HOTEL LINK"].copy()

    # normalizacja
    df["karta_norm"] = df[kol_karta].astype(str).apply(normalizuj_numer_karty)
    df["nazwisko"] = df[kol_posiadacz].astype(str).apply(wyodrebnij_nazwisko)
    df["ops_kwota_raw"] = df[kol_przychod_hotel].astype(str).apply(przecinek_na_kropke)
    df["ops_kwota"] = pd.to_numeric(df["ops_kwota_raw"], errors="coerce")

    # punkty tylko do reguły częstotliwości (nie sumujemy w raporcie)
    if kol_punkty in df.columns:
        df["ops_punkty_raw"] = df[kol_punkty].astype(str).apply(przecinek_na_kropke)
        df["ops_punkty"] = pd.to_numeric(df["ops_punkty_raw"], errors="coerce")
    else:
        df["ops_punkty"] = df["ops_kwota"].where(df["ops_kwota"].notna(), 0.0)

    return df

# =======================
# Wejście: XLS (lojalność)
# =======================

def wczytaj_lojalnosc_xls(sciezka: str) -> pd.DataFrame:
    """
    Czyta loyalty.xls. Nagłówki w 13. wierszu => header=12.
    """
    try:
        df = pd.read_excel(sciezka, dtype=str, engine="xlrd", header=12)
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
            f"W XLS brakuje kolumn: {', '.join(brakujace)}. "
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
    tolerancja_scisla: float = 0.01,
    tolerancja_miekka: float = 1.0
) -> Dict[str, pd.DataFrame]:
    """
    Zwraca słownik nazw_sekcji -> DataFrame z wynikami do zapisania w Excelu.
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

    # pojemniki na wiersze dla arkuszy
    pelna, toler, nazw_rozne, roznaliczb, niezg, brak_w_csv, csv_brak_w_loyal = [], [], [], [], [], [], []
    freq_ok, freq_warn, freq_info = [], [], []

    # porównania karta-po-karcie
    for karta, l in loj_map.items():
        goscie = l["goscie"]
        loj_kwoty = l["loj_kwoty"]
        loj_nazwiska = l["loj_nazwiska"]

        if karta in ops_map:
            o = ops_map[karta]
            ops_kwoty = o["ops_kwoty"]
            ops_nazwiska = o["ops_nazwiska"]

            # karta OK, ale nazwisko z Loyalty w ogóle nie występuje w CSV -> pokaż oba
            if not (loj_nazwiska & wszystkie_ops_nazwiska):
                nazw_rozne.append({
                    "Karta": karta,
                    "Nazwiska_Loyalty": fmt_set(loj_nazwiska),
                    "Nazwiska_CSV": fmt_set(ops_nazwiska),
                    "Kwoty_Loyalty": fmt_list(loj_kwoty),
                    "Kwoty_CSV": fmt_list(ops_kwoty),
                    "Różnice_Δ": fmt_deltas(loj_kwoty, ops_kwoty)
                })

            if len(loj_kwoty) != len(ops_kwoty):
                roznaliczb.append({
                    "Karta": karta,
                    "Nazwiska_Loyalty": fmt_set(loj_nazwiska),
                    "Nazwiska_CSV": fmt_set(ops_nazwiska),
                    "Kwoty_Loyalty": fmt_list(loj_kwoty),
                    "Kwoty_CSV": fmt_list(ops_kwoty)
                })
            else:
                roznice = [abs(lv - ov) for lv, ov in zip(loj_kwoty, ops_kwoty)]
                if all(d <= tolerancja_scisla for d in roznice):
                    pelna.append({
                        "Karta": karta,
                        "Nazwiska_Loyalty": fmt_set(loj_nazwiska),
                        "Nazwiska_CSV": fmt_set(ops_nazwiska),
                        "Kwoty_Loyalty": fmt_list(loj_kwoty),
                        "Kwoty_CSV": fmt_list(ops_kwoty),
                        "Różnice_Δ": fmt_deltas(loj_kwoty, ops_kwoty)
                    })
                    # jeżeli mimo zgodności kwot nazwiska różne – pokażmy też w arkuszu różnic nazwisk
                    if not (loj_nazwiska & ops_nazwiska):
                        nazw_rozne.append({
                            "Karta": karta,
                            "Nazwiska_Loyalty": fmt_set(loj_nazwiska),
                            "Nazwiska_CSV": fmt_set(ops_nazwiska),
                            "Kwoty_Loyalty": fmt_list(loj_kwoty),
                            "Kwoty_CSV": fmt_list(ops_kwoty),
                            "Różnice_Δ": fmt_deltas(loj_kwoty, ops_kwoty)
                        })
                elif all(d <= 1.0 for d in roznice):
                    toler.append({
                        "Karta": karta,
                        "Nazwiska_Loyalty": fmt_set(loj_nazwiska),
                        "Nazwiska_CSV": fmt_set(ops_nazwiska),
                        "Kwoty_Loyalty": fmt_list(loj_kwoty),
                        "Kwoty_CSV": fmt_list(ops_kwoty),
                        "Różnice_Δ": ", ".join(f"Δ={d:.2f}" for d in roznice)
                    })
                else:
                    niezg.append({
                        "Karta": karta,
                        "Nazwiska_Loyalty": fmt_set(loj_nazwiska),
                        "Nazwiska_CSV": fmt_set(ops_nazwiska),
                        "Kwoty_Loyalty": fmt_list(loj_kwoty),
                        "Kwoty_CSV": fmt_list(ops_kwoty),
                        "Różnice_Δ": ", ".join(f"Δ={d:.2f}" for d in roznice)
                    })
        else:
            brak_w_csv.append({
                "Karta": karta,
                "Nazwiska_Loyalty": fmt_set(loj_nazwiska),
                "Kwoty_Loyalty": fmt_list(loj_kwoty)
            })

    # karty obecne w CSV, brak w Loyalty
    brakujace_karty = set(ops_map.keys()) - set(loj_map.keys())
    for k in sorted(brakujace_karty):
        csv_brak_w_loyal.append({
            "Karta": k,
            "Nazwiska_CSV": fmt_set(ops_map[k]["ops_nazwiska"]),
            "Kwoty_CSV": fmt_list(ops_map[k]["ops_kwoty"])
        })

    # --- Reguła #1: częstotliwość nazwisk (punkty tylko do klasyfikacji)
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
            freq_ok.append({"Nazwisko": nazw, "Wiersze": rows, "Wiersze_z_punktami": z_pkt, "Uwagi": "Dozwolone (3 wpisy, punkty za 2)."})
        elif z_pkt >= rows:
            freq_warn.append({"Nazwisko": nazw, "Wiersze": rows, "Wiersze_z_punktami": z_pkt, "Uwagi": "Punkty za wszystkie — możliwe duplikaty."})
        else:
            freq_info.append({"Nazwisko": nazw, "Wiersze": rows, "Wiersze_z_punktami": z_pkt, "Uwagi": "Do weryfikacji."})

    # Konwersja do DataFrame
    wyniki = {
        "00_PODSUMOWANIE": pd.DataFrame(
            [{"Sekcja": "01_PELNA_ZGODNOSC", "Wierszy": len(pelna)},
             {"Sekcja": "02_ZGODNOSC_W_TOLERANCJI_±1", "Wierszy": len(toler)},
             {"Sekcja": "03_KARTA_OK_NAZWISKA_ROZNE", "Wierszy": len(nazw_rozne)},
             {"Sekcja": "04_ROZNA_LICZBA_POZYCJI", "Wierszy": len(roznaliczb)},
             {"Sekcja": "05_NIEZGODNOSC_KWOT_>±1", "Wierszy": len(niezg)},
             {"Sekcja": "06_BRAK_KARTY_W_CSV", "Wierszy": len(brak_w_csv)},
             {"Sekcja": "07_KARTY_W_CSV_BRAK_W_LOYALTY", "Wierszy": len(csv_brak_w_loyal)},
             {"Sekcja": "08_FREQ_OK", "Wierszy": len(freq_ok)},
             {"Sekcja": "09_FREQ_OSTRZEZENIE", "Wierszy": len(freq_warn)},
             {"Sekcja": "10_FREQ_INFO", "Wierszy": len(freq_info)}]
        ),
        "01_PELNA_ZGODNOSC": pd.DataFrame(pelna),
        "02_ZGODNOSC_W_TOLERANCJI_±1": pd.DataFrame(toler),
        "03_KARTA_OK_NAZWISKA_ROZNE": pd.DataFrame(nazw_rozne),
        "04_ROZNA_LICZBA_POZYCJI": pd.DataFrame(roznaliczb),
        "05_NIEZGODNOSC_KWOT_>±1": pd.DataFrame(niezg),
        "06_BRAK_KARTY_W_CSV": pd.DataFrame(brak_w_csv),
        "07_KARTY_W_CSV_BRAK_W_LOYALTY": pd.DataFrame(csv_brak_w_loyal),
        "08_FREQ_OK": pd.DataFrame(freq_ok),
        "09_FREQ_OSTRZEZENIE": pd.DataFrame(freq_warn),
        "10_FREQ_INFO": pd.DataFrame(freq_info),
    }
    return wyniki

# =======================
# Zapis do Excela
# =======================

def zapisz_do_excela(wyniki: Dict[str, pd.DataFrame], plik: Path):
    with pd.ExcelWriter(plik, engine="xlsxwriter") as writer:
        wb = writer.book
        fmt_header = wb.add_format({"bold": True, "bg_color": "#DDEBF7", "border": 1})
        fmt_wrap = wb.add_format({"text_wrap": True})
        fmt_title = wb.add_format({"bold": True, "font_size": 14})

        # Najpierw podsumowanie
        df_podsum = wyniki["00_PODSUMOWANIE"]
        df_podsum.to_excel(writer, sheet_name="00_PODSUMOWANIE", index=False)
        ws0 = writer.sheets["00_PODSUMOWANIE"]
        # nagłówki już zapisane, dociśnij styl
        for col_idx, col_name in enumerate(df_podsum.columns):
            ws0.write(0, col_idx, col_name, fmt_header)
        ws0.set_column(0, 0, 34)
        ws0.set_column(1, 1, 12)
        ws0.write(2 + len(df_podsum), 0, "Legenda:", fmt_title)
        ws0.write(3 + len(df_podsum), 0,
                  "• Pełna zgodność: Δ ≤ 0.01\n"
                  "• Zgodność w tolerancji: 0.01 < Δ ≤ 1.00 (pokazujemy Δ)\n"
                  "• Niezgodność kwot: Δ > 1.00\n"
                  "• „KARTA_OK_NAZWISKA_ROZNE” – karta zgodna, ale nazwiska różne (pokazujemy oba).",
                  fmt_wrap)

        # Pozostałe arkusze
        for nazwa, df in wyniki.items():
            if nazwa == "00_PODSUMOWANIE":
                continue
            df = df.copy()
            for col in df.columns:
                df[col] = df[col].astype(str)
            df.to_excel(writer, sheet_name=nazwa, index=False)
            ws = writer.sheets[nazwa]
            # nagłówki
            for col_idx, col_name in enumerate(df.columns):
                ws.write(0, col_idx, col_name, fmt_header)
            # szerokości + zawijanie
            szer = {
                "Karta": 22,
                "Nazwiska_Loyalty": 30,
                "Nazwiska_CSV": 30,
                "Kwoty_Loyalty": 30,
                "Kwoty_CSV": 30,
                "Różnice_Δ": 20,
                "Wiersze": 10,
                "Wiersze_z_punktami": 18,
                "Uwagi": 36,
                "Sekcja": 34,
            }
            for i, col in enumerate(df.columns):
                ws.set_column(i, i, szer.get(col, 22), fmt_wrap)
            if not df.empty:
                ws.autofilter(0, 0, len(df), len(df.columns) - 1)
            ws.freeze_panes(1, 0)

    print(f"✅ Raport Excel zapisany do pliku: {plik}")

# =======================
# Główna funkcja
# =======================

def porownaj_punkty_z_kartami():
    root = base_dir()
    loyalty_path = root / "loyalty.xls"
    operations_path = root / "operations.csv"
    output_path = root / "wynik_porownania.xlsx"

    # Proste sprawdzenie obecności plików wejściowych
    missing = [p.name for p in [loyalty_path, operations_path] if not p.exists()]
    if missing:
        print(f"❌ Brak plików wejściowych w folderze: {root}")
        for m in missing:
            print(f"   - {m}")
        print("\nUmieść 'loyalty.xls' i 'operations.csv' obok pliku .exe i uruchom ponownie.")
        return

    # wczytaj źródła
    lojal_df = wczytaj_lojalnosc_xls(str(loyalty_path))
    ops_df = wczytaj_operacje_csv(str(operations_path))

    # porównanie (Δ <= 0.01 pełna zgodność; do ±1 zgodność w tolerancji)
    wyniki = porownaj(
        lojal_df,
        ops_df,
        tolerancja_scisla=0.01,
        tolerancja_miekka=1.0
    )

    # zapis do Excela obok .exe
    zapisz_do_excela(wyniki, plik=output_path)

    print("\n✅ Gotowe. Możesz otworzyć plik:", output_path.name)

if __name__ == "__main__":
    porownaj_punkty_z_kartami()
