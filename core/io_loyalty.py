# core/io_loyalty.py
# -*- coding: utf-8 -*-

from __future__ import annotations
import pandas as pd
from pathlib import Path
from typing import Iterable

from .config import COLS_L
from .utils import (
    read_excel_safe,               # <— kluczowy bezpieczny odczyt
    normalizuj_numer_karty,
    wyciagnij_pmid_z_karty,
    wyodrebnij_nazwisko,
    przecinek_na_kropke,
    parse_date_any,
    fmt_date,
)

def wczytaj_loyalty(path: str) -> pd.DataFrame:
    """
    Czyta pojedynczy plik loyaltyexport (nagłówki od 13. wiersza -> header=12),
    wyprowadza PMID z numeru karty i normalizuje kluczowe kolumny.
    """
    c = COLS_L
    engine = "xlrd" if str(path).lower().endswith(".xls") else "openpyxl"
    try:
        df = read_excel_safe(path, dtype=str, header=12, engine=engine)
    except ImportError as e:
        print("❌ Brak biblioteki do odczytu Excela:", e)
        print("Zainstaluj: pip install openpyxl et-xmlfile  (dla .xlsx) oraz/lub xlrd (dla .xls).")
        raise

    # Nagłówki potrafią nie być str (np. daty) — wymuś str i strip
    df.columns = [(x if isinstance(x, str) else str(x)).strip() for x in df.columns]

    missing = [c[k] for k in ("card", "guest", "rev") if c[k] not in df.columns]
    if missing:
        raise ValueError(f"W Loyalty brakuje kolumn: {missing}.")

    df = df[[c["card"], c["guest"], c["rev"]] + ([c["dep"]] if c["dep"] in df.columns else [])].copy()

    # PMID z numeru karty
    df["karta_norm"] = df[c["card"]].astype(str).apply(normalizuj_numer_karty)
    df["pmid"]       = df["karta_norm"].apply(wyciagnij_pmid_z_karty)

    # Nazwiska
    df["gosc_nazwa_raw"] = df[c["guest"]].astype(str).str.strip().str.upper()
    df["gosc_nazwisko"]  = df["gosc_nazwa_raw"].apply(wyodrebnij_nazwisko)

    # Kwoty
    df["loyal_kwota_raw"] = df[c["rev"]].astype(str).apply(przecinek_na_kropke)
    df["loyal_kwota"]     = pd.to_numeric(df["loyal_kwota_raw"], errors="coerce")

    # Daty (opcjonalnie)
    if c["dep"] in df.columns:
        df["loyal_data"] = df[c["dep"]].apply(parse_date_any)
    else:
        df["loyal_data"] = pd.NaT
    df["loyal_data_str"] = df["loyal_data"].apply(fmt_date)

    return df


def wczytaj_loyalty_many(paths: Iterable[str | Path]) -> pd.DataFrame:
    """
    Scala wiele plików Loyalty w jeden DataFrame (dodaje kolumnę „Źródło”).
    Zwraca pustą ramkę z wymaganymi kolumnami, jeśli lista ścieżek jest pusta.
    """
    paths = [Path(p) for p in paths]
    if not paths:
        return pd.DataFrame(columns=["pmid", "gosc_nazwisko", "loyal_kwota", "loyal_data_str"])

    frames: list[pd.DataFrame] = []
    for p in paths:
        df = wczytaj_loyalty(str(p))
        df["Źródło"] = p.name
        frames.append(df)

    return pd.concat(frames, ignore_index=True)
