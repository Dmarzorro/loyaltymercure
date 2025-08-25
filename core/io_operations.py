# -*- coding: utf-8 -*-

from pathlib import Path
import pandas as pd

from .config import COLS_O
from .utils import (
    read_excel_safe, normalizuj_numer_karty, normalizuj_pmid,
    przecinek_na_kropke, parse_date_any, fmt_date
)

def _normalize_ops(df: pd.DataFrame) -> pd.DataFrame:
    c = COLS_O

    # 1) filtr: tylko "Hotel Stay"
    if c["credit"] not in df.columns:
        raise ValueError(f"Brak kolumny w Operations: '{c['credit']}'")
    mask_hotel = df[c["credit"]].astype(str).str.strip().str.upper() == "HOTEL STAY"
    df = df[mask_hotel].copy()

    # 2) PMID
    if c["pmid"] not in df.columns:
        raise ValueError(f"Brak kolumny w Operations: '{c['pmid']}'")
    df["pmid"] = df[c["pmid"]].astype(str).apply(normalizuj_pmid)

    # 3) Nazwisko
    if c["holder"] not in df.columns:
        raise ValueError(f"Brak kolumny w Operations: '{c['holder']}'")
    df["nazwisko"] = df[c["holder"]].astype(str).str.strip().str.upper()

    # 4) Kwota
    if c["rev_hotel"] not in df.columns:
        raise ValueError(f"Brak kolumny w Operations: '{c['rev_hotel']}'")
    df["ops_kwota_raw"] = df[c["rev_hotel"]].astype(str).apply(przecinek_na_kropke)
    df["ops_kwota"] = pd.to_numeric(df["ops_kwota_raw"], errors="coerce")

    # 5) Data (opcjonalnie)
    if c["dep"] in df.columns:
        df["ops_data"] = df[c["dep"]].apply(parse_date_any)
    else:
        df["ops_data"] = pd.NaT
    df["ops_data_str"] = df["ops_data"].apply(fmt_date)

    # 6) Punkty (jeśli są)
    points_col = (
        c["points1"] if c["points1"] in df.columns
        else (c["points2"] if c["points2"] in df.columns else None)
    )
    if points_col:
        df["ops_punkty_raw"] = df[points_col].astype(str).apply(przecinek_na_kropke)
        df["ops_punkty"] = pd.to_numeric(df["ops_punkty_raw"], errors="coerce")
    else:
        df["ops_punkty"] = df["ops_kwota"].where(df["ops_kwota"].notna(), 0.0)

    # dodatkowe
    if c.get("card") in df.columns:
        df["karta_norm"] = df[c["card"]].astype(str).apply(normalizuj_numer_karty)
    else:
        df["karta_norm"] = ""

    return df

def wczytaj_operations(path: str) -> pd.DataFrame:
    """Czyta pojedynczy plik Operations (nagłówki w 3. wierszu)."""
    engine = "xlrd" if str(path).lower().endswith(".xls") else "openpyxl"
    df = read_excel_safe(path, dtype=str, header=2, engine=engine)  # <— KLUCZOWE
    df.columns = [(x if isinstance(x, str) else str(x)).strip() for x in df.columns]
    return _normalize_ops(df)

def wczytaj_operations_many(paths: list[str | Path]) -> pd.DataFrame:
    """Scala wiele plików Operations."""
    frames: list[pd.DataFrame] = []
    for p in paths:
        df = wczytaj_operations(str(p))  # już czyści i normalizuje
        df["Źródło"] = Path(str(p)).name
        frames.append(df)
    if not frames:
        return pd.DataFrame()
    return pd.concat(frames, ignore_index=True)
