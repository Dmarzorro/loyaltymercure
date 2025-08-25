# -*- coding: utf-8 -*-

import sys
import re
from pathlib import Path
from typing import List, Set, Tuple
import pandas as pd

from .config import COLS_L, COLS_O  # w razie potrzeby


# ============ Ścieżki / Wyszukiwanie ============

def base_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent.parent

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
    return _find_latest(folder, (".xlsx", ".xls"), ("operation", "operations"))

def znajdz_plik_loyalty(folder: Path) -> Path:
    return _find_latest(folder, (".xlsx", ".xls"), ("loyalty", "loyaltyexport"))


# ============ Normalizacja / helpers ============

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

def normalizuj_pmid(x) -> str:
    s = "" if x is None else str(x)
    return re.sub(r"\s+", "", s).strip().upper()

def wyciagnij_pmid_z_karty(z_karty: str) -> str:
    """
    Jeśli długość >= 9 → 8 znaków przed ostatnim znakiem (np. 30810324975248MC → 4975248M),
    w innym wypadku → ostatnie 8 znaków (fallback).
    """
    if z_karty is None:
        return ""
    s = str(z_karty).strip().upper()
    if len(s) >= 9:
        return s[-9:-1]
    return s[-8:] if len(s) >= 8 else s

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
    if x is None or (isinstance(x, float) and pd.isna(x)) or (isinstance(x, str) and not x.strip()):
        return pd.NaT
    try:
        if isinstance(x, (int, float)) or (isinstance(x, str) and re.fullmatch(r"\d+(\.\d+)?", x.strip())):
            val = float(x)
            return pd.to_datetime(val, origin="1899-12-30", unit="D", errors="coerce")
    except Exception:
        pass
    return pd.to_datetime(str(x).strip().replace("  ", " "), dayfirst=True, errors="coerce")

def fmt_date(dt) -> str:
    return "—" if pd.isna(dt) else pd.Timestamp(dt).strftime("%Y-%m-%d")


# ============ Excel: ścieżka wyjściowa ============

def wybierz_sciezke_wyjsciowa(folder: Path, limit: int = 31, ext: str = ".xlsx") -> Path:
    candidates = [folder / f"{i:02d}{ext}" for i in range(1, limit+1)]
    for p in candidates:
        if not p.exists():
            return p
    return min(candidates, key=lambda x: x.stat().st_mtime)
