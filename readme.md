# Raport PMID — **Operations ↔ Loyalty**

> Porównywanie danych hotelowych **po PMID** i generowanie kolorowego raportu **XLSX**.  
> Aplikacja ma **GUI** (PL, drag-and-drop) oraz **CLI** (tryb automatyczny).

---

## Spis treści

- [Funkcje](#funkcje)
- [Wymagania](#wymagania)
- [Instalacja](#instalacja)
- [Uruchomienie](#uruchomienie)
  - [GUI](#gui)
  - [CLI](#cli)
- [Format wejścia](#format-wejścia)
  - [Operations](#operations)
  - [Loyalty](#loyalty)
- [Raport XLSX](#raport-xlsx)
- [Konfiguracja](#konfiguracja)
- [Budowanie EXE (Windows)](#budowanie-exe-windows)
- [Prywatność](#prywatność)


---

## Funkcje

- ✅ **PMID jako klucz**:
  - *Operations:* czytane z kolumny `PMID`.
  - *Loyalty:* wyprowadzane z `Loyalty Card Number` → **8 znaków przed ostatnim znakiem**  
    (np. `30810324975248MC` → `4975248M`).
- ✅ **Filtr danych**: z Operations brane są tylko rekordy z `Credit type = Hotel Stay`.
- ✅ **Porównywane kwoty**:
  - *Operations:* `Revenue Hotel currency`
  - *Loyalty:* `Total Revenue (Net of VAT)`
- ✅ **Tolerancja różnic**: domyślnie `|Δ| ≤ 0.10` (zmienialna w GUI).
- ✅ **Wiele plików**: łączenie **kilku** plików Operations i/lub Loyalty w jeden zbiór przed porównaniem.
- ✅ **Raport XLSX**: osobne arkusze, filtry, dopasowane szerokości kolumn, **kolorowanie** statusów, lista wyboru w kolumnie `Status_Manual`.
- ✅ **GUI (PL)**: *ttkbootstrap*, **drag-and-drop**, pasek postępu, logi.
- ✅ **CLI**: tryb automatyczny (sam znajduje najnowsze pliki w folderze programu i tworzy `01.xlsx…31.xlsx` cyklicznie).

---

## Wymagania

- **Python 3.9+**
- Pakiety (w `requirements.txt`):
  - `pandas`, `openpyxl`, `xlrd`, `xlsxwriter`
  - `ttkbootstrap`
  - *(opcjonalnie dla drag-and-drop w GUI)* `tkinterdnd2`

---

## Instalacja

```bash
python -m venv .venv
# Windows (PowerShell):
. .venv\Scripts\Activate.ps1
# macOS/Linux:
# source .venv/bin/activate

pip install --upgrade pip
pip install -r requirements.txt
# (opcjonalnie do DnD)
pip install tkinterdnd2

```

## Uruchomienie 
### Gui

```bash
python app.py --gui
```

Wskaż Operations (.xls/.xlsx) — jeden lub wiele plików (można przeciągnąć).

Wskaż Loyalty (.xls/.xlsx) — jeden lub wiele plików (można przeciągnąć).

Ustaw Tolerancję Δ i ewentualnie ścieżkę Plik wyjściowy (.xlsx).

Kliknij **Generuj raport.**

Podpowiedź: jeśli przeciąganie dodało cudzysłowy do ścieżki, program je usuwa; w razie wątpliwości wybierz plik przyciskiem **Wybierz…**.

### CLI

Połóż najnowsze pliki Operations i Loyalty obok programu i uruchom:

```bash
python app.py
```
Skrypt sam znajdzie pliki i zapisze raport cyklicznie jako 01.xlsx … 31.xlsx
(nadpisuje najstarszy z istniejących).

## Format wejścia

### Operations

- **Excel** (`.xls`/`.xlsx`), nagłówki w **3. wierszu** → `header=2`.
- **Wymagane kolumny:**
  - `PMID`
  - `Last name`
  - `Revenue Hotel currency`
  - `Credit type` *(tylko `Hotel Stay` jest brane do porównania)*
- **Opcjonalne kolumny:**
  - `Rewards Points` **lub** `Reward points`
  - `Check-out date`
  - `ALL card number`

### Loyalty

- **Excel** (`.xls`/`.xlsx`), nagłówki od **13. wiersza** → `header=12`.
- **Wymagane kolumny:**
  - `Loyalty Card Number`
  - `Guest Name`
  - `Total Revenue (Net of VAT)`
- **Opcjonalne:** `Departure`


## Raport XLSX

**Tworzone arkusze:**

- `00_PODSUMOWANIE` — zliczenia sekcji + legenda.
- `01_ZGODNE_≤0.10` — kwoty zgodne i nazwiska zgodne.
- `02_NIEZGODNE_>0.10` — różnice kwot większe niż tolerancja.
- `03_INNE_NAZWISKA` — kwoty zgodne, nazwiska różne.
- `04_ROZNA_LICZBA` — różna liczba transakcji na PMID między źródłami.
- `05_BRAK_W_OPER` — rekordy tylko w Loyalty.
- `06_TYLKO_OPER` — rekordy tylko w Operations.
- `07_FREQ` — szybka analiza częstości nazwisk vs. punktów.
- `99_PRZEGLAD` — pełna lista porównań; kolumny: `Status_Auto`, `Status_Manual`, `Status_Final`, komentarze, daty.

**Kolory w `99_PRZEGLAD`:**

- `ZGODNE` — zielony  
- `INNE_NAZWISKA` — żółty  
- `ROZNICA_KWOT`, `ROZNA_LICZBA_TRANSAKCJI`, `BRAK_W_OPERATIONS`, `BRAK_W_LOYALTY` — czerwony


## Konfiguracja

- **GUI:** zmień pole **„Tolerancja Δ”**.
- **Kod/CLI:** argument `tolerancja` w wywołaniu:

```python 
wyniki = porownaj(lojal_df, ops_df, tolerancja=0.10)
```

## Budowanie EXE (Windows)

> Buduj wewnątrz **aktywnego wirtualnego środowiska** z zainstalowanymi zależnościami.

**Z GUI (bez konsoli):**
```bash
pyinstaller --onefile --windowed --name loyaltymercure ^
  --hidden-import openpyxl --hidden-import et_xmlfile ^
  --collect-data ttkbootstrap --collect-data tkinterdnd2 ^
  app.py
```
Plik pojawi się w dist/loyaltymercure.exe.

## Prywatność
- Aplikacja działa lokalnie — dane nie są wysyłane do Internetu. 
- Raport zawiera nazwiska/PMID/kwoty — przechowuj zgodnie z polityką firmy.
