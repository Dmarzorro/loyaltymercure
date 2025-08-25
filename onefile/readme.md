# Loyalty vs Operations – porównanie i raport XLSX

**Wersja uproszczona.**  
Skrypt porównuje transakcje z plików **Loyalty** i **Operations** (kwoty netto, nazwiska, karty) i generuje raport w Excelu z czytelnymi arkuszami oraz kolorami statusów.

## Co potrafi

- **Nie jest wrażliwy na nazwy plików.**  
  Sam znajduje najnowsze pliki w folderze programu:
  - **Operations:** `*.csv`, `*.xls`, `*.xlsx` ze słowem `operation` lub `operations` w nazwie (np. `REPORT_OPERATION...csv`, `operations_2025.xlsx`).
  - **Loyalty:** `*.xls`, `*.xlsx` ze słowem `loyalty` lub `loyaltyexport` w nazwie (np. `H3417_LoyaltyExport_...xls`).
- Obsługuje **CSV** (auto-wykrycie separatora i wiersza nagłówka) oraz **Excel** (XLS/XLSX).
- **Tolerancja** różnic kwot: **±0,10** (Δ ≤ 0,10 = zgodne; Δ > 0,10 = niezgodne).
- Wykrywa:
  - różne nazwiska przy tej samej karcie,
  - różną liczbę pozycji,
  - karty obecne tylko w jednym z plików.
- Raport Excel z arkuszem przeglądowym, **posortowanym po ważności** (❌/⚠️ nad ✓) i z **kolorami** całych wierszy.
- Nazwy arkuszy dostosowane do limitu Excela (≤ 31 znaków).

---

## Jak przygotować pliki

1. Z systemu **Hotellink** pobierz plik **z jednego dnia** (wyjazdy tego dnia) – **CSV lub XLS/XLSX**.  
2. Z systemu **FOLS** pobierz **Loyalty Export** z tym samym dniem wyjazdu – **XLS/XLSX**.  
3. **Nie musisz zmieniać nazw plików.** Wystarczy, że w nazwie występuje odpowiednie **słowo**:
   - Operations: `operation` lub `operations`
   - Loyalty: `loyalty` lub `loyaltyexport`
4. Umieść oba pliki w **tym samym folderze**, co program (`.exe` albo skrypt `.py`).

---

## Jak uruchomić (gotowe .exe)

1. Skopiuj pliki **Operations** i **Loyalty** do folderu z `LoyaltyComparator.exe`.
2. Uruchom `LoyaltyComparator.exe` / `MajaExport.exe` (podwójny klik lub z wiersza poleceń).
3. Program wypisze, jakie pliki znalazł, i zapisze raport jako **`01.xlsx … 31.xlsx`** (cyklicznie; nadpisuje najstarszy z nich).
4. Otwórz wynikowy plik.

> Jeśli antywirus ostrzeże (typowe dla samodzielnie zbudowanych `.exe`), uruchom z konsoli lub poproś IT o dodanie wyjątku dla lokalnego pliku.

---

## Jak uruchomić (Python)

```bash
python -m venv .venv
.venv\Scripts\activate
pip install --upgrade pip
pip install pandas xlrd openpyxl xlsxwriter
python loyaltyexport.py
```

Co jest w raporcie (arkusze)
00_PODSUMOWANIE – liczba wierszy w każdej sekcji + legenda.

01_ZGODNE_≤0,10 – karty, gdzie kwoty są zgodne (w granicy tolerancji) i nazwiska się pokrywają.

02_NIEZGODNE_>0,10 – karty, gdzie różnice kwot przekraczają tolerancję.

03_KARTA_OK_INNE_NAZWISKA – kwoty zgodne, ale inne nazwiska (np. literówki, inne osoby).

04_RÓŻNA LICZBA POZYCJI – różna liczba transakcji w Loyalty vs Operations.

05_BRAK KARTY W OPERATIONS – karta jest w Loyalty, brak w Operations.

06_KARTY W OPERATIONS BRAK W LOYALTY – karta jest w Operations, brak w Loyalty.

07_FREQ – statystyka częstotliwości nazwisk (np. nazwisko pojawia się 3×, punkty 2× – dozwolone; ostrzeżenia o możliwych duplikatach).

99_PRZEGLĄD_TRANSAKCJI – pełna lista par porównanych kwot z opisem:

Status:

❌ RÓŻNICA KWOT / BRAK KARTY / RÓŻNA LICZBA TRANSAKCJI – czerwony wiersz, najwyższy priorytet.

⚠️ INNE NAZWISKA – żółty wiersz (tolerancja kwot spełniona, ale nazwiska różne).

✓ ZGODNE – zielony wiersz.

Uwaga: dodatkowy komentarz, np. „Różne nazwiska: Loyalty=… vs Operations=…”, lub informacja, że nazwisko z Loyalty w ogóle nie występuje w Operations.

W arkuszach włączone są filtry i zamrożony wiersz nagłówków.

Zmiana tolerancji
Domyślnie 0,10.
Aby zmienić, edytuj wywołanie w funkcji głównej:


wyniki = porownaj(lojal_df, ops_df, tolerancja=0.10)
Typowe komunikaty / rozwiązywanie problemów
„Brak wymaganych kolumn” – w źródłach muszą być co najmniej:

Loyalty: Loyalty Card Number, Guest Name, Total Revenue (Net of VAT) (nagłówek od 13. wiersza).

Operations: Card no., Cardholder (stamped), Revenue hotel currency (+ opcjonalnie Rewards Points, Earn Media).

CSV z nietypowym separatorem lub liniami nad nagłówkiem – skrypt sam wykrywa separator i wiersz nagłówka; błędne linie są pomijane.

Za długie nazwy arkuszy – skrypt automatycznie skraca zgodnie z limitem Excela.

Brak wykrytych plików – sprawdź, czy w nazwie występuje odpowiednie słowo (operation(s), loyalty, loyaltyexport) i czy pliki leżą w folderze programu.

Uwaga dot. prywatności
Skrypt działa lokalnie: nie wysyła danych do Internetu.
Raport zawiera nazwiska i numery kart – przechowuj pliki zgodnie z zasadami firmy.

FAQ
Czy muszę zmieniać nazwy plików wejściowych?
Nie. Wystarczy, że w nazwie występuje odpowiednie słowo: operation/operations lub loyalty/loyaltyexport.

Czy Operations musi być CSV?
Nie. Działa CSV oraz Excel (XLS/XLSX).

Co jeśli nazwiska się różnią, ale kwoty są zgodne?
Zapis trafia do 03_KARTA_OK_INNE_NAZWISKA oraz do przeglądu (⚠️), z komentarzem, które nazwiska się różnią.

Jak nazywa się plik wyjściowy?
Cyklicznie: 01.xlsx … 31.xlsx. Gdy wszystkie istnieją, nadpisywany jest najstarszy z nich.

