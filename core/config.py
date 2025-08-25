COLS_L = {
    "card": "Loyalty Card Number",
    "guest": "Guest Name",
    "rev": "Total Revenue (Net of VAT)",
    "dep": "Departure",  # opcjonalnie
}

# Operations (Excel; nagłówek w 3. wierszu = header=2)
COLS_O = {
    "pmid": "PMID",
    "card": "ALL card number",
    "holder": "Last name",
    "rev_hotel": "Revenue Hotel currency",
    "points1": "Rewards Points",
    "points2": "Reward points",
    "credit": "Credit type",
    "dep": "Check-out date",  # opcjonalnie
}

STATUS_ALLOWED = [
    "ZGODNE",
    "INNE_NAZWISKA",
    "ROZNICA_KWOT",
    "ROZNA_LICZBA_TRANSAKCJI",
    "BRAK_W_OPERATIONS",
    "BRAK_W_LOYALTY",
]
