import pandas as pd
import re

def extract_surname(full_name):
    if pd.isnull(full_name) or not str(full_name).strip():
        return ""
    parts = str(full_name).strip().split()
    return parts[-1].upper() if parts else ""

def fix_card_code(x: str) -> str:
    if not isinstance(x, str):
        x = str(x)
    x = x.strip()
    x = re.sub(r'\s+', '', x)
    try:
        return str(int(float(x))).upper()
    except ValueError:
        return x.upper()

def replace_comma_with_dot(x: str) -> str:
    return x.replace(',', '.')

def porownaj_punkty_z_kartami():
    full_match = []
    fallback_match = []
    count_mismatch = []
    points_mismatch = []
    no_match = []
    output_missing_cards = []

    loyalty_df = pd.read_excel('loyalty.xls', dtype=str, engine='xlrd')
    operations_df = pd.read_csv('operations.csv', sep=';', dtype=str)