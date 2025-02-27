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

    loyalty_df['Loyalty Card Number'] = loyalty_df['Loyalty Card Number'].astype(str).apply(fix_card_code)
    loyalty_df['Guest Name'] = loyalty_df['Guest Name'].astype(str).str.strip().str.upper()
    loyalty_df['Total Revenue (Net of VAT)'] = loyalty_df['Total Revenue (Net of VAT)'].astype(str).apply(replace_comma_with_dot)
    loyalty_df['loyalty_rev'] = pd.to_numeric(loyalty_df['Total Revenue (Net of VAT)'], errors='coerce')
    loyalty_df['card_code_cleaned'] = loyalty_df['Loyalty Card Number'].str.replace(r'\s+', '', regex=True).str.upper()

    operations_df['Card no.'] = operations_df['Card no.'].astype(str).apply(fix_card_code)
    operations_df['surname'] = operations_df['Cardholder (stamped)'].apply(extract_surname)
    operations_df['Revenue hotel currency'] = operations_df['Revenue hotel currency'].astype(str).apply(replace_comma_with_dot)
    operations_df['ops_rev'] = pd.to_numeric(operations_df['Revenue hotel currency'], errors='coerce')
    operations_df['card_no_cleaned'] = operations_df['Card no.'].str.replace(r'\s+', '', regex=True).str.upper()