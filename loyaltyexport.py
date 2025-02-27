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

    if "Earn Media" in operations_df.columns:
        operations_df["Earn Media"] = operations_df["Earn Media"].astype(str).str.strip().str.upper()
        operations_df = operations_df[operations_df["Earn Media"] != "HOTEL LINK"]

    tolerance = 0.01

    loyalty_group = loyalty_df.groupby('card_code_cleaned').agg({
        'loyalty_rev': list,
        'Guest Name': lambda x: set(x)
    }).reset_index()

    ops_group = operations_df.groupby('card_no_cleaned').agg({
        'ops_rev': list,
        'surname': lambda x: set(x)
    }).reset_index()

    loyalty_dict = { row['card_code_cleaned']: {'loyalty_rev': sorted(row['loyalty_rev']),
                                                  'guest_names': row['Guest Name']}
                     for _, row in loyalty_group.iterrows() }
    ops_dict = { row['card_no_cleaned']: {'ops_rev': sorted(row['ops_rev']),
                                          'surnames': row['surname']}
                 for _, row in ops_group.iterrows() }

    for card, ldata in loyalty_dict.items():
        guest_names = ldata['guest_names']
        loyalty_revs = ldata['loyalty_rev']
        key = card
        if key in ops_dict:
            odata = ops_dict[key]
            ops_revs = odata['ops_rev']
            op_surnames = odata['surnames']
            if len(loyalty_revs) != len(ops_revs):
                count_mismatch.append(
                    f"Dla {', '.join(guest_names)} (karta: {card}): liczba operacji nie zgadza się (Loyalty: {loyalty_revs}, Operations: {ops_revs})."
                )
            else:
                pairs_match = [abs(l - o) < tolerance for l, o in zip(loyalty_revs, ops_revs)]
                if all(pairs_match):
                    if guest_names == op_surnames:
                        full_match.append(
                            f"Dla {', '.join(guest_names)} (karta: {card}): punkty się zgadzają ({loyalty_revs})."
                        )
                    else:
                        fallback_match.append(
                            f"Dla {', '.join(guest_names)} (karta: {card}): Nazwiska się nie zgadzają, ale punkty się zgadzają ({loyalty_revs})."
                        )
                else:
                    points_mismatch.append(
                        f"Dla {', '.join(guest_names)} (karta: {card}): Nazwiska się zgadzają, ale punkty się nie zgadzają (Loyalty: {loyalty_revs}, Operations: {ops_revs})."
                    )
        else:
            no_match.append(
                f"Dla {', '.join(guest_names)} (karta: {card}): brak dopasowania w Operations."
            )

    missing_cards = set(ops_dict.keys()) - set(loyalty_dict.keys())
    if missing_cards:
        output_missing_cards.append("Karty obecne w Operations, ale brak w Loyalty:")
        for card in sorted(missing_cards):
            names = ', '.join(ops_dict[card]['surnames'])
            output_missing_cards.append(f"  {names} (karta: {card})")

    output_lines = []
    output_lines.append("=== FULL MATCHES ===")
    output_lines.extend(full_match)
    output_lines.append("\n=== FALLBACK MATCHES (surname mismatch, но punkty się zgadzają) ===")
    output_lines.extend(fallback_match)
    output_lines.append("\n=== COUNT MISMATCH ===")
    output_lines.extend(count_mismatch)
    output_lines.append("\n=== POINTS MISMATCH ===")
    output_lines.extend(points_mismatch)
    output_lines.append("\n=== NO MATCH ===")
    output_lines.extend(no_match)
    if output_missing_cards:
        output_lines.append("\n=== MISSING CARDS (Operations, ale brak w Loyalty) ===")
        output_lines.extend(output_missing_cards)

    output_filename = "comparison_result.txt"
    with open(output_filename, "w", encoding="utf-8") as f:
        for line in output_lines:
            f.write(line + "\n")
    print("\n".join(output_lines))
    print(f"\n✅ Wynik zapisany w pliku {output_filename}")

if __name__ == "__main__":
    porownaj_punkty_z_kartami()