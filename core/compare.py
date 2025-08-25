# -*- coding: utf-8 -*-

from typing import Dict, Set
import pandas as pd

from .utils import fmt_set, fmt_list, fmt_list_s, fmt_deltas


def porownaj(lojal_df: pd.DataFrame, ops_df: pd.DataFrame, tolerancja: float = 0.10) -> Dict[str, pd.DataFrame]:
    lojal_df = lojal_df.copy()
    ops_df   = ops_df.copy()

    lojal_df["pair_loyal"] = lojal_df.apply(
        lambda r: (r["loyal_kwota"], r["loyal_data_str"]) if pd.notna(r["loyal_kwota"]) else None, axis=1
    )
    ops_df["pair_ops"] = ops_df.apply(
        lambda r: (r["ops_kwota"], r["ops_data_str"]) if pd.notna(r["ops_kwota"]) else None, axis=1
    )

    # grupy po PMID
    loj_grp = lojal_df.groupby("pmid").agg(
        loj_pary=("pair_loyal", lambda x: sorted([p for p in x if p is not None], key=lambda t: t[0])),
        loj_nazwiska=("gosc_nazwisko", lambda x: set(s for s in x if s))
    ).reset_index()
    ops_grp = ops_df.groupby("pmid").agg(
        ops_pary=("pair_ops",  lambda x: sorted([p for p in x if p is not None], key=lambda t: t[0])),
        ops_nazwiska=("nazwisko", lambda x: set(s for s in x if s))
    ).reset_index()

    # mapy
    loj_map, ops_map = {}, {}
    for _, r in loj_grp.iterrows():
        kw, dt = zip(*r["loj_pary"]) if r["loj_pary"] else ([], [])
        loj_map[r["pmid"]] = {"kw": list(kw), "daty": list(dt), "naz": r["loj_nazwiska"]}
    for _, r in ops_grp.iterrows():
        kw, dt = zip(*r["ops_pary"]) if r["ops_pary"] else ([], [])
        ops_map[r["pmid"]] = {"kw": list(kw), "daty": list(dt), "naz": r["ops_nazwiska"]}

    wszystkie_ops_nazwiska: Set[str] = set(ops_df["nazwisko"].dropna().astype(str).tolist())
    wszystkie_pmid = sorted(set(loj_map) | set(ops_map))

    # sekcje
    zgodne, niezgodne, inne_naz = [], [], []
    roznaliczb, brak_w_ops, ops_brak_w_loyal = [], [], []
    freq_rows, przeglad_rows = [], []

    # porównanie
    for pmid in wszystkie_pmid:
        L = loj_map.get(pmid)
        O = ops_map.get(pmid)

        if L is None and O is not None:
            przeglad_rows.append({
                "PMID": pmid,
                "Kwota_Loyalty": "—", "Kwota_Operations": fmt_list(O["kw"]), "Δ": "—",
                "Data_Loyalty": "—", "Data_Operations": fmt_list_s(O["daty"]),
                "Nazwiska_Loyalty": "—", "Nazwiska_Operations": fmt_set(O["naz"]),
                "Status_Auto": "BRAK_W_LOYALTY", "Uwaga": "Brak transakcji w Loyalty."
            })
            ops_brak_w_loyal.append({
                "PMID": pmid,
                "Nazwiska_Operations": fmt_set(O["naz"]),
                "Kwoty_Operations": fmt_list(O["kw"]),
                "Daty_Operations": fmt_list_s(O["daty"])
            })
            continue

        if L is not None and O is None:
            przeglad_rows.append({
                "PMID": pmid,
                "Kwota_Loyalty": fmt_list(L["kw"]), "Kwota_Operations": "—", "Δ": "—",
                "Data_Loyalty": fmt_list_s(L["daty"]), "Data_Operations": "—",
                "Nazwiska_Loyalty": fmt_set(L["naz"]), "Nazwiska_Operations": "—",
                "Status_Auto": "BRAK_W_OPERATIONS", "Uwaga": "Brak transakcji w Operations."
            })
            brak_w_ops.append({
                "PMID": pmid,
                "Nazwiska_Loyalty": fmt_set(L["naz"]),
                "Kwoty_Loyalty": fmt_list(L["kw"]),
                "Daty_Loyalty": fmt_list_s(L["daty"])
            })
            continue

        # pmid w obu
        loj_kw, ops_kw = L["kw"], O["kw"]
        loj_dt, ops_dt = L["daty"], O["daty"]
        loj_naz, ops_naz = L["naz"], O["naz"]
        globalnie_brak_naz = not (loj_naz & wszystkie_ops_nazwiska)

        if len(loj_kw) != len(ops_kw):
            roznaliczb.append({
                "PMID": pmid,
                "Nazwiska_Loyalty": fmt_set(loj_naz), "Nazwiska_Operations": fmt_set(ops_naz),
                "Kwoty_Loyalty": fmt_list(loj_kw),  "Kwoty_Operations": fmt_list(ops_kw),
                "Daty_Loyalty": fmt_list_s(loj_dt), "Daty_Operations": fmt_list_s(ops_dt),
            })
            przeglad_rows.append({
                "PMID": pmid,
                "Kwota_Loyalty": fmt_list(loj_kw), "Kwota_Operations": fmt_list(ops_kw), "Δ": "—",
                "Data_Loyalty": fmt_list_s(loj_dt), "Data_Operations": fmt_list_s(ops_dt),
                "Nazwiska_Loyalty": fmt_set(loj_naz), "Nazwiska_Operations": fmt_set(ops_naz),
                "Status_Auto": "ROZNA_LICZBA_TRANSAKCJI",
                "Uwaga": "Nazwisko z Loyalty nie występuje w Operations (globalnie)." if globalnie_brak_naz else "—"
            })
            continue

        roznice = [abs(lv - ov) for lv, ov in zip(loj_kw, ops_kw)]
        wszystkie_ok = all(d <= tolerancja for d in roznice)

        for lv, ov, dl, do in zip(loj_kw, ops_kw, loj_dt, ops_dt):
            d = abs(lv - ov)
            if d <= tolerancja:
                if (loj_naz & ops_naz):
                    status = "ZGODNE"
                    uwaga = "Nazwisko z Loyalty nie występuje w Operations (globalnie)." if globalnie_brak_naz else "—"
                else:
                    status = "INNE_NAZWISKA"
                    uwaga = "Nazwisko z Loyalty nie występuje w Operations (globalnie)." if globalnie_brak_naz \
                        else f"Różne nazwiska: Loyalty={fmt_set(loj_naz)} vs Operations={fmt_set(ops_naz)}"
            else:
                status = "ROZNICA_KWOT"
                uwaga = "Nazwisko z Loyalty nie występuje w Operations (globalnie)." if globalnie_brak_naz else "—"

            przeglad_rows.append({
                "PMID": pmid,
                "Kwota_Loyalty": f"{lv:.2f}", "Kwota_Operations": f"{ov:.2f}", "Δ": f"{d:.2f}",
                "Data_Loyalty": dl, "Data_Operations": do,
                "Nazwiska_Loyalty": fmt_set(loj_naz), "Nazwiska_Operations": fmt_set(ops_naz),
                "Status_Auto": status, "Uwaga": uwaga
            })

        if wszystkie_ok:
            target = zgodne if (loj_naz & ops_naz) else inne_naz
            target.append({
                "PMID": pmid,
                "Nazwiska_Loyalty": fmt_set(loj_naz), "Nazwiska_Operations": fmt_set(ops_naz),
                "Kwoty_Loyalty": fmt_list(loj_kw),  "Kwoty_Operations": fmt_list(ops_kw),
                "Daty_Loyalty": fmt_list_s(loj_dt), "Daty_Operations": fmt_list_s(ops_dt),
                "Różnice_Δ": fmt_deltas(loj_kw, ops_kw)
            })
        else:
            niezgodne.append({
                "PMID": pmid,
                "Nazwiska_Loyalty": fmt_set(loj_naz), "Nazwiska_Operations": fmt_set(ops_naz),
                "Kwoty_Loyalty": fmt_list(loj_kw),  "Kwoty_Operations": fmt_list(ops_kw),
                "Daty_Loyalty": fmt_list_s(loj_dt), "Daty_Operations": fmt_list_s(ops_dt),
                "Różnice_Δ": fmt_deltas(loj_kw, ops_kw)
            })

    # FREQ
    ops_tmp = ops_df.copy()
    ops_tmp["ma_punkty"] = ops_tmp["ops_punkty"].fillna(0) > 0
    freq = ops_tmp.groupby("nazwisko").agg(
        Wiersze=("nazwisko","size"),
        Wiersze_z_punktami=("ma_punkty","sum")
    ).reset_index()
    for _, r in freq.iterrows():
        nazw, rows, zpkt = r["nazwisko"] or "—", int(r["Wiersze"]), int(r["Wiersze_z_punktami"])
        if rows <= 2: continue
        if rows == 3 and zpkt == 2: status, uw = "OK", "3 wpisy, punkty za 2 — dozwolone."
        elif zpkt >= rows:          status, uw = "OSTRZEŻENIE", "Punkty za wszystkie — możliwe duplikaty."
        else:                       status, uw = "INFO", "Inny przypadek — do weryfikacji."
        freq_rows.append({"Nazwisko": nazw, "Wiersze": rows, "Wiersze_z_punktami": zpkt, "Status": status, "Uwagi": uw})

    # PRZEGLĄD
    df_przeglad = pd.DataFrame(przeglad_rows)
    if not df_przeglad.empty:
        def _kat(s):  return "OK" if s=="ZGODNE" else "PROBLEM"
        def _prio(s):
            if s in ("ROZNICA_KWOT", "ROZNA_LICZBA_TRANSAKCJI", "BRAK_W_OPERATIONS", "BRAK_W_LOYALTY"): return 1
            if s=="INNE_NAZWISKA": return 2
            return 3
        df_przeglad["Kategoria"] = df_przeglad["Status_Auto"].map(_kat)
        df_przeglad["Priorytet"] = df_przeglad["Status_Auto"].map(_prio)
        df_przeglad["Status_Manual"] = ""
        df_przeglad["Status_Final"]  = df_przeglad["Status_Auto"]

        order = ["Kategoria","Priorytet","Status_Auto","Status_Manual","Status_Final",
                 "PMID","Nazwiska_Loyalty","Nazwiska_Operations",
                 "Kwota_Loyalty","Kwota_Operations","Δ",
                 "Data_Loyalty","Data_Operations","Uwaga"]
        df_przeglad = df_przeglad[order].sort_values(
            ["Kategoria","Priorytet","PMID"], ascending=[True,True,True], kind="mergesort"
        )
    else:
        df_przeglad = pd.DataFrame(columns=[
            "Kategoria","Priorytet","Status_Auto","Status_Manual","Status_Final",
            "PMID","Nazwiska_Loyalty","Nazwiska_Operations",
            "Kwota_Loyalty","Kwota_Operations","Δ","Data_Loyalty","Data_Operations","Uwaga"
        ])

    wyniki = {
        "00_PODSUMOWANIE": pd.DataFrame([
            {"Sekcja":"01_ZGODNE_≤0,10","Wierszy":len(zgodne)},
            {"Sekcja":"02_NIEZGODNE_>0,10","Wierszy":len(niezgodne)},
            {"Sekcja":"03_KARTA_OK_INNE_NAZWISKA","Wierszy":len(inne_naz)},
            {"Sekcja":"04_RÓŻNA_LICZBA_POZYCJI","Wierszy":len(roznaliczb)},
            {"Sekcja":"05_BRAK_KARTY_W_OPERATIONS","Wierszy":len(brak_w_ops)},
            {"Sekcja":"06_KARTY_W_OPERATIONS_BRAK_W_LOYALTY","Wierszy":len(ops_brak_w_loyal)},
            {"Sekcja":"07_FREQ","Wierszy":len(freq_rows)},
            {"Sekcja":"99_PRZEGLAD_TRANSAKCJI","Wierszy":len(df_przeglad)},
        ]),
        "01_ZGODNE_≤0,10": pd.DataFrame(zgodne),
        "02_NIEZGODNE_>0,10": pd.DataFrame(niezgodne),
        "03_KARTA_OK_INNE_NAZWISKA": pd.DataFrame(inne_naz),
        "04_RÓŻNA_LICZBA_POZYCJI": pd.DataFrame(roznaliczb),
        "05_BRAK_KARTY_W_OPERATIONS": pd.DataFrame(brak_w_ops),
        "06_KARTY_W_OPERATIONS_BRAK_W_LOYALTY": pd.DataFrame(ops_brak_w_loyal),
        "07_FREQ": pd.DataFrame(freq_rows),
        "99_PRZEGLAD_TRANSAKCJI": df_przeglad,
    }
    return wyniki
