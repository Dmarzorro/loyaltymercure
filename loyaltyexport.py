import pandas as pd
import re

def extract_surname(full_name):
    if pd.isnull(full_name) or not str(full_name).strip():
        return ""
    parts = str(full_name).strip().split()
    return parts[-1].upper() if parts else ""

