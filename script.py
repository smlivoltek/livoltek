import pandas as pd
import json
from datetime import datetime

# ==============================
# CONFIG
# ==============================

INPUT_FILE = "transit.xlsx"
OUTPUT_FILE = "data.json"

# ==============================
# CLASSIFICAÇÃO MANUAL
# ==============================

CLASSIFICATION = {
    # MAR/26
    ("C04010E2EAFER11003", "Mar/26", 20): "E-mobility",
    ("2080703590", "Mar/26", 1): "Projects",

    # APR/26 - SPTRANS
    ("B1U2ME2E0M5SR11002", "Apr/26", 3): "Projects E-mobility",
    ("B1U2ME2E0M5SR12002", "Apr/26", 10): "Projects E-mobility",

    # APR/26 - SPENCER
    ("9122003638", "Apr/26", 180): "E-mobility Projects",
    ("9122003639", "Apr/26", 60): "E-mobility Projects",
    ("9122003640", "Apr/26", 450): "E-mobility Projects",
    ("9122003641", "Apr/26", 150): "E-mobility Projects",

    # APR/26 - DISTRIBUTION
    ("5250301010", "Apr/26", 40): "Distribution",
    ("5250301030", "Apr/26", 40): "Distribution",

    # MAY/26
    ("EDP007SR10001", "May/26", 100): "Projects",
    ("EDP007LR10001", "May/26", 25): "Projects",
    ("BLF-B5150R11001", "May/26", 375): "Projects",
    ("2160100160L", "May/26", 125): "Projects",

    # JUN/26
    ("M16010E2GYGHR11001", "Jun/26", 1): "Mobility Projects",

    # JUL/26
    ("9192220364", "Jul/26", 5): "E-mobility",

    # AUG/26
    ("HXEDE081R10001", "Aug/26", 50): "Projects",

    # SEP/26
    ("A0220400E11R11005", "Sep/26", 300): "E-mobility",
    ("C04010E2EAFER11003", "Sep/26", 20): "E-mobility",
    ("C06010E2EAFGR11003", "Sep/26", 10): "E-mobility",
    ("M0601000E2EYR11001", "Sep/26", 20): "E-mobility",
    ("M1201000E2EYR11001", "Sep/26", 40): "E-mobility",
    ("M18010E2EYR11003", "Sep/26", 5): "E-mobility",
    ("M18010E2GYFIR11001", "Sep/26", 1): "E-mobility",
}

def classify_directorate(code, necessity, qty):
    key = (
        str(code).strip(),
        str(necessity).strip().title(),
        int(qty) if str(qty).isdigit() else 0
    )

    if key in CLASSIFICATION:
        return CLASSIFICATION[key]

    return "Distribution"

# ==============================
# FUNÇÃO DELAY
# ==============================

def calculate_delay(eta):
    try:
        eta_date = pd.to_datetime(eta, errors='coerce')
        if pd.isna(eta_date):
            return 0

        today = datetime.now()
        delay = (today - eta_date).days

        return delay if delay > 0 else 0
    except:
        return 0

# ==============================
# LEITURA EXCEL
# ==============================

df = pd.read_excel(INPUT_FILE)

data = []

for _, r in df.iterrows():
    try:
        necessity = str(r.iloc[6])    # G
        code = str(r.iloc[7])         # H
        qty = r.iloc[9]               # J
        desc = str(r.iloc[8])         # I (descrição)
        po = str(r.iloc[6])           # se não tiver PO separado
        etd = str(r.iloc[23])         # X
        eta = str(r.iloc[27])         # AB
        status = str(r.iloc[28])      # AC

        # FILTRO: só 2026
        if "26" not in necessity:
            continue

        delay_days = calculate_delay(eta)

        data.append({
            "po": po,
            "directorate": classify_directorate(code, necessity, qty),
            "status": status,
            "etd": etd,
            "eta": eta,
            "necessity": necessity,
            "code": code,
            "description": desc,
            "qty": qty,
            "delay_days": delay_days
        })

    except Exception as e:
        print(f"Erro na linha: {e}")

# ==============================
# SALVAR JSON
# ==============================

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print("JSON atualizado com sucesso!")
