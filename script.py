import pandas as pd
import json
from datetime import datetime

# ==============================
# CONFIG
# ==============================

TRANSIT_FILE = "transit.xlsx"
STOCK_FILE = "estoque.xlsx"
OUTPUT_FILE = "data.json"

# ==============================
# CLASSIFICAÇÃO (MÊS + PRODUTO + QTD)
# ==============================

CLASSIFICATION = {
    ("C04010E2EAFER11003", "Mar/26", 20): "E-mobility",
    ("2080703590", "Mar/26", 1): "Projects",

    ("B1U2ME2E0M5SR11002", "Apr/26", 3): "Projects E-mobility",
    ("B1U2ME2E0M5SR12002", "Apr/26", 10): "Projects E-mobility",

    ("9122003638", "Apr/26", 180): "E-mobility Projects",
    ("9122003639", "Apr/26", 60): "E-mobility Projects",
    ("9122003640", "Apr/26", 450): "E-mobility Projects",
    ("9122003641", "Apr/26", 150): "E-mobility Projects",

    ("5250301010", "Apr/26", 40): "Distribution",
    ("5250301030", "Apr/26", 40): "Distribution",

    ("EDP007SR10001", "May/26", 100): "Projects",
    ("EDP007LR10001", "May/26", 25): "Projects",
    ("BLF-B5150R11001", "May/26", 375): "Projects",
    ("2160100160L", "May/26", 125): "Projects",

    ("M16010E2GYGHR11001", "Jun/26", 1): "Mobility Projects",

    ("9192220364", "Jul/26", 5): "E-mobility",

    ("HXEDE081R10001", "Aug/26", 50): "Projects",

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

    return CLASSIFICATION.get(key, "Distribution")


# ==============================
# CALCULAR DIAS DESDE PO SENT
# ==============================

def days_since_po(po_sent):
    try:
        po_date = pd.to_datetime(po_sent, errors='coerce')
        if pd.isna(po_date):
            return 0

        return (datetime.now() - po_date).days
    except:
        return 0


# ==============================
# CARREGAR ARQUIVOS
# ==============================

df = pd.read_excel(TRANSIT_FILE)
df_stock = pd.read_excel(STOCK_FILE)

# códigos válidos (estoque)
stock_codes = df_stock.iloc[:, 0].astype(str).str.strip().unique()


# ==============================
# PROCESSAMENTO
# ==============================

data = []

for _, r in df.iterrows():
    try:
        necessity = str(r.iloc[6])   # G
        code = str(r.iloc[7]).strip()  # H
        desc = str(r.iloc[8])       # I
        qty = r.iloc[9]             # J
        directorate_manual = str(r.iloc[10])  # K
        cliente = str(r.iloc[11])   # L
        ip_number = str(r.iloc[5])  # F
        po_sent = str(r.iloc[19])   # T
        etd = str(r.iloc[23])       # X
        eta = str(r.iloc[27])       # AB
        status = str(r.iloc[28])    # AC

        # ==============================
        # FILTROS
        # ==============================

        # somente 2026
        if "26" not in necessity:
            continue

        # somente itens do estoque
        if code not in stock_codes:
            continue

        # ==============================
        # PROCESSAMENTO
        # ==============================

        days_po = days_since_po(po_sent)

        directorate = classify_directorate(code, necessity, qty)

        data.append({
            "ip_number": ip_number,
            "client": cliente,
            "directorate": directorate,
            "status": status,
            "po_sent": po_sent,
            "days_since_po_sent": days_po,
            "etd": etd,
            "eta": eta,
            "necessity": necessity,
            "code": code,
            "description": desc,
            "qty": qty
        })

    except Exception as e:
        print(f"Erro: {e}")


# ==============================
# EXPORTAR JSON
# ==============================

with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print("Dashboard atualizado com sucesso!")
