import pandas as pd
import json

# ===== FUNÇÃO: CLASSIFICA DIRETORIA =====
def classify_directorate(cliente, desc, code):
    text = (str(cliente) + " " + str(desc) + " " + str(code)).lower()

    # CLIENTES PRIORITÁRIOS
    if "sptrans" in text:
        return "Projects E-mobility"

    if "piracicabana" in text:
        return "Projects E-mobility"

    if "vinicio carrara" in text:
        return "E-mobility"

    if "mateus gomes" in text:
        return "Projects"

    if "boat engine" in text:
        return "Projects"

    if "kelly li" in text:
        return "Distribution"

    # PRODUTOS
    if "motor de popa" in text:
        return "Projects"

    if "bess" in text or "pcs" in text:
        return "Projects"

    if "charger" in text or "ev" in text:
        return "E-mobility"

    if "medidor" in text or "hexing" in text:
        return "Hexing"

    return "Distribution"


# ===== LER EXCEL =====
df = pd.read_excel("transit.xlsx")

data = {}

for _, r in df.iterrows():

    # FILTRO: SOMENTE 2026
    if "26" not in str(r.iloc[2]):
        continue

    po = str(r.iloc[6]).strip()
    if po == "nan":
        continue

    cliente = str(r.iloc[1])
    desc = str(r.iloc[12])
    code = str(r.iloc[10])

    item = {
        "code": code,
        "qty": int(r.iloc[13]) if not pd.isna(r.iloc[13]) else 0,
        "necessity": str(r.iloc[2]),
        "desc": desc
    }

    if po not in data:
        data[po] = {
            "po": po,
            "directorate": classify_directorate(cliente, desc, code),
            "status": str(r.iloc[28]),
            "etd": str(r.iloc[23]),
            "eta_dest": str(r.iloc[27]),
            "po_sent": str(r.iloc[19]),
            "items": [],
            "total_qty": 0,
            "wh": "",
            "risk": False,
            "impact": "medium"
        }

    data[po]["items"].append(item)
    data[po]["total_qty"] += item["qty"]


# ===== REGRA DE RISCO =====
for po in data:
    if not data[po]["etd"] or data[po]["etd"] == "nan":
        data[po]["risk"] = True


# ===== SALVAR JSON =====
with open("data.json", "w", encoding="utf-8") as f:
    json.dump(list(data.values()), f, ensure_ascii=False, indent=2)
