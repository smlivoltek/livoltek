import pandas as pd
import json

df = pd.read_excel("transit.xlsx")

data = {}

for _, r in df.iterrows():
    po = str(r.iloc[6]).strip()
    if po == "nan":
        continue

    item = {
        "code": str(r.iloc[10]),
        "qty": int(r.iloc[13]) if not pd.isna(r.iloc[13]) else 0,
        "necessity": str(r.iloc[2]),
        "desc": str(r.iloc[12])
    }

    if po not in data:
        data[po] = {
            "po": po,
            "directorate": "",
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

for po in data:
    if not data[po]["etd"] or data[po]["etd"] == "nan":
        data[po]["risk"] = True

with open("data.json", "w", encoding="utf-8") as f:
    json.dump(list(data.values()), f, ensure_ascii=False, indent=2)
