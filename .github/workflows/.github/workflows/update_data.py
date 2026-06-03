#!/usr/bin/env python3
"""
Baixa os CSVs das planilhas do Google Sheets e salva em data/
Roda via GitHub Actions semanalmente ou manualmente
"""
import urllib.request
import os

SHEETS = {
    "CWB":     "https://docs.google.com/spreadsheets/d/e/2PACX-1vQXyqYNC-Rv_Hva2AT0Mid_HzzMt1FeqXYnHCiIH8rThS-9wBVMEc4KShgLYSv-JkoIgFh6tCczQjXX/pub?gid=0&single=true&output=csv",
    "FOR":     "https://docs.google.com/spreadsheets/d/e/2PACX-1vQXyqYNC-Rv_Hva2AT0Mid_HzzMt1FeqXYnHCiIH8rThS-9wBVMEc4KShgLYSv-JkoIgFh6tCczQjXX/pub?gid=411796829&single=true&output=csv",
    "MAO-LIV": "https://docs.google.com/spreadsheets/d/e/2PACX-1vQXyqYNC-Rv_Hva2AT0Mid_HzzMt1FeqXYnHCiIH8rThS-9wBVMEc4KShgLYSv-JkoIgFh6tCczQjXX/pub?gid=1879495818&single=true&output=csv",
    "MAO-ELE": "https://docs.google.com/spreadsheets/d/e/2PACX-1vQXyqYNC-Rv_Hva2AT0Mid_HzzMt1FeqXYnHCiIH8rThS-9wBVMEc4KShgLYSv-JkoIgFh6tCczQjXX/pub?gid=2006971885&single=true&output=csv",
    "RESERVAS":"https://docs.google.com/spreadsheets/d/e/2PACX-1vQXyqYNC-Rv_Hva2AT0Mid_HzzMt1FeqXYnHCiIH8rThS-9wBVMEc4KShgLYSv-JkoIgFh6tCczQjXX/pub?gid=577818960&single=true&output=csv",
}

os.makedirs("data", exist_ok=True)

for name, url in SHEETS.items():
    print(f"Baixando {name}...", end=" ")
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30) as r:
            content = r.read().decode("utf-8")
        if content.strip().startswith("<"):
            print(f"ERRO: retornou HTML")
            continue
        with open(f"data/{name}.csv", "w", encoding="utf-8") as f:
            f.write(content)
        lines = content.strip().count("\n") + 1
        print(f"OK ({lines} linhas)")
    except Exception as e:
        print(f"ERRO: {e}")

print("\nDone.")
