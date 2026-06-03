#!/usr/bin/env python3
"""
Livoltek | Atualização automática dos dados de estoque
Baixa os CSVs do Google Sheets e salva em data/
Executado pelo GitHub Actions automaticamente
"""
import urllib.request
import os

SHEETS = {
    "CWB.csv":     "https://docs.google.com/spreadsheets/d/e/2PACX-1vQXyqYNC-Rv_Hva2AT0Mid_HzzMt1FeqXYnHCiIH8rThS-9wBVMEc4KShgLYSv-JkoIgFh6tCczQjXX/pub?gid=0&single=true&output=csv",
    "FOR.csv":     "https://docs.google.com/spreadsheets/d/e/2PACX-1vQXyqYNC-Rv_Hva2AT0Mid_HzzMt1FeqXYnHCiIH8rThS-9wBVMEc4KShgLYSv-JkoIgFh6tCczQjXX/pub?gid=411796829&single=true&output=csv",
    "MAO-LIV.csv": "https://docs.google.com/spreadsheets/d/e/2PACX-1vQXyqYNC-Rv_Hva2AT0Mid_HzzMt1FeqXYnHCiIH8rThS-9wBVMEc4KShgLYSv-JkoIgFh6tCczQjXX/pub?gid=1879495818&single=true&output=csv",
    "MAO-ELE.csv": "https://docs.google.com/spreadsheets/d/e/2PACX-1vQXyqYNC-Rv_Hva2AT0Mid_HzzMt1FeqXYnHCiIH8rThS-9wBVMEc4KShgLYSv-JkoIgFh6tCczQjXX/pub?gid=2006971885&single=true&output=csv",
    "RESERVAS.csv":"https://docs.google.com/spreadsheets/d/e/2PACX-1vQXyqYNC-Rv_Hva2AT0Mid_HzzMt1FeqXYnHCiIH8rThS-9wBVMEc4KShgLYSv-JkoIgFh6tCczQjXX/pub?gid=577818960&single=true&output=csv",
}

os.makedirs("data", exist_ok=True)

ok = 0
for filename, url in SHEETS.items():
    print(f"Baixando {filename}...", end=" ", flush=True)
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30) as r:
            content = r.read().decode("utf-8")
        if content.strip().startswith("<"):
            print("ERRO: retornou HTML (planilha não está publicada como CSV)")
            continue
        with open(f"data/{filename}", "w", encoding="utf-8") as f:
            f.write(content)
        lines = content.strip().count("\n")
        print(f"OK ({lines} linhas)")
        ok += 1
    except Exception as e:
        print(f"ERRO: {e}")

print(f"\n{ok}/{len(SHEETS)} arquivos atualizados.")
if ok == 0:
    raise SystemExit("Nenhum arquivo foi baixado. Verifique se as abas estão publicadas.")
