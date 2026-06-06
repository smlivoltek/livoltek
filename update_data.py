#!/usr/bin/env python3
"""
Livoltek | Atualização automática dos dados de estoque
Baixa os CSVs do Google Sheets e salva em data/
"""
import urllib.request, os

SHEETS = {
    "ESTOQUE.csv":     "https://docs.google.com/spreadsheets/d/e/2PACX-1vQXyqYNC-Rv_Hva2AT0Mid_HzzMt1FeqXYnHCiIH8rThS-9wBVMEc4KShgLYSv-JkoIgFh6tCczQjXX/pub?gid=641758287&single=true&output=csv",
    "RESERVAS.csv":    "https://docs.google.com/spreadsheets/d/e/2PACX-1vQXyqYNC-Rv_Hva2AT0Mid_HzzMt1FeqXYnHCiIH8rThS-9wBVMEc4KShgLYSv-JkoIgFh6tCczQjXX/pub?gid=577818960&single=true&output=csv",
    "CONFERENCIA.csv": "https://docs.google.com/spreadsheets/d/e/2PACX-1vQXyqYNC-Rv_Hva2AT0Mid_HzzMt1FeqXYnHCiIH8rThS-9wBVMEc4KShgLYSv-JkoIgFh6tCczQjXX/pub?gid=935931187&single=true&output=csv",
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
            print(f"ERRO: retornou HTML"); continue
        with open(f"data/{filename}", "w", encoding="utf-8") as f:
            f.write(content)
        print(f"OK ({content.strip().count(chr(10))+1} linhas)")
        ok += 1
    except Exception as e:
        print(f"ERRO: {e}")

print(f"\n{ok}/{len(SHEETS)} arquivos atualizados.")
if ok == 0:
    raise SystemExit("Nenhum arquivo baixado.")
