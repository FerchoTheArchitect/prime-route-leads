"""
Construye legacy_aliases.json: claves viejas de "Mark as contacted" (prd_contacted) que
paginas anteriores guardaban con otro formato (numero DOT, o un MC distinto con el mismo
telefono) -> MC del carrier actual en leads_verificados.xlsx.

Asi un carrier marcado en una pagina vieja nunca aparece como "Not contacted" en la nueva.
Lee todas las versiones de index.html en el historial de git. Solo contiene numeros MC/DOT
publicos (no dice a quien se contacto). Correr de nuevo despues de cada batch no hace falta;
generar_html.py usa el archivo si existe.

Uso: python generar_aliases.py
"""
import json
import re
import subprocess

import pandas as pd


def digits10(p):
    d = "".join(c for c in str(p).split(".")[0] if c.isdigit())
    return d[-10:] if len(d) >= 10 else ""


df = pd.read_excel("leads_verificados.xlsx", dtype=str).fillna("")
cur_mc = set(df["MC Number"])
mc_by_dot = {d.replace(".0", "").strip(): m for d, m in zip(df["DOT Number"], df["MC Number"]) if d}
mc_by_phone = {digits10(p): m for p, m in zip(df["Telefono"], df["MC Number"]) if digits10(p)}

commits = subprocess.run(["git", "log", "--format=%h", "--", "index.html"],
                         capture_output=True, text=True, check=True).stdout.split()

aliases = {}
for h in commits:
    html = subprocess.run(["git", "show", f"{h}:index.html"], capture_output=True,
                          text=True, encoding="utf-8", errors="ignore").stdout
    for block in re.split(r'(?=<div class="card[^"]*" data-dot=)', html)[1:]:
        key = re.search(r'data-dot="([^"]*)"', block).group(1)
        if not key or key in cur_mc or key in aliases:
            continue
        usdot = re.search(r'USDOT</span><span>(\d+)<', block)
        phone = re.search(r'(?:tel:|sms:)\+?1?(\d{10})\b', block)
        target = (mc_by_dot.get(key) or (usdot and mc_by_dot.get(usdot.group(1)))
                  or (phone and mc_by_phone.get(phone.group(1))))
        if target:
            aliases[key] = target

with open("legacy_aliases.json", "w", encoding="utf-8") as f:
    json.dump(dict(sorted(aliases.items())), f, indent=2)
    f.write("\n")
print(f"{len(commits)} versiones revisadas, {len(aliases)} claves viejas apuntan a carriers actuales:")
for k, v in sorted(aliases.items()):
    print(f"  {k} -> {v}")
