import pandas as pd
from collections import Counter

df = pd.read_excel('leads_verificados.xlsx', dtype=str).fillna('')
nuevos = df.tail(284)

print("=== NUEVOS CARRIERS AGREGADOS ===")
print(f"Total nuevos: {len(nuevos)}")
print()

print("--- TIPO VEHICULO ---")
for val, cnt in nuevos['Tipo Vehiculo'].value_counts().items():
    label = val if val.strip() else "(sin dato)"
    print(f"  {label}: {cnt}")

tt = len(nuevos[nuevos['Tipo Vehiculo'].str.contains('Truck Tractor', na=False)])
print(f"\nCon Truck Tractor: {tt}/{len(nuevos)} ({tt*100//len(nuevos)}%)")
print()

print("--- FECHAS DE REGISTRO ---")
fechas = nuevos['Fecha Registro'].astype(str)
print(f"  Mas reciente: {fechas.max()}")
print(f"  Mas vieja:    {fechas.min()}")
print()

meses = Counter()
for f in fechas:
    f = f.strip()
    if len(f) >= 7:
        meses[f[:7]] += 1
    elif len(f) == 8 and f.isdigit():
        meses[f[:4] + "-" + f[4:6]] += 1
    else:
        meses[f] += 1

print("POR MES:")
for m in sorted(meses.keys(), reverse=True):
    print(f"  {m}: {meses[m]}")

print()
print("--- TIPO DE CARGA ---")
for val, cnt in nuevos['Tipo de Carga'].value_counts().items():
    print(f"  {val}: {cnt}")
