"""
Genera la pagina de leads (index.html) para usar desde el celular.
- Los textos de los mensajes vienen de messages.json (V0-V3 rotan).
- El diseno/JS de la pagina esta en plantilla_leads.html.
- El log de envios, respuestas y notas se guarda SOLO en el celular (localStorage),
  nunca en este repo ni en la pagina publica.

Uso: python generar_html.py            (abre el HTML al terminar)
     python generar_html.py --no-open
"""
import json
import os
import re
import sys

import pandas as pd

ARCHIVO_EXCEL     = "leads_verificados.xlsx"
ARCHIVO_MENSAJES  = "messages.json"
ARCHIVO_PLANTILLA = "plantilla_leads.html"
ARCHIVOS_HTML     = ["index.html", "leads_carriers.html"]


def friendly_name(raw):
    """ELITE TRUCKING LLC  →  Elite Trucking"""
    name = raw.strip()
    # Quitar sufijos legales
    name = re.sub(r'\b(LLC|L\.L\.C\.?|INC\.?|INCORPORATED|CORP\.?|CORPORATION|CO\.?|LTD\.?|LP|L\.P\.)\b',
                  '', name, flags=re.IGNORECASE)
    # Quitar guiones/comas sueltos al final
    name = re.sub(r'[,\-\s]+$', '', name).strip()
    # Title Case
    name = name.title()
    return name if name else raw.strip().title()


def clean_phone(p):
    # Si es float (ej: 9012658635.0), convertir a int primero
    try:
        if float(p) == int(float(p)):
            p = str(int(float(p)))
    except (ValueError, TypeError, OverflowError):
        pass
    digits = "".join(c for c in str(p) if c.isdigit())
    return digits[-10:] if len(digits) >= 10 else ""


def clean_num(v):
    v = str(v).strip()
    return v[:-2] if v.endswith(".0") else v


with open(ARCHIVO_MENSAJES, encoding="utf-8") as f:
    mensajes = json.load(f)
for v in mensajes["rotation"]:
    if v not in mensajes["messages"]:
        sys.exit(f"messages.json: falta el texto de {v}")
    texto = mensajes["messages"][v]
    if texto and "{company_name}" not in texto:
        print(f"Aviso: {v} no tiene {{company_name}}")

print("Leyendo Excel...")
df = pd.read_excel(ARCHIVO_EXCEL, dtype=str)
df = df.fillna("")

# Solo los que tienen telefono
df = df[df["Telefono"].str.strip() != ""].copy()
df["_digits"] = df["Telefono"].apply(clean_phone)
df = df[df["_digits"] != ""].copy()

print(f"Generando HTML con {len(df):,} carriers...")

leads = []
for _, row in df.iterrows():
    nombre    = row.get("Nombre de la Empresa", "")
    direccion = row.get("Direccion", "")
    usdot     = clean_num(row.get("DOT Number", ""))
    camiones  = clean_num(row.get("Num Camiones", ""))
    leads.append({
        "mc":  row.get("MC Number", ""),
        "dot": usdot if usdot and usdot != "nan" else "",
        "n":   nombre,
        "f":   friendly_name(nombre) if nombre else "your company",
        "p":   row["_digits"],
        "c":   direccion.split(",")[1].strip() if "," in direccion else "",
        "s":   row.get("Estado", ""),
        "t":   int(camiones) if camiones.isdigit() else camiones,
        "r":   clean_num(row.get("Fecha Registro", ""))[:8],
        "g":   row.get("Tipo de Carga", ""),
        "v":   row.get("Tipo Vehiculo", ""),
    })


def as_script_json(obj):
    # Seguro dentro de <script>: evita que "</script>" en un nombre cierre la etiqueta
    return json.dumps(obj, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")


with open(ARCHIVO_PLANTILLA, encoding="utf-8") as f:
    html = f.read()
html = (html.replace("__LEADS_JSON__", as_script_json(leads))
            .replace("__MESSAGES_JSON__", as_script_json(mensajes)))

for archivo in ARCHIVOS_HTML:
    with open(archivo, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"HTML guardado: {os.path.abspath(archivo)}")

if "--no-open" not in sys.argv:
    os.startfile(os.path.abspath(ARCHIVOS_HTML[0]))
