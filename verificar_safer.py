"""
Verifica carriers contra SAFER usando DOT number.
Busca carriers registrados hace 2-12 meses (ya tuvieron tiempo de obtener autoridad).
Solo conserva los AUTHORIZED FOR Property.
"""

import pandas as pd
import requests
import time
import os
from io import StringIO
from datetime import datetime, timedelta

# Rango: 6 a 18 meses atras
fecha_max = (datetime.now() - timedelta(days=150)).strftime("%Y%m%d")   # hace 5 meses
fecha_min = (datetime.now() - timedelta(days=540)).strftime("%Y%m%d")   # hace 18 meses

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml",
}

print(f"Buscando carriers registrados entre {fecha_min} y {fecha_max}...")

base_url = "https://data.transportation.gov/resource/az4n-8mr2.csv"
CANADA_PROVINCES = "('ON','QC','BC','AB','MB','SK','NS','NB','PE','NL','NT','YT','NU')"

where = (
    f"add_date >= '{fecha_min}' "
    f"AND add_date <= '{fecha_max}' "
    f"AND status_code = 'A' "
    f"AND docket1prefix = 'MC' "
    f"AND docket1_status_code = 'A' "
    f"AND interstate_beyond_100_miles != '0' "
    f"AND power_units::number = 1 "
    f"AND phy_state NOT IN {CANADA_PROVINCES} "
    f"AND (crgo_genfreight = 'X' OR crgo_metalsheet = 'X' OR crgo_bldgmat = 'X' OR crgo_construct = 'X') "
    f"AND (owntract::number > 0 OR trmtract::number > 0 OR trptract::number > 0 "
    f"OR owntrail::number > 0 OR trmtrail::number > 0 OR trptrail::number > 0)"
)

todos = []
offset = 0
while True:
    params = {
        "$select": "dot_number,docket1,legal_name,phone,phy_street,phy_city,phy_state,phy_zip,power_units,add_date,"
                   "crgo_genfreight,crgo_metalsheet,crgo_bldgmat,crgo_drybulk,crgo_construct,"
                   "owntract,trmtract,trptract,owntrail,trmtrail,trptrail",
        "$where": where,
        "$order": "add_date DESC",
        "$limit": 50000,
        "$offset": offset,
    }
    r = requests.get(base_url, params=params, timeout=120)
    r.raise_for_status()
    chunk = pd.read_csv(StringIO(r.text))
    if len(chunk) == 0:
        break
    todos.append(chunk)
    offset += len(chunk)
    print(f"  Descargados: {offset:,} carriers...")
    if len(chunk) < 50000:
        break

df = pd.concat(todos, ignore_index=True)
print(f"Total FMCSA: {len(df):,}")

# Filtrar con telefono
df["phone"] = df["phone"].fillna("").astype(str).str.strip()
df = df[df["phone"] != ""].copy()
print(f"Con telefono: {len(df):,}")

# Cargar carriers ya verificados para no repetir
ARCHIVO_SALIDA = "leads_verificados.xlsx"
ya_verificados = set()
df_previo = None
if os.path.exists(ARCHIVO_SALIDA):
    df_previo = pd.read_excel(ARCHIVO_SALIDA)
    ya_verificados = set(df_previo["MC Number"].astype(str).str.replace("MC-", ""))
    print(f"Ya verificados previamente: {len(ya_verificados)}")

# Saltar los ya verificados y tomar los siguientes 500
df["_mc"] = df["docket1"].astype(str).str.replace(".0", "", regex=False)
df_pendientes = df[~df["_mc"].isin(ya_verificados)].copy()
df_check = df_pendientes.head(500).copy()
print(f"A verificar contra SAFER: {len(df_check)} (saltando {len(ya_verificados)} ya verificados)")


def check_safer(dot_number):
    """Consulta Motus por DOT y verifica propertyChk + operatingAuthorityStatus Active."""
    try:
        resp = requests.get(
            f"https://motus.dot.gov/api/carriers/{dot_number}",
            headers=HEADERS,
            timeout=15,
        )

        if resp.status_code == 404:
            return False, "Not found"

        resp.raise_for_status()
        data = resp.json()

        regs = data.get("entityRegistrations", [])
        for reg in regs:
            if not reg.get("propertyChk") or not reg.get("forHireChk"):
                continue
            for eoa in reg.get("entityRegistrationOperatingAuthorities", []):
                authority = eoa.get("entityOperatingAuthority", {})
                status_obj = authority.get("operatingAuthorityStatus", {})
                status_name = status_obj.get("operatingAuthorityStatusName", "")
                if status_name.lower() == "active":
                    return True, "Authorized"

        if data.get("entityId"):
            return False, "Not Authorized"

        return False, "Unknown"

    except requests.exceptions.Timeout:
        return False, "Timeout"
    except Exception:
        return False, "Error"


# Verificar
auth_indices = []
counts = {"Authorized": 0, "Not Authorized": 0, "Not found": 0, "Error": 0, "Timeout": 0, "Unknown": 0}

for idx, (i, row) in enumerate(df_check.iterrows()):
    dot = row["dot_number"]
    nombre = str(row.get("legal_name", ""))[:40]
    mc = row.get("docket1", "")

    is_auth, status = check_safer(dot)
    counts[status] = counts.get(status, 0) + 1

    symbol = "OK" if is_auth else "X "
    print(f"  [{idx+1}/{len(df_check)}] {symbol} DOT:{dot} MC:{mc} - {nombre} -> {status}")

    if is_auth:
        auth_indices.append(i)

    time.sleep(0.3)

print(f"\n{'='*60}")
for k, v in counts.items():
    if v > 0:
        print(f"  {k}: {v}")
print(f"  TOTAL AUTORIZADOS: {len(auth_indices)}")
print(f"{'='*60}")

# Generar Excel con nuevos autorizados + previos
df_auth_new = df.loc[auth_indices].copy()

# Preparar columnas para los nuevos
df_auth_new["Direccion"] = (
    df_auth_new["phy_street"].fillna("") + ", " +
    df_auth_new["phy_city"].fillna("") + ", " +
    df_auth_new["phy_state"].fillna("") + " " +
    df_auth_new["phy_zip"].fillna("")
)

cargo_cols = {
    "crgo_genfreight": "General Freight / Dry Van",
    "crgo_metalsheet": "Metal / Acero",
    "crgo_bldgmat": "Materiales de Construccion",
    "crgo_drybulk": "Dry Bulk",
    "crgo_construct": "Maquinaria de Construccion",
}
def tipo_carga(row):
    tipos = [label for col, label in cargo_cols.items() if str(row.get(col, "")).strip().upper() in ("Y", "X")]
    return ", ".join(tipos) if tipos else "Otro"

df_auth_new["Tipo de Carga"] = df_auth_new.apply(tipo_carga, axis=1)
df_auth_new["MC Number"] = "MC-" + df_auth_new["docket1"].astype(str).str.replace(".0", "", regex=False)

def tipo_vehiculo(row):
    def num(v):
        try: return int(float(v))
        except: return 0
    t = num(row.get("owntract", 0)) + num(row.get("trmtract", 0)) + num(row.get("trptract", 0))
    s = num(row.get("owntruck", 0)) + num(row.get("trmtruck", 0)) + num(row.get("trptruck", 0))
    if t > 0 and s == 0: return "Truck Tractor"
    if t > 0 and s > 0: return "Truck Tractor + Straight"
    if s > 0: return "Straight Truck"
    return ""

df_auth_new["Tipo Vehiculo"] = df_auth_new.apply(tipo_vehiculo, axis=1)

def make_sms_link(phone):
    digits = "".join(c for c in str(phone) if c.isdigit())
    if len(digits) < 10:
        return ""
    return f"tel:{digits[-10:]}"

df_auth_new["Enviar Mensaje"] = df_auth_new["phone"].apply(make_sms_link)

resultado_new = df_auth_new[[
    "MC Number", "legal_name", "phone",
    "Enviar Mensaje", "Direccion", "phy_state",
    "power_units", "add_date", "Tipo de Carga", "dot_number", "Tipo Vehiculo"
]].copy()

resultado_new.columns = [
    "MC Number", "Nombre de la Empresa", "Telefono",
    "Enviar Mensaje", "Direccion", "Estado",
    "Num Camiones", "Fecha Registro", "Tipo de Carga", "DOT Number", "Tipo Vehiculo"
]

# Combinar con previos
if df_previo is not None and len(df_previo) > 0:
    resultado = pd.concat([df_previo, resultado_new], ignore_index=True)
    resultado = resultado.drop_duplicates(subset=["MC Number"])
    print(f"\nPrevios: {len(df_previo)} + Nuevos: {len(resultado_new)} = Total: {len(resultado)}")
else:
    resultado = resultado_new

# Eliminar duplicados por telefono (mismo numero = mismo dueno, queda el primero)
def norm_phone(p):
    digits = "".join(c for c in str(p) if c.isdigit())
    return digits[-10:] if len(digits) >= 10 else ""

resultado["_phone_norm"] = resultado["Telefono"].apply(norm_phone)
antes = len(resultado)
resultado = resultado[resultado["_phone_norm"] != ""]  # quitar sin telefono
resultado = resultado.drop_duplicates(subset=["_phone_norm"], keep="first")
resultado = resultado.drop(columns=["_phone_norm"])
print(f"Duplicados por telefono eliminados: {antes - len(resultado)}")

resultado.to_excel(ARCHIVO_SALIDA, index=False)
ruta = os.path.abspath(ARCHIVO_SALIDA)
print(f"Archivo: {ruta}")
print(f"Total carriers autorizados: {len(resultado):,}")
