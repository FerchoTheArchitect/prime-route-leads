"""
Verifica carriers contra SAFER usando DOT number.
Por defecto busca carriers registrados hace 2-6 meses con 1 camion (como siempre).
Solo conserva los AUTHORIZED FOR Property.

Opciones (todas opcionales):
  --min-months 12 --max-months 36   ventana de registro (meses atras)
  --min-trucks 1  --max-trucks 3    numero de camiones (power units)
  --batch 500                       cuantos verificar contra Motus
  --dry-run                         solo muestra el filtro y la lista do-not-contact, no descarga nada
Ejemplo: python verificar_safer.py --min-months 12 --max-months 36 --max-trucks 3
"""

import argparse
import glob
import pandas as pd
import requests
import sys
import time
import os
from io import StringIO
from datetime import datetime, timedelta

ap = argparse.ArgumentParser(description="Verifica carriers nuevos contra Motus.")
ap.add_argument("--min-months", type=int, default=2, help="registrados hace al menos N meses (default 2)")
ap.add_argument("--max-months", type=int, default=6, help="registrados hace como maximo N meses (default 6)")
ap.add_argument("--min-trucks", type=int, default=1, help="minimo de camiones (default 1)")
ap.add_argument("--max-trucks", type=int, default=1, help="maximo de camiones (default 1)")
ap.add_argument("--batch", type=int, default=500, help="cuantos verificar contra Motus (default 500)")
ap.add_argument("--private-dir", default="private", help="carpeta local con do_not_contact.csv / exports del celular")
ap.add_argument("--dry-run", action="store_true", help="no descarga ni verifica nada")
args = ap.parse_args()
if args.min_months > args.max_months or args.min_trucks > args.max_trucks:
    sys.exit("Rango invalido: min debe ser <= max")

# Rango de registro (1 mes = 30 dias, igual que antes: 2-6 meses = 60-180 dias)
fecha_max = (datetime.now() - timedelta(days=args.min_months * 30)).strftime("%Y%m%d")
fecha_min = (datetime.now() - timedelta(days=args.max_months * 30)).strftime("%Y%m%d")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml",
}


def norm_phone(p):
    # 9012658635.0 -> 9012658635
    try:
        if float(p) == int(float(p)):
            p = str(int(float(p)))
    except (ValueError, TypeError, OverflowError):
        pass
    digits = "".join(c for c in str(p) if c.isdigit())
    return digits[-10:] if len(digits) >= 10 else ""


def norm_mc(m):
    return str(m).upper().replace("MC-", "").replace(".0", "").strip()


DNC_OUTCOMES = {"STOP", "Not interested", "Called me scammer"}

def load_do_not_contact(folder):
    """Lee los CSV exportados del celular (no se publican, la carpeta esta en .gitignore).
    - do_not_contact*.csv: todas las filas son do-not-contact
    - cualquier otro CSV con columna 'outcome': filas con STOP / Not interested / Called me scammer
    """
    phones, mcs = set(), set()
    for path in sorted(glob.glob(os.path.join(folder, "*.csv"))):
        try:
            d = pd.read_csv(path, dtype=str, encoding="utf-8-sig").fillna("")
        except Exception as e:
            print(f"  (no se pudo leer {path}: {e})")
            continue
        if os.path.basename(path).lower().startswith("do_not_contact"):
            rows = d
        elif "outcome" in d.columns:
            rows = d[d["outcome"].isin(DNC_OUTCOMES)]
        else:
            continue
        for _, r in rows.iterrows():
            ph = norm_phone(r.get("phone", ""))
            mc = norm_mc(r.get("mc", ""))
            if ph: phones.add(ph)
            if mc: mcs.add(mc)
    return phones, mcs


dnc_phones, dnc_mcs = load_do_not_contact(args.private_dir)
print(f"Do-not-contact: {len(dnc_phones)} telefonos, {len(dnc_mcs)} MCs (de {args.private_dir}/)")

print(f"Buscando carriers registrados entre {fecha_min} y {fecha_max}, "
      f"{args.min_trucks}-{args.max_trucks} camion(es)...")

base_url = "https://data.transportation.gov/resource/az4n-8mr2.csv"
CANADA_PROVINCES = "('ON','QC','BC','AB','MB','SK','NS','NB','PE','NL','NT','YT','NU')"

where = (
    f"add_date >= '{fecha_min}' "
    f"AND add_date <= '{fecha_max}' "
    f"AND status_code = 'A' "
    f"AND docket1prefix = 'MC' "
    f"AND docket1_status_code = 'A' "
    f"AND interstate_beyond_100_miles != '0' "
    f"AND power_units::number >= {args.min_trucks} "
    f"AND power_units::number <= {args.max_trucks} "
    f"AND phy_state NOT IN {CANADA_PROVINCES} "
    f"AND (crgo_genfreight = 'X' OR crgo_metalsheet = 'X' OR crgo_bldgmat = 'X' OR crgo_construct = 'X') "
    f"AND (owntract::number > 0 OR trmtract::number > 0 OR trptract::number > 0 "
    f"OR owntrail::number > 0 OR trmtrail::number > 0 OR trptrail::number > 0)"
)

if args.dry_run:
    print("\n[DRY RUN] Filtro FMCSA:\n  " + where)
    print(f"[DRY RUN] Verificaria hasta {args.batch} carriers contra Motus. No se descargo nada.")
    sys.exit(0)

todos = []
offset = 0
while True:
    params = {
        "$select": "dot_number,docket1,legal_name,phone,phy_street,phy_city,phy_state,phy_zip,power_units,add_date,"
                   "crgo_genfreight,crgo_metalsheet,crgo_bldgmat,crgo_drybulk,crgo_construct,"
                   "owntract,trmtract,trptract,owntrail,trmtrail,trptrail,owntruck,trmtruck,trptruck",
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

# Saltar los ya verificados, telefonos repetidos y do-not-contact; tomar los siguientes N
telefonos_previos = set()
if df_previo is not None:
    telefonos_previos = set(df_previo["Telefono"].apply(norm_phone)) - {""}
df["_mc"] = df["docket1"].astype(str).str.replace(".0", "", regex=False)
df["_phone_norm"] = df["phone"].apply(norm_phone)
es_dnc = df["_mc"].isin(dnc_mcs) | df["_phone_norm"].isin(dnc_phones)
print(f"Saltados por do-not-contact: {int(es_dnc.sum())}")
df_pendientes = df[~df["_mc"].isin(ya_verificados)
                   & ~df["_phone_norm"].isin(telefonos_previos)
                   & ~es_dnc].copy()
df_pendientes = df_pendientes.drop_duplicates(subset=["_phone_norm"], keep="first")
df_check = df_pendientes.head(args.batch).copy()
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
