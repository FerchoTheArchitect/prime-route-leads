"""
Lote 2: Verifica carriers contra SAFER.
Filtra por nombre (trucking-related) y excluye non-trucking.
Genera un HTML separado en lote2/index.html
"""

import pandas as pd
import requests
import time
import os
from io import StringIO
from datetime import datetime, timedelta

# Rango: 4 a 12 meses atras
fecha_max = (datetime.now() - timedelta(days=120)).strftime("%Y%m%d")
fecha_min = (datetime.now() - timedelta(days=365)).strftime("%Y%m%d")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml",
}

# Keywords
INCLUDE = ["trucking", "freight", "transport", "logistics", "hauling", "carrier",
           "express", "dispatch", "loads", "shipping", "cargo", "movers", "trux",
           "haulage", "trucker", "flatbed", "drayage", "intermodal", "reefer"]
EXCLUDE = ["construction", "towing", "moving", "landscaping", "roofing", "painting",
           "cleaning", "plumbing", "electric", "auto", "car ", "farm", "dental",
           "medical", "remodeling", "renovation", "hvac", "solar", "paving",
           "concrete", "excavat", "demolition", "fencing", "flooring", "handyman",
           "welding", "limo", "taxi", "bus ", "charter", "school", "daycare",
           "restaurant", "catering", "lawn", "tree service", "pest", "junk",
           "appliance", "furniture", "gun", "firearm", "security"]

print(f"Buscando carriers registrados entre {fecha_min} y {fecha_max}...")

base_url = "https://data.transportation.gov/resource/az4n-8mr2.csv"
where = (
    f"add_date >= '{fecha_min}' "
    f"AND add_date <= '{fecha_max}' "
    f"AND status_code = 'A' "
    f"AND docket1prefix = 'MC' "
    f"AND docket1_status_code = 'A' "
    f"AND interstate_beyond_100_miles != '0' "
    f"AND power_units::number >= 1 "
    f"AND power_units::number <= 3"
)

todos = []
offset = 0
while True:
    params = {
        "$select": "dot_number,docket1,legal_name,phone,phy_street,phy_city,phy_state,phy_zip,power_units,add_date,"
                   "crgo_genfreight,crgo_metalsheet,crgo_bldgmat,crgo_drybulk,crgo_construct",
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

# Filtrar por nombre
def name_filter(name):
    name_lower = str(name).lower()
    has_include = any(kw in name_lower for kw in INCLUDE)
    has_exclude = any(kw in name_lower for kw in EXCLUDE)
    return has_include and not has_exclude

df = df[df["legal_name"].apply(name_filter)].copy()
print(f"Filtrado por nombre: {len(df):,}")

# Excluir carriers ya verificados (lote 1)
ARCHIVO_PREVIO = "leads_verificados.xlsx"
ya_verificados = set()
if os.path.exists(ARCHIVO_PREVIO):
    df_previo = pd.read_excel(ARCHIVO_PREVIO)
    ya_verificados = set(df_previo["MC Number"].astype(str).str.replace("MC-", ""))
    print(f"Ya verificados previamente: {len(ya_verificados)}")

df["_mc"] = df["docket1"].astype(str).str.replace(".0", "", regex=False)
df = df[~df["_mc"].isin(ya_verificados)].copy()
print(f"Pendientes de verificar: {len(df):,}")

# Tomar primeros 500
df_check = df.head(500).copy()
print(f"A verificar contra SAFER: {len(df_check)}")


def check_safer(dot_number):
    try:
        resp = requests.get(
            "https://safer.fmcsa.dot.gov/query.asp",
            params={
                "searchtype": "ANY",
                "query_type": "queryCarrierSnapshot",
                "query_param": "USDOT",
                "query_string": str(dot_number),
            },
            headers=HEADERS,
            timeout=15,
        )
        text = resp.text
        if "No records matching" in text:
            return False, "Not found"
        if "AUTHORIZED FOR Property" in text:
            return True, "Authorized"
        if "Entity Type:" in text:
            return False, "Not Authorized"
        return False, "Unknown"
    except:
        return False, "Error"


auth_indices = []
counts = {"Authorized": 0, "Not Authorized": 0, "Not found": 0, "Error": 0, "Unknown": 0}

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

if len(auth_indices) == 0:
    print("\nNingun carrier autorizado encontrado.")
    exit()

df_auth = df.loc[auth_indices].copy()

# Preparar columnas
df_auth["Direccion"] = (
    df_auth["phy_street"].fillna("") + ", " +
    df_auth["phy_city"].fillna("") + ", " +
    df_auth["phy_state"].fillna("") + " " +
    df_auth["phy_zip"].fillna("")
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

df_auth["Tipo de Carga"] = df_auth.apply(tipo_carga, axis=1)
df_auth["MC Number"] = "MC-" + df_auth["docket1"].astype(str).str.replace(".0", "", regex=False)

resultado = df_auth[[
    "MC Number", "legal_name", "phone",
    "Direccion", "phy_state",
    "power_units", "add_date", "Tipo de Carga"
]].copy()

resultado.columns = [
    "MC Number", "Nombre de la Empresa", "Telefono",
    "Direccion", "Estado",
    "Num Camiones", "Fecha Registro", "Tipo de Carga"
]

# Guardar Excel
ARCHIVO_LOTE2 = "leads_lote2.xlsx"
resultado.to_excel(ARCHIVO_LOTE2, index=False)

# ── GENERAR HTML ─────────────────────────────────────────────
import urllib.parse

MESSAGE = (
    "Hi! I'm Fernando, a bilingual truck dispatcher for semi trucks. First 5 loads free, "
    "no strings attached. If this sounds interesting we can jump on a quick call. "
    "If not, no worries at all, I won't bother you again. Have a great day!"
)

encoded_msg = urllib.parse.quote(MESSAGE)

def clean_phone(p):
    try:
        if float(p) == int(float(p)):
            p = str(int(float(p)))
    except (ValueError, TypeError, OverflowError):
        pass
    digits = "".join(c for c in str(p) if c.isdigit())
    return digits[-10:] if len(digits) >= 10 else ""

rows_html = ""
for i, row in resultado.iterrows():
    digits = clean_phone(row["Telefono"])
    if not digits:
        continue
    nombre   = row.get("Nombre de la Empresa", "")
    estado   = row.get("Estado", "")
    ciudad   = row.get("Direccion", "").split(",")[1].strip() if "," in row.get("Direccion", "") else ""
    camiones = row.get("Num Camiones", "")
    carga    = row.get("Tipo de Carga", "")
    fecha    = row.get("Fecha Registro", "")
    mc       = row.get("MC Number", "")
    fmt_phone = f"({digits[:3]}) {digits[3:6]}-{digits[6:]}" if len(digits) == 10 else row["Telefono"]

    is_dry    = "Dry Van" in str(carga)
    row_class = "dry" if is_dry else "other"
    badge     = '<span class="badge">Dry Van</span>' if is_dry else ""

    sms_link  = f"sms:+1{digits}?body={encoded_msg}"

    rows_html += f"""
    <div class="card {row_class}" data-dot="{mc}">
      <div class="card-header">
        <span class="company">{nombre}</span>
        {badge}
      </div>
      <div class="card-body">
        <div class="info"><span class="label">Phone</span><a href="tel:+1{digits}" style="color:#1F4E79;text-decoration:none;font-weight:600">{fmt_phone}</a></div>
        <div class="info"><span class="label">Location</span><span>{ciudad}, {estado}</span></div>
        <div class="info"><span class="label">Trucks</span><span>{camiones}</span></div>
        <div class="info"><span class="label">Cargo</span><span>{carga}</span></div>
        <div class="info"><span class="label">Registered</span><span>{fecha}</span></div>
        <div class="info"><span class="label">MC</span><span>{mc}</span></div>
      </div>
      <div class="card-footer">
        <a href="{sms_link}" class="btn-msg">💬 Send Message</a>
        <a href="tel:+1{digits}" class="btn-call">📞 Call</a>
      </div>
      <button class="btn-done" onclick="toggleContacted(this)">Mark as contacted</button>
    </div>
    """

total = len(resultado)

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Prime Route Dispatch — Leads Lote 2</title>
  <style>
    * {{ box-sizing: border-box; margin: 0; padding: 0; }}
    body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
            background: #f0f2f5; color: #333; }}

    .topbar {{
      background: #1F4E79; color: white; padding: 14px 16px;
      position: sticky; top: 0; z-index: 100;
      display: flex; align-items: center; justify-content: space-between;
    }}
    .topbar h1 {{ font-size: 16px; font-weight: 700; }}
    .topbar .count {{ font-size: 13px; opacity: .8; }}

    .search-bar {{
      background: white; padding: 10px 16px;
      border-bottom: 1px solid #ddd; position: sticky; top: 49px; z-index: 99;
    }}
    .search-bar input {{
      width: 100%; padding: 8px 12px; border: 1px solid #ccc;
      border-radius: 8px; font-size: 15px;
    }}

    .filters {{
      display: flex; gap: 8px; padding: 10px 16px;
      background: white; border-bottom: 1px solid #ddd;
      position: sticky; top: 97px; z-index: 98;
      overflow-x: auto;
    }}
    .filter-btn {{
      padding: 6px 14px; border-radius: 20px; border: 1px solid #ccc;
      background: white; font-size: 13px; white-space: nowrap; cursor: pointer;
    }}
    .filter-btn.active {{ background: #1F4E79; color: white; border-color: #1F4E79; }}

    .progress {{
      background: white; padding: 8px 16px; font-size: 13px; color: #666;
      border-bottom: 1px solid #eee; text-align: center;
    }}
    .progress b {{ color: #1F4E79; }}

    .list {{ padding: 12px 16px; display: flex; flex-direction: column; gap: 12px; }}

    .card {{
      background: white; border-radius: 12px;
      box-shadow: 0 1px 4px rgba(0,0,0,.08); overflow: hidden;
    }}
    .card.dry {{ border-left: 4px solid #70AD47; }}
    .card.other {{ border-left: 4px solid #BDD7EE; }}
    .card.contacted {{ opacity: 0.5; }}

    .card-header {{
      padding: 12px 14px 8px; display: flex;
      align-items: flex-start; justify-content: space-between; gap: 8px;
    }}
    .company {{ font-weight: 700; font-size: 15px; line-height: 1.3; }}
    .badge {{
      background: #E2EFDA; color: #375623; font-size: 11px;
      padding: 2px 8px; border-radius: 10px; white-space: nowrap; font-weight: 600;
    }}

    .card-body {{ padding: 0 14px 8px; }}
    .info {{
      display: flex; justify-content: space-between;
      padding: 3px 0; font-size: 13px; border-bottom: 1px solid #f5f5f5;
    }}
    .label {{ color: #888; font-weight: 500; }}

    .card-footer {{
      display: flex; gap: 8px; padding: 10px 14px;
      border-top: 1px solid #f0f0f0;
    }}
    .btn-msg {{
      flex: 1; background: #1F4E79; color: white; text-align: center;
      padding: 10px; border-radius: 8px; font-size: 14px; font-weight: 600;
      text-decoration: none; display: block;
    }}
    .btn-call {{
      background: #E2EFDA; color: #375623; text-align: center;
      padding: 10px 14px; border-radius: 8px; font-size: 14px; font-weight: 600;
      text-decoration: none; display: block;
    }}
    .btn-done {{
      width: 100%; margin-top: 6px; padding: 8px;
      background: #f0f0f0; color: #888; border: none;
      border-radius: 8px; font-size: 13px; cursor: pointer;
    }}

    .no-results {{ text-align: center; padding: 40px 16px; color: #aaa; font-size: 15px; }}
  </style>
</head>
<body>

<div class="topbar">
  <h1>Prime Route — Lote 2</h1>
  <span class="count" id="total-count">{total:,} carriers</span>
</div>

<div class="search-bar">
  <input type="text" id="search" placeholder="Search by name, state, city..." oninput="filterCards()">
</div>

<div class="filters">
  <button class="filter-btn active" onclick="setFilter('all', this)">All</button>
  <button class="filter-btn" onclick="setFilter('dry', this)">Dry Van</button>
  <button class="filter-btn" onclick="setFilter('other', this)">Other</button>
  <button class="filter-btn" onclick="setFilter('pending', this)">Not Contacted</button>
  <button class="filter-btn" onclick="setFilter('contacted', this)">Contacted</button>
</div>

<div class="progress" id="progress-bar">
  Contacted: <b id="contacted-count">0</b> / <b>{total:,}</b>
</div>

<div class="list" id="card-list">
{rows_html}
</div>

<div class="no-results" id="no-results" style="display:none">No results found.</div>

<script>
  let currentFilter = 'all';
  const contacted = new Set(JSON.parse(localStorage.getItem('prd_contacted_lote2') || '[]'));

  document.querySelectorAll('.card').forEach((card, i) => {{
    card.dataset.id = i;
    if (contacted.has(card.getAttribute('data-dot'))) {{
      card.classList.add('contacted');
      const btn = card.querySelector('.btn-done');
      if (btn) btn.textContent = 'Contacted';
    }}
  }});

  updateProgress();

  function toggleContacted(btn) {{
    const card = btn.closest('.card');
    const dot = card.getAttribute('data-dot');
    if (card.classList.contains('contacted')) {{
      card.classList.remove('contacted');
      contacted.delete(dot);
      btn.textContent = 'Mark as contacted';
    }} else {{
      card.classList.add('contacted');
      contacted.add(dot);
      btn.textContent = 'Contacted';
    }}
    localStorage.setItem('prd_contacted_lote2', JSON.stringify([...contacted]));
    updateProgress();
    filterCards();
  }}

  function updateProgress() {{
    document.getElementById('contacted-count').textContent = contacted.size.toLocaleString();
  }}

  function setFilter(filter, btn) {{
    currentFilter = filter;
    document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    filterCards();
  }}

  function filterCards() {{
    const query = document.getElementById('search').value.toLowerCase();
    let visible = 0;
    document.querySelectorAll('.card').forEach((card, i) => {{
      const text = card.textContent.toLowerCase();
      const isDry = card.classList.contains('dry');
      const isContacted = card.classList.contains('contacted');

      const matchSearch = !query || text.includes(query);
      const matchFilter =
        currentFilter === 'all' ? true :
        currentFilter === 'dry' ? isDry :
        currentFilter === 'other' ? !isDry :
        currentFilter === 'pending' ? !isContacted :
        currentFilter === 'contacted' ? isContacted : true;

      const show = matchSearch && matchFilter;
      card.style.display = show ? '' : 'none';
      if (show) visible++;
    }});
    document.getElementById('total-count').textContent = visible.toLocaleString() + ' carriers';
    document.getElementById('no-results').style.display = visible === 0 ? 'block' : 'none';
  }}
</script>

</body>
</html>"""

# Guardar HTML
os.makedirs("lote2", exist_ok=True)
html_path = os.path.join("lote2", "index.html")
with open(html_path, "w", encoding="utf-8") as f:
    f.write(html)

ruta = os.path.abspath(html_path)
print(f"\nHTML: {ruta}")
print(f"Excel: {os.path.abspath(ARCHIVO_LOTE2)}")
print(f"Total carriers autorizados: {len(resultado):,}")
