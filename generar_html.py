"""
Genera un archivo HTML con los leads para usar desde el celular.
Boton "Send Message" abre la app de mensajes con el numero y mensaje pre-escrito.
"""
import pandas as pd
import urllib.parse
import os
import html

ARCHIVO_EXCEL  = "leads_verificados.xlsx"
ARCHIVO_HTML   = "leads_carriers.html"

import re

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

def build_message(company_name):
    return (
        f"Hey, I know you probably get a lot of calls, so I figured a text was better. I came across {company_name} and wanted to reach out about truck dispatching services. "
        "My name's Fernando, I'm an independent dispatcher and I'm currently offering a free trial of my services, no commitment. "
        "If you're not interested, no worries at all, just don't reply and I won't reach out again. "
        "Thanks for your time, God bless you."
    )

print("Leyendo Excel...")
df = pd.read_excel(ARCHIVO_EXCEL, dtype=str)
df = df.fillna("")

# Solo los que tienen telefono
df = df[df["Telefono"].str.strip() != ""].copy()

# Ya viene filtrado del verificador

def clean_phone(p):
    # Si es float (ej: 9012658635.0), convertir a int primero
    try:
        if float(p) == int(float(p)):
            p = str(int(float(p)))
    except (ValueError, TypeError, OverflowError):
        pass
    digits = "".join(c for c in str(p) if c.isdigit())
    return digits[-10:] if len(digits) >= 10 else ""

df["_digits"] = df["Telefono"].apply(clean_phone)
df = df[df["_digits"] != ""].copy()

print(f"Generando HTML con {len(df):,} carriers...")

rows_html = ""
for i, row in df.iterrows():
    digits     = row["_digits"]
    nombre     = row.get("Nombre de la Empresa", "")
    telefono   = row.get("Telefono", "")
    estado     = row.get("Estado", "")
    ciudad     = row.get("Direccion", "").split(",")[1].strip() if "," in row.get("Direccion", "") else ""
    camiones   = row.get("Num Camiones", "")
    carga      = row.get("Tipo de Carga", "")
    fecha      = str(row.get("Fecha Registro", ""))[:10]
    dot        = row.get("MC Number", "")
    usdot      = str(row.get("DOT Number", "")).replace(".0", "").strip()
    motus_link = f"https://motus.dot.gov/customer/{usdot}/account" if usdot and usdot != "nan" else ""
    vehiculo   = row.get("Tipo Vehiculo", "")

    is_dry     = "Dry Van" in carga
    row_class  = "dry" if is_dry else "other"
    badge      = '<span class="badge">Dry Van</span>' if is_dry else ""

    personal_msg = build_message(friendly_name(nombre) if nombre else "your company")
    encoded_msg  = urllib.parse.quote(personal_msg)
    sms_link     = f"sms:+1{digits}?body={encoded_msg}"
    fmt_phone  = f"({digits[:3]}) {digits[3:6]}-{digits[6:]}" if len(digits) == 10 else telefono

    rows_html += f"""
    <div class="card {row_class}" data-dot="{dot}">
      <div class="card-header">
        <span class="company">{nombre}</span>
        {badge}
      </div>
      <div class="card-body">
        <div class="info"><span class="label">Phone</span><span>{fmt_phone}</span></div>
        <div class="info"><span class="label">Location</span><span>{ciudad}, {estado}</span></div>
        <div class="info"><span class="label">Trucks</span><span>{camiones}</span></div>
        <div class="info"><span class="label">Cargo</span><span>{carga}</span></div>
        <div class="info"><span class="label">Registered</span><span>{fecha}</span></div>
        <div class="info"><span class="label">MC</span><span>{dot}</span></div>
        <div class="info"><span class="label">USDOT</span><span>{usdot if usdot and usdot != "nan" else "—"}</span></div>
        <div class="info"><span class="label">Vehicle</span><span>{vehiculo if vehiculo else "—"}</span></div>
      </div>
      <div class="card-footer">
        <a href="{sms_link}" class="btn-msg">💬 Send Message</a>
        <a href="tel:+1{digits}" class="btn-call">📞 Call</a>
      </div>
      {'<a href="' + motus_link + '" class="btn-motus" target="_blank">🔍 Motus</a>' if motus_link else ''}
      <div class="btn-copy-row">
        <button class="btn-copy" onclick="copyText(this, '+1{digits}')">📋 Copy Number</button>
        <button class="btn-copy" onclick="copyText(this, this.dataset.msg)" data-msg="{html.escape(personal_msg, quote=True)}">📋 Copy Message</button>
      </div>
      <button class="btn-done" onclick="toggleContacted(this)">Mark as contacted</button>
    </div>
    """

html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Prime Route Dispatch — Leads</title>
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
    .btn-motus {{
      display: block; width: calc(100% - 28px); margin: 0 14px 6px;
      background: #EBF3FB; color: #1F4E79; text-align: center;
      padding: 8px; border-radius: 8px; font-size: 13px; font-weight: 600;
      text-decoration: none;
    }}
    .btn-copy-row {{
      display: flex; gap: 8px; padding: 0 14px 8px;
    }}
    .btn-copy {{
      flex: 1; padding: 8px; border: none; border-radius: 8px;
      background: #f5f5f5; color: #555; font-size: 13px; font-weight: 600;
      cursor: pointer; text-align: center;
    }}
    .btn-copy.copied {{ background: #E2EFDA; color: #375623; }}
    .btn-done {{
      width: 100%; margin-top: 2px; padding: 8px;
      background: #f0f0f0; color: #888; border: none;
      border-radius: 8px; font-size: 13px; cursor: pointer;
    }}
    .done-label {{
      font-size: 11px; color: #aaa; text-align: center;
      padding: 4px 0 8px; display: none;
    }}
    .card.contacted .done-label {{ display: block; }}

    .no-results {{ text-align: center; padding: 40px 16px; color: #aaa; font-size: 15px; }}

    .list {{ padding: 12px 16px; display: flex; flex-direction: column; gap: 12px; }}
  </style>
</head>
<body>

<div class="topbar">
  <h1>🚛 Prime Route — Leads</h1>
  <span class="count" id="total-count">{len(df):,} carriers</span>
</div>

<div class="search-bar">
  <input type="text" id="search" placeholder="Search by name, state, city..." oninput="filterCards()">
</div>

<div class="filters">
  <button class="filter-btn active" onclick="setFilter('all', this)">All</button>
  <button class="filter-btn" onclick="setFilter('dry', this)">Dry Van ✅</button>
  <button class="filter-btn" onclick="setFilter('other', this)">Other</button>
  <button class="filter-btn" onclick="setFilter('pending', this)">Not Contacted</button>
  <button class="filter-btn" onclick="setFilter('contacted', this)">Contacted ✓</button>
</div>

<div class="progress" id="progress-bar">
  Contacted: <b id="contacted-count">0</b> / <b>{len(df):,}</b>
</div>

<div class="list" id="card-list">
{rows_html}
</div>

<div class="no-results" id="no-results" style="display:none">No results found.</div>

<script>
  let currentFilter = 'all';
  const contacted = new Set(JSON.parse(localStorage.getItem('prd_contacted') || '[]'));

  // On load: restore contacted state from localStorage
  document.querySelectorAll('.card').forEach((card) => {{
    const dot = card.getAttribute('data-dot');
    if (contacted.has(dot)) {{
      card.classList.add('contacted');
      const btn = card.querySelector('.btn-done');
      if (btn) btn.textContent = '✓ Contacted';
    }}
  }});

  updateProgress();

  function copyText(btn, text) {{
    navigator.clipboard.writeText(text).then(() => {{
      const original = btn.textContent;
      btn.textContent = '✓ Copied!';
      btn.classList.add('copied');
      setTimeout(() => {{ btn.textContent = original; btn.classList.remove('copied'); }}, 2000);
    }});
  }}

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
      btn.textContent = '✓ Contacted';
    }}
    localStorage.setItem('prd_contacted', JSON.stringify([...contacted]));
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

with open(ARCHIVO_HTML, "w", encoding="utf-8") as f:
    f.write(html)

ruta = os.path.abspath(ARCHIVO_HTML)
print(f"\nHTML guardado: {ruta}")
print(f"Sube este archivo a Google Drive y abrelo en Chrome desde tu celular.")
os.startfile(ruta)
