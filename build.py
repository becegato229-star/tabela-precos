"""
MUBEC — Gerador de Tabela de Preços
Lê os arquivos Excel em /dados e gera index.html em /docs
"""
import pandas as pd
import json
import re
import base64
from pathlib import Path

# ── Caminhos ─────────────────────────────────────────────────────────────────
BASE       = Path(__file__).parent
DADOS      = BASE / "dados"
DOCS       = BASE / "docs"
DOCS.mkdir(exist_ok=True)

FILE_SUPORTES  = DADOS / "precos_suportes.xlsx"
FILE_OUTROS    = DADOS / "precos_outros.xlsx"
FILE_EMBS      = DADOS / "embalagens.xlsx"
FILE_LOGO      = DADOS / "logo.jpg"

# ── Helpers ───────────────────────────────────────────────────────────────────
def to_float(v):
    try:
        s = str(v).strip().replace("-", "").strip()
        return float(s) if s else None
    except:
        return None

def safe_cod(c):
    try: return int(float(str(c)))
    except: return None

def normalize_m(s):
    s = re.sub(r"\s*LINHA LEVE\s*", "", s, flags=re.I)
    s = re.sub(r"(\d)X(\s|$)", r"\1 X\2", s)
    return re.sub(r"\s+", " ", s).strip().upper()

def clean_m(s):
    return re.sub(r"\s*LINHA LEVE\s*", "", s, flags=re.I).strip()

def rename(t):
    """Apply brand text substitutions."""
    t = re.sub(r"\bBRAÇADEIRA\b",  "ABRAÇADEIRA", t)
    t = re.sub(r"\bOLHAL\b",       "PARAFUSO OLHAL", t)
    t = re.sub(r"\bG\. FOGO\b",    "GF", t)
    t = re.sub(r"\bZINCAD[OA]\b",  "GE", t)
    return t

def fmt(v):
    try:
        n = float(v)
        i = f"{int(abs(n)):,}".replace(",", ".")
        d = f"{abs(n):.2f}".split(".")[1]
        s = "-" if n < 0 else ""
        return f'<span class="curr">R$</span><span class="amt">{s}{i},{d}</span>'
    except:
        return '<span class="curr"></span><span class="amt">—</span>'

def fmt_emb(v):
    try:
        n = float(v)
        return str(int(n)) if n == int(n) else str(n)
    except:
        return "A DEFINIR"

# ── 1. Carregar precos_suportes.xlsx ─────────────────────────────────────────
df_s = pd.read_excel(FILE_SUPORTES)
df_s.columns = ["codigo", "item", "preco_normal", "preco_minimo"]
df_s["codigo"] = df_s["codigo"].apply(lambda x: int(x) if pd.notna(x) else "")

def get_group(item):
    m = re.match(r"^((?:[A-Z]+\s+)+?)(?:\d+\s+)?(?:\d+/\d+|M\d+|\d+\s+X\b)", item)
    if m: return m.group(1).strip()
    return item.split(" X ")[0].strip()

def get_medida(item, grupo):
    remainder = item[len(grupo):].strip() if item.startswith(grupo) else item
    return remainder.strip() or item

df_s["grupo"]  = df_s["item"].apply(get_group)
df_s["medida"] = df_s.apply(lambda r: get_medida(r["item"], r["grupo"]), axis=1)

old_groups = {}
for g, grp in df_s.groupby("grupo"):
    old_groups[g] = grp[["codigo","medida","preco_normal","preco_minimo"]].to_dict("records")

# ── 2. Carregar precos_outros.xlsx ────────────────────────────────────────────
raw = pd.read_excel(FILE_OUTROS, sheet_name="Plan1", header=None)

new_groups = {}

def parse_price_cell(val):
    if pd.isna(val): return None
    s = str(val).strip()
    s = re.sub(r"R\$\s*", "", s)
    s = re.sub(r"\s+", "", s)
    s = s.replace(".", "").replace(",", ".")
    try: return float(s)
    except: return None

def process_side(col_cat, col_cod, col_med, col_price, col_unit_check=None):
    current_cat = None
    current_unit = "PC"
    rows = []

    for idx, row in raw.iterrows():
        c0 = str(row[col_cat]).strip() if pd.notna(row[col_cat]) else ""
        c1 = str(row[col_med]).strip() if pd.notna(row[col_med]) else ""
        c2 = row[col_price] if pd.notna(row[col_price]) else None
        cod = row[col_cod] if pd.notna(row[col_cod]) else None

        if c0 and c0 not in ("nan","COD") and c1 in ("","nan") and c2 is None:
            if current_cat and rows:
                if current_cat not in new_groups:
                    new_groups[current_cat] = {"unit": current_unit, "items": []}
                new_groups[current_cat]["items"].extend(rows)
            current_cat = c0; rows = []
        elif c0 == "COD" or c1 == "Medida":
            ph = str(c2).strip() if c2 is not None else ""
            current_unit = "CT" if "CT" in ph else "PC"
        elif current_cat and c1 and c1 != "nan" and c2 is not None:
            p = parse_price_cell(c2)
            if p:
                try: c = int(float(str(cod)))
                except: c = ""
                rows.append({"codigo": c, "medida": c1, "preco": p})

    if current_cat and rows:
        if current_cat not in new_groups:
            new_groups[current_cat] = {"unit": current_unit, "items": []}
        new_groups[current_cat]["items"].extend(rows)

process_side(0, 0, 1, 2)
process_side(7, 7, 8, 9)

# ── 3. Carregar embalagens ─────────────────────────────────────────────────────
raw_emb = pd.read_excel(FILE_EMBS, sheet_name="Plan1", header=None)
emb_lookup  = {}   # cod -> qty
medida_emb  = {}   # normalized_medida -> qty

def process_emb_side(col_cod, col_emb):
    for _, row in raw_emb.iterrows():
        cod_val = row[col_cod]; emb_val = row[col_emb]
        if pd.isna(cod_val) or pd.isna(emb_val): continue
        try:
            cod = int(float(str(cod_val)))
            if cod == 0: continue
            emb_lookup[cod] = float(str(emb_val).strip().replace(",","."))
        except: pass

process_emb_side(0, 3)
process_emb_side(7, 10)
process_emb_side(5, 8)

# Build medida->emb from groups that have full data
donor_groups = ["SUPORTE RT NAT", "SUPORTE RT GE", "ABRACADEIRA U GE",
                "KIT ABRACADEIRA COMPLETO U GE", "ABRACADEIRA U INOX",
                "KIT ABRACADEIRA COMPLETO U INOX"]
for g in donor_groups:
    for r in old_groups.get(g, []):
        c = safe_cod(r["codigo"])
        if c and c in emb_lookup:
            norm = normalize_m(r["medida"])
            if norm not in medida_emb:
                medida_emb[norm] = emb_lookup[c]

# Fix generic 1/2 entries
medida_emb.setdefault("1/2 X 1000", 25.0)
medida_emb.setdefault("1/2 X 3000", 10.0)

INOX_ABRAC = {"ABRACADEIRA U INOX", "KIT ABRACADEIRA COMPLETO U INOX"}

def get_emb(cod, med, force_indef=False):
    if force_indef: return "A DEFINIR"
    c = safe_cod(cod)
    if c and c in emb_lookup:
        return fmt_emb(emb_lookup[c])
    norm = normalize_m(med)
    if norm in medida_emb:
        return fmt_emb(medida_emb[norm])
    return "A DEFINIR"

# ── 4. Logo em base64 ─────────────────────────────────────────────────────────
with open(FILE_LOGO, "rb") as f:
    logo_b64 = base64.b64encode(f.read()).decode()

# ── 5. Layout ─────────────────────────────────────────────────────────────────
LAYOUT = [
    ("SEC",  "SUPORTES"),
    ("ROW",  "SUPORTE RT NAT",                           "SUPORTE RT GE"),
    ("ROW",  "SUPORTE RT GF",                            "SUPORTE RT INOX"),
    ("SEC",  "ABRAÇADEIRAS U"),
    ("ROW",  "ABRACADEIRA U NAT",                        "ABRACADEIRA U GE"),
    ("ROW",  None,                                       "KIT ABRACADEIRA COMPLETO U GE"),
    ("ROW",  "ABRACADEIRA U GF",                         "ABRACADEIRA U INOX"),
    ("ROW",  "KIT ABRACADEIRA COMPLETO U GF",            "KIT ABRACADEIRA COMPLETO U INOX"),
    ("ROW",  'BRAÇADEIRA TIPO "D" C/ PARAFUSO ZINCADA',  'BRAÇADEIRA TIPO "D" C/ CUNHA ZINCADA'),
    ("ROW",  "BRAÇADEIRA TIPO ECONÔMICA ZINCADA",        'BRAÇADEIRA TIPO "U" PERFIL C/ PARAUSO'),
    ("ROW",  "BRAÇADEIRA TIPO UNIÃO HORIZINTAL ZINCADA", "BRAÇADEIRA TIPO UNIÃO VERTICAL ZINCADA"),
    ("FULL", "BRAÇADEIRA TIPO ÔMEGA ZINCADA"),
    ("SEC",  "OLHAIS"),
    ("ROW",  "OLHAL NAT",                                "OLHAL GE"),
    ("ROW",  "OLHAL GF",                                 "PARAFUSO OLHAL ZINCADO"),
    ("FULL", "PARAFUSO OLHAL ZINCADO C/ SOLDA"),
    ("SEC",  "PORCAS"),
    ("ROW",  "PORCA SEXTAVADA ZINCADA",                  "PORCA SEXTAVADA G. FOGO"),
    ("FULL", "PORCA SEXTAVADA INOX 304"),
    ("ROW",  "PORCA LOSANGULAR C/ ROSCA ZINCADA",        "PORCA LOSANGULAR C/ MOLA ZINCADA"),
    ("FULL", "PORCA LOSANGULAR C/ PINO ZINCADO"),
    ("SEC",  "ARRUELAS"),
    ("ROW",  "ARRUELA LISA ZINCADA",                     "ARRUELA LISA G. FOGO"),
    ("FULL", "ARRUELA PRESSÃO ZINCADA"),
    ("SEC",  "PROLONGADORES"),
    ("FULL", "PROLONGADOR SEXTAVADO ZINCADO"),
    ("SEC",  "OUTROS"),
    ("ROW",  'CHUMBADOR "CB" C/ PARAFUSO ZINCADO',       "JAQUETA E CONE ZINCADO"),
    ("ROW",  "PARAFUSO LENTILHA FENDA ZINCADO",          "PARAFUSO LENTILHA TRAVA ZINCADO"),
    ("FULL", "PARAFUSO SEXTAVADO ZINCADO"),
]

# ── 6. HTML builders ──────────────────────────────────────────────────────────
COLGROUP = ('<colgroup>'
            '<col style="width:58px"><col>'
            '<col style="width:90px"><col style="width:112px">'
            '</colgroup>')

def make_rows(group_name, rows_src, price_fn, force_indef=False):
    html = []
    for r in rows_src:
        cod  = str(r["codigo"]) if r.get("codigo") else "—"
        med  = clean_m(r.get("medida",""))
        desc = rename(f"{group_name} {med}".strip().upper())
        emb  = get_emb(r.get("codigo"), r.get("medida",""), force_indef=force_indef)
        ph   = price_fn(r)
        indef = " indefinido" if emb == "A DEFINIR" else ""
        html.append(
            f'<tr>'
            f'<td class="td-cod">{cod}</td>'
            f'<td class="td-desc">{desc}</td>'
            f'<td class="td-emb{indef}">{emb}</td>'
            f'<td class="td-price">{ph}</td>'
            f'</tr>'
        )
    return html

def build_table(group_name, price_key):
    if group_name is None: return ""
    is_suporte   = group_name.startswith("SUPORTE RT")
    force_indef  = group_name in INOX_ABRAC
    unit = "/PC" if is_suporte else None

    if group_name in old_groups:
        src = old_groups[group_name]
        if unit is None:
            unit = "/CT" if any("1000" in r.get("medida","") or "3000" in r.get("medida","") for r in src) else "/PC"
        rows_html = make_rows(group_name, src, lambda r: fmt(r[price_key]), force_indef)
    elif group_name in new_groups:
        data = new_groups[group_name]
        if unit is None: unit = f'/{data["unit"]}'
        rows_html = make_rows(group_name, data["items"], lambda r: fmt(r["preco"]), force_indef)
    else:
        return ""

    title = rename(group_name.upper())
    thead = (f'<thead><tr>'
             f'<th class="th-cod">COD</th><th class="th-desc">DESCRIÇÃO</th>'
             f'<th class="th-emb">Embalagem</th><th class="th-price">PREÇO {unit}</th>'
             f'</tr></thead>')
    return (f'<div class="cat-block">'
            f'<div class="cat-title">{title}</div>'
            f'<table>{COLGROUP}{thead}<tbody>{"".join(rows_html)}</tbody></table>'
            f'</div>')

def build_page(price_key):
    parts = []
    for entry in LAYOUT:
        kind = entry[0]
        if kind == "SEC":
            parts.append(f'<div class="sec-header">{entry[1]}</div>')
        elif kind == "FULL":
            tbl = build_table(entry[1], price_key)
            if tbl: parts.append(f'<div class="full-row">{tbl}</div>')
        elif kind == "ROW":
            left  = build_table(entry[1], price_key)
            right = build_table(entry[2], price_key)
            lc = f'<div class="half">{left}</div>'  if left  else '<div class="half empty"></div>'
            rc = f'<div class="half">{right}</div>' if right else '<div class="half empty"></div>'
            parts.append(f'<div class="two-row">{lc}{rc}</div>')
    return "".join(parts)

page_normal = build_page("preco_normal")
page_minimo  = build_page("preco_minimo")

total_items = sum(len(v) for v in old_groups.values()) + sum(len(v["items"]) for v in new_groups.values())
total_groups = len(old_groups) + len(new_groups)

# ── 7. HTML completo ──────────────────────────────────────────────────────────
HTML = f"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>MUBEC — Tabela de Preços</title>
<style>
  @import url("https://fonts.googleapis.com/css2?family=Barlow:wght@400;600;700;900&display=swap");
  :root{{--verde:#48D597;--cinza:#717C7D;--preto:#000;--borda:#ccc;--th-bg:#f0f0f0;}}
  *{{margin:0;padding:0;box-sizing:border-box;}}
  body{{font-family:"Barlow",Arial,sans-serif;background:#fff;color:#000;font-size:13px;padding:24px 32px 40px;}}
  .page-header{{display:flex;align-items:center;justify-content:space-between;border-bottom:3px solid var(--preto);padding-bottom:14px;margin-bottom:20px;}}
  .logo img{{height:44px;}}
  .header-right{{text-align:right;font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--cinza);line-height:1.6;}}
  .header-right .big{{font-size:14px;color:#000;}}
  .tabs{{display:flex;gap:4px;margin-bottom:20px;}}
  .tab-btn{{padding:9px 28px;font-family:"Barlow",Arial,sans-serif;font-weight:900;font-size:12px;letter-spacing:.12em;text-transform:uppercase;cursor:pointer;border:2px solid var(--preto);background:#fff;color:var(--preto);transition:all .15s;}}
  .tab-btn.active{{background:var(--preto);color:var(--verde);}}
  .tab-btn:hover:not(.active){{background:var(--th-bg);}}
  .page{{display:none;}}
  /* Search */
  .search-bar{{display:flex;align-items:center;gap:16px;margin-bottom:20px;}}
  .search-wrap{{display:flex;align-items:center;flex:1;border:2px solid var(--preto);background:#fff;transition:border-color .15s;}}
  .search-wrap:focus-within{{border-color:var(--verde);}}
  .search-icon{{width:16px;height:16px;margin:0 10px;color:var(--cinza);flex-shrink:0;}}
  #search-input{{flex:1;border:none;outline:none;font-family:"Barlow",Arial,sans-serif;font-size:13px;padding:9px 0;background:transparent;color:var(--preto);}}
  #search-input::placeholder{{color:#aaa;}}
  #search-clear{{border:none;background:none;cursor:pointer;padding:0 12px;font-size:13px;color:var(--cinza);display:none;line-height:1;}}
  #search-clear:hover{{color:var(--preto);}}
  #search-count{{font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--cinza);white-space:nowrap;min-width:80px;text-align:right;}}
  .two-row.hidden,.full-row.hidden,.cat-block.hidden,tr.hidden{{display:none;}}
  mark{{background:#d4f7e8;color:inherit;border-radius:2px;padding:0 1px;}}
  /* Layout */
  .sec-header{{background:var(--verde);color:var(--preto);font-weight:900;font-size:13px;letter-spacing:.18em;text-transform:uppercase;padding:10px 14px;margin-top:24px;margin-bottom:14px;}}
  .sec-header:first-child{{margin-top:0;}}
  .two-row{{display:grid;grid-template-columns:1fr 1fr;gap:0 28px;margin-bottom:20px;align-items:start;}}
  .half{{min-width:0;}}
  .full-row{{margin-bottom:20px;}}
  /* Tables */
  .cat-block{{border:1px solid var(--borda);}}
  .cat-title{{background:var(--preto);color:#fff;font-weight:900;font-size:11px;letter-spacing:.12em;text-transform:uppercase;padding:7px 10px;border-bottom:2px solid var(--verde);}}
  table{{width:100%;border-collapse:collapse;table-layout:fixed;}}
  thead tr{{background:var(--th-bg);}}
  thead th{{padding:6px 8px;font-weight:700;font-size:11px;letter-spacing:.06em;text-transform:uppercase;color:var(--cinza);border-bottom:1px solid var(--borda);white-space:nowrap;overflow:hidden;}}
  th.th-cod{{text-align:center;}}th.th-desc{{text-align:left;}}th.th-emb{{text-align:center;}}th.th-price{{text-align:right;}}
  tbody tr{{border-bottom:1px solid #e8e8e8;}}
  tbody tr:last-child{{border-bottom:none;}}
  tbody tr:nth-child(even){{background:#fafafa;}}
  tbody tr:hover{{background:#f0fdf7;}}
  td{{padding:5px 8px;vertical-align:middle;font-size:12px;}}
  td.td-cod{{color:var(--cinza);font-weight:600;font-size:11px;text-align:center;white-space:nowrap;overflow:hidden;}}
  td.td-desc{{font-weight:400;text-align:left;text-transform:uppercase;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}
  td.td-emb{{text-align:center;font-weight:600;font-size:11px;color:#333;white-space:nowrap;}}
  td.td-emb.indefinido{{color:#bbb;font-style:italic;font-size:10px;font-weight:400;}}
  td.td-price{{font-weight:700;white-space:nowrap;color:var(--preto);padding-left:6px;padding-right:8px;}}
  td.td-price .curr{{float:left;color:var(--cinza);font-weight:600;font-size:11px;padding-right:4px;line-height:inherit;}}
  td.td-price .amt{{display:block;text-align:right;}}
  .footer{{margin-top:32px;padding-top:12px;border-top:2px solid var(--preto);display:flex;justify-content:space-between;font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--cinza);}}
  .verde-dot{{display:inline-block;width:8px;height:8px;background:var(--verde);border-radius:50%;margin-right:4px;vertical-align:middle;}}
</style>
</head>
<body>
<div class="page-header">
  <div class="logo"><img src="data:image/jpeg;base64,{logo_b64}" alt="MUBEC"></div>
  <div class="header-right">Tabela de Preços<br><span class="big">{total_items} Itens · {total_groups} Categorias</span></div>
</div>
<div class="tabs">
  <button class="tab-btn active" onclick="showPage('normal',this)">Preço Normal</button>
  <button class="tab-btn" onclick="showPage('minimo',this)">Preço Mínimo</button>
</div>
<div class="search-bar">
  <div class="search-wrap">
    <svg class="search-icon" viewBox="0 0 20 20" fill="none" xmlns="http://www.w3.org/2000/svg">
      <circle cx="8.5" cy="8.5" r="5.5" stroke="currentColor" stroke-width="1.8"/>
      <line x1="12.5" y1="12.5" x2="17" y2="17" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/>
    </svg>
    <input type="text" id="search-input" placeholder="Buscar por código ou descrição…" autocomplete="off">
    <button id="search-clear" title="Limpar">✕</button>
  </div>
  <div id="search-count"></div>
</div>
<div id="page-normal" class="page">{page_normal}</div>
<div id="page-minimo" class="page">{page_minimo}</div>
<div class="footer">
  <span><span class="verde-dot"></span>MUBEC — Tabela de Preços · Abril 2026</span>
  <span>Uso Interno</span>
</div>
<script>
document.getElementById('page-normal').style.display='block';
function showPage(id,btn){{
  document.querySelectorAll('.page').forEach(p=>p.style.display='none');
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));
  document.getElementById('page-'+id).style.display='block';
  btn.classList.add('active');
}}
(function(){{
  const input=document.getElementById('search-input');
  const clearBtn=document.getElementById('search-clear');
  const countEl=document.getElementById('search-count');
  function getPage(){{for(const p of document.querySelectorAll('.page'))if(p.style.display!=='none')return p;return document.getElementById('page-normal');}}
  function escRe(s){{return s.replace(/[.*+?^${{}}()|[\\]\\\\]/g,'\\\\$&');}}
  function highlight(el,re){{
    const walker=document.createTreeWalker(el,NodeFilter.SHOW_TEXT,null);
    const nodes=[];let n;while((n=walker.nextNode()))nodes.push(n);
    nodes.forEach(node=>{{if(!re.test(node.textContent))return;const frag=document.createDocumentFragment();let last=0,m;re.lastIndex=0;const txt=node.textContent;while((m=re.exec(txt))!==null){{if(m.index>last)frag.appendChild(document.createTextNode(txt.slice(last,m.index)));const mk=document.createElement('mark');mk.textContent=m[0];frag.appendChild(mk);last=re.lastIndex;if(!re.global)break;}}if(last<txt.length)frag.appendChild(document.createTextNode(txt.slice(last)));node.parentNode.replaceChild(frag,node);}});
  }}
  function doSearch(){{
    const q=input.value.trim();clearBtn.style.display=q?'block':'none';
    const page=getPage();
    page.querySelectorAll('mark').forEach(m=>m.replaceWith(document.createTextNode(m.textContent)));
    page.normalize();page.querySelectorAll('.hidden').forEach(el=>el.classList.remove('hidden'));
    if(!q){{countEl.textContent='';return;}}
    const re=new RegExp(escRe(q),'gi');let total=0;
    page.querySelectorAll('tbody tr').forEach(tr=>{{
      const cod=tr.querySelector('.td-cod');const desc=tr.querySelector('.td-desc');
      const ct=cod?cod.textContent:'';const dt=desc?desc.textContent:'';
      re.lastIndex=0;const ok=re.test(ct)||(re.lastIndex=0,re.test(dt));re.lastIndex=0;
      if(ok){{total++;if(cod)highlight(cod,new RegExp(escRe(q),'gi'));if(desc)highlight(desc,new RegExp(escRe(q),'gi'));tr.classList.remove('hidden');}}
      else tr.classList.add('hidden');
    }});
    page.querySelectorAll('.cat-block').forEach(b=>{{b.classList.toggle('hidden',b.querySelectorAll('tbody tr:not(.hidden)').length===0);}});
    page.querySelectorAll('.two-row').forEach(r=>{{r.classList.toggle('hidden',[...r.querySelectorAll('.cat-block')].filter(b=>!b.classList.contains('hidden')).length===0);}});
    page.querySelectorAll('.full-row').forEach(r=>{{r.classList.toggle('hidden',[...r.querySelectorAll('.cat-block')].filter(b=>!b.classList.contains('hidden')).length===0);}});
    page.querySelectorAll('.sec-header').forEach(sec=>{{let el=sec.nextElementSibling,any=false;while(el&&!el.classList.contains('sec-header')){{if(!el.classList.contains('hidden')&&(el.classList.contains('two-row')||el.classList.contains('full-row'))){{any=true;break;}}el=el.nextElementSibling;}}sec.classList.toggle('hidden',!any);}});
    countEl.textContent=total===0?'Sem resultados':total===1?'1 item':total+' itens';
  }}
  input.addEventListener('input',doSearch);
  clearBtn.addEventListener('click',()=>{{input.value='';doSearch();input.focus();}});
  document.querySelectorAll('.tab-btn').forEach(btn=>btn.addEventListener('click',()=>setTimeout(doSearch,50)));
}})();
</script>
</body>
</html>"""

out = DOCS / "index.html"
out.write_text(HTML, encoding="utf-8")
print(f"✅ Gerado: {out} ({len(HTML):,} bytes)")
