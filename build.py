"""
MUBEC — Gerador de Tabela de Preços
Lê precos_todos.xlsx em /dados e gera index.html em /docs
"""
import pandas as pd
import re
import base64
from pathlib import Path

BASE  = Path(__file__).parent
DADOS = BASE / "dados"
DOCS  = BASE / "docs"
DOCS.mkdir(exist_ok=True)

FILE_TODOS = DADOS / "precos_todos.xlsx"
FILE_LOGO  = DADOS / "logo.jpg"

def clean_m(s):
    return re.sub(r"\s*LINHA LEVE\s*", "", str(s), flags=re.I).strip()

def rename(t):
    t = re.sub(r"\bBRAÇADEIRA\b",  "ABRAÇADEIRA", t)
    t = re.sub(r"\bOLHAL\b",       "PARAFUSO OLHAL", t)
    t = re.sub(r"\bG\. FOGO\b",    "GF", t)
    t = re.sub(r"\bZINCAD[OA]\b",  "GE", t)
    return t

def fmt_price(v):
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
        if pd.isna(v): return None
        n = float(v)
        return str(int(n)) if n == int(n) else str(n)
    except:
        return None

def fmt_ncm(v):
    try:
        if pd.isna(v): return "—"
        n = int(float(str(v)))
        s = f"{n:08d}"
        return f"{s[:4]}.{s[4:6]}.{s[6:]}"
    except:
        return "—"

def fmt_ipi(v):
    try:
        if pd.isna(v): return "—"
        n = float(v)
        return f"{n*100:.1f}%".replace(".", ",")
    except:
        return "—"

KNOWN_PREFIXES = [
    "KIT ABRACADEIRA D C/ PARAFUSO GE",
    "ABRACADEIRA D C/ CUNHA GE",
    "KIT ABRACADEIRA U PERFIL C/ PARAFUSO GE",
    "PORCA LOSANGULAR C/ ROSCA GE",
    "PORCA LOSANGULAR C/ MOLA GE",
    "PORCA LOSANGULAR C/ PINO GE",
]

def get_group(item):
    for prefix in KNOWN_PREFIXES:
        if item.startswith(prefix):
            return prefix
    m = re.match(r"^((?:[A-Z]+\s+)+?)(?:\d+\s+)?(?:\d+/\d+|M\d+|\d+\s+X\b|\d+\s+\d+/\d+|\b\d+\b(?:\s+(?:NC|WW|X\b))?)", item)
    if m: return m.group(1).strip()
    return item.split(" X ")[0].strip()

def get_medida(item, grupo):
    remainder = item[len(grupo):].strip() if item.startswith(grupo) else item
    return remainder.strip() or item

# ── Load ──────────────────────────────────────────────────────────────────────
df = pd.read_excel(FILE_TODOS, sheet_name="Sheet2")
df.columns = ["codigo","item","preco_normal","preco_minimo","embalagem","ncm","ipi"]
df["codigo"] = df["codigo"].apply(lambda x: int(x) if pd.notna(x) else "")
df["grupo"]  = df["item"].apply(get_group)
df["medida"] = df.apply(lambda r: get_medida(r["item"], r["grupo"]), axis=1)

all_groups = {}
for g, grp in df.groupby("grupo"):
    all_groups[g] = grp.to_dict("records")

with open(FILE_LOGO, "rb") as f:
    logo_b64 = base64.b64encode(f.read()).decode()

# ── Layout ────────────────────────────────────────────────────────────────────
LAYOUT = [
    ("SEC",  "SUPORTES"),
    ("ROW",  "SUPORTE RT NAT",                           "SUPORTE RT GE"),
    ("ROW",  "SUPORTE RT GF",                            "SUPORTE RT INOX"),
    ("SEC",  "ABRAÇADEIRAS U"),
    ("ROW",  "ABRACADEIRA U NAT",                        "ABRACADEIRA U GE"),
    ("FULL", "KIT ABRACADEIRA COMPLETO U GE"),
    ("ROW",  "ABRACADEIRA U GF",                         "ABRACADEIRA U INOX"),
    ("ROW",  "KIT ABRACADEIRA COMPLETO U GF",            "KIT ABRACADEIRA COMPLETO U INOX"),
    ("FULL", "KIT ABRACADEIRA U PERFIL C/ PARAFUSO GE"),
    ("SEC",  "ABRAÇADEIRAS PERFIL"),
    ("ROW",  "KIT ABRACADEIRA D C/ PARAFUSO GE",         "ABRACADEIRA D C/ CUNHA GE"),
    ("FULL", "ABRACADEIRA ECONOMICA GE"),
    ("ROW",  "KIT ABRACADEIRA UNIAO HORIZONTAL GE",      "KIT ABRACADEIRA UNIAO VERTICAL GE"),
    ("FULL", "ABRACADEIRA OMEGA GE"),
    ("SEC",  "OLHAIS"),
    ("ROW",  "OLHAL NAT",                                "OLHAL GE"),
    ("ROW",  "OLHAL GF",                                 "PARAFUSO OLHAL GE"),
    ("FULL", "PARAFUSO OLHAL GE C/ SOLDA"),
    ("SEC",  "PORCAS"),
    ("ROW",  "PORCA SEXTAVADA GE",                  "PORCA SEXTAVADA GF"),
    ("ROW",  "PORCA SEXTAVADA GF",                       "PORCA SEXTAVADA INOX"),
    ("ROW",  "PORCA LOSANGULAR C/ ROSCA GE",        "PORCA LOSANGULAR C/ MOLA GE"),
    ("FULL", "PORCA LOSANGULAR C/ PINO GE"),
    ("SEC",  "ARRUELAS"),
    ("ROW",  "ARRUELA LISA GE",                     "ARRUELA LISA GF"),
    ("ROW",  "ARRUELA LISA INOX",                        None),
    ("FULL", "ARRUELA DE PRESSAO GE"),
    ("SEC",  "PROLONGADORES"),
    ("FULL", "PROLONGADOR GE"),
    ("SEC",  "OUTROS"),
    ("ROW",  "CHUMBADOR CB GE",                           "JAQUETA E CONE GE"),
    ("ROW",  "PARAFUSO LENTILHA FENDA GE",               "PARAFUSO LENTILHA TRAVA GE"),
    ("FULL", "PARAFUSO SEXTAVADO GE"),
    ("FULL", "PARAFUSO AUTO BROCANTE GE"),
    ("SEC",  "GRAMPOS"),
    ("ROW",  "GRAMPO C GE",                              "GRAMPO C COMPLETO GE"),
    ("FULL", "BALANCIM P/ GRAMPO C GE"),
    ("SEC",  "SUPORTES ESPECIAIS"),
    ("ROW",  "SUPORTE MEDAJOIST",                        "SUPORTE MEDAJOIST CURTO"),
]

INOX_ABRAC = {"ABRACADEIRA U INOX", "KIT ABRACADEIRA COMPLETO U INOX"}

COLGROUP = ('<colgroup>'
            '<col style="width:52px">'   # COD
            '<col>'                       # DESCRIÇÃO (flex)
            '<col style="width:72px">'   # EMB
            '<col style="width:100px">'  # PREÇO
            '<col style="width:82px">'   # NCM
            '<col style="width:44px">'   # IPI
            '</colgroup>')

def make_rows(group_name, rows, price_key, force_indef=False):
    html = []
    for r in rows:
        cod  = str(r["codigo"]) if r.get("codigo") else "—"
        med  = clean_m(r.get("medida", ""))
        desc = rename(f"{group_name} {med}".strip().upper())
        ph   = fmt_price(r[price_key])
        ncm  = fmt_ncm(r.get("ncm"))
        ipi  = fmt_ipi(r.get("ipi"))
        if force_indef:
            emb_str, emb_cls = "A DEFINIR", " indefinido"
        else:
            emb_str = fmt_emb(r.get("embalagem"))
            if emb_str is None:
                emb_str, emb_cls = "A DEFINIR", " indefinido"
            else:
                emb_cls = ""
        html.append(
            f'<tr>'
            f'<td class="td-cod">{cod}</td>'
            f'<td class="td-desc">{desc}</td>'
            f'<td class="td-emb{emb_cls}">{emb_str}</td>'
            f'<td class="td-price">{ph}</td>'
            f'<td class="td-ncm">{ncm}</td>'
            f'<td class="td-ipi">{ipi}</td>'
            f'</tr>'
        )
    return html

def build_table(group_name, price_key):
    if group_name is None: return ""
    rows = all_groups.get(group_name)
    if not rows:
        for k in all_groups:
            if rename(k.upper()) == rename(group_name.upper()):
                rows = all_groups[k]; break
    if not rows: return ""
    is_suporte  = group_name.startswith("SUPORTE RT")
    force_indef = group_name in INOX_ABRAC
    unit = "/PC" if is_suporte else (
        "/CT" if any("1000" in r.get("medida","") or "3000" in r.get("medida","") for r in rows) else "/PC"
    )
    rows_html = make_rows(group_name, rows, price_key, force_indef)
    title = rename(group_name.upper())
    thead = (f'<thead><tr>'
             f'<th class="th-cod">COD</th>'
             f'<th class="th-desc">DESCRIÇÃO</th>'
             f'<th class="th-emb">Embalagem</th>'
             f'<th class="th-price">PREÇO {unit}</th>'
             f'<th class="th-ncm">NCM</th>'
             f'<th class="th-ipi">IPI</th>'
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
total_items  = len(df)
total_groups = df["grupo"].nunique()

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
  .sec-header{{background:var(--verde);color:var(--preto);font-weight:900;font-size:13px;letter-spacing:.18em;text-transform:uppercase;padding:10px 14px;margin-top:24px;margin-bottom:14px;}}
  .sec-header:first-child{{margin-top:0;}}
  .two-row{{display:grid;grid-template-columns:1fr 1fr;gap:0 16px;margin-bottom:20px;align-items:start;}}
  .half{{min-width:0;}}
  .full-row{{margin-bottom:20px;}}
  .cat-block{{border:1px solid var(--borda);}}
  .cat-title{{background:var(--preto);color:#fff;font-weight:900;font-size:11px;letter-spacing:.12em;text-transform:uppercase;padding:7px 10px;border-bottom:2px solid var(--verde);}}
  table{{width:100%;border-collapse:collapse;table-layout:fixed;}}
  thead tr{{background:var(--th-bg);}}
  thead th{{padding:6px 8px;font-weight:700;font-size:11px;letter-spacing:.06em;text-transform:uppercase;color:var(--cinza);border-bottom:1px solid var(--borda);white-space:nowrap;overflow:hidden;}}
  th.th-cod   {{ text-align:center; }}
  th.th-desc  {{ text-align:left; }}
  th.th-emb   {{ text-align:center; }}
  th.th-price {{ text-align:right; }}
  th.th-ncm   {{ text-align:center; }}
  th.th-ipi   {{ text-align:center; }}
  tbody tr{{border-bottom:1px solid #e8e8e8;}}
  tbody tr:last-child{{border-bottom:none;}}
  tbody tr:nth-child(even){{background:#fafafa;}}
  tbody tr:hover{{background:#f0fdf7;}}
  td{{padding:5px 8px;vertical-align:middle;font-size:12px;}}
  td.td-cod  {{color:var(--cinza);font-weight:600;font-size:11px;text-align:center;white-space:nowrap;overflow:hidden;}}
  td.td-desc {{font-weight:400;text-align:left;text-transform:uppercase;white-space:normal;overflow:hidden;word-break:break-word;}}
  td.td-emb  {{text-align:center;font-weight:600;font-size:11px;color:#333;white-space:nowrap;}}
  td.td-emb.indefinido{{color:#bbb;font-style:italic;font-size:10px;font-weight:400;}}
  td.td-price{{font-weight:700;white-space:nowrap;color:var(--preto);padding-left:6px;padding-right:8px;}}
  td.td-price .curr{{float:left;color:var(--cinza);font-weight:600;font-size:11px;padding-right:4px;line-height:inherit;}}
  td.td-price .amt{{display:block;text-align:right;}}
  td.td-ncm  {{text-align:center;font-size:11px;color:var(--cinza);white-space:nowrap;letter-spacing:.03em;}}
  td.td-ipi  {{text-align:center;font-size:11px;color:var(--cinza);white-space:nowrap;font-weight:600;}}
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
    <svg class="search-icon" viewBox="0 0 20 20" fill="none"><circle cx="8.5" cy="8.5" r="5.5" stroke="currentColor" stroke-width="1.8"/><line x1="12.5" y1="12.5" x2="17" y2="17" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>
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
function showPage(id,btn){{document.querySelectorAll('.page').forEach(p=>p.style.display='none');document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('active'));document.getElementById('page-'+id).style.display='block';btn.classList.add('active');}}
(function(){{const input=document.getElementById('search-input');const clearBtn=document.getElementById('search-clear');const countEl=document.getElementById('search-count');function getPage(){{for(const p of document.querySelectorAll('.page'))if(p.style.display!=='none')return p;return document.getElementById('page-normal');}}function escRe(s){{return s.replace(/[.*+?^${{}}()|[\\]\\\\]/g,'\\\\$&');}}function highlight(el,re){{const walker=document.createTreeWalker(el,NodeFilter.SHOW_TEXT,null);const nodes=[];let n;while((n=walker.nextNode()))nodes.push(n);nodes.forEach(node=>{{if(!re.test(node.textContent))return;const frag=document.createDocumentFragment();let last=0,m;re.lastIndex=0;const txt=node.textContent;while((m=re.exec(txt))!==null){{if(m.index>last)frag.appendChild(document.createTextNode(txt.slice(last,m.index)));const mk=document.createElement('mark');mk.textContent=m[0];frag.appendChild(mk);last=re.lastIndex;if(!re.global)break;}}if(last<txt.length)frag.appendChild(document.createTextNode(txt.slice(last)));node.parentNode.replaceChild(frag,node);}});}}function doSearch(){{const q=input.value.trim();clearBtn.style.display=q?'block':'none';const page=getPage();page.querySelectorAll('mark').forEach(m=>m.replaceWith(document.createTextNode(m.textContent)));page.normalize();page.querySelectorAll('.hidden').forEach(el=>el.classList.remove('hidden'));if(!q){{countEl.textContent='';return;}}const re=new RegExp(escRe(q),'gi');let total=0;page.querySelectorAll('tbody tr').forEach(tr=>{{const cod=tr.querySelector('.td-cod');const desc=tr.querySelector('.td-desc');const ct=cod?cod.textContent:'';const dt=desc?desc.textContent:'';re.lastIndex=0;const ok=re.test(ct)||(re.lastIndex=0,re.test(dt));re.lastIndex=0;if(ok){{total++;if(cod)highlight(cod,new RegExp(escRe(q),'gi'));if(desc)highlight(desc,new RegExp(escRe(q),'gi'));tr.classList.remove('hidden');}}else tr.classList.add('hidden');}});page.querySelectorAll('.cat-block').forEach(b=>{{b.classList.toggle('hidden',b.querySelectorAll('tbody tr:not(.hidden)').length===0);}});page.querySelectorAll('.two-row').forEach(r=>{{r.classList.toggle('hidden',[...r.querySelectorAll('.cat-block')].filter(b=>!b.classList.contains('hidden')).length===0);}});page.querySelectorAll('.full-row').forEach(r=>{{r.classList.toggle('hidden',[...r.querySelectorAll('.cat-block')].filter(b=>!b.classList.contains('hidden')).length===0);}});page.querySelectorAll('.sec-header').forEach(sec=>{{let el=sec.nextElementSibling,any=false;while(el&&!el.classList.contains('sec-header')){{if(!el.classList.contains('hidden')&&(el.classList.contains('two-row')||el.classList.contains('full-row'))){{any=true;break;}}el=el.nextElementSibling;}}sec.classList.toggle('hidden',!any);}});countEl.textContent=total===0?'Sem resultados':total===1?'1 item':total+' itens';}}input.addEventListener('input',doSearch);clearBtn.addEventListener('click',()=>{{input.value='';doSearch();input.focus();}});document.querySelectorAll('.tab-btn').forEach(btn=>btn.addEventListener('click',()=>setTimeout(doSearch,50)));}}
)();
</script>
</body>
</html>"""

out = DOCS / "index.html"
out.write_text(HTML, encoding="utf-8")
print(f"✅ Gerado: {out} ({len(HTML):,} bytes)")
print(f"   {total_items} itens | {total_groups} categorias")
