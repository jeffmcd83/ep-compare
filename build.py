#!/usr/bin/env python3
"""Regenerate the EP ablation catheter comparison site from the xlsx.
Usage: python3 build.py [xlsx] [guides.json]   -> writes index.html + data/ep.json
"""
import sys, json, re, html, datetime, os, openpyxl
XLSX = sys.argv[1] if len(sys.argv) > 1 else '/cursor/stores/store-d369e846-de0e-411a-8a83-9b0cbb73d3f7/ep_ablation_catheter_comparison.xlsx'
GUIDES = sys.argv[2] if len(sys.argv) > 2 else '/workspace/guides.json'
OUT = os.path.dirname(os.path.abspath(__file__))
SELF_ID = 'ep'
SELF_URL = 'https://jeffmcd83.github.io/ep-compare/'
TITLE = 'EP Ablation Catheter Comparison Guide'
e = lambda s: html.escape('' if s is None else str(s))

wb = openpyxl.load_workbook(XLSX, data_only=True)
def rows(name):
    it = wb[name].iter_rows(values_only=True); hdr = [str(h).strip() if h else '' for h in next(it)]
    return hdr, [dict(zip(hdr, r)) for r in it if any(v not in (None, '') for v in r)]
ch, comp = rows('Comparison'); nh, nonus = rows('Non-US & pipeline'); sh, srcs = rows('Sources'); _, notes = rows('Notes & gaps')
src = {str(s['Source ID']).strip(): s for s in srcs}

def guides_menu():
    try: g = json.load(open(GUIDES))['guides']
    except Exception: g = [{'id': SELF_ID, 'title': 'EP', 'url': SELF_URL, 'status': 'live'}]
    li = []
    for x in g:
        url = x.get('url') or (SELF_URL if x['id'] == SELF_ID else None)
        live = x.get('status') == 'live' or x['id'] == SELF_ID
        if x['id'] == SELF_ID: li.append(f'<li class="cur"><a href="{e(url)}" aria-current="page">{e(x["title"])}</a></li>')
        elif live and url: li.append(f'<li><a href="{e(url)}">{e(x["title"])}</a></li>')
        else: li.append(f'<li class="soon" aria-disabled="true">{e(x["title"])}<small>coming soon</small></li>')
    return '<details id="gd"><summary aria-label="All comparison guides">Guides ▾</summary><ul role="list">' + ''.join(li) + '</ul></details>'

GAP = re.compile(r'\[GAP[^\]]*\]', re.I)
CONF = re.compile(r'conflict|discrepan|\bvs\.?\b|disagree', re.I)
def linkify(t):
    t = e(t)
    t = re.sub(r'(https?://[^\s<]+)', r'<a href="\1" target="_blank" rel="noopener">\1</a>', t)
    t = GAP.sub(lambda m: f'<mark class="gp">{m.group(0)}</mark>', t)
    return t.replace('\n', '<br>')
def srclinks(ids):
    out = []
    for i in re.findall(r'S\d+', str(ids or '')):
        s = src.get(i)
        out.append(f'<a class="sid" href="{e(s["URL"])}" target="_blank" rel="noopener" title="{e(s["Title / description"])}">{i}</a>' if s else i)
    return ' '.join(out)

SKIP = {'ID', 'Source URLs (all for this row)', 'Source IDs - specs', 'Source IDs - regulatory', 'Source IDs - recalls/safety'}
cols = [c for c in ch if c and c not in SKIP]
energies = sorted({re.split(r'[ (:;,]', str(r.get('Energy type') or ''))[0] for r in comp} - {''})
mfrs = sorted({re.sub(r'\s*\(.*', '', str(r.get('Manufacturer') or '')) for r in comp} - {''})

th = ''.join(f'<th data-c="{i}">{e(c)}<span class="ar"></span></th>' for i, c in enumerate(cols)) + '<th>Sources</th>'
trs = []; ngap = nconf = 0
for r in comp:
    notes_txt = str(r.get('Notes / gaps for this row') or '')
    conf = bool(CONF.search(notes_txt)); nconf += conf
    tds = []
    for c in cols:
        v = r.get(c); s = '' if v is None else str(v)
        g = (not s.strip()) or bool(GAP.search(s)); ngap += g
        cls = ' class="gap"' if g else ''
        if c == 'Notes / gaps for this row' and conf: cls = ' class="cf"'
        body = linkify(s) if s.strip() else '<mark class="gp">[GAP]</mark>'
        if c == 'Product (catheter)': body = f'<b>{body}</b>' + (' <span class="cfb" title="Sources disagree; see notes">⚠ conflict</span>' if conf else '')
        tds.append(f'<td{cls}>{body}</td>')
    sl = ''.join(f'<div><small>{k.split(" - ")[1]}:</small> {srclinks(r.get(k))}</div>' for k in ('Source IDs - specs', 'Source IDs - regulatory', 'Source IDs - recalls/safety') if r.get(k))
    tds.append(f'<td class="sl">{sl}</td>')
    en = re.split(r'[ (:;,]', str(r.get('Energy type') or ''))[0]; mf = re.sub(r'\s*\(.*', '', str(r.get('Manufacturer') or ''))
    trs.append(f'<tr data-e="{e(en)}" data-m="{e(mf)}" data-g="{int(any(GAP.search(str(r.get(c) or "")) or not str(r.get(c) or "").strip() for c in cols))}" data-cf="{int(conf)}">' + ''.join(tds) + '</tr>')

nu = ''.join('<tr>' + ''.join(f'<td>{linkify(x.get(c))}</td>' for c in nh) + '</tr>' for x in nonus)
nuh = ''.join(f'<th>{e(c)}</th>' for c in nh)
cats = {}
for n in notes: cats.setdefault(n['Category'], []).append(n['Note'])
nt = ''.join(f'<div class="nc"><h3>{e(k)}</h3><ul>' + ''.join(f'<li class="{"cfn" if k=="Discrepancy" else ""}">{linkify(v)}</li>' for v in vs) + '</ul></div>' for k, vs in cats.items())
sr = ''.join(f'<tr><td><b>{e(s["Source ID"])}</b></td><td>{e(s["Title / description"])}</td><td><a href="{e(s["URL"])}" target="_blank" rel="noopener">link</a></td><td>{e(s.get("Source type"))}</td><td>{e(s.get("Used for device rows"))}</td><td>{e(s.get("Accessed"))}</td></tr>' for s in srcs)
updated = datetime.date.today().isoformat()
opt = lambda xs: ''.join(f'<option>{e(x)}</option>' for x in xs)

crm = open(os.path.join(OUT, 'base.css')).read()
page = f'''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{TITLE}</title><link rel="icon" href="favicon.svg" type="image/svg+xml"><meta name="theme-color" content="#111827"><meta name="color-scheme" content="light">
<style>{crm}</style></head><body>
<header><h1>{TITLE}</h1>{guides_menu()}<span class="pass">{len(comp)} US-approved catheters · updated {updated}</span>
<nav class="tabs"><button data-t="cmp" class="on">Comparison</button><button data-t="nus">Non-US &amp; pipeline</button><button data-t="nts">Notes &amp; gaps</button><button data-t="src">Sources</button></nav></header>
<main>
<section id="cmp" class="pane on">
<div class="flt"><input id="q" placeholder="Filter (product, spec, indication…)" aria-label="Filter"><select id="fe"><option value="">All energy types</option>{opt(energies)}</select><select id="fm"><option value="">All manufacturers</option>{opt(mfrs)}</select>
<label class="tg"><input type="checkbox" id="fg"> Rows with gaps</label><label class="tg"><input type="checkbox" id="fc"> Rows with conflicts</label><span class="cnt" id="n"></span></div>
<p class="lg"><mark class="gp">[GAP]</mark> not found in an authoritative source (no value inferred) · <span class="cfb">⚠ conflict</span> sources disagree; both values shown · click a header to sort</p>
<div class="wrap rwrap"><table id="rt"><thead><tr>{th}</tr></thead><tbody>{''.join(trs)}</tbody></table></div></section>
<section id="nus" class="pane"><div class="nc"><p>Not FDA approved as of the data date; listed for awareness only.</p></div><div class="wrap"><table id="nt"><thead><tr>{nuh}</tr></thead><tbody>{nu}</tbody></table></div></section>
<section id="nts" class="pane">{nt}</section>
<section id="src" class="pane"><div class="wrap rwrap"><table id="st"><thead><tr><th>ID</th><th>Title</th><th>URL</th><th>Type</th><th>Rows</th><th>Accessed</th></tr></thead><tbody>{sr}</tbody></table></div></section>
</main><footer>Public sources only (FDA SSED/PMA/recall databases, manufacturer spec sheets). Not medical advice; verify against current IFU. Updated {updated}.</footer>
<script>
const $=s=>document.querySelector(s),rs=[...document.querySelectorAll('#rt tbody tr')];
document.querySelectorAll('nav.tabs button').forEach(b=>b.onclick=()=>{{document.querySelectorAll('nav.tabs button,.pane').forEach(x=>x.classList.remove('on'));b.classList.add('on');$('#'+b.dataset.t).classList.add('on')}});
function f(){{const q=$('#q').value.toLowerCase(),en=$('#fe').value,m=$('#fm').value,g=$('#fg').checked,c=$('#fc').checked;let n=0;rs.forEach(r=>{{const v=(!q||r.textContent.toLowerCase().includes(q))&&(!en||r.dataset.e==en)&&(!m||r.dataset.m==m)&&(!g||r.dataset.g=='1')&&(!c||r.dataset.cf=='1');r.hidden=!v;n+=v}});$('#n').textContent=n+' of '+rs.length+' shown'}}
['#q','#fe','#fm','#fg','#fc'].forEach(s=>$(s).addEventListener('input',f));f();
let sc=-1,sd=1;document.querySelectorAll('#rt th[data-c]').forEach(h=>h.onclick=()=>{{const i=+h.dataset.c;sd=sc==i?-sd:1;sc=i;document.querySelectorAll('#rt .ar').forEach(a=>a.textContent='');h.querySelector('.ar').textContent=sd>0?' ▲':' ▼';const tb=$('#rt tbody');rs.sort((a,b)=>sd*a.cells[i].textContent.localeCompare(b.cells[i].textContent,undefined,{{numeric:true}})).forEach(r=>tb.appendChild(r))}});
</script></body></html>'''
open(os.path.join(OUT, 'index.html'), 'w').write(page)
os.makedirs(os.path.join(OUT, 'data'), exist_ok=True)
json.dump({'updated': updated, 'comparison': comp, 'non_us': nonus, 'sources': srcs, 'notes': notes}, open(os.path.join(OUT, 'data/ep.json'), 'w'), default=str, indent=1)
print(f'built: {len(comp)} catheters, {ngap} gap cells, {nconf} conflict rows, {len(nonus)} non-US, {len(srcs)} sources')
