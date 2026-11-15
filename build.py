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
cath, catrows = rows('Categories') if 'Categories' in wb.sheetnames else ([], [])
gh, gens = rows('Generators') if 'Generators' in wb.sheetnames else ([], [])
mh, maps = rows('Mapping systems') if 'Mapping systems' in wb.sheetnames else ([], [])
shh, shs = rows('Sheaths & transseptal') if 'Sheaths & transseptal' in wb.sheetnames else ([], [])
xh, xs = rows('Diagnostic & ICE') if 'Diagnostic & ICE' in wb.sheetnames else ([], [])
rh, reps = rows('Reprocessed devices') if 'Reprocessed devices' in wb.sheetnames else ([], [])
ph, stacks = rows('Procedure stacks') if 'Procedure stacks' in wb.sheetnames else ([], [])
ah, accs = rows('Accessories') if 'Accessories' in wb.sheetnames else ([], [])
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

def fam_attrs(r):
    fam = str(r.get('Family') or '').strip()
    cats = [x.strip() for x in str(r.get('Category') or '').split(';') if x.strip()]
    return f' data-fam="{e(fam)}" data-cat="{e("|".join(cats))}"'
# Category tree (sheet order if present, else a stable fallback)
if catrows:
    cat_tree = []
    seen=set()
    for r in catrows:
        fam, cat = str(r.get('Family') or '').strip(), str(r.get('Category') or '').strip()
        if fam and cat and (fam, cat) not in seen:
            cat_tree.append((fam, cat)); seen.add((fam, cat))
    families = []
    [families.append(f) for f,_ in cat_tree if f not in families]
else:
    families, cat_tree = [], []
cat_tree_js = {}
for fam, cat in cat_tree:
    cat_tree_js.setdefault(fam, []).append(cat)
CATREE = json.dumps(cat_tree_js)

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
    trs.append(f'<tr data-e="{e(en)}" data-m="{e(mf)}" data-g="{int(any(GAP.search(str(r.get(c) or "")) or not str(r.get(c) or "").strip() for c in cols))}" data-cf="{int(conf)}"{fam_attrs(r)}>' + ''.join(tds) + '</tr>')

GSKIP = {'Source URLs (all for this row)', 'Source IDs'}
gcols = [c for c in gh if c and c not in GSKIP]
ggap = gconf = 0; gtr = []
for r in gens:
    cf = bool(CONF.search(str(r.get('Notes / gaps for this row') or ''))); gconf += cf; tds = []
    for c in gcols:
        s_ = '' if r.get(c) is None else str(r.get(c))
        g = bool(GAP.search(s_)); ggap += g
        if not s_.strip(): body = '' if c == 'Notes / gaps for this row' else '<mark class="gp">[GAP]</mark>'; ggap += c != 'Notes / gaps for this row'
        else: body = linkify(s_)
        if c == 'Generator / console': body = f'<b>{body}</b>'
        tds.append(f'<td{" class=gap" if g else ""}>{body}</td>')
    tds.append(f'<td class="sl">{srclinks(r.get("Source IDs"))}</td>')
    gtr.append(f'<tr{fam_attrs(r)}>' + ''.join(tds) + '</tr>')
gth = ''.join(f'<th>{e(c)}</th>' for c in gcols) + '<th>Sources</th>'
mcols = [c for c in mh if c and c not in GSKIP]
mgap = 0; mtr = []
for r in maps:
    tds = []
    for c in mcols:
        s_ = '' if r.get(c) is None else str(r.get(c))
        g = bool(GAP.search(s_)); mgap += g
        if not s_.strip(): body = '' if c == 'Notes / gaps for this row' else '<mark class="gp">[GAP]</mark>'; mgap += c != 'Notes / gaps for this row'
        else: body = linkify(s_)
        if c == 'System': body = f'<b>{body}</b>'
        tds.append(f'<td{" class=gap" if g else ""}>{body}</td>')
    tds.append(f'<td class="sl">{srclinks(r.get("Source IDs"))}</td>')
    mtr.append(f'<tr{fam_attrs(r)}>' + ''.join(tds) + '</tr>')
mth = ''.join(f'<th>{e(c)}</th>' for c in mcols) + '<th>Sources</th>'
scols = [c for c in shh if c and c not in GSKIP]
sgap = 0; str_ = []
for r in shs:
    tds = []
    for c in scols:
        s_ = '' if r.get(c) is None else str(r.get(c))
        g = bool(GAP.search(s_)); sgap += g
        if not s_.strip(): body = '' if c == 'Notes / gaps for this row' else '<mark class="gp">[GAP]</mark>'; sgap += c != 'Notes / gaps for this row'
        else: body = linkify(s_)
        if c == 'Product': body = f'<b>{body}</b>'
        tds.append(f'<td{" class=gap" if g else ""}>{body}</td>')
    tds.append(f'<td class="sl">{srclinks(r.get("Source IDs"))}</td>')
    str_.append(f'<tr{fam_attrs(r)}>' + ''.join(tds) + '</tr>')
sth = ''.join(f'<th>{e(c)}</th>' for c in scols) + '<th>Sources</th>'
xcols = [c for c in xh if c and c not in GSKIP]
xgap = 0; xtr = []
for r in xs:
    tds = []
    for c in xcols:
        s_ = '' if r.get(c) is None else str(r.get(c))
        g = bool(GAP.search(s_)); xgap += g
        if not s_.strip(): body = '' if c == 'Notes / gaps for this row' else '<mark class="gp">[GAP]</mark>'; xgap += c != 'Notes / gaps for this row'
        else: body = linkify(s_)
        if c == 'Product': body = f'<b>{body}</b>'
        tds.append(f'<td{" class=gap" if g else ""}>{body}</td>')
    tds.append(f'<td class="sl">{srclinks(r.get("Source IDs"))}</td>')
    xtr.append(f'<tr{fam_attrs(r)}>' + ''.join(tds) + '</tr>')
xth = ''.join(f'<th>{e(c)}</th>' for c in xcols) + '<th>Sources</th>'
RSKIP = {'Source URLs (all for this row)', 'Source IDs'}
rcols = [c for c in rh if c and c not in RSKIP]
rgap = 0; rtr = []
reprocs = sorted({str(r.get('Reprocessor') or '').strip() for r in reps} - {''})
rtypes = sorted({str(r.get('Device type') or '').strip() for r in reps} - {''})
for r in reps:
    tds = []
    for c in rcols:
        s_ = '' if r.get(c) is None else str(r.get(c))
        g = bool(GAP.search(s_)); rgap += g
        if not s_.strip():
            body = '' if c in ('Notes / gaps for this row', 'Linked guide device ID') else '<mark class="gp">[GAP]</mark>'
            rgap += c not in ('Notes / gaps for this row', 'Linked guide device ID')
        else:
            body = linkify(s_)
        if c == 'Reprocessed device name (as cleared)': body = f'<b>{body}</b>'
        tds.append(f'<td{" class=gap" if g else ""}>{body}</td>')
    tds.append(f'<td class="sl">{srclinks(r.get("Source IDs"))}</td>')
    rp = e(str(r.get('Reprocessor') or '').strip()); dt = e(str(r.get('Device type') or '').strip())
    rtr.append(f'<tr data-rp="{rp}" data-dt="{dt}"{fam_attrs(r)}>' + ''.join(tds) + '</tr>')
rth = ''.join(f'<th>{e(c)}</th>' for c in rcols) + '<th>Sources</th>'
# Accessories
ASKIP = {'Source URLs (all for this row)', 'Source IDs'}
acols = [c for c in ah if c and c not in ASKIP]
agap = 0; atr = []
for r in accs:
    tds = []
    for c in acols:
        s_ = '' if r.get(c) is None else str(r.get(c))
        g = bool(GAP.search(s_)); agap += g
        if not s_.strip():
            body = '' if c == 'Notes / gaps for this row' else '<mark class="gp">[GAP]</mark>'
            agap += c != 'Notes / gaps for this row'
        else:
            body = linkify(s_)
        if c == 'Product': body = f'<b>{body}</b>'
        tds.append(f'<td{" class=gap" if g else ""}>{body}</td>')
    tds.append(f'<td class="sl">{srclinks(r.get("Source IDs"))}</td>')
    atr.append(f'<tr{fam_attrs(r)}>' + ''.join(tds) + '</tr>')
ath = ''.join(f'<th>{e(c)}</th>' for c in acols) + '<th>Sources</th>'
# Procedure stacks (grouped by Stack ID; filters: procedure, indication, supplier, energy)
PSKIP = {'Source URLs'}
pcols = [c for c in ph if c and c not in PSKIP]
def split_inds(s):
    return [x.strip() for x in str(s or '').split(';') if x.strip()]
procs, inds, sups, ens = [], [], [], []
for r in stacks:
    v = str(r.get('Procedure') or '').strip()
    if v and v not in procs: procs.append(v)
    ind_key = 'Indication (as labeled, e.g. paroxysmal AF, persistent AF, CTI-dependent flutter)'
    for tok in split_inds(r.get(ind_key)):
        if tok not in inds: inds.append(tok)
    v = str(r.get('Supplier') or '').strip()
    if v and v not in sups: sups.append(v)
    v = str(r.get('Energy approach') or '').strip()
    if v and v not in ens: ens.append(v)
STATUS_CLS = {'Labeled pairing': 'lab', 'Typical, not labeled': 'typ', 'Optional': 'opt'}
pth = ''.join(f'<th>{e(c)}</th>' for c in pcols)
ptr = []; pgap = 0; cur = None
ind_key = 'Indication (as labeled, e.g. paroxysmal AF, persistent AF, CTI-dependent flutter)'
model_key = 'Exact model (with guide device ID link, e.g. D05)'
for r in stacks:
    sid = str(r.get('Stack ID') or '').strip()
    if sid and sid != cur:
        cur = sid
        proc = str(r.get('Procedure') or '').strip()
        ind = str(r.get(ind_key) or '').strip()
        sup = str(r.get('Supplier') or '').strip()
        en = str(r.get('Energy approach') or '').strip()
        title = f'{sid} · {sup} · {en} · {proc} · {ind}'
        ptr.append(f'<tr class="stk"><td colspan="{len(pcols)}">{e(title)}</td></tr>')
    tds = []
    status = str(r.get('Labeled status') or '').strip()
    stcls = STATUS_CLS.get(status, '')
    for c in pcols:
        s_ = '' if r.get(c) is None else str(r.get(c))
        g = bool(GAP.search(s_))
        blank = not s_.strip()
        intentional_blank = blank and c in (model_key, 'Family > Category', 'Source IDs', 'Notes', 'Labeled status')
        if blank:
            body = '' if intentional_blank else '<mark class="gp">[GAP]</mark>'
            if not intentional_blank:
                g = True
        else:
            body = linkify(s_)
        if c == model_key and s_.strip():
            body = f'<b>{body}</b>'
        if c == 'Labeled status' and stcls:
            body = f'<span class="stb {stcls}">{e(status)}</span>'
        if g: pgap += 1
        extra = ''
        if g: extra = ' class="gap"'
        elif c == 'Labeled status' and stcls: extra = f' class="{stcls}"'
        tds.append(f'<td{extra}>{body}</td>')
    famcat = str(r.get('Family > Category') or '')
    fam, catv = (famcat.split('>', 1) + [''])[:2] if '>' in famcat else (famcat, '')
    fam, catv = fam.strip(), catv.strip()
    proc = str(r.get('Procedure') or '').strip()
    ind = str(r.get(ind_key) or '').strip()
    sup = str(r.get('Supplier') or '').strip()
    en = str(r.get('Energy approach') or '').strip()
    ind_attr = e('|'.join(split_inds(ind)))
    ptr.append(
        f'<tr data-proc="{e(proc)}" data-ind="{ind_attr}" data-sup="{e(sup)}" data-en="{e(en)}" data-fam="{e(fam)}" data-cat="{e(catv)}" data-stack="{e(sid)}">'
        + ''.join(tds) + '</tr>'
    )
nu = ''.join(f'<tr{fam_attrs(x)}>' + ''.join(f'<td>{linkify(x.get(c))}</td>' for c in nh) + '</tr>' for x in nonus)
nuh = ''.join(f'<th>{e(c)}</th>' for c in nh)
cats = {}
for n in notes: cats.setdefault(n['Category'], []).append(n['Note'])
nt = ''.join(f'<div class="nc"><h3>{e(k)}</h3><ul>' + ''.join(f'<li class="{"cfn" if k=="Discrepancy" else ""}">{linkify(v)}</li>' for v in vs) + '</ul></div>' for k, vs in cats.items())
sr = ''.join(f'<tr><td><b>{e(s["Source ID"])}</b></td><td>{e(s["Title / description"])}</td><td><a href="{e(s["URL"])}" target="_blank" rel="noopener">link</a></td><td>{e(s.get("Source type"))}</td><td>{e(s.get("Used for device rows"))}</td><td>{e(s.get("Accessed"))}</td></tr>' for s in srcs)
updated = datetime.date.today().isoformat()
opt = lambda xs: ''.join(f'<option>{e(x)}</option>' for x in xs)

crm = open(os.path.join(OUT, 'base.css')).read()
page = f'''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{TITLE}</title><link rel="icon" href="favicon.svg" type="image/svg+xml"><link rel="manifest" href="manifest.webmanifest"><meta name="theme-color" content="#111827"><meta name="color-scheme" content="light">
<style>{crm}#pst tr.stk td{{background:#1F3864;color:#fff;font-weight:700;letter-spacing:.01em;padding:.65rem .8rem}}#pst td.lab,#act td.lab{{background:#e8f5e9}}#pst td.typ{{background:#fff8e1}}#pst td.opt{{background:#e3f2fd}}.stb{{display:inline-block;padding:.1rem .45rem;border-radius:999px;font-size:.85em;font-weight:600}}.stb.lab{{background:#c8e6c9;color:#1b5e20}}.stb.typ{{background:#ffe082;color:#5d4037}}.stb.opt{{background:#bbdefb;color:#0d47a1}}</style></head><body>
<header><h1>{TITLE}</h1>{guides_menu()}<span class="pass">{len(comp)} US-approved catheters · updated {updated}</span>
<nav class="tabs"><button data-t="cmp" class="on">Comparison</button><button data-t="gen">Generators</button><button data-t="map">Mapping systems</button><button data-t="sh">Sheaths &amp; transseptal</button><button data-t="dx">Diagnostic &amp; ICE</button><button data-t="rep">Reprocessed</button><button data-t="ps">Procedure stacks</button><button data-t="acc">Accessories</button><button data-t="nus">Non-US &amp; pipeline</button><button data-t="nts">Notes &amp; gaps</button><button data-t="src">Sources</button></nav>
<div class="flt catf" id="catflt"><select id="ffam" aria-label="Family"><option value="">All families</option>{opt(families)}</select><select id="fcat" aria-label="Category"><option value="">All categories</option></select><span class="cnt" id="fcn"></span></div></header>
<main>
<section id="cmp" class="pane on">
<div class="flt"><input id="q" placeholder="Filter (product, spec, indication…)" aria-label="Filter"><select id="fe"><option value="">All energy types</option>{opt(energies)}</select><select id="fm"><option value="">All manufacturers</option>{opt(mfrs)}</select>
<label class="tg"><input type="checkbox" id="fg"> Rows with gaps</label><label class="tg"><input type="checkbox" id="fc"> Rows with conflicts</label><span class="cnt" id="n"></span></div>
<p class="lg"><mark class="gp">[GAP]</mark> not found in an authoritative source (no value inferred) · <span class="cfb">⚠ conflict</span> sources disagree; both values shown · click a header to sort</p>
<div class="wrap rwrap"><table id="rt"><thead><tr>{th}</tr></thead><tbody>{''.join(trs)}</tbody></table></div></section>
<section id="gen" class="pane"><div class="nc"><p>US FDA-approved ablation generators and consoles. <mark class="gp">[GAP]</mark> = not found in an authoritative source.</p></div><div class="wrap rwrap"><table id="gt"><thead><tr>{gth}</tr></thead><tbody>{"".join(gtr)}</tbody></table></div></section>
<section id="map" class="pane"><div class="nc"><p>US FDA-cleared 3D electroanatomic mapping / navigation systems. <mark class="gp">[GAP]</mark> = not found in an authoritative source.</p></div><div class="wrap rwrap"><table id="mt"><thead><tr>{mth}</tr></thead><tbody>{"".join(mtr)}</tbody></table></div></section>
<section id="sh" class="pane"><div class="nc"><p>US FDA-cleared EP access tools: steerable and fixed-curve sheaths, transseptal needles, dilators and RF wires. <mark class="gp">[GAP]</mark> = not found in an authoritative source.</p></div><div class="wrap rwrap"><table id="st"><thead><tr>{sth}</tr></thead><tbody>{"".join(str_)}</tbody></table></div></section>
<section id="dx" class="pane"><div class="nc"><p>US FDA-cleared EP diagnostic/mapping catheters and intracardiac echo (ICE) catheters. <mark class="gp">[GAP]</mark> = not found in an authoritative source.</p></div><div class="wrap rwrap"><table id="xt"><thead><tr>{xth}</tr></thead><tbody>{"".join(xtr)}</tbody></table></div></section>
<section id="rep" class="pane">
<div class="flt"><input id="rq" placeholder="Filter reprocessed devices…" aria-label="Filter reprocessed"><select id="fr"><option value="">All reprocessors</option>{opt(reprocs)}</select><select id="fd"><option value="">All device types</option>{opt(rtypes)}</select><label class="tg"><input type="checkbox" id="rfg"> Rows with gaps</label><span class="cnt" id="rn"></span></div>
<p class="lg">US FDA 510(k) clearances for reprocessed EP devices. Catalog numbers are recorded only when a 510(k) summary (or a reprocessor catalog page) names them. <mark class="gp">[GAP]</mark> = not found in an authoritative source. Linked guide device IDs point to an exact match on another tab.</p>
<div class="wrap rwrap"><table id="rpt"><thead><tr>{rth}</tr></thead><tbody>{"".join(rtr)}</tbody></table></div></section>
<section id="ps" class="pane">
<div class="flt"><select id="sp" aria-label="Procedure"><option value="">All procedures</option>{opt(procs)}</select><select id="si" aria-label="Indication"><option value="">All indications</option>{opt(inds)}</select><select id="ss" aria-label="Supplier"><option value="">All suppliers</option>{opt(sups)}</select><select id="se" aria-label="Energy approach"><option value="">All energy approaches</option>{opt(ens)}</select><span class="cnt" id="pn"></span></div>
<p class="lg">First-batch PVI / AF stacks, one row per component slot. <span class="stb lab">Labeled pairing</span> is named in an SSED, IFU, 510(k), or already-cited OEM spec. <span class="stb typ">Typical, not labeled</span> (ICE and transseptal tools in this batch) means the slot is usual in practice but no model is named. <span class="stb opt">Optional</span> is a labeled accessory that the source calls optional. Blank models stay blank when no source names one — not inferred.</p>
<div class="wrap rwrap"><table id="pst"><thead><tr>{pth}</tr></thead><tbody>{"".join(ptr)}</tbody></table></div></section>
<section id="acc" class="pane"><div class="nc"><p>Named accessories pulled from the same SSED / IFU / 510(k) citations used in the procedure stacks (cables, guidewires, pumps and tubing, patches, foot switches, sensors). <mark class="gp">[GAP]</mark> = a catalog or brand was not named.</p></div><div class="wrap rwrap"><table id="act"><thead><tr>{ath}</tr></thead><tbody>{"".join(atr)}</tbody></table></div></section>
<section id="nus" class="pane"><div class="nc"><p>Not FDA approved as of the data date; listed for awareness only.</p></div><div class="wrap"><table id="nt"><thead><tr>{nuh}</tr></thead><tbody>{nu}</tbody></table></div></section>
<section id="nts" class="pane">{nt}</section>
<section id="src" class="pane"><div class="wrap rwrap"><table id="st"><thead><tr><th>ID</th><th>Title</th><th>URL</th><th>Type</th><th>Rows</th><th>Accessed</th></tr></thead><tbody>{sr}</tbody></table></div></section>
</main><footer>Public sources only (FDA SSED/PMA/recall databases, manufacturer spec sheets). Not medical advice; verify against current IFU. Updated {updated}.</footer>
<script>
const $=s=>document.querySelector(s),rs=[...document.querySelectorAll('#rt tbody tr')];
const CATREE={CATREE};
const rrs=[...document.querySelectorAll('#rpt tbody tr')];
function catMatch(tr){{
  const fam=$('#ffam')?$('#ffam').value:'', cat=$('#fcat')?$('#fcat').value:'';
  if(fam && (tr.dataset.fam||'')!==fam) return false;
  if(cat){{const cs=(tr.dataset.cat||'').split('|').filter(Boolean); if(!cs.includes(cat)) return false;}}
  return true;
}}
function fillCats(){{
  const fam=$('#ffam').value, sel=$('#fcat'), cur=sel.value;
  const cats=fam?(CATREE[fam]||[]):Object.keys(CATREE).reduce((a,k)=>a.concat(CATREE[k]||[]),[]);
  sel.innerHTML='<option value="">All categories</option>'+cats.map(c=>'<option>'+c+'</option>').join('');
  if(cats.includes(cur)) sel.value=cur;
}}
function f(){{const q=$('#q').value.toLowerCase(),en=$('#fe').value,m=$('#fm').value,g=$('#fg').checked,c=$('#fc').checked;let n=0;rs.forEach(r=>{{const v=(!q||r.textContent.toLowerCase().includes(q))&&(!en||r.dataset.e==en)&&(!m||r.dataset.m==m)&&(!g||r.dataset.g=='1')&&(!c||r.dataset.cf=='1')&&catMatch(r);r.hidden=!v;n+=v}});$('#n').textContent=n+' of '+rs.length+' shown'}}
function rf(){{const q=($('#rq')&&$('#rq').value||'').toLowerCase(),rp=$('#fr')?$('#fr').value:'',dt=$('#fd')?$('#fd').value:'',g=$('#rfg')&&$('#rfg').checked;let n=0;rrs.forEach(r=>{{const hasGap=!!r.querySelector('.gap,.gp');const v=(!q||r.textContent.toLowerCase().includes(q))&&(!rp||r.dataset.rp==rp)&&(!dt||r.dataset.dt==dt)&&(!g||hasGap)&&catMatch(r);r.hidden=!v;n+=v}});if($('#rn'))$('#rn').textContent=n+' of '+rrs.length+' shown'}}
function filterOther(){{
  document.querySelectorAll('#gen tbody tr, #map tbody tr, #sh tbody tr, #dx tbody tr, #acc tbody tr, #nus tbody tr').forEach(r=>{{r.hidden=!catMatch(r)}});
}}
function recount(){{
  const pane=document.querySelector('.pane.on'); if(!pane||!$('#fcn')) return;
  if(pane.id==='nts'||pane.id==='src'){{$('#fcn').textContent=''; $('#catflt').hidden=true; return;}}
  $('#catflt').hidden=false;
  const rows=[...pane.querySelectorAll('tbody tr')];
  const vis=rows.filter(r=>!r.hidden).length;
  $('#fcn').textContent=vis+' of '+rows.length+' shown';
}}
function sf(){{
  if(!document.querySelector('#pst')) return;
  const p=$('#sp')?$('#sp').value:'', i=$('#si')?$('#si').value:'', s=$('#ss')?$('#ss').value:'', en=$('#se')?$('#se').value:'';
  let rows=0,vis=0,stacks=0,vstk=0;
  document.querySelectorAll('#pst tbody tr.stk').forEach(h=>{{
    stacks++;
    let show=false;
    let n=h.nextElementSibling;
    while(n && !n.classList.contains('stk')){{
      rows++;
      const toks=(n.dataset.ind||'').split('|').filter(Boolean);
      const ok=(!p||n.dataset.proc===p)&&(!i||toks.includes(i))&&(!s||n.dataset.sup===s)&&(!en||n.dataset.en===en)&&catMatch(n);
      n.hidden=!ok; if(ok){{show=true;vis++}} n=n.nextElementSibling;
    }}
    h.hidden=!show; if(show)vstk++;
  }});
  if($('#pn')) $('#pn').textContent=vstk+' of '+stacks+' stacks · '+vis+' of '+rows+' rows';
}}
function applyAll(){{f();rf();sf();filterOther();recount();}}
document.querySelectorAll('nav.tabs button').forEach(b=>b.onclick=()=>{{document.querySelectorAll('nav.tabs button,.pane').forEach(x=>x.classList.remove('on'));b.classList.add('on');$('#'+b.dataset.t).classList.add('on');applyAll()}});
['#q','#fe','#fm','#fg','#fc'].forEach(s=>$(s).addEventListener('input',applyAll));
['#rq','#fr','#fd','#rfg'].forEach(s=>{{const el=$(s);if(el)el.addEventListener('input',applyAll)}});
['#sp','#si','#ss','#se'].forEach(s=>{{const el=$(s);if(el)el.addEventListener('input',applyAll)}});
$('#ffam').addEventListener('input',()=>{{fillCats();applyAll()}});
$('#fcat').addEventListener('input',applyAll);
fillCats();applyAll();
let sc=-1,sd=1;document.querySelectorAll('#rt th[data-c]').forEach(h=>h.onclick=()=>{{const i=+h.dataset.c;sd=sc==i?-sd:1;sc=i;document.querySelectorAll('#rt .ar').forEach(a=>a.textContent='');h.querySelector('.ar').textContent=sd>0?' ▲':' ▼';const tb=$('#rt tbody');rs.sort((a,b)=>sd*a.cells[i].textContent.localeCompare(b.cells[i].textContent,undefined,{{numeric:true}})).forEach(r=>tb.appendChild(r))}});
if('serviceWorker'in navigator)navigator.serviceWorker.register('sw.js');
</script></body></html>'''
open(os.path.join(OUT, 'index.html'), 'w').write(page)
# Offline support: cache name changes every build so browsers pick up new data
open(os.path.join(OUT, 'sw.js'), 'w').write("const P='ep-',C=P+'%s';self.addEventListener('install',e=>{self.skipWaiting();e.waitUntil(caches.open(C).then(c=>c.addAll(['./','./favicon.svg','./data/ep.json','./manifest.webmanifest'])))});self.addEventListener('activate',e=>e.waitUntil(caches.keys().then(k=>Promise.all(k.filter(x=>x.startsWith(P)&&x!==C).map(x=>caches.delete(x)))).then(()=>self.clients.claim())));self.addEventListener('fetch',e=>{if(e.request.method!=='GET')return;const u=new URL(e.request.url);if(u.origin!==location.origin)return;e.respondWith(caches.open(C).then(c=>fetch(e.request,{cache:'no-cache'}).then(r=>{if(r.ok)c.put(e.request,r.clone());return r}).catch(()=>c.match(e.request,{ignoreSearch:true}))))});" % datetime.datetime.now().strftime('%Y%m%d%H%M%S'))
json.dump({'name': TITLE, 'short_name': 'EP Guide', 'start_url': './', 'scope': './', 'display': 'standalone', 'background_color': '#111827', 'theme_color': '#111827', 'icons': [{'src': 'favicon.svg', 'sizes': 'any', 'type': 'image/svg+xml'}]}, open(os.path.join(OUT, 'manifest.webmanifest'), 'w'), indent=1)
os.makedirs(os.path.join(OUT, 'data'), exist_ok=True)
json.dump({'updated': updated, 'comparison': comp, 'generators': gens, 'mapping_systems': maps, 'sheaths': shs, 'diagnostic_ice': xs, 'reprocessed': reps, 'procedure_stacks': stacks, 'accessories': accs, 'non_us': nonus, 'sources': srcs, 'notes': notes, 'categories': catrows}, open(os.path.join(OUT, 'data/ep.json'), 'w'), default=str, indent=1)
print(f'built: {len(comp)} catheters, {ngap} gap cells, {nconf} conflict rows, {len(gens)} generators ({ggap} gap cells), {len(maps)} mapping systems ({mgap} gap cells), {len(shs)} sheaths ({sgap} gap cells), {len(xs)} diagnostic/ICE ({xgap} gap cells), {len(reps)} reprocessed ({rgap} gap cells), {len(stacks)} stack rows ({pgap} gap cells), {len(accs)} accessories ({agap} gap cells), {len(nonus)} non-US, {len(srcs)} sources')
