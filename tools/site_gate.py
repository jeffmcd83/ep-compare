#!/usr/bin/env python3
"""Deploy gate: serve a site folder locally and smoke-test it in Chrome.
Usage: site_gate.py SITE_DIR [subpath ...]   (default subpath: "")
Exit 0 = safe to publish. Exit 1 = do NOT publish (page error, empty page, bad data, crash).
Checks per page: loads with no JS errors, renders >= 1 table row or card, every data/*.json parses,
manifest.webmanifest is valid JSON (if linked), and the service worker registers."""
import sys, os, json, glob, socket, subprocess, time, traceback
def main():
    site = os.path.abspath(sys.argv[1]); subs = sys.argv[2:] or ['']
    bad = []
    for sub in subs:
        for f in glob.glob(os.path.join(site, sub, 'data', '*.json')):
            try:
                d = json.load(open(f))
                if not d: bad.append(f'empty data {f}')
            except Exception as e: bad.append(f'bad json {f}: {e}')
    s = socket.socket(); s.bind(('127.0.0.1', 0)); port = s.getsockname()[1]; s.close()
    srv = subprocess.Popen([sys.executable, '-m', 'http.server', str(port), '--bind', '127.0.0.1'], cwd=site, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        import urllib.request
        for _ in range(30):
            try: urllib.request.urlopen(f'http://127.0.0.1:{port}/', timeout=2); break
            except Exception: time.sleep(0.5)
        from playwright.sync_api import sync_playwright
        with sync_playwright() as p:
            b = p.chromium.launch(executable_path='/usr/bin/google-chrome')
            for sub in subs:
                url = f'http://127.0.0.1:{port}/{sub}'
                for vp in [{'width': 1300, 'height': 900}, {'width': 390, 'height': 844}]:
                    ctx = b.new_context(viewport=vp); pg = ctx.new_page(); errs = []
                    pg.on('pageerror', lambda e: errs.append(str(e)))
                    r = pg.goto(url); pg.wait_for_timeout(2500)
                    if not r or not r.ok: bad.append(f'{url} HTTP {r.status if r else "none"}')
                    n = pg.evaluate("document.querySelectorAll('tbody tr, .card, table tr').length")
                    if n < 1: bad.append(f'{url} {vp["width"]}px rendered no rows')
                    if errs: bad.append(f'{url} JS errors: {errs[:3]}')
                    if pg.locator('link[rel=manifest]').count():
                        m = pg.evaluate("fetch(document.querySelector('link[rel=manifest]').href).then(r=>r.json()).then(j=>!!j.name).catch(()=>false)")
                        if not m: bad.append(f'{url} manifest invalid')
                    sw = pg.evaluate("navigator.serviceWorker?navigator.serviceWorker.getRegistration().then(r=>!!r):false")
                    if not sw: bad.append(f'{url} service worker not registered')
                    print(f'{url} {vp["width"]}px rows={n} sw={sw} errors={len(errs)}')
                    ctx.close()
            b.close()
    except Exception:
        traceback.print_exc(); bad.append('gate crashed')
    finally:
        srv.kill()
    if bad:
        print('GATE FAILED — do not publish:'); [print(' -', x) for x in bad]; sys.exit(1)
    print('GATE PASSED')
main()
