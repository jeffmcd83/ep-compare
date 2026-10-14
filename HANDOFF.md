# HANDOFF — EP Ablation Catheter Comparison guide
Live: https://jeffmcd83.github.io/ep-compare/  ·  Written 2026-10-14

## How it's built
- **Source data:** the EP xlsx (default path in `build.py`, EP Bot's store; copy in the backup repo) + `/workspace/ep/` FDA PDFs.
- **Build:** `python3 build.py [xlsx] [guides.json]` writes `index.html`, `data/ep.json`, `sw.js`, `manifest.webmanifest`.
- **Hosting:** GitHub Pages, branch `main`, folder `/`. Static files only.
- **Offline:** `sw.js` (network-first, cache fallback; cache name = build timestamp) + `manifest.webmanifest`.

## Data rules
Verified sources only, naming the exact model/catalog number; every value links to its source; blanks stay blank; conflicts are logged, never overwritten. See STANDING_DECISIONS.md.

## How it's updated (deploy safety)
- `./publish.sh` — rebuild, then `tools/site_gate.py .` (headless Chrome: no JS errors, rows render desktop+mobile, data JSON valid, manifest valid, service worker registers). Gate failure/crash => build reverted, nothing pushed.

## How to roll back
See ROLLBACK.md. Tag `stable` = last known-good live state (2026-10-14).

## Backups
Off-box backup of working files (CSVs, scripts, logs, STANDING_DECISIONS, PDF manifest): private repo `jeffmcd83/device-guides-backup`, refreshed by `bash /workspace/backup_guides.sh`.

## If the bot is gone
The site keeps serving as-is from GitHub Pages. To maintain: clone this repo, restore working files from the backup repo into `/workspace/`, run the build + gate commands above (needs Python 3, openpyxl, playwright + Chrome).
