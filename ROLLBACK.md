# ROLLBACK — EP Ablation Catheter Comparison guide

Current known-good live state is tagged **`stable`** (also `stable-2026-10-14`). List tags: `git tag -l`.

```bash
cd /workspace/ep-compare
git checkout stable -- index.html data sw.js manifest.webmanifest favicon.svg   # or v1.0 / v1.0.1 / v1.1
git commit -qam "Rollback to stable" && git push -q
```
(`v1.0`/`v1.0.1` predate offline support; rolling back to them also removes it — browsers that installed the worker keep a cached copy until it updates.)

Verify: open https://jeffmcd83.github.io/ep-compare/ (refresh once), check the version/pass label in the header.
