#!/bin/bash
# Weekly republish: rebuild from xlsx, run the deploy gate, push to GitHub Pages.
# If the gate fails or crashes, local changes are reverted and NOTHING is pushed (live site stays as is).
set -euo pipefail; cd "$(dirname "$0")"; python3 build.py "$@"
if ! python3 tools/site_gate.py .; then
  echo "ABORT: deploy gate failed — reverting build, nothing pushed"; git checkout -- . ; exit 1; fi
git add -A; git diff --cached --quiet || { git commit -qm "Data update $(date +%F)"; git push -q origin main; }
