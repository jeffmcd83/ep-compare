#!/bin/bash
# Weekly republish: rebuild from xlsx and push to GitHub Pages
set -e; cd "$(dirname "$0")"; python3 build.py "$@"
git add -A; git diff --cached --quiet || { git commit -qm "Data update $(date +%F)"; git push -q origin main; }
