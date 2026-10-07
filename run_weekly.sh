#!/usr/bin/env bash
# Weekly refresh: rebuild data/ and push. Run from the repo root.
set -euo pipefail
pip install -q --break-system-packages -r pipeline/requirements.txt 2>/dev/null || pip install -q -r pipeline/requirements.txt
python3 pipeline/mf_pipeline.py --out data --workers 16
rm -rf data/_raw
python3 pipeline/validate.py data
git add data
git -c user.name="mf-screener bot" -c user.email="noreply@anthropic.com" commit -m "Weekly data refresh $(date -u +%Y-%m-%d)" || echo "no changes"
git push origin HEAD:main
