#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"
[[ "$(uname -s)" == Darwin ]] || { echo 'Run this installer on macOS.' >&2; exit 1; }
command -v python3.13 >/dev/null || { echo 'Install Python 3.13 from python.org, then reopen Terminal.' >&2; exit 1; }
python3.13 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python -m pip check
[[ -f config.json ]] || cp config.example.json config.json
mkdir -p workbooks runtime
for test in test_wheat_corn.py test_parser.py test_monthly_snapshots.py; do
    .venv/bin/python "$test"
done
echo 'PASS: Installation complete. Follow MAC_KURULUM.md to copy the workbooks and test Excel.'
