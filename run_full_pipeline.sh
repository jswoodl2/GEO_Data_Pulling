#!/usr/bin/env bash
set -euo pipefail

# Run the full GEO pipeline end-to-end (R + Python).
# All scripts live in this folder (flat layout) — no parent-directory hops.
#
# Steps:
#   1) GEO_Extraction.R        — pull GEO metadata for GSE IDs in ids.csv
#   2) fix.R                   — second pass for failed GSEs
#   3) geo_oa_counter.py       — annotate placenta sheet with OA / text-mining info
#   4) download_papers_copy.py — download OA papers (PMC / Europe PMC / Unpaywall / Elsevier)
#   5) download_supplements.py — download PMC / Europe PMC / Elsevier supplementary files
#   6) chunk_xml_papers.py     — extract text from papers into processed_papers.json
#   7) llm_parser.py           — Gemini extraction into Answers / Evidence / Failures sheets
#
# Usage:
#   cd /Users/jjoseph/Desktop/geo_scrap_2026_april28th
#   bash run_full_pipeline.sh

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

# Load secrets (.env is gitignored) — sets GEMINI_API_KEY, ELSEVIER_API_KEY, UNPAYWALL_EMAIL.
if [ -f "$ROOT_DIR/.env" ]; then
  set -a
  # shellcheck disable=SC1091
  source "$ROOT_DIR/.env"
  set +a
  echo "Loaded .env"
else
  echo "Warning: no .env file in $ROOT_DIR — make sure GEMINI_API_KEY etc. are exported"
fi

echo "Root dir: $ROOT_DIR"
echo

# Default R GEO extraction to one worker for stability on macOS.
# Override only if you accept the risk of parallel R worker crashes.
export GEO_R_WORKERS="${GEO_R_WORKERS:-1}"
echo "R GEO workers: $GEO_R_WORKERS"

# Canonical pipeline handoff files. Python scripts read these env vars.
export GEO_METADATA_INPUT="$ROOT_DIR/gse_metadata_full_checkpoint_MERGED.xlsx"
export PIPELINE_EXCEL_FILE="$ROOT_DIR/geo_master_access.xlsx"

# Prefer a local venv; otherwise use the existing geo_scrap venv where these deps are installed.
if [ -x "$ROOT_DIR/.venv/bin/python" ]; then
  PYTHON_BIN="$ROOT_DIR/.venv/bin/python"
elif [ -x "$ROOT_DIR/../geo_scrap/.venv/bin/python" ]; then
  PYTHON_BIN="$ROOT_DIR/../geo_scrap/.venv/bin/python"
else
  PYTHON_BIN="python3"
fi
echo "Python: $PYTHON_BIN"


run_r_step() {
  local script="$1"
  echo "============================================================"
  echo "Running R script: $script"
  echo "============================================================"
  Rscript "$script"
}

run_py_step() {
  local script="$1"
  echo "============================================================"
  echo "Running Python script: $script"
  echo "============================================================"
  "$PYTHON_BIN" "$script"
}

echo "=== 1) R: GEO metadata extraction (first pass) ==="
run_r_step "GEO_Extraction.R"

echo "=== 2) R: Fix failed GEO entries ==="
run_r_step "fix.R"

# If fix.R had nothing to patch, it may not create the merged workbook.
# Fall back to the first-pass checkpoint/full workbook for downstream Python.
if [ ! -f "$GEO_METADATA_INPUT" ]; then
  if [ -f "$ROOT_DIR/gse_metadata_full_checkpoint.xlsx" ]; then
    export GEO_METADATA_INPUT="$ROOT_DIR/gse_metadata_full_checkpoint.xlsx"
  elif [ -f "$ROOT_DIR/gse_metadata_full.xlsx" ]; then
    export GEO_METADATA_INPUT="$ROOT_DIR/gse_metadata_full.xlsx"
  fi
fi

echo "Using GEO metadata workbook: $GEO_METADATA_INPUT"
echo "Pipeline access workbook:    $PIPELINE_EXCEL_FILE"

echo "=== 3) Python: GEO access / OA status ==="
run_py_step "geo_oa_counter.py"

echo "=== 4) Python: Download open-access papers (PMC / Europe PMC / Unpaywall / Elsevier) ==="
run_py_step "download_papers_copy.py"

echo "=== 5) Python: Download supplementary files (PMC / Europe PMC / Elsevier) ==="
# Playwright needs a browser binary on first run; check + install if missing.
if ! "$PYTHON_BIN" -c "from playwright.sync_api import sync_playwright" >/dev/null 2>&1; then
  echo "Playwright not installed. Run: $PYTHON_BIN -m pip install playwright && $PYTHON_BIN -m playwright install chromium"
  exit 1
fi
run_py_step "download_supplements.py"

echo "=== 6) Python: Extract + chunk paper text into JSON ==="
run_py_step "chunk_xml_papers.py"

echo "=== 7) Python: LLM parsing of chunked papers ==="
run_py_step "llm_parser.py"

echo
echo "Pipeline completed successfully."
