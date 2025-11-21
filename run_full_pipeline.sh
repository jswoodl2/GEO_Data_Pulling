#!/usr/bin/env bash
set -euo pipefail

# End-to-end GEO pipeline (R + Python)
# ------------------------------------
# 1) R_Scripts/GEO_Extraction.R   - pull GEO metadata for GSE IDs in ids.csv
# 2) R_Scripts/fix.R              - optional second pass on failed GSEs
# 3) Python_Scripts/geo_oa_counter.py
#    - annotate placenta_sheet.xlsx with OA / text-mining info
# 4) Python_Scripts/download_papers_copy.py
#    - download open-access papers (PMC / Europe PMC / Unpaywall) to downloaded_papers/
# 5) Python_Scripts/chunk_xml_papers.py
#    - turn XML into processed_papers.json (chunked text)
# 6) Python_Scripts/llm_parser.py
#    - call Gemini to extract placenta metadata -> final_paper_analysis_results_2.xlsx
#
# Usage:
#   cd GEO_Data_Pulling
#   bash run_full_pipeline.sh

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$ROOT_DIR"

run_r() {
  echo "============================================================"
  echo "Running R script: $1"
  echo "============================================================"
  Rscript "$1"
}

run_py() {
  echo "============================================================"
  echo "Running Python script: $*"
  echo "============================================================"
  python3 "$@"
}

echo "=== 1) R: GEO metadata extraction ==="
run_r "R_Scripts/GEO_Extraction.R"

echo "=== 2) R: Fix failed GEO entries (if any) ==="
if [ -f "R_Scripts/fix.R" ]; then
  run_r "R_Scripts/fix.R"
else
  echo "fix.R not found in R_Scripts/; skipping."
fi

echo "=== 3) Python: GEO OA/Access counter ==="
run_py "Python_Scripts/geo_oa_counter.py"

echo "=== 4) Python: Download open-access papers ==="
run_py "Python_Scripts/download_papers_copy.py"

echo "=== 5) Python: Chunk XML papers into JSON ==="
run_py "Python_Scripts/chunk_xml_papers.py"

echo "=== 6) Python: LLM parsing of papers ==="
run_py "Python_Scripts/llm_parser.py"

echo
echo "Pipeline completed."

