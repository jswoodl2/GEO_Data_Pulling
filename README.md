# GEO_Data_Pulling

Pipeline for pulling GEO study metadata (R) and connecting it to full‑text articles and placenta‑specific annotations (Python).

---

## R scripts (`R_Scripts/`)

### `GEO_Extraction.R`

Reads a list of GEO Series IDs (GSEs), retrieves study‑level and sample‑level metadata from NCBI GEO, maps PubMed → PMCID/DOI using the NIH ID Converter API, and writes results to an Excel file.

**Inputs**
- `R_Scripts/ids.csv` — one ID per line (`GSE12345`, `gse12345`, or just `12345`).

**Outputs**
- `R_Scripts/gse_metadata_full.xlsx` with two sheets:  
  - **Metadata** — one row per GSE in the requested column order  
  - **Failed** — errors for GSEs that could not be fetched  
- Other helper files:  
  - `R_Scripts/debug_failures.csv` — errors for failed GSEs  
  - `R_Scripts/discarded_ids.csv` (if present) — rows skipped during cleaning  
- Cache folder: `R_Scripts/geo_cache/` (speeds reruns; safe to delete to refetch).

**R requirements**
- R ≥ 4.1  
- Packages:  
  `GEOquery, readxl, writexl, dplyr, stringr, purrr, lubridate, progressr, tibble, furrr, data.table, rvest, httr, jsonlite`

First‑time install:
```r
install.packages(c("readxl","writexl","dplyr","stringr","purrr","lubridate",
                   "progressr","tibble","furrr","data.table","rvest","httr","jsonlite"))
if (!requireNamespace("BiocManager", quietly = TRUE)) install.packages("BiocManager")
BiocManager::install("GEOquery")
```

### `fix.R`

Optional second pass script. Reads the list of failed GSEs (from the `Failed` sheet / CSV) and retries metadata extraction and cleaning, then merges corrected rows back into `gse_metadata_full.xlsx`.

Run this after `GEO_Extraction.R` if you see non‑trivial numbers of failed IDs.

---

## Python scripts (`Python_Scripts/`)

These scripts are designed to work with the placenta metadata sheet and full‑text papers. They currently assume you have:

- A placenta metadata Excel file (e.g. `placenta_sheet.xlsx`) on disk.
- Network access to NCBI / PMC / Europe‑PMC / Unpaywall.
- A Gemini API key for the LLM parsing step.

All scripts can be run from the repository root (see bash pipeline below).

### `geo_oa_counter.py`

Counts how many GEO entries have PMIDs/DOIs and checks which are open‑access and text‑minable.

- **Inputs:**  
  - Placenta metadata Excel (by default `~/Desktop/placenta_sheet.xlsx` — adjust in the script or pass a path when you customize it).
- **Outputs:**  
  - `geo_master_access.xlsx` — copy of the sheet with columns like `pmcid`, `pmc_oa_subset`, `unpaywall_oa_status`, `license`, and `ok_to_text_mine`.  
  - Prints counts such as “Has PMID”, “In PMC OA subset”, “Accessible for LLM”.

### `download_papers_copy.py`

Downloads open‑access full‑text papers using PMCID / PMID / DOI and only from legitimate OA sources (PMC, Europe PMC, Unpaywall).

- **Inputs:**  
  - Same placenta Excel (for PMIDs/DOIs/PMCIDs).  
- **Outputs:**  
  - `downloaded_papers/` — folder of XML/HTML/PDF (primarily PMC XML).  
  - Summary CSVs listing successes and failures.

### `chunk_xml_papers.py`

Turns downloaded PMC XML files into chunked plain text suitable for LLM input.

- **Inputs:**  
  - `downloaded_papers/` — XML files as created by `download_papers_copy.py`.
- **Outputs:**  
  - `processed_papers.json` — list of objects of the form  
    `{"pmcid": "PMCxxxxxx", "chunks": ["...", "...", ...]}`.

### `llm_parser.py`

Uses Gemini 2.0 Flash to read `processed_papers.json`, extract placenta‑specific metadata for each paper, and merge into your Excel schema.

- **Inputs:**  
  - `processed_papers.json`  
  - Placenta Excel template (default `~/Desktop/placenta_sheet.xlsx`).  
- **Outputs:**  
  - `final_paper_analysis_results_2.xlsx` — AI‑filled placenta metadata sheet.  
  - `raw_outputs/` — per‑paper JSON + prompt artifacts for debugging and auditing.
- **Requires:**  
  - `GEMINI_API_KEY` set in your environment.  
  - Python packages: `google-genai`, `pandas`, `openpyxl`.

### `merge_geo_metadata.py`

Merges GEO metadata (`gse_metadata_full_checkpoint_MERGED.xlsx`) with the AI‑annotated placenta sheet on GEO Series ID, producing a single wide table with both GEO and paper‑level fields.

- **Inputs:** metadata Excel + final paper analysis Excel.  
- **Outputs:** merged Excel with one row per GSE and all placenta annotations appended.

### `ai_vs_students_comparison.py`

Compares student annotations vs AI annotations per GSE and generates both Excel summaries and plots.

- **Inputs:**  
  - `Placenta_Study_Information-5.xlsx` (student sheet).  
  - `gse_metadata_full_checkpoint_MERGED_with_final_cleaned.xlsx` (AI sheet).  
- **Outputs:**  
  - `ai_vs_students_comparison.xlsx` and `ai_vs_students_comparison_metadata.xlsx`.  
  - PNG figures: agreement by field, yes/no confusion matrices, per‑study agreement histogram, trimester distribution plots.

---

## Full pipeline runner

### `run_full_pipeline.sh`

Convenience script to run the full R → Python workflow from the repository root:

```bash
cd GEO_Data_Pulling
bash run_full_pipeline.sh
```

Steps executed:
1. `R_Scripts/GEO_Extraction.R` — pull GEO metadata for IDs in `ids.csv`.  
2. `R_Scripts/fix.R` (if present) — retry failed GSEs and update metadata.  
3. `Python_Scripts/geo_oa_counter.py` — compute OA / text‑mining eligibility.  
4. `Python_Scripts/download_papers_copy.py` — download full‑text papers.  
5. `Python_Scripts/chunk_xml_papers.py` — chunk XML into `processed_papers.json`.  
6. `Python_Scripts/llm_parser.py` — run Gemini on chunks and write `final_paper_analysis_results_2.xlsx`.

Before running the pipeline, update any hard‑coded paths inside the Python scripts (e.g. the location of your placenta Excel file) and ensure the required Python packages are installed.
