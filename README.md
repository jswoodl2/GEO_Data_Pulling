# GEO_Data_Pulling


## Overview on GEO_Extraction.R
This script (`GEO_Extraction.R`) reads a list of GEO Series IDs (GSEs), pulls study‐level and sample‐level metadata from NCBI GEO, enriches it with PubMed→PMCID/DOI using the NIH ID Converter API, and writes a tidy Excel workbook.

- **Input:** `ids.csv` (one ID per line; accepts `GSE12345`, `gse12345`, or just `12345`)
- **Output:** `gse_metadata_full.xlsx` with two sheets:
  - **Metadata** — one row per GSE in the requested column order
  - **Failed** — errors for GSEs that could not be fetched
- **Cache:** `geo_cache/` folder (speeds up reruns; safe to delete to force refetch)

## Requirements
- R ≥ 4.1
- Packages:
  `GEOquery, readxl, writexl, dplyr, stringr, purrr, lubridate, progressr, tibble, furrr, data.table, rvest, httr, jsonlite`

First-time installation:
```r
install.packages(c("readxl","writexl","dplyr","stringr","purrr","lubridate",
                   "progressr","tibble","furrr","data.table","rvest","httr","jsonlite"))
if (!requireNamespace("BiocManager", quietly = TRUE)) install.packages("BiocManager")
BiocManager::install("GEOquery")
```

## Usage
From terminal:
```bash
Rscript GEO_Extraction.R
```
By default, the script processes **all IDs** in `ids.csv`.  
(For testing, you can limit to first 10 IDs by uncommenting `geo_ids = geo_ids[1:10]`).

## Input format
- File: `ids.csv`
- One ID per line, e.g.:
  ```
  GSE12345
  12346
  gse12347
  ```
- The script normalizes IDs (adds “GSE”, uppercases).  
- Rows without digits are discarded.

## Output
- `gse_metadata_full.xlsx` (Excel workbook)
- `debug_failures.csv` (errors for failed GSEs)
- `discarded_ids.csv` (rows skipped during cleaning)

