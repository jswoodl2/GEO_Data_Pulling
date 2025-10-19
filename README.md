# GEO_Data_Pulling


### GEO_Extraction.R

This script (`GEO_Extraction.R`) reads a list of GEO Series IDs (GSEs), retrieves study‐level and sample‐level metadata from NCBI GEO, maps PubMed → PMCID/DOI using the NIH ID Converter API, and writes results to an Excel file.

### Inputs
- **File:** `ids.csv`  
  - One ID per line; accepts `GSE12345`, `gse12345`, or just `12345`.

### Outputs
- **Excel workbook:** `gse_metadata_full.xlsx` with two sheets:  
  - **Metadata** — one row per GSE in the requested column order  
  - **Failed** — errors for GSEs that could not be fetched  
- **Other files:**  
  - `debug_failures.csv` — errors for failed GSEs  
  - `discarded_ids.csv` — rows skipped during cleaning  
- **Cache:** `geo_cache/` folder (speeds reruns; can be safely deleted to refetch).

### Requirements
- **R version:** ≥ 4.1  
- **R packages:**  
  `GEOquery, readxl, writexl, dplyr, stringr, purrr, lubridate, progressr, tibble, furrr, data.table, rvest, httr, jsonlite`


**First-time installation:**
```r
install.packages(c("readxl","writexl","dplyr","stringr","purrr","lubridate",
                   "progressr","tibble","furrr","data.table","rvest","httr","jsonlite"))
if (!requireNamespace("BiocManager", quietly = TRUE)) install.packages("BiocManager")
BiocManager::install("GEOquery")```

## Usage
From terminal:
```bash
Rscript GEO_Extraction.R
```
Or in Rstudio:  
```bash
 source("GEO_Extraction.R
```

By default, the script processes **all IDs** in `ids.csv`.  
(For testing, you can limit to first 10 IDs by uncommenting `geo_ids = geo_ids[1:10]`).
