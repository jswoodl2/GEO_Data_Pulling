# GEO Scrape Pipeline

This branch contains the runnable 2026 GEO extraction pipeline. It starts from
`ids.csv`, pulls GEO metadata, checks article access, downloads papers and
supplements, chunks text, and writes Gemini-extracted answers to
`ai_annotated.xlsx`.

## Run

```bash
cd /Users/jjoseph/Desktop/geo_scrap
cp .env.example .env
# edit .env with real keys
bash run_full_pipeline.sh
```

The pipeline uses the local `.venv` if present. If not, it falls back to
`python3`; install dependencies with:

```bash
python3 -m pip install -r requirements.txt
python3 -m playwright install chromium
```

## Inputs

- `ids.csv`: one GEO Series ID per line. The current file contains the GSE IDs
  selected from the long accession list where `Is_GSE == 1`.
- `.env`: local secrets, not committed.

Required environment values:

- `GEMINI_API_KEY`: Gemini API key for `llm_parser.py`.
- `UNPAYWALL_EMAIL`: email sent to Unpaywall.
- `ELSEVIER_API_KEY`: Elsevier key for article and supplement fallback.

## Steps

`run_full_pipeline.sh` runs these steps in order:

1. `GEO_Extraction.R`: pulls GEO metadata for all IDs in `ids.csv`.
2. `fix.R`: retries and merges failed GEO metadata rows.
3. `geo_oa_counter.py`: checks PMCID/OA/download status.
4. `download_papers_copy.py`: downloads papers from PMC, Europe PMC,
   Unpaywall, and Elsevier where available.
5. `download_supplements.py`: downloads supplementary files from PMC,
   Europe PMC, and Elsevier, with short-file flags in
   `downloaded_supplements/_report.csv`.
6. `chunk_xml_papers.py`: extracts and chunks paper text.
7. `llm_parser.py`: sends GEO metadata, paper text, and supplement text to
   Gemini and writes `ai_annotated.xlsx`.

## Outputs

- `gse_metadata_full*.xlsx`: GEO metadata checkpoints.
- `geo_master_access.xlsx`: metadata plus access/download decisions.
- `downloaded_papers/`: downloaded article files.
- `downloaded_supplements/`: downloaded supplement files and `_report.csv`.
- `processed_papers.json`: chunked paper text.
- `ai_annotated.xlsx`: final workbook with `Answers`, `Evidence`, and
  `Failures` sheets.

Generated outputs, caches, raw model outputs, downloaded papers, supplements,
virtual environments, and `.env` are ignored by git.
