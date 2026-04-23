#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Batch paper reader using Gemini 2.0 Flash (google-genai) with retries.
- Reads chunked paper texts from INPUT_JSON_FILE (processed_papers.json).
- Prompts Gemini to return a strict JSON object for each paper.
- Merges results into your Excel template by PMCID.
- Saves a finalized Excel with all extracted fields aligned to your schema.
- I added retry logic for 429/500/502/503/504 with capped exponential backoff + jitter
- Client re-creation on mid/backoff attempts
- Auto-resume: skip PMCID if normalized JSON already exists (toggle OVERWRITE)
- Persist .error.txt per PMCID on failure; keep goinng

Prereqs:
  pip install -U google-genai pandas openpyxl

Auth:
  export GEMINI_API_KEY="YOUR_KEY"
  # add your key here only if you are not using an env var
"""

import os
import json
import time
import random
import re
import pathlib
from typing import Any, Dict, List, Optional

import pandas as pd
from google import genai
from google.genai import types

# config 
INPUT_JSON_FILE = "processed_papers.json"                     # Output of your process_papers.py
EXCEL_TEMPLATE_FILE = os.path.expanduser("~/Desktop/placenta_sheet.xlsx")
OUTPUT_EXCEL_FILE = "final_paper_analysis_results_2.xlsx"

MODEL_NAME = "gemini-2.0-flash"                               # Keep your model
REQUEST_INTERVAL_SEC = 1.0                                    # Base pacing between *successful* calls

MAX_RETRIES = 12                                              # More patience under load
BACKOFF_BASE = 2.0                                            # Exponential base
BACKOFF_JITTER = (0.0, 0.6)                                   # Extra random delay
BACKOFF_CAP_SEC = 60.0                                        # Max sleep between retries
RECREATE_CLIENT_EVERY_N_ERRORS = 3                            # Proactively refresh client during flurries
POST_BURST_COOLOFF_SEC = 5.0                                  # After a failure burst, cool down

# set a fallback key here only if you are not using an env var
FALLBACK_API_KEY = ""  # Enter your own key here!

OUTPUT_RAW_DIR = "raw_outputs"                                # folder for per-PMCID artifacts
SAVE_PROMPT_PER_PAPER = True                                  # also store the prompt for auditing
OVERWRITE = True                                             # If False, skip PMCID that already has .normalized.json

# treat these status codes as retryable
RETRYABLE_STATUS_CODES = {429, 500, 502, 503, 504}

# extraction schema

JSON_STRUCTURE: Dict[str, Any] = {
  "Supervisor/Contact/Corresponding author name": "Find the designated corresponding author. If not explicit, use the last author.",
  "Supervisor/Contact/Corresponding author email": "Find the corresponding author's email.",
  "Main topic of the publication": "A brief, 1-2 sentence summary of the paper's main topic.",
  "Pregnancy trimester (1st, 2nd, 3rd, term (for full-term delivery), premature (for early delivery due to complications)": 
  "Return '1st', '2nd', '3rd', 'Term', 'Premature', or a list if multiple are studied. Example: ['1st', '3rd'], Please answer question best according to paper.",
  "Birthweight of offspring provided (yes/no)": "Yes or No.",
  "Gestational Age at delivery provided (yes/no)": "Yes or No.",
  "GA at delivery (weeks)": "Return the value as a number (e.g., 38.5) or 'Not Provided'.Answer question best according to paper.",
  "Gestational Age at sample collection provided (yes/no)": "Yes or No.",
  "GA at sample collection (weeks)": "Return the value as a number (e.g., 34.2) or 'Not Provided'.Answer question best according to paper.",
  "Sex of Offspring Provided (yes/no)": "Yes or No.",
  "Parity provided (yes/no)": "Yes or No.",
  "Gravidity provided (yes/no)": "Yes or No.",
  "Number of offspring per pregnancy provided (yes/no)": "Yes or No.",
  "Self-reported race/ethnicity of mother provided (yes/no)": "Yes or No.",
  "Genetic ancestry or genetic strain provided (yes/no)": "Yes or No.",
  "Maternal Height provided (yes/no)": "Yes or No.",
  "Maternal Pre-pregnancy Weight provided (yes/no)": "Yes or No.",
  "Paternal Height provided (yes/no)": "Yes or No.",
  "Paternal Weight provided (yes/no)": "Yes or No.",
  "Maternal age at sample collection provided (yes/no)": "Yes or No.",
  "Paternal age at sample collection provided (yes/no)": "Yes or No.",
  "Samples from pregnancy complications collected": "Yes or No.",
  "Mode of delivery provided (yes/no)": "Yes or No.",
  "Pregnancy complications in data set (list)": "Return a list of strings.",
  "Fetal complications listed (yes/no)": "Yes or No.",
  "Fetal complications in data set (list)": "Return a list of strings.",
  "Other Phenotypes Provided (list)": "Return a list of other key phenotypes studied.",
  "Hospital/Center where samples were collected": "The name of the institution.",
  "Country where samples were collected": "The name of the country."
}

# utilities

def make_client() -> genai.Client:
    api_key = os.environ.get("GEMINI_API_KEY") or FALLBACK_API_KEY
    if not api_key:
        raise RuntimeError(
            "Missing API key. Set environment variable GEMINI_API_KEY or populate FALLBACK_API_KEY."
        )
    return genai.Client(api_key=api_key)

def ensure_dir(path: str) -> None:
    pathlib.Path(path).mkdir(parents=True, exist_ok=True)

def safe_join(*parts: str) -> str:
    return str(pathlib.Path(*parts))

def write_text(path: str, text: str) -> None:
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)

def write_json(path: str, obj: dict) -> None:
    with open(path, "w", encoding="utf-8") as f:
        json.dump(obj, f, ensure_ascii=False, indent=2)

def extract_status_code(e: Exception) -> Optional[int]:
    """
    Try to robustly recover an HTTP/gRPC-ish status code from various exception flavors.
    """
    # check direct status-code attributes first
    for attr in ("status_code", "code"):
        try:
            val = getattr(e, attr)
            if isinstance(val, int):
                return val
        except Exception:
            pass
    # check a nested response object next
    try:
        resp = getattr(e, "response", None)
        if resp is not None:
            sc = getattr(resp, "status_code", None)
            if isinstance(sc, int):
                return sc
    except Exception:
        pass
    # fall back to parsing a status code from the error text
    m = re.search(r"\b(4\d\d|5\d\d)\b", str(e))
    if m:
        try:
            return int(m.group(1))
        except Exception:
            return None
    return None

def is_retryable_error(e: Exception) -> bool:
    code = extract_status_code(e)
    if code is None:
        # treat common transport errors as retryable
        txt = str(e).lower()
        keywords = ("unavailable", "temporarily", "timeout", "timed out", "connection", "reset by peer", "try again")
        return any(k in txt for k in keywords)
    return code in RETRYABLE_STATUS_CODES

#  prompt 

BASE_INSTRUCTIONS = """
You are an expert biomedical data extractor.

Return ONLY a single valid JSON object (no surrounding text, no code fences).
STRICT OUTPUT CONTRACT:
- Yes/No keys: return exactly "Yes" or "No". If it is not explicitly reported, return "No".
- List keys: return a JSON array of strings ([]) if none.
- Pregnancy trimester: return one of "1st","2nd","3rd","Term","Premature", OR an array of those (e.g., ["1st","3rd"]). If unclear, return "No".
- GA fields (weeks): return a bare number (e.g., 38.5). If a value appears like "38 weeks" return 38. If there are multiple samples, return a list, otherwise use the main value. If unclear, return "No".
- Do not invent values. Prefer "No" if the paper does not explicitly state it.

EXAMPLES (follow exactly):
- "Gestational Age at delivery provided (yes/no)": "Yes"
- "GA at delivery (weeks)": 38.5
- "Pregnancy trimester": "Term"
- "Pregnancy complications in data set (list)": ["Preeclampsia","Gestational diabetes"]
- "Country where samples were collected": "United States"
- Missing/unclear: "No", use your best judgment

Fields to extract (keys) and their instructions (values):
{json_instructions}

---
PAPER TEXT:
{paper_text}
---
"""

def build_prompt(paper_text: str) -> str:
    json_instructions = json.dumps(JSON_STRUCTURE, indent=2)
    return BASE_INSTRUCTIONS.format(json_instructions=json_instructions, paper_text=paper_text)

# post-process helpers

UNKNOWN_STRINGS = {"unknown", "not known", "not reported", "n/a", "na", "none stated"}
GA_NUMERIC_KEYS = {"GA at delivery (weeks)", "GA at sample collection (weeks)"}

def light_fix_unknown(val):
    if val is None:
        return "Not Provided"
    if isinstance(val, str):
        s = val.strip().lower()
        if s in UNKNOWN_STRINGS or s == "not provided":
            return "Not Provided"
    return val

def light_fix_ga_number(val):
    if isinstance(val, (int, float)):
        return float(val)
    if val is None:
        return "Not Provided"
    s = str(val).strip()
    if s.lower() in UNKNOWN_STRINGS or s.lower() == "not provided":
        return "Not Provided"
    m = re.search(r"(-?\d+(?:\.\d+)?)", s)
    if m:
        try:
            return float(m.group(1))
        except Exception:
            pass
    return s

def minimal_postprocess(obj: Dict[str, Any]) -> Dict[str, Any]:
    out = {}
    for key, instr in JSON_STRUCTURE.items():
        val = obj.get(key, None)
        val = light_fix_unknown(val)
        if key in GA_NUMERIC_KEYS:
            val = light_fix_ga_number(val)

        if ("(list)" in key) or ("Return a list" in instr):
            if val is None or val == "Not Provided":
                val = []
            elif isinstance(val, list):
                val = [str(x).strip() for x in val if str(x).strip()]
            elif isinstance(val, str):
                parts = [p.strip() for p in val.replace(";", ",").split(",")]
                val = [p for p in parts if p]
            else:
                val = [str(val).strip()]

        if val is None or (isinstance(val, str) and val.strip() == ""):
            val = "Not Provided"
        out[key] = val
    return out

# model call

def call_model_with_retries(client_factory, prompt: str) -> Optional[str]:
    """
    Calls Gemini with exponential backoff. Returns raw response.text if successful, else None.
    `client_factory` is a zero-arg function so we can recreate clients in-flight.
    """
    client = client_factory()
    consecutive_errors = 0

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = client.models.generate_content(
                model=MODEL_NAME,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json"
                ),
            )
            text = (resp.text or "").strip()
            if not text:
                raise ValueError("Empty response.text from model.")
            return text

        except Exception as e:
            consecutive_errors += 1
            retryable = is_retryable_error(e)
            status_code = extract_status_code(e)
            print(f"    > API error (attempt {attempt}/{MAX_RETRIES})"
                  f"{' [retryable]' if retryable else ''}: {e}")

            if not retryable or attempt == MAX_RETRIES:
                return None

            # use capped exponential backoff with jitter
            backoff = min(BACKOFF_BASE ** (attempt - 1), BACKOFF_CAP_SEC)
            jitter = random.uniform(*BACKOFF_JITTER)
            sleep_s = backoff + jitter
            time.sleep(sleep_s)

            # recreate the client during repeated failures
            if consecutive_errors % RECREATE_CLIENT_EVERY_N_ERRORS == 0:
                try:
                    client = client_factory()
                except Exception as ce:
                    print(f"    > Client re-creation failed: {ce}")

            # add a short cooldown after overload-style errors
            if status_code in {429, 503}:
                time.sleep(POST_BURST_COOLOFF_SEC)

    return None

# pipeline 
def extract_info_from_paper(
    client_factory,
    full_paper_text: str,
    pmcid: str
) -> Optional[Dict[str, Any]]:
    prompt = build_prompt(full_paper_text)
    raw = call_model_with_retries(client_factory, prompt)
    if raw is None:
        return None

    # try a direct json parse first
    parsed = None
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError:
        cleaned = raw.replace("```json", "").replace("```", "").strip()
        try:
            parsed = json.loads(cleaned)
            raw = cleaned
        except Exception:
            print("    > ERROR: Could not decode JSON after cleanup.")
            return None

    if not isinstance(parsed, dict):
        print("    > ERROR: Model returned non-object JSON.")
        return None

    normalized = minimal_postprocess(parsed)
    normalized["PMCID"] = pmcid

    ensure_dir(OUTPUT_RAW_DIR)
    base = safe_join(OUTPUT_RAW_DIR, pmcid)

    write_text(base + ".raw.json", raw)
    write_json(base + ".normalized.json", normalized)
    if SAVE_PROMPT_PER_PAPER:
        write_text(base + ".prompt.txt", prompt)

    return normalized

def merge_with_template(original_df: pd.DataFrame, results_df: pd.DataFrame) -> pd.DataFrame:
    if "PMCID" not in original_df.columns:
        raise ValueError("Template Excel must contain a 'PMCID' column.")

    final_df = pd.merge(original_df, results_df, on="PMCID", how="left", suffixes=("_x", "_y"))

    for col in JSON_STRUCTURE.keys():
        col_x = f"{col}_x"
        col_y = f"{col}_y"
        if col_y in final_df.columns:
            final_df[col] = final_df[col_y].where(final_df[col_y].notna(), final_df.get(col_x))
            for c in (col_x, col_y):
                if c in final_df.columns:
                    final_df.drop(columns=c, inplace=True, errors="ignore")
        else:
            if col not in final_df.columns:
                final_df[col] = pd.NA

    final_columns = list(original_df.columns)
    for col in JSON_STRUCTURE.keys():
        if col not in final_columns:
            final_columns.append(col)

    if "PMCID" in final_columns:
        final_columns = ["PMCID"] + [c for c in final_columns if c != "PMCID"]

    final_df = final_df.reindex(columns=final_columns)
    return final_df

def stringify_lists_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    for col in df.columns:
        if df[col].apply(lambda x: isinstance(x, list)).any():
            df[col] = df[col].apply(lambda x: ", ".join(map(str, x)) if isinstance(x, list) else x)
    return df

def already_done(pmcid: str) -> bool:
    path = safe_join(OUTPUT_RAW_DIR, f"{pmcid}.normalized.json")
    return pathlib.Path(path).exists()

def mark_failure(pmcid: str, err_text: str) -> None:
    ensure_dir(OUTPUT_RAW_DIR)
    write_text(safe_join(OUTPUT_RAW_DIR, f"{pmcid}.error.txt"), err_text)

def main() -> None:
    # load the paper chunks json
    try:
        with open(INPUT_JSON_FILE, "r", encoding="utf-8") as f:
            all_papers_data = json.load(f)
    except FileNotFoundError:
        print(f"Error: The input file '{INPUT_JSON_FILE}' was not found.")
        return
    except json.JSONDecodeError as e:
        print(f"Error: Failed to parse '{INPUT_JSON_FILE}': {e}")
        return

    if not isinstance(all_papers_data, list):
        print(f"Error: '{INPUT_JSON_FILE}' must contain a JSON array.")
        return

    total = len(all_papers_data)
    print(f"Starting analysis of {total} papers.")

    # use a factory so retries can recreate the client
    def client_factory():
        return make_client()

    results: List[Dict[str, Any]] = []
    processed = 0
    skipped = 0
    failures = 0

    for idx, paper in enumerate(all_papers_data, start=1):
        pmcid = paper.get("pmcid", "UNKNOWN")
        chunks = paper.get("chunks", [])
        print(f"\nProcessing paper {idx}/{total} (PMCID: {pmcid})...")

        if not chunks or not isinstance(chunks, list):
            print("    > Skipping: no 'chunks' list found.")
            skipped += 1
            continue

        # skip finished papers unless overwrite is enabled
        if not OVERWRITE and already_done(pmcid):
            print("    > Already extracted. Skipping (set OVERWRITE=True to redo).")
            skipped += 1
            continue

        full_text = " ".join(chunks)

        try:
            extracted = extract_info_from_paper(client_factory, full_text, pmcid)
            if extracted is None:
                failures += 1
                mark_failure(pmcid, "Extraction failed after retries.")
                print(f"    > Failed to extract data for {pmcid}. Skipping.")
            else:
                results.append(extracted)
                processed += 1
                print(f"    > Successfully extracted data for {pmcid}")
        except Exception as e:
            failures += 1
            mark_failure(pmcid, f"Unhandled exception: {e}")
            print(f"    > Unhandled exception for {pmcid}: {e}")

        # only apply base pacing after successful calls
        time.sleep(REQUEST_INTERVAL_SEC)

    print("\nAnalysis complete! Assembling the final Excel sheet...")
    print(f"Summary: processed={processed}, skipped={skipped}, failures={failures}")

    if not results:
        print("Warning: No results extracted; saving empty results file with only headers.")
        empty_cols = ["PMCID"] + list(JSON_STRUCTURE.keys())
        final_df = pd.DataFrame(columns=empty_cols)
        final_df.to_excel(OUTPUT_EXCEL_FILE, index=False)
        print(f"Saved: {OUTPUT_EXCEL_FILE}")
        return

    results_df = pd.DataFrame(results)
    results_df = stringify_lists_for_excel(results_df)

    # merge results into the excel template
    try:
        original_df = pd.read_excel(EXCEL_TEMPLATE_FILE)
        if "PMCID" not in original_df.columns:
            print("Warning: Template missing 'PMCID' column. Saving extracted results only.")
            final_df = results_df
        else:
            final_df = merge_with_template(original_df, results_df)
    except FileNotFoundError:
        print(f"Warning: Original spreadsheet '{EXCEL_TEMPLATE_FILE}' not found. Saving extracted results only.")
        final_df = results_df

    # convert columns to strings for excel compatibility
    final_df.columns = [str(c) for c in final_df.columns]

    final_df.to_excel(OUTPUT_EXCEL_FILE, index=False)
    print(f"\nSuccess! Your spreadsheet is ready: '{OUTPUT_EXCEL_FILE}'")
    print(f"Per-paper artifacts saved to: '{OUTPUT_RAW_DIR}/' directory")

if __name__ == "__main__":
    main()
