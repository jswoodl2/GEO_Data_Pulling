"""
Subset Re-run: AI annotation with Confidence Scores + Evidence Quotes
----------------------------------------------------------------------
Selects N papers with the most AI-vs-student disagreements (from comparison),
re-queries Gemini 2.0 Flash asking for:
  - answer: "Yes" or "No"
  - confidence: 0.0 – 1.0
  - evidence: direct quote from the paper (or "Not found")

Outputs:
  subset_rerun_with_confidence.xlsx   — per-paper × per-question results
  subset_rerun_summary.xlsx           — pivot summary

Set GEMINI_API_KEY env var before running:
  export GEMINI_API_KEY="your_key_here"
  python3 rerun_subset_with_confidence.py
"""

import os, json, time, random, re, pathlib
import pandas as pd
import warnings
warnings.filterwarnings("ignore")

try:
    from google import genai
    from google.genai import types
    GENAI_AVAILABLE = True
except ImportError:
    GENAI_AVAILABLE = False
    print("WARNING: google-genai not installed. Run: pip install -U google-genai")

COMPARISON_XLSX   = "comparison_ai_vs_student.xlsx"
AI_XLSX           = "ai_annotated.xlsx"
PAPERS_JSON       = "processed_papers.json"
OUTPUT_XLSX       = "subset_rerun_with_confidence.xlsx"
SUMMARY_XLSX      = "subset_rerun_summary.xlsx"

MODEL_NAME        = "gemini-2.0-flash"
N_PAPERS          = 25       # how many papers to re-run (top N by disagreement count)
REQUEST_INTERVAL  = 1.5      # seconds between calls
MAX_RETRIES       = 6
BACKOFF_BASE      = 2.0
BACKOFF_CAP       = 45.0
FALLBACK_API_KEY  = ""       # set here OR via GEMINI_API_KEY env var

YN_QUESTIONS = [
    "Birthweight of offspring provided (yes/no)",
    "Gestational Age at delivery provided (yes/no)",
    "Gestational Age at sample collection provided (yes/no)",
    "Sex of Offspring Provided (yes/no)",
    "Parity provided (yes/no)",
    "Gravidity provided (yes/no)",
    "Number of offspring per pregnancy provided (yes/no)",
    "Self-reported race/ethnicity of mother provided (yes/no)",
    "Genetic ancestry or genetic strain provided (yes/no)",
    "Maternal Height provided (yes/no)",
    "Maternal Pre-pregnancy Weight provided (yes/no)",
    "Paternal Height provided (yes/no)",
    "Paternal Weight provided (yes/no)",
    "Maternal age at sample collection provided (yes/no)",
    "Paternal age at sample collection provided (yes/no)",
    "Samples from pregnancy complications collected",
    "Mode of delivery provided (yes/no)",
    "Fetal complications listed (yes/no)",
]

SHORT_LABELS = {
    "Birthweight of offspring provided (yes/no)":                   "Birthweight",
    "Gestational Age at delivery provided (yes/no)":                "GA at delivery",
    "Gestational Age at sample collection provided (yes/no)":       "GA at collection",
    "Sex of Offspring Provided (yes/no)":                           "Sex of offspring",
    "Parity provided (yes/no)":                                     "Parity",
    "Gravidity provided (yes/no)":                                  "Gravidity",
    "Number of offspring per pregnancy provided (yes/no)":          "N offspring/preg",
    "Self-reported race/ethnicity of mother provided (yes/no)":     "Race/ethnicity",
    "Genetic ancestry or genetic strain provided (yes/no)":         "Genetic ancestry",
    "Maternal Height provided (yes/no)":                            "Maternal height",
    "Maternal Pre-pregnancy Weight provided (yes/no)":              "Maternal pre-preg wt",
    "Paternal Height provided (yes/no)":                            "Paternal height",
    "Paternal Weight provided (yes/no)":                            "Paternal weight",
    "Maternal age at sample collection provided (yes/no)":          "Maternal age",
    "Paternal age at sample collection provided (yes/no)":          "Paternal age",
    "Samples from pregnancy complications collected":                "Complication samples",
    "Mode of delivery provided (yes/no)":                           "Mode of delivery",
    "Fetal complications listed (yes/no)":                          "Fetal complications",
}

SYSTEM_PROMPT = """You are an expert biomedical data extractor specializing in placental research publications.

For each question below, return a JSON object with EXACTLY these fields:
  "answer":     "Yes" or "No"
  "confidence": a float from 0.0 (very uncertain) to 1.0 (completely certain)
  "evidence":   a SHORT verbatim quote (≤ 50 words) from the paper that directly
                supports your answer, or "Not found in text" if you cannot locate one.

Rules:
- Answer "Yes" only if the data item is EXPLICITLY reported/provided in the paper.
- If not clearly stated, answer "No" with confidence ≥ 0.5.
- confidence = 1.0 means you found a clear, unambiguous statement.
- confidence < 0.5 means you are guessing; explain briefly in evidence field.
- Do NOT invent quotes. Copy text verbatim.

Return ONLY a single valid JSON object (no markdown, no code fences) with this structure:
{
  "Q_Birthweight":         {"answer": "Yes", "confidence": 0.95, "evidence": "Birth weight was recorded for all neonates (mean 3.2 kg)."},
  "Q_GA_delivery":         {"answer": "No",  "confidence": 0.9,  "evidence": "Not found in text"},
  ...
}

Questions (use these exact keys):
{question_block}

---
PAPER TEXT:
{paper_text}
---
"""

def build_question_block():
    lines = []
    for q in YN_QUESTIONS:
        key = "Q_" + SHORT_LABELS[q].replace(" ", "_").replace("/", "_").replace("-", "_")
        lines.append(f'  "{key}": "{q}"')
    return "{\n" + ",\n".join(lines) + "\n}"

def get_question_key_map():
    """Returns {short_key: full_question_name}"""
    m = {}
    for q in YN_QUESTIONS:
        key = "Q_" + SHORT_LABELS[q].replace(" ", "_").replace("/", "_").replace("-", "_")
        m[key] = q
    return m

def make_client():
    api_key = os.environ.get("GEMINI_API_KEY") or FALLBACK_API_KEY
    if not api_key:
        raise RuntimeError("Set GEMINI_API_KEY env var or FALLBACK_API_KEY in script.")
    return genai.Client(api_key=api_key)

def extract_status_code(e):
    for attr in ("status_code", "code"):
        try:
            v = getattr(e, attr)
            if isinstance(v, int):
                return v
        except Exception:
            pass
    try:
        r = getattr(e, "response", None)
        if r:
            sc = getattr(r, "status_code", None)
            if isinstance(sc, int):
                return sc
    except Exception:
        pass
    m = re.search(r"\b(4\d\d|5\d\d)\b", str(e))
    return int(m.group(1)) if m else None

def call_model(client_factory, prompt):
    client = client_factory()
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = client.models.generate_content(
                model=MODEL_NAME,
                contents=prompt,
                config=types.GenerateContentConfig(response_mime_type="application/json"),
            )
            text = (resp.text or "").strip()
            if not text:
                raise ValueError("Empty response")
            return text
        except Exception as e:
            code = extract_status_code(e)
            retryable = code in {429, 500, 502, 503, 504} if code else True
            print(f"      API error attempt {attempt}/{MAX_RETRIES}: {e}")
            if not retryable or attempt == MAX_RETRIES:
                return None
            sleep = min(BACKOFF_BASE ** (attempt - 1) + random.uniform(0, 0.5), BACKOFF_CAP)
            time.sleep(sleep)
            if code in {429, 503}:
                time.sleep(5)
            client = client_factory()
    return None

def main():
    print("=" * 65)
    print("Subset AI Re-run with Confidence Scores + Evidence Quotes")
    print("=" * 65)

    # 1. load comparison to find top disagreement papers
    print("\n[1/5] Loading comparison sheet...")
    comp = pd.read_excel(COMPARISON_XLSX, sheet_name="Comparison")
    agree_cols = [c for c in comp.columns if "__agreement" in c]
    comp["n_disagree"] = (comp[agree_cols] == "disagree").sum(axis=1)
    top = comp.nlargest(N_PAPERS * 3, "n_disagree")  # over-sample; we'll filter by PMCID

    # 2. load ai sheet to get pmcids
    print("[2/5] Loading AI annotated sheet for PMCIDs...")
    ai = pd.read_excel(AI_XLSX)
    ai["_geo"] = ai["GEO Series ID (GSE___)"].astype(str).str.strip().str.upper()
    geo_to_pmcid = ai.set_index("_geo")["PMCID"].dropna().to_dict()

    # 3. load processed papers json
    print("[3/5] Loading processed papers JSON...")
    with open(PAPERS_JSON, "r", encoding="utf-8") as f:
        all_papers = json.load(f)
    pmcid_to_text = {
        p["pmcid"]: " ".join(p.get("chunks", []))
        for p in all_papers
        if p.get("pmcid") and p.get("chunks")
    }
    print(f"      {len(pmcid_to_text)} papers with text available.")

    # 4. select target papers
    selected = []
    for _, row in top.iterrows():
        geo = str(row["GEO_ID"]).strip().upper()
        pmcid = geo_to_pmcid.get(geo)
        if pmcid and str(pmcid) in pmcid_to_text:
            selected.append({
                "geo_id":      geo,
                "pmcid":       str(pmcid),
                "title":       str(row.get("PlTitle", "")),
                "n_disagree":  int(row["n_disagree"]),
            })
            if len(selected) >= N_PAPERS:
                break

    print(f"      Selected {len(selected)} papers (target: {N_PAPERS}).")
    if not selected:
        print("ERROR: No papers found with both a PMCID and text. Check your files.")
        return

    # 5. run model
    if not GENAI_AVAILABLE:
        print("\nERROR: google-genai not available. Install it and re-run.")
        return

    api_key = os.environ.get("GEMINI_API_KEY") or FALLBACK_API_KEY
    if not api_key:
        print("\nERROR: No API key found. Set GEMINI_API_KEY env var.")
        return

    client_factory = make_client
    question_block = build_question_block()
    key_map = get_question_key_map()

    print(f"\n[4/5] Running Gemini {MODEL_NAME} on {len(selected)} papers...\n")

    all_rows = []
    for i, paper in enumerate(selected, 1):
        pmcid = paper["pmcid"]
        geo   = paper["geo_id"]
        title = paper["title"][:80]
        print(f"  [{i:2d}/{len(selected)}] {geo} | {pmcid} | {title}...")

        paper_text = pmcid_to_text[pmcid]
        prompt = SYSTEM_PROMPT.format(
            question_block=question_block,
            paper_text=paper_text[:40000],  # cap at 40k chars to stay within limits
        )

        raw = call_model(client_factory, prompt)
        if raw is None:
            print(f"        FAILED – skipping.")
            continue

        # parse json
        try:
            parsed = json.loads(raw)
        except json.JSONDecodeError:
            cleaned = raw.replace("```json", "").replace("```", "").strip()
            try:
                parsed = json.loads(cleaned)
            except Exception:
                print(f"        JSON parse error – skipping.")
                continue

        if not isinstance(parsed, dict):
            print(f"        Non-dict response – skipping.")
            continue

        # build the output row
        row = {
            "GEO_ID":       geo,
            "PMCID":        pmcid,
            "Paper title":  paper["title"],
            "N_disagree (AI vs Student)": paper["n_disagree"],
        }

        # pull the original ai and student answers
        comp_row = comp[comp["GEO_ID"] == geo]

        for short_key, full_q in key_map.items():
            label = SHORT_LABELS[full_q]
            entry = parsed.get(short_key, {})
            if not isinstance(entry, dict):
                entry = {}

            answer     = str(entry.get("answer",     "")).strip()
            confidence = entry.get("confidence", None)
            evidence   = str(entry.get("evidence",  "Not found")).strip()

            # keep the original answers for comparison
            if not comp_row.empty:
                ai_orig  = comp_row.iloc[0].get(full_q + "__norm__ai",      "")
                stu_orig = comp_row.iloc[0].get(full_q + "__norm__student",  "")
                agree    = comp_row.iloc[0].get(label + "__agreement",       "")
            else:
                ai_orig = stu_orig = agree = ""

            row[f"{label} | New AI answer"]    = answer
            row[f"{label} | Confidence"]       = confidence
            row[f"{label} | Evidence quote"]   = evidence
            row[f"{label} | Orig AI answer"]   = ai_orig
            row[f"{label} | Student answer"]   = stu_orig
            row[f"{label} | Orig agreement"]   = agree

            # flag low-confidence answers
            try:
                low_conf = float(confidence) < 0.6
            except (TypeError, ValueError):
                low_conf = True
            row[f"{label} | Low confidence?"] = "YES" if low_conf else ""

        all_rows.append(row)
        print(f"        OK. Answers for {len(key_map)} questions.")
        time.sleep(REQUEST_INTERVAL)

    if not all_rows:
        print("No results collected. Exiting.")
        return

    # 6. write excel
    print(f"\n[5/5] Writing outputs...")
    df = pd.DataFrame(all_rows)

    # summary: per question, flag low-confidence rows
    summary_rows = []
    for full_q in YN_QUESTIONS:
        label = SHORT_LABELS[full_q]
        conf_col = f"{label} | Confidence"
        ans_col  = f"{label} | New AI answer"
        low_col  = f"{label} | Low confidence?"
        agree_col = f"{label} | Orig agreement"

        if conf_col not in df.columns:
            continue

        confs = pd.to_numeric(df[conf_col], errors="coerce").dropna()
        answers = df[ans_col].str.lower()

        summary_rows.append({
            "Question":                full_q,
            "Short label":             label,
            "N papers":                len(df),
            "Avg confidence":          round(confs.mean(), 3) if len(confs) else None,
            "Min confidence":          round(confs.min(),  3) if len(confs) else None,
            "N low confidence (<0.6)": (df[low_col] == "YES").sum(),
            "New AI % yes":            round(100*(answers == "yes").sum()/len(df), 1),
            "N questions w/ change from orig AI":
                ((df[ans_col].str.lower() != df[f"{label} | Orig AI answer"].str.lower()) &
                 df[ans_col].notna()).sum(),
        })
    summary_df = pd.DataFrame(summary_rows)

    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Per-paper Results", index=False)
        summary_df.to_excel(writer, sheet_name="Question Summary", index=False)

        # highlight low-confidence cells
        from openpyxl.styles import PatternFill, Font
        ws = writer.sheets["Per-paper Results"]
        ORANGE = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        RED    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        GREEN  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

        header = {cell.value: cell.column for cell in ws[1]}
        for col_name, col_idx in header.items():
            if col_name is None:
                continue
            if "| Low confidence?" in str(col_name):
                for row_idx in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if cell.value == "YES":
                        cell.fill = ORANGE
                        cell.font = Font(bold=True)
            elif "| Orig agreement" in str(col_name):
                for row_idx in range(2, ws.max_row + 1):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    if cell.value == "agree":
                        cell.fill = GREEN
                    elif cell.value == "disagree":
                        cell.fill = RED

        ws.freeze_panes = "E2"

    summary_df.to_excel(SUMMARY_XLSX, index=False)

    print(f"\nDone. {len(df)} papers processed.")
    print(f"  {OUTPUT_XLSX}")
    print(f"  {SUMMARY_XLSX}")
    print()
    print("Questions with lowest average confidence:")
    if not summary_df.empty:
        print(summary_df.nsmallest(5, "Avg confidence")[["Short label", "Avg confidence", "N low confidence (<0.6)"]].to_string(index=False))

if __name__ == "__main__":
    main()
