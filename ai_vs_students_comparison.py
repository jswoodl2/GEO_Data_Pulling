#!/usr/bin/env python3
"""
Compare student placenta-study annotations with AI-generated metadata
and produce both an Excel summary and a set of figures.

Default paths (can be overridden via CLI):
  - Student sheet: ../geo_2025_Oct_31/Placenta_Study_Information-5.xlsx
  - AI sheet:      ../geo_2025_Oct_31/gse_metadata_full_checkpoint_MERGED_with_final_cleaned.xlsx
  - Outputs (XLSX + PNGs) are written next to those files.
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

import matplotlib.pyplot as plt
import pandas as pd
import seaborn as sns
import textwrap


# GEO ID column name (both sheets use this)
GEO_COL = "GEO Series ID (GSE___)"

# Canonical question fields we want to compare (exact column names in AI sheet)
CANON_QUESTIONS: List[str] = [
    "Pregnancy trimester (1st, 2nd, 3rd, term (for full-term delivery), premature (for early delivery due to complications)",
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
    "Pregnancy complications in data set (list)",
    "Fetal complications listed (yes/no)",
    "Fetal complications in data set (list)",
    "Other Phenotypes Provided (list)",
    "Hospital/Center where samples were collected",
    "Country where samples were collected",
]

TRIMESTER_FIELD = CANON_QUESTIONS[0]

# Choose which of the above are interpreted as Yes/No style
YESNO_FIELDS = {
    q
    for q in CANON_QUESTIONS
    if "(yes/no)" in q.lower() or "listed (yes/no)" in q.lower()
}
YESNO_FIELDS.add("Samples from pregnancy complications collected")

# Pairs of metadata fields (label, student-column-name, ai-column-name)
METADATA_FIELD_PAIRS: List[Tuple[str, str, str]] = [
    ("Data type", "Data Type", "Data type"),
    (
        "SuperSeries list",
        "If SuperSeries, list GEO Series that are part of the SuperSeries",
        "SuperSeries, list GEO Series that are part of the SuperSeries",
    ),
    ("Sample size (placenta)", "Sample size (placenta)", "Sample size (placenta)"),
    ("Title", "Title", "PlTitle"),
    ("Organism", "Organism", "Organism"),
    ("Characteristics", "Characteristics", "Characteristics"),
    ("Extracted molecule", "Extracted molecule", "Extracted molecule"),
    ("Extraction protocol", "Extraction protocol", "Extraction protocol"),
    ("Library Strategy", "Library Strategy", "Library Strategy"),
    ("Library source", "Library source", "Library source"),
    ("Library selection", "Library selection", "Library selection"),
    ("Instrument model", "Instrument model", "Instrument model"),
    ("Assay description", "Assay description", "Assay description"),
    ("Data processing", "Data processing", "Data processing"),
    ("Platform ID (list)", "Platform ID (list)", "Platform ID (list)"),
    ("SRA Study ID (raw data)", "SRA Study ID (raw data)", "SRA Study ID (raw data)"),
    ("BioSample/BioProject ID", "BioSample/BioProject ID", "BioSample/BioProject ID"),
    (
        "File types/resources provided (list)",
        "File types/resources provided (list)",
        "File types/resources provided (list)",
    ),
    ("Submission date", "Submission date", "Submission date"),
    ("Last update date", "Last update date", "Last update date"),
    ("Organization name", "Organization name", "Organization name"),
    ("Contact name", "Contact name", "Contact name"),
    ("E-mail(s)", "E-mail(s)", "E-mail(s)"),
    ("Country", "Country", "Country"),
]


_TRIMESTER_MAP: Dict[str, Iterable[str]] = {
    "1st": {"1", "1st", "first"},
    "2nd": {"2", "2nd", "second"},
    "3rd": {"3", "3rd", "third"},
    "term": {"term", "full-term", "full term"},
    "premature": {"preterm", "premature", "early"},
}


def _wrap_labels(labels: Iterable[str], width: int = 38) -> List[str]:
    return ["\n".join(textwrap.wrap(str(x), width=width)) for x in labels]


def _norm_yesno(x) -> str:
    s = str(x).lower()
    return "Yes" if "yes" in s else "No"


def _norm_trimester(x) -> str:
    s = str(x).strip().lower()
    if not s or s in {"nan", "none"}:
        return "unknown/other"
    for canon, keys in _TRIMESTER_MAP.items():
        for k in keys:
            if f" {k} " in f" {s} ":
                return canon
    if s in {"1", "2", "3"}:
        return {"1": "1st", "2": "2nd", "3": "3rd"}[s]
    return "unknown/other"


def _norm_text(x) -> str:
    if pd.isna(x):
        return ""
    s = str(x).strip().lower()
    return re.sub(r"\s+", " ", s)


def _norm_for_compare(field: str, value) -> str:
    if field in YESNO_FIELDS:
        return _norm_yesno(value)
    if field == TRIMESTER_FIELD:
        return _norm_trimester(value)
    return _norm_text(value)


def _display_value(field: str, value) -> str:
    """Value to show in the spreadsheet; normalize yes/no + trimester only."""
    if field in YESNO_FIELDS:
        return _norm_yesno(value)
    if field == TRIMESTER_FIELD:
        return _norm_trimester(value)
    if pd.isna(value):
        return ""
    return str(value)


def _norm_gse(x) -> str:
    sx = "" if pd.isna(x) else str(x).strip().upper()
    m = re.search(r"(GSE\d{3,})", sx)
    return m.group(1) if m else sx


def _load_and_merge(
    student_path: Path, ai_path: Path, verbose: bool = True
) -> pd.DataFrame:
    student_df = pd.read_excel(student_path)
    ai_df = pd.read_excel(ai_path)

    if GEO_COL not in student_df.columns:
        raise SystemExit(f"{GEO_COL!r} not found in student sheet.")
    if GEO_COL not in ai_df.columns:
        raise SystemExit(f"{GEO_COL!r} not found in AI sheet.")

    student_df = student_df.copy()
    ai_df = ai_df.copy()
    student_df["GEO_KEY"] = student_df[GEO_COL].map(_norm_gse)
    ai_df["GEO_KEY"] = ai_df[GEO_COL].map(_norm_gse)

    student_df = student_df.drop_duplicates(subset=["GEO_KEY"])
    ai_df = ai_df.drop_duplicates(subset=["GEO_KEY"])

    merged = pd.merge(
        student_df,
        ai_df,
        on="GEO_KEY",
        how="inner",
        suffixes=(" — Student", " — AI"),
    )

    if verbose:
        print(f"Student rows: {len(student_df)}, AI rows: {len(ai_df)}")
        print(f"Joined rows (both student + AI): {len(merged)}")

    return merged


def build_comparison(
    student_path: Path, ai_path: Path
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    merged = _load_and_merge(student_path, ai_path, verbose=True)

    # Per-study comparison table
    records: List[dict] = []
    stats = {q: {"n": 0, "matches": 0} for q in CANON_QUESTIONS}

    for _, row in merged.iterrows():
        base = {"GEO_KEY": row["GEO_KEY"]}
        for field in CANON_QUESTIONS:
            s_col = f"{field} — Student"
            a_col = f"{field} — AI"
            s_raw = row.get(s_col, "")
            a_raw = row.get(a_col, "")

            s_norm = _norm_for_compare(field, s_raw)
            a_norm = _norm_for_compare(field, a_raw)
            match = s_norm == a_norm

            base[f"{field} — Student"] = _display_value(field, s_raw)
            base[f"{field} — AI"] = _display_value(field, a_raw)
            base[f"{field} — Match"] = "Yes" if match else "No"

            stats[field]["n"] += 1
            stats[field]["matches"] += int(match)

        records.append(base)

    per_study_df = pd.DataFrame.from_records(records)

    # Field-level summary
    summary_rows = []
    for field, d in stats.items():
        n = d["n"]
        k = d["matches"]
        rate = (k / n) if n else 0.0
        summary_rows.append(
            {"Field": field, "Compared": n, "Matches": k, "AgreementRate": round(rate, 4)}
        )

    summary_df = pd.DataFrame(summary_rows).sort_values(
        "AgreementRate", ascending=False
    )
    return per_study_df, summary_df


def build_metadata_comparison(
    student_path: Path, ai_path: Path
) -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Compare higher-level metadata fields like Data type, Title, etc.
    Uses explicit (student-column, ai-column) pairs defined in METADATA_FIELD_PAIRS.
    """
    merged = _load_and_merge(student_path, ai_path, verbose=False)

    records: List[dict] = []
    stats = {label: {"n": 0, "matches": 0} for label, _, _ in METADATA_FIELD_PAIRS}

    for _, row in merged.iterrows():
        base = {"GEO_KEY": row["GEO_KEY"]}
        for label, stu_col_name, ai_col_name in METADATA_FIELD_PAIRS:
            s_col = f"{stu_col_name} — Student"
            a_col = f"{ai_col_name} — AI"
            s_raw = row.get(s_col, "")
            a_raw = row.get(a_col, "")

            s_norm = _norm_text(s_raw)
            a_norm = _norm_text(a_raw)
            match = s_norm == a_norm

            base[f"{label} — Student"] = "" if pd.isna(s_raw) else str(s_raw)
            base[f"{label} — AI"] = "" if pd.isna(a_raw) else str(a_raw)
            base[f"{label} — Match"] = "Yes" if match else "No"

            stats[label]["n"] += 1
            stats[label]["matches"] += int(match)

        records.append(base)

    per_study_df = pd.DataFrame.from_records(records)

    summary_rows = []
    for label, d in stats.items():
        n = d["n"]
        k = d["matches"]
        rate = (k / n) if n else 0.0
        summary_rows.append(
            {"Field": label, "Compared": n, "Matches": k, "AgreementRate": round(rate, 4)}
        )

    summary_df = pd.DataFrame(summary_rows).sort_values(
        "AgreementRate", ascending=False
    )
    return per_study_df, summary_df


def make_figures(
    per_study_df: pd.DataFrame, summary_df: pd.DataFrame, output_prefix: str = "comparison_figs"
) -> None:
    # Defensive copy
    df = per_study_df.copy()

    # 1) Agreement Rate by Field (bigger canvas + wrapped labels)
    plt.figure(figsize=(12, max(8, 0.55 * len(summary_df))))
    plot_df = summary_df.copy().sort_values("AgreementRate", ascending=True)
    sns.barplot(data=plot_df, y="Field", x="AgreementRate")
    plt.yticks(
        ticks=range(len(plot_df)),
        labels=_wrap_labels(plot_df["Field"].tolist(), width=45),
        fontsize=9,
    )
    plt.title("Agreement Rate by Metadata Field", fontsize=14, pad=12)
    plt.xlabel("Agreement Rate")
    plt.ylabel("")
    plt.tight_layout()
    plt.savefig(f"{output_prefix}_agreement_by_field.png", dpi=300)
    plt.close()

    # 2) Yes/No confusion matrix (counts + percentages)
    yesno_match_cols = [
        c
        for c in df.columns
        if " — Match" in c and "(yes/no)" in c.split(" — ")[0].lower()
    ]
    cm_frames: List[pd.DataFrame] = []
    for mcol in yesno_match_cols:
        base = mcol.split(" — ")[0]
        s_col = f"{base} — Student"
        a_col = f"{base} — AI"
        if s_col in df.columns and a_col in df.columns:
            tmp = df[[s_col, a_col]].copy()
            tmp.columns = ["Student", "AI"]
            tmp["Student"] = tmp["Student"].map(_norm_yesno)
            tmp["AI"] = tmp["AI"].map(_norm_yesno)
            cm_frames.append(tmp)

    if cm_frames:
        big = pd.concat(cm_frames, ignore_index=True)
        cm_counts = pd.crosstab(big["Student"], big["AI"])
        cm_perc = cm_counts / cm_counts.to_numpy().sum() * 100.0

        plt.figure(figsize=(8, 6))
        ax = sns.heatmap(cm_counts, annot=True, fmt="d", cbar=False, cmap="Blues")
        ax.set_title("AI vs Student (Yes/No Fields Combined) — Counts", pad=12)
        plt.savefig(
            f"{output_prefix}_yesno_confusion_counts.png",
            dpi=300,
            bbox_inches="tight",
        )
        plt.close()

        plt.figure(figsize=(8, 6))
        ax = sns.heatmap(cm_perc, annot=True, fmt=".1f", cbar=False, cmap="Blues")
        ax.set_title("AI vs Student (Yes/No Fields Combined) — Percent", pad=12)
        plt.savefig(
            f"{output_prefix}_yesno_confusion_percent.png",
            dpi=300,
            bbox_inches="tight",
        )
        plt.close()

    # 3) Per-study agreement histogram
    match_cols = [c for c in df.columns if " — Match" in c]
    if match_cols:
        rates = (
            df[match_cols].apply(lambda r: sum(v == "Yes" for v in r), axis=1)
            / len(match_cols)
        )
        plt.figure(figsize=(10, 6))
        sns.histplot(rates, bins=12)
        plt.title("Distribution of Agreement Rates per Study", fontsize=16, pad=12)
        plt.xlabel("Agreement Rate")
        plt.ylabel("Number of Studies")
        plt.xlim(0, 1)
        plt.tight_layout()
        plt.savefig(f"{output_prefix}_perstudy_hist.png", dpi=300)
        plt.close()

    # 4) Trimester distribution — normalized and tidy
    field = TRIMESTER_FIELD
    s_col = f"{field} — Student"
    a_col = f"{field} — AI"
    if s_col in df.columns and a_col in df.columns:
        tidy = pd.DataFrame(
            {
                "Student": df[s_col].map(_norm_trimester),
                "AI": df[a_col].map(_norm_trimester),
            }
        )
        long = tidy.melt(var_name="Source", value_name="Trimester")

        # Count per (Source, Trimester), then convert to percent within each Source.
        counts = (
            long.groupby(["Source", "Trimester"])
            .size()
            .reset_index(name="Count")
        )
        counts["Percent"] = (
            counts.groupby("Source")["Count"].transform(lambda s: s / s.sum() * 100)
        )

        order = ["1st", "2nd", "3rd", "term", "premature", "unknown/other"]

        plt.figure(figsize=(12, 6))
        sns.barplot(
            data=counts, x="Trimester", y="Percent", hue="Source", order=order
        )
        plt.title("Trimester Distribution — Student vs AI (Percent)", fontsize=14, pad=10)
        plt.xlabel("")
        plt.ylabel("Percent of Studies")
        plt.ylim(0, 100)
        plt.legend(title="", loc="upper right")
        plt.tight_layout()
        plt.savefig(f"{output_prefix}_trimester_percent.png", dpi=300)
        plt.close()

    print("Figures saved with prefix:", output_prefix)


def main() -> None:
    default_student = Path("..") / "geo_2025_Oct_31" / "Placenta_Study_Information-5.xlsx"
    default_ai = (
        Path("..")
        / "geo_2025_Oct_31"
        / "gse_metadata_full_checkpoint_MERGED_with_final_cleaned.xlsx"
    )
    default_out_xlsx = (
        Path("..") / "geo_2025_Oct_31" / "ai_vs_students_comparison.xlsx"
    )

    parser = argparse.ArgumentParser(
        description=(
            "Compare student placenta-study annotations with AI metadata "
            "and generate an Excel summary plus plots."
        )
    )
    parser.add_argument(
        "--student",
        type=Path,
        default=default_student,
        help=f"Path to student annotations (default: {default_student})",
    )
    parser.add_argument(
        "--ai",
        type=Path,
        default=default_ai,
        help=f"Path to AI annotations (default: {default_ai})",
    )
    parser.add_argument(
        "--out-xlsx",
        type=Path,
        default=default_out_xlsx,
        help=f"Path for Excel output (default: {default_out_xlsx})",
    )
    parser.add_argument(
        "--out-prefix",
        type=str,
        default=str(default_out_xlsx.with_suffix("")),
        help="Prefix for figure filenames (default: Excel path without extension).",
    )

    args = parser.parse_args()

    per_study_df, summary_df = build_comparison(args.student, args.ai)

    # Save questionnaire-style comparison Excel
    args.out_xlsx.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(args.out_xlsx, engine="xlsxwriter") as xw:
        per_study_df.to_excel(xw, index=False, sheet_name="PerStudyComparison")
        summary_df.to_excel(xw, index=False, sheet_name="FieldSummary")

    print(f"Wrote comparison workbook to: {args.out_xlsx}")

    # Make plots for questionnaire-style fields
    make_figures(per_study_df, summary_df, output_prefix=args.out_prefix)

    # ---- Additional comparison for high-level metadata fields ----
    meta_per_study_df, meta_summary_df = build_metadata_comparison(args.student, args.ai)

    meta_out_xlsx = args.out_xlsx.with_name(args.out_xlsx.stem + "_metadata.xlsx")
    with pd.ExcelWriter(meta_out_xlsx, engine="xlsxwriter") as xw:
        meta_per_study_df.to_excel(xw, index=False, sheet_name="PerStudyComparison")
        meta_summary_df.to_excel(xw, index=False, sheet_name="FieldSummary")

    print(f"Wrote metadata comparison workbook to: {meta_out_xlsx}")

    make_figures(
        meta_per_study_df,
        meta_summary_df,
        output_prefix=args.out_prefix + "_metadata",
    )


if __name__ == "__main__":
    main()
