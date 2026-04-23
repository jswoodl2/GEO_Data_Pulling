"""
AI vs Student Annotation Comparison
Produces:
  1. comparison_ai_vs_student.xlsx  - full merged comparison + agreement per paper
  2. second_human_review_template.xlsx - blank template for 2nd human pass
  3. summary_stats.xlsx - per-question agreement, bias, confusion counts
  4. plots/ - bar charts for agreement rate and yes-rate bias
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import os
import warnings
warnings.filterwarnings("ignore")

print("Loading spreadsheets...")
ai = pd.read_excel("ai_annotated.xlsx")
student = pd.read_excel("student_annotated.xlsx")

def clean_geo_id(s):
    if pd.isna(s):
        return np.nan
    return str(s).strip().upper()

ai["_geo"] = ai["GEO Series ID (GSE___)"].apply(clean_geo_id)
student["_geo"] = student["GEO Series ID (GSE___)"].apply(clean_geo_id)

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

YES_TOKENS = {"yes", "y", "true", "1"}
NO_TOKENS  = {"no",  "n", "false", "0"}

def normalise_yn(raw):
    """
    Returns 'yes', 'no', or 'unclear'.
    For multi-value entries (e.g. 'yes; no; no') we take the FIRST token.
    """
    if pd.isna(raw):
        return "unclear"
    s = str(raw).strip().lower()
    # take first token if semicolon-separated
    first = s.split(";")[0].strip()
    # strip trailing punctuation / extra spaces
    first = first.strip(" .,-()")
    if first in YES_TOKENS:
        return "yes"
    if first in NO_TOKENS:
        return "no"
    # partial match
    if first.startswith("yes"):
        return "yes"
    if first.startswith("no"):
        return "no"
    return "unclear"

for col in YN_QUESTIONS:
    if col in ai.columns:
        ai[col + "__norm"] = ai[col].apply(normalise_yn)
    if col in student.columns:
        student[col + "__norm"] = student[col].apply(normalise_yn)

# keep only first occurrence per geo id in each dataframe (ai rows are unique)
ai_dedup     = ai.drop_duplicates(subset=["_geo"])
student_dedup = student.drop_duplicates(subset=["_geo"])

# identify columns to bring in
ai_keep = ["_geo", "GEO Series ID (GSE___)", "PlTitle"] + \
          [col + "__norm" for col in YN_QUESTIONS if col in ai.columns] + \
          [col for col in YN_QUESTIONS if col in ai.columns]

student_keep = ["_geo"] + \
               ["Student Assigned (Data Annotation)", "Date of first data annotation"] + \
               [col + "__norm" for col in YN_QUESTIONS if col in student.columns] + \
               [col for col in YN_QUESTIONS if col in student.columns]

merged = pd.merge(
    ai_dedup[ai_keep],
    student_dedup[student_keep],
    on="_geo",
    suffixes=("__ai", "__student"),
    how="inner"
)

print(f"Overlapping GEO IDs matched: {len(merged)}")

for col in YN_QUESTIONS:
    ai_col  = col + "__norm__ai"
    stu_col = col + "__norm__student"
    if ai_col in merged.columns and stu_col in merged.columns:
        agree = []
        for a, s in zip(merged[ai_col], merged[stu_col]):
            if a == "unclear" or s == "unclear":
                agree.append("unclear")
            elif a == s:
                agree.append("agree")
            else:
                agree.append("disagree")
        merged[SHORT_LABELS[col] + "__agreement"] = agree

rows = []
for col in YN_QUESTIONS:
    ai_col  = col + "__norm__ai"
    stu_col = col + "__norm__student"
    label   = SHORT_LABELS[col]
    if ai_col not in merged.columns or stu_col not in merged.columns:
        continue

    ai_vals  = merged[ai_col]
    stu_vals = merged[stu_col]

    # filter to rows where both are clear
    mask = (ai_vals != "unclear") & (stu_vals != "unclear")
    n_comparable = mask.sum()
    if n_comparable == 0:
        continue

    ai_c  = ai_vals[mask]
    stu_c = stu_vals[mask]

    agree_n       = (ai_c == stu_c).sum()
    agreement_pct = 100 * agree_n / n_comparable

    # confusion counts
    tp = ((ai_c == "yes") & (stu_c == "yes")).sum()   # AI yes, Student yes
    fp = ((ai_c == "yes") & (stu_c == "no")).sum()    # AI yes, Student no
    fn = ((ai_c == "no")  & (stu_c == "yes")).sum()   # AI no,  Student yes
    tn = ((ai_c == "no")  & (stu_c == "no")).sum()    # AI no,  Student no

    ai_yes_pct  = 100 * (ai_c  == "yes").sum() / n_comparable
    stu_yes_pct = 100 * (stu_c == "yes").sum() / n_comparable
    bias        = ai_yes_pct - stu_yes_pct          # positive = AI more permissive

    n_unclear_ai  = (ai_vals  == "unclear").sum()
    n_unclear_stu = (stu_vals == "unclear").sum()

    rows.append({
        "Question (short)":          label,
        "Question (full)":           col,
        "N comparable":              n_comparable,
        "N unclear (AI)":            n_unclear_ai,
        "N unclear (Student)":       n_unclear_stu,
        "Agreement (n)":             agree_n,
        "Agreement (%)":             round(agreement_pct, 1),
        "AI % yes":                  round(ai_yes_pct, 1),
        "Student % yes":             round(stu_yes_pct, 1),
        "Bias (AI minus Student %)": round(bias, 1),
        "TP (both yes)":             tp,
        "FP (AI yes, Stu no)":       fp,
        "FN (AI no,  Stu yes)":      fn,
        "TN (both no)":              tn,
    })

summary = pd.DataFrame(rows)

# rename for readability
out_cols = {"_geo": "GEO_ID", "GEO Series ID (GSE___)": "GEO Series ID"}
merged_out = merged.rename(columns=out_cols)

# reorder: id, title, student, then per-question triples (ai_raw, stu_raw, agreement)
base_cols = ["GEO_ID", "GEO Series ID", "PlTitle",
             "Student Assigned (Data Annotation)", "Date of first data annotation"]
question_cols = []
for col in YN_QUESTIONS:
    label = SHORT_LABELS[col]
    for suffix, tag in [("__norm__ai",      "AI answer"),
                        ("__norm__student",  "Student answer"),
                        ("__agreement",      "Agreement")]:
        cname = (col + suffix) if suffix != "__agreement" else (label + suffix)
        if cname in merged_out.columns:
            question_cols.append(cname)

all_cols = [c for c in base_cols if c in merged_out.columns] + \
           [c for c in question_cols if c in merged_out.columns]
merged_out = merged_out[all_cols]

template_rows = []
for _, row in merged_out.iterrows():
    base = {
        "GEO_ID":          row.get("GEO_ID", ""),
        "GEO Series ID":   row.get("GEO Series ID", ""),
        "Paper Title":     row.get("PlTitle", ""),
        "Original Student": row.get("Student Assigned (Data Annotation)", ""),
        "Reviewer 2 Name": "",
        "Reviewer 2 Date": "",
    }
    for col in YN_QUESTIONS:
        label = SHORT_LABELS[col]
        ai_ans  = row.get(col + "__norm__ai",     "")
        stu_ans = row.get(col + "__norm__student", "")
        agree   = row.get(label + "__agreement",  "")
        base[f"{label} | AI"]         = ai_ans
        base[f"{label} | Student"]    = stu_ans
        base[f"{label} | Agreement"]  = agree
        base[f"{label} | Reviewer 2"] = ""   # blank for human to fill
        base[f"{label} | Notes"]      = ""
    template_rows.append(base)

template_df = pd.DataFrame(template_rows)

print("Writing comparison_ai_vs_student.xlsx ...")
with pd.ExcelWriter("comparison_ai_vs_student.xlsx", engine="openpyxl") as writer:
    merged_out.to_excel(writer, sheet_name="Comparison", index=False)
    summary.to_excel(writer, sheet_name="Summary Stats", index=False)

    ws = writer.sheets["Comparison"]
    # colour-code agreement column: agree=green, disagree=red, unclear=yellow
    from openpyxl.styles import PatternFill
    GREEN  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    RED    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    YELLOW = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")

    header = {cell.value: cell.column for cell in ws[1]}
    agree_cols = [c for c in header if "__agreement" in str(c)]
    for col_name in agree_cols:
        col_idx = header[col_name]
        for row_idx in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if cell.value == "agree":
                cell.fill = GREEN
            elif cell.value == "disagree":
                cell.fill = RED
            elif cell.value == "unclear":
                cell.fill = YELLOW

print("Writing second_human_review_template.xlsx ...")
with pd.ExcelWriter("second_human_review_template.xlsx", engine="openpyxl") as writer:
    template_df.to_excel(writer, sheet_name="Second Review", index=False)
    ws = writer.sheets["Second Review"]

    from openpyxl.styles import PatternFill, Font, Alignment
    BLUE_FILL   = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    ORANGE_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    GREY_FILL   = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    GREEN_FILL  = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    header = {cell.value: cell.column for cell in ws[1]}
    for col_name, col_idx in header.items():
        if col_name is None:
            continue
        col_name_s = str(col_name)
        hdr_cell = ws.cell(row=1, column=col_idx)
        hdr_cell.font = Font(bold=True)
        hdr_cell.alignment = Alignment(wrap_text=True)
        if "| AI" in col_name_s:
            hdr_cell.fill = BLUE_FILL
        elif "| Student" in col_name_s:
            hdr_cell.fill = ORANGE_FILL
        elif "| Agreement" in col_name_s:
            hdr_cell.fill = GREY_FILL
        elif "| Reviewer 2" in col_name_s:
            hdr_cell.fill = GREEN_FILL

    ws.row_dimensions[1].height = 60
    ws.freeze_panes = "G2"

print("Writing summary_stats.xlsx ...")
summary.to_excel("summary_stats.xlsx", index=False)

os.makedirs("plots", exist_ok=True)

labels = summary["Question (short)"].tolist()
x      = np.arange(len(labels))

# plot 1: agreement % per question
fig, ax = plt.subplots(figsize=(14, 6))
bars = ax.bar(x, summary["Agreement (%)"], color="steelblue", edgecolor="white", width=0.6)
ax.axhline(summary["Agreement (%)"].mean(), color="crimson", linestyle="--", lw=1.5,
           label=f"Mean = {summary['Agreement (%)'].mean():.1f}%")
ax.set_xticks(x)
ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=9)
ax.set_ylabel("Agreement (%)", fontsize=11)
ax.set_title("AI vs Student Agreement Rate per Yes/No Question\n(overlapping papers, n={})"
             .format(len(merged)), fontsize=12)
ax.set_ylim(0, 105)
ax.legend()
for bar, val in zip(bars, summary["Agreement (%)"]):
    ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.8,
            f"{val:.0f}%", ha="center", va="bottom", fontsize=7.5)
plt.tight_layout()
plt.savefig("plots/agreement_rate_per_question.png", dpi=150)
plt.close()
print("Saved plots/agreement_rate_per_question.png")

# plot 2: % yes bias (ai vs student)
fig, ax = plt.subplots(figsize=(14, 6))
width = 0.35
bars_ai  = ax.bar(x - width/2, summary["AI % yes"],      width, label="AI",      color="#4472C4", edgecolor="white")
bars_stu = ax.bar(x + width/2, summary["Student % yes"], width, label="Student",  color="#ED7D31", edgecolor="white")
ax.set_xticks(x)
ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=9)
ax.set_ylabel("% Yes", fontsize=11)
ax.set_title("% 'Yes' Responses: AI vs Student per Question\n(overlapping papers, n={})"
             .format(len(merged)), fontsize=12)
ax.set_ylim(0, 105)
ax.legend()
plt.tight_layout()
plt.savefig("plots/yes_rate_bias_ai_vs_student.png", dpi=150)
plt.close()
print("Saved plots/yes_rate_bias_ai_vs_student.png")

# plot 3: bias bar (ai minus student)
fig, ax = plt.subplots(figsize=(14, 5))
colors = ["#4472C4" if b >= 0 else "#ED7D31" for b in summary["Bias (AI minus Student %)"]]
bars = ax.bar(x, summary["Bias (AI minus Student %)"], color=colors, edgecolor="white", width=0.6)
ax.axhline(0, color="black", lw=0.8)
ax.set_xticks(x)
ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=9)
ax.set_ylabel("AI % Yes  −  Student % Yes", fontsize=11)
ax.set_title("Bias: AI minus Student 'Yes' Rate\n(positive = AI more permissive, negative = Student more permissive)",
             fontsize=11)
blue_patch   = mpatches.Patch(color="#4472C4", label="AI more permissive")
orange_patch = mpatches.Patch(color="#ED7D31", label="Student more permissive")
ax.legend(handles=[blue_patch, orange_patch])
for bar, val in zip(bars, summary["Bias (AI minus Student %)"]):
    ypos = bar.get_height() + 0.3 if val >= 0 else bar.get_height() - 1.5
    ax.text(bar.get_x() + bar.get_width()/2, ypos, f"{val:+.1f}",
            ha="center", va="bottom", fontsize=7.5)
plt.tight_layout()
plt.savefig("plots/bias_ai_minus_student.png", dpi=150)
plt.close()
print("Saved plots/bias_ai_minus_student.png")

# plot 4: stacked agreement breakdown
fig, ax = plt.subplots(figsize=(14, 5))
agree_col   = label + "__agreement"

# recount per question
agree_n     = summary["TP (both yes)"] + summary["TN (both no)"]
disagree_fp = summary["FP (AI yes, Stu no)"]
disagree_fn = summary["FN (AI no,  Stu yes)"]
n_comp      = summary["N comparable"]

pct_agree   = 100 * agree_n   / n_comp
pct_fp      = 100 * disagree_fp / n_comp
pct_fn      = 100 * disagree_fn / n_comp

ax.bar(x, pct_agree, label="Agree",               color="#70AD47", edgecolor="white")
ax.bar(x, pct_fp,    bottom=pct_agree,             label="Disagree: AI yes, Student no", color="#FF0000", edgecolor="white")
ax.bar(x, pct_fn,    bottom=pct_agree + pct_fp,   label="Disagree: AI no, Student yes", color="#FF7F00", edgecolor="white")

ax.set_xticks(x)
ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=9)
ax.set_ylabel("% of comparable rows", fontsize=11)
ax.set_title("Agreement Breakdown per Question (stacked)", fontsize=12)
ax.legend(loc="lower right")
ax.set_ylim(0, 108)
plt.tight_layout()
plt.savefig("plots/agreement_breakdown_stacked.png", dpi=150)
plt.close()
print("Saved plots/agreement_breakdown_stacked.png")

print("\n" + "="*70)
print("SUMMARY: AI vs Student Agreement")
print("="*70)
print(f"{'Question':<35} {'N':>5} {'Agree%':>8} {'AI%yes':>8} {'Stu%yes':>8} {'Bias':>7}")
print("-"*70)
for _, r in summary.iterrows():
    print(f"{r['Question (short)']:<35} {r['N comparable']:>5} "
          f"{r['Agreement (%)']:>7.1f}% {r['AI % yes']:>7.1f}% "
          f"{r['Student % yes']:>7.1f}% {r['Bias (AI minus Student %)']:>+6.1f}%")
print("="*70)
print(f"\nOverall mean agreement: {summary['Agreement (%)'].mean():.1f}%")
print(f"Questions with <80% agreement: {(summary['Agreement (%)'] < 80).sum()}")
print(f"Questions with >10pp AI bias:  {(summary['Bias (AI minus Student %)'].abs() > 10).sum()}")
print("\nOutputs written:")
print("  comparison_ai_vs_student.xlsx")
print("  second_human_review_template.xlsx")
print("  summary_stats.xlsx")
print("  plots/agreement_rate_per_question.png")
print("  plots/yes_rate_bias_ai_vs_student.png")
print("  plots/bias_ai_minus_student.png")
print("  plots/agreement_breakdown_stacked.png")
