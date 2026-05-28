"""
Creates manual_review_top10.xlsx — clean, simple annotation sheet.
"""

import pandas as pd
import warnings
warnings.filterwarnings("ignore")
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

comp = pd.read_excel("comparison_ai_vs_student.xlsx", sheet_name="Comparison")
ai   = pd.read_excel("ai_annotated.xlsx")
ai["_geo"] = ai["GEO Series ID (GSE___)"].astype(str).str.strip().str.upper()

agree_cols = [c for c in comp.columns if "__agreement" in c]
comp["n_disagree"] = (comp[agree_cols] == "disagree").sum(axis=1)
top10_geo = comp.nlargest(10, "n_disagree")["GEO_ID"].tolist()

QUESTIONS = [
    ("Birthweight",          "Birthweight of offspring provided (yes/no)"),
    ("GA at delivery",       "Gestational Age at delivery provided (yes/no)"),
    ("GA at collection",     "Gestational Age at sample collection provided (yes/no)"),
    ("Sex of offspring",     "Sex of Offspring Provided (yes/no)"),
    ("N offspring/preg",     "Number of offspring per pregnancy provided (yes/no)"),
    ("Parity",               "Parity provided (yes/no)"),
    ("Gravidity",            "Gravidity provided (yes/no)"),
    ("Race / ethnicity",     "Self-reported race/ethnicity of mother provided (yes/no)"),
    ("Genetic ancestry",     "Genetic ancestry or genetic strain provided (yes/no)"),
    ("Maternal height",      "Maternal Height provided (yes/no)"),
    ("Maternal pre-preg wt", "Maternal Pre-pregnancy Weight provided (yes/no)"),
    ("Maternal age",         "Maternal age at sample collection provided (yes/no)"),
    ("Paternal height",      "Paternal Height provided (yes/no)"),
    ("Paternal weight",      "Paternal Weight provided (yes/no)"),
    ("Paternal age",         "Paternal age at sample collection provided (yes/no)"),
    ("Complication samples", "Samples from pregnancy complications collected"),
    ("Mode of delivery",     "Mode of delivery provided (yes/no)"),
    ("Fetal complications",  "Fetal complications listed (yes/no)"),
]

rows = []
for geo in top10_geo:
    ai_row   = ai[ai["_geo"] == geo]
    comp_row = comp[comp["GEO_ID"] == geo]
    raw_pmid = ai_row.iloc[0].get("PMID","") if not ai_row.empty else ""
    pmid_clean = str(int(float(raw_pmid))) if str(raw_pmid).strip() not in ("","nan") else ""
    row = {
        "GEO ID":           geo,
        "GEO Link":         f"https://www.ncbi.nlm.nih.gov/geo/query/acc.cgi?acc={geo}",
        "PMID":             pmid_clean,
        "PubMed Link":      f"https://pubmed.ncbi.nlm.nih.gov/{pmid_clean}/" if pmid_clean else "",
        "GEO Dataset Title": str(ai_row.iloc[0]["PlTitle"]) if not ai_row.empty else "",
        "# Disagreements":  int(comp_row.iloc[0]["n_disagree"]) if not comp_row.empty else 0,
    }
    for short, _ in QUESTIONS:
        row[short] = ""
    rows.append(row)

df = pd.DataFrame(rows)

NAVY   = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
LGREY  = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
WHITE  = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
YELLOW = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT   = Alignment(horizontal="left",   vertical="center", wrap_text=True)

META_COLS = ["GEO ID", "GEO Link", "PMID", "PubMed Link", "GEO Dataset Title", "# Disagreements"]
ALL_COLS  = META_COLS + [s for s, _ in QUESTIONS]
COL_IDX   = {name: i + 1 for i, name in enumerate(ALL_COLS)}

with pd.ExcelWriter("manual_review_top10.xlsx", engine="openpyxl") as writer:
    df.to_excel(writer, sheet_name="My Annotations", index=False, startrow=1)
    ws = writer.sheets["My Annotations"]

    # row 1: instruction banner
    ws.insert_rows(1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(ALL_COLS))
    c = ws.cell(row=1, column=1)
    c.value     = "Open each paper via PubMed Link or GEO Link → read it → type  yes / no / unclear  in each column.  Note: 'GEO Dataset Title' is the submitter's name for the dataset — the actual paper title on PubMed may be worded differently."
    c.fill      = YELLOW
    c.font      = Font(italic=True, size=10, color="5D4037", name="Calibri")
    c.alignment = LEFT
    ws.row_dimensions[1].height = 22

    # row 2: column headers
    for col_name, col_idx in COL_IDX.items():
        cell = ws.cell(row=2, column=col_idx)
        cell.value     = col_name
        cell.fill      = NAVY
        cell.font      = Font(bold=True, color="FFFFFF", size=9, name="Calibri")
        cell.alignment = CENTER
        cell.border    = thin_border()
    ws.row_dimensions[2].height = 48

    # data rows
    for row_idx in range(3, 3 + len(rows)):
        even = (row_idx % 2 == 0)
        ws.row_dimensions[row_idx].height = 32
        for col_name, col_idx in COL_IDX.items():
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = thin_border()
            if col_name in META_COLS:
                cell.fill      = LGREY if even else WHITE
                cell.font      = Font(size=9, name="Calibri", bold=(col_name == "GEO ID"))
                cell.alignment = LEFT if col_name == "Paper Title" else CENTER
            else:
                cell.fill      = LGREY if even else WHITE
                cell.font      = Font(size=10, name="Calibri")
                cell.alignment = CENTER

    # column widths
    ws.column_dimensions[get_column_letter(COL_IDX["GEO ID"])].width             = 12
    ws.column_dimensions[get_column_letter(COL_IDX["GEO Link"])].width            = 14
    ws.column_dimensions[get_column_letter(COL_IDX["PMID"])].width                = 11
    ws.column_dimensions[get_column_letter(COL_IDX["PubMed Link"])].width         = 16
    ws.column_dimensions[get_column_letter(COL_IDX["GEO Dataset Title"])].width   = 44
    ws.column_dimensions[get_column_letter(COL_IDX["# Disagreements"])].width     = 12
    for short, _ in QUESTIONS:
        ws.column_dimensions[get_column_letter(COL_IDX[short])].width = 15

    ws.freeze_panes = "F3"   # freeze metadata, scroll through questions
    ws.sheet_view.zoomScale = 90

print("Done: manual_review_top10.xlsx")
