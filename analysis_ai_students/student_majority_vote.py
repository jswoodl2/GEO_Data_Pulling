"""
Build a majority-vote consensus from the 3 student annotation sets
in manual_review_annotated.xlsx.

For yes/no questions: takes the majority (2 out of 3 wins).
For free-text: takes the most common answer, or the longest if all differ.
Outputs: student_majority_vote.xlsx
"""

import os
from collections import Counter
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill

INPUT_FILE = os.path.join(os.path.dirname(__file__), "manual_review_annotated.xlsx")
OUTPUT_FILE = os.path.join(os.path.dirname(__file__), "student_majority_vote.xlsx")

YESNO_QUESTIONS = {
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
}


def normalize_yesno(val):
    """Normalize yes/no answers."""
    if val is None:
        return None
    s = str(val).strip().lower()
    if s in ("yes", "y", "true", "1"):
        return "Yes"
    if s in ("no", "n", "false", "0"):
        return "No"
    return None  # not a clear yes/no


def majority_vote_yesno(answers):
    """Take majority vote for yes/no. Returns (consensus, vote_detail)."""
    normalized = [normalize_yesno(a) for a in answers]
    valid = [a for a in normalized if a is not None]
    if not valid:
        return "N/A", "all blank"

    counts = Counter(valid)
    winner, count = counts.most_common(1)[0]
    total = len(valid)
    detail = f"{count}/{total} {winner}"
    return winner, detail


def majority_vote_freetext(answers):
    """Take majority for free text. If all differ, pick longest."""
    cleaned = []
    for a in answers:
        if a is None:
            cleaned.append("")
        else:
            cleaned.append(str(a).strip())

    # normalize for comparison (lowercase, strip)
    normalized = [a.lower().replace("n/a", "").strip() for a in cleaned]

    # count normalized versions
    counts = Counter(n for n in normalized if n)
    if not counts:
        return "Not Provided", "all blank"

    winner_norm, count = counts.most_common(1)[0]
    total = len([n for n in normalized if n])

    # return the original (non-lowered) version of the winner
    for a, n in zip(cleaned, normalized):
        if n == winner_norm:
            return a, f"{count}/{total} agree"

    return cleaned[0], f"1/{total}"


def main():
    wb_in = openpyxl.load_workbook(INPUT_FILE)
    sheets = wb_in.sheetnames  # ['Set1', 'Set2', 'Set3']
    print(f"Sheets: {sheets}")

    # read all data from all sheets
    all_data = {}  # {sheet_name: [{header: value, ...}, ...]}
    headers = None

    for sheet_name in sheets:
        ws = wb_in[sheet_name]
        # row 2 = headers
        sheet_headers = [c.value for c in ws[2]]
        if headers is None:
            headers = sheet_headers

        rows = []
        for row in ws.iter_rows(min_row=3, max_row=ws.max_row, values_only=True):
            geo_id = row[0]
            if not geo_id:
                continue
            row_dict = {}
            for h, v in zip(headers, row):
                if h:
                    row_dict[h] = v
            rows.append(row_dict)
        all_data[sheet_name] = rows
        print(f"  {sheet_name}: {len(rows)} papers")

    # get the question columns (skip geo id, geo link, pmid, geo title)
    question_headers = [h for h in headers if h and h not in ("GEO ID", "GEO Link", "PMID", "GEO Title")]

    # get paper list from first sheet
    geo_ids = [r["GEO ID"] for r in all_data[sheets[0]]]

    # build majority vote
    wb_out = openpyxl.Workbook()

    # sheet 1: consensus answers
    ws_consensus = wb_out.active
    ws_consensus.title = "Majority Vote"

    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, size=10)
    agree_fill = PatternFill(start_color="C6EFCE", fill_type="solid")  # green
    split_fill = PatternFill(start_color="FFEB9C", fill_type="solid")  # yellow
    disagree_fill = PatternFill(start_color="FFC7CE", fill_type="solid")  # red

    out_headers = ["GEO ID", "GEO Link", "PMID", "GEO Title"] + question_headers
    for col, h in enumerate(out_headers, 1):
        cell = ws_consensus.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    for paper_idx, geo_id in enumerate(geo_ids):
        row_num = paper_idx + 2

        # get this paper's row from each set
        paper_rows = []
        for sheet_name in sheets:
            for r in all_data[sheet_name]:
                if r.get("GEO ID") == geo_id:
                    paper_rows.append(r)
                    break

        if not paper_rows:
            continue

        # write identifiers from first set
        ws_consensus.cell(row=row_num, column=1, value=paper_rows[0].get("GEO ID"))
        ws_consensus.cell(row=row_num, column=2, value=paper_rows[0].get("GEO Link"))
        ws_consensus.cell(row=row_num, column=3, value=paper_rows[0].get("PMID"))
        ws_consensus.cell(row=row_num, column=4, value=paper_rows[0].get("GEO Title"))

        for q_idx, q in enumerate(question_headers):
            answers = [r.get(q) for r in paper_rows]

            # check if this is a yes/no question
            is_yesno = q in YESNO_QUESTIONS or any(
                normalize_yesno(a) is not None for a in answers if a
            )

            if is_yesno and q in YESNO_QUESTIONS:
                consensus, detail = majority_vote_yesno(answers)
            else:
                consensus, detail = majority_vote_freetext(answers)

            cell = ws_consensus.cell(row=row_num, column=5 + q_idx, value=consensus)
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    # sheet 2: detailed vote breakdown
    ws_detail = wb_out.create_sheet("Vote Details")
    detail_headers = ["GEO ID", "Question", "Set1", "Set2", "Set3", "Consensus", "Agreement"]
    for col, h in enumerate(detail_headers, 1):
        cell = ws_detail.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font

    detail_row = 2
    total_questions = 0
    unanimous_count = 0
    majority_count = 0
    split_count = 0

    for geo_id in geo_ids:
        paper_rows = []
        for sheet_name in sheets:
            for r in all_data[sheet_name]:
                if r.get("GEO ID") == geo_id:
                    paper_rows.append(r)
                    break

        if len(paper_rows) < 2:
            continue

        for q in question_headers:
            answers = [r.get(q) for r in paper_rows]
            is_yesno = q in YESNO_QUESTIONS

            if is_yesno:
                consensus, detail = majority_vote_yesno(answers)
                normed = [normalize_yesno(a) or str(a) for a in answers]
                valid = [a for a in normed if a in ("Yes", "No")]
                if len(set(valid)) <= 1 and len(valid) >= 2:
                    agreement = "Unanimous"
                    unanimous_count += 1
                elif len(valid) >= 2:
                    agreement = "Majority"
                    majority_count += 1
                else:
                    agreement = "Insufficient"
                    split_count += 1
            else:
                consensus, detail = majority_vote_freetext(answers)
                agreement = detail

            total_questions += 1

            ws_detail.cell(row=detail_row, column=1, value=geo_id)
            ws_detail.cell(row=detail_row, column=2, value=q)
            for s_idx, a in enumerate(answers):
                ws_detail.cell(row=detail_row, column=3 + s_idx, value=str(a) if a else "")
            ws_detail.cell(row=detail_row, column=6, value=consensus)

            cell = ws_detail.cell(row=detail_row, column=7, value=agreement)
            if "Unanimous" in agreement or "3/3" in agreement:
                cell.fill = agree_fill
            elif "Majority" in agreement or "2/3" in agreement:
                cell.fill = split_fill
            else:
                cell.fill = disagree_fill

            detail_row += 1

    # auto-width
    for ws in [ws_consensus, ws_detail]:
        for col in ws.columns:
            max_len = max(len(str(c.value or "")[:60]) for c in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 45)

    wb_out.save(OUTPUT_FILE)
    print(f"\nSaved: {OUTPUT_FILE}")
    print(f"Total question/paper combos: {total_questions}")
    if total_questions:
        yn_total = unanimous_count + majority_count + split_count
        if yn_total:
            print(f"Yes/No questions:")
            print(f"  Unanimous (3/3): {unanimous_count} ({100*unanimous_count/yn_total:.0f}%)")
            print(f"  Majority  (2/1): {majority_count} ({100*majority_count/yn_total:.0f}%)")
            print(f"  Split/Other:     {split_count} ({100*split_count/yn_total:.0f}%)")


if __name__ == "__main__":
    main()
