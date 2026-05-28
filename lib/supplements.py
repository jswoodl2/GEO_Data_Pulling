"""
Supplement text extraction for placenta papers.

Reads every supplement file in downloaded_supplements/<PMCID>/ and
concatenates the readable text. Handles .docx, .doc, .xlsx, .xls,
.pdf, .csv, .txt. Skips the tiny "Preparing your download..." HTML
interstitials that PMC sometimes serves with a .pdf filename.

Public API:
    load_supplements(pmcid, base_dir=None) -> str

Tunables (change at any time):
    XLSX_ROW_CAP            max rows we read per spreadsheet sheet
    SUPPLEMENT_CHAR_CAP     total char cap on the concatenated output
    INTERSTITIAL_BYTE_LIMIT files smaller than this AND starting with html
                            markers are skipped (PMC interstitial pages)
"""

from __future__ import annotations

import os
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Iterable, Optional

import openpyxl

# config
XLSX_ROW_CAP = 200
SUPPLEMENT_CHAR_CAP = 80_000
INTERSTITIAL_BYTE_LIMIT = 3000


def extract_docx_text(path: str) -> str:
    """Pull text + tables out of a .docx file."""
    try:
        with zipfile.ZipFile(path) as z:
            with z.open("word/document.xml") as f:
                tree = ET.parse(f)
                root = tree.getroot()
                ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
                lines: list[str] = []
                # paragraphs
                for p in root.findall(".//w:p", ns):
                    texts = [t.text for t in p.findall(".//w:t", ns) if t.text]
                    line = "".join(texts).strip()
                    if line:
                        lines.append(line)
                # tables
                for t_idx, table in enumerate(root.findall(".//w:tbl", ns)):
                    lines.append(f"\n[Table {t_idx + 1}]")
                    for row in table.findall(".//w:tr", ns):
                        cells: list[str] = []
                        for cell in row.findall(".//w:tc", ns):
                            ct = "".join(t.text for t in cell.findall(".//w:t", ns) if t.text)
                            cells.append(ct.strip())
                        lines.append(" | ".join(cells))
                return "\n".join(lines)
    except Exception as e:
        return f"[Error reading {os.path.basename(path)}: {e}]"


def extract_xlsx_text(path: str, max_rows: int = XLSX_ROW_CAP) -> str:
    """Pull cell values from an xlsx file, capped at max_rows per sheet."""
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        parts: list[str] = []
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            parts.append(f"\n[Sheet: {sheet_name}]")
            n = 0
            for row in ws.iter_rows(values_only=True):
                if n >= max_rows:
                    parts.append(f"  ... (truncated at {max_rows} rows)")
                    break
                cells = [str(c) if c is not None else "" for c in row]
                if any(cells):
                    parts.append(" | ".join(cells))
                n += 1
        wb.close()
        return "\n".join(parts)
    except Exception as e:
        return f"[Error reading {os.path.basename(path)}: {e}]"


def extract_pdf_text(path: str) -> str:
    """Pull text from a PDF using pymupdf (fitz)."""
    try:
        import fitz
        doc = fitz.open(path)
        parts: list[str] = []
        for page_num, page in enumerate(doc, 1):
            text = page.get_text()
            if text.strip():
                parts.append(f"[Page {page_num}]\n{text}")
        doc.close()
        return "\n\n".join(parts)
    except Exception as e:
        return f"[Error reading {os.path.basename(path)}: {e}]"


def _is_interstitial(path: str) -> bool:
    """PMC sometimes serves a tiny 'Preparing your download...' placeholder
    instead of the real supplement. Two flavors exist:
      1. HTML page misnamed .pdf (starts with <!DOCTYPE / <html)
      2. Real but tiny 1-page PDF whose only content is "Preparing to download"
    Both are useless for the model."""
    try:
        size = os.path.getsize(path)
    except Exception:
        return False
    if size >= INTERSTITIAL_BYTE_LIMIT:
        return False
    try:
        with open(path, "rb") as f:
            head = f.read(2048)
    except Exception:
        return False
    # html placeholder
    if head.startswith(b"<!DOCTYPE") or head.startswith(b"<html"):
        return True
    # tiny pdf placeholder: check the body bytes for the marker
    if head.startswith(b"%PDF-") and b"Preparing" in head:
        return True
    return False


def _looks_like_interstitial_text(text: str) -> bool:
    """A real supplement extracts to thousands of chars. If a PDF extracts
    to <300 chars AND mentions 'Preparing to download', it's a placeholder."""
    if not text:
        return True
    if len(text) < 300 and "Preparing to download" in text:
        return True
    return False


def load_supplements(
    pmcid: str,
    base_dir: Optional[Path] = None,
    char_cap: int = SUPPLEMENT_CHAR_CAP,
    verbose: bool = True,
) -> str:
    """Concatenate text from every supplement file we have for `pmcid`.

    Args:
        pmcid:    e.g. "PMC6390544"
        base_dir: project root. If None, assumes cwd has downloaded_supplements/
        char_cap: truncate concatenated text at this many chars
        verbose:  print which files were used / skipped

    Returns "(No supplementary files available)" if nothing is found.
    """
    if base_dir is None:
        base_dir = Path.cwd()
    suppl_path = Path(base_dir) / "downloaded_supplements" / pmcid

    if not suppl_path.exists():
        return "(No supplementary files available)"

    parts: list[str] = []
    skipped_interstitials: list[str] = []
    for fname in sorted(os.listdir(suppl_path)):
        fpath = str(suppl_path / fname)
        ext = os.path.splitext(fname)[1].lower()

        # skip pmc interstitial html pages disguised as .pdf
        if _is_interstitial(fpath):
            skipped_interstitials.append(fname)
            continue

        if ext in (".docx", ".doc"):
            body = extract_docx_text(fpath)
        elif ext in (".xlsx", ".xls"):
            body = extract_xlsx_text(fpath)
        elif ext == ".pdf":
            body = extract_pdf_text(fpath)
            # post-extract check: tiny "Preparing to download" placeholders
            if _looks_like_interstitial_text(body):
                skipped_interstitials.append(fname)
                continue
        elif ext in (".csv", ".txt"):
            try:
                with open(fpath, encoding="utf-8", errors="replace") as f:
                    body = f.read()[:50_000]
            except Exception as e:
                body = f"[Error reading {fname}: {e}]"
        else:
            body = f"[Skipped: unsupported format {ext}]"

        parts.append(f"\n--- Supplement file: {fname} ---")
        parts.append(body)

    if verbose and skipped_interstitials:
        print(f"  [load_supplements:{pmcid}] skipped {len(skipped_interstitials)} interstitial files: {skipped_interstitials}")

    text = "\n".join(parts)
    if len(text) > char_cap:
        text = text[:char_cap] + f"\n\n... (supplementary text truncated at {char_cap:,} chars)"
    return text or "(No supplementary files available)"
