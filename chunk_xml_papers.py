#!/usr/bin/env python3
"""
Extract text from downloaded PMC papers and split into overlapping chunks.

Supports XML, HTML, and PDF files in the input directory.
For XML files that are truncated (publisher restriction), automatically
scrapes the full text from the PMC HTML page via Playwright.

Output: processed_papers.json with structure:
    [{"pmcid": "PMCxxxxxx", "chunks": ["...", ...]}, ...]

Usage:
    source .venv/bin/activate
    python chunk_xml_papers.py
"""

import json
import os
import re
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Optional

# config
INPUT_DIR = Path("downloaded_papers")
OUTPUT_JSON_FILE = Path("processed_papers.json")

CHUNK_SIZE = 5000       # characters per chunk
OVERLAP_SIZE = 500      # character overlap between consecutive chunks

# if a paper has less than this many chars or is missing key section
# headings, we treat it as truncated and try scraping the html version
MIN_FULLTEXT_CHARS = 5000
FULLTEXT_KEYWORDS = {"methods", "results", "discussion", "materials and methods"}


# text extraction functions

def extract_text_from_xml(xml_file_path: Path) -> Optional[str]:
    """Parse a PMC XML file and pull out all readable text."""
    try:
        tree = ET.parse(xml_file_path)
        root = tree.getroot()
        all_text = " ".join(node.text for node in root.iter() if node.text)
        return " ".join(all_text.split())
    except ET.ParseError as e:
        print(f"    > Could not parse XML {xml_file_path.name}: {e}")
        return None
    except Exception as e:
        print(f"    > Error with {xml_file_path.name}: {e}")
        return None


def extract_text_from_html_file(html_path: Path) -> Optional[str]:
    """Extract article text from a local HTML file using BeautifulSoup."""
    try:
        from bs4 import BeautifulSoup
        html = html_path.read_text(encoding="utf-8", errors="replace")
        soup = BeautifulSoup(html, "html.parser")
        article = soup.find("article") or soup.find("main") or soup
        for tag in article.find_all(["script", "style", "nav", "header", "footer"]):
            tag.decompose()
        text = article.get_text(" ", strip=True)
        return " ".join(text.split())
    except Exception as e:
        print(f"    > Error reading HTML {html_path.name}: {e}")
        return None


def extract_text_from_pdf(pdf_path: Path) -> Optional[str]:
    """Extract text from a PDF using pymupdf (fitz)."""
    try:
        import fitz
        doc = fitz.open(str(pdf_path))
        parts = []
        for page in doc:
            text = page.get_text()
            if text.strip():
                parts.append(text)
        doc.close()
        return " ".join(" ".join(parts).split())
    except Exception as e:
        print(f"    > Error reading PDF {pdf_path.name}: {e}")
        return None


def scrape_pmc_html(pmcid: str) -> Optional[str]:
    """Scrape full article text from the PMC website using playwright."""
    try:
        from playwright.sync_api import sync_playwright
        from bs4 import BeautifulSoup
    except ImportError:
        print("    > playwright or bs4 not installed, skipping HTML scrape")
        return None

    url = f"https://pmc.ncbi.nlm.nih.gov/articles/{pmcid}/"
    try:
        with sync_playwright() as p:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context()
            page = context.new_page()
            page.goto(url, wait_until="networkidle", timeout=60000)
            html = page.content()
            browser.close()

        soup = BeautifulSoup(html, "html.parser")
        article = soup.find("article") or soup.find("main") or soup
        for tag in article.find_all(["script", "style", "nav", "header", "footer"]):
            tag.decompose()
        text = article.get_text(" ", strip=True)
        return " ".join(text.split())
    except Exception as e:
        print(f"    > HTML scrape failed for {pmcid}: {e}")
        return None


def is_truncated(text: Optional[str]) -> bool:
    """Check if extracted text is too short or missing key sections."""
    if not text:
        return True
    if len(text) < MIN_FULLTEXT_CHARS:
        return True
    # need at least 2 section keywords to count as full text
    text_lower = text.lower()
    keyword_hits = sum(1 for kw in FULLTEXT_KEYWORDS if kw in text_lower)
    if keyword_hits < 2:
        return True
    return False


# chunking

def split_text_into_chunks(text: str) -> List[str]:
    """Split a long text into overlapping chunks."""
    if not text:
        return []
    chunks: List[str] = []
    start = 0
    n = len(text)
    while start < n:
        end = start + CHUNK_SIZE
        chunks.append(text[start:end])
        start += CHUNK_SIZE - OVERLAP_SIZE
    return chunks


# main

def main() -> None:
    if not INPUT_DIR.exists():
        raise SystemExit(f"Input directory not found: {INPUT_DIR}")

    print(f"Processing files from '{INPUT_DIR}'...")

    # gather all paper files, grouped by pmcid
    # prefer xml over html over pdf since xml parses fastest
    paper_files: dict[str, Path] = {}
    for f in sorted(INPUT_DIR.iterdir()):
        if f.suffix.lower() in (".xml", ".html", ".pdf"):
            pmcid = f.stem
            if pmcid not in paper_files or f.suffix.lower() == ".xml":
                paper_files[pmcid] = f

    total = len(paper_files)
    print(f"Found {total} papers to process.\n")

    all_papers_data = []
    stats = {"xml_ok": 0, "html_fallback": 0, "html_file": 0, "pdf_file": 0, "failed": 0}

    # collect truncated papers to scrape in bulk later
    truncated_pmcids = []

    # first pass: try extracting from local files
    local_results: dict[str, Optional[str]] = {}
    for index, (pmcid, fpath) in enumerate(sorted(paper_files.items()), 1):
        ext = fpath.suffix.lower()
        print(f"[{index}/{total}] {fpath.name}", end="")

        if ext == ".xml":
            text = extract_text_from_xml(fpath)
            if is_truncated(text):
                if pmcid.startswith("PMC"):
                    print(f"  -> TRUNCATED ({len(text or ''):,} chars), will scrape HTML")
                    truncated_pmcids.append(pmcid)
                    local_results[pmcid] = None  # placeholder
                elif text:
                    print(f"  -> SHORT/NON-PMC ({len(text):,} chars), keeping extracted text")
                    local_results[pmcid] = text
                    stats["xml_ok"] += 1
                else:
                    print("  -> no extractable XML text")
                    stats["failed"] += 1
            else:
                print(f"  -> OK ({len(text):,} chars)")
                local_results[pmcid] = text
                stats["xml_ok"] += 1
        elif ext == ".html":
            text = extract_text_from_html_file(fpath)
            if text and len(text) > MIN_FULLTEXT_CHARS:
                print(f"  -> OK HTML ({len(text):,} chars)")
                local_results[pmcid] = text
                stats["html_file"] += 1
            else:
                if pmcid.startswith("PMC"):
                    print(f"  -> HTML too short ({len(text or ''):,} chars), will scrape PMC")
                    truncated_pmcids.append(pmcid)
                    local_results[pmcid] = None
                elif text:
                    print(f"  -> SHORT/NON-PMC HTML ({len(text):,} chars), keeping extracted text")
                    local_results[pmcid] = text
                    stats["html_file"] += 1
                else:
                    print("  -> HTML extraction failed")
                    stats["failed"] += 1
        elif ext == ".pdf":
            text = extract_text_from_pdf(fpath)
            if text and len(text) > MIN_FULLTEXT_CHARS:
                print(f"  -> OK PDF ({len(text):,} chars)")
                local_results[pmcid] = text
                stats["pdf_file"] += 1
            else:
                if pmcid.startswith("PMC"):
                    print(f"  -> PDF too short ({len(text or ''):,} chars), will scrape PMC")
                    truncated_pmcids.append(pmcid)
                    local_results[pmcid] = None
                elif text:
                    print(f"  -> SHORT/NON-PMC PDF ({len(text):,} chars), keeping extracted text")
                    local_results[pmcid] = text
                    stats["pdf_file"] += 1
                else:
                    print("  -> PDF extraction failed")
                    stats["failed"] += 1

    # second pass: scrape truncated papers from pmc html using playwright
    if truncated_pmcids:
        print(f"\n--- Scraping {len(truncated_pmcids)} truncated papers from PMC HTML ---")
        try:
            from playwright.sync_api import sync_playwright
            from bs4 import BeautifulSoup
            import time

            with sync_playwright() as p:
                browser = p.chromium.launch(headless=True)
                context = browser.new_context()
                page = context.new_page()

                for i, pmcid in enumerate(truncated_pmcids, 1):
                    print(f"  [{i}/{len(truncated_pmcids)}] Scraping {pmcid}...", end="", flush=True)
                    url = f"https://pmc.ncbi.nlm.nih.gov/articles/{pmcid}/"
                    try:
                        page.goto(url, wait_until="networkidle", timeout=60000)
                        html = page.content()
                        soup = BeautifulSoup(html, "html.parser")
                        article = soup.find("article") or soup.find("main") or soup
                        for tag in article.find_all(["script", "style", "nav", "header", "footer"]):
                            tag.decompose()
                        text = article.get_text(" ", strip=True)
                        text = " ".join(text.split())

                        if len(text) > MIN_FULLTEXT_CHARS:
                            print(f" OK ({len(text):,} chars)")
                            local_results[pmcid] = text
                            stats["html_fallback"] += 1
                        else:
                            print(f" still short ({len(text):,} chars)")
                            stats["failed"] += 1
                    except Exception as e:
                        print(f" ERROR: {e}")
                        stats["failed"] += 1
                    time.sleep(1)

                browser.close()
        except ImportError:
            print("  playwright/bs4 not installed — cannot scrape. Install with:")
            print("    pip install playwright beautifulsoup4 && playwright install chromium")
            stats["failed"] += len(truncated_pmcids)

    # build final output json with chunks for each paper
    for pmcid in sorted(paper_files.keys()):
        text = local_results.get(pmcid)
        if text:
            chunks = split_text_into_chunks(text)
            all_papers_data.append({"pmcid": pmcid, "chunks": chunks})

    OUTPUT_JSON_FILE.write_text(
        json.dumps(all_papers_data, indent=4),
        encoding="utf-8",
    )

    print(f"\n{'='*60}")
    print(f"DONE — {len(all_papers_data)} papers processed")
    print(f"{'='*60}")
    print(f"  XML (full text):      {stats['xml_ok']}")
    print(f"  HTML fallback (PMC):  {stats['html_fallback']}")
    print(f"  HTML file:            {stats['html_file']}")
    print(f"  PDF file:             {stats['pdf_file']}")
    print(f"  Failed:               {stats['failed']}")
    print(f"\nSaved to: {OUTPUT_JSON_FILE}")


if __name__ == "__main__":
    main()
