#!/usr/bin/env python3
"""
Extract text from downloaded PMC XML files and split into overlapping chunks.

This is the script version of the XML chunking section from download_papers.ipynb.
It reads XML files from a folder (default: downloaded_papers) and writes a JSON
file containing, for each paper:
    {
        "pmcid": "PMCxxxxxx",
        "chunks": ["....", "...", ...]
    }
"""

import json
import os
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Optional

# --- CONFIGURATION ---
# Folder with the downloaded PMC XML files
INPUT_DIR = Path("downloaded_papers")

# Output JSON file with all chunked texts
OUTPUT_JSON_FILE = Path("processed_papers.json")

# Settings for how to chunk the text
CHUNK_SIZE = 5000  # characters per chunk (roughly 1000-1200 words)
OVERLAP_SIZE = 500  # characters of overlap between chunks


def extract_text_from_xml(xml_file_path: Path) -> Optional[str]:
    """Parse an XML file and extract all human-readable text as a single string."""
    try:
        tree = ET.parse(xml_file_path)
        root = tree.getroot()
        all_text = " ".join(node.text for node in root.iter() if node.text)
        cleaned_text = " ".join(all_text.split())
        return cleaned_text
    except ET.ParseError as e:
        print(f"    > Could not parse XML file {xml_file_path.name}. Error: {e}")
        return None
    except Exception as e:
        print(f"    > Unexpected error with file {xml_file_path.name}: {e}")
        return None


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


def main() -> None:
    all_papers_data = []

    if not INPUT_DIR.exists():
        raise SystemExit(f"Input directory not found: {INPUT_DIR}")

    print(f"Starting to process files from the '{INPUT_DIR}' directory...")
    xml_files = sorted(f for f in INPUT_DIR.iterdir() if f.suffix.lower() == ".xml")

    for index, xml_path in enumerate(xml_files, start=1):
        print(f"\nProcessing file {index}/{len(xml_files)}: {xml_path.name}...")

        full_text = extract_text_from_xml(xml_path)

        if full_text:
            text_chunks = split_text_into_chunks(full_text)
            print(f"  > Extracted text and created {len(text_chunks)} chunks.")

            pmcid = xml_path.stem
            all_papers_data.append({"pmcid": pmcid, "chunks": text_chunks})

    OUTPUT_JSON_FILE.write_text(
        json.dumps(all_papers_data, indent=4),
        encoding="utf-8",
    )

    print("\n--- Process Complete ---")
    print(f"Successfully processed {len(all_papers_data)} papers.")
    print(f"All chunked text has been saved to '{OUTPUT_JSON_FILE}'.")


if __name__ == "__main__":
    main()

