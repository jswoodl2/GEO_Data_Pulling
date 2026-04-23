"""
Download supplementary files from PMC for all downloaded papers.
Uses Playwright to bypass PMC's proof-of-work JS challenge.

Usage:
    # activate venv first: source .venv/bin/activate
    python download_supplements.py                    # download all
    python download_supplements.py --limit 10         # first 10 papers
    python download_supplements.py --pmcid PMC10843761  # single paper
    python download_supplements.py --types xlsx,docx  # only these file types
"""

import os
import re
import sys
import json
import time
import argparse
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
from playwright.sync_api import sync_playwright


PAPERS_DIR = "downloaded_papers"
SUPPL_DIR = "downloaded_supplements"
PROGRESS_FILE = os.path.join(SUPPL_DIR, "_progress.json")

# file types worth downloading (structured/text data + pdfs)
DEFAULT_TYPES = {".xlsx", ".xls", ".docx", ".doc", ".csv", ".txt", ".pdf"}

BASE_URL = "https://pmc.ncbi.nlm.nih.gov/articles/instance/{pmcid_num}/bin/{filename}"
EFETCH_URL = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pmc&id={pmcid_num}&rettype=xml"
PMC_ARTICLE_URL = "https://pmc.ncbi.nlm.nih.gov/articles/instance/{pmcid_num}/bin/"

API_DELAY = 0.35  # NCBI rate limit: ~3 requests/sec


def fetch_supplement_list_efetch(pmcid_num: str) -> list[dict]:
    """Fetch supplement file list from NCBI efetch API."""
    url = EFETCH_URL.format(pmcid_num=pmcid_num)
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30) as r:
            xml_str = r.read().decode("utf-8", errors="replace")
    except Exception as e:
        print(f"    efetch failed: {e}")
        return []

    return _parse_supplements_from_xml_str(xml_str)


def fetch_supplement_list_html(page, pmcid_num: str) -> list[dict]:
    """Scrape supplement file links from the PMC article HTML page."""
    url = f"https://pmc.ncbi.nlm.nih.gov/articles/PMC{pmcid_num}/"
    try:
        page.goto(url, wait_until="networkidle", timeout=30000)
        html = page.content()
    except Exception as e:
        print(f"    HTML scrape failed: {e}")
        return []

    supplements = []
    seen = set()
    for match in re.finditer(
        rf'/articles/instance/{pmcid_num}/bin/([^\s"\'<>]+)', html
    ):
        filename = match.group(1)
        if filename in seen:
            continue
        seen.add(filename)
        ext = os.path.splitext(filename)[1].lower()
        if ext:
            supplements.append({"filename": filename, "ext": ext})

    return supplements


def _parse_supplements_from_xml_str(xml_str: str) -> list[dict]:
    """Parse supplement filenames from an XML string (efetch or local)."""
    supplements = []
    seen = set()

    # find media href attributes
    for match in re.finditer(r'href="([^"]+\.\w{2,5})"', xml_str):
        href = match.group(1)
        # skip non-file refs (e.g., urls to other papers)
        if "/" in href or href in seen:
            continue
        seen.add(href)
        ext = os.path.splitext(href)[1].lower()
        if ext:
            supplements.append({"filename": href, "ext": ext})

    return supplements


def parse_supplements_from_local_xml(xml_path: str) -> list[dict]:
    """Extract supplement file references from a local PMC XML file."""
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
        xml_str = ET.tostring(root, encoding="unicode")
    except ET.ParseError:
        return []

    return _parse_supplements_from_xml_str(xml_str)


def get_pmcid_number(pmcid: str) -> str:
    """Extract numeric part from PMCID (e.g., PMC10843761 -> 10843761)."""
    return pmcid.replace("PMC", "").strip()


def load_progress() -> dict:
    """Load download progress to allow resuming."""
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE) as f:
            return json.load(f)
    return {"downloaded": {}, "failed": {}}


def save_progress(progress: dict):
    with open(PROGRESS_FILE, "w") as f:
        json.dump(progress, f, indent=2)


def download_file_playwright(page, context, pmcid_num: str, filename: str, out_path: str) -> bool:
    """Download a single supplement file. PDFs render inline in the browser,
    so we fetch them via the browser context's request API (reusing PoW cookies)."""
    url = BASE_URL.format(pmcid_num=pmcid_num, filename=filename)
    ext = os.path.splitext(filename)[1].lower()

    # pdfs are served inline by chromium — use request api instead of download event
    if ext == ".pdf":
        try:
            resp = context.request.get(url, timeout=60000)
            if not resp.ok:
                print(f"      HTTP {resp.status}")
                return False
            body = resp.body()
            if body[:20].startswith(b"<!DOCTYPE") or body[:20].startswith(b"<html"):
                return False
            with open(out_path, "wb") as f:
                f.write(body)
            return True
        except Exception as e:
            print(f"      Error: {e}")
            return False

    try:
        with page.expect_download(timeout=45000) as dl_info:
            page.goto(url, wait_until="commit", timeout=30000)
        download = dl_info.value
        download.save_as(out_path)

        with open(out_path, "rb") as f:
            header = f.read(20)
        if header.startswith(b"<!DOCTYPE") or header.startswith(b"<html"):
            os.remove(out_path)
            return False
        return True
    except Exception as e:
        try:
            time.sleep(5)
            with page.expect_download(timeout=30000) as dl_info:
                pass
            download = dl_info.value
            download.save_as(out_path)
            with open(out_path, "rb") as f:
                header = f.read(20)
            if header.startswith(b"<!DOCTYPE") or header.startswith(b"<html"):
                os.remove(out_path)
                return False
            return True
        except Exception:
            pass
        print(f"      Error: {e}")
        return False


def main():
    parser = argparse.ArgumentParser(description="Download PMC supplement files")
    parser.add_argument("--limit", type=int, default=0, help="Limit number of papers")
    parser.add_argument("--pmcid", type=str, default="", help="Single PMCID to process")
    parser.add_argument(
        "--types", type=str, default="",
        help="Comma-separated file extensions (e.g., xlsx,docx). Default: xlsx,xls,docx,doc,csv,txt"
    )
    parser.add_argument("--all-types", action="store_true", help="Download all file types including pdf/pptx/zip")
    args = parser.parse_args()

    allowed_types = DEFAULT_TYPES
    if args.all_types:
        allowed_types = None  # no filter
    elif args.types:
        allowed_types = {"." + t.strip(".") for t in args.types.split(",")}

    os.makedirs(SUPPL_DIR, exist_ok=True)
    progress = load_progress()

    # gather papers to process
    if args.pmcid:
        xml_files = [f"{args.pmcid}.xml"]
    else:
        xml_files = sorted(f for f in os.listdir(PAPERS_DIR) if f.endswith(".xml"))

    if args.limit:
        xml_files = xml_files[: args.limit]

    # build download plan — use efetch api to discover supplement files
    print("Discovering supplement files via NCBI efetch API...")
    plan = []
    for idx, xml_file in enumerate(xml_files):
        pmcid = xml_file.replace(".xml", "")
        pmcid_num = get_pmcid_number(pmcid)

        # always use efetch — local xmls often miss pdf references
        supplements = fetch_supplement_list_efetch(pmcid_num)
        time.sleep(API_DELAY)

        if not supplements:
            continue

        for supp in supplements:
            if allowed_types and supp["ext"] not in allowed_types:
                continue
            key = f"{pmcid}/{supp['filename']}"
            if key in progress["downloaded"]:
                continue
            plan.append((pmcid, pmcid_num, supp["filename"], supp["ext"]))

        if (idx + 1) % 50 == 0:
            print(f"  Scanned {idx+1}/{len(xml_files)} papers, {len(plan)} files queued...")

    print(f"Papers: {len(xml_files)}, Files to download: {len(plan)}")
    if not plan:
        print("Nothing to download.")
        return

    # download with playwright
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        # solve pow once by visiting an article page — cookies carry over for pdf fetches
        if plan:
            first_pmcid_num = plan[0][1]
            try:
                page.goto(f"https://pmc.ncbi.nlm.nih.gov/articles/PMC{first_pmcid_num}/",
                          wait_until="networkidle", timeout=45000)
            except Exception:
                pass

        for i, (pmcid, pmcid_num, filename, ext) in enumerate(plan):
            key = f"{pmcid}/{filename}"
            paper_dir = os.path.join(SUPPL_DIR, pmcid)
            os.makedirs(paper_dir, exist_ok=True)
            out_path = os.path.join(paper_dir, filename)

            print(f"  [{i+1}/{len(plan)}] {pmcid} / {filename} ...", end=" ", flush=True)

            ok = download_file_playwright(page, context, pmcid_num, filename, out_path)
            if ok:
                size = os.path.getsize(out_path)
                print(f"OK ({size:,} bytes)")
                progress["downloaded"][key] = {"size": size, "path": out_path}
            else:
                print("FAILED")
                progress["failed"][key] = {"attempts": progress.get("failed", {}).get(key, {}).get("attempts", 0) + 1}

            # save progress every 10 files
            if (i + 1) % 10 == 0:
                save_progress(progress)

            # be polite to pmc
            time.sleep(1)

        browser.close()

    save_progress(progress)
    dl_count = len(progress["downloaded"])
    fail_count = len(progress["failed"])
    print(f"\nDone. Downloaded: {dl_count}, Failed: {fail_count}")


if __name__ == "__main__":
    main()
