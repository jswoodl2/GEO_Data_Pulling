"""
Download supplementary files for downloaded papers.

Three sequential passes per paper:
  1. PMC      — NCBI efetch + Playwright (bypasses PMC's proof-of-work JS challenge).
                Runs for any paper saved as PMC<num>.xml in downloaded_papers/.
  2. EuropePMC — supplementaryFiles endpoint (returns a zip). Tried for any paper
                 that has a PMCID but got nothing in pass 1, AND for PMCID-shaped
                 papers we couldn't pull via PMC.
  3. Elsevier — parses each Elsevier-saved article XML for <ce:e-component> /
                <xocs:attachment> / <ce:object-ref> supplement refs, then fetches
                each one via the Elsevier TDM API (uses ELSEVIER_API_KEY).

Outputs:
  downloaded_supplements/<PMCID-or-doi>/<filename>     — actual files
  downloaded_supplements/_progress.json                — resume state
  downloaded_supplements/_report.csv                   — per-file outcome incl.
                                                          short-flag for stub files

Short-flag: files smaller than SUPPLEMENT_MIN_USEFUL_BYTES are saved but flagged
so you can spot stubs (e.g. publisher wrapper pages) that downloaded but contain
no real metadata.

Usage:
    # activate venv first: source .venv/bin/activate
    python download_supplements.py                     # all three passes
    python download_supplements.py --skip-elsevier     # PMC + EPMC only
    python download_supplements.py --pmcid PMC10843761 # single paper
    python download_supplements.py --types xlsx,docx   # filter to these types
"""

import argparse
import csv
import io
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Optional, List, Dict, Tuple

from playwright.sync_api import sync_playwright


# ============================================================================
# CONFIGURATION
# ============================================================================
PAPERS_DIR = "downloaded_papers"
SUPPL_DIR = "downloaded_supplements"
PROGRESS_FILE = os.path.join(SUPPL_DIR, "_progress.json")
REPORT_FILE = os.path.join(SUPPL_DIR, "_report.csv")

# File extensions worth downloading (structured/text data + pdfs)
DEFAULT_TYPES = {".xlsx", ".xls", ".docx", ".doc", ".csv", ".txt", ".pdf"}

# Per-file size threshold below which we suspect the file is empty / wrapper / stub
SUPPLEMENT_MIN_USEFUL_BYTES = 2000

API_DELAY = 0.35  # NCBI rate limit: ~3 requests/sec

BASE_URL = "https://pmc.ncbi.nlm.nih.gov/articles/instance/{pmcid_num}/bin/{filename}"
EFETCH_URL = "https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pmc&id={pmcid_num}&rettype=xml"
EPMC_SUPP_URL = "https://www.ebi.ac.uk/europepmc/webservices/rest/PMC/{pmcid_num}/supplementaryFiles"

ELSEVIER_API_KEY = os.environ.get("ELSEVIER_API_KEY", "").strip()


# ============================================================================
# REPORTING (per-file outcomes flow into _report.csv at the end)
# ============================================================================
REPORT_ROWS: List[Dict] = []


def log_outcome(paper_key: str, source: str, filename: str, status: str,
                bytes_: int, short: bool, reason: str = "") -> None:
    REPORT_ROWS.append({
        "PaperKey": paper_key,
        "Source": source,
        "Filename": filename,
        "Status": status,
        "Bytes": bytes_,
        "ShortFlag": short,
        "Reason": reason,
    })


def validate_file(path: str, min_bytes: int = SUPPLEMENT_MIN_USEFUL_BYTES) -> Tuple[int, bool]:
    if not os.path.exists(path):
        return 0, True
    size = os.path.getsize(path)
    return size, size < min_bytes


def write_report() -> None:
    os.makedirs(SUPPL_DIR, exist_ok=True)
    with open(REPORT_FILE, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["PaperKey", "Source", "Filename", "Status", "Bytes", "ShortFlag", "Reason"],
        )
        writer.writeheader()
        writer.writerows(REPORT_ROWS)
    print(f"Report written: {REPORT_FILE} ({len(REPORT_ROWS)} rows)")


# ============================================================================
# PROGRESS (resume state)
# ============================================================================
def load_progress() -> dict:
    if os.path.exists(PROGRESS_FILE):
        with open(PROGRESS_FILE) as f:
            return json.load(f)
    return {"downloaded": {}, "failed": {}}


def save_progress(progress: dict) -> None:
    os.makedirs(SUPPL_DIR, exist_ok=True)
    with open(PROGRESS_FILE, "w") as f:
        json.dump(progress, f, indent=2)


def get_pmcid_number(pmcid: str) -> str:
    return pmcid.replace("PMC", "").strip()


def sanitize_paper_key(value: str) -> str:
    value = str(value or "").strip().replace(" ", "_")
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", value)


def load_download_summary() -> Dict[str, Dict[str, str]]:
    """Map downloaded file stems / PMCID / sanitized DOI to download-summary rows."""
    path = os.path.join(PAPERS_DIR, "download_summary.csv")
    out: Dict[str, Dict[str, str]] = {}
    if not os.path.exists(path):
        return out
    try:
        with open(path, encoding="utf-8") as f:
            for row in csv.DictReader(f):
                keys = []
                saved = row.get("saved_path") or ""
                if saved:
                    keys.append(os.path.splitext(os.path.basename(saved))[0])
                pmcid = (row.get("pmcid") or "").strip()
                doi = (row.get("doi") or "").strip()
                pmid = (row.get("pmid") or "").strip()
                if pmcid:
                    keys.append(pmcid)
                if doi:
                    keys.append(sanitize_paper_key(doi))
                if pmid:
                    keys.append(pmid)
                for key in keys:
                    out[key] = row
    except Exception as e:
        print(f"  warning: could not read paper download summary: {e}")
    return out


def europepmc_search_record(meta: Dict[str, str]) -> Optional[Dict[str, str]]:
    """Find a Europe PMC record from DOI/PMID/PMCID metadata."""
    queries = []
    doi = (meta.get("doi") or "").strip()
    pmid = (meta.get("pmid") or "").strip()
    pmcid = (meta.get("pmcid") or "").strip()
    if doi:
        queries.append(f'DOI:"{doi}"')
    if pmid:
        queries.append(f"EXT_ID:{pmid}")
    if pmcid:
        queries.append(f"EXT_ID:{pmcid}")
    for q in queries:
        url = "https://www.ebi.ac.uk/europepmc/webservices/rest/search?" + urllib.parse.urlencode({
            "query": q,
            "format": "json",
            "pageSize": 1,
        })
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=30) as r:
                data = json.loads(r.read().decode("utf-8", errors="replace"))
            results = (data.get("resultList") or {}).get("result") or []
            if results:
                return results[0]
        except Exception:
            continue
    return None


# ============================================================================
# PASS 1: PMC (NCBI efetch + Playwright PoW bypass)
# ============================================================================
def fetch_supplement_list_efetch(pmcid_num: str) -> List[dict]:
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


def _parse_supplements_from_xml_str(xml_str: str) -> List[dict]:
    """Parse supplement filenames from an XML string (efetch or local)."""
    supplements = []
    seen = set()
    for match in re.finditer(r'href="([^"]+\.\w{2,5})"', xml_str):
        href = match.group(1)
        if "/" in href or href in seen:
            continue
        seen.add(href)
        ext = os.path.splitext(href)[1].lower()
        if ext:
            supplements.append({"filename": href, "ext": ext})
    return supplements


def download_file_playwright(page, context, pmcid_num: str, filename: str, out_path: str) -> bool:
    """Download a single PMC supplement. PDFs render inline so we fetch them via
    the browser context's request API (reusing PoW cookies)."""
    url = BASE_URL.format(pmcid_num=pmcid_num, filename=filename)
    ext = os.path.splitext(filename)[1].lower()

    # PDFs are served inline by chromium — use request API instead of download event
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


def run_pmc_pass(papers: List[str], progress: dict, allowed_types: Optional[set]) -> None:
    print("\n" + "=" * 70)
    print("PASS 1: PMC supplementary files (Playwright)")
    print("=" * 70)

    plan = []
    for pmcid in papers:
        if not pmcid.startswith("PMC"):
            continue
        pmcid_num = get_pmcid_number(pmcid)
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

    print(f"  {len(plan)} PMC supplement files queued")
    if not plan:
        return

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        # Solve PoW once — cookies carry over
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
                size, short = validate_file(out_path)
                if short:
                    print(f"OK ({size:,} bytes) — FLAGGED SHORT")
                    log_outcome(pmcid, "PMC", filename, "ok_short", size, True, "below threshold")
                else:
                    print(f"OK ({size:,} bytes)")
                    log_outcome(pmcid, "PMC", filename, "ok", size, False)
                progress["downloaded"][key] = {"size": size, "path": out_path}
            else:
                print("FAILED")
                log_outcome(pmcid, "PMC", filename, "failed", 0, True, "download error")
                progress["failed"][key] = {
                    "attempts": progress.get("failed", {}).get(key, {}).get("attempts", 0) + 1
                }

            if (i + 1) % 10 == 0:
                save_progress(progress)
            time.sleep(1)

        browser.close()


# ============================================================================
# PASS 2: Europe PMC supplementary files endpoint
# ============================================================================
def run_europepmc_pass(papers: List[str], progress: dict, allowed_types: Optional[set]) -> None:
    print("\n" + "=" * 70)
    print("PASS 2: Europe PMC supplementary files")
    print("=" * 70)

    metadata = load_download_summary()
    fetched_for = 0
    for paper_key in papers:
        # Still try Europe PMC even if PMC found files; it can expose additional supplements.
        url = None
        request_key = paper_key
        if paper_key.startswith("PMC"):
            pmcid_num = get_pmcid_number(paper_key)
            url = EPMC_SUPP_URL.format(pmcid_num=pmcid_num)
        else:
            meta = metadata.get(paper_key, {})
            record = europepmc_search_record(meta)
            if record:
                source = (record.get("source") or "").strip()
                ext_id = (record.get("id") or record.get("pmid") or "").strip()
                pmcid = (record.get("pmcid") or "").strip()
                if pmcid:
                    url = EPMC_SUPP_URL.format(pmcid_num=get_pmcid_number(pmcid))
                elif source and ext_id:
                    url = f"https://www.ebi.ac.uk/europepmc/webservices/rest/{urllib.parse.quote(source)}/{urllib.parse.quote(ext_id)}/supplementaryFiles"
            if not url:
                continue

        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
            with urllib.request.urlopen(req, timeout=60) as r:
                data = r.read()
        except Exception as e:
            print(f"  {request_key}: EPMC fetch failed: {e}")
            log_outcome(request_key, "EuropePMC", "(zip)", "failed", 0, True, str(e))
            time.sleep(API_DELAY)
            continue

        if not data or len(data) < 200:
            time.sleep(API_DELAY)
            continue

        paper_dir = os.path.join(SUPPL_DIR, request_key)
        os.makedirs(paper_dir, exist_ok=True)
        saved = 0
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                for name in zf.namelist():
                    if name.endswith("/"):
                        continue
                    ext = os.path.splitext(name)[1].lower()
                    if allowed_types and ext and ext not in allowed_types:
                        continue
                    filename = os.path.basename(name)
                    out_path = os.path.join(paper_dir, filename)
                    with open(out_path, "wb") as f:
                        f.write(zf.read(name))
                    size, short = validate_file(out_path)
                    status = "ok_short" if short else "ok"
                    reason = "below threshold" if short else ""
                    log_outcome(request_key, "EuropePMC", filename, status, size, short, reason)
                    progress["downloaded"][f"{request_key}/{filename}"] = {"size": size, "path": out_path}
                    saved += 1
            if saved > 0:
                fetched_for += 1
                print(f"  {request_key}: saved {saved} EPMC supplements")
        except zipfile.BadZipFile:
            print(f"  {request_key}: EPMC response not a valid zip ({len(data)} bytes)")
            log_outcome(request_key, "EuropePMC", "(zip)", "failed", len(data), True, "not a zip")

        time.sleep(API_DELAY)

    print(f"  EPMC pass: supplements fetched for {fetched_for} paper(s)")


# ============================================================================
# PASS 3: Elsevier (parse downloaded Elsevier article XMLs)
# ============================================================================
def parse_elsevier_supplement_refs(xml_path: str) -> List[Dict[str, str]]:
    """Find supplement file references in an Elsevier article XML.

    Defensive parse — Elsevier puts supplement refs in several places depending on
    journal:
      - <ce:e-component> with @xlink:href or child <ce:link>
      - <xocs:attachment> with @filename, @ref-type
      - <ce:object-ref> with @xlink:href
    Returns [{href, filename}] (either or both may be empty).
    """
    refs: List[Dict[str, str]] = []
    try:
        tree = ET.parse(xml_path)
        root = tree.getroot()
    except Exception:
        return refs

    seen = set()
    for elem in root.iter():
        tag = elem.tag.split("}")[-1]
        if tag not in ("e-component", "attachment", "object-ref", "object"):
            continue

        attrs = elem.attrib
        href = (
            attrs.get("{http://www.w3.org/1999/xlink}href")
            or attrs.get("href")
            or attrs.get("eid")
            or attrs.get("refid")
            or ""
        )
        filename = attrs.get("filename") or attrs.get("name") or ""

        # Look through children for filename / link tags
        for child in elem.iter():
            ctag = child.tag.split("}")[-1]
            if ctag in ("filename", "label", "caption") and child.text:
                t = child.text.strip()
                if not filename and "." in t and len(t) < 200:
                    filename = t
            if ctag == "link":
                lh = child.attrib.get("{http://www.w3.org/1999/xlink}href")
                if lh and not href:
                    href = lh

        if not href and not filename:
            continue
        key = (href, filename)
        if key in seen:
            continue
        seen.add(key)
        refs.append({"href": href, "filename": filename})

    return refs


def run_elsevier_pass(papers: List[str], progress: dict, allowed_types: Optional[set]) -> None:
    print("\n" + "=" * 70)
    print("PASS 3: Elsevier supplementary files (TDM API)")
    print("=" * 70)

    if not ELSEVIER_API_KEY:
        print("  ELSEVIER_API_KEY not set in environment; skipping pass.")
        return

    # Identify Elsevier-saved XMLs using download_summary.csv when available.
    if not os.path.isdir(PAPERS_DIR):
        print(f"  {PAPERS_DIR}/ not found; nothing to scan.")
        return
    metadata = load_download_summary()
    elsevier_xmls = []
    for fname in sorted(os.listdir(PAPERS_DIR)):
        if not fname.endswith(".xml"):
            continue
        base = fname[:-4]
        if base.startswith("PMC"):
            continue
        meta = metadata.get(base, {})
        if metadata and (meta.get("source") or "").lower() != "elsevier":
            continue
        elsevier_xmls.append(fname)

    if not elsevier_xmls:
        print("  No Elsevier-saved XMLs detected.")
        return

    print(f"  Scanning {len(elsevier_xmls)} Elsevier XML(s) for supplement refs...")

    for fname in elsevier_xmls:
        xml_path = os.path.join(PAPERS_DIR, fname)
        doi_key = fname[:-4]  # sanitized DOI used as filename
        refs = parse_elsevier_supplement_refs(xml_path)

        if not refs:
            print(f"  {doi_key}: no supplement refs in XML")
            continue

        paper_dir = os.path.join(SUPPL_DIR, doi_key)
        os.makedirs(paper_dir, exist_ok=True)
        print(f"  {doi_key}: {len(refs)} supplement ref(s)")

        for ref in refs:
            href = (ref.get("href") or "").strip()
            filename = (ref.get("filename") or "").strip()
            if not filename and href:
                filename = href.rsplit("/", 1)[-1]
            if not filename:
                filename = f"elsevier_supp_{abs(hash(href)) % 10**8}.bin"
            ext = os.path.splitext(filename)[1].lower()
            if allowed_types and ext and ext not in allowed_types:
                continue

            # Resolve URL
            if href.startswith("http"):
                url = href
            elif href:
                url = f"https://api.elsevier.com/content/object/eid/{urllib.parse.quote(href)}"
            else:
                log_outcome(doi_key, "Elsevier", filename, "failed", 0, True, "no href / eid")
                continue

            out_path = os.path.join(paper_dir, filename)
            try:
                req = urllib.request.Request(
                    url,
                    headers={
                        "X-ELS-APIKey": ELSEVIER_API_KEY,
                        "Accept": "*/*",
                        "User-Agent": "Mozilla/5.0",
                    },
                )
                with urllib.request.urlopen(req, timeout=60) as r:
                    data = r.read()
            except Exception as e:
                print(f"    {filename}: Elsevier fetch failed: {e}")
                log_outcome(doi_key, "Elsevier", filename, "failed", 0, True, str(e))
                time.sleep(API_DELAY)
                continue

            if not data:
                log_outcome(doi_key, "Elsevier", filename, "failed", 0, True, "empty response")
                continue

            with open(out_path, "wb") as f:
                f.write(data)
            size, short = validate_file(out_path)
            status = "ok_short" if short else "ok"
            reason = "below threshold" if short else ""
            flag_str = " — FLAGGED SHORT" if short else ""
            print(f"    {filename}: {size:,} bytes{flag_str}")
            log_outcome(doi_key, "Elsevier", filename, status, size, short, reason)
            progress["downloaded"][f"{doi_key}/{filename}"] = {"size": size, "path": out_path}
            time.sleep(API_DELAY)


# ============================================================================
# MAIN
# ============================================================================
def main():
    parser = argparse.ArgumentParser(description="Download supplements (PMC + EPMC + Elsevier)")
    parser.add_argument("--limit", type=int, default=0, help="Limit number of papers")
    parser.add_argument("--pmcid", type=str, default="", help="Single PMCID to process")
    parser.add_argument(
        "--types", type=str, default="",
        help="Comma-separated extensions (e.g., xlsx,docx). Default: xlsx,xls,docx,doc,csv,txt,pdf",
    )
    parser.add_argument("--all-types", action="store_true", help="Download all file types")
    parser.add_argument("--skip-pmc", action="store_true", help="Skip PMC pass")
    parser.add_argument("--skip-europepmc", action="store_true", help="Skip Europe PMC pass")
    parser.add_argument("--skip-elsevier", action="store_true", help="Skip Elsevier pass")
    args = parser.parse_args()

    allowed_types: Optional[set]
    if args.all_types:
        allowed_types = None
    elif args.types:
        allowed_types = {"." + t.strip(".") for t in args.types.split(",")}
    else:
        allowed_types = DEFAULT_TYPES

    os.makedirs(SUPPL_DIR, exist_ok=True)
    progress = load_progress()

    if not os.path.isdir(PAPERS_DIR):
        sys.exit(f"Papers directory not found: {PAPERS_DIR}. Run download_papers_copy.py first.")

    # Discover papers: file stems of XML/HTML/PDF in downloaded_papers/
    if args.pmcid:
        papers = [args.pmcid]
    else:
        stems = set()
        for f in os.listdir(PAPERS_DIR):
            if f.endswith((".xml", ".html", ".pdf")):
                stems.add(os.path.splitext(f)[0])
        papers = sorted(stems)

    if args.limit:
        papers = papers[: args.limit]

    print(f"Papers to scan: {len(papers)}")
    print(f"Allowed types: {sorted(allowed_types) if allowed_types else 'ALL'}")
    print(f"Elsevier API key set: {bool(ELSEVIER_API_KEY)}")

    if not args.skip_pmc:
        run_pmc_pass(papers, progress, allowed_types)
        save_progress(progress)

    if not args.skip_europepmc:
        run_europepmc_pass(papers, progress, allowed_types)
        save_progress(progress)

    if not args.skip_elsevier:
        run_elsevier_pass(papers, progress, allowed_types)
        save_progress(progress)

    save_progress(progress)
    write_report()

    dl_count = len(progress["downloaded"])
    fail_count = len(progress["failed"])
    short_count = sum(1 for r in REPORT_ROWS if r["ShortFlag"] and r["Status"].startswith("ok"))
    print("\n" + "=" * 70)
    print("FINAL SUMMARY")
    print("=" * 70)
    print(f"  Files downloaded:     {dl_count}")
    print(f"  Files failed:         {fail_count}")
    print(f"  Files flagged SHORT:  {short_count}  (saved but suspiciously small)")
    print(f"  Report:               {REPORT_FILE}")


if __name__ == "__main__":
    main()
