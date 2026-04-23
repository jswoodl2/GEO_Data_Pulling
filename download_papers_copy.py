"""
SIMPLE OPEN ACCESS PAPER DOWNLOADER (PMC / EUROPE PMC / UNPAYWALL ONLY)

Only uses legitimate, free open access sources:
- PubMed Central (PMC)
- Europe PMC
- Unpaywall

INSTALLATION:
pip install requests pandas biopython openpyxl
"""

import os
import re
import time
import json
import pathlib
import pandas as pd
import requests
from typing import Optional, Dict, Any, Tuple, List
from Bio import Entrez


# configuration

NCBI_EMAIL = "jjosep31@asu.edu"         # email for NCBI Entrez & PMC idconv
UNPAYWALL_EMAIL = "jjosep31@asu.edu"    # required email for Unpaywall API
EXCEL_FILE = os.path.expanduser("~/Desktop/placenta_sheet.xlsx")
OUTPUT_DIR = "downloaded_papers"

# networking & retry
API_DELAY = 0.3
MAX_RETRIES = 5
BASE_BACKOFF = 0.75
TIMEOUT = 30

# behavior
USE_PMID_TO_PMCID = True

# content validation thresholds (to avoid saving stubs/tocs)
MIN_HTML_BYTES = 2000
MIN_PTAG_COUNT = 20
MIN_WORDS = 1500

Entrez.email = NCBI_EMAIL


# http session & retry logic

SESSION = requests.Session()
SESSION.headers.update({
    # some repositories are picky about ua/accept headers
    "User-Agent": f"oa-scraper/1.0 (+{UNPAYWALL_EMAIL}) Mozilla/5.0",
    "Accept": "*/*",
    "Accept-Language": "en-US,en;q=0.9",
})

def backoff_sleep(attempt: int) -> None:
    # exponential backoff with a tiny linear term
    delay = BASE_BACKOFF * (2 ** (attempt - 1)) + (0.05 * (attempt - 1))
    time.sleep(delay)

def retrying_get(
    url: str,
    *,
    stream: bool = False,
    expected_status: int = 200,
    headers: Optional[Dict[str, str]] = None
) -> Optional[requests.Response]:
    last_err = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            r = SESSION.get(
                url,
                timeout=TIMEOUT,
                allow_redirects=True,
                stream=stream,
                headers=headers
            )
            # handle soft rate-limits
            if r.status_code == expected_status:
                return r
            if r.status_code in (429, 500, 502, 503, 504):
                # respect retry-after if present
                ra = r.headers.get("Retry-After")
                if ra:
                    try:
                        time.sleep(float(ra))
                    except Exception:
                        pass
                last_err = RuntimeError(f"HTTP {r.status_code}")
            else:
                return None
        except requests.RequestException as e:
            last_err = e
        if attempt < MAX_RETRIES:
            backoff_sleep(attempt)
    return None

def retrying_entrez_efetch(db: str, id_: str, rettype: str, retmode: str) -> Optional[str]:
    last_err = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            handle = Entrez.efetch(db=db, id=id_, rettype=rettype, retmode=retmode)
            txt = handle.read()
            handle.close()
            if isinstance(txt, bytes):
                txt = txt.decode("utf-8", errors="ignore")
            if txt:
                return txt
            last_err = RuntimeError("Empty efetch body")
        except Exception as e:
            last_err = e
        if attempt < MAX_RETRIES:
            backoff_sleep(attempt)
    return None


# utilities

def ensure_dir(path: str) -> None:
    pathlib.Path(path).mkdir(parents=True, exist_ok=True)

def sanitize_filename(s: str) -> str:
    s = s.strip().replace(" ", "_")
    return re.sub(r"[^A-Za-z0-9_.-]+", "_", s)

def save_binary(content: bytes, path: str) -> None:
    with open(path, "wb") as f:
        f.write(content)

def save_text(content: str, path: str) -> None:
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)

def as_clean_str(x) -> Optional[str]:
    if pd.isna(x):
        return None
    s = str(x).strip()
    return s if s else None

def is_jats_xml(xml_text: str) -> bool:
    if not xml_text:
        return False
    low = xml_text.lower()
    return ("<article" in low) and ("<body" in low or "journal-meta" in low or "article-meta" in low)

# doi normalization
_DOI_PAT = re.compile(r'(10\.\d{4,9}/\S+)', re.IGNORECASE)
def normalize_doi(raw: Optional[str]) -> Optional[str]:
    if not raw or pd.isna(raw):
        return None
    s = str(raw).strip()
    # strip common url wrappers
    for pref in ("https://doi.org/", "http://doi.org/", "https://dx.doi.org/", "http://dx.doi.org/"):
        if s.lower().startswith(pref):
            s = s[len(pref):]
            break
    # trim common trailing punctuation/brackets
    s = s.strip().strip(").,;]}>")
    m = _DOI_PAT.search(s)
    return m.group(1) if m else (s if s.lower().startswith("10.") else None)

def force_https(url: str) -> str:
    return url.replace("http://", "https://", 1) if url.startswith("http://") else url

def doi_resolver_url(doi: str) -> str:
    return f"https://doi.org/{doi}"


# html/pdf validation & sniffing

_META_CIT_PDF = re.compile(
    r'<meta[^>]+name=["\']citation_pdf_url["\'][^>]+content=["\']([^"\']+)["\']',
    re.I
)
_META_CIT_FTHTML = re.compile(
    r'<meta[^>]+name=["\']citation_fulltext_html_url["\'][^>]+content=["\']([^"\']+)["\']',
    re.I
)
_PDF_HREF = re.compile(r'href=["\']([^"\']+\.pdf(?:\?[^"\']*)?)["\']', re.I)

def sniff_citation_meta(html: str) -> Tuple[Optional[str], Optional[str]]:
    pdf = None
    htm = None
    m1 = _META_CIT_PDF.search(html or "")
    if m1:
        pdf = m1.group(1).strip()
    m2 = _META_CIT_FTHTML.search(html or "")
    if m2:
        htm = m2.group(1).strip()
    return pdf, htm

def find_pdf_href(html: str) -> Optional[str]:
    m = _PDF_HREF.search(html or "")
    return m.group(1) if m else None

def is_pdf_bytes(data: bytes, content_type: str = "") -> bool:
    if not data:
        return False
    if data[:5] == b"%PDF-":
        return True
    return "pdf" in (content_type or "").lower()

def is_pdfish_url(url: str) -> bool:
    u = (url or "").lower()
    return u.endswith(".pdf") or u.endswith("/pdf") or "/pdf?" in u

def looks_like_fulltext_html(html: str, url: str = "") -> bool:
    if not html or len(html) < MIN_HTML_BYTES:
        return False
    low = html.lower()
    signals = 0
    if "<article" in low or 'role="article"' in low:
        signals += 1
    if low.count("<p") >= MIN_PTAG_COUNT:
        signals += 1
    if ("references" in low and ("<ol" in low or "<ul" in low)):
        signals += 1
    section_hits = 0
    for token in ("methods", "results", "discussion", "conclusion", "introduction", "materials and methods"):
        if token in low:
            section_hits += 1
    if section_hits >= 2:
        signals += 1
    text_only = re.sub(r"<[^>]+>", " ", low)
    word_count = len(re.findall(r"\b\w+\b", text_only))
    if word_count >= MIN_WORDS:
        signals += 1
    return signals >= 2


# pmc id converter

def pmcid_from_pmid_via_idconv(pmid: str) -> Optional[str]:
    url = (
        "https://www.ncbi.nlm.nih.gov/pmc/utils/idconv/v1.0/"
        f"?format=json&tool=script&email={NCBI_EMAIL}&ids={pmid}"
    )
    r = retrying_get(url)
    if not r:
        return None
    try:
        data = r.json()
    except json.JSONDecodeError:
        return None
    recs = data.get("records", [])
    if not recs:
        return None
    pmcid = recs[0].get("pmcid")
    return pmcid or None


# pmc (entrez)

def fetch_pmc_xml_via_entrez(pmcid: str) -> Optional[str]:
    xml_text = retrying_entrez_efetch(db="pmc", id_=pmcid, rettype="xml", retmode="text")
    if xml_text and is_jats_xml(xml_text):
        return xml_text
    return None

def download_via_pmc(pmcid: str, out_dir: str) -> Optional[str]:
    print(f"  > PMC: trying JATS XML for {pmcid} ...")
    xml_text = fetch_pmc_xml_via_entrez(pmcid)
    if xml_text:
        out = os.path.join(out_dir, f"{sanitize_filename(pmcid)}.xml")
        save_text(xml_text, out)
        print(f"    OK PMC XML: {out}")
        return out
    print("    No PMC XML.")
    return None


# europe pmc

def europe_pmc_search(doi: Optional[str], pmid: Optional[str], pmcid: Optional[str]) -> Optional[Dict[str, Any]]:
    base = "https://www.ebi.ac.uk/europepmc/webservices/rest/search"
    queries = []
    if doi:
        queries.append(f"DOI:{doi}")
    if pmcid:
        queries.append(f"EXT_ID:{pmcid}")
    if pmid:
        queries.append(f"EXT_ID:{pmid}")
    for q in queries:
        url = f"{base}?query={q}&format=json"
        r = retrying_get(url)
        time.sleep(API_DELAY)
        if not r:
            continue
        try:
            data = r.json()
        except json.JSONDecodeError:
            continue
        results = (data.get("resultList") or {}).get("result") or []
        if results:
            return results[0]
    return None

def fetch_europe_pmc_fulltext_xml(pmcid: str) -> Optional[bytes]:
    url = f"https://www.ebi.ac.uk/europepmc/webservices/rest/{pmcid}/fullTextXML"
    r = retrying_get(url)
    if r and r.text and is_jats_xml(r.text):
        return r.content
    return None

def pick_epmc_fulltext_url(result: Dict[str, Any]) -> Tuple[Optional[str], Optional[str]]:
    ft = result.get("fullTextUrlList", {})
    urls = ft.get("fullTextUrl", []) if isinstance(ft, dict) else []
    html_oa = [u for u in urls if (u.get("documentStyle","") or "").lower()=="html" and "open" in (u.get("availability","") or "").lower()]
    pdf_oa  = [u for u in urls if (u.get("documentStyle","") or "").lower()=="pdf"  and "open" in (u.get("availability","") or "").lower()]
    if html_oa: return html_oa[0].get("url"), "html"
    if pdf_oa:  return pdf_oa[0].get("url"), "pdf"
    pdf_any = [u for u in urls if (u.get("documentStyle","") or "").lower()=="pdf"]
    if pdf_any: return pdf_any[0].get("url"), "pdf"
    html_any = [u for u in urls if (u.get("documentStyle","") or "").lower()=="html"]
    if html_any: return html_any[0].get("url"), "html"
    return None, None

def download_via_europe_pmc(pmcid: Optional[str], doi: Optional[str], pmid: Optional[str], out_dir: str) -> Optional[str]:
    if pmcid:
        print(f"  > Europe PMC: trying XML for {pmcid} ...")
        xml_bytes = fetch_europe_pmc_fulltext_xml(pmcid)
        if xml_bytes:
            out = os.path.join(out_dir, f"{sanitize_filename(pmcid)}.xml")
            save_binary(xml_bytes, out)
            print(f"    OK EPMC XML: {out}")
            return out
        print("    No Europe PMC XML for this PMCID.")

    print("  > Europe PMC: searching for OA links ...")
    result = europe_pmc_search(doi=doi, pmid=pmid, pmcid=pmcid)
    if not result:
        print("    No Europe PMC record.")
        return None

    url, kind = pick_epmc_fulltext_url(result)
    if not url:
        print("    No full-text URL in Europe PMC record.")
        return None

    url = force_https(url)
    print(f"    Europe PMC: downloading {url}")
    r = retrying_get(url)
    if not r or not r.content:
        print("    Europe PMC download failed.")
        return None

    ct = (r.headers.get("Content-Type") or "").lower()
    data = r.content

    if kind == "html" and "html" in ct and not is_pdfish_url(url):
        html = data.decode("utf-8", errors="ignore")
        if looks_like_fulltext_html(html, url):
            base = pmcid or doi or pmid or "paper"
            out = os.path.join(out_dir, f"{sanitize_filename(base)}.html")
            save_text(html, out)
            print(f"    OK EPMC full-text HTML: {out}")
            return out
        pdf_meta, _ = sniff_citation_meta(html)
        if pdf_meta:
            r2 = retrying_get(force_https(pdf_meta))
            if r2 and r2.content and is_pdf_bytes(r2.content, r2.headers.get("Content-Type","")):
                base = pmcid or doi or pmid or "paper"
                out = os.path.join(out_dir, f"{sanitize_filename(base)}.pdf")
                save_binary(r2.content, out)
                print(f"    OK EPMC PDF via meta: {out}")
                return out
        print("    Europe PMC HTML did not look like full text.")
        return None

    if is_pdf_bytes(data, ct) or is_pdfish_url(url):
        base = pmcid or doi or pmid or "paper"
        out = os.path.join(out_dir, f"{sanitize_filename(base)}.pdf")
        save_binary(data, out)
        print(f"    OK EPMC PDF: {out}")
        return out

    print("    Europe PMC content not recognized.")
    return None


# unpaywall (repository-priority)

def unpaywall_lookup(doi: str) -> Optional[Dict[str, Any]]:
    url = f"https://api.unpaywall.org/v2/{doi}?email={UNPAYWALL_EMAIL}"
    # use a slightly more browser-like header set to avoid some hosts blocking
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "application/json,text/plain,*/*",
    }
    return _retrying_json(url, headers=headers)

def _retrying_json(url: str, headers: Optional[Dict[str, str]] = None) -> Optional[Dict[str, Any]]:
    last_err = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            r = SESSION.get(url, timeout=TIMEOUT, allow_redirects=True, headers=headers)
            if r.status_code == 200:
                try:
                    return r.json()
                except json.JSONDecodeError:
                    return None
            if r.status_code in (429, 500, 502, 503, 504):
                ra = r.headers.get("Retry-After")
                if ra:
                    try:
                        time.sleep(float(ra))
                    except Exception:
                        pass
                last_err = RuntimeError(f"HTTP {r.status_code}")
            else:
                return None
        except requests.RequestException as e:
            last_err = e
        if attempt < MAX_RETRIES:
            backoff_sleep(attempt)
    return None

def unpaywall_candidate_locations(data: Dict[str, Any]) -> List[Tuple[str, str]]:
    """
    Get OA locations, prioritizing repositories over publishers.
    Returns a list of (kind, url) where kind in {repo_html, repo_pdf, pub_html, pub_pdf}.
    """
    cands: List[Tuple[str, str]] = []
    locs = data.get("oa_locations") or []

    def _push(kind: str, url: Optional[str]):
        if not url:
            return
        cands.append((kind, force_https(url)))

    for loc in locs:
        host = (loc.get("host_type") or "").lower()  # repository/publisher
        url_pdf  = loc.get("url_for_pdf")
        url_html = loc.get("url")
        if url_html:
            if is_pdfish_url(url_html):
                _push("repo_pdf" if host == "repository" else "pub_pdf", url_html)
            else:
                _push("repo_html" if host == "repository" else "pub_html", url_html)
        if url_pdf:
            _push("repo_pdf" if host == "repository" else "pub_pdf", url_pdf)

    best = data.get("best_oa_location") or {}
    if best:
        bu = best.get("url")
        bp = best.get("url_for_pdf")
        bh = (best.get("host_type") or "").lower()
        if bu:
            kind = "repo_html" if bh == "repository" else "pub_html"
            if is_pdfish_url(bu):
                kind = "repo_pdf" if bh == "repository" else "pub_pdf"
            _push(kind, bu)
        if bp:
            _push("repo_pdf" if bh == "repository" else "pub_pdf", bp)

    # deduplicate and prioritize
    seen = set()
    ordered: List[Tuple[str, str]] = []
    for k, u in cands:
        if u not in seen:
            seen.add(u)
            ordered.append((k, u))

    priority = {"repo_html": 0, "repo_pdf": 1, "pub_html": 2, "pub_pdf": 3}
    ordered.sort(key=lambda kv: priority.get(kv[0], 99))
    return ordered

def download_via_unpaywall(doi: Optional[str], out_dir: str) -> Optional[str]:
    if not doi:
        return None

    print(f"  > Unpaywall: querying {doi} ...")
    data = unpaywall_lookup(doi)
    time.sleep(API_DELAY)
    if not data or not data.get("is_oa"):
        print("    No OA in Unpaywall.")
        return None

    candidates = unpaywall_candidate_locations(data)
    if not candidates:
        print("    OA found but no candidate URLs.")
        return None

    base_path = os.path.join(out_dir, sanitize_filename(doi))

    for kind, url in candidates:
        print(f"    Trying {kind}: {url}")
        headers = {
            "User-Agent": "Mozilla/5.0",
            "Accept": "*/*",
        }
        r = retrying_get(url, headers=headers)
        if not r or r.status_code != 200 or not r.content:
            print(f"      Failed (status: {r.status_code if r else 'no response'})")
            continue

        ct = (r.headers.get("Content-Type") or "").lower()
        data_bytes = r.content

        # html path
        if "html" in ct and not is_pdfish_url(url):
            html = data_bytes.decode("utf-8", errors="ignore")
            if looks_like_fulltext_html(html, url):
                out = base_path + ".html"
                save_text(html, out)
                print(f"      OK full-text HTML: {out}")
                return out
            # try meta/href pdf from the html page
            pdf_meta, _ = sniff_citation_meta(html)
            pdf_href = find_pdf_href(html)
            pdf_url = pdf_meta or pdf_href
            if pdf_url:
                if not pdf_url.startswith("http"):
                    # make relative links absolute (simple heuristic)
                    try:
                        base = url.rsplit("/", 1)[0]
                        pdf_url = base + "/" + pdf_url.lstrip("/")
                    except Exception:
                        pass
                pdf_url = force_https(pdf_url)
                r2 = retrying_get(pdf_url, headers={"Referer": url, "User-Agent": "Mozilla/5.0"})
                if r2 and r2.content and is_pdf_bytes(r2.content, r2.headers.get("Content-Type", "")):
                    out = base_path + ".pdf"
                    save_binary(r2.content, out)
                    print(f"      OK PDF from HTML link: {out}")
                    return out
            print("      HTML not full-text (or no PDF found), continuing...")
            continue

        # pdf path
        if is_pdf_bytes(data_bytes, ct) or is_pdfish_url(url):
            out = base_path + ".pdf"
            save_binary(data_bytes, out)
            print(f"      OK PDF: {out}")
            return out

        print("      Content not recognized, continuing...")

    print("    All Unpaywall candidates exhausted.")
    return None


# orchestration per row

def process_row(pmcid: Optional[str], pmid: Optional[str], doi: Optional[str]) -> Dict[str, Any]:
    outcome = {
        "pmcid": pmcid,
        "pmid": pmid,
        "doi": doi,
        "saved_path": None,
        "source": None,
        "format": None,
        "status": "failed",
        "notes": ""
    }

    # normalize pmcid
    if pmcid:
        pmcid = pmcid.strip().upper().replace(" ", "")
        if pmcid.startswith("PMCPMC"):
            pmcid = pmcid.replace("PMCPMC", "PMC", 1)
        if not pmcid.startswith("PMC") and pmcid.isdigit():
            pmcid = "PMC" + pmcid
        outcome["pmcid"] = pmcid

    # map pmid to pmcid if needed
    if not pmcid and pmid and USE_PMID_TO_PMCID:
        conv = pmcid_from_pmid_via_idconv(pmid)
        time.sleep(API_DELAY)
        if conv:
            print(f"    PMID {pmid} → PMCID {conv}")
            pmcid = conv
            outcome["pmcid"] = pmcid

    # strategy 1: pmc xml
    if pmcid:
        path = download_via_pmc(pmcid, OUTPUT_DIR)
        time.sleep(API_DELAY)
        if path:
            outcome.update({"saved_path": path, "source": "PMC", "format": "xml", "status": "ok"})
            return outcome

    # strategy 2: europe pmc (xml / html / pdf)
    path = download_via_europe_pmc(pmcid=pmcid, doi=doi, pmid=pmid, out_dir=OUTPUT_DIR)
    time.sleep(API_DELAY)
    if path:
        fmt = "xml" if path.lower().endswith(".xml") else ("pdf" if path.lower().endswith(".pdf") else "html")
        outcome.update({"saved_path": path, "source": "Europe PMC", "format": fmt, "status": "ok"})
        return outcome

    # strategy 3: unpaywall (repository-first)
    if doi:
        path = download_via_unpaywall(doi=doi, out_dir=OUTPUT_DIR)
        time.sleep(API_DELAY)
        if path:
            fmt = "pdf" if path.lower().endswith(".pdf") else "html"
            outcome.update({"saved_path": path, "source": "Unpaywall", "format": fmt, "status": "ok"})
            return outcome

    outcome["notes"] = "No full text found via PMC, Europe PMC, or Unpaywall."
    return outcome


# main

def main():
    ensure_dir(OUTPUT_DIR)
    print(f"Loading data from {EXCEL_FILE} ...")
    df = pd.read_excel(EXCEL_FILE)

    # normalize columns
    df.columns = df.columns.str.strip().str.lower()
    if "doi (link)" in df.columns and "doi" not in df.columns:
        df = df.rename(columns={"doi (link)": "doi"})

    df = df.replace(r"^\s*$", pd.NA, regex=True)

    # keep rows with at least one identifier
    need_cols = [c for c in ["pmcid", "pmid", "doi"] if c in df.columns]
    if need_cols:
        df = df.dropna(subset=need_cols, how="all")

    total = len(df)
    print(f"Found {total} rows to process.")
    print("\n" + "="*70)
    print("SOURCES ENABLED:")
    print("  - PubMed Central (PMC)")
    print("  - Europe PMC")
    print("  - Unpaywall (repository priority)")
    print("="*70 + "\n")

    results = []
    for i, row in df.iterrows():
        print(f"\n{'='*70}")
        print(f"Processing {i+1} / {total}")
        print(f"{'='*70}")

        pmcid = as_clean_str(row.get("pmcid"))
        pmid  = as_clean_str(row.get("pmid"))
        doi   = normalize_doi(as_clean_str(row.get("doi")))

        if pmid and pmid.isdigit():
            pmid = str(int(pmid))

        print(f"IDs -> PMCID:{bool(pmcid)} | PMID:{bool(pmid)} | DOI:{bool(doi)}")

        outcome = process_row(pmcid=pmcid, pmid=pmid, doi=doi)
        results.append(outcome)

    # save results
    summary_csv = os.path.join(OUTPUT_DIR, "download_summary.csv")
    results_df = pd.DataFrame(results)
    results_df.to_csv(summary_csv, index=False)

    ok_df   = results_df[results_df["status"] == "ok"].copy()
    fail_df = results_df[results_df["status"] != "ok"].copy()

    successes_csv = os.path.join(OUTPUT_DIR, "download_successes.csv")
    failures_csv  = os.path.join(OUTPUT_DIR, "download_failures.csv")
    ok_df.to_csv(successes_csv, index=False)
    fail_df.to_csv(failures_csv,  index=False)

    # stats by source/format
    by_source = ok_df["source"].value_counts().to_dict() if not ok_df.empty else {}
    by_format = ok_df["format"].value_counts().to_dict() if not ok_df.empty else {}

    print("\n" + "="*70)
    print("FINAL SUMMARY")
    print("="*70)
    print(f"Total papers processed:     {len(results_df)}")
    print(f"Successfully downloaded:    {len(ok_df)} ({(len(ok_df)/max(1,len(results_df)))*100:.1f}%)")
    print(f"Failed to download:         {len(fail_df)} ({(len(fail_df)/max(1,len(results_df)))*100:.1f}%)")

    print("\nBreakdown by source:")
    for source, count in by_source.items():
        print(f"  {source:<20} {count:>4} papers")

    print("\nBreakdown by format:")
    for fmt, count in by_format.items():
        print(f"  {fmt.upper():<20} {count:>4} files")

    if not fail_df.empty and len(fail_df) <= 20:
        print(f"\nFailed papers (showing all {len(fail_df)}):")
        for _, r in fail_df.iterrows():
            ids = []
            if r.get('pmcid'): ids.append(f"PMCID:{r['pmcid']}")
            if r.get('pmid'): ids.append(f"PMID:{r['pmid']}")
            if r.get('doi'): ids.append(f"DOI:{r['doi']}")
            print(f"  {' | '.join(ids)}")
    elif not fail_df.empty:
        print(f"\nFailed papers (showing first 10 of {len(fail_df)}):")
        for _, r in fail_df.head(10).iterrows():
            ids = []
            if r.get('pmcid'): ids.append(f"PMCID:{r['pmcid']}")
            if r.get('pmid'): ids.append(f"PMID:{r['pmid']}")
            if r.get('doi'): ids.append(f"DOI:{r['doi']}")
            print(f"  {' | '.join(ids)}")

    print(f"\n{'='*70}")
    print("OUTPUT FILES:")
    print(f"{'='*70}")
    print(f"Summary (all papers):   {summary_csv}")
    print(f"Successes only:         {successes_csv}")
    print(f"Failures only:          {failures_csv}")
    print(f"Downloaded papers in:   {os.path.abspath(OUTPUT_DIR)}")
    print(f"{'='*70}\n")

if __name__ == "__main__":
    main()
