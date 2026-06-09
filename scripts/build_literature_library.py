#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build watershed-based runoff mechanism PDF library from Literature.xlsx.
Processes 20 watersheds per batch (use --batch N).
"""

from __future__ import annotations

import argparse
import re
import time
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from urllib.parse import urlparse

import httpx
import pandas as pd
from pypdf import PdfReader
from rapidfuzz import fuzz

WORKSPACE = Path("/workspace")
INPUT_FILE = WORKSPACE / "Literature.xlsx"
LIBRARY_ROOT = WORKSPACE / "Runoff_Mechanism_Library"
CUMULATIVE_LOG = WORKSPACE / "PDF_Download_Log_cumulative.xlsx"
BATCH_SIZE = 20
ACCESS_DATE = date.today().isoformat()
EMAIL = "literature@github.com"
HEADERS = {"User-Agent": f"ExperimentalCatchment/1.0 (mailto:{EMAIL})"}


@dataclass
class Paper:
    watershed: str
    watershed_folder: str
    citation: str
    source_field: str  # McMillan or Penna
    url_hint: str = ""
    doi: str = ""
    title: str = ""
    authors: str = ""
    year: str = ""
    first_author: str = ""
    dedup_key: str = ""


@dataclass
class DownloadLogEntry:
    watershed: str
    paper_title: str
    authors: str
    year: str
    doi: str
    source_field: str
    download_result: str
    pdf_file_path: str = ""
    notes: str = ""


GITKEEP = ".gitkeep"


def _slug_part(text: str) -> str:
    text = re.sub(r"\b(experimental|catchment|watershed|basin|river|united states)\b", "", text, flags=re.I)
    text = re.sub(r"[^\w\s-]", "", text)
    text = re.sub(r"\s+", "_", text.strip())
    return re.sub(r"_+", "_", text).strip("_")


def sanitize_folder_name(watershed: str) -> str:
    """Short filesystem-safe watershed folder name (single watershed; may collide)."""
    name = watershed.strip()
    if " - " in name:
        parts = [p.strip() for p in name.split(" - ")]
        if len(parts[-1]) < 40:
            name = parts[-1]
    comma_parts = [p.strip() for p in name.split(",")]
    slug = _slug_part(comma_parts[0]) if comma_parts else ""
    if len(slug) <= 8 and len(comma_parts) > 1:
        extra = _slug_part(comma_parts[1])
        if extra:
            slug = f"{slug}_{extra}" if slug else extra
    slug = slug[:60]
    return slug or "Unknown_Watershed"


def _folder_name_candidates(watershed: str) -> list[str]:
    """Progressively more specific folder names for collision resolution."""
    name = watershed.strip()
    if " - " in name:
        dash_parts = [p.strip() for p in name.split(" - ")]
        if len(dash_parts[-1]) < 40:
            name = dash_parts[-1]
    parts = [p.strip() for p in name.split(",") if p.strip()]
    candidates: list[str] = []
    if parts:
        candidates.append(_slug_part(parts[0]))
    if len(parts) >= 2:
        candidates.append(_slug_part(f"{parts[0]}_{parts[1]}"))
    if len(parts) >= 3:
        candidates.append(_slug_part(f"{parts[0]}_{parts[1]}_{parts[2]}"))
    candidates.append(_slug_part(name.replace(",", " ")))
    seen: list[str] = []
    for c in candidates:
        c = (c or "Unknown_Watershed")[:60]
        if c not in seen:
            seen.append(c)
    return seen or ["Unknown_Watershed"]


def build_watershed_folder_map(watersheds: list[str]) -> dict[str, str]:
    """One unique folder name per watershed row (243 folders for 243 rows)."""
    assigned: dict[str, str] = {}
    used: set[str] = set()
    for ws in watersheds:
        chosen = None
        for candidate in _folder_name_candidates(ws):
            if candidate not in used:
                chosen = candidate
                break
        if not chosen:
            base = _folder_name_candidates(ws)[-1][:50]
            n = 2
            while f"{base}_{n}"[:60] in used:
                n += 1
            chosen = f"{base}_{n}"[:60]
        assigned[ws] = chosen
        used.add(chosen)
    return assigned


def extract_doi(text: str) -> str:
    if not text or str(text) == "nan":
        return ""
    m = re.search(r"10\.\d{4,9}/[^\s,;\)\]>\"']+", str(text), re.I)
    return m.group(0).rstrip(".") if m else ""


def parse_citation(citation: str) -> dict:
    cit = re.sub(r"\s+", " ", str(citation).strip())
    year_m = re.search(r"\b(19|20)\d{2}\b", cit)
    year = year_m.group(0) if year_m else ""
    doi = extract_doi(cit)
    # First author: text before first comma or before " et al"
    author_part = cit.split(" et al")[0].split(",")[0].strip()
    first_author = ""
    if author_part:
        bits = author_part.split()
        first_author = re.sub(r"[^A-Za-z]", "", bits[-1]) if bits else ""
    # Title: between year+comma patterns or after author block
    title = ""
    if year:
        after_year = cit.split(year, 1)[-1]
        after_year = re.sub(r"^[\s.,]+", "", after_year)
        title = after_year.split(".")[0].strip() if after_year else ""
        title = re.sub(r"^[\d\s.,]+", "", title)
    if not title and "." in cit:
        parts = cit.split(".")
        if len(parts) > 1:
            title = parts[1].strip()[:120]
    return {
        "year": year,
        "doi": doi,
        "first_author": first_author or "Unknown",
        "title": title[:200],
        "authors": author_part[:120],
    }


def short_title(title: str, max_len: int = 40) -> str:
    t = re.sub(r"[^\w\s]", "", title)
    words = t.split()[:5]
    s = "".join(w.capitalize() for w in words if w)
    return s[:max_len] or "Paper"


def pdf_filename(paper: Paper) -> str:
    fa = re.sub(r"[^A-Za-z]", "", paper.first_author) or "Unknown"
    st = re.sub(r"[^A-Za-z0-9]", "", short_title(paper.title)) or "Paper"
    return f"{fa}_{paper.year}_{st}.pdf"


def split_papers(cell: str) -> list[str]:
    if not cell or str(cell) == "nan":
        return []
    return [p.strip() for p in re.split(r"\s*---\s*", str(cell)) if p.strip()]


def dedup_key(paper: Paper) -> str:
    if paper.doi:
        return f"doi:{paper.doi.lower()}"
    return f"title:{paper.title.lower()[:80]}|{paper.year}|{paper.first_author.lower()}"


def collect_papers_for_row(row: pd.Series, folder_map: dict[str, str] | None = None) -> list[Paper]:
    ws = str(row["Watershed"])
    folder = (folder_map or {}).get(ws, sanitize_folder_name(ws))
    papers: list[Paper] = []
    m_cits = split_papers(row.get("McMillan_Literature", ""))
    m_urls = split_papers(row.get("McMillan_Literature_URL", ""))
    p_cits = split_papers(row.get("Penna_Full_Citation", ""))

    for i, cit in enumerate(m_cits):
        meta = parse_citation(cit)
        url_hint = m_urls[i] if i < len(m_urls) else (m_urls[0] if len(m_urls) == 1 else "")
        doi = meta["doi"] or extract_doi(url_hint)
        p = Paper(
            watershed=ws, watershed_folder=folder, citation=cit, source_field="McMillan",
            url_hint=url_hint, doi=doi, **{k: meta[k] for k in ("title", "authors", "year", "first_author")},
        )
        p.dedup_key = dedup_key(p)
        papers.append(p)

    for cit in p_cits:
        meta = parse_citation(cit)
        doi = meta["doi"]
        p = Paper(
            watershed=ws, watershed_folder=folder, citation=cit, source_field="Penna",
            url_hint="", doi=doi, **{k: meta[k] for k in ("title", "authors", "year", "first_author")},
        )
        p.dedup_key = dedup_key(p)
        papers.append(p)

    seen: set[str] = set()
    unique: list[Paper] = []
    for p in papers:
        if p.dedup_key in seen:
            continue
        # citation similarity dedup
        dup = False
        for u in unique:
            if p.doi and u.doi and p.doi.lower() == u.doi.lower():
                dup = True
                break
            if p.title and u.title and fuzz.ratio(p.title.lower(), u.title.lower()) > 92:
                dup = True
                break
            if fuzz.ratio(p.citation[:120], u.citation[:120]) > 95:
                dup = True
                break
        if dup:
            continue
        seen.add(p.dedup_key)
        unique.append(p)
    return unique


def ensure_watershed_folder(folder_name: str) -> Path:
    """Create watershed folder and .gitkeep so empty folders are tracked in git."""
    folder = LIBRARY_ROOT / folder_name
    folder.mkdir(parents=True, exist_ok=True)
    gitkeep = folder / GITKEEP
    if not gitkeep.exists():
        gitkeep.write_text("", encoding="utf-8")
    return folder


def create_all_watershed_folders(df: pd.DataFrame, save_mapping: bool = False) -> dict[str, str]:
    """Create one folder per watershed row (including those with zero PDFs)."""
    LIBRARY_ROOT.mkdir(parents=True, exist_ok=True)
    folder_map = build_watershed_folder_map(df["Watershed"].astype(str).tolist())
    for folder_name in folder_map.values():
        ensure_watershed_folder(folder_name)
    df["PDF_Library_Folder"] = df["Watershed"].astype(str).map(folder_map)
    if save_mapping:
        df.to_excel(INPUT_FILE, index=False)
    return folder_map


def search_web_pdf_urls(paper: Paper) -> list[tuple[str, str]]:
    """Mandatory supplemental web search via DuckDuckGo for legal PDF hosts."""
    try:
        from ddgs import DDGS
    except ImportError:
        try:
            from duckduckgo_search import DDGS
        except ImportError:
            return []
    queries = []
    if paper.doi:
        queries.append(f"{paper.doi} pdf")
    if paper.title and paper.year:
        queries.append(
            f'"{paper.first_author}" {paper.year} {paper.title[:70]} pdf site:zenodo.org OR site:osf.io OR site:hydroshare.org'
        )
        queries.append(
            f'"{paper.title[:60]}" {paper.year} pdf repository'
        )
    allowed_hosts = (
        "zenodo.org", "osf.io", "hydroshare.org", "edu", "gov", "arxiv.org",
        "copernicus.org", "frontiersin.org", "mdpi.com", "usgs.gov", "usda.gov",
        "fs.fed.us", "lter", "digitalcommons", "repository", "hal.", "ssrn.com",
    )
    out: list[tuple[str, str]] = []
    seen: set[str] = set()
    ddgs = DDGS()
    for q in queries[:3]:
        try:
            for r in ddgs.text(q, max_results=6):
                url = r.get("href", "")
                if not url or url in seen:
                    continue
                ul = url.lower()
                if ".pdf" in ul or any(h in ul for h in allowed_hosts):
                    if any(b in ul for b in ("sci-hub", "libgen", "z-lib")):
                        continue
                    seen.add(url)
                    out.append((url, "WebSearch"))
        except Exception:
            pass
        time.sleep(1.0)
    return out


def crossref_pdf_link(client: httpx.Client, doi: str) -> str | None:
    try:
        r = client.get(
            f"https://api.crossref.org/works/{doi}",
            headers={**HEADERS, "mailto": EMAIL},
            timeout=25,
        )
        if r.status_code != 200:
            return None
        for link in r.json().get("message", {}).get("link", []):
            if link.get("content-type") == "application/pdf" and link.get("URL"):
                return link["URL"]
    except Exception:
        pass
    return None


def zenodo_search_pdf(client: httpx.Client, paper: Paper) -> list[tuple[str, str]]:
    out: list[tuple[str, str]] = []
    q = paper.doi or f"{paper.title} {paper.year}"
    try:
        r = client.get(
            "https://zenodo.org/api/records",
            params={"q": q, "size": 3},
            timeout=25,
        )
        if r.status_code == 200:
            for rec in r.json().get("hits", {}).get("hits", []):
                for f in rec.get("files", []):
                    key = f.get("key", "")
                    if key.lower().endswith(".pdf"):
                        out.append((f.get("links", {}).get("self", ""), "Zenodo"))
    except Exception:
        pass
    return [(u, s) for u, s in out if u]


def copernicus_pdf_url(doi: str) -> str | None:
    m = re.match(r"10\.5194/([a-z]+)-(\d+)-(\d+)-(\d+)", doi, re.I)
    if not m:
        return None
    journal, vol, page, year = m.groups()
    return f"https://{journal}.copernicus.org/articles/{vol}/{page}/{year}/{journal}-{vol}-{page}-{year}.pdf"


def frontiers_pdf_url(doi: str) -> str | None:
    if not doi.lower().startswith("10.3389/"):
        return None
    return f"https://doi.org/{doi}/pdf"


def plos_pdf_url(doi: str) -> str | None:
    if not doi.lower().startswith("10.1371/"):
        return None
    return f"https://journals.plos.org/plosone/article/file?id={doi}&type=printable"


def mdpi_pdf_url(doi: str) -> str | None:
    m = re.match(r"10\.3390/([a-z0-9]+)", doi, re.I)
    if not m:
        return None
    return f"https://www.mdpi.com/{m.group(1)}/pdf"


def wiley_agu_pdf_url(doi: str) -> str | None:
    if doi.lower().startswith(("10.1029/", "10.1002/", "10.1111/")):
        return f"https://doi.org/{doi}/pdf"
    return None


def essd_pdf_url(doi: str) -> str | None:
    return copernicus_pdf_url(doi)  # earth system science data etc.


def find_oa_urls(client: httpx.Client, paper: Paper) -> list[tuple[str, str]]:
    """Return list of (url, source_name) to try."""
    candidates: list[tuple[str, str]] = []
    doi = paper.doi

    if paper.url_hint and paper.url_hint.lower().endswith(".pdf"):
        candidates.append((paper.url_hint, "McMillan_URL_direct"))

    if doi:
        for fn, name in [
            (copernicus_pdf_url, "Copernicus"),
            (frontiers_pdf_url, "Frontiers"),
            (plos_pdf_url, "PLOS"),
            (mdpi_pdf_url, "MDPI"),
            (wiley_agu_pdf_url, "Publisher_PDF"),
        ]:
            u = fn(doi)
            if u:
                candidates.append((u, name))

        # Unpaywall
        try:
            r = client.get(f"https://api.unpaywall.org/v2/{doi}", params={"email": EMAIL}, timeout=25)
            if r.status_code == 200:
                j = r.json()
                loc = j.get("best_oa_location") or {}
                pdf = loc.get("url_for_pdf") or loc.get("url")
                if pdf:
                    candidates.append((pdf, "Unpaywall"))
        except Exception:
            pass

        # OpenAlex
        try:
            r = client.get(f"https://api.openalex.org/works/https://doi.org/{doi}", timeout=25)
            if r.status_code == 200:
                j = r.json()
                oa = j.get("open_access", {}).get("oa_url")
                if oa:
                    candidates.append((oa, "OpenAlex_OA"))
                for loc in j.get("locations", []) or []:
                    pdf = loc.get("pdf_url")
                    if pdf:
                        candidates.append((pdf, "OpenAlex_location"))
        except Exception:
            pass

        # Semantic Scholar
        try:
            r = client.get(
                f"https://api.semanticscholar.org/graph/v1/paper/DOI:{doi}",
                params={"fields": "openAccessPdf,externalIds"},
                timeout=25,
            )
            if r.status_code == 200:
                oa = r.json().get("openAccessPdf") or {}
                if oa.get("url"):
                    candidates.append((oa["url"], "SemanticScholar"))
        except Exception:
            pass

        # CORE API
        try:
            r = client.get(
                "https://api.core.ac.uk/v3/search/works",
                params={"q": f"doi:{doi}", "limit": 1},
                headers={**HEADERS, "Authorization": "Bearer"},
                timeout=25,
            )
        except Exception:
            pass

        cr = crossref_pdf_link(client, doi)
        if cr:
            candidates.append((cr, "Crossref"))

        # Scrape DOI landing page for PDF links
        try:
            r = client.get(f"https://doi.org/{doi}", timeout=30, follow_redirects=True)
            if r.status_code == 200 and "html" in r.headers.get("content-type", ""):
                pdfs = re.findall(r'href=["\']([^"\']+\.pdf[^"\']*)["\']', r.text, re.I)
                for pdf in pdfs[:5]:
                    if pdf.startswith("/"):
                        base = f"{urlparse(str(r.url)).scheme}://{urlparse(str(r.url)).netloc}"
                        pdf = base + pdf
                    candidates.append((pdf, "DOI_landing_pdf"))
        except Exception:
            pass

        candidates.extend(zenodo_search_pdf(client, paper))

    candidates.extend(search_web_pdf_urls(paper))

    # dedupe urls
    seen_u: set[str] = set()
    out: list[tuple[str, str]] = []
    for u, s in candidates:
        if u and u not in seen_u:
            seen_u.add(u)
            out.append((u, s))
    return out


def validate_pdf(content: bytes, paper: Paper) -> tuple[bool, str]:
    if len(content) < 5000:
        return False, "File too small for full paper PDF"
    if not content[:5].startswith(b"%PDF"):
        return False, "Not a PDF file"
    try:
        import io
        reader = PdfReader(io.BytesIO(content))
        if len(reader.pages) < 2:
            return False, "PDF has fewer than 2 pages (likely abstract-only)"
        text = ""
        for i in range(min(3, len(reader.pages))):
            text += (reader.pages[i].extract_text() or "").lower()
        if len(text) < 200:
            return False, "Insufficient extractable text"
        score_title = fuzz.partial_ratio(paper.title.lower()[:60], text) if paper.title else 50
        score_year = 100 if paper.year and paper.year in text else 30
        score_author = 100 if paper.first_author.lower() in text else 40
        doi_ok = paper.doi and paper.doi.lower().replace("/", "") in text.replace("/", "")
        if paper.title and score_title < 35 and not doi_ok:
            return False, f"Title mismatch (score={score_title})"
        if paper.year and score_year < 50 and not doi_ok:
            return False, f"Year {paper.year} not found in PDF"
        if paper.first_author != "Unknown" and score_author < 40 and not doi_ok:
            return False, f"Author {paper.first_author} not found in PDF"
        return True, "Validated"
    except Exception as e:
        return False, f"PDF read error: {e}"


def download_pdf(client: httpx.Client, url: str) -> tuple[bytes | None, str]:
    try:
        r = client.get(url, timeout=90, follow_redirects=True)
        if r.status_code != 200:
            return None, f"HTTP {r.status_code}"
        ct = r.headers.get("content-type", "").lower()
        if "pdf" in ct or r.content[:5].startswith(b"%PDF"):
            return r.content, "OK"
        if "html" in ct:
            pdfs = re.findall(r'href=["\']([^"\']+\.pdf[^"\']*)["\']', r.text, re.I)
            if pdfs:
                return download_pdf(client, pdfs[0])
        return None, f"Not PDF content-type: {ct}"
    except Exception as e:
        return None, str(e)


def process_paper(client: httpx.Client, paper: Paper) -> DownloadLogEntry:
    entry = DownloadLogEntry(
        watershed=paper.watershed,
        paper_title=paper.title,
        authors=paper.authors,
        year=paper.year,
        doi=paper.doi,
        source_field=paper.source_field,
        download_result="Missing",
        notes="",
    )

    dest_dir = LIBRARY_ROOT / paper.watershed_folder
    dest_dir.mkdir(parents=True, exist_ok=True)
    fname = pdf_filename(paper)
    dest_path = dest_dir / fname

    if dest_path.exists() and dest_path.stat().st_size > 5000:
        entry.download_result = "Downloaded"
        entry.pdf_file_path = str(dest_path.relative_to(WORKSPACE))
        entry.notes = "Already exists"
        return entry

    candidates = find_oa_urls(client, paper)
    if not candidates:
        entry.notes = "No legal OA PDF URL found"
        if not paper.doi:
            entry.download_result = "Ambiguous"
            entry.notes = "No DOI; could not locate legal PDF"
        return entry

    for url, source in candidates:
        content, err = download_pdf(client, url)
        if not content:
            entry.notes += f" [{source}: {err}]"
            continue
        ok, reason = validate_pdf(content, paper)
        if ok:
            dest_path.write_bytes(content)
            entry.download_result = "Downloaded"
            entry.pdf_file_path = str(dest_path.relative_to(WORKSPACE))
            entry.notes = f"Downloaded via {source}. {reason}"
            return entry
        entry.notes += f" [{source}: rejected - {reason}]"

    if paper.doi and "paywall" not in entry.notes.lower():
        entry.notes = (entry.notes + " Paywall or no legal OA PDF.").strip()
    return entry


def update_watershed_row(row: pd.Series, papers: list[Paper], logs: list[DownloadLogEntry]) -> dict:
    ws_logs = [l for l in logs if l.watershed == row["Watershed"]]
    downloaded = sum(1 for l in ws_logs if l.download_result == "Downloaded")
    missing = sum(1 for l in ws_logs if l.download_result == "Missing")
    ambiguous = sum(1 for l in ws_logs if l.download_result == "Ambiguous")
    total = len(ws_logs)

    if total == 0:
        status = "Complete"
    elif missing == 0 and ambiguous == 0:
        status = "Complete"
    else:
        status = "Incomplete"

    missing_cits = []
    for l in ws_logs:
        if l.download_result in ("Missing", "Ambiguous"):
            # find citation
            for p in papers:
                if p.watershed == l.watershed and (p.title == l.paper_title or p.doi == l.doi):
                    missing_cits.append(p.citation)
                    break
            else:
                missing_cits.append(l.paper_title or l.doi or "Unknown citation")

    notes_parts = []
    if downloaded:
        notes_parts.append(f"{downloaded} papers downloaded")
    if missing:
        notes_parts.append(f"{missing} paywalled or no legal PDF")
    if ambiguous:
        notes_parts.append(f"{ambiguous} citation ambiguous")

    return {
        "Downloaded_PDF_Count": downloaded,
        "Missing_PDF_Count": missing + ambiguous,
        "PDF_Status": status,
        "Download_Notes": "; ".join(notes_parts) + "." if notes_parts else "No papers identified.",
        "Missing_PDF_Citations": "\n---\n".join(missing_cits),
    }


def load_cumulative_log() -> pd.DataFrame:
    cols = [
        "Watershed", "Paper_Title", "Authors", "Year", "DOI", "Source_Field",
        "Download_Result", "PDF_File_Path", "Notes",
    ]
    if CUMULATIVE_LOG.exists():
        return pd.read_excel(CUMULATIVE_LOG)
    return pd.DataFrame(columns=cols)


def save_cumulative_log(log_df: pd.DataFrame) -> None:
    log_df.to_excel(CUMULATIVE_LOG, index=False)


def run_batch(batch_num: int, df: pd.DataFrame | None = None, cumulative: pd.DataFrame | None = None) -> tuple[pd.DataFrame, pd.DataFrame]:
    start = (batch_num - 1) * BATCH_SIZE
    end = start + BATCH_SIZE

    if df is None:
        df = pd.read_excel(INPUT_FILE)
    if cumulative is None:
        cumulative = load_cumulative_log()

    new_cols = {
        "Downloaded_PDF_Count": 0,
        "Missing_PDF_Count": 0,
        "PDF_Status": "",
        "Download_Notes": "",
        "Missing_PDF_Citations": "",
    }
    for col, default in new_cols.items():
        if col not in df.columns:
            df[col] = default
    df["Downloaded_PDF_Count"] = pd.to_numeric(df["Downloaded_PDF_Count"], errors="coerce").fillna(0).astype(int)
    df["Missing_PDF_Count"] = pd.to_numeric(df["Missing_PDF_Count"], errors="coerce").fillna(0).astype(int)

    batch_df = df.iloc[start:end].copy()
    all_logs: list[DownloadLogEntry] = []
    LIBRARY_ROOT.mkdir(parents=True, exist_ok=True)
    folder_map = build_watershed_folder_map(df["Watershed"].astype(str).tolist())
    df["PDF_Library_Folder"] = df["Watershed"].astype(str).map(folder_map)

    print(f"Batch {batch_num:02d}: watersheds {start+1}-{min(end, len(df))} ({len(batch_df)} rows)")

    with httpx.Client(headers=HEADERS) as client:
        for idx, row in batch_df.iterrows():
            ws = str(row["Watershed"])
            ensure_watershed_folder(folder_map[ws])
            papers = collect_papers_for_row(row, folder_map)
            print(f"  {row['Watershed'][:50]}: {len(papers)} papers")
            for paper in papers:
                entry = process_paper(client, paper)
                all_logs.append(entry)
                status = entry.download_result
                print(f"    [{status}] {paper.first_author} {paper.year} | {paper.doi or 'no-doi'}")
                time.sleep(0.5)

            ws_logs = [l for l in all_logs if l.watershed == row["Watershed"]]
            updates = update_watershed_row(row, papers, ws_logs)
            for k, v in updates.items():
                df.at[idx, k] = v
            # checkpoint log after each watershed
            log_df = pd.DataFrame([l.__dict__ for l in all_logs])
            log_df = log_df.rename(columns={
                "watershed": "Watershed", "paper_title": "Paper_Title",
                "authors": "Authors", "year": "Year", "doi": "DOI",
                "source_field": "Source_Field", "download_result": "Download_Result",
                "pdf_file_path": "PDF_File_Path", "notes": "Notes",
            })
            checkpoint = WORKSPACE / f"Literature_batch{batch_num:02d}.xlsx"
            with pd.ExcelWriter(checkpoint, engine="openpyxl") as w:
                df.to_excel(w, sheet_name="Literature", index=False)
                log_df.to_excel(w, sheet_name="PDF_Download_Log", index=False)

    # Save batch output
    out_xlsx = WORKSPACE / f"Literature_batch{batch_num:02d}.xlsx"
    log_df = pd.DataFrame([l.__dict__ for l in all_logs])
    log_df = log_df.rename(columns={
        "watershed": "Watershed",
        "paper_title": "Paper_Title",
        "authors": "Authors",
        "year": "Year",
        "doi": "DOI",
        "source_field": "Source_Field",
        "download_result": "Download_Result",
        "pdf_file_path": "PDF_File_Path",
        "notes": "Notes",
    })

    # Merge into cumulative log (replace same watershed+doi entries from this batch)
    if not cumulative.empty:
        batch_keys = set(
            (r["Watershed"], str(r.get("DOI", "")))
            for _, r in log_df.iterrows()
        )
        cumulative = cumulative[
            ~cumulative.apply(
                lambda r: (r["Watershed"], str(r.get("DOI", "") if pd.notna(r.get("DOI")) else "")) in batch_keys,
                axis=1,
            )
        ]
    cumulative = pd.concat([cumulative, log_df], ignore_index=True)
    save_cumulative_log(cumulative)

    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Literature", index=False)
        log_df.to_excel(w, sheet_name="PDF_Download_Log", index=False)
        cumulative.to_excel(w, sheet_name="PDF_Download_Log_Cumulative", index=False)

    # Also update master Literature.xlsx for processed rows only
    master = pd.read_excel(INPUT_FILE)
    for col, default in new_cols.items():
        if col not in master.columns:
            master[col] = default
    master["Downloaded_PDF_Count"] = pd.to_numeric(master["Downloaded_PDF_Count"], errors="coerce").fillna(0).astype(int)
    master["Missing_PDF_Count"] = pd.to_numeric(master["Missing_PDF_Count"], errors="coerce").fillna(0).astype(int)
    for idx in batch_df.index:
        for col in new_cols:
            master.at[idx, col] = df.at[idx, col]
    master.to_excel(INPUT_FILE, index=False)

    print(f"\nSaved {out_xlsx}")
    print(f"PDFs in {LIBRARY_ROOT}")
    downloaded = sum(1 for l in all_logs if l.download_result == "Downloaded")
    print(f"Downloaded: {downloaded}/{len(all_logs)}")
    return df, cumulative


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--batch", type=int, default=1, help="Batch number (1-based, 20 watersheds each)")
    parser.add_argument("--from-batch", type=int, default=None, help="Run batches from N to --to-batch")
    parser.add_argument("--to-batch", type=int, default=None, help="Run batches through M inclusive")
    parser.add_argument("--create-folders-only", action="store_true", help="Create all watershed folders and exit")
    parser.add_argument("--sync-folders", action="store_true", help="Create all folders + .gitkeep, update PDF_Library_Folder")
    args = parser.parse_args()

    df = pd.read_excel(INPUT_FILE)
    if args.sync_folders or args.create_folders_only:
        folder_map = create_all_watershed_folders(df, save_mapping=True)
        n_folders = len(set(folder_map.values()))
        print(f"Created/verified {n_folders} watershed folders under {LIBRARY_ROOT} (one per row)")
        return

    create_all_watershed_folders(df, save_mapping=False)

    if args.from_batch and args.to_batch:
        cumulative = load_cumulative_log()
        for b in range(args.from_batch, args.to_batch + 1):
            print(f"\n{'='*60}\nRunning batch {b}\n{'='*60}")
            df, cumulative = run_batch(b, df=df, cumulative=cumulative)
        return

    run_batch(args.batch, df=df)


if __name__ == "__main__":
    main()
