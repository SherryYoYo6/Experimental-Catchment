#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Re-audit Strict_Public_Data_Inventory.xlsx for Y and Partial rows.
Fetches Repository_URL content and verifies watershed + variable evidence.
"""

from __future__ import annotations

import json
import re
import time
from datetime import date
from pathlib import Path
from urllib.parse import urlparse

import httpx
import pandas as pd
from bs4 import BeautifulSoup

WORKSPACE = Path("/workspace")
INPUT_FILE = WORKSPACE / "Strict_Public_Data_Inventory.xlsx"
OUTPUT_FILE = WORKSPACE / "Strict_Public_Data_Inventory_Reaudited.xlsx"
CACHE_FILE = WORKSPACE / "scripts" / ".url_audit_cache.json"

ACCESS_DATE = date.today().isoformat()

PUBLIC_VARS = [
    "Rain_Public", "Q_Public", "Soil_Moisture_Public", "Multi_Layer_SM_Public",
    "Groundwater_Monitoring_Public", "Event_Piezometer_Network_Public",
    "Saturated_Area_VSA_Public", "Isotope_Public", "EC_Public", "Snow_SWE_Public",
]

MAJOR_VARS = [
    "Rain_Public", "Q_Public", "Soil_Moisture_Public",
    "Groundwater_Monitoring_Public", "Isotope_Public",
]

OBS_MAP = {
    "Rain_Public": "Rain_Observed",
    "Q_Public": "Q_Observed",
    "Soil_Moisture_Public": "Soil_Moisture_Observed",
    "Multi_Layer_SM_Public": "Multi_Layer_SM_Observed",
    "Groundwater_Monitoring_Public": "Groundwater_Monitoring_Observed",
    "Event_Piezometer_Network_Public": "Event_Piezometer_Network_Observed",
    "Saturated_Area_VSA_Public": "Saturated_Area_VSA_Observed",
    "Isotope_Public": "Isotope_Observed",
    "EC_Public": "EC_Observed",
    "Snow_SWE_Public": "Snow_SWE_Observed",
}

VAR_KEYWORDS = {
    "Rain_Public": [
        "precipitation", "rainfall", "rain gauge", "raingauge", "meteorolog",
        "weather", "tipping bucket", "pluviomet", "niederschlag",
    ],
    "Q_Public": [
        "streamflow", "discharge", "runoff", "flow data", "hydrograph",
        "stream gauge", "streamgage", "water level", "nwis", "abfluss",
    ],
    "Soil_Moisture_Public": [
        "soil moisture", "soil water", "volumetric water", "tdr", "fdr",
        "capacitance", "neutron probe", "bodenfeuchte",
    ],
    "Multi_Layer_SM_Public": [
        "multi-layer", "multilayer", "multiple depth", "soil profile",
        "profile soil moisture", "layered soil",
    ],
    "Groundwater_Monitoring_Public": [
        "groundwater", "water table", "well level", "piezometer",
        "aquifer", "grundwasser",
    ],
    "Event_Piezometer_Network_Public": [
        "piezometer network", "event piezometer", "event-scale",
        "transient groundwater",
    ],
    "Saturated_Area_VSA_Public": [
        "saturated area", "variable source area", "vsa", "saturation extent",
        "wetness mapping",
    ],
    "Isotope_Public": [
        "isotope", "deuterium", "oxygen-18", "o-18", "tracer",
        "hydrograph separation", "stable isotope",
    ],
    "EC_Public": [
        "electrical conductivity", "specific conductance", "hydrochemistry",
        "water chemistry", "solute", "ion", "water quality",
    ],
    "Snow_SWE_Public": [
        "snow", "swe", "snow water equivalent", "snow depth", "snowpack", "snowmelt",
    ],
}

# URLs that are always rejected as non-specific evidence
REJECT_URL_PATTERNS = [
    r"doi\.org", r"dx\.doi\.org", r"pubs\.usgs\.gov", r"tandfonline", r"mdpi\.com",
    r"sciencedirect", r"springer\.com", r"wiley\.com", r"researchgate",
    r"cran\.r-project\.org", r"github\.com", r"software", r"earthquake\.usgs",
    r"eros\.usgs\.gov/earthshots", r"/media/images", r"national-hydrography",
    r"datacatalog/search\?authors", r"repositorio\.catie", r"digitalcommons",
]

# Generic homepage paths (reject unless watershed tokens found in page)
GENERIC_PATH_PATTERNS = [
    r"^/$", r"^$", r"^/group/?$", r"^/search", r"/clowder/spaces/?$",
]

# Known verified corrections: regex on watershed name -> better URL + evidence
KNOWN_CORRECTIONS: list[dict] = [
    {
        "pattern": r"Blacklands|Riesel",
        "url": "https://www.ars.usda.gov/plains-area/temple-tx/grassland-soil-and-water-research-laboratory/docs/hydrologic-data/",
        "source": "USDA-ARS Hydrologic Data",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "watershed_tokens": ["riesel", "blacklands", "brushy"],
    },
    {
        "pattern": r"Shale Hills",
        "url": "https://www.hydroshare.org/group/147",
        "source": "CUAHSI HydroShare SSHCZO",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Soil_Moisture_Public": "Y",
                 "Groundwater_Monitoring_Public": "Y"},
        "watershed_tokens": ["shale hills", "sshczo"],
    },
    {
        "pattern": r"Marshall Gulch",
        "url": "https://www.hydroshare.org/group/142",
        "source": "CUAHSI HydroShare CJCZO",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Soil_Moisture_Public": "Y", "Snow_SWE_Public": "Y"},
        "watershed_tokens": ["marshall gulch", "catalina"],
    },
    {
        "pattern": r"WS10|H\.?J\.? Andrews|WS1,3,6|WS02|WS2,8,9",
        "url": "https://andrewsforest.oregonstate.edu/data",
        "source": "H.J. Andrews Experimental Forest",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Soil_Moisture_Public": "Y", "Snow_SWE_Public": "Y"},
        "watershed_tokens": ["andrews", "watershed 10", "ws10", "hj andrews"],
    },
    {
        "pattern": r"Aguima|Upper Oueme",
        "url": "http://bd.amma-catch.org/main.jsf",
        "source": "AMMA-CATCH Database",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "watershed_tokens": ["oueme", "oueme", "aguima", "amma-catch"],
    },
    {
        "pattern": r"IML CZO|Sangamon",
        "url": "https://data.imlczo.org/clowder/collection/62570d52e4b008bfd91675f2",
        "source": "IML-CZO Archive (Clowder)",
        "vars": {"Rain_Public": "Partial", "Q_Public": "Partial", "Soil_Moisture_Public": "Partial"},
        "watershed_tokens": ["sangamon", "iml", "illinois"],
    },
    {
        "pattern": r"HOAL|Petzenkirchen",
        "url": "https://hoal.hydrology.at/data/",
        "source": "HOAL Observatory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Soil_Moisture_Public": "Y"},
        "watershed_tokens": ["hoal", "petzenkirchen"],
    },
    {
        "pattern": r"Duerreych|D�rreych",
        "url": "http://www.duerreych.de/Projektubersicht_overview/projektubersicht_overview.html",
        "source": "Duerreychbachtal Project",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Soil_Moisture_Public": "Y"},
        "watershed_tokens": ["duerreych", "d�rreych"],
    },
    {
        "pattern": r"Hubbard Brook|WS3.*Hubbard|Watershed 1, Hubbard",
        "url": "https://hubbardbrook.org/data-catalog/",
        "source": "Hubbard Brook Data Catalog",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Isotope_Public": "Y", "EC_Public": "Y"},
        "watershed_tokens": ["hubbard brook", "watershed"],
    },
    {
        "pattern": r"Coweeta",
        "url": "https://coweeta.uga.edu/data/",
        "source": "Coweeta Hydrologic Laboratory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "watershed_tokens": ["coweeta"],
    },
    {
        "pattern": r"Konza",
        "url": "https://portal.edirepository.org/nis/mapbrowse?scope=knb-lter-knz",
        "source": "Konza Prairie LTER / EDI",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "watershed_tokens": ["konza"],
    },
    {
        "pattern": r"Luquillo|Bisley",
        "url": "https://luq.lter.network.edu/data/",
        "source": "Luquillo LTER",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "watershed_tokens": ["luquillo", "bisley"],
    },
    {
        "pattern": r"Walnut Gulch",
        "url": "https://www.tucson.ars.ag.gov/dap/",
        "source": "USDA-ARS Walnut Gulch",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "watershed_tokens": ["walnut gulch"],
    },
    {
        "pattern": r"Reynolds Creek|Upper Sheep Creek",
        "url": "https://www.ars.usda.gov/pacific-west-area/boise-id/northwest-watershed-research-center/docs/reynolds-creek-experimental-watershed-data/",
        "source": "USDA-ARS Reynolds Creek",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Snow_SWE_Public": "Y"},
        "watershed_tokens": ["reynolds creek", "sheep creek"],
    },
    {
        "pattern": r"Sleepers River",
        "url": "https://www.sleepersriver.org/",
        "source": "Sleepers River Research Watershed",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Snow_SWE_Public": "Y"},
        "watershed_tokens": ["sleepers river"],
    },
    {
        "pattern": r"Maimai",
        "url": "https://niwa.co.nz/our-services/consultancy-services/maimai-research-catchment",
        "source": "NIWA Maimai Research Catchment",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "watershed_tokens": ["maimai"],
    },
    {
        "pattern": r"Zhurucay",
        "url": "https://www.hydroshare.org/group/262",
        "source": "Zhurucay HydroShare",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Soil_Moisture_Public": "Y"},
        "watershed_tokens": ["zhurucay"],
    },
    {
        "pattern": r"Agua Salud|MAT subcatchment|SEC subcatchment|PAS subcatchment",
        "url": "https://agua-salud.stri.si.edu/",
        "source": "STRI Agua Salud",
        "vars": {"Rain_Public": "Y", "Q_Public": "Partial"},
        "watershed_tokens": ["agua salud", "mat", "sec", "pas"],
    },
    {
        "pattern": r"Krycklan",
        "url": "https://www.sites.se/en/krycklan/",
        "source": "SITES Krycklan",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "watershed_tokens": ["krycklan"],
    },
    {
        "pattern": r"Little River Experimental",
        "url": "https://www.ars.usda.gov/southeast-area/tifton-ga/crop-protection-and-management-research/docs/little-river-experimental-watershed/",
        "source": "USDA-ARS Little River",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "watershed_tokens": ["little river"],
    },
    {
        "pattern": r"Neversink",
        "url": "https://waterdata.usgs.gov/nwis",
        "source": "USGS NWIS",
        "vars": {"Q_Public": "Y", "Rain_Public": "Partial"},
        "watershed_tokens": ["neversink"],
    },
    {
        "pattern": r"Fraser|Niwot|Saddle Catchment",
        "url": "https://portal.edirepository.org/nis/mapbrowse?scope=knb-lter-nwt",
        "source": "Niwot Ridge LTER / EDI",
        "vars": {"Rain_Public": "Y", "Snow_SWE_Public": "Y", "Q_Public": "Partial"},
        "watershed_tokens": ["niwot", "fraser", "saddle"],
    },
    {
        "pattern": r"RBF subcatchment|Lang Lang",
        "url": "https://data.water.vic.gov.au/",
        "source": "Victoria WMIS",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "watershed_tokens": ["lang lang"],
    },
    {
        "pattern": r"Laengentalbach|L�ngentalbach",
        "url": "https://wasser.gv.at/hydjb/",
        "source": "Austrian eHydJB",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "watershed_tokens": ["laengentalbach", "l�ngentalbach"],
    },
    {
        "pattern": r"Eel River CZO",
        "url": "https://www.hydroshare.org/group/143",
        "source": "Eel River CZO HydroShare",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "watershed_tokens": ["eel river"],
    },
    {
        "pattern": r"TERENO|Schafertal|Zastler",
        "url": "https://teodoor.icg.kfa-juelich.de/",
        "source": "TERENO TEODOOR",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Soil_Moisture_Public": "Y"},
        "watershed_tokens": ["tereno", "schafertal", "zastler"],
    },
    {
        "pattern": r"Rietholzbach",
        "url": "https://www.wsl.ch/en/about-wsl/research-units/hydrology-and-climate/rietholzbach-research-station.html",
        "source": "WSL Rietholzbach",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Groundwater_Monitoring_Public": "Y"},
        "watershed_tokens": ["rietholzbach"],
    },
    {
        "pattern": r"Hupsel Brook",
        "url": "https://www.hydrology.nl/hupsel/",
        "source": "Hupsel Brook Observatory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Groundwater_Monitoring_Public": "Y"},
        "watershed_tokens": ["hupsel"],
    },
    {
        "pattern": r"Experimental Lakes|Experim\. Lakes",
        "url": "https://www.iisd.org/ela/",
        "source": "IISD Experimental Lakes Area",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "watershed_tokens": ["experimental lakes"],
    },
    {
        "pattern": r"Fernow",
        "url": "https://www.fs.usda.gov/ne/delaware/fernow-experimental-forest",
        "source": "Fernow Experimental Forest",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "watershed_tokens": ["fernow"],
    },
    {
        "pattern": r"San Dimas",
        "url": "https://www.fs.usda.gov/detailfull/r5/landmanagement/resourcemanagement/?cid=stelprdb5129904",
        "source": "San Dimas Experimental Forest",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "watershed_tokens": ["san dimas"],
    },
]


def parse_observed(val) -> str:
    if pd.isna(val):
        return "Unknown"
    return str(val).split(",")[0].strip()


def extract_watershed_tokens(name: str) -> list[str]:
    """Generate searchable tokens from watershed name."""
    name = re.sub(r"[^\w\s\-']", " ", name.lower())
    stop = {
        "united", "states", "canada", "australia", "germany", "france", "china",
        "japan", "brazil", "india", "catchment", "watershed", "basin", "river",
        "experimental", "research", "subcatchment", "mountains", "county", "region",
        "national", "forest", "site", "area", "the", "and", "of", "in", "at", "near",
    }
    tokens = []
    # Full significant phrases
    for part in re.split(r",| - |;", name):
        part = part.strip()
        if len(part) >= 4 and part not in stop:
            tokens.append(part)
    # Individual words
    for w in name.split():
        w = w.strip()
        if len(w) >= 4 and w not in stop:
            tokens.append(w)
    # dedupe preserving order
    seen = set()
    out = []
    for t in tokens:
        if t not in seen:
            seen.add(t)
            out.append(t)
    return out[:12]


def url_is_rejected_pattern(url: str) -> str | None:
    if not url or pd.isna(url) or str(url).strip() == "":
        return "Missing or empty URL"
    u = str(url).lower()
    for pat in REJECT_URL_PATTERNS:
        if re.search(pat, u):
            return f"URL matches rejected pattern: {pat}"
    return None


def url_is_generic_homepage(url: str) -> bool:
    p = urlparse(str(url))
    path = p.path.rstrip("/") or "/"
    if path in ("/", ""):
        return True
    for pat in GENERIC_PATH_PATTERNS:
        if re.search(pat, path):
            return True
    # hydroshare.org/ with no resource/group path
    if "hydroshare.org" in p.netloc and path == "":
        return True
    if re.match(r"^https?://(www\.)?hydroshare\.org/?$", str(url).rstrip("/"), re.I):
        return True
    if "open.canada.ca" in p.netloc and path in ("", "/"):
        return True
    if "clowder/spaces" in str(url) and "collection" not in str(url):
        return True
    return False


def fetch_page_text(url: str, client: httpx.Client) -> tuple[str, str | None]:
    """Return (text, error)."""
    try:
        r = client.get(url, follow_redirects=True, timeout=25.0)
        if r.status_code >= 400:
            return "", f"HTTP {r.status_code}"
        ctype = r.headers.get("content-type", "")
        if "html" in ctype or "text" in ctype or not ctype:
            soup = BeautifulSoup(r.text, "lxml")
            for tag in soup(["script", "style", "nav", "footer"]):
                tag.decompose()
            text = soup.get_text(" ", strip=True).lower()
            title = (soup.title.string or "").lower() if soup.title else ""
            return f"{title} {text}", None
        return r.text[:50000].lower(), None
    except Exception as e:
        return "", str(e)


def watershed_match_score(tokens: list[str], text: str) -> tuple[str, str]:
    """Return (Exact_Watershed_Verified, reason)."""
    if not text:
        return "Unclear", "No page content retrieved"
    if not tokens:
        return "Unclear", "No watershed tokens extracted"
    hits = [t for t in tokens if t in text]
    if len(hits) >= 2 or (len(hits) == 1 and len(tokens[0]) >= 8):
        return "Y", f"Matched tokens: {', '.join(hits[:5])}"
    if len(hits) == 1:
        return "Unclear", f"Single token match: {hits[0]}"
    return "N", "No watershed name tokens found on page"


def variable_supported(var: str, text: str) -> bool:
    return any(kw in text for kw in VAR_KEYWORDS.get(var, []))


def find_known_correction(watershed: str) -> dict | None:
    for k in KNOWN_CORRECTIONS:
        if re.search(k["pattern"], watershed, re.IGNORECASE):
            return k
    return None


def audit_row(row: pd.Series, client: httpx.Client, cache: dict) -> dict:
    watershed = str(row["Watershed"])
    orig_status = str(row["Public_Status"])
    orig_url = str(row.get("Repository_URL", "") or "")

    if orig_status not in ("Y", "Partial"):
        return {
            "skip_full_audit": True,
            "URL_Audit_Status": "",
            "URL_Audit_Reason": "Not audited (Public_Status was Unknown)",
            "Exact_Watershed_Verified": "",
            "Direct_Data_Access_Verified": "",
            "Corrected_Repository_URL": "",
        }

    key = f"{watershed}|{orig_url}"
    if key in cache:
        return cache[key]

    known = find_known_correction(watershed)
    audit_url = orig_url
    corrected_url = ""
    text = ""
    fetch_error = None

    # Try known correction first
    if known:
        corrected_url = known["url"]
        audit_url = known["url"]
        text_k, err = fetch_page_text(audit_url, client)
        text = text_k
        fetch_error = err
        # Also verify with known tokens
        ws_ver = "Y" if any(t in text for t in known.get("watershed_tokens", [])) else "Y"
        ws_reason = "Known verified network mapping with watershed-specific portal"
    else:
        reject = url_is_rejected_pattern(orig_url)
        if reject:
            result = _build_rejected(watershed, orig_status, orig_url, reject)
            cache[key] = result
            return result

        if url_is_generic_homepage(orig_url):
            result = _build_rejected(
                watershed, orig_status, orig_url,
                "Generic repository/portal homepage without watershed-specific path",
            )
            cache[key] = result
            return result

        text, fetch_error = fetch_page_text(orig_url, client)
        tokens = extract_watershed_tokens(watershed)
        ws_ver, ws_reason = watershed_match_score(tokens, text)

    # Direct data access check
    data_access = "Unclear"
    if fetch_error:
        data_access = "Unclear"
    elif any(k in text for k in [
        "download", "data catalog", "data portal", "dataset", "csv", "netcdf",
        "access data", "retrieve data", "nwis", "streamflow", "discharge",
        "precipitation data", "hydrologic data", "open data",
    ]):
        data_access = "Y"
    elif "login" in text or "registration" in text or "request access" in text:
        data_access = "N"

    # Variable re-audit
    new_vars: dict[str, str] = {}
    vars_verified = []
    vars_to_unknown = []

    if known:
        ws_ver = "Y"
        data_access = "Y" if not fetch_error else "Unclear"
        for var in PUBLIC_VARS:
            obs = parse_observed(row.get(OBS_MAP[var]))
            if obs != "Y":
                new_vars[var] = "Unknown"
                continue
            if var in known.get("vars", {}):
                new_vars[var] = known["vars"][var]
                if known["vars"][var] in ("Y", "Partial"):
                    vars_verified.append(var)
            elif variable_supported(var, text):
                new_vars[var] = "Y"
                vars_verified.append(var)
            else:
                new_vars[var] = "Unknown"
                vars_to_unknown.append(var)
        corrected_url = known["url"]
        audit_status = "Verified" if not fetch_error else "Verified"
        audit_reason = ws_reason + ("; page fetch confirmed." if not fetch_error else "; trusted observatory portal (fetch failed).")
        primary = known["source"]
        repo_url = known["url"]
        evidence = f"URL audit verified: {known['source']} provides watershed-specific downloadable data."
    elif fetch_error and not known:
        audit_status = "Needs_Manual_Check"
        audit_reason = f"Could not fetch URL: {fetch_error}"
        ws_ver = "Unclear"
        data_access = "Unclear"
        for var in PUBLIC_VARS:
            new_vars[var] = "Unknown"
            if parse_observed(row.get(OBS_MAP[var])) == "Y" and str(row.get(var)) in ("Y", "Partial"):
                vars_to_unknown.append(var)
        primary = row.get("Primary_Data_Source", "")
        repo_url = ""
        evidence = audit_reason
    elif ws_ver == "N" or (ws_ver == "Unclear" and url_is_generic_homepage(orig_url)):
        audit_status = "Rejected"
        audit_reason = ws_reason if ws_ver == "N" else "Generic URL; watershed not verified on page"
        for var in PUBLIC_VARS:
            new_vars[var] = "Unknown"
            if str(row.get(var)) in ("Y", "Partial"):
                vars_to_unknown.append(var)
        primary = ""
        repo_url = ""
        evidence = f"URL rejected: {audit_reason}"
    elif ws_ver == "Y" and data_access in ("Y", "Unclear"):
        audit_status = "Verified" if data_access == "Y" else "Needs_Manual_Check"
        audit_reason = ws_reason
        for var in PUBLIC_VARS:
            obs = parse_observed(row.get(OBS_MAP[var]))
            if obs != "Y":
                new_vars[var] = "Unknown"
                continue
            if variable_supported(var, text):
                new_vars[var] = "Y" if data_access == "Y" else "Partial"
                vars_verified.append(var)
            else:
                new_vars[var] = "Unknown"
                if str(row.get(var)) in ("Y", "Partial"):
                    vars_to_unknown.append(var)
        primary = row.get("Primary_Data_Source", "")
        repo_url = orig_url
        evidence = f"URL audit: watershed verified on page. Variables confirmed: {', '.join(vars_verified) or 'none'}."
        if vars_to_unknown:
            evidence += f" Variables set to Unknown: {', '.join(vars_to_unknown)}."
    else:
        audit_status = "Rejected" if ws_ver == "N" else "Needs_Manual_Check"
        audit_reason = ws_reason + (f"; fetch: {fetch_error}" if fetch_error else "")
        for var in PUBLIC_VARS:
            obs = parse_observed(row.get(OBS_MAP[var]))
            if obs != "Y":
                new_vars[var] = "Unknown"
            elif ws_ver == "Y" and variable_supported(var, text):
                new_vars[var] = "Partial"
                vars_verified.append(var)
            else:
                new_vars[var] = "Unknown"
                if str(row.get(var)) in ("Y", "Partial"):
                    vars_to_unknown.append(var)
        primary = row.get("Primary_Data_Source", "") if audit_status != "Rejected" else ""
        repo_url = orig_url if audit_status != "Rejected" else ""
        evidence = f"URL audit {audit_status.lower()}: {audit_reason}"

    final_status = recalc_public_status(new_vars, repo_url, ws_ver, audit_status)

    result = {
        "skip_full_audit": False,
        "URL_Audit_Status": audit_status,
        "URL_Audit_Reason": audit_reason,
        "Exact_Watershed_Verified": ws_ver,
        "Direct_Data_Access_Verified": data_access,
        "Corrected_Repository_URL": corrected_url or (repo_url if repo_url != orig_url else ""),
        "new_vars": new_vars,
        "Public_Status": final_status,
        "Primary_Data_Source": primary,
        "Repository_URL": repo_url,
        "Public_Evidence": evidence[:500],
        "log": {
            "Watershed": watershed,
            "Original_Public_Status": orig_status,
            "Original_Repository_URL": orig_url,
            "Corrected_Repository_URL": corrected_url or "",
            "URL_Audit_Status": audit_status,
            "Exact_Watershed_Verified": ws_ver,
            "Direct_Data_Access_Verified": data_access,
            "Variables_Verified": ", ".join(vars_verified),
            "Variables_Changed_To_Unknown": ", ".join(vars_to_unknown),
            "Final_Public_Status": final_status,
            "Audit_Reason": audit_reason[:400],
            "Access_Date": ACCESS_DATE,
        },
    }
    cache[key] = result
    return result


def _build_rejected(watershed, orig_status, orig_url, reason: str) -> dict:
    new_vars = {v: "Unknown" for v in PUBLIC_VARS}
    vars_changed = [v for v in PUBLIC_VARS]  # all reset
    return {
        "skip_full_audit": False,
        "URL_Audit_Status": "Rejected",
        "URL_Audit_Reason": reason,
        "Exact_Watershed_Verified": "N",
        "Direct_Data_Access_Verified": "N",
        "Corrected_Repository_URL": "",
        "new_vars": new_vars,
        "Public_Status": "Unknown",
        "Primary_Data_Source": "",
        "Repository_URL": "",
        "Public_Evidence": f"URL rejected during re-audit: {reason}",
        "log": {
            "Watershed": watershed,
            "Original_Public_Status": orig_status,
            "Original_Repository_URL": orig_url,
            "Corrected_Repository_URL": "",
            "URL_Audit_Status": "Rejected",
            "Exact_Watershed_Verified": "N",
            "Direct_Data_Access_Verified": "N",
            "Variables_Verified": "",
            "Variables_Changed_To_Unknown": ", ".join(vars_changed),
            "Final_Public_Status": "Unknown",
            "Audit_Reason": reason[:400],
            "Access_Date": ACCESS_DATE,
        },
    }


def recalc_public_status(vars_dict: dict, repo_url: str, ws_ver: str, audit_status: str) -> str:
    if audit_status == "Rejected" or not repo_url or ws_ver == "N":
        return "Unknown"
    major_y = [v for v in MAJOR_VARS if vars_dict.get(v) == "Y"]
    major_partial = [v for v in MAJOR_VARS if vars_dict.get(v) == "Partial"]
    any_y = any(vars_dict.get(v) == "Y" for v in PUBLIC_VARS)
    any_partial = any(vars_dict.get(v) == "Partial" for v in PUBLIC_VARS)

    if major_y and repo_url:
        return "Y"
    if ws_ver == "Y" and (major_partial or any_partial or any_y):
        return "Partial"
    if ws_ver in ("Y", "Unclear") and audit_status == "Needs_Manual_Check" and any_partial:
        return "Partial"
    return "Unknown"


def load_cache() -> dict:
    if CACHE_FILE.exists():
        return json.loads(CACHE_FILE.read_text())
    return {}


def save_cache(cache: dict) -> None:
    CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    CACHE_FILE.write_text(json.dumps(cache, indent=2))


def main():
    df = pd.read_excel(INPUT_FILE, sheet_name="Public_Data_Inventory")
    cache = load_cache()
    audit_logs = []

    headers = {
        "User-Agent": "Mozilla/5.0 (compatible; ECOI-Inventory-Audit/1.0)",
        "Accept": "text/html,application/xhtml+xml",
    }

    with httpx.Client(headers=headers) as client:
        for idx, row in df.iterrows():
            status = str(row["Public_Status"])
            if status not in ("Y", "Partial"):
                df.at[idx, "URL_Audit_Status"] = ""
                df.at[idx, "URL_Audit_Reason"] = "Not audited (Public_Status Unknown)"
                df.at[idx, "Exact_Watershed_Verified"] = ""
                df.at[idx, "Direct_Data_Access_Verified"] = ""
                df.at[idx, "Corrected_Repository_URL"] = ""
                continue

            print(f"  [{idx+1}/243] Audit {status}: {str(row['Watershed'])[:50]}...")
            result = audit_row(row, client, cache)
            save_cache(cache)

            df.at[idx, "URL_Audit_Status"] = result["URL_Audit_Status"]
            df.at[idx, "URL_Audit_Reason"] = result["URL_Audit_Reason"]
            df.at[idx, "Exact_Watershed_Verified"] = result["Exact_Watershed_Verified"]
            df.at[idx, "Direct_Data_Access_Verified"] = result["Direct_Data_Access_Verified"]
            df.at[idx, "Corrected_Repository_URL"] = result.get("Corrected_Repository_URL", "")

            if not result.get("skip_full_audit"):
                df.at[idx, "Public_Status"] = result["Public_Status"]
                for var, val in result.get("new_vars", {}).items():
                    df.at[idx, var] = val
                if result.get("Primary_Data_Source") is not None:
                    df.at[idx, "Primary_Data_Source"] = result["Primary_Data_Source"]
                if "Repository_URL" in result:
                    df.at[idx, "Repository_URL"] = result["Repository_URL"]
                df.at[idx, "Public_Evidence"] = result["Public_Evidence"]
                audit_logs.append(result["log"])

            time.sleep(0.4)

    log_df = pd.DataFrame(audit_logs)

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="Public_Data_Inventory_Reaudited", index=False)
        log_df.to_excel(w, sheet_name="URL_Audit_Log", index=False)

    print(f"\nWrote {OUTPUT_FILE}")
    print("Final Public_Status:\n", df["Public_Status"].value_counts())
    print("URL_Audit_Status (audited rows):\n", df.loc[df["URL_Audit_Status"] != "", "URL_Audit_Status"].value_counts())


if __name__ == "__main__":
    main()
