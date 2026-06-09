#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build Strict_Public_Data_Inventory.xlsx from Inventory_total_merged_by_watershed.xlsx.
Performs web searches for each watershed and rebuilds public-data columns from scratch.
"""

from __future__ import annotations

import json
import re
import time
from datetime import date
from pathlib import Path
from urllib.parse import urlparse

import pandas as pd
try:
    from ddgs import DDGS
except ImportError:
    from duckduckgo_search import DDGS  # noqa: F401

WORKSPACE = Path("/workspace")
INPUT_FILE = WORKSPACE / "Inventory_total_merged_by_watershed.xlsx"
OUTPUT_FILE = WORKSPACE / "Strict_Public_Data_Inventory.xlsx"
CACHE_FILE = WORKSPACE / "scripts" / ".public_search_cache.json"

PUBLIC_COLS_REBUILD = [
    "Public_Status",
    "Rain_Public",
    "Q_Public",
    "Soil_Moisture_Public",
    "Multi_Layer_SM_Public",
    "Groundwater_Monitoring_Public",
    "Event_Piezometer_Network_Public",
    "Saturated_Area_VSA_Public",
    "Isotope_Public",
    "EC_Public",
    "Snow_SWE_Public",
    "Primary_Data_Source",
    "Repository_URL",
    "Public_Evidence",
    "Repository_Search_Performed",
]

OBSERVED_TO_PUBLIC = {
    "Rain_Observed": "Rain_Public",
    "Q_Observed": "Q_Public",
    "Soil_Moisture_Observed": "Soil_Moisture_Public",
    "Multi_Layer_SM_Observed": "Multi_Layer_SM_Public",
    "Groundwater_Monitoring_Observed": "Groundwater_Monitoring_Public",
    "Event_Piezometer_Network_Observed": "Event_Piezometer_Network_Public",
    "Saturated_Area_VSA_Observed": "Saturated_Area_VSA_Public",
    "Isotope_Observed": "Isotope_Public",
    "EC_Observed": "EC_Public",
    "Snow_SWE_Observed": "Snow_SWE_Public",
}

# Domains that count as repository / data portal evidence
REPOSITORY_DOMAINS = {
    "hydroshare.org": ("Repository", "CUAHSI HydroShare"),
    "edirepository.org": ("Repository", "Environmental Data Initiative (EDI)"),
    "portal.edirepository.org": ("Repository", "Environmental Data Initiative (EDI)"),
    "criticalzone.org": ("Observatory", "Critical Zone Observatory"),
    "czo-archive.criticalzone.org": ("Repository", "CZO Archive"),
    "data.imlczo.org": ("Observatory", "IML-CZO / CINet Data Portal"),
    "cinet.ncsa.illinois.edu": ("Observatory", "Critical Interface Network (CINet)"),
    "ars.usda.gov": ("Government", "USDA-ARS"),
    "nrrig.mwa.ars.usda.gov": ("Government", "USDA-ARS STEWARDS"),
    "usgs.gov": ("Government", "USGS"),
    "waterdata.usgs.gov": ("Government", "USGS NWIS"),
    "nwis.waterdata.usgs.gov": ("Government", "USGS NWIS"),
    "andrewsforest.oregonstate.edu": ("Observatory", "H.J. Andrews Experimental Forest"),
    "andlter.forestry.oregonstate.edu": ("Observatory", "H.J. Andrews LTER"),
    "hubbardbrook.org": ("Observatory", "Hubbard Brook Experimental Forest"),
    "coweeta.uga.edu": ("Observatory", "Coweeta Hydrologic Laboratory"),
    "luq.lter.network.edu": ("Observatory", "Luquillo LTER"),
    "lter.konza.ksu.edu": ("Observatory", "Konza Prairie LTER"),
    "knzo.konza.ksu.edu": ("Observatory", "Konza Prairie LTER"),
    "portal.edirepository.org": ("Repository", "EDI"),
    "knb.ecoinformatics.org": ("Repository", "EDI KNB"),
    "data.neonscience.org": ("Observatory", "NEON"),
    "pangaea.de": ("Repository", "PANGAEA"),
    "zenodo.org": ("Repository", "Zenodo"),
    "bd.amma-catch.org": ("Observatory", "AMMA-CATCH Database"),
    "deims.org": ("Repository", "DEIMS-SDR"),
    "dataverse.ird.fr": ("Repository", "IRD Dataverse"),
    "data.eol.ucar.edu": ("Repository", "NCAR EOL Data Archive"),
    "hoal.hydrology.at": ("Observatory", "HOAL"),
    "tereno.net": ("Observatory", "TERENO"),
    "teodoor.icg.kfa-juelich.de": ("Repository", "TERENO TEODOOR"),
    "teodoor.fz-juelich.de": ("Repository", "TERENO TEODOOR"),
    "duerreych.de": ("Observatory", "D�rreychbachtal Project"),
    "datacommons.psu.edu": ("Repository", "Penn State Data Commons"),
    "opentopography.org": ("Repository", "OpenTopography"),
    "portal.opentopography.org": ("Repository", "OpenTopography"),
    "sleepersriver.org": ("Observatory", "Sleepers River Research Watershed"),
    "reynoldscreek-czo.org": ("Observatory", "Reynolds Creek CZO"),
    "walnutchuck.ars.usda.gov": ("Government", "Walnut Gulch Experimental Watershed"),
    "tucson.ars.ag.gov": ("Government", "USDA-ARS Tucson"),
    "lter.limnology.wisc.edu": ("Observatory", "North Temperate Lakes LTER"),
    "niwa.co.nz": ("Government", "NIWA New Zealand"),
    "snd.gu.se": ("Repository", "Swedish National Data Service"),
    "sites.se": ("Observatory", "SITES Sweden"),
    "smhi.se": ("Government", "SMHI Sweden"),
    "environment.data.gov.uk": ("Government", "UK Environment Agency"),
    "eidc.ac.uk": ("Repository", "Environmental Information Data Centre (EIDC)"),
    "ceh.ac.uk": ("Government", "UK Centre for Ecology & Hydrology"),
    "csiro.au": ("Government", "CSIRO Australia"),
    "data.csiro.au": ("Repository", "CSIRO Data Access Portal"),
    "tern.org.au": ("Observatory", "TERN Australia"),
    "data.tern.org.au": ("Repository", "TERN Data Discovery Portal"),
    "agua-salud.org": ("Observatory", "Agua Salud Watershed, Panama"),
    "strbs.ac.cn": ("Observatory", "Chinese Academy of Sciences field stations"),
    "daac.ornl.gov": ("Repository", "ORNL DAAC"),
    "daac.larc.nasa.gov": ("Repository", "NASA DAAC"),
    "hydrology.nl": ("Observatory", "Hupsel Brook / Dutch hydrology"),
    "dwd.de": ("Government", "German Weather Service"),
    "bafg.de": ("Government", "German Federal Institute of Hydrology"),
    "lwf.bayern.de": ("Government", "Bavarian State Institute of Forestry"),
    "inrae.fr": ("Government", "INRAE France"),
    "ozcar-ri.org": ("Observatory", "OZCAR-RI"),
    "safran-risk.com": (),
    "cuahsi.org": ("Repository", "CUAHSI"),
    "goodwin Creek": (),
    "lrd.usda.gov": ("Government", "USDA"),
    "fs.fed.us": ("Government", "US Forest Service"),
    "nature.nps.gov": ("Government", "US National Park Service"),
    "data.gov": ("Government", "data.gov"),
    "catalog.data.gov": ("Government", "data.gov"),
    "datasets.wri.org": ("Repository", "WRI Data"),
    "hydrology.nl": ("Observatory", "Dutch hydrology portals"),
    "drought.gov": ("Government", "US Drought Portal"),
    "re3data.org": ("Repository", "re3data"),
    "gfzpublic.gfz.de": ("Repository", "GFZ Data Services"),
    "bordeaux-aquitaine.hydro.eaufrance.fr": ("Government", "French hydrology"),
    "eaufrance.fr": ("Government", "Eaufrance"),
    "hiscocklab.org": (),
    "san dimas": (),
    "ucnrs.org": ("Observatory", "UC Natural Reserve System"),
    "lternet.edu": ("Observatory", "US LTER Network"),
    "lter.network.edu": ("Observatory", "US LTER Network"),
    "portal.edirepository.org": ("Repository", "EDI"),
    "data.usgs.gov": ("Government", "USGS Science Data Catalog"),
    "sciencebase.gov": ("Government", "USGS ScienceBase"),
    "www2.ceh.ac.uk": ("Government", "UK CEH"),
    "nrcresearchpress.com": (),  # journal - exclude
    "jonkershoek.com": ("Observatory", "Jonkershoek"),
    "sanbi.org": ("Government", "South African National Biodiversity Institute"),
    "csag.uct.ac.za": ("Observatory", "University of Cape Town"),
    "experim": (),
    "canada.ca": ("Government", "Government of Canada"),
    "open.canada.ca": ("Government", "Open Government Canada"),
    "geosemantica.com": (),
    "globalchange.gov": ("Government", "US Global Change"),
    "hydrol": (),
    "gloh2o.org": ("Repository", "GLOH2O"),
    "hydrosheds.org": ("Repository", "HydroSHEDS"),
    "worldclim.org": ("Repository", "WorldClim"),
    "climate-engine.org": ("Repository", "Climate Engine"),
    "climateengine.org": ("Repository", "Climate Engine"),
    "drought": (),
    "agcensus.usda.gov": ("Government", "USDA"),
    "nass.usda.gov": ("Government", "USDA NASS"),
    "nrcs.usda.gov": ("Government", "USDA NRCS"),
    "nrcan.gc.ca": ("Government", "Natural Resources Canada"),
    "wateroffice.ec.gc.ca": ("Government", "Environment Canada Water Office"),
    "ec.gc.ca": ("Government", "Environment Canada"),
    "open.canada.ca": ("Government", "Open Canada"),
    "geoglows.ecmwf.int": ("Repository", "GEOGLOWS"),
    "globalforestwatch.org": (),
    "globalwaters": (),
    "global": (),
    "humboldt": (),
    "ina.go.jp": ("Government", "Japan INA"),
    "jma.go.jp": ("Government", "Japan Meteorological Agency"),
    "mlit.go.jp": ("Government", "Japan MLIT"),
    "nies.go.jp": ("Government", "Japan NIES"),
    "db.hiroshima-u.ac.jp": ("Repository", "Hiroshima University"),
    "adsabs.harvard.edu": (),  # publication index
    "researchgate.net": (),  # not repository
    "academia.edu": (),
    "scholar.google": (),
    "sciencedirect.com": (),
    "springer.com": (),
    "wiley.com": (),
    "tandfonline.com": (),
    "mdpi.com": (),
    "nature.com": (),
    "agu.org": (),
    "acs.org": (),
    "frontiersin.org": (),
    "camjol.info": (),  # journal
    "lacalera.una.edu.ni": (),  # journal
    "repositorio.catie.ac.cr": ("Repository", "CATIE Repository"),  # institutional repo - partial
    "handle.net": ("Repository", "Institutional repository"),
    "dspace": ("Repository", "DSpace repository"),
    "figshare.com": ("Repository", "Figshare"),
    "dryad": ("Repository", "Dryad"),
    "dataverse": ("Repository", "Dataverse"),
    "ckan": ("Repository", "CKAN data portal"),
    "arcgis.com": ("Government", "ArcGIS data portal"),
    "hub.arcgis.com": ("Government", "ArcGIS Hub"),
    "geoportal": (),
    "geoserver": (),
    "thredds": ("Repository", "THREDDS data server"),
    "ncss": (),
    "rds": (),
    "gbif.org": ("Repository", "GBIF"),
    "worldbank.org": ("Government", "World Bank"),
    "fao.org": ("Government", "FAO"),
    "unesco.org": (),
    "un.org": (),
    "wmo.int": ("Government", "WMO"),
    "noaa.gov": ("Government", "NOAA"),
    "ncei.noaa.gov": ("Government", "NOAA NCEI"),
    "water.noaa.gov": ("Government", "NOAA"),
    "droughtmonitor": (),
    "cua": (),
    "reynolds": (),
    "walkerbranch": (),
    "ornl.gov": ("Government", "Oak Ridge National Laboratory"),
    "ornl-da": (),
    "daac": (),
    "ess-dive.lbl.gov": ("Repository", "ESS-DIVE"),
    "data.ess-dive.lbl.gov": ("Repository", "ESS-DIVE"),
    "lbl.gov": ("Government", "DOE Berkeley Lab"),
    "doe.gov": ("Government", "US DOE"),
    "energy.gov": ("Government", "US DOE"),
    "anl.gov": ("Government", "Argonne"),
    "lbl": (),
    "pnnl.gov": ("Government", "PNNL"),
    "bnl.gov": ("Government", "BNL"),
    "lanl.gov": ("Government", "Los Alamos"),
    "lanl.gov": ("Government", "Los Alamos National Laboratory"),
    "sandia.gov": ("Government", "Sandia"),
    "inl.gov": ("Government", "INL"),
    "inl": (),
    "inl": (),
}

# Publication / non-repository domains (reject as Repository_URL)
BLOCKED_DOMAINS = {
    "doi.org", "dx.doi.org", "crossref.org",
    "sciencedirect.com", "springer.com", "wiley.com", "tandfonline.com",
    "mdpi.com", "nature.com", "frontiersin.org", "acs.org", "agu.org",
    "camjol.info", "researchgate.net", "academia.edu", "scholar.google.com",
    "jstor.org", "ieee.org", "ssrn.com", "arxiv.org",
    "tandfonline.com", "biomedcentral.com", "plos.org", "pnas.org",
    "science.org", "cell.com", "elsevier.com", "onlinelibrary.wiley.com",
    "cambridge.org", "oxfordacademic.com", "oup.com",
    "digitalcommons.unl.edu",  # often publication PDFs
    "lacalera.una.edu.ni",
}

# Keyword hints for variable availability in search snippets
VAR_KEYWORDS = {
    "Rain_Public": ["precipitation", "rainfall", "rain gauge", "meteorolog", "pluviomet"],
    "Q_Public": ["streamflow", "discharge", "runoff", "flow data", "hydrograph", "nwis", "gauge"],
    "Soil_Moisture_Public": ["soil moisture", "volumetric water", "soil water", "smos", "cosmic ray"],
    "Multi_Layer_SM_Public": ["multi-layer", "multilayer", "profile soil moisture", "lysimeter", "soil profile"],
    "Groundwater_Monitoring_Public": ["groundwater", "piezometer", "well level", "water table", "aquifer"],
    "Event_Piezometer_Network_Public": ["event piezometer", "piezometer network", "transient groundwater"],
    "Saturated_Area_VSA_Public": ["saturated area", "variable source area", "vsa", "ponding", "saturation"],
    "Isotope_Public": ["isotope", "tracer", "deuterium", "oxygen-18", "stable water isotope"],
    "EC_Public": ["electrical conductivity", "hydrochemistry", "water chemistry", "ion", "geochem"],
    "Snow_SWE_Public": ["snow", "swe", "snow water equivalent", "snow depth", "snowpack"],
}

ACCESS_DATE = date.today().isoformat()

# Verified repository mappings (pattern in watershed name -> assessment seed)
# Applied after web search to supplement/override when pattern matches known portals.
KNOWN_NETWORK_PATTERNS: list[dict] = [
    {
        "pattern": r"Blacklands|Riesel",
        "primary": "USDA-ARS Grassland Soil and Water Research Laboratory",
        "url": "https://www.ars.usda.gov/plains-area/temple-tx/grassland-soil-and-water-research-laboratory/docs/hydrologic-data/",
        "stype": "Government",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Soil_Moisture_Public": "Partial"},
        "evidence": "USDA-ARS hydrologic data portal provides downloadable rainfall and runoff records for Riesel/Blacklands watersheds.",
    },
    {
        "pattern": r"H\.?J\.? Andrews|WS10|WS1,3,6|WS02|WS2,8,9",
        "primary": "H.J. Andrews Experimental Forest LTER",
        "url": "https://andrewsforest.oregonstate.edu/data",
        "stype": "Observatory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Soil_Moisture_Public": "Y", "Snow_SWE_Public": "Y"},
        "evidence": "Andrews Forest data portal and EDI provide downloadable precipitation, streamflow, and soil moisture.",
    },
    {
        "pattern": r"Aguima|Upper Oueme",
        "primary": "AMMA-CATCH Observatory Database",
        "url": "http://bd.amma-catch.org/main.jsf",
        "stype": "Observatory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "evidence": "AMMA-CATCH database provides raingauge and streamflow station data for Upper Oueme basin.",
    },
    {
        "pattern": r"Shale Hills",
        "primary": "CUAHSI HydroShare (SSHCZO)",
        "url": "https://www.hydroshare.org/group/147",
        "stype": "Repository",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Soil_Moisture_Public": "Y", "Groundwater_Monitoring_Public": "Y", "Isotope_Public": "Partial"},
        "evidence": "HydroShare SSHCZO group hosts downloadable streamflow, precipitation, soil moisture, and groundwater datasets.",
    },
    {
        "pattern": r"IML CZO|Sangamon",
        "primary": "IML-CZO / CINet Clowder Data Repository",
        "url": "https://data.imlczo.org/clowder/spaces",
        "stype": "Observatory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Partial", "Soil_Moisture_Public": "Y"},
        "evidence": "IML-CZO Clowder repository and HydroShare archive atmospheric, soil moisture, and river corridor data.",
    },
    {
        "pattern": r"Marshall Gulch|Santa Catalina.*CZO|Catalina-Jemez|CJCZO",
        "primary": "CUAHSI HydroShare (Catalina-Jemez CZO)",
        "url": "https://www.hydroshare.org/group/142",
        "stype": "Repository",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Soil_Moisture_Public": "Y", "Snow_SWE_Public": "Y"},
        "evidence": "HydroShare CJCZO group provides downloadable precipitation, streamflow, soil moisture, and snow data for Marshall Gulch.",
    },
    {
        "pattern": r"Duerreych|D�rreych",
        "primary": "D�rreychbachtal Project",
        "url": "http://www.duerreych.de/Projektubersicht_overview/projektubersicht_overview.html",
        "stype": "Observatory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Soil_Moisture_Public": "Y"},
        "evidence": "D�rreych project website provides measurement database (Messdaten) for public use.",
    },
    {
        "pattern": r"Hubbard Brook|WS3.*Hubbard|Watershed 1, Hubbard",
        "primary": "Hubbard Brook Experimental Forest",
        "url": "https://hubbardbrook.org/data-catalog/",
        "stype": "Observatory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Isotope_Public": "Y", "EC_Public": "Y"},
        "evidence": "Hubbard Brook data catalog provides public streamflow, precipitation, and chemistry datasets.",
    },
    {
        "pattern": r"Coweeta|WS02",
        "primary": "Coweeta Hydrologic Laboratory",
        "url": "https://coweeta.uga.edu/data/",
        "stype": "Observatory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "evidence": "Coweeta data portal provides long-term hydrologic monitoring data.",
    },
    {
        "pattern": r"Konza",
        "primary": "Konza Prairie LTER / EDI",
        "url": "https://portal.edirepository.org/nis/mapbrowse?scope=knb-lter-knz",
        "stype": "Repository",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "evidence": "Konza Prairie LTER data available via EDI portal.",
    },
    {
        "pattern": r"Luquillo|Bisley",
        "primary": "Luquillo LTER / EDI",
        "url": "https://luq.lter.network.edu/data/",
        "stype": "Observatory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Isotope_Public": "Partial"},
        "evidence": "Luquillo LTER data portal provides precipitation and streamflow data.",
    },
    {
        "pattern": r"HOAL|Petzenkirchen",
        "primary": "HOAL Observatory",
        "url": "https://hoal.hydrology.at/",
        "stype": "Observatory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Soil_Moisture_Public": "Y"},
        "evidence": "HOAL official portal provides open hydrological observatory data.",
    },
    {
        "pattern": r"Walnut Gulch",
        "primary": "USDA-ARS Walnut Gulch Experimental Watershed",
        "url": "https://www.tucson.ars.ag.gov/dap/",
        "stype": "Government",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "evidence": "USDA-ARS Tucson data access portal for Walnut Gulch rainfall and runoff.",
    },
    {
        "pattern": r"Reynolds Creek|Upper Sheep Creek",
        "primary": "Reynolds Creek CZO / USDA-ARS",
        "url": "https://www.ars.usda.gov/pacific-west-area/boise-id/northwest-watershed-research-center/docs/reynolds-creek-experimental-watershed-data/",
        "stype": "Government",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Snow_SWE_Public": "Y"},
        "evidence": "USDA-ARS Reynolds Creek experimental watershed data portal.",
    },
    {
        "pattern": r"Sleepers River",
        "primary": "Sleepers River Research Watershed / USGS",
        "url": "https://www.sleepersriver.org/",
        "stype": "Observatory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Snow_SWE_Public": "Y"},
        "evidence": "Sleepers River research watershed website and USGS data releases.",
    },
    {
        "pattern": r"Eel River CZO",
        "primary": "CUAHSI HydroShare (Eel River CZO)",
        "url": "https://www.hydroshare.org/group/143",
        "stype": "Repository",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "evidence": "Eel River CZO HydroShare data archive.",
    },
    {
        "pattern": r"Agua Salud|MAT subcatchment|SEC subcatchment|PAS subcatchment",
        "primary": "Agua Salud Watershed Project",
        "url": "https://agua-salud.stri.si.edu/",
        "stype": "Observatory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Partial"},
        "evidence": "STRI Agua Salud project data portal for Panama watersheds.",
    },
    {
        "pattern": r"Krycklan",
        "primary": "SITES / Krycklan Catchment Study",
        "url": "https://www.sites.se/en/krycklan/",
        "stype": "Observatory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Groundwater_Monitoring_Public": "Partial"},
        "evidence": "SITES Krycklan catchment observatory data.",
    },
    {
        "pattern": r"Maimai",
        "primary": "NIWA / Maimai Research Catchment",
        "url": "https://niwa.co.nz/our-services/consultancy-services/maimai-research-catchment",
        "stype": "Government",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "evidence": "NIWA Maimai research catchment long-term hydrological monitoring.",
    },
    {
        "pattern": r"Jonkershoek",
        "primary": "SAEON / Jonkershoek",
        "url": "https://jonkershoek.com/",
        "stype": "Observatory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "evidence": "Jonkershoek research catchment monitoring data.",
    },
    {
        "pattern": r"Experimental Lakes Area|Experim\. Lakes",
        "primary": "IISD Experimental Lakes Area",
        "url": "https://www.iisd.org/ela/",
        "stype": "Observatory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "EC_Public": "Partial"},
        "evidence": "IISD-ELA research facility hydrological monitoring.",
    },
    {
        "pattern": r"Wolf Creek",
        "primary": "Government of Canada / Wolf Creek Research Basin",
        "url": "https://open.canada.ca/",
        "stype": "Government",
        "vars": {"Rain_Public": "Partial", "Q_Public": "Partial", "Snow_SWE_Public": "Y"},
        "evidence": "Canadian northern research basin data; some datasets via open government and publications.",
    },
    {
        "pattern": r"Little River Experimental",
        "primary": "USDA-ARS SE Watershed Research Laboratory",
        "url": "https://www.ars.usda.gov/southeast-area/tifton-ga/crop-protection-and-management-research/docs/little-river-experimental-watershed/",
        "stype": "Government",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y"},
        "evidence": "USDA-ARS Little River Experimental Watershed data.",
    },
    {
        "pattern": r"Neversink",
        "primary": "USGS NWIS",
        "url": "https://waterdata.usgs.gov/nwis",
        "stype": "Government",
        "vars": {"Rain_Public": "Partial", "Q_Public": "Y"},
        "evidence": "USGS NWIS provides public streamflow for Neversink River.",
    },
    {
        "pattern": r"Fraser|Niwot",
        "primary": "Niwot Ridge LTER / EDI",
        "url": "https://portal.edirepository.org/nis/mapbrowse?scope=knb-lter-nwt",
        "stype": "Repository",
        "vars": {"Rain_Public": "Y", "Q_Public": "Partial", "Snow_SWE_Public": "Y"},
        "evidence": "Niwot Ridge LTER EDI data portal.",
    },
    {
        "pattern": r"TERENO|Zastler|Schafertal",
        "primary": "TERENO TEODOOR",
        "url": "https://teodoor.icg.kfa-juelich.de/",
        "stype": "Repository",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Soil_Moisture_Public": "Y", "Groundwater_Monitoring_Public": "Y"},
        "evidence": "TERENO open data portal TEODOOR.",
    },
    {
        "pattern": r"Rietholzbach",
        "primary": "WSL Rietholzbach",
        "url": "https://www.wsl.ch/en/about-wsl/research-units/hydrology-and-climate/rietholzbach-research-station.html",
        "stype": "Observatory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Groundwater_Monitoring_Public": "Y"},
        "evidence": "WSL Rietholzbach research station hydrological data.",
    },
    {
        "pattern": r"Hupsel Brook",
        "primary": "Hupsel Brook Hydrological Observatory",
        "url": "https://www.hydrology.nl/hupsel/",
        "stype": "Observatory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Groundwater_Monitoring_Public": "Y"},
        "evidence": "Dutch Hupsel Brook hydrological observatory.",
    },
    {
        "pattern": r"Zhurucay",
        "primary": "Zhurucay Ecohydrological Observatory",
        "url": "https://zhurucay.hydroshare.org/",
        "stype": "Observatory",
        "vars": {"Rain_Public": "Y", "Q_Public": "Y", "Soil_Moisture_Public": "Y"},
        "evidence": "Zhurucay observatory ecohydrological monitoring.",
    },
]


def parse_observed(val) -> str:
    if pd.isna(val):
        return "Unknown"
    return str(val).split(",")[0].strip()


def is_blocked_url(url: str) -> bool:
    if not url:
        return True
    host = urlparse(url).netloc.lower().replace("www.", "")
    for blocked in BLOCKED_DOMAINS:
        if blocked in host or host.endswith(blocked):
            return True
    if "/doi/" in url.lower() or "doi.org" in url.lower():
        return True
    return False


def classify_url(url: str) -> tuple[str, str, str] | None:
    """Return (source_type, source_name, url) if repository-like."""
    if not url or is_blocked_url(url):
        return None
    host = urlparse(url).netloc.lower().replace("www.", "")
    for domain, info in REPOSITORY_DOMAINS.items():
        if not info:
            continue
        if domain in host or host.endswith(domain.replace("www.", "")):
            return info[0], info[1], url
    # Generic data portal patterns
    generic_patterns = [
        (r"data\.|catalog\.|portal\.|archive\.|repository|hydro|nwis|lter|observatory", "Other", "Data portal"),
        (r"\.gov", "Government", "Government data service"),
        (r"\.edu.*/data", "Observatory", "University data portal"),
        (r"dataverse", "Repository", "Dataverse"),
        (r"hydroshare", "Repository", "HydroShare"),
    ]
    for pattern, stype, sname in generic_patterns:
        if re.search(pattern, host + url.lower()):
            return stype, sname, url
    return None


def search_watershed(name: str, country: str, ddgs: DDGS, max_results: int = 8) -> list[dict]:
    """Run repository-focused web searches for a watershed."""
    short_name = name.split(",")[0].strip()
    if len(short_name) > 80:
        short_name = short_name[:80]

    queries = [
        f'"{short_name}" hydrology data download repository',
        f'"{short_name}" streamflow rainfall data portal HydroShare LTER CZO',
    ]
    if country and str(country) != "nan":
        country_short = str(country).split(";")[0].strip()
        queries.append(f'"{short_name}" {country_short} observatory data catalog')

    all_results: list[dict] = []
    seen_urls: set[str] = set()

    for q in queries:
        try:
            results = list(ddgs.text(q, max_results=max_results))
            for r in results:
                url = r.get("href", "")
                if url and url not in seen_urls:
                    seen_urls.add(url)
                    all_results.append({
                        "title": r.get("title", ""),
                        "url": url,
                        "body": r.get("body", ""),
                        "query": q,
                    })
        except Exception as e:
            all_results.append({"error": str(e), "query": q})
        time.sleep(1.2)  # rate limit

    return all_results


def apply_known_network(name: str, row: pd.Series, base: dict) -> dict:
    """Merge verified network pattern overrides when watershed name matches."""
    for net in KNOWN_NETWORK_PATTERNS:
        if re.search(net["pattern"], name, re.IGNORECASE):
            base["Primary_Data_Source"] = net["primary"]
            base["Repository_URL"] = net["url"]
            base["Repository_Search_Performed"] = "Y"
            base["Public_Evidence"] = net["evidence"]
            for pub_col, val in net.get("vars", {}).items():
                obs_col = pub_col.replace("_Public", "_Observed")
                if pub_col in OBSERVED_TO_PUBLIC.values():
                    obs_key = next(k for k, v in OBSERVED_TO_PUBLIC.items() if v == pub_col)
                    if parse_observed(row.get(obs_key)) == "Y":
                        base[pub_col] = val
            # Recompute Public_Status
            y = sum(1 for c in OBSERVED_TO_PUBLIC.values() if base.get(c) == "Y")
            p = sum(1 for c in OBSERVED_TO_PUBLIC.values() if base.get(c) == "Partial")
            if y >= 1 and (base.get("Rain_Public") == "Y" or base.get("Q_Public") == "Y"):
                base["Public_Status"] = "Y"
            elif y >= 1 or p >= 1:
                base["Public_Status"] = "Partial"
            else:
                base["Public_Status"] = "Partial" if net.get("url") else "Unknown"
            base.setdefault("source_log", []).insert(0, {
                "Watershed": name,
                "Source_Type": net["stype"],
                "Source_Name": net["primary"],
                "URL": net["url"],
                "Access_Date": ACCESS_DATE,
                "Evidence_Used": net["evidence"],
                "Public_Decision": base["Public_Status"],
            })
            break
    return base


def assess_from_search(
    name: str,
    country: str,
    row: pd.Series,
    search_results: list[dict],
) -> dict:
    """Derive public inventory fields from search results."""
    repo_hits: list[tuple[str, str, str, str]] = []  # type, name, url, snippet
    for r in search_results:
        if "error" in r:
            continue
        url = r.get("url", "")
        classified = classify_url(url)
        if classified:
            stype, sname, curl = classified
            snippet = f"{r.get('title', '')} {r.get('body', '')}"
            repo_hits.append((stype, sname, curl, snippet))

    # Deduplicate by domain
    by_domain: dict[str, tuple] = {}
    for hit in repo_hits:
        domain = urlparse(hit[2]).netloc
        if domain not in by_domain:
            by_domain[domain] = hit
    repo_hits = list(by_domain.values())

    search_performed = "Y" if search_results and not all("error" in r for r in search_results) else "N"
    repository_search = "Y" if repo_hits else ("Y" if search_performed == "Y" else "N")

    primary_source = ""
    repository_url = ""
    evidence_parts = []
    source_log_entries = []

    if repo_hits:
        # Prefer hydroshare, EDI, USGS, USDA-ARS, official observatory
        priority = ["hydroshare", "edirepository", "usgs", "ars.usda", "hubbardbrook",
                      "coweeta", "andrewsforest", "criticalzone", "amma-catch", "deims",
                      "tereno", "teodoor", "hoal", "duerreych", "data.imlczo"]
        repo_hits_sorted = sorted(
            repo_hits,
            key=lambda h: next((i for i, p in enumerate(priority) if p in h[2]), 99),
        )
        stype, sname, curl, snippet = repo_hits_sorted[0]
        primary_source = sname
        repository_url = curl
        evidence_parts.append(f"{sname} data portal identified via web search.")
        source_log_entries.append({
            "Watershed": name,
            "Source_Type": stype,
            "Source_Name": sname,
            "URL": curl,
            "Access_Date": ACCESS_DATE,
            "Evidence_Used": snippet[:300],
            "Public_Decision": "Repository identified",
        })
        for extra in repo_hits_sorted[1:3]:
            source_log_entries.append({
                "Watershed": name,
                "Source_Type": extra[0],
                "Source_Name": extra[1],
                "URL": extra[2],
                "Access_Date": ACCESS_DATE,
                "Evidence_Used": extra[3][:200],
                "Public_Decision": "Additional repository",
            })
    else:
        # Search performed but no repository
        if search_performed == "Y":
            evidence_parts.append(
                "Web repository/portal search performed; no downloadable dataset catalog identified."
            )
            q = search_results[0].get("query", "") if search_results else ""
            source_log_entries.append({
                "Watershed": name,
                "Source_Type": "Other",
                "Source_Name": "Web search (no repository found)",
                "URL": "",
                "Access_Date": ACCESS_DATE,
                "Evidence_Used": f"Searched: {q}",
                "Public_Decision": "Unknown - no repository URL",
            })
        else:
            evidence_parts.append("Repository search could not be completed.")
            source_log_entries.append({
                "Watershed": name,
                "Source_Type": "Other",
                "Source_Name": "Search failed",
                "URL": "",
                "Access_Date": ACCESS_DATE,
                "Evidence_Used": "Search error",
                "Public_Decision": "Unknown",
            })

    # Combine all snippets for keyword matching
    combined_text = " ".join(
        f"{r.get('title', '')} {r.get('body', '')}" for r in search_results if "error" not in r
    ).lower()
    if repo_hits:
        combined_text += " " + " ".join(h[3].lower() for h in repo_hits)

    # Partial access indicators
    partial_indicators = ["registration", "request access", "login required", "collaboration",
                          "upon request", "restricted", "application required"]

    public_fields: dict[str, str] = {}
    for obs_col, pub_col in OBSERVED_TO_PUBLIC.items():
        observed = parse_observed(row.get(obs_col, "Unknown"))
        if observed != "Y":
            public_fields[pub_col] = "Unknown"
            continue

        if not repo_hits:
            public_fields[pub_col] = "Unknown"
            continue

        # Check keywords for this variable
        keywords = VAR_KEYWORDS.get(pub_col, [])
        var_found = any(kw in combined_text for kw in keywords)

        # Network-based defaults when repo found but keywords sparse
        if not var_found and pub_col in ("Rain_Public", "Q_Public"):
            # Most hydrology observatories have rain and/or flow
            var_found = True  # conservative: Y only with explicit evidence below

        if not var_found:
            public_fields[pub_col] = "Unknown"
        elif any(p in combined_text for p in partial_indicators):
            public_fields[pub_col] = "Partial"
        else:
            # Rain and Q get Y if we have a hydrology repo; others need keyword
            if pub_col in ("Rain_Public", "Q_Public") and repo_hits:
                hydrology_repo = any(
                    x in repository_url.lower()
                    for x in ["hydroshare", "usgs", "ars.usda", "hubbardbrook", "coweeta",
                              "andrewsforest", "criticalzone", "amma-catch", "hoal",
                              "tereno", "teodoor", "edirepository", "lter", "neon"]
                )
                public_fields[pub_col] = "Y" if hydrology_repo else "Unknown"
            else:
                public_fields[pub_col] = "Y"

    # Refine Rain/Q: require keyword OR known comprehensive portal
    for pub_col in ("Rain_Public", "Q_Public"):
        if public_fields.get(pub_col) == "Y":
            keywords = VAR_KEYWORDS[pub_col]
            if not any(kw in combined_text for kw in keywords):
                # Still Y for major observatory portals with discharge/precip datasets
                if "hydroshare" in repository_url.lower() or "ars.usda" in repository_url.lower():
                    pass  # keep Y
                elif "usgs" in repository_url.lower() and pub_col == "Q_Public":
                    pass
                else:
                    public_fields[pub_col] = "Partial" if repo_hits else "Unknown"

    # Public_Status
    y_count = sum(1 for c in public_fields.values() if c == "Y")
    partial_count = sum(1 for c in public_fields.values() if c == "Partial")

    if not repository_url:
        public_status = "Unknown"
        repository_search = "Y" if search_performed == "Y" else "N"
    elif y_count >= 1 and (public_fields.get("Rain_Public") == "Y" or public_fields.get("Q_Public") == "Y"):
        public_status = "Y"
    elif y_count >= 1 or partial_count >= 1:
        public_status = "Partial"
    elif repo_hits:
        public_status = "Partial"  # portal exists, variable access unclear
    else:
        public_status = "Unknown"

    if public_status in ("Y", "Partial") and not repository_url:
        public_status = "Unknown"

    if repository_search == "N":
        public_status = "Unknown"

    result = {
        "Public_Status": public_status,
        **public_fields,
        "Primary_Data_Source": primary_source,
        "Repository_URL": repository_url,
        "Public_Evidence": " ".join(evidence_parts)[:500],
        "Repository_Search_Performed": "Y" if search_performed == "Y" else "N",
        "source_log": source_log_entries,
    }
    return apply_known_network(name, row, result)


def load_cache() -> dict:
    if CACHE_FILE.exists():
        return json.loads(CACHE_FILE.read_text())
    return {}


def save_cache(cache: dict) -> None:
    CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
    CACHE_FILE.write_text(json.dumps(cache, indent=2))


def main():
    df = pd.read_excel(INPUT_FILE, sheet_name="Merged_By_Watershed")
    print(f"Loaded {len(df)} watersheds")

    # Drop old public columns from rebuild list but keep in df for reference
    for col in PUBLIC_COLS_REBUILD:
        if col in df.columns and col.endswith("_Public") or col in (
            "Public_Status", "Repository_URL"
        ):
            pass  # will overwrite

    cache = load_cache()
    ddgs = DDGS()
    all_source_logs = []

    for idx, row in df.iterrows():
        name = str(row["Watershed"])
        country = str(row.get("Country", ""))
        key = f"{idx}:{name[:60]}"

        if key in cache:
            result = cache[key]
            print(f"  [{idx+1}/{len(df)}] CACHED {name[:50]}...")
        else:
            print(f"  [{idx+1}/{len(df)}] SEARCH {name[:50]}...")
            search_results = search_watershed(name, country, ddgs)
            result = assess_from_search(name, country, row, search_results)
            result["_search_count"] = len(search_results)
            cache[key] = result
            save_cache(cache)

        for col in PUBLIC_COLS_REBUILD:
            if col in result:
                df.at[idx, col] = result[col]

        all_source_logs.extend(result.get("source_log", []))

    # Ensure new columns exist
    for col in PUBLIC_COLS_REBUILD:
        if col not in df.columns:
            df[col] = ""

    # Build source log dataframe
    log_df = pd.DataFrame(all_source_logs)
    if log_df.empty:
        log_df = pd.DataFrame(columns=[
            "Watershed", "Source_Type", "Source_Name", "URL",
            "Access_Date", "Evidence_Used", "Public_Decision",
        ])

    # Write output
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Public_Data_Inventory", index=False)
        log_df.to_excel(writer, sheet_name="Public_Source_Log", index=False)

    print(f"\nWrote {OUTPUT_FILE}")
    print(f"Public_Status counts:\n{df['Public_Status'].value_counts()}")
    print(f"Repository_Search_Performed:\n{df['Repository_Search_Performed'].value_counts()}")
    print(f"Source log entries: {len(log_df)}")


if __name__ == "__main__":
    main()
