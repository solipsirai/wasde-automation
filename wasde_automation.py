#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""USDA WASDE -> five source Excel workbooks (Microsoft Excel on Windows).

Free stack: Python stdlib for retrieval/parsing, pywin32 to preserve Excel templates.
Release-day poll is supplied by Windows Task Scheduler (see README_TR.md).

Fail closed: does not create a monthly sheet unless ALL required USDA metrics
and ALL requested countries/regions are present and parseable.
"""
from __future__ import annotations

import argparse
import calendar
from contextlib import contextmanager
from datetime import date, datetime
from decimal import Decimal
import hashlib
from html.parser import HTMLParser
import json
import logging
import os
from pathlib import Path
import re
import shutil
import sys
import time
from urllib.parse import urljoin, urlsplit
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo

HERE = Path(__file__).resolve().parent
STATE_FILE = HERE / "runtime" / "wasde_state.json"
LOG_FILE = HERE / "runtime" / "wasde_automation.log"
TIMEOUT_SECONDS = 18
ARCHIVE = "https://esmis.nal.usda.gov/publication/world-agricultural-supply-and-demand-estimates"
RELEASE_DATES = HERE / "release-dates.json"

ABBR = {i: calendar.month_abbr[i] for i in range(1, 13)}

# The templates' exact country/region order. This is the contract, NOT a
# guess based on the number of rows USDA happens to publish next month.
WHEAT_REGIONS = [
    "Argentina", "Australia", "Canada", "European Union", "Russia", "Ukraine",
    "Major Importers", "Brazil", "China", "Japan", "N. Africa", "Nigeria",
    "Sel. Mideast", "Southeast Asia", "India", "Kazakhstan", "World",
]
CORN_REGIONS = [
    "World", "World Less China", "United States", "Argentina", "Brazil",
    "Ukraine", "Russia", "Egypt", "European Union", "Japan", "Mexico",
    "Southeast Asia", "South Korea", "South Africa", "Canada", "China",
]
WHEAT_US_ROWS = {
    6: "areaplanted", 7: "areaharvested", 9: "yieldperharvestedacre",
    11: "beginningstocks", 12: "production", 13: "imports", 14: "supplytotal",
    15: "food", 16: "seed", 17: "feedandresidual", 18: "domestictotal",
    19: "exports", 20: "usetotal", 21: "endingstocks", 22: "avgfarmprice",
}
CORN_US_ROWS = {
    6: "areaplanted", 7: "areaharvested", 9: "yieldperharvestedacre",
    11: "beginningstocks", 12: "production", 13: "imports", 14: "supplytotal",
    15: "feedandresidual", 16: "foodseedindustrial", 17: "ethanolbyproducts",
    18: "domestictotal", 19: "exports", 20: "usetotal", 21: "endingstocks",
    22: "avgfarmprice",
}

MONTH_NAMES = {i: calendar.month_name[i] for i in range(1, 13)}
# Field tails: 4 US values, 7 world values; parse with Decimal for accuracy.
NUMBER = r"-?\d+(?:,\d{3})*(?:\.\d+)?"
RE_US = re.compile(rf"^(?P<label>.*?)\s{{2,}}(?P<values>{NUMBER}(?:\s+{NUMBER}){{3}})\s*$")
RE_WORLD = re.compile(rf"^(?P<label>.*?)\s{{2,}}(?P<values>{NUMBER}(?:\s+{NUMBER}){{6}})\s*$")
RE_FOOT = re.compile(r"\s+\d+/\s*$")
RE_SEASON = re.compile(r"\b(20\d{2}/\d{2})\s+Proj\.?", re.I)


def norm(value: str) -> str:
    """Normalize field/country labels; remove USDA footnote tokens."""
    return re.sub(r"[^a-z0-9]+", "", RE_FOOT.sub("", value).lower())


def parse_values(raw: str, count: int) -> tuple[Decimal, ...]:
    values = tuple(Decimal(v.replace(",", "")) for v in raw.split())
    if len(values) != count:
        raise ValueError(f"Expected {count} metrics but got {values}")
    return values


def section_between(full: str, begin: str, end: str) -> str:
    """Use line-level section headings to avoid incidental text mentions."""
    start = re.search(r"(?m)^\s*" + re.escape(begin) + r"\s*(?:\d+/|\(Contd\.\))?.*$", full)
    if not start:
        raise ValueError(f"Missing USDA section: {begin}")
    finish = re.search(r"(?m)^\s*" + re.escape(end) + r"\b.*$", full[start.end():])
    if not finish:
        raise ValueError(f"Missing boundary after {begin}: {end}")
    return full[start.end():start.end()+finish.start()]


def parse_us(full: str, crop: str) -> tuple[dict[str, tuple[Decimal, ...]], list[str]]:
    if crop == "wheat":
        body = section_between(full, "U.S. Wheat Supply and Use", "U.S. Wheat by Class")
    else:
        body = section_between(full, "U.S. Feed Grain and Corn Supply and Use", "U.S. Sorghum, Barley, and Oats Supply and Use")
    year_line = next((line for line in body.splitlines() if len(re.findall(r"20\d{2}/\d{2}", line)) >= 3), "")
    years = re.findall(r"20\d{2}/\d{2}", year_line)
    if crop == "corn":
        corn_marker = re.search(r"(?m)^\s*CORN\s*$", body)
        if not corn_marker:
            raise ValueError("US feed grains table has no CORN block")
        body = body[corn_marker.end():]
    if len(years) != 4:
        raise ValueError(f"Expected four marketing-year columns in US {crop}, got {years}")
    out: dict[str, tuple[Decimal, ...]] = {}
    for line in body.splitlines():
        m = RE_US.match(line)
        if not m:
            continue
        label = norm(m.group("label"))
        if label.startswith("avgfarmprice"):
            label = "avgfarmprice"
        elif label.startswith("foodseedindustrial"):
            label = "foodseedindustrial"
        elif label.startswith("ethanolbyproducts"):
            label = "ethanolbyproducts"
        if label:
            out[label] = parse_values(m.group("values"), 4)
    requested = (WHEAT_US_ROWS if crop == "wheat" else CORN_US_ROWS)
    missing = [key for key in requested.values() if key not in out]
    if missing:
        raise ValueError(f"US {crop} metrics missing: {missing}")
    return out, years


def parse_world(full: str, crop: str, previous: str, current: str) -> tuple[dict[str, dict[str, tuple[Decimal, ...]]], str]:
    if crop == "wheat":
        body = section_between(full, "World Wheat Supply and Use", "World Coarse Grain Supply and Use")
    else:
        body = section_between(full, "World Corn Supply and Use", "World Rice Supply and Use")
    season_matches = list(RE_SEASON.finditer(body))
    if not season_matches:
        raise ValueError(f"World {crop} report lacks projected marketing year")
    latest = season_matches[-1]
    season = latest.group(1)
    projected = body[latest.end():]
    out: dict[str, dict[str, tuple[Decimal, ...]]] = {}
    region = None
    for line in projected.splitlines():
        m = RE_WORLD.match(line)
        if m:
            label = RE_FOOT.sub("", m.group("label").strip())
            nums = parse_values(m.group("values"), 7)
            if label in (previous, current):
                if region is None:
                    raise ValueError(f"Found {label} measurements before a region in {crop}")
                out.setdefault(norm(region), {})[label] = nums
            elif label:
                region = label
            continue
        # Headers on standalone lines, e.g. "World  3/", "Argentina".
        stripped = RE_FOOT.sub("", line.strip())
        if not stripped or not re.search(r"[A-Za-z]", stripped):
            continue
        if (stripped.startswith("WASDE -") or "(Contd.)" in stripped
            or "Supply and Use" in stripped or "Metric Tons" in stripped
            or "Stocks" in stripped or "Produc-" in stripped
            or "Domestic" in stripped or stripped.startswith("Region")
            or stripped in ("Supply", "Use", "Selected Other", "Major Exporters", "Major Importers")):
            # Major Importers/Exporters are actual region keys; keep them.
            if stripped in ("Major Exporters", "Major Importers"):
                region = stripped
            continue
        # Footnoted group headers (e.g. "Major Importers  6/").
        candidates = (WHEAT_REGIONS if crop == 'wheat' else CORN_REGIONS)
        candidates = set(norm(x) for x in candidates) | {
            'world','worldlesschina','unitedstates','totalforeign',
            'majorexporters','majorimporters','selectedother',
        }
        if norm(stripped) in candidates:
            region = stripped
    required = WHEAT_REGIONS if crop == "wheat" else CORN_REGIONS
    missing = [r for r in required if not all(m in out.get(norm(r), {}) for m in (previous, current))]
    if missing:
        raise ValueError(f"World {crop} missing regions or monthly pairs: {missing}")
    return out, season



# Extra tables follow the user's uploaded October 2026 files.
# The original worksheets are preserved. New sheets are built from a copy.
WORKBOOK_NAMES = {
    "wheat": "Wheat Wasde.xlsx",
    "corn": "Corn Wasde(1).xlsx",
    "soybean": "Soybean Wasde(1).xlsx",
    "soyoil": "Soyoil Wasde(1).xlsx",
    "sugar": "Sugar Wasde(1).xlsx",
}
FILES = {crop: HERE / "workbooks" / name for crop, name in WORKBOOK_NAMES.items()}


def configure(path: Path, require_workbooks: bool) -> None:
    """Resolve the workbook directory relative to the local config file."""
    if not path.is_file():
        if require_workbooks:
            raise FileNotFoundError(f"Create {path} from config.example.json")
        return
    config = json.loads(path.read_text(encoding="utf-8-sig"))
    directory = (path.resolve().parent / os.path.expanduser(os.path.expandvars(config["workbook_directory"]))).resolve()
    overrides = config.get("workbooks", {})
    if overrides.keys() - WORKBOOK_NAMES.keys():
        raise ValueError("Unknown crop in config.workbooks")
    names = WORKBOOK_NAMES | overrides
    for name in names.values():
        if Path(name).name != name or Path(name).suffix.lower() != ".xlsx":
            raise ValueError("Workbook values must be plain .xlsx filenames")
    if len({name.casefold() for name in names.values()}) != len(names):
        raise ValueError("Each crop must use a distinct workbook")
    FILES.update({crop: directory / name for crop, name in names.items()})


US_ROWS = {
    'wheat': WHEAT_US_ROWS,
    'corn': CORN_US_ROWS,
    'soybean': {5:'areaplanted', 6:'areaharvested', 8:'yieldperharvestedacre',
                10:'beginningstocks',11:'production',12:'imports',13:'supplytotal',
                14:'crushings',15:'exports',16:'seed',17:'residual',18:'usetotal',
                19:'endingstocks',20:'avgfarmprice'},
    'soyoil': {6:'beginningstocks',7:'production',8:'imports',9:'supplytotal',
               10:'domesticdisappearance',11:'biofuel',12:'industrial',
               13:'exports',14:'usetotal',15:'endingstocks',16:'avgprice'},
    'sugar': {6:'beginningstocks',7:'production',8:'beetsugar',9:'canesugar',
              10:'florida',11:'louisiana',12:'texas',13:'imports',14:'trq',
              15:'otherprogram',16:'nonprogram',17:'mexico',
              18:'hightiertariffother',19:'totalsupply',20:'exports',
              21:'deliveries',22:'food',23:'other',24:'miscellaneous',
              25:'totaluse',26:'endingstocks',27:'stockstouseratio'},
}
WORLD_REGIONS = {
    'wheat': WHEAT_REGIONS, 'corn': CORN_REGIONS,
    'soybean': ['World','United States','Argentina','Brazil','Paraguay',
                'China','European Union','Southeast Asia','Mexico'],
    # Uploaded Soyoil world sheets mix in unrelated grain rows (Russia,
    # South Africa, Egypt) and duplicate World/US. The new month uses
    # the actual World Soybean Oil supply/use entities in the USDA report.
    'soyoil': ['World','World Less China','United States','Total Foreign',
               'Major Exporters','Argentina','Brazil','European Union',
               'Major Importers','China','India','North Africa'],
}
WORLD_COLUMNS = {'wheat':(0,1,2,5,6),'corn':tuple(range(7)),
                 'soybean':tuple(range(7)), 'soyoil':tuple(range(6))}
MONTH_ABBR_TO_FULL = {v:MONTH_NAMES[k] for k,v in ABBR.items()}


def extract_us_rows(body: str, crop: str):
    """Read the US 4-column balance; do not reuse labels from another crop."""
    # Annual columns appear twice for the same projected season.
    yline=next((line for line in body.splitlines()
                if len(re.findall(r'20\d{2}/\d{2}',line))>=4),'')
    years=re.findall(r'20\d{2}/\d{2}', yline)
    if len(years)!=4:
        raise ValueError(f'US {crop}: expected four marketing-year headers, got {years}')
    result={}
    for line in body.splitlines():
        match=RE_US.match(line)
        if not match:
            continue
        key=norm(match.group('label'))
        # USDA oil label is line-wrapped: "Food, Feed & other" / "Industrial".
        if key=='industrial':
            result['industrial']=parse_values(match.group('values'),4)
        elif key.startswith('avgfarmprice'):
            result['avgfarmprice']=parse_values(match.group('values'),4)
        elif key.startswith('avgprice'):
            result['avgprice']=parse_values(match.group('values'),4)
        else:
            result[key]=parse_values(match.group('values'),4)
    missing=[f'{row}: {metric}' for row,metric in US_ROWS[crop].items()
             if metric not in result]
    if missing:
        raise ValueError(f'US {crop} missing required rows: {missing}')
    return result, years


def parse_extra_us(full: str, crop: str):
    if crop in ('soybean','soyoil'):
        text=section_between(full, 'U.S. Soybeans and Products Supply and Use',
                              'U.S. Sugar Supply and Use')
        markers={m.group(1):m for m in re.finditer(
            r'(?m)^\s*(SOYBEANS|SOYBEAN OIL|SOYBEAN MEAL)\s*$',text)}
        first=markers['SOYBEANS'].start()
        oil=markers['SOYBEAN OIL'].start()
        meal=markers['SOYBEAN MEAL'].start()
        headings=text[:first]
        body=headings+(text[first:oil] if crop=='soybean' else text[oil:meal])
        return extract_us_rows(body,crop)
    if crop=='sugar':
        return extract_us_rows(section_between(
            full, 'U.S. Sugar Supply and Use','Mexico Sugar Supply and Use'),crop)
    raise ValueError('Unexpected crop '+crop)


def parse_extra_world(full: str, crop: str, previous: str, current: str):
    stop = ('World Soybean Meal Supply and Use' if crop == 'soybean'
            else 'U.S. Quarterly Animal Product Production')
    start = ('World Soybean Supply and Use' if crop == 'soybean'
             else 'World Soybean Oil Supply and Use')
    body=section_between(full,start,stop)
    seasons=list(RE_SEASON.finditer(body))
    if not seasons:
        raise ValueError(f'World {crop}: projected season not present')
    latest=seasons[-1]
    season=latest.group(1)
    region=None
    data={}
    expected=set(norm(x) for x in WORLD_REGIONS[crop]) | {'totalforeign'}
    numfields=7 if crop=='soybean' else 6
    pattern=re.compile(rf'^\s*(?P<label>.*?)\s{{2,}}(?P<values>{NUMBER}(?:\s+{NUMBER}){{{numfields-1}}})\s*$')
    for line in body[latest.end():].splitlines():
        m=pattern.match(line)
        if m:
            label=RE_FOOT.sub('',m.group('label').strip())
            if label in (previous,current):
                if region is None:
                    raise ValueError(f'{crop}: period line without region')
                vals=parse_values(m.group('values'),numfields)
                data.setdefault(norm(region),{})[label]=vals
            # No country is expected on same line as its month in projected section.
            continue
        label=RE_FOOT.sub('',line.strip())
        if norm(label) in expected:
            region=label
    missing=[r for r in WORLD_REGIONS[crop]
             if not all(m in data.get(norm(r),{}) for m in (previous,current))]
    if missing:
        raise ValueError(f'World {crop}: missing monthly region pairs: {missing}')
    return data,season


def parse_mexico(full: str, previous: str, current: str):
    body=section_between(full,'Mexico Sugar Supply and Use','U.S. Cotton Supply and Use')
    # The Mexico table contains both an estimated and projected marketing year.
    headers=list(re.finditer(r'(?m)^\s*(20\d{2}/\d{2})\s+(Est\.|Proj\.)\s*$',body))
    if len(headers)<2:
        raise ValueError('Mexico sugar: expected two marketing years')
    result={}
    row_re=re.compile(rf'^\s*(?P<month>[A-Z][a-z]{{2}})\s+(?P<values>{NUMBER}(?:\s+{NUMBER}){{5}})\s*$')
    for i,h in enumerate(headers):
        season=h.group(1)
        segment=body[h.end():headers[i+1].start() if i+1<len(headers) else len(body)]
        rows={}
        for line in segment.splitlines():
            match=row_re.match(line)
            if match and match.group('month') in (previous,current):
                rows[match.group('month')]=parse_values(match.group('values'),6)
        if len(rows)==2:
            result[season]=rows
    if len(result)!=2:
        raise ValueError(f'Mexico sugar: missing {previous}/{current} pairs for two seasons; found {result.keys()}')
    return result


def parse_wasde(txt: str, year: int, month: int) -> dict:
    name=MONTH_NAMES[month]
    required=('U.S. Wheat Supply and Use','World Corn Supply and Use',
              'U.S. Soybeans and Products Supply and Use','World Soybean Oil Supply and Use',
              'U.S. Sugar Supply and Use','Mexico Sugar Supply and Use')
    if len(txt)<40_000 or 'WASDE -' not in txt or f'{name} {year}' not in txt or any(s not in txt for s in required):
        raise ValueError(f'Response is not the complete published {name} {year} USDA WASDE TXT')
    prev=ABBR[12 if month==1 else month-1]; curr=ABBR[month]
    data={'date':f'{year:04d}-{month:02d}','prev_month':prev,'current_month':curr}
    for crop in FILES:
        if crop in ('wheat','corn'):
            us, years=parse_us(txt,crop)
            world, season=parse_world(txt,crop,prev,curr)
        elif crop=='sugar':
            us,years=parse_extra_us(txt,crop)
            data[crop]={'us':us,'us_years':years,'mexico':parse_mexico(txt,prev,curr)}
            continue
        else:
            us,years=parse_extra_us(txt,crop)
            world,season=parse_extra_world(txt,crop,prev,curr)
        if season!=years[-1]:
            raise ValueError(f'{crop}: World {season} and US {years[-1]} marketing years disagree')
        if world['world'][curr][1]<=0 or world['world'][curr][-1]<0:
            raise ValueError(f'{crop}: invalid world production or ending stocks')
        data[crop]={'us':us,'us_years':years,'world':world,'season':season}
    validate_report_values(data)
    return data


class ArchiveLinks(HTMLParser):
    def __init__(self, filename):
        super().__init__()
        self.filename = filename
        self.urls = []

    def handle_starttag(self, tag, attrs):
        href = dict(attrs).get("href") if tag == "a" else None
        if not href:
            return
        url = urlsplit(urljoin(ARCHIVE, href))
        if (url.scheme == "https" and url.netloc == "esmis.nal.usda.gov"
                and url.path.startswith("/sites/default/release-files/")
                and url.path.rsplit("/", 1)[-1].casefold() == self.filename):
            self.urls.append(url.geturl())


def fetch(url, accept):
    req = Request(url, headers={"User-Agent": "Mozilla/5.0", "Accept": accept})
    with urlopen(req, timeout=TIMEOUT_SECONDS) as response:
        return response.read(2500000).decode("utf-8-sig", errors="replace"), response.headers.get("Content-Type", "")


def archive_report_urls(filename):
    page, _ = fetch(ARCHIVE, "text/html")
    links = ArchiveLinks(filename)
    links.feed(page)
    return list(dict.fromkeys(links.urls))


def get_report(year: int, month: int):
    filename = f"wasde{month:02d}{year%100:02d}.txt"
    urls = [f"https://www.usda.gov/oce/commodity/wasde/{filename}",
            f"https://usda.gov/oce/commodity/wasde/{filename}"]
    errors = []
    for candidates in (lambda: urls, lambda: archive_report_urls(filename)):
        try:
            for url in candidates():
                try:
                    txt, content_type = fetch(url, "text/plain, */*")
                    if ("html" in content_type.lower() or len(txt) < 40000 or "WASDE -" not in txt
                            or f"{MONTH_NAMES[month]} {year}" not in txt):
                        raise ValueError("not the complete requested WASDE TXT")
                    return txt, url
                except (ValueError, OSError) as exc:
                    errors.append(f"{url}: {exc}")
        except OSError as exc:
            errors.append(f"USDA archive: {exc}")
    raise RuntimeError("USDA report unavailable: " + " | ".join(errors))


# Version 3: monthly snapshots, immutable earlier months, clean World/Mexico layouts.
# The only sheets a user can deliberately rebuild are from the explicitly selected
# report month, and only after a full workbook backup.

US_LABELS = {
 'wheat': 'U.S. WHEAT', 'corn': 'U.S. CORN', 'soybean': 'U.S. SOYBEANS',
 'soyoil': 'U.S. SOYBEAN OIL', 'sugar': 'U.S. SUGAR'
}
WORLD_HEADINGS = {
 'wheat': ['Beginning Stocks', 'Production', 'Imports', 'Exports', 'Ending Stocks'],
 'corn': ['Beginning Stocks', 'Production', 'Imports', 'Domestic Feed',
          'Domestic Total', 'Exports', 'Ending Stocks'],
 'soybean': ['Beginning Stocks', 'Production', 'Imports', 'Domestic Crush',
             'Domestic Total', 'Exports', 'Ending Stocks'],
 'soyoil': ['Beginning Stocks', 'Production', 'Imports', 'Domestic Total',
            'Exports', 'Ending Stocks'],
}


def validate_report_values(report: dict) -> None:
    """Fail closed on common parsing mistakes; all values must come from target report."""
    curr = report['current_month']
    for crop in ('wheat', 'corn', 'soybean', 'soyoil'):
        regions=report[crop]['world']
        world=regions['world'][curr]
        us=regions.get('unitedstates', {}).get(curr)
        foreign=regions.get('totalforeign', {}).get(curr)
        if us is None or foreign is None:
            raise ValueError(f'{crop}: USDA World, United States and Total Foreign groups are all required')
        if len(world) != len(us) or len(world) != len(foreign):
            raise ValueError(f'{crop}: inconsistent World / US / Foreign columns')
        # Production is always field index 1; rounding is at most a few 0.01 MT.
        if abs(world[1] - (us[1] + foreign[1])) > Decimal('0.10'):
            raise ValueError(f'{crop}: World production does not equal US + Total Foreign: '
                             f'{world[1]} != {us[1]} + {foreign[1]}')
        if world[1] <= max(us[1], foreign[1]):
            raise ValueError(f'{crop}: World production incorrectly mapped to subregion')
        for region in WORLD_REGIONS[crop]:
            for period in (report['prev_month'],curr):
                vals=regions[norm(region)][period]
                if len(vals) not in (6,7):
                    raise ValueError(f'{crop}/{region}: invalid metric count')
                if any(not v.is_finite() for v in vals):
                    raise ValueError(f'{crop}/{region}: non-finite USDA metric')
    # US sugar stock identity; rounding to whole 1,000 short tons.
    sugar=report['sugar']['us']
    for j in (2,3):
        supply=sugar['totalsupply'][j]
        use=sugar['totaluse'][j]
        ending=sugar['endingstocks'][j]
        if abs(supply-use-ending)>2:
            raise ValueError('U.S. sugar balance fails: Total Supply - Total Use != Ending Stocks')
    for season,pair in report['sugar']['mexico'].items():
        for period,nums in pair.items():
            if abs(nums[0]+nums[1]+nums[2]-nums[3]-nums[4]-nums[5])>2:
                raise ValueError(f'Mexico sugar {season}/{period}: stocks do not reconcile')


# Excel COM may reject calls temporarily when busy; only retry transient errors.
COM_RETRYABLE_HRESULTS = {
 -2147418111,  # RPC_E_CALL_REJECTED
 -2147417846,  # RPC_E_SERVERCALL_RETRYLATER
 -2147417845,  # RPC_E_SERVERCALL_REJECTED
}

def com_retry(action, what: str, timeout: float = 45):
    deadline=time.monotonic()+timeout
    while True:
        try:
            return action()
        except Exception as exc:
            # Only retry the known, transient Windows COM errors. Do not
            # swallow AttributeError, merged-cell errors, etc.
            if sys.platform != 'win32' and 'pywintypes' not in sys.modules:
                raise
            import pywintypes, pythoncom
            if not isinstance(exc, pywintypes.com_error):
                raise
            code=getattr(exc,'hresult',exc.args[0] if exc.args else None)
            if code not in COM_RETRYABLE_HRESULTS or time.monotonic()>=deadline:
                raise
            try: pythoncom.PumpWaitingMessages()
            except Exception: pass
            logging.warning('Excel busy; retry %s: %s',code,what)
            time.sleep(0.50)


def put(ws,row:int,col:int,value,formula:bool=False):
    prop='Formula' if formula else 'Value2'
    com_retry(lambda:setattr(ws.Cells(row,col),prop,value),f'{prop} R{row}C{col}')


def read(ws,row:int,col:int):
    return com_retry(lambda:ws.Cells(row,col).Value2,f'get R{row}C{col}')


def col_letter(index:int)->str:
    out=''
    while index:
        index,r=divmod(index-1,26)
        out=chr(65+r)+out
    return out


def tab_names(year:int,month:int,crop:str):
    base=f"{MONTH_NAMES[month]}'{year%100:02d}"
    return base+' Balance Sheet',base+( ' Mexico Sugar' if crop=='sugar' else
        ' World '+{'wheat':'Wheat','corn':'Corn','soybean':'Soybean','soyoil':'SoyOil'}[crop])


def previous_month(year,month):
    return (year-1,12) if month==1 else (year,month-1)


def compatible_sheet_names(names,year,month,crop,kind):
    """Find historical user's quirky month names; no matching by month alone."""
    short=str(year%100).zfill(2)
    month_s=MONTH_NAMES[month].lower()
    if kind=='us': terms=('balance sheet',)
    elif crop=='sugar': terms=('mexico sugar',)
    else: terms=('world '+('soyoil' if crop=='soyoil' else crop),)
    matches=[]
    for name in names:
        s=name.strip().lower()
        # October'26, October 26', October26, November'26 etc.
        if not s.startswith(month_s): continue
        prefix=s[:len(month_s)+7]
        if not re.match(r"^"+re.escape(month_s)+r"\s*'?"+short+r"\s*'?",prefix): continue
        if all(term not in s for term in terms):continue
        matches.append(name)
    return matches


def all_sheet_names(wb):
    return com_retry(lambda:[ws.Name for ws in wb.Worksheets], 'list workbook sheets')


def find_previous_sheet(wb,year,month,crop,kind):
    names=all_sheet_names(wb)
    matches=compatible_sheet_names(names,year,month,crop,kind)
    desired=tab_names(year,month,crop)[0 if kind=='us' else 1]
    for name in matches:
        if name.strip().casefold()==desired.casefold():return name
    if len(matches)!=1:
        raise ValueError(f'{crop}: expected one {year:04d}-{month:02d} {kind} template; got {matches}')
    return matches[0]


def status_sheet(wb,year,month,url,digest,crop):
    names=all_sheet_names(wb)
    if 'WASDE_STATUS' in names:ws=wb.Worksheets('WASDE_STATUS')
    else:
        ws=wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
        ws.Name='WASDE_STATUS'
    fields=[('Status','UPDATED / VERIFIED'),('Report',f'{year:04d}-{month:02d}'),
            ('Product',crop.upper()),('Source',url),('SHA256',digest),
            ('Last refresh',datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
            ('Policy','Every release month = two distinct tabs; older months immutable'),
            ('Change','Current USDA month - previous USDA month, same USDA report'),
            ('Repair','Only explicitly selected report-month tabs may be recreated')]
    for i,(key,value) in enumerate(fields,2):
        put(ws,i,1,key);put(ws,i,2,value)
    ws.Range('A2:A10').Font.Bold=True
    ws.Columns('A:A').ColumnWidth=22
    ws.Columns('B:B').ColumnWidth=86


def check_near(actual, expected, label, tol=0.005):
    try:
        if actual is None or abs(float(actual)-float(expected))>tol:
            raise ValueError(f'Excel validation failed {label}: {actual!r}, expected {expected!r}')
    except (TypeError,ValueError,OverflowError) as exc:
        raise ValueError(f'Excel validation failed {label}: {actual!r}, expected {expected!r}') from exc


def fill_us(ws,data,crop,prev,curr):
    year_headers=data['us_years']
    header=2 if crop=='soybean' else 3
    put(ws,header,2,US_LABELS[crop])
    put(ws,header,3,year_headers[0])
    put(ws,header,4,year_headers[1]+' Est.')
    put(ws,header,5,f'{year_headers[-1]} Proj. {MONTH_NAMES[next(i for i,v in ABBR.items() if v==prev)]}')
    put(ws,header,6,f'{year_headers[-1]} Proj. {MONTH_NAMES[next(i for i,v in ABBR.items() if v==curr)]}')
    put(ws,header,7,'Change')
    for row,key in US_ROWS[crop].items():
        vals=data['us'][key]
        for k,v in enumerate(vals,3):put(ws,row,k,float(v))
        # Any inherited shared formula needs to be explicitly replaced with
        # a standalone formula. This fixes stale/wrong G rows in Sugar.
        com_retry(lambda:ws.Cells(row,7).ClearContents(),f'clear change {crop} {row}')
        put(ws,row,7,f'=F{row}-E{row}',formula=True)
    com_retry(lambda:ws.Calculate(),f'calculate US {crop}')
    for row,key in US_ROWS[crop].items():
        values=data['us'][key]
        for col,v in enumerate(values,3):
            check_near(read(ws,row,col),v,f'{crop} US {key} {col}')
        expected=values[3]-values[2]
        check_near(read(ws,row,7),expected,f'{crop} US {key} Change')
        formula=com_retry(lambda:ws.Cells(row,7).Formula,f'read formula {crop} {row}')
        if str(formula).upper().replace('$','')!=f'=F{row}-E{row}':
            raise ValueError(f'{crop} US {key}: wrong Change formula {formula!r}')


def create_world_sheet(wb,name,data,crop,prev,curr):
    ws=wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name=name
    regions=WORLD_REGIONS[crop]
    fields=WORLD_HEADINGS[crop]
    indices=WORLD_COLUMNS[crop]
    width=3+len(fields)
    last=col_letter(width)
    # Clean NEW worksheet. No inherited merged cells, stale labels or formulas.
    ws.Range(f'B2:{last}2').Interior.Color=0x402814
    ws.Range(f'B2:{last}2').Font.Color=0xFFFFFF
    ws.Range(f'B2:{last}2').Font.Bold=True
    ws.Range(f'B2:{last}2').RowHeight=37
    ws.Range(f'D2:{last}2').WrapText=True
    ws.Columns('B:B').ColumnWidth=23
    ws.Columns('C:C').ColumnWidth=12
    ws.Range(f'D:{last}').ColumnWidth=15
    put(ws,2,2,f'{crop.upper()} | {data["season"]} Proj.')
    put(ws,2,3,'Release')
    for j,label in enumerate(fields,4):put(ws,2,j,label)
    for i,region in enumerate(regions):
        row=3+3*i
        ws.Range(f'B{row}:B{row+2}').Merge()
        put(ws,row,2,region)
        put(ws,row,3,prev);put(ws,row+1,3,curr);put(ws,row+2,3,'Change')
        pair=data['world'][norm(region)]
        for j,idx in enumerate(indices,4):
            put(ws,row,j,float(pair[prev][idx]))
            put(ws,row+1,j,float(pair[curr][idx]))
            put(ws,row+2,j,f'={col_letter(j)}{row+1}-{col_letter(j)}{row}',formula=True)
        ws.Range(f'C{row+2}:{last}{row+2}').Interior.Color=0xF1F5F9
        if i%2==1:
            ws.Range(f'B{row}:{last}{row+1}').Interior.Color=0xF8FAFC
        ws.Range(f'B{row}:{last}{row+2}').Borders(9).LineStyle=1  # xlEdgeBottom
    ws.Range(f'D3:{last}{3+3*len(regions)}').NumberFormat='0.00;[Red]-0.00;0.00'
    ws.Range(f'D3:{last}{3+3*len(regions)}').HorizontalAlignment=-4108
    ws.Range('B2:C2').HorizontalAlignment=-4131
    com_retry(lambda:ws.Calculate(),f'calc World {crop}')
    for i,region in enumerate(regions):
        row=3+3*i;pair=data['world'][norm(region)]
        for j,idx in enumerate(indices,4):
            check_near(read(ws,row,j),pair[prev][idx],f'{crop}/{region}/{prev}/{j}')
            check_near(read(ws,row+1,j),pair[curr][idx],f'{crop}/{region}/{curr}/{j}')
            check_near(read(ws,row+2,j),pair[curr][idx]-pair[prev][idx],f'{crop}/{region}/Change/{j}')
    return ws


def create_mexico_sheet(wb,name,data,prev,curr):
    ws=wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
    ws.Name=name
    ws.Columns('B:B').ColumnWidth=19
    ws.Columns('C:C').ColumnWidth=14
    ws.Range('D:I').ColumnWidth=17
    ws.Range('B1:I1').Merge()
    ws.Range('B1:I1').Interior.Color=0x402814
    ws.Range('B1:I1').Font.Color=0xFFFFFF
    ws.Range('B1:I1').Font.Bold=True
    ws.Range('B1:I1').RowHeight=28
    put(ws,1,2,'MEXICO SUGAR | USDA WASDE SUPPLY & USE')
    ws.Range('B3:C3').Merge()
    put(ws,3,2,'MEXICO')
    for col,label in enumerate(('Beginning Stocks','Production','Imports','Domestic Total',
                                'Exports','Ending Stocks'),4):put(ws,3,col,label)
    ws.Range('D3:I3').Font.Bold=True
    ws.Range('D3:I3').WrapText=True
    ws.Range('D3:I3').RowHeight=35
    ws.Range('D10:I10').Merge()
    put(ws,10,4,'1,000 Metric Tons, Actual Weight')
    for i,(season,pair) in enumerate(sorted(data['mexico'].items())):
        row=11+3*i
        ws.Range(f'B{row}:B{row+2}').Merge()
        put(ws,row,2,f'{season} '+('Est.' if i==0 else 'Proj.'))
        put(ws,row,3,prev);put(ws,row+1,3,curr);put(ws,row+2,3,'Change')
        for j,(p,c) in enumerate(zip(pair[prev],pair[curr]),4):
            put(ws,row,j,float(p));put(ws,row+1,j,float(c))
            put(ws,row+2,j,f'={col_letter(j)}{row+1}-{col_letter(j)}{row}',formula=True)
        ws.Range(f'C{row+2}:I{row+2}').Interior.Color=0xF1F5F9
        ws.Range(f'B{row}:I{row+2}').Borders(9).LineStyle=1
    ws.Range('D11:I16').NumberFormat='0;[Red]-0;0'
    com_retry(lambda:ws.Calculate(),'calculate Mexico')
    for i,(season,pair) in enumerate(sorted(data['mexico'].items())):
        row=11+3*i
        for col,(p,c) in enumerate(zip(pair[prev],pair[curr]),4):
            check_near(read(ws,row,col),p,f'Mexico {season} {prev} {col}')
            check_near(read(ws,row+1,col),c,f'Mexico {season} {curr} {col}')
            check_near(read(ws,row+2,col),c-p,f'Mexico {season} delta {col}')
    return ws


def update_one_excel(excel,filepath:Path,crop:str,data:dict,year:int,month:int,
                     url:str,digest:str,prev:str,curr:str,repair:bool):
    wb=com_retry(lambda:excel.Workbooks.Open(str(filepath.resolve()),UpdateLinks=0,ReadOnly=False),
                 f'open {crop}',timeout=60)
    try:
        if wb.ReadOnly:raise PermissionError(f'{crop}: close Excel; file is read-only')
        names=all_sheet_names(wb)
        usname,worldname=tab_names(year,month,crop)
        existing_us=compatible_sheet_names(names,year,month,crop,'us')
        existing_world=compatible_sheet_names(names,year,month,crop,'world')
        if (existing_us or existing_world) and not repair:
            raise FileExistsError(f'{crop}: report-month tabs already exist: '
                f'{existing_us+existing_world}. Use --repair-current to rebuild that month ONLY.')
        if repair:
            # Backups were created by apply_excel(). Delete ONLY target-month pages.
            for sheet_name in existing_us+existing_world:
                wb.Worksheets(sheet_name).Delete()
            logging.info('Repair %s: replaced current-month sheets %s',crop,
                         existing_us+existing_world)
        py,pm=previous_month(year,month)
        source=find_previous_sheet(wb,py,pm,crop,'us')
        copied=wb.Worksheets(source).Copy(After=wb.Worksheets(wb.Worksheets.Count))
        us_ws=copied if sys.platform == 'darwin' else wb.ActiveSheet
        us_ws.Name=usname
        fill_us(us_ws,data,crop,prev,curr)
        if crop=='sugar':create_mexico_sheet(wb,worldname,data,prev,curr)
        else:create_world_sheet(wb,worldname,data,crop,prev,curr)
        status_sheet(wb,year,month,url,digest,crop)
        com_retry(lambda:wb.Save(),f'save {crop}',timeout=90)
        final_names=all_sheet_names(wb)
        if final_names.count(usname)!=1 or final_names.count(worldname)!=1:
            raise ValueError(f'{crop}: monthly tabs not created exactly once')
        logging.info('VERIFIED %s %s / %s',crop,usname,worldname)
    finally:
        com_retry(lambda:wb.Close(SaveChanges=False),f'close {crop}',timeout=60)


def verify_saved_excel(excel, filepath: Path, crop: str, data: dict, year: int, month: int, prev: str, curr: str):
    """Reopen saved staging workbook and verify it actually persisted correct month/deltas."""
    wb = com_retry(lambda: excel.Workbooks.Open(str(filepath.resolve()), UpdateLinks=0, ReadOnly=True),
                   f'verify-open {crop}', timeout=60)
    try:
        us_name, other_name = tab_names(year, month, crop)
        names = all_sheet_names(wb)
        if names.count(us_name)!=1 or names.count(other_name)!=1:
            raise ValueError(f'{crop}: saved workbook missing unique monthly tab pair')
        ws=wb.Worksheets(us_name)
        hdr=2 if crop=='soybean' else 3
        prev_label=f'{data["us_years"][-1]} Proj. {MONTH_ABBR_TO_FULL[prev]}'
        curr_label=f'{data["us_years"][-1]} Proj. {MONTH_ABBR_TO_FULL[curr]}'
        if read(ws,hdr,5)!=prev_label or read(ws,hdr,6)!=curr_label:
            raise ValueError(f'{crop}: wrong saved report month labels')
        ws.Calculate()
        # All US rows, including the previously broken Sugar last six rows.
        for row,key in US_ROWS[crop].items():
            values=data['us'][key]
            check_near(read(ws,row,5),values[2],f'{crop} {key} prev saved')
            check_near(read(ws,row,6),values[3],f'{crop} {key} curr saved')
            check_near(read(ws,row,7),values[3]-values[2],f'{crop} {key} change saved')
        ws2=wb.Worksheets(other_name)
        ws2.Calculate()
        if crop=='sugar':
            for i,(season,pair) in enumerate(sorted(data['mexico'].items())):
                row=11+3*i
                if read(ws2,row,3)!=prev or read(ws2,row+1,3)!=curr:
                    raise ValueError(f'{crop} Mexico: month labels wrong in saved file')
                for col,(p,c) in enumerate(zip(pair[prev],pair[curr]),4):
                    check_near(read(ws2,row+2,col),c-p,f'Mexico {season} {col} saved')
        else:
            regions=WORLD_REGIONS[crop]
            indices=WORLD_COLUMNS[crop]
            # Verify the aggregate World and final region, plus a representative country.
            for region in {regions[0],regions[-1],regions[len(regions)//2]}:
                row=3+3*regions.index(region)
                if read(ws2,row,3)!=prev or read(ws2,row+1,3)!=curr:
                    raise ValueError(f'{crop}/{region}: month labels wrong in saved file')
                for col,idx in enumerate(indices,4):
                    pair=data['world'][norm(region)]
                    check_near(read(ws2,row+2,col),pair[curr][idx]-pair[prev][idx],
                               f'{crop}/{region} change col {col} saved')
        logging.info('POST-SAVE VERIFIED %s %04d-%02d',crop,year,month)
    finally:
        com_retry(lambda: wb.Close(SaveChanges=False),f'verify-close {crop}',timeout=60)


def read_state():
    try:return json.loads(STATE_FILE.read_text(encoding='utf-8'))
    except FileNotFoundError:return {}


def save_state(state):
    temp=STATE_FILE.with_suffix('.json.tmp')
    temp.write_text(json.dumps(state,indent=2,ensure_ascii=False)+'\n',encoding='utf-8')
    os.replace(temp,STATE_FILE)


@contextmanager
def windows_excel_session():
    try:
        import win32com.client
    except ImportError as exc:
        raise RuntimeError('Run install.ps1 to install pywin32 in the project environment') from exc
    excel=win32com.client.DispatchEx('Excel.Application')
    calc_mode=None
    control_book=None
    try:
        excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        excel.Visible=False
        excel.DisplayAlerts=False
        excel.AskToUpdateLinks=False
        excel.EnableEvents=False
        excel.ScreenUpdating=False
        # No recalculation of historical sheets by Application.Calculate.
        control_book=excel.Workbooks.Add()
        calc_mode=excel.Calculation
        excel.Calculation=-4135   # xlCalculationManual
        yield excel
    finally:
        if calc_mode is not None:
            try:excel.Calculation=calc_mode
            except Exception:logging.warning('Could not reset Excel Calculation')
        try:
            if control_book is not None:
                control_book.Close(SaveChanges=False)
        finally:
            com_retry(lambda:excel.Quit(),'quit Excel',timeout=60)


def apply_excel(report,year,month,source,digest,repair=False):
    for crop,path in FILES.items():
        if not path.is_file():raise FileNotFoundError(f'Missing {crop}: {path.name}')
    stamp=datetime.now().strftime('%Y%m%d_%H%M%S_%f')
    backups=STATE_FILE.parent/'backups'/f'{year:04d}-{month:02d}_{stamp}'
    staging=next(iter(FILES.values())).parent/'_staging'/f'{year:04d}-{month:02d}_{stamp}'
    backups.mkdir(parents=True,exist_ok=False)
    if STATE_FILE.is_file():
        shutil.copy2(STATE_FILE, backups / STATE_FILE.name)
    staging.mkdir(parents=True,exist_ok=False)
    swapped=[]
    try:
        for original in FILES.values():
            shutil.copy2(original,backups/original.name)
            shutil.copy2(original,staging/original.name)
        if sys.platform == 'darwin':
            from excel_mac import session
        else:
            session = windows_excel_session
        with session() as excel:
            for crop,path in FILES.items():
                logging.info('Building fresh %04d-%02d: %s',year,month,crop)
                update_one_excel(excel,staging/path.name,crop,report[crop],year,month,source,digest,
                                 report['prev_month'],report['current_month'],repair)
                verify_saved_excel(excel,staging/path.name,crop,report[crop],year,month,
                                   report['prev_month'],report['current_month'])
        try:
            for original in FILES.values():
                os.replace(staging/original.name,original)
                swapped.append(original)
        except Exception:
            for original in swapped:
                shutil.copy2(backups/original.name,original)
            raise
    finally:
        shutil.rmtree(staging,ignore_errors=True)
    return backups


def main(argv=None):
    cli=argparse.ArgumentParser(description='USDA WASDE month-by-month five-crop Excel snapshotter')
    cli.add_argument('--report',help='YYYY-MM month to download, e.g. 2026-11')
    cli.add_argument('--dry-run',action='store_true',help='Validate source without changing Excel')
    cli.add_argument('--now',action='store_true',help='Check latest published report immediately')
    cli.add_argument('--repair-current',action='store_true',help='Backup and rebuild ONLY selected report-month tabs')
    cli.add_argument('--force',action='store_true',help='Deprecated unsafe option: refused')
    cli.add_argument('--file',type=Path,help='Offline USDA txt file for validation / testing')
    cli.add_argument('--config', type=Path, default=HERE / 'config.json')
    cli.add_argument('--scheduled', action='store_true', help='Run on configured release days only')
    cli.add_argument('--check-install', action='store_true', help='Check workbook paths and Excel COM')
    args=cli.parse_args(argv)
    LOG_FILE.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(filename=str(LOG_FILE),level=logging.INFO,
                        format='%(asctime)s %(levelname)s %(message)s',encoding='utf-8')
    if args.force:
        cli.error('--force disabled: use --repair-current to rebuild only the selected month')
    try:
        configure(args.config, require_workbooks=not args.dry_run)
        if args.check_install:
            for crop, path in FILES.items():
                if not path.is_file():
                    raise FileNotFoundError(f"{crop}: {path}")
            if sys.platform == 'darwin':
                from excel_mac import check_install
                check_install()
                return 0
            import win32com.client
            excel = win32com.client.DispatchEx("Excel.Application")
            try:
                print(f"PASS: Excel COM {excel.Version}; five workbook paths found")
            finally:
                excel.Quit()
            return 0
    except Exception:
        logging.exception('Installation/configuration check failed')
        print(f'Configuration or installation error; see {LOG_FILE}', file=sys.stderr)
        return 1
    now = datetime.now(ZoneInfo("Europe/Istanbul"))
    today = now.date()
    release_dates = json.loads(RELEASE_DATES.read_text(encoding='utf-8'))
    release_day = release_dates.get(f'{today:%Y-%m}')
    if args.scheduled:
        if today.isoformat() != release_day:
            return 0
        if not ((18, 55) <= (now.hour, now.minute) <= (21, 55)):
            return 0
    if args.report:
        m=re.fullmatch(r'(20\d{2})-(0[1-9]|1[0-2])',args.report)
        if not m:cli.error('--report requires YYYY-MM')
        year,month=int(m.group(1)),int(m.group(2))
        if (year,month)>(today.year,today.month):
            cli.error('future report cannot be processed')
    else:
        year,month=today.year,today.month
        if release_day is None:
            print(f'No release calendar for {year}; use --report YYYY-MM.',file=sys.stderr)
            return 2
        if today < date.fromisoformat(release_day):
            if not args.now:
                return 0
            year,month=previous_month(year,month)
    tag=f'{year:04d}-{month:02d}'
    try:
        state=read_state()
        if not (args.now or args.file or args.report or args.repair_current or args.dry_run) and state.get(tag):
            if (today-date.fromisoformat(release_day)).days>2:return 0
            if time.time()-state.get(tag+':last_checked',0)<1800:return 0
        if args.file:
            txt=args.file.read_text(encoding='utf-8-sig')
            source="Local USDA report (offline input)"
        else:
            txt,source=get_report(year,month)
        digest=hashlib.sha256(txt.encode('utf-8')).hexdigest()
        if not (args.dry_run or args.repair_current) and state.get(tag)==digest:
            state[tag+':last_checked']=time.time()
            save_state(state)
            print(f'NO CHANGE: USDA {tag} already recorded. Use --repair-current if month tabs are wrong.')
            return 0
        report=parse_wasde(txt,year,month)
        print(f'VALIDATED USDA WASDE {tag}, comparing {report["prev_month"]} -> {report["current_month"]}')
        for crop in FILES:
            if crop=='sugar':
                est=report['sugar']['us']['endingstocks']
                print(f'  Sugar US ending stocks: {est[2]} -> {est[3]} (change {est[3]-est[2]})')
            else:
                val=report[crop]['world']['world'][report['current_month']][1]
                print(f'  {crop} world production: {val} MMT')
        if args.dry_run:return 0
        backups=apply_excel(report,year,month,source,digest,repair=args.repair_current)
        state[tag]=digest
        state[tag+':last_checked']=time.time()
        save_state(state)
        logging.info('SUCCESS %s: 10 monthly sheets verified; backups %s',tag,backups)
        print(f'UPDATED five Excel workbooks; TWO verified monthly tabs per crop. BACKUPS: {backups}')
        return 0
    except Exception:
        logging.exception('WASDE update failed')
        print(f'Update aborted; see {LOG_FILE}',file=sys.stderr)
        return 1

if __name__=='__main__':
    if os.name in ('nt', 'posix'):

        STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
        with (STATE_FILE.parent / 'wasde.lock').open('a+b') as lock:
            lock.seek(0)
            try:
                if os.name == 'nt':
                    import msvcrt
                    msvcrt.locking(lock.fileno(), msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(lock.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            except OSError:
                sys.exit('Another WASDE run is in progress in this installation.')
            sys.exit(main())
    sys.exit(main())
