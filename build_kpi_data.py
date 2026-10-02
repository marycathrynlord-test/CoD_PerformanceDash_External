#!/usr/bin/env python3
"""
Builds kpi-data.json (read by index.html) from KPI_Tracking.xlsx.

Run from the repository folder:   python3 build_kpi_data.py

What it does
  1. Reads the "KPI Catalog" and "KPI Data" sheets of the workbook.
  2. Checks them (unknown IDs, text where a number should be, duplicates...).
  3. If anything is wrong, lists every problem and stops WITHOUT touching
     kpi-data.json, so the live dashboard keeps its last good numbers.
  4. Otherwise writes kpi-data.json.

It only uses Python's standard library, so there is nothing to install.
The numbers themselves never live in this file or in index.html.
"""

import json
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

# The script, the workbook, data.json and index.html normally all sit together
# in the top level of the repository. (A scripts/ + data-source/ folder layout
# also works.)
HERE = Path(__file__).resolve().parent
ROOT = HERE if (HERE / "data.json").exists() else HERE.parent
WORKBOOK = next(
    (p for p in (ROOT / "KPI_Tracking.xlsx", ROOT / "data-source" / "KPI_Tracking.xlsx") if p.exists()),
    ROOT / "KPI_Tracking.xlsx",
)
STRUCTURE = ROOT / "data.json"
OUTPUT = ROOT / "kpi-data.json"

CATALOG_SHEET = "KPI Catalog"
DATA_SHEET = "KPI Data"

# Status rule, matching the dashboard's Progress Key. A KPI can override
# these two numbers in its own catalog row.
DEFAULT_ON_TRACK_WITHIN_PCT = 5.0    # at/above target, or short by up to this % -> On Track
DEFAULT_OFF_TRACK_BEYOND_PCT = 15.0  # short by more than this %                 -> Off Track
                                     # anything in between                       -> At Risk

BLANK_WORDS = {"", "n/a", "na", "-", "--", "tbd"}

# The three levels of the dashboard a KPI can belong to, as typed in the
# "Level" column, and the list in data.json each one refers to.
LEVELS = {"work system": "workSystems", "key service": "keyServices", "program": "programs"}
LEVEL_LABELS = {"work system": "Work System", "key service": "Key Service", "program": "Program"}


# --------------------------------------------------------------------------
# Minimal .xlsx reader (an .xlsx file is a zip of XML files)
# --------------------------------------------------------------------------
NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "p": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def _col_index(cell_ref):
    letters = re.match(r"[A-Z]+", cell_ref).group(0)
    n = 0
    for ch in letters:
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _si_text(si):
    # Plain text sits in <t>; formatted text is split across <r><t> runs.
    # Phonetic guides (<rPh>) are deliberately not read.
    parts = [t.text or "" for t in si.findall("m:t", NS)]
    parts += [t.text or "" for t in si.findall("m:r/m:t", NS)]
    return "".join(parts)


def read_workbook(path):
    """Returns {sheet name: [(row number, [cell values]), ...]}."""
    sheets = {}
    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        shared = []
        if "xl/sharedStrings.xml" in names:
            root = ET.fromstring(z.read("xl/sharedStrings.xml"))
            shared = [_si_text(si) for si in root.findall("m:si", NS)]

        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        targets = {rel.get("Id"): rel.get("Target") for rel in rels.findall("p:Relationship", NS)}

        wb = ET.fromstring(z.read("xl/workbook.xml"))
        for sheet in wb.find("m:sheets", NS).findall("m:sheet", NS):
            target = targets[sheet.get(f"{{{NS['r']}}}id")]
            target = target.lstrip("/") if target.startswith("/") else "xl/" + target
            root = ET.fromstring(z.read(target))
            rows = []
            for row in root.iter(f"{{{NS['m']}}}row"):
                values = []
                for c in row.findall("m:c", NS):
                    idx = _col_index(c.get("r"))
                    ctype = c.get("t", "n")
                    v = c.find("m:v", NS)
                    value = None
                    if ctype == "s" and v is not None:
                        value = shared[int(v.text)]
                    elif ctype == "inlineStr":
                        is_node = c.find("m:is", NS)
                        value = _si_text(is_node) if is_node is not None else None
                    elif ctype in ("str", "e"):
                        value = v.text if v is not None else None
                    elif ctype == "b":
                        value = bool(int(v.text)) if v is not None else None
                    elif v is not None and v.text not in (None, ""):
                        value = float(v.text)
                    while len(values) <= idx:
                        values.append(None)
                    values[idx] = value
                rows.append((int(row.get("r", len(rows) + 1)), values))
            sheets[sheet.get("name")] = rows
    return sheets


def table(rows, sheet_name, problems):
    """Turns a sheet (header row first) into [(row number, {lower-case header: value}), ...]."""
    rows = [(n, r) for n, r in rows if any(c not in (None, "") for c in r)]
    if not rows:
        problems.append(f'"{sheet_name}" sheet is empty.')
        return []
    headers = [clean(h).lower() for h in rows[0][1]]
    out = []
    for n, r in rows[1:]:
        r = list(r) + [None] * (len(headers) - len(r))
        out.append((n, {h: r[i] for i, h in enumerate(headers) if h}))
    return out


# --------------------------------------------------------------------------
# Small helpers
# --------------------------------------------------------------------------
def clean(v):
    """Cell value -> trimmed text. Whole numbers lose their '.0' (2026.0 -> '2026')."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return "Yes" if v else "No"
    if isinstance(v, float):
        return str(int(v)) if v == int(v) else repr(v)
    return str(v).strip()


def is_blank(v):
    return clean(v).lower() in BLANK_WORDS


def to_number(v):
    """Cell value -> float, or None if blank / N/A. Raises ValueError for other text."""
    if is_blank(v):
        return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    text = clean(v).replace(",", "").replace("$", "").replace("%", "").strip()
    return float(text)  # ValueError if it isn't a number


def yes_no(v, default):
    text = clean(v).lower()
    if text == "":
        return default
    if text in ("yes", "y", "true", "1"):
        return True
    if text in ("no", "n", "false", "0"):
        return False
    raise ValueError(text)


def tidy(n):
    """Rounds away floating point noise (0.793 * 100 = 79.30000000000001)."""
    n = round(n, 6)
    return int(n) if n == int(n) else n


def status_for(value, target, higher_is_better, on_within, off_beyond):
    if value is None or target is None or target == 0:
        return None
    shortfall = (target - value) if higher_is_better else (value - target)
    shortfall_pct = shortfall / abs(target) * 100
    if shortfall_pct <= on_within:
        return "on-track"
    if shortfall_pct <= off_beyond:
        return "at-risk"
    return "off-track"


# --------------------------------------------------------------------------
# Build
# --------------------------------------------------------------------------
def build():
    problems, warnings, notes = [], [], []

    if not WORKBOOK.exists():
        sys.exit(f"Can't find the workbook at {WORKBOOK}")
    try:
        sheets = read_workbook(WORKBOOK)
    except Exception as err:  # corrupt or non-xlsx file
        sys.exit(f"Couldn't read {WORKBOOK.name}: {err}")

    for name in (CATALOG_SHEET, DATA_SHEET):
        if name not in sheets:
            problems.append(f'The workbook has no sheet named "{name}".')
    if problems:
        fail(problems)

    structure = json.loads(STRUCTURE.read_text(encoding="utf-8"))
    item_names, item_levels = {}, {}
    for level, list_name in LEVELS.items():
        for item in structure.get(list_name, []):
            item_names[item["id"]] = item["name"]
            item_levels[item["id"]] = level
    objective_names = {o["id"]: o["name"] for o in structure.get("objectives", [])}

    # ---- Catalog ----------------------------------------------------------
    catalog = {}
    for n, row in table(sheets[CATALOG_SHEET], CATALOG_SHEET, problems):
        where = f"{CATALOG_SHEET}, row {n}"
        kpi_id = clean(row.get("kpi id"))
        if not kpi_id:
            continue
        if kpi_id in catalog:
            problems.append(f'{where}: KPI ID "{kpi_id}" is used more than once.')
            continue
        name = clean(row.get("kpi name"))
        attach = clean(row.get("dashboard item id"))
        if not name:
            problems.append(f"{where}: KPI Name is blank.")
        level = " ".join(clean(row.get("level")).lower().split())
        if level and level not in LEVELS:
            problems.append(f'{where}: Level "{clean(row.get("level"))}" must be Work System, Key Service or Program.')
            level = ""
        if attach not in item_names:
            problems.append(
                f'{where}: Dashboard Item ID "{attach}" is not a Work System, Key Service or Program ID in data.json.'
            )
        elif level and item_levels[attach] != level:
            problems.append(
                f'{where}: Level says {LEVEL_LABELS[level]}, but {attach} is the '
                f'{LEVEL_LABELS[item_levels[attach]]} "{item_names[attach]}". Change one so they agree.'
            )
        level = item_levels.get(attach, level)

        # Optional: Strategic Plan Objectives this KPI is ALSO shown under
        # (one or more Objective IDs, separated by commas).
        objectives = []
        for ob in re.split(r"[,;]", clean(row.get("also show under objective"))):
            ob = ob.strip()
            if not ob:
                continue
            if ob not in objective_names:
                problems.append(f'{where}: "Also Show Under Objective" has "{ob}", which is not an Objective ID in data.json.')
            elif ob not in objectives:
                objectives.append(ob)

        unit = clean(row.get("unit")) or "number"
        if unit.lower() in ("percent", "number", "currency"):
            unit = unit.lower()

        entry = {
            "id": kpi_id,
            "name": name,
            "level": LEVEL_LABELS.get(level, ""),
            "attachTo": attach,
            "objectives": objectives,
            "unit": unit,
            "decimals": 1 if unit == "percent" else 0,
            "higherIsBetter": True,
            "compareTo": "target",
            "drivesStatus": False,
            "frequency": clean(row.get("update frequency")),
            "owner": clean(row.get("owner")),
            "source": clean(row.get("source")),
            "description": clean(row.get("description")),
            "footnote": clean(row.get("footnote")),
        }
        # Which number decides the status: the Target (default) or the Benchmark.
        # ("Compare Against" is the older name for this column.)
        basis = clean(row.get("status based on") or row.get("compare against")).lower()
        if basis in ("", "target", "benchmark"):
            entry["compareTo"] = basis or "target"
        else:
            problems.append(f'{where}: "Status Based On" must be Target or Benchmark.')
        on_within, off_beyond = DEFAULT_ON_TRACK_WITHIN_PCT, DEFAULT_OFF_TRACK_BEYOND_PCT
        try:
            if not is_blank(row.get("decimals")):
                entry["decimals"] = int(to_number(row.get("decimals")))
        except ValueError:
            problems.append(f"{where}: Decimals must be a whole number.")
        for key, label in (("higherIsBetter", "higher is better"), ("drivesStatus", "drives status")):
            try:
                entry[key] = yes_no(row.get(label), entry[key])
            except ValueError:
                problems.append(f'{where}: "{label.title()}" must be Yes or No.')
        try:
            if not is_blank(row.get("on track within %")):
                on_within = to_number(row.get("on track within %"))
            if not is_blank(row.get("off track beyond %")):
                off_beyond = to_number(row.get("off track beyond %"))
        except ValueError:
            problems.append(f"{where}: the two status threshold columns must be numbers.")
        if entry["drivesStatus"] and level == "work system":
            warnings.append(
                f"{where}: Drives Status is ignored for Work System KPIs. A Work System's color "
                "always comes from its Key Services."
            )
            entry["drivesStatus"] = False
        if on_within > off_beyond:
            problems.append(f'{where}: "On Track Within %" is larger than "Off Track Beyond %".')
        entry["_thresholds"] = (on_within, off_beyond)
        entry["_rows"] = []
        catalog[kpi_id] = entry

    # ---- Data -------------------------------------------------------------
    seen = set()
    for n, row in table(sheets[DATA_SHEET], DATA_SHEET, problems):
        where = f"{DATA_SHEET}, row {n}"
        kpi_id = clean(row.get("kpi id"))
        if not kpi_id:
            continue
        if kpi_id not in catalog:
            problems.append(f'{where}: KPI ID "{kpi_id}" is not listed in the {CATALOG_SHEET} sheet.')
            continue
        period = clean(row.get("period"))
        segment = clean(row.get("segment"))
        if not period:
            problems.append(f"{where}: Period is blank.")
            continue
        if re.fullmatch(r"[3-6]\d{4}", period):
            warnings.append(
                f'{where}: Period "{period}" looks like a date Excel converted to a number. '
                "Format the cell as Text and retype it."
            )
        key = (kpi_id, period, segment.lower())
        if key in seen:
            problems.append(f'{where}: duplicate row for {kpi_id} / {period} / {segment or "(headline)"}.')
            continue
        seen.add(key)

        try:
            value = to_number(row.get("value"))
        except ValueError:
            problems.append(f'{where}: Value "{clean(row.get("value"))}" is not a number.')
            continue
        comparisons = {}
        bad = False
        for column in ("target", "benchmark"):
            try:
                comparisons[column] = to_number(row.get(column))
            except ValueError:
                problems.append(f'{where}: {column.title()} "{clean(row.get(column))}" is not a number.')
                bad = True
        # Older workbooks had one combined "Target or Benchmark" column; it
        # fills whichever of the two this KPI's status is based on.
        if "target or benchmark" in row and not bad:
            try:
                legacy = to_number(row.get("target or benchmark"))
                basis = catalog[kpi_id]["compareTo"]
                if legacy is not None and comparisons[basis] is None:
                    comparisons[basis] = legacy
            except ValueError:
                problems.append(f'{where}: Target or Benchmark "{clean(row.get("target or benchmark"))}" is not a number.')
                bad = True
        if bad:
            continue
        if value is None:
            continue  # blank or N/A: nothing to publish for this row
        if catalog[kpi_id]["unit"] == "percent" and 0 < abs(value) < 1:
            warnings.append(
                f"{where}: {value} is being read as {value}%, not {value * 100:g}%. "
                "Percent KPIs are typed as 79.3, not 0.793."
            )
        catalog[kpi_id]["_rows"].append(
            {"period": period, "segment": segment, "value": value,
             "target": comparisons["target"], "benchmark": comparisons["benchmark"],
             "note": clean(row.get("note"))}
        )

    if problems:
        fail(problems)

    # ---- Assemble output --------------------------------------------------
    kpis = []
    for entry in catalog.values():
        rows = entry.pop("_rows")
        on_within, off_beyond = entry.pop("_thresholds")
        headline = [r for r in rows if not r["segment"]]
        if not headline:
            warnings.append(f'{entry["id"]}: no headline rows (blank Segment) with a value yet, so it is left off the dashboard.')
            continue
        if all(re.fullmatch(r"\d{4}", r["period"]) for r in headline):
            headline.sort(key=lambda r: int(r["period"]))  # plain years: sort them

        series = []
        for r in headline:
            point = {"period": r["period"], "value": tidy(r["value"])}
            for column in ("target", "benchmark"):
                if r[column] is not None:
                    point[column] = tidy(r[column])
            # Status is judged against the Target or the Benchmark, whichever
            # the catalog names. If that number is missing there is no status.
            point["status"] = status_for(r["value"], r[entry["compareTo"]], entry["higherIsBetter"], on_within, off_beyond)
            if r["note"]:
                point["note"] = r["note"]
            series.append(point)

        latest = series[-1]
        for column in ("target", "benchmark"):
            if column not in latest:
                notes.append(
                    f'{entry["id"]}: no {column} for {latest["period"]}. The dashboard will show "{column.title()} not set"'
                    + (" and no status." if column == entry["compareTo"] else ".")
                )
        segments = [
            {"label": r["segment"], "value": tidy(r["value"]), **({"note": r["note"]} if r["note"] else {})}
            for r in rows
            if r["segment"] and r["period"] == latest["period"]
        ]
        entry = {k: v for k, v in entry.items() if v != "" and v != []}
        entry["status"] = latest["status"]
        entry["onTrackWithinPct"] = tidy(on_within)
        entry["offTrackBeyondPct"] = tidy(off_beyond)
        entry["series"] = series
        entry["segments"] = segments
        kpis.append(entry)

    output = {
        "_comment": "Generated by build_kpi_data.py from KPI_Tracking.xlsx. Do not edit by hand.",
        "kpis": kpis,
    }
    OUTPUT.write_text(json.dumps(output, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")

    for w in warnings:
        print(f"WARNING  {w}")
    for n in notes:
        print(f"NOTE     {n}")
    print(f"Wrote {OUTPUT.name}: {len(kpis)} KPI(s).")
    for k in kpis:
        last = k["series"][-1]
        print(f'  {k["id"]:<12} -> {k["level"]:<11} {k["attachTo"]:<5} {item_names[k["attachTo"]]}: '
              f'{last["value"]} in {last["period"]} ({last["status"] or "no target"})')
        for ob in k.get("objectives", []):
            print(f'  {"":<12}    also under Objective {ob} {objective_names[ob]}')


def fail(problems):
    print(f"kpi-data.json was NOT updated. {len(problems)} problem(s) to fix in the spreadsheet:\n")
    for p in problems:
        print(f"  - {p}")
    sys.exit(1)


if __name__ == "__main__":
    build()
