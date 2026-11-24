"""
Excel → data converter for the WTO Digitalization Compendium
(Empowering Women to Trade through Digitalization: Policy Insights Compendium).

Reads the two workbooks and writes the four data files used by the page:

    downloads/digitalization-compendium.xlsx
        sheets "Questions", "WTO Members", "Other Stakeholders"
        → data/digitalization_data.json       the survey responses, as loaded by index.html
        → data/digitalization_data.js         the same data as a script (see below)

    downloads/digitalization-policies.xlsx
        sheets "Policies", "Labels"
        → data/digitalization_policies.json   the policy catalogue (Policy view), as loaded by index.html
        → data/digitalization_policies.js     the same catalogue as a script (see below)

The .js files are used only when the page is opened from disk: browsers do not let a page opened
from disk read a JSON file.

Cell values are copied verbatim. The only processing: multiple-choice cells (one option per
line) become lists of options, dates become "YYYY-MM-DD hh:mm:ss" text, and in the policies
workbook the labels (instrument type, policy areas, scope, target groups) become the short codes
listed in its Labels sheet, items ("• " lines) and links (one per line) become lists.

The policies workbook is checked against the compendium: every respondent, source question and
label must exist, or nothing is written. A policy whose description no longer appears word for
word in that respondent's answers is reported as a warning.

Usage (from any folder):
    python3 tools/excel_to_json.py
    python3 tools/excel_to_json.py --input path/to/compendium.xlsx --policies path/to/policies.xlsx

Requirements:
    Python 3.8+ and openpyxl  (pip install -r tools/requirements.txt)

Run it every time the Excel file changes, then refresh index.html.
"""

import argparse
import json
import os
import re
import sys
from collections import Counter
from datetime import date, datetime

try:
    import openpyxl
except ImportError:
    sys.exit("ERROR: openpyxl is required.  Install it with:  pip install openpyxl")


PACKAGE_DIR    = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DEFAULT_INPUT  = os.path.join(PACKAGE_DIR, "downloads", "digitalization-compendium.xlsx")
DEFAULT_POLICIES = os.path.join(PACKAGE_DIR, "downloads", "digitalization-policies.xlsx")
DEFAULT_OUTDIR = os.path.join(PACKAGE_DIR, "data")
JS_VARIABLE    = "DIGITALIZATION_COMPENDIUM_DATA"   # index.html reads window.DIGITALIZATION_COMPENDIUM_DATA
POLICIES_JS_VARIABLE = "DIGITALIZATION_POLICIES"    # index.html reads window.DIGITALIZATION_POLICIES

# Policies workbook: Labels sheet "Label" column → key in the catalogue's meta (in this order)
LABEL_KINDS = [("Instrument type", "types"), ("Policy area", "areas"), ("Target group", "groups"), ("Scope", "scopes")]

TITLE = "Empowering Women to Trade through Digitalization: Policy Insights Compendium"

# Respondent sheets: sheet name → (group id, name of the added region/category column)
SHEETS = {
    "WTO Members":        ("members",      "Region"),
    "Other Stakeholders": ("stakeholders", "Category"),
}

# Column header "Q1 | …" (answer), "Q1.1 | …" or "S1.a | …" (explanation / details)
CODE_HEADER = re.compile(r"^(?P<code>[A-Z]\w*?)(?:\.(?P<part>1|[a-h]))? \| ")


def text(v):
    """Cell value → text, verbatim ('' for empty cells, dates as YYYY-MM-DD hh:mm:ss)."""
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def read_questions(wb):
    ws = wb["Questions"]
    header = [text(c.value) for c in ws[1]]
    idx = {h: i for i, h in enumerate(header)}
    questions = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[idx["Code"]]:
            continue
        prompts = text(row[idx["Explanation prompt(s)"]])
        questions.append({
            "code":     text(row[idx["Code"]]),
            "group":    "members" if text(row[idx["Respondents"]]) == "WTO Members" else "stakeholders",
            "theme":    text(row[idx["Theme"]]),
            "question": text(row[idx["Question"]]),
            "prompts":  [p for p in prompts.split("\n") if p.strip()],
            "kind":     "text" if text(row[idx["Answer type"]]).lower().startswith("free") else "options",
        })
    return questions


def read_respondents(ws, group, area_header, questions):
    """One sheet → respondents. Each respondent has:
         info     the response's own identification fields, verbatim (e.g. Time, Ministry/Institution:)
         answers  {code: {"options": [...] (option questions), "main": "..." (free-text questions),
                          "notes": ["...", ...] (explanations, in the order of the question's prompts)}}"""
    header = [text(c.value) for c in ws[1]]
    kinds = {q["code"]: q["kind"] for q in questions if q["group"] == group}
    respondents = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        if not any(v not in (None, "") for v in values):
            continue
        r = {"id": "", "group": group, "name": "", "area": "", "info": [], "answers": {}}
        for h, v in zip(header, values):
            m = CODE_HEADER.match(h)
            if h == "Respondent":
                r["name"] = text(v)
            elif h == area_header:
                r["area"] = text(v)
            elif not m:
                r["info"].append({"label": h, "value": text(v)})
            elif m.group("code") not in kinds:
                print(f"WARNING: column '{h[:60]}' in '{ws.title}' has no entry in the Questions sheet; skipped")
            else:
                a = r["answers"].setdefault(m.group("code"), {"options": [], "main": "", "notes": []})
                if m.group("part"):
                    a["notes"].append(text(v))
                elif kinds[m.group("code")] == "text":
                    a["main"] = text(v)
                else:
                    a["options"] = [o.strip() for o in text(v).split("\n") if o.strip()]
        r["id"] = group[0] + "-" + re.sub(r"[^a-z0-9]+", "-", r["name"].lower()).strip("-")
        respondents.append(r)
    return respondents


def norm_ws(s):
    return re.sub(r"\s+", " ", s or "").strip()


def read_policies(path, data):
    """Policies workbook → the Policy view catalogue {"meta": {...}, "policies": [...]}.
    Stops with an error listing every problem if a respondent, question or label is unknown."""
    wb = openpyxl.load_workbook(path, data_only=True)
    for sheet in ("Policies", "Labels"):
        if sheet not in wb.sheetnames:
            sys.exit(f"ERROR: sheet '{sheet}' not found in {path}")

    meta = {key: {} for _, key in LABEL_KINDS}
    kinds = dict(LABEL_KINDS)
    for row in wb["Labels"].iter_rows(min_row=2, values_only=True):
        if row and row[0]:
            kind, code, meaning = text(row[0]), text(row[1]), text(row[2])
            if kind not in kinds:
                sys.exit(f"ERROR: Labels sheet: unknown label kind '{kind}'")
            meta[kinds[kind]][code] = meaning
    code_of = {key: {v: k for k, v in labels.items()} for key, labels in meta.items()}

    by_name = {r["name"]: r for r in data["respondents"]}
    q_order = [q["code"] for q in data["meta"]["questions"]]
    answers_text = {r["id"]: norm_ws(" ".join(" ".join([a["main"]] + a["notes"]) for a in r["answers"].values()))
                    for r in data["respondents"]}

    ws = wb["Policies"]
    header = [text(c.value) for c in ws[1]]
    need = {"ID": "id", "Respondent": "respondent", "WTO Member or place where the policy applies": "place",
            "Policy name (as written by the respondent)": "name", "Instrument type": "type", "Policy areas": "areas",
            "Scope": "scope", "Target groups": "groups", "Description (quoted from the answer)": "description",
            "Items listed by the respondent": "items", "Links given by the respondent": "links",
            "Source question code": "source", "Policy also named in answers to": "also"}
    missing = [h for h in need if h not in header]
    if missing:
        sys.exit("ERROR: Policies sheet is missing the column(s): " + ", ".join(missing))
    col = {need[h]: header.index(h) for h in need}

    errors, warnings, policies, seen = [], [], [], set()
    for n, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        v = {k: text(row[i]) for k, i in col.items()}
        if not v["id"] and not v["name"]:
            continue
        where = f"Policies row {n} ({v['id'] or v['name'][:40]})"

        def codes(kind, cell, many):
            parts = [x.strip() for x in cell.split(";")] if many else [cell.strip()]
            out = []
            for x in parts:
                if not x:
                    continue
                if x not in code_of[kind]:
                    errors.append(f"{where}: unknown {kind[:-1]} '{x}' (not in the Labels sheet)")
                else:
                    out.append(code_of[kind][x])
            return out

        r = by_name.get(v["respondent"])
        if not r:
            errors.append(f"{where}: respondent '{v['respondent']}' is not in the compendium workbook")
            continue
        if v["id"] in seen:
            errors.append(f"{where}: duplicate ID")
        seen.add(v["id"])
        if v["source"] not in q_order:
            errors.append(f"{where}: unknown source question '{v['source']}'")
            continue
        also = [c.strip() for c in v["also"].split(",") if c.strip()]
        bad = [c for c in also if c not in q_order]
        if bad:
            errors.append(f"{where}: unknown question(s) in 'Policy also named in answers to': {', '.join(bad)}")
            continue
        type_ = codes("types", v["type"], False)
        scope = codes("scopes", v["scope"], False)
        if not (v["name"] and v["description"] and v["type"] and v["scope"]):
            errors.append(f"{where}: name, description, instrument type and scope are required")
            continue
        if not (type_ and scope):
            continue
        items = [x[2:] if x.startswith("• ") else x for x in v["items"].split("\n") if x.strip()]
        for t in [v["name"], v["description"]] + items:
            if norm_ws(t) not in answers_text[r["id"]]:
                warnings.append(f"{where}: not found word for word in the answers of {r['name']}: '{t[:60]}…'")
        policies.append({
            "id": v["id"],
            "respondent": r["id"],
            "name": v["name"],
            "type": type_[0],
            "areas": codes("areas", v["areas"], True),
            "groups": codes("groups", v["groups"], True),
            "scope": scope[0],
            # for WTO Members the place column repeats the Member's name; only stakeholders' places are stored
            "country": "" if r["group"] == "members" and v["place"] in ("", r["name"]) else v["place"],
            "source": v["source"],
            "description": v["description"],
            "items": items,
            "links": [x.strip() for x in v["links"].split("\n") if x.strip()],
            "mentioned_in": sorted(set(also + [v["source"]]), key=q_order.index),
        })
    if errors:
        sys.exit("ERROR: nothing written. Fix these in " + os.path.basename(path) + ":\n  " + "\n  ".join(errors))
    for w in warnings:
        print("WARNING: " + w)
    return {"meta": meta, "policies": policies}


def write_data(outdir, stem, variable, comment, obj):
    json_str = json.dumps(obj, ensure_ascii=False, indent=1)
    os.makedirs(outdir, exist_ok=True)
    json_path = os.path.join(outdir, stem + ".json")
    js_path = os.path.join(outdir, stem + ".js")
    with open(json_path, "w", encoding="utf-8") as f:
        f.write(json_str + "\n")
    with open(js_path, "w", encoding="utf-8") as f:
        f.write(comment)
        f.write(f"window.{variable} = {json_str};\n")
    print(f"Written:  {json_path}")
    print(f"Written:  {js_path}")


def read_compendium(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    questions = read_questions(wb)
    respondents = []
    for sheet, (group, area_header) in SHEETS.items():
        respondents += read_respondents(wb[sheet], group, area_header, questions)
    themes = []
    for q in questions:
        if q["theme"] not in themes:
            themes.append(q["theme"])
    return {
        "meta": {
            "title": TITLE,
            "generated_at": date.today().isoformat(),
            "groups": {"members": "WTO Members", "stakeholders": "WTO stakeholders"},
            "themes": themes,
            "questions": questions,
        },
        "respondents": respondents,
    }


def main():
    parser = argparse.ArgumentParser(description="Convert the two Digitalization Compendium Excel files into the page's data files (data/*.json and .js)")
    parser.add_argument("--input",    "-i", default=DEFAULT_INPUT,    help="compendium Excel file (default: downloads/digitalization-compendium.xlsx)")
    parser.add_argument("--policies", "-p", default=DEFAULT_POLICIES, help="policies Excel file (default: downloads/digitalization-policies.xlsx)")
    parser.add_argument("--outdir",   "-o", default=DEFAULT_OUTDIR,   help="output folder (default: data/)")
    args = parser.parse_args()
    for path in (args.input, args.policies):
        if not os.path.isfile(path):
            sys.exit(f"ERROR: Input file not found: {path}")

    print(f"Reading:  {args.input}")
    data = read_compendium(args.input)
    print(f"Reading:  {args.policies}")
    catalogue = read_policies(args.policies, data)   # checked before anything is written

    write_data(args.outdir, "digitalization_data", JS_VARIABLE,
               "/* Digitalization Compendium data.\n"
               "   Generated from downloads/digitalization-compendium.xlsx by tools/excel_to_json.py: do not edit by hand. */\n", data)
    write_data(args.outdir, "digitalization_policies", POLICIES_JS_VARIABLE,
               "/* Digitalization Compendium policy catalogue.\n"
               "   Generated from downloads/digitalization-policies.xlsx by tools/excel_to_json.py: do not edit by hand. */\n", catalogue)

    groups = Counter(r["group"] for r in data["respondents"])
    print(f"Responses: {groups['members']} WTO Members, {groups['stakeholders']} other stakeholders; "
          f"{len(data['meta']['questions'])} questions in {len(data['meta']['themes'])} themes")
    print(f"Policies:  {len(catalogue['policies'])} from {len({p['respondent'] for p in catalogue['policies']})} respondents")


if __name__ == "__main__":
    main()
