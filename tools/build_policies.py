"""Build the Policy view catalogue from the curated files in tools/policy_catalogue/ (Trade and Gender Office use).

The catalogue is curated by hand in tools/policy_catalogue/*.py (one file per respondent).
Every entry points at the respondent's own words; nothing is paraphrased:
  n  policy name, exactly as written by the respondent (must appear verbatim in their answers)
  t  instrument type code (TYPES)
  a  policy area codes, space separated (AREAS)
  g  optional target group codes, space separated (GROUPS)
  s  scope: "D" domestic, "I" international / development cooperation
  q  answer code the description is taken from (e.g. "Q9")
  x  (start, end): description = the verbatim text in that answer from `start` through `end`
  i  optional (code, start, end): a listed block in an answer, split into one item per line
  c  optional: country or region the policy applies to, when it is not the respondent (stakeholder answers)
  l  optional URL fragments; each must match exactly one URL written in the respondent's answers

It writes downloads/digitalization-policies.xlsx: one row per policy with its labels, the quoted
description and, next to it, the full source question and the respondent's complete answer, verbatim.
That Excel is the source of the page's catalogue: the script converts it with tools/excel_to_json.py
(data/digitalization_policies.json and .js), after checking the round trip is exact. Later edits can be
made in the Excel itself and converted with tools/excel_to_json.py.

Run:  python3 tools/build_policies.py
The script stops with an error if any name, description, item block or link is not found verbatim.
"""
import glob, importlib.util, json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
DATA = os.path.join(ROOT, 'data', 'digitalization_data.json')
OUT_JSON = os.path.join(ROOT, 'data', 'digitalization_policies.json')
OUT_JS = os.path.join(ROOT, 'data', 'digitalization_policies.js')
OUT_XLSX = os.path.join(ROOT, 'downloads', 'digitalization-policies.xlsx')

TYPES = {
    'STR': 'Strategy, plan or policy framework',
    'LAW': 'Law, regulation or budget measure',
    'FTA': 'Trade agreement provision',
    'PRG': 'Programme or initiative',
    'TRN': 'Training or capacity building',
    'FIN': 'Finance, funding or investment',
    'PLT': 'Platform, portal or digital tool',
    'DAT': 'Data, statistics or research',
    'NET': 'Network, partnership or coordination body',
}
AREAS = {
    'ecom':   'E-commerce and market access',
    'skills': 'Digital skills and training',
    'fin':    'Access to finance and digital payments',
    'ent':    'Entrepreneurship, start-ups and innovation',
    'tf':     'Trade facilitation and logistics',
    'infra':  'Digital infrastructure and connectivity',
    'stem':   'STEM, ICT and digital careers',
    'net':    'Mentoring and business networks',
    'rules':  'Trade agreements and digital trade rules',
    'gender': 'Gender mainstreaming and equality frameworks',
    'data':   'Data, evidence and monitoring',
    'cyber':  'Online safety and cybersecurity',
    'coord':  'Coordination and governance',
}
GROUPS = {
    'msme':   'Women entrepreneurs and women-led MSMEs',
    'rural':  'Rural women',
    'indig':  'Indigenous women',
    'youth':  'Young women and girls',
    'disab':  'Women with disabilities',
    'work':   'Women workers and professionals',
    'gov':    'Policymakers and officials',
}
SCOPES = {'D': 'Domestic', 'I': 'International cooperation'}

URL_RE = re.compile(r'https?://[^\s<>"]+[^\s<>".,;:)\]]')
WWW_RE = re.compile(r'(?<![/\w.@])(www\.[\w-]+(?:\.[\w-]+)+(?:/[^\s<>"\])]*)?)')
EMAIL_RE = re.compile(r'[\w.+-]+@[\w-]+\.[\w.-]+')
PHONE_RE = re.compile(r'\+?\d[\d ()-]{8,}\d')


def norm(s):
    return re.sub(r'\s+', ' ', s).strip()


def answer_texts(a):
    return [t for t in [a.get('main', '')] + a.get('notes', []) if t and t.strip()]


def find_span(texts, start, end):
    """Return the verbatim text from `start` through `end` within one of `texts` (whitespace-tolerant)."""
    for t in texts:
        nt = norm(t)
        i = nt.find(norm(start))
        if i == -1:
            continue
        j = nt.find(norm(end), i)
        if j == -1:
            continue
        return nt[i:j + len(norm(end))]
    return None


def load_specs():
    specs = {}
    for path in sorted(glob.glob(os.path.join(HERE, 'policy_catalogue', '*.py'))):
        name = os.path.splitext(os.path.basename(path))[0]
        spec = importlib.util.spec_from_file_location(name, path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        specs[mod.RESPONDENT] = mod.P
    return specs


def main():
    data = json.load(open(DATA, encoding='utf-8'))
    resp = {r['id']: r for r in data['respondents']}
    errors, out = [], []
    for rid, entries in load_specs().items():
        r = resp.get(rid)
        if not r:
            errors.append('unknown respondent ' + rid); continue
        all_text = norm(' '.join(t for a in r['answers'].values() for t in answer_texts(a)))
        blob = ' '.join(t for a in r['answers'].values() for t in answer_texts(a))
        urls, seen = [], set()
        for u in URL_RE.findall(blob) + ['https://' + w for w in WWW_RE.findall(blob)]:
            key = re.sub(r'^https?://', '', u).rstrip('/').lower()
            if key not in seen:
                seen.add(key); urls.append(u)
        for k, e in enumerate(entries, 1):
            tag = '%s #%d %r' % (rid, k, e.get('n'))
            if norm(e['n']) not in all_text:
                errors.append(tag + ': name not found verbatim')
            if e['t'] not in TYPES: errors.append(tag + ': bad type ' + e['t'])
            for c in e['a'].split():
                if c not in AREAS: errors.append(tag + ': bad area ' + c)
            for c in e.get('g', '').split():
                if c not in GROUPS: errors.append(tag + ': bad group ' + c)
            if e['s'] not in SCOPES: errors.append(tag + ': bad scope')
            a = r['answers'].get(e['q'])
            desc = find_span(answer_texts(a), *e['x']) if a else None
            if not desc:
                errors.append(tag + ': description not found in ' + e['q']); continue
            items = []
            if e.get('i'):
                ic, ist, ien = e['i']
                ia = r['answers'].get(ic)
                block = None
                if ia:
                    for t in answer_texts(ia):
                        i0 = t.find(ist)
                        j0 = t.find(ien, i0) if i0 != -1 else -1
                        if i0 != -1 and j0 != -1:
                            block = t[i0:j0 + len(ien)]; break
                if not block:
                    errors.append(tag + ': item block not found in ' + ic)
                else:
                    items = [re.sub(r'^[\s\-•·*•\d.)]+', '', ln).strip() for ln in block.split('\n')]
                    items = [norm(x) for x in items if x.strip() and not re.match(r'^https?://\S+;?$', x.strip())]
            links = []
            for frag in e.get('l', []):
                hits = [u for u in urls if (u.endswith(frag[:-1]) if frag.endswith('$') else frag in u)]
                if len(hits) != 1:
                    errors.append('%s: link fragment %r matched %d URLs' % (tag, frag, len(hits)))
                else:
                    links.append(hits[0])
            for txt in [desc] + items:
                if EMAIL_RE.search(txt) or PHONE_RE.search(txt):
                    errors.append(tag + ': possible personal data in text')
            mentioned = [c for c, ans in r['answers'].items()
                         if any(norm(e['n']) in norm(t) for t in answer_texts(ans))]
            out.append({
                'id': '%s-%d' % (rid, k),
                'respondent': rid,
                'name': norm(e['n']),
                'type': e['t'],
                'areas': e['a'].split(),
                'groups': e.get('g', '').split(),
                'scope': e['s'],
                'country': e.get('c', ''),
                'source': e['q'],
                'description': desc,
                'items': items,
                'links': links,
                'mentioned_in': sorted(set(mentioned + [e['q']]), key=lambda c: list(r['answers']).index(c)),
            })
    if errors:
        print('\n'.join(errors)); print('\n%d error(s); nothing written.' % len(errors)); sys.exit(1)
    catalogue = {'meta': {'types': TYPES, 'areas': AREAS, 'groups': GROUPS, 'scopes': SCOPES}, 'policies': out}
    # The Excel is the source of the page's catalogue (as for the compendium): write it, then convert it
    # with tools/excel_to_json.py and check the round trip is exact.
    write_excel(catalogue, data, OUT_XLSX)
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    import excel_to_json
    back = excel_to_json.read_policies(OUT_XLSX, data)
    if back != catalogue:
        sys.exit('ERROR: the policies Excel does not convert back to the catalogue; data files not written.')
    excel_to_json.write_data(os.path.dirname(OUT_JSON), 'digitalization_policies', excel_to_json.POLICIES_JS_VARIABLE,
                             '/* Digitalization Compendium policy catalogue.\n'
                             '   Generated from downloads/digitalization-policies.xlsx by tools/excel_to_json.py: do not edit by hand. */\n',
                             back)
    print('%d policies from %d respondents written.' % (len(out), len({p["respondent"] for p in out})))


def write_excel(catalogue, data, path):
    """One row per policy, with the full source question and the respondent's complete answer (verbatim)."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    resp = {r['id']: r for r in data['respondents']}
    qs = {q['code']: q for q in data['meta']['questions']}
    groups = data['meta']['groups']
    m = catalogue['meta']
    head_fill, head_font = PatternFill('solid', fgColor='006072'), Font(bold=True, color='FFFFFF')
    wrap = Alignment(wrap_text=True, vertical='top')

    def sheet(ws, header, rows, widths):
        ws.append(header)
        for c in ws[1]:
            c.fill, c.font, c.alignment = head_fill, head_font, Alignment(wrap_text=True, vertical='center')
        for row in rows:
            ws.append(row)
        for col, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = w
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.alignment = wrap
        ws.freeze_panes = 'C2' if len(header) > 6 else 'A2'
        ws.auto_filter.ref = ws.dimensions

    def full_answer(r, code):
        a, q = r['answers'].get(code, {}), qs[code]
        parts = []
        if a.get('options'):
            parts.append('Answer: ' + ' | '.join(a['options']))
        if (a.get('main') or '').strip():
            parts.append(a['main'].strip())
        for i, t in enumerate(a.get('notes', [])):
            if t and t.strip():
                prompt = q['prompts'][i] if i < len(q.get('prompts', [])) else 'Explanation'
                parts.append(prompt + '\n' + t.strip())
        return '\n\n'.join(parts)

    wb = Workbook()
    ws = wb.active
    ws.title = 'Policies'
    rows = []
    for p in catalogue['policies']:
        r = resp[p['respondent']]
        q = qs[p['source']]
        rows.append([
            p['id'], r['name'], groups[r['group']], r['area'], p['country'] or (r['name'] if r['group'] == 'members' else ''),
            p['name'], m['types'][p['type']], '; '.join(m['areas'][a] for a in p['areas']), m['scopes'][p['scope']],
            '; '.join(m['groups'][g] for g in p['groups']),
            p['description'], '\n'.join('• ' + x for x in p['items']), '\n'.join(p['links']),
            p['source'], q['theme'], q['question'], ' | '.join(r['answers'].get(p['source'], {}).get('options', [])),
            full_answer(r, p['source']), ', '.join(c for c in p['mentioned_in'] if c != p['source']),
        ])
    sheet(ws, ['ID', 'Respondent', 'Respondent type', 'Region / Category', 'WTO Member or place where the policy applies',
               'Policy name (as written by the respondent)', 'Instrument type', 'Policy areas', 'Scope', 'Target groups',
               'Description (quoted from the answer)', 'Items listed by the respondent', 'Links given by the respondent',
               'Source question code', 'Survey theme', 'Source question (full wording)', 'Answer options selected',
               'Full answer to the source question (verbatim)', 'Policy also named in answers to'],
          rows, [14, 24, 16, 20, 22, 40, 24, 32, 18, 26, 70, 45, 40, 10, 24, 55, 40, 100, 18])

    ws2 = wb.create_sheet('Labels')
    lab = [['Instrument type', k, v] for k, v in m['types'].items()] + [['Policy area', k, v] for k, v in m['areas'].items()] + \
          [['Scope', k, v] for k, v in m['scopes'].items()] + [['Target group', k, v] for k, v in m['groups'].items()]
    sheet(ws2, ['Label', 'Code', 'Meaning'], lab, [18, 10, 50])

    ws3 = wb.create_sheet('About', 0)
    about = [
        ['Empowering Women to Trade through Digitalisation: Policy Insights'],
        ['Policy catalogue'],
        [''],
        ['This workbook lists the policies, strategies, laws, trade agreement provisions, programmes, platforms, funds and networks that WTO Members and WTO stakeholders name in their survey responses (the Policy view of the Compendium).'],
        ['Each row is one policy. The policy name, the description, the listed items and the links are quoted word for word from the response. The instrument type, policy areas, scope and target groups are labels added to help browse the catalogue; they are not the respondents\' words (see the Labels sheet).'],
        ['For each policy, the last columns give the survey question its description is taken from, in full, and the respondent\'s complete answer to that question, verbatim: the options selected and every explanation. The same policy may also be named in other answers (last column).'],
        ['Personal data has been removed from the responses.'],
        [''],
        ['Policies: %d from %d respondents.' % (len(catalogue['policies']), len({p['respondent'] for p in catalogue['policies']}))],
        ['Generated by tools/build_policies.py from the Compendium data (%s).' % data['meta'].get('generated_at', '')],
    ]
    for row in about:
        ws3.append(row)
    ws3.column_dimensions['A'].width = 120
    ws3['A1'].font = Font(bold=True, size=14, color='003E4A')
    ws3['A2'].font = Font(bold=True, size=12, color='006072')
    for row in ws3.iter_rows(min_row=3):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical='top')
    wb.active = 1
    wb.save(path)


if __name__ == '__main__':
    main()
