"""Build the Policy view catalogue: data/digitalization_policies.json (+ .js for file:// use).

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

Run:  python3 tools/build_policies.py
The script stops with an error if any name, description, item block or link is not found verbatim.
"""
import glob, importlib.util, json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
DATA = os.path.join(ROOT, 'data', 'digitalization_data.json')
OUT_JSON = os.path.join(ROOT, 'data', 'digitalization_policies.json')
OUT_JS = os.path.join(ROOT, 'data', 'digitalization_policies.js')

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
    with open(OUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(catalogue, f, ensure_ascii=False, indent=1)
    with open(OUT_JS, 'w', encoding='utf-8') as f:
        f.write('/* Digitalisation Compendium policy catalogue. Generated by tools/build_policies.py: do not edit by hand. */\n')
        f.write('window.DIGITALIZATION_POLICIES = ' + json.dumps(catalogue, ensure_ascii=False, indent=1) + ';\n')
    print('%d policies from %d respondents written.' % (len(out), len({p["respondent"] for p in out})))


if __name__ == '__main__':
    main()
