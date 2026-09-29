"""Pull applied MFN ad valorem equivalents (WITS TRAINS, AVE-estimated, falling back to
reported) at HS-6 for the EU (imports) and each Mercosur party (Greek exports)."""
import json, re, time, subprocess, os, sys
os.chdir(os.path.dirname(os.path.abspath(__file__)))
L = json.load(open('lines44.json'))['lines']
REPORTERS = {'EU': '918', 'ARG': '032', 'BRA': '076', 'PRY': '600', 'URY': '858'}
need = {'EU': sorted({l['hs6'] for l in L if l['flow'] == 'M'})}
for p in ['ARG', 'BRA', 'PRY', 'URY']:
    need[p] = sorted({l['hs6'] for l in L if l['flow'] == 'X'})
out = json.load(open('wits_rates.json')) if os.path.exists('wits_rates.json') else {}


def fetch(rep, prods, year, dtype):
    url = (f"https://wits.worldbank.org/API/V1/SDMX/V21/datasource/TRN/reporter/{rep}/partner/000/"
           f"product/{';'.join(prods)}/year/{year}/datatype/{dtype}")
    t = subprocess.run(['curl', '-s', '--max-time', '120', url], capture_output=True, text=True).stdout
    res = {}
    for s in re.findall(r'<Series.*?</Series>', t, re.S):
        p = re.search(r'PRODUCTCODE="(\d+)"', s).group(1)
        o = re.search(r'<Obs[^>]*/>', s)
        if not o:
            continue
        o = o.group(0)
        g = lambda k: (re.search(k + r'="([^"]*)"', o) or [None, None])[1]
        res[p] = {'rate': float(g('OBS_VALUE')), 'lines': int(g('TOTALNOOFLINES') or 0), 'na': int(g('NBR_NA_LINES') or 0),
                  'min': g('MIN_RATE'), 'max': g('MAX_RATE'), 'year': year, 'dtype': dtype}
    return res


for party, prods in need.items():
    rep = REPORTERS[party]
    out.setdefault(party, {})
    todo = [p for p in prods if p not in out[party]]
    for year in ['2023', '2022', '2021', '2024']:
        todo = [p for p in prods if p not in out[party]]
        for i in range(0, len(todo), 20):
            chunk = todo[i:i + 20]
            res = fetch(rep, chunk, year, 'aveestimated')
            # fall back to reported where AVE missing
            miss = [p for p in chunk if p not in res]
            if miss:
                rep_res = fetch(rep, miss, year, 'reported')
                for k, v in rep_res.items():
                    if v['na'] == 0:  # only accept fully ad valorem reported rates
                        res[k] = v
            out[party].update(res)
            time.sleep(1)
        json.dump(out, open('wits_rates.json', 'w'), indent=0)
    missing = [p for p in prods if p not in out[party]]
    print(party, 'rates for', len(out[party]), 'of', len(prods), 'lines; missing', len(missing), missing[:15])
    sys.stdout.flush()
