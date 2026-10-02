"""Parse the COMEXT JSON-stat pulls into one tidy table: year, flow, partner, code (CN8 only), value_eur, qty_100kg."""
import json, gzip, glob, os, sys, io, csv
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
rows = []
for f in sorted(glob.glob('comext_raw/gr_*_*_*.json.gz')):
    _, flow, year, p = os.path.basename(f)[:-8].split('_')
    j = json.load(gzip.open(f, 'rt', encoding='utf-8'))
    ids, size = j['id'], j['size']
    dims = {d: j['dimension'][d]['category']['index'] for d in ids}
    inv = {d: {v: k for k, v in dims[d].items()} for d in ids}
    # strides for row-major linear index
    strides = []
    acc = 1
    for s in reversed(size):
        strides.insert(0, acc); acc *= s
    ip, ii = ids.index('product'), ids.index('indicators')
    rec = {}
    for lin, val in j['value'].items():
        lin = int(lin); pos = []
        for st, s in zip(strides, size):
            pos.append((lin // st) % s)
        prod = inv['product'][pos[ip]]; ind = inv['indicators'][pos[ii]]
        if len(prod) != 8 or 'X' in prod:
            continue   # keep CN8 leaves only (aggregates are recomputed)
        rec.setdefault(prod, {})[ind] = val
    for prod, d in rec.items():
        rows.append((int(year), flow, p, prod, d.get('VALUE_IN_EUROS', 0) or 0, d.get('QUANTITY_IN_100KG', 0) or 0))
with open('comext_cn8_tidy.csv', 'w', newline='', encoding='utf-8') as f:
    w = csv.writer(f); w.writerow(['year', 'flow', 'partner', 'cn8', 'value_eur', 'qty_100kg']); w.writerows(rows)
print('rows', len(rows), 'files', len(glob.glob('comext_raw/gr_*_*_*.json.gz')))
# quick chapter check: 2023 imports by chapter (EUR)
import collections
t = collections.defaultdict(float)
for y, fl, p, c, v, q in rows:
    if y == 2023 and fl == 'M' and p != 'BO':
        t[c[:2]] += v
print('2023 imports (EUR m) by chapter, 4 Mercosur parties:')
for k, v in sorted(t.items(), key=lambda x: -x[1])[:12]:
    print(' ', k, round(v / 1e6, 2))
