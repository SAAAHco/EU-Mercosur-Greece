"""Build the HS-6 line table of Greek trade with Mercosur (2022-2024 mean) for the
25 modelled chapters, and compare its chapter sums with the model's baselines."""
import json, glob, os, collections
os.chdir(os.path.dirname(os.path.abspath(__file__)))
CFG = r"F:/2025/Kostas/EU-Mercosur/Analysis/05172026/Final/Updated/FInal Outputs/ecological-economics/EU-Mercosur-Greece_github_upload/EU-Mercosur-Greece/data/v5_config.json"
cfg = json.load(open(CFG))
chapters = {c['hs']: c for c in cfg['chapters']}
PARTNERS = {32: 'ARG', 68: 'BOL', 76: 'BRA', 600: 'PRY', 858: 'URY'}
rows = collections.defaultdict(float)  # (flow, hs6, partner) -> sum over 3 years
for f in glob.glob('ct6_*_*_*.json'):
    _, flow, year, p = f[:-5].split('_')
    d = json.load(open(f))
    for r in d.get('data', []):
        rows[(flow, r['cmdCode'], PARTNERS[int(p)])] += (r.get('primaryValue') or 0) / 3.0
lines = []
for (flow, hs6, p), v in rows.items():
    if hs6[:2] in chapters and v > 0:
        lines.append({'flow': flow, 'hs6': hs6, 'hs2': hs6[:2], 'partner': p, 'value': v})
json.dump(lines, open('lines_2022_2024.json', 'w'), indent=0)
agg = collections.defaultdict(float)
for l in lines:
    agg[(l['flow'], l['hs2'])] += l['value']
print(f"{'HS':4s} {'model_imp':>10s} {'hs6_imp':>10s} {'model_exp':>10s} {'hs6_exp':>10s}")
for hs, c in sorted(chapters.items(), key=lambda kv: -kv[1]['baseline_imp']):
    print(f"{hs:4s} {c['baseline_imp']/1e6:10.2f} {agg[('M', hs)]/1e6:10.2f} {c['baseline_exp']/1e6:10.2f} {agg[('X', hs)]/1e6:10.2f}")
print('lines', len(lines), 'imports lines', sum(1 for l in lines if l['flow'] == 'M'), 'exports lines', sum(1 for l in lines if l['flow'] == 'X'))
