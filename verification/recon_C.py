"""Reconcile the independent COMEXT extract with the UN Comtrade extract used by both papers, and show the CN8
composition of the chapters that drove the published result. Writes recon_C.json and prints tables."""
import json, csv, glob, os, sys, io, collections
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import pandas as pd
R = os.environ.get('EUMG_PACKAGE', os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
FX = {int(k): v for k, v in json.load(open('ecb_usd_per_eur.json')).items()}
M = 1e6
PARTNERS = {32: 'AR', 68: 'BO', 76: 'BR', 600: 'PY', 858: 'UY'}

# COMEXT
cx = pd.read_csv('comext_cn8_tidy.csv', dtype={'cn8': str})
cx['usd'] = cx['value_eur'] * cx['year'].map(FX)
cx['hs2'] = cx['cn8'].str[:2]; cx['hs6'] = cx['cn8'].str[:6]
# Comtrade HS2 raw extract (both papers) and HS6 files (v8.1)
raw = pd.read_excel(os.path.join(R, 'model/inputs/Greece__Mar_Total_TradeData_Balance_Sheet.xlsx'), 'Sheet1')
raw['hs2'] = raw['cmdCode'].astype(str).str.zfill(2); raw['partner'] = raw['partnerCode'].map(PARTNERS)
ct6 = []
for f in glob.glob(os.path.join(R, 'model/inputs/comtrade_hs6/ct6_*_*_*.json')):
    _, flow, year, p = os.path.basename(f)[:-5].split('_')
    for r in json.load(open(f, encoding='utf-8')).get('data', []):
        ct6.append({'flow': flow, 'year': int(year), 'partner': PARTNERS[int(p)], 'hs6': r['cmdCode'], 'usd': r.get('primaryValue') or 0, 'kg': r.get('netWgt') or 0})
ct6 = pd.DataFrame(ct6)

out = {}
print('=== 1. Chapter totals 2022-24 mean, million USD: COMEXT (all 5 partners) vs Comtrade extract ===')
a = cx[(cx.year.between(2022, 2024))].groupby(['hs2', 'flow'])['usd'].sum() / 3 / M
b = raw[raw.refYear.between(2022, 2024)].groupby(['hs2', 'flowCode'])['primaryValue'].sum() / 3 / M
tab = pd.concat([a.rename('COMEXT'), b.rename('Comtrade')], axis=1).fillna(0)
tab['diff_pct'] = (tab.COMEXT / tab.Comtrade.replace(0, float('nan')) - 1) * 100
for fl in ['M', 'X']:
    t = tab.xs(fl, level=1).sort_values('Comtrade', ascending=False)
    print(f'-- {"Imports" if fl=="M" else "Exports"} (top 20 by Comtrade):  total COMEXT {t.COMEXT.sum():.1f}  Comtrade {t.Comtrade.sum():.1f}')
    print(t.head(20).round(2).to_string())
    out[f'chapter_{fl}'] = t.round(3).reset_index().to_dict('records')
# totals over the full panel
print('\n=== 2. Annual totals, million USD (5 partners) ===')
ay = cx.groupby(['year', 'flow'])['usd'].sum().unstack() / M
by = raw.groupby(['refYear', 'flowCode'])['primaryValue'].sum().unstack() / M
print(pd.concat([ay.add_prefix('COMEXT_'), by.add_prefix('Comtrade_')], axis=1).round(1).to_string())

print('\n=== 3. CN8 composition of the chapters that carried the published result (imports, 2022-24 mean, 4 parties) ===')
comp = {}
for hs in ['23', '09', '12', '24', '27', '26', '47', '03', '20', '22', '08', '02']:
    d = cx[(cx.hs2 == hs) & (cx.flow == 'M') & cx.year.between(2022, 2024) & (cx.partner != 'BO')]
    g = d.groupby('cn8')['usd'].sum() / 3 / M
    tot = g.sum()
    top = g.sort_values(ascending=False).head(6)
    comp[hs] = [(c, round(v, 3), round(v / tot * 100, 1)) for c, v in top.items()]
    print(f'HS {hs}: total {tot:.2f}m; ' + '; '.join(f'{c} {v:.2f}m ({v/tot*100:.1f}%)' for c, v in top.items()))
out['composition'] = comp

print('\n=== 4. HS6 check vs Comtrade HS6 (2022-24 mean imports, 4 parties): top lines ===')
c6 = cx[(cx.flow == 'M') & cx.year.between(2022, 2024) & (cx.partner != 'BO')].groupby('hs6')['usd'].sum() / 3 / M
t6 = ct6[(ct6.flow == 'M') & (ct6.partner != 'BO')].groupby('hs6')['usd'].sum() / 3 / M
tt = pd.concat([c6.rename('COMEXT'), t6.rename('Comtrade')], axis=1).fillna(0).sort_values('Comtrade', ascending=False)
print(tt.head(25).round(2).to_string())
out['hs6_top'] = tt.head(40).round(3).reset_index().to_dict('records')
json.dump(out, open('recon_C.json', 'w'), indent=1, default=str)
