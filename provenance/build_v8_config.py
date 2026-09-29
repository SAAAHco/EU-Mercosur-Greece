"""Build v8_config.json: 44 chapters, line-level (HS-6) tariff paths from the EU-Mercosur
Interim Agreement on Trade schedules (OJ L 2026/184, Appendices 2-A-1 and 2-A-2).

Design (documented in CHANGELOG_v8.md):
- Chapter set: every HS-2 chapter with Greece-Mercosur trade above 1 million USD in either
  direction in any year 2014-2024 (the stated rule, 43 chapters), plus HS 04 (dairy) kept for
  the dairy tariff-rate-quota test.
- Baselines: 2022-2024 means from the raw UN Comtrade extract; growth: 2014-2024 endpoint CAGR
  winsorized at +/-15% with no exceptions (zero endpoints: +15% / -15% / 0).
- EU side: tau0 = applied MFN ad valorem equivalent (WITS TRAINS, AVE-estimated, 2023).
  Each HS-6 line's remaining-tariff fraction r(y) is the mean over its CN-8 rows in the EU
  schedule of the category path (first cut at entry into force, per Annex 2-A Sec A para 6).
  TRQ lines: in-quota rate applies only to the share of the line that fits a notional Greek
  slice of the EU quota (2.4% population share x quota volume in year y x Greek import unit
  value); out-of-quota imports keep the MFN rate.
- Mercosur side: tau0 = partner base rate in the Mercosur schedule (ad valorem), averaged over
  NCM rows; same category logic. Bolivia: no change (not a party).
- Lines with no schedule match at HS-6 fall back to HS-4 rows; tariffs missing from WITS fall
  back to the chapter's trade-weighted rate over matched lines.
"""
import json, csv, os, collections, re, math
import pandas as pd
os.chdir(os.path.dirname(os.path.abspath(__file__)))
RAW = r"F:/2025/Kostas/EU-Mercosur/Analysis/Final Analysis/Greece__Mar_Total_TradeData_Balance_Sheet.xlsx"
V5 = r"F:/2025/Kostas/EU-Mercosur/Analysis/05172026/Final/Updated/FInal Outputs/ecological-economics/EU-Mercosur-Greece_github_upload/EU-Mercosur-Greece/data/v5_config.json"
GREEK_SHARE = 0.024
HORIZON = range(11)

D = json.load(open('lines44.json'))
CH = D['chapters']
LINES = D['lines']
W = json.load(open('wits_rates.json'))
v5 = {c['hs']: c for c in json.load(open(V5))['chapters']}
eu_rows = list(csv.DictReader(open('schedules/appendix_2-A-1_EU_schedule_from_OJ_xhtml.csv', encoding='utf-8')))
me_rows = list(csv.DictReader(open('schedules/appendix_2-A-2_MERCOSUR_schedule_from_OJ_xhtml.csv', encoding='utf-8')))
trq = list(csv.DictReader(open('schedules/trq_summary.csv', encoding='utf-8')))
for r in eu_rows:
    r['code'] = re.sub(r'\D', '', r['CN 2013'])
for r in me_rows:
    r['code'] = re.sub(r'\D', '', r['NCM'])

# ---------- unit values (USD per kg) from the Comtrade HS-6 pulls ----------
uv_num, uv_den = collections.defaultdict(float), collections.defaultdict(float)
import glob
for f in glob.glob('ct6_M_*_*.json'):
    for r in json.load(open(f)).get('data', []):
        if r.get('netWgt'):
            uv_num[r['cmdCode']] += r['primaryValue'] or 0
            uv_den[r['cmdCode']] += r['netWgt']
UV = {k: uv_num[k] / uv_den[k] for k in uv_num if uv_den[k] > 0}

# ---------- EU TRQ parameters (Annex 2-A Sec B) ----------
def trq_param(code):
    rows = [t for t in trq if t['opening_party'] == 'EU' and t['notation'].split(' ')[0] == code]
    return rows
EU_TRQ = {
    # code: (in-quota rule, [(year0 t, final t, final year)])
    'BF1': ('ad_valorem', 7.5, [(9075, 54450, 5)]),
    'BF2': ('ad_valorem', 7.5, [(7425, 44550, 5)]),
    'PK': ('specific_ratio', None, [(4167, 25000, 5)]),
    'PY1': ('zero', 0, [(15000, 90000, 5)]),
    'PY2': ('zero', 0, [(15000, 90000, 5)]),
    'MP': ('pref_ramp', (0.10, 10), [(1000, 10000, 10)]),
    'CE': ('pref_ramp', (0.10, 10), [(3000, 30000, 10)]),
    'IF': ('pref_ramp', (0.10, 10), [(500, 5000, 10)]),
    'ME': ('zero', 0, [(166667, 1000000, 5)]),
    'RE': ('zero', 0, [(10000, 60000, 5)]),
    'SR': ('zero', 0, [(180000, 180000, 0)]),
    'OS': ('fraction', 0.5, [(2000, 2000, 0)]),
    'EG1': ('zero', 0, [(500, 3000, 5)]),
    'EG2': ('zero', 0, [(500, 3000, 5)]),
    'HY': ('zero', 0, [(7500, 45000, 5)]),
    'RM': ('zero', 0, [(400, 2400, 5)]),
    'SC': ('zero', 0, [(1000, 1000, 0)]),
    'SH1': ('fraction', 0.5, [(1500, 1500, 0)]),
    'SH2': ('zero', 0, [(100, 600, 5)]),
    'EL': ('fraction', 1 / 3, [(33333, 200000, 5)]),   # all-uses quota: in-quota = one third of MFN (6.4/19.2; 3.4/10.2 EUR/hl)
    'GC': ('pref_ramp', (0.30, 7), [(1875, 15000, 7)]),
}


def quota_t(code, y):
    tot = 0.0
    for q0, qf, fy in EU_TRQ[code][2]:
        tot += qf if fy == 0 else q0 + (qf - q0) * min(1.0, y / fy)
    return tot


def r_linear(P, y):
    """Remaining fraction of the base duty under linear elimination with first cut at EIF."""
    if P <= 0:
        return 0.0
    return max(0.0, 1 - (y + 1) / (P + 1))


def r_category(cat, y):
    """Remaining fraction of the base duty for a non-TRQ staging category; None if TRQ/unknown."""
    c = cat.replace(' ', '')
    if c in ('0',):
        return 0.0
    m = re.fullmatch(r'(\d+)(/EP)?', c)
    if m:
        return r_linear(int(m.group(1)), y)
    m = re.fullmatch(r'0/EP', c)
    if c == '0/EP':
        return 0.0
    if c == 'E':
        return 1.0
    if c.startswith('FP30'):
        return 0.7
    if c.startswith('FP50'):
        return 0.5
    if c == '50%':
        return 1 - 0.5 * min(1.0, (y + 1) / 5)
    if c.startswith('4-EG'):
        return r_linear(4, y)
    if c.startswith('0+EA/10') or c.startswith('10/OS'):
        return r_linear(10, y)
    if c == 'SW/12':
        return 0.0
    if c == 'CE/E':
        return None
    if c in ('15V',):
        return 1.0 if y <= 6 else max(0.0, 1 - [0.19, 0.381, 0.571, 0.643, 0.714, 0.786, 0.857, 0.929, 1.0][min(y - 7, 8)])
    return None


def eu_line_path(hs6, tau0, value_by_year):
    rows = [r for r in eu_rows if r['code'].startswith(hs6)]
    matched = 'hs6'
    if not rows:
        rows = [r for r in eu_rows if r['code'].startswith(hs6[:4])]
        matched = 'hs4' if rows else 'none'
    if not rows:
        return [1.0] * 11, {'matched': 'none', 'cats': []}
    cats = [r['Staging category'].strip() for r in rows]
    paths = []
    trq_codes = []
    for cat in cats:
        code = cat.replace(' ', '').split('/')[0] if cat.replace(' ', '') not in EU_TRQ else cat.replace(' ', '')
        if cat.replace(' ', '') in EU_TRQ:
            trq_codes.append(cat.replace(' ', ''))
            paths.append(('TRQ', cat.replace(' ', '')))
            continue
        p = [r_category(cat, y) for y in HORIZON]
        if p[0] is None:
            p = [1.0] * 11  # unknown category: conservatively no preference
        paths.append(('LIN', p))
    return paths, {'matched': matched, 'cats': cats, 'trq': sorted(set(trq_codes))}


def resolve_eu(hs6, tau0, cf_line_by_year):
    """Return remaining-fraction path r(y) for the line, after TRQ fill."""
    paths, info = eu_line_path(hs6, tau0, None)
    if paths == [1.0] * 11:
        return paths, info
    uv = UV.get(hs6, None)  # USD per kg
    out = []
    for y in HORIZON:
        vals = []
        for kind, p in paths:
            if kind == 'LIN':
                vals.append(p[y])
            else:
                rule, par, _ = EU_TRQ[p]
                if rule == 'zero':
                    r_in = 0.0
                elif rule == 'fraction':
                    r_in = par
                elif rule == 'ad_valorem':
                    r_in = min(1.0, (par / 100) / tau0) if tau0 > 0 else 1.0
                elif rule == 'pref_ramp':
                    p0, fy = par
                    pref = p0 + (1 - p0) * min(1.0, y / fy)
                    r_in = 1 - pref
                elif rule == 'specific_ratio':
                    r_in = 0.5  # pigmeat in-quota 83 EUR/t vs base ~ half; not material for Greece
                else:
                    r_in = 1.0
                slice_usd = GREEK_SHARE * quota_t(p, y) * 1000 * (uv if uv else 3.0)
                m = cf_line_by_year[y]
                f = 1.0 if m <= 0 else min(1.0, slice_usd / m)
                vals.append(1 - f * (1 - r_in))
        out.append(sum(vals) / len(vals))
    return out, info


def me_line_path(hs6, partner):
    col = {'ARG': 'Base rate of Argentina', 'BRA': 'Base rate of Brazil', 'PRY': 'Base rate of Paraguay', 'URY': 'Base rate of Uruguay'}[partner]
    rows = [r for r in me_rows if r['code'].startswith(hs6)]
    matched = 'hs6'
    if not rows:
        rows = [r for r in me_rows if r['code'].startswith(hs6[:4])]
        matched = 'hs4' if rows else 'none'
    if not rows:
        return None, None, {'matched': 'none'}
    bases, paths = [], []
    for r in rows:
        try:
            b = float(r[col].replace(',', '.').replace('%', '').strip())
        except ValueError:
            b = None
        cat = r['Staging category'].strip()
        c = cat.replace(' ', '')
        if c.startswith('TRQ-'):
            # EU-origin dairy/garlic quota: preference 10% (30% garlic) rising to 100%; Greek exports far below any notional slice
            p0, fy = (0.30, 7) if c == 'TRQ-4' else (0.10, 10)
            p = [1 - (p0 + (1 - p0) * min(1.0, y / fy)) for y in HORIZON]
        elif c in ('CH1',):
            p = [r_linear(9, y) for y in HORIZON]
        elif c in ('CH2',):
            p = [r_linear(14, y) for y in HORIZON]
        elif c == 'T1':
            p = [r_linear(9, y) for y in HORIZON]
        elif c.lower().startswith('see') or c == '':
            p = [r_linear(18, y) for y in HORIZON]
        else:
            p = [r_category(cat, y) for y in HORIZON]
            if p[0] is None:
                p = [r_linear(10, y) for y in HORIZON]
        if b is not None:
            bases.append(b)
        paths.append(p)
    base = sum(bases) / len(bases) / 100 if bases else None
    path = [sum(p[y] for p in paths) / len(paths) for y in HORIZON]
    return base, path, {'matched': matched, 'cats': [r['Staging category'] for r in rows][:8]}


# ---------- baselines and CAGRs from the raw extract ----------
raw = pd.read_excel(RAW, 'Sheet1')
raw['cmd'] = raw['cmdCode'].astype(str).str.zfill(2)
g = raw.groupby(['cmd', 'flowCode', 'refYear'])['primaryValue'].sum()


def val(ch, f, y):
    try:
        return float(g.loc[(ch, f, y)])
    except KeyError:
        return 0.0


def cagr(ch, f):
    a, b = val(ch, f, 2014), val(ch, f, 2024)
    if a > 0 and b > 0:
        c = (b / a) ** (1 / 10) - 1
    elif a == 0 and b > 0:
        c = 0.15
    elif a > 0 and b == 0:
        c = -0.15
    else:
        c = 0.0
    return max(-0.15, min(0.15, c))


NAMES = {'02': 'Meat and edible offal', '03': 'Fish and crustaceans', '07': 'Edible vegetables', '10': 'Cereals',
         '16': 'Meat and fish preparations', '17': 'Sugars and sugar confectionery', '21': 'Miscellaneous edible preparations',
         '28': 'Inorganic chemicals', '29': 'Organic chemicals', '32': 'Tanning and dyeing extracts, paints',
         '33': 'Essential oils, perfumery, cosmetics', '40': 'Rubber and articles', '44': 'Wood and articles',
         '48': 'Paper and paperboard', '49': 'Printed books and products', '64': 'Footwear', '74': 'Copper and articles',
         '90': 'Optical, medical instruments', '94': 'Furniture, lighting'}
NEW_WEDGE = {'02': (0.01, 'EUDR cattle (Rijk and Kuepper 2025)'), '40': (0.01, 'EUDR rubber'), '44': (0.01, 'EUDR wood'),
             '48': (0.01, 'EUDR paper'), '49': (0.01, 'EUDR printed products'), '07': (0.01, 'Import-standards compliance, perishables')}

chapters = []
report = []
for ch in CH:
    base_imp = sum(val(ch, 'M', y) for y in (2022, 2023, 2024)) / 3
    base_exp = sum(val(ch, 'X', y) for y in (2022, 2023, 2024)) / 3
    rec = {'hs': ch, 'name': (v5[ch]['name'] if ch in v5 else NAMES.get(ch, ch)),
           'baseline_imp': base_imp, 'baseline_exp': base_exp,
           'cagr_imp': cagr(ch, 'M'), 'cagr_exp': cagr(ch, 'X'),
           'imp_elast': -3.5, 'exp_elast': -2.5,
           'wedge': (v5[ch]['wedge'] if ch in v5 else NEW_WEDGE.get(ch, (0.0, ''))[0]),
           'tier': (v5[ch].get('tier', '') if ch in v5 else NEW_WEDGE.get(ch, (0.0, 'None'))[1]),
           'reduction_type': 'Full', 'target_eu': 0.0, 'sensitivity': v5[ch]['sensitivity'] if ch in v5 else 'n/a'}
    # ----- EU side -----
    ml = [l for l in LINES if l['hs2'] == ch and l['flow'] == 'M']
    tot = sum(l['value'] for l in ml)
    by6 = collections.defaultdict(lambda: {'party': 0.0, 'bol': 0.0})
    for l in ml:
        by6[l['hs6']]['bol' if l['partner'] == 'BOL' else 'party'] += l['value']
    # chapter fallback tau
    matched = [(W['EU'][h]['rate'] / 100, v['party'] + v['bol']) for h, v in by6.items() if h in W['EU']]
    tau_fb = sum(t * w for t, w in matched) / sum(w for _, w in matched) if matched and sum(w for _, w in matched) > 0 else 0.0
    eu_lines = []
    for h, v in by6.items():
        tau0 = W['EU'][h]['rate'] / 100 if h in W['EU'] else tau_fb
        cf_line = [0] * 11
        share_party = v['party'] / tot if tot > 0 else 0
        share_bol = v['bol'] / tot if tot > 0 else 0
        cf_line = [share_party * base_imp * (1 + rec['cagr_imp']) ** (2026 + y - 2023) for y in HORIZON]
        path, info = resolve_eu(h, tau0, cf_line)
        if share_party > 0:
            eu_lines.append({'hs6': h, 'share': share_party, 'tau0': tau0, 'r_path': path, 'info': info, 'partner': 'MER4'})
        if share_bol > 0:
            eu_lines.append({'hs6': h, 'share': share_bol, 'tau0': tau0, 'r_path': [1.0] * 11, 'info': {'matched': 'BOL not party'}, 'partner': 'BOL'})
    rec['eu_lines'] = eu_lines
    # ----- Mercosur side -----
    xl = [l for l in LINES if l['hs2'] == ch and l['flow'] == 'X']
    xtot = sum(l['value'] for l in xl)
    mer_lines = []
    fb_pairs = []
    for l in xl:
        if l['partner'] == 'BOL':
            continue
        b, p, info = me_line_path(l['hs6'], l['partner'])
        if b is None:
            wr = W.get(l['partner'], {}).get(l['hs6'])
            b = wr['rate'] / 100 if wr else None
        if b is not None:
            fb_pairs.append((b, l['value']))
        mer_lines.append({'hs6': l['hs6'], 'partner': l['partner'], 'share': l['value'] / xtot if xtot else 0, 'tau0': b,
                          'r_path': p if p is not None else [r_linear(10, y) for y in HORIZON], 'info': info})
    mfb = sum(b * w for b, w in fb_pairs) / sum(w for _, w in fb_pairs) if fb_pairs and sum(w for _, w in fb_pairs) > 0 else 0.0
    for m in mer_lines:
        if m['tau0'] is None:
            m['tau0'] = mfb
            m['info']['tau_fallback'] = 'chapter mean'
    bol_share = sum(l['value'] for l in xl if l['partner'] == 'BOL') / xtot if xtot else 0
    if bol_share > 0:
        mer_lines.append({'hs6': 'BOL', 'partner': 'BOL', 'share': bol_share, 'tau0': 0.0, 'r_path': [1.0] * 11, 'info': {'matched': 'BOL not party'}})
    rec['mer_lines'] = mer_lines
    # display fields
    rec['eu_mfn'] = sum(l['share'] * l['tau0'] for l in eu_lines)
    rec['mer_mfn'] = sum(l['share'] * l['tau0'] for l in mer_lines)
    ph = [(l['share'] * l['tau0'], next((y for y in HORIZON if l['r_path'][y] <= 1e-9), 10)) for l in eu_lines if l['r_path'][0] < 1 or l['r_path'][10] < 1]
    rec['phase_eu'] = max(1, round(sum(w * p for w, p in ph) / sum(w for w, _ in ph))) if ph and sum(w for w, _ in ph) > 0 else 7
    phm = [(l['share'] * l['tau0'], next((y for y in HORIZON if l['r_path'][y] <= 1e-9), 15)) for l in mer_lines if l['r_path'][10] < 1]
    rec['phase_mer'] = max(1, round(sum(w * p for w, p in phm) / sum(w for w, _ in phm))) if phm and sum(w for w, _ in phm) > 0 else 10
    chapters.append(rec)
    dut = sum(l['share'] for l in eu_lines if l['tau0'] > 0)
    y10 = sum(l['share'] * l['tau0'] * l['r_path'][10] for l in eu_lines)
    report.append((ch, rec['name'][:28], base_imp / 1e6, rec['eu_mfn'] * 100, y10 * 100, dut * 100, base_exp / 1e6, rec['mer_mfn'] * 100,
                   sum(l['share'] * l['tau0'] * l['r_path'][10] for l in mer_lines) * 100, rec['cagr_imp'] * 100, rec['cagr_exp'] * 100))

json.dump({'chapters': chapters, 'meta': {'greek_quota_share': GREEK_SHARE, 'wits_year': 2023,
           'schedule_source': 'OJ L 2026/184, Appendices 2-A-1 and 2-A-2'}}, open('v8_config.json', 'w'), indent=0)
print(f"{'HS':3s} {'name':28s} {'impM':>7s} {'EU t0':>6s} {'EU y10':>6s} {'dut%':>5s} | {'expM':>6s} {'Mer t0':>6s} {'Mer y10':>7s} | {'gM':>5s} {'gX':>5s}")
for r in report:
    print(f"{r[0]:3s} {r[1]:28s} {r[2]:7.2f} {r[3]:6.1f} {r[4]:6.1f} {r[5]:5.0f} | {r[6]:6.2f} {r[7]:6.1f} {r[8]:7.1f} | {r[9]:5.1f} {r[10]:5.1f}")
