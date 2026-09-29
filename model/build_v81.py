"""v8.1 builder: line-level tariffs from the EU-Mercosur Interim Agreement on Trade schedules
(OJ L 2026/184), revised after the three-way verification of v8.

Changes from v8 (see CHANGELOG):
 1. TRQ lines carry their quota data to the engine, which applies marginal-price logic:
    once the (pooled) notional Greek slice binds, additional imports pay the out-of-quota
    duty, so there is no quantity response; below the slice the response is capped at it.
 2. Quota tonnes in carcass-weight equivalent are converted to product weight
    (Annex 2-A Sec E: 130% boneless beef, 140% boneless poultry); slices are pooled across
    the HS-6 lines that share a quota; year-0 quota volumes are pro-rated to 245/365.
 3. For the lines that carry the result, tau0 and staging come from the traded CN-8 row,
    valued at Greek unit values (verified against TARIC; tariffcheck/tariff_check_lines.csv).
 4. Elsewhere, CN-8/NCM row paths are weighted by their ad valorem base rate; Mercosur
    base = min(schedule base, applied MFN) per Art 2.4(7).
 5. Spike screen: a line-partner flow whose single largest year exceeds 80% of its 2022-2024
    sum and USD 1 million is treated as a one-off or spike; that year is replaced by the mean
    of the other two years (a pure one-off therefore drops out). Flows that first appear in
    2024 are kept, since a new flow cannot be shown to be a one-off. Baselines and the 2024
    endpoint are adjusted accordingly. Growth follows the published endpoint rule
    (2014 to 2024, winsorized at +/-15%); three-year-average growth is a sensitivity.
 6. Year 0 runs from 1 May: r_eff(0) = 1 - (8/12)(1 - r(0)).
"""
import json, csv, os, re, glob, collections, copy
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
RAW = os.path.join(HERE, 'inputs', 'Greece__Mar_Total_TradeData_Balance_Sheet.xlsx')
V5 = os.path.join(HERE, 'inputs', 'v5_config.json')
GREEK_SHARE = 0.024
H = range(11)
PARTNERS = {32: 'ARG', 68: 'BOL', 76: 'BRA', 600: 'PRY', 858: 'URY'}

DEFAULTS = dict(growth_rule='endpoint', exclude_oneoffs=True, tobacco_tau0=0.0404, hilton_quota=False, drop_new_2024=False,
                fish_atq_zero=False, ethanol_chem_quota=False, growth_cap=0.15, static_growth=False,
                year0_prorate=True)

# Verified traded-row overrides for the lines that carry the result (TARIC, Greek unit values)
OVERRIDES = {
    '240120': ('tobacco', '4'), '080550': (0.0651, '7'), '271019': (0.0, '0'), '030617': (0.12, '4'),
    '030743': (0.08, '0'), '380610': (0.05, '4'), '220720': (0.161, 'EL'), '020230': (0.20, 'BF2'),
    '020130': (0.20, 'BF1'), '200919': (0.122, '7'), '030474': (0.075, '0'), '640220': (0.17, '10'),
    '170199': (0.65, 'E'), '441239': (0.07, '7'), '080830': (0.0186, '0'), '200969': (0.40, '10'),
    '200911': (0.152, '10'), '230990': (0.0273, '10'), '020714': (0.379, 'PY1'), '030366': (0.15, '0'),
    '020629': (0.0, '0'),   # traded row 0206 29 99 is Free at the Greek unit value
}
FISH_ATQ = {'030617', '030743', '030474', '030366'}

EU_TRQ = {  # code: (rule, parameter, [(year0 t, final t, final year)], cwe_factor)
    'BF1': ('ad_valorem', 0.075, [(9075, 54450, 5)], 1.3), 'BF2': ('ad_valorem', 0.075, [(7425, 44550, 5)], 1.3),
    'PK': ('fraction', 0.1, [(4167, 25000, 5)], 1.2), 'PY1': ('zero', 0, [(15000, 90000, 5)], 1.4),
    'PY2': ('zero', 0, [(15000, 90000, 5)], 1.0), 'MP': ('pref_step', 'dairy', [(1000, 10000, 10)], 1.0),
    'CE': ('pref_step', 'dairy', [(3000, 30000, 10)], 1.0), 'IF': ('pref_step', 'dairy', [(500, 5000, 10)], 1.0),
    'ME': ('zero', 0, [(166667, 1000000, 5)], 1.0), 'RE': ('zero', 0, [(10000, 60000, 5)], 1.0),
    'SR': ('zero', 0, [(180000, 180000, 0)], 1.0), 'OS': ('fraction', 0.5, [(2000, 2000, 0)], 1.0),
    'EG1': ('zero', 0, [(500, 3000, 5)], 1.0), 'EG2': ('zero', 0, [(500, 3000, 5)], 1.0),
    'HY': ('zero', 0, [(7500, 45000, 5)], 1.0), 'RM': ('zero', 0, [(400, 2400, 5)], 1.0),
    'SC': ('zero', 0, [(1000, 1000, 0)], 1.0), 'SH1': ('fraction', 0.5, [(1500, 1500, 0)], 1.0),
    'SH2': ('zero', 0, [(100, 600, 5)], 1.0), 'EL': ('fraction', 1 / 3, [(33333, 200000, 5)], 1.0),
    'ELCHEM': ('zero', 0, [(75000, 450000, 5)], 1.0),
    'HILTON': ('zero', 0, [(46876, 46876, 0)], 1.0),   # existing WTO Hilton quotas (AR 29,500; BR 10,000; UY 6,376; PY 1,000 t), in-quota 20% -> 0 at EIF
    'GC': ('pref_step', 'garlic', [(1875, 15000, 7)], 1.0),
}
DAIRY_PREF = [0.10, 0.20, 0.30, 0.40, 0.50, 0.60, 0.70, 0.80, 0.90, 0.95, 1.0]
DAIRY_VOL = [3000, 6000, 9000, 12000, 15000, 18000, 21000, 24000, 27000, 28500, 30000]


def r_linear(P, y):
    return 0.0 if P <= 0 else max(0.0, 1 - (y + 1) / (P + 1))


def r_category(cat, y):
    c = cat.replace(' ', '')
    if c in ('0', '0/EP', 'SW/12'):
        return 0.0
    m = re.fullmatch(r'(\d+)(/EP)?', c)
    if m:
        return r_linear(int(m.group(1)), y)
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
    if c == '15V':
        return 1.0 if y <= 6 else max(0.0, 1 - [0.19, 0.381, 0.571, 0.643, 0.714, 0.786, 0.857, 0.929, 1.0][min(y - 7, 8)])
    return None


def quota_t(code, y, prorate):
    q0, qf, fy = EU_TRQ[code][2][0]
    q = qf if fy == 0 else q0 + (qf - q0) * min(1.0, y / fy)
    if code in ('CE', 'MP', 'IF'):
        q = DAIRY_VOL[y] / 30000 * qf
    if y == 0 and prorate:
        q *= 245 / 365
    return q


def r_in_path(code, tau0):
    rule, par = EU_TRQ[code][0], EU_TRQ[code][1]
    out = []
    for y in H:
        if rule == 'zero':
            out.append(0.0)
        elif rule == 'fraction':
            out.append(par)
        elif rule == 'ad_valorem':
            out.append(min(1.0, par / tau0) if tau0 > 0 else 1.0)
        elif rule == 'pref_step':
            pref = DAIRY_PREF[y] if par == 'dairy' else 0.30 + 0.70 * min(1.0, y / 7)
            out.append(1 - pref)
        else:
            out.append(1.0)
    return out


def num_rate(s):
    s = s.replace(',', '.').strip()
    return float(s) / 100 if re.fullmatch(r'\d+(\.\d+)?', s) else None


def load_inputs():
    eu_rows = list(csv.DictReader(open(os.path.join(HERE, 'inputs', 'schedules/appendix_2-A-1_EU_schedule_from_OJ_xhtml.csv'), encoding='utf-8')))
    me_rows = list(csv.DictReader(open(os.path.join(HERE, 'inputs', 'schedules/appendix_2-A-2_MERCOSUR_schedule_from_OJ_xhtml.csv'), encoding='utf-8')))
    for r in eu_rows:
        r['code'] = re.sub(r'\D', '', r['CN 2013'])
    for r in me_rows:
        r['code'] = re.sub(r'\D', '', r['NCM'])
    W = json.load(open(os.path.join(HERE, 'inputs', 'wits_rates.json')))
    # HS-6 by year
    yearly = collections.defaultdict(lambda: collections.defaultdict(float))
    uvn, uvd = collections.defaultdict(float), collections.defaultdict(float)
    for f in glob.glob(os.path.join(HERE, 'inputs', 'comtrade_hs6', 'ct6_*_*_*.json')):
        _, flow, year, p = os.path.basename(f)[:-5].split('_')
        for r in json.load(open(f)).get('data', []):
            yearly[(flow, r['cmdCode'], PARTNERS[int(p)])][int(year)] += (r.get('primaryValue') or 0)
            if flow == 'M' and r.get('netWgt'):
                uvn[r['cmdCode']] += r['primaryValue'] or 0
                uvd[r['cmdCode']] += r['netWgt']
    UV = {k: uvn[k] / uvd[k] for k in uvn if uvd[k] > 0}
    raw = pd.read_excel(RAW, 'Sheet1')
    raw['cmd'] = raw['cmdCode'].astype(str).str.zfill(2)
    g = raw.groupby(['cmd', 'flowCode', 'refYear'])['primaryValue'].sum()
    return eu_rows, me_rows, W, yearly, UV, g


INPUTS = None


def build(opts=None):
    global INPUTS
    o = dict(DEFAULTS)
    o.update(opts or {})
    if INPUTS is None:
        INPUTS = load_inputs()
    eu_rows, me_rows, W, yearly, UV, g = INPUTS
    D = json.load(open(os.path.join(HERE, 'inputs', 'lines44.json')))
    CH = D['chapters']
    v5 = {c['hs']: c for c in json.load(open(V5))['chapters']}

    def rawv(ch, f, y):
        try:
            return float(g.loc[(ch, f, y)])
        except KeyError:
            return 0.0

    # one-off detection
    oneoffs = {}   # (flow, h6, p) -> {year: amount removed}
    for (flow, h6, p), ys in yearly.items():
        vals = {y: ys.get(y, 0.0) for y in (2022, 2023, 2024)}
        tot3 = sum(vals.values())
        ymax = max(vals, key=vals.get)
        new_flow_2024 = ymax == 2024 and vals[2022] == 0 and vals[2023] == 0
        if tot3 > 0 and vals[ymax] > 0.8 * tot3 and vals[ymax] > 1e6 and not (new_flow_2024 and not o['drop_new_2024']):
            others = [v for y, v in vals.items() if y != ymax]
            oneoffs[(flow, h6, p)] = {ymax: vals[ymax] - sum(others) / 2}
    excl = oneoffs if o['exclude_oneoffs'] else {}

    def line_base(flow, h6, p):
        ys = yearly.get((flow, h6, p), {})
        adj = excl.get((flow, h6, p), {})
        return sum(ys.get(y, 0.0) - adj.get(y, 0.0) for y in (2022, 2023, 2024)) / 3

    def removed(ch, flow, year=None):
        tot = 0.0
        for (f, h6, p), adj in excl.items():
            if f == flow and h6[:2] == ch:
                tot += sum(adj.values()) if year is None else adj.get(year, 0.0)
        return tot

    def chap_totals(ch, flow):
        base = sum(rawv(ch, flow, y) for y in (2022, 2023, 2024)) / 3
        drop = removed(ch, flow) / 3
        return base - drop, drop

    def growth(ch, flow, base_now):
        if o['static_growth']:
            return 0.0
        cap = o['growth_cap']
        if o['growth_rule'] == 'endpoint':
            a, b, n = rawv(ch, flow, 2014), rawv(ch, flow, 2024), 10
            b -= removed(ch, flow, 2024)
        else:
            a, b, n = sum(rawv(ch, flow, y) for y in (2014, 2015, 2016)) / 3, base_now, 8
        if a > 0 and b > 0:
            c = (b / a) ** (1 / n) - 1
        elif a == 0 and b > 0:
            c = cap
        elif a > 0 and b <= 0:
            c = -cap
        else:
            c = 0.0
        return max(-cap, min(cap, c))

    def prorate0(path):
        if not o['year0_prorate']:
            return path
        p = list(path)
        p[0] = 1 - (8 / 12) * (1 - p[0])
        return p

    NAMES = {'02': 'Meat and edible offal', '03': 'Fish and crustaceans', '07': 'Edible vegetables', '10': 'Cereals',
             '16': 'Meat and fish preparations', '17': 'Sugars and sugar confectionery', '21': 'Miscellaneous edible preparations',
             '28': 'Inorganic chemicals', '29': 'Organic chemicals', '32': 'Tanning and dyeing extracts, paints',
             '33': 'Essential oils, perfumery, cosmetics', '40': 'Rubber and articles', '44': 'Wood and articles',
             '48': 'Paper and paperboard', '49': 'Printed books and products', '64': 'Footwear', '74': 'Copper and articles',
             '90': 'Optical, medical instruments', '94': 'Furniture, lighting'}
    NEW_WEDGE = {'02': (0.01, 'EUDR cattle'), '40': (0.01, 'EUDR rubber'), '44': (0.01, 'EUDR wood'),
                 '48': (0.01, 'EUDR paper'), '49': (0.01, 'EUDR printed products'), '07': (0.01, 'Import-standards compliance, perishables')}

    chapters, trq_lines = [], collections.defaultdict(list)
    for ch in CH:
        base_imp, drop_imp = chap_totals(ch, 'M')
        base_exp, drop_exp = chap_totals(ch, 'X')
        rec = {'hs': ch, 'name': v5[ch]['name'] if ch in v5 else NAMES.get(ch, ch), 'baseline_imp': max(base_imp, 0.0),
               'baseline_exp': max(base_exp, 0.0), 'oneoff_excluded_imp': drop_imp, 'oneoff_excluded_exp': drop_exp,
               'cagr_imp': growth(ch, 'M', base_imp), 'cagr_exp': growth(ch, 'X', base_exp), 'imp_elast': -3.5, 'exp_elast': -2.5,
               'wedge': v5[ch]['wedge'] if ch in v5 else NEW_WEDGE.get(ch, (0.0, ''))[0],
               'tier': v5[ch].get('tier', '') if ch in v5 else NEW_WEDGE.get(ch, (0.0, 'None'))[1],
               'reduction_type': 'Full', 'target_eu': 0.0, 'sensitivity': v5[ch]['sensitivity'] if ch in v5 else 'n/a'}
        # ---------- EU side ----------
        lines_m = collections.defaultdict(lambda: {'party': 0.0, 'bol': 0.0})
        for (flow, h6, p) in yearly:
            if flow == 'M' and h6[:2] == ch:
                lines_m[h6]['bol' if p == 'BOL' else 'party'] += line_base(flow, h6, p)
        tot = sum(v['party'] + v['bol'] for v in lines_m.values())
        wits_fb = [(W['EU'][h]['rate'] / 100, v['party'] + v['bol']) for h, v in lines_m.items() if h in W['EU']]
        tau_fb = sum(t * w for t, w in wits_fb) / sum(w for _, w in wits_fb) if wits_fb and sum(w for _, w in wits_fb) > 0 else 0.0
        eu_lines = []
        for h6, v in lines_m.items():
            if tot <= 0 or (v['party'] + v['bol']) <= 0:
                continue
            share_p, share_b = v['party'] / tot, v['bol'] / tot
            trq_code = None
            if h6 in OVERRIDES:
                t0, cat = OVERRIDES[h6]
                if t0 == 'tobacco':
                    t0 = o['tobacco_tau0']

                if o['fish_atq_zero'] and h6 in FISH_ATQ:
                    t0 = 0.0
                if cat in EU_TRQ:
                    trq_code, path = cat, [1.0] * 11
                else:
                    path = [r_category(cat, y) for y in H]
                info = {'source': 'traded CN-8 override', 'cat': cat}
            else:
                t0 = W['EU'][h6]['rate'] / 100 if h6 in W['EU'] else tau_fb
                rows = [r for r in eu_rows if r['code'].startswith(h6)] or [r for r in eu_rows if r['code'].startswith(h6[:4])]
                info = {'source': 'schedule rows (base-weighted)', 'cats': [r['Staging category'] for r in rows][:8]}
                if not rows:
                    path = [1.0] * 11
                else:
                    num, den = [0.0] * 11, 0.0
                    trqs = [r['Staging category'].replace(' ', '') for r in rows if r['Staging category'].replace(' ', '') in EU_TRQ]
                    lin = [r for r in rows if r['Staging category'].replace(' ', '') not in EU_TRQ]
                    if trqs and len(trqs) >= len(lin):
                        trq_code = collections.Counter(trqs).most_common(1)[0][0]
                        path = [1.0] * 11
                    else:
                        for r in lin:
                            p = [r_category(r['Staging category'], y) for y in H]
                            if p[0] is None:
                                continue
                            br = r['Base rate'].strip().lower()
                            w = 0.0 if br.startswith('free') else (num_rate(r['Base rate']) or 0.05)
                            num = [a + w * b for a, b in zip(num, p)]
                            den += w
                        path = [a / den for a in num] if den > 0 else ([0.0] * 11 if all(r['Base rate'].strip().lower().startswith('free') for r in lin) else [1.0] * 11)
            if share_p > 0:
                ln = {'hs6': h6, 'share': share_p, 'tau0': t0, 'r_path': prorate0(path), 'partner': 'MER4', 'info': info}
                if trq_code:
                    ln['trq'] = {'code': trq_code, 'r_in': r_in_path(trq_code, t0)}
                    trq_lines[trq_code].append((ch, h6, v['party']))
                    if trq_code in ('BF1', 'BF2') and o['hilton_quota']:
                        ln['trq']['code2'] = 'HILTON'
                        ln['trq']['r_in2'] = r_in_path('HILTON', t0)
                        trq_lines['HILTON'].append((ch, h6, v['party']))
                    if trq_code == 'EL' and o['ethanol_chem_quota']:
                        ln['trq']['code2'] = 'ELCHEM'
                        ln['trq']['r_in2'] = r_in_path('ELCHEM', t0)
                        trq_lines['ELCHEM'].append((ch, h6, v['party']))
                eu_lines.append(ln)
            if share_b > 0:
                eu_lines.append({'hs6': h6, 'share': share_b, 'tau0': t0, 'r_path': [1.0] * 11, 'partner': 'BOL', 'info': {'source': 'BOL not party'}})
        rec['eu_lines'] = eu_lines
        # ---------- Mercosur side ----------
        col = {'ARG': 'Base rate of Argentina', 'BRA': 'Base rate of Brazil', 'PRY': 'Base rate of Paraguay', 'URY': 'Base rate of Uruguay'}
        xl = [(h6, p, line_base('X', h6, p)) for (flow, h6, p) in yearly if flow == 'X' and h6[:2] == ch]
        xtot = sum(v for _, _, v in xl)
        mer_lines, fb = [], []
        for h6, p, val in xl:
            if val <= 0 or xtot <= 0:
                continue
            if p == 'BOL':
                mer_lines.append({'hs6': h6, 'partner': 'BOL', 'share': val / xtot, 'tau0': 0.0, 'r_path': [1.0] * 11, 'info': {'source': 'BOL not party'}})
                continue
            rows = [r for r in me_rows if r['code'].startswith(h6)] or [r for r in me_rows if r['code'].startswith(h6[:4])]
            bases, num, den = [], [0.0] * 11, 0.0
            for r in rows:
                b = num_rate(r[col[p]])
                c = r['Staging category'].replace(' ', '')
                if c.startswith('TRQ-'):
                    pth = [1 - DAIRY_PREF[y] for y in H] if c != 'TRQ-4' else [1 - (0.30 + 0.70 * min(1.0, y / 7)) for y in H]
                elif c == 'CH1' or c == 'T1':
                    pth = [r_linear(9, y) for y in H]
                elif c == 'CH2':
                    pth = [r_linear(14, y) for y in H]
                else:
                    pth = [r_category(r['Staging category'], y) for y in H]
                    if pth[0] is None:
                        pth = [r_linear(18, y) for y in H]
                if b is not None:
                    bases.append(b)
                w = b if (b is not None and b > 0) else 0.0
                num = [a + w * q for a, q in zip(num, pth)]
                den += w
            base = sum(bases) / len(bases) if bases else None
            path = [a / den for a in num] if den > 0 else ([r_linear(10, y) for y in H] if rows else [r_linear(10, y) for y in H])
            applied = W.get(p, {}).get(h6, {}).get('rate')
            applied = applied / 100 if applied is not None else None
            if base is None:
                t0 = applied
            else:
                t0 = min(base, applied) if applied is not None else base
            if t0 is not None:
                fb.append((t0, val))
            mer_lines.append({'hs6': h6, 'partner': p, 'share': val / xtot, 'tau0': t0, 'r_path': prorate0(path),
                              'info': {'source': 'schedule (base-weighted)' if rows else 'no schedule row'}})
        mfb = sum(t * w for t, w in fb) / sum(w for _, w in fb) if fb and sum(w for _, w in fb) > 0 else 0.0
        for m in mer_lines:
            if m['tau0'] is None:
                m['tau0'] = mfb
        rec['mer_lines'] = mer_lines
        rec['eu_mfn'] = sum(l['share'] * l['tau0'] for l in eu_lines)
        rec['mer_mfn'] = sum(l['share'] * l['tau0'] for l in mer_lines)
        rec['phase_eu'] = 7
        rec['phase_mer'] = 10
        chapters.append(rec)

    # pooled TRQ slices: allocate each quota's notional Greek slice across its lines by baseline value
    for code, members in trq_lines.items():
        totv = sum(v for _, _, v in members)
        for ch, h6, v in members:
            c = next(c for c in chapters if c['hs'] == ch)
            for ln in c['eu_lines']:
                if ln['hs6'] == h6 and 'trq' in ln:
                    uv = UV.get(h6, 3.0)  # USD per kg, product weight
                    cwe = EU_TRQ[code][3]
                    sl = [GREEK_SHARE * quota_t(code, y, o['year0_prorate']) * 1000 / cwe * uv * (v / totv if totv else 1) for y in H]
                    if code in ('ELCHEM', 'HILTON'):
                        ln['trq']['slice2'] = sl
                    else:
                        ln['trq']['slice'] = sl
    meta = {'opts': o, 'oneoffs': {f"{k[0]}|{k[1]}|{k[2]}": {str(y): a for y, a in v.items()} for k, v in oneoffs.items()},
            'excluded_oneoffs': sorted(f"{k[0]}|{k[1]}|{k[2]}" for k in excl)}
    return {'chapters': chapters, 'meta': meta}


if __name__ == '__main__':
    cfg = build()
    json.dump(cfg, open(os.path.join(HERE, 'v81_config.json'), 'w'), indent=0)
    print('one-offs detected:', cfg['meta']['oneoffs'])
    for c in cfg['chapters']:
        if c['baseline_imp'] > 1e6 or c['baseline_exp'] > 1e6:
            print(f"{c['hs']} imp {c['baseline_imp']/1e6:7.2f} g {c['cagr_imp']*100:5.1f} | exp {c['baseline_exp']/1e6:6.2f} g {c['cagr_exp']*100:5.1f} | EU t0 {c['eu_mfn']*100:5.1f} Mer t0 {c['mer_mfn']*100:5.1f}")
