"""Analysis C: build line-level model inputs from INDEPENDENT sources.

Trade:    Eurostat COMEXT (ds-045409), Greece reporter, CN8, 2014-2024, partners BR AR UY PY (+BO, no preference),
          EUR converted to USD at ECB annual average rates.
Tariffs:  EU TARIC third-country duty at CN8 (SimDate 2023-07-01) for every import line with a material baseline, ad
          valorem equivalents at Greek unit values (COMEXT value / quantity); WITS-TRAINS 2023 AVE at HS6 elsewhere.
          Mercosur: schedule base rates (Appendix 2-A-2, per party) capped at the WITS applied MFN (Art. 2.4(7)).
Staging:  Appendix 2-A-1 / 2-A-2 of OJ L 2026/184 (CSV parsed from the OJ XHTML), exact CN8 match first, HS6 fallback.
Method:   the published rules (chapter inclusion > USD 1m in any year; baselines 2022-24 mean; endpoint CAGR 2014-24
          capped +/-15%; spike screen; Year 0 pro-rated 8/12; notional Greek TRQ slice 2.4%).
Output:   config_C.json (same schema as the v8.1 config, readable by engine_C.py and by pe_v81.py) + lines_C.csv diagnostics.
"""
import json, csv, re, os, sys, io, collections, math
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import taric, wits

HERE = os.path.dirname(os.path.abspath(__file__))
R = os.environ.get('EUMG_PACKAGE', os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
SCHED_EU = os.path.join(R, 'model/inputs/schedules/appendix_2-A-1_EU_schedule_from_OJ_xhtml.csv')
SCHED_ME = os.path.join(R, 'model/inputs/schedules/appendix_2-A-2_MERCOSUR_schedule_from_OJ_xhtml.csv')
V5 = os.path.join(R, 'model/inputs/v5_config.json')
FX = {int(k): v for k, v in json.load(open('ecb_usd_per_eur.json')).items()}
PARTY = {'BR': 'BRA', 'AR': 'ARG', 'UY': 'URY', 'PY': 'PRY', 'BO': 'BOL'}
H = range(11)
GREEK_SHARE = 0.024
OPTS = dict(spike_screen=True, growth_cap=0.15, year0_prorate=True, taric_min_usd=100_000, tobacco='certified',
            fish_atq_zero=False, include_bolivia_in_totals=True)

# ---------------- quota parameters (Annex 2-A Sections B and E of the agreement; see trq_summary.csv) ----------------
# code: (in-quota rule, parameter, (year0 t, final t, final year), carcass-weight factor to product weight)
EU_TRQ = {'BF1': ('ad_valorem', 0.075, (9075, 54450, 5), 1.3), 'BF2': ('ad_valorem', 0.075, (7425, 44550, 5), 1.3),
          'PK': ('fraction', 0.1, (4167, 25000, 5), 1.2), 'PY1': ('zero', 0, (15000, 90000, 5), 1.4), 'PY2': ('zero', 0, (15000, 90000, 5), 1.0),
          'MP': ('pref_step', 'dairy', (1000, 10000, 10), 1.0), 'CE': ('pref_step', 'dairy', (3000, 30000, 10), 1.0),
          'IF': ('pref_step', 'dairy', (500, 5000, 10), 1.0), 'ME': ('zero', 0, (166667, 1000000, 5), 1.0), 'RE': ('zero', 0, (10000, 60000, 5), 1.0),
          'SR': ('zero', 0, (180000, 180000, 0), 1.0), 'OS': ('fraction', 0.5, (2000, 2000, 0), 1.0), 'EG1': ('zero', 0, (500, 3000, 5), 1.0),
          'EG2': ('zero', 0, (500, 3000, 5), 1.0), 'HY': ('zero', 0, (7500, 45000, 5), 1.0), 'RM': ('zero', 0, (400, 2400, 5), 1.0),
          'SC': ('zero', 0, (1000, 1000, 0), 1.0), 'SH1': ('fraction', 0.5, (1500, 1500, 0), 1.0), 'SH2': ('zero', 0, (100, 600, 5), 1.0),
          'EL': ('fraction', 1 / 3, (33333, 200000, 5), 1.0), 'GC': ('pref_step', 'garlic', (1875, 15000, 7), 1.0)}
DAIRY_PREF = [0.10, 0.20, 0.30, 0.40, 0.50, 0.60, 0.70, 0.80, 0.90, 0.95, 1.0]


def r_linear(P, y):
    return 0.0 if P <= 0 else max(0.0, 1 - (y + 1) / (P + 1))


def r_category(cat, y):
    c = (cat or '').replace(' ', '')
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


def quota_t(code, y):
    q0, qf, fy = EU_TRQ[code][2]
    q = qf if fy == 0 else q0 + (qf - q0) * min(1.0, y / fy)
    if code in ('CE', 'MP', 'IF'):
        q = [3000, 6000, 9000, 12000, 15000, 18000, 21000, 24000, 27000, 28500, 30000][y] / 30000 * qf
    if y == 0 and OPTS['year0_prorate']:
        q *= 8 / 12      # the Commission opened the 2026 quotas at 8/12 of the year-0 quantity (1 May - 31 Dec)
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
    return out


def prorate0(path):
    p = list(path)
    if OPTS['year0_prorate']:
        p[0] = 1 - (8 / 12) * (1 - p[0])
    return p


def num_rate(s):
    s = (s or '').replace(',', '.').strip()
    return float(s) / 100 if re.fullmatch(r'\d+(\.\d+)?', s) else None


# ---------------- 1. trade data ----------------
rows = list(csv.DictReader(open('comext_cn8_tidy.csv', encoding='utf-8')))
for r in rows:
    r['year'] = int(r['year']); r['v_eur'] = float(r['value_eur']); r['q'] = float(r['qty_100kg']); r['v_usd'] = r['v_eur'] * FX[r['year']]
print('COMEXT rows', len(rows), 'years', sorted({r['year'] for r in rows}), 'partners', sorted({r['partner'] for r in rows}))

def parties(r):
    return OPTS['include_bolivia_in_totals'] or r['partner'] != 'BO'

chap_year = collections.defaultdict(float)      # (hs2, flow, year) -> USD, all partners in scope
line_year = collections.defaultdict(float)      # (flow, cn8, partner, year) -> USD
uv_num, uv_den = collections.defaultdict(float), collections.defaultdict(float)   # CN8 -> EUR, 100kg (imports 2022-24)
for r in rows:
    if not parties(r):
        continue
    chap_year[(r['cn8'][:2], r['flow'], r['year'])] += r['v_usd']
    line_year[(r['flow'], r['cn8'], r['partner'], r['year'])] += r['v_usd']
    if r['flow'] == 'M' and 2022 <= r['year'] <= 2024:
        uv_num[r['cn8']] += r['v_eur']; uv_den[r['cn8']] += r['q']
UV100 = {c: uv_num[c] / uv_den[c] for c in uv_num if uv_den[c] > 0}     # EUR per 100 kg
UVUSDKG = {c: uv_num[c] * FX[2023] / (uv_den[c] * 100) for c in uv_num if uv_den[c] > 0}   # USD per kg (approx.)

chapters_all = sorted({k[0] for k in chap_year})
included = sorted({hs for hs in chapters_all if any(chap_year[(hs, f, y)] > 1e6 for f in 'MX' for y in range(2014, 2025))})
if '04' not in included:
    included.append('04'); included.sort()      # dairy kept for the quota/designation tests, as in the published set
print('chapters meeting the USD 1m rule:', len(included), included)

# spike screen at (flow, HS6, partner) level, as documented
h6_year = collections.defaultdict(lambda: collections.defaultdict(float))
for (f, c, p, y), v in line_year.items():
    if 2022 <= y <= 2024:
        h6_year[(f, c[:6], p)][y] += v
spikes = {}
for k, ys in h6_year.items():
    vals = {y: ys.get(y, 0.0) for y in (2022, 2023, 2024)}
    tot = sum(vals.values()); ymax = max(vals, key=vals.get)
    new24 = ymax == 2024 and vals[2022] == 0 and vals[2023] == 0
    if tot > 0 and vals[ymax] > 0.8 * tot and vals[ymax] > 1e6 and not new24:
        others = [v for y, v in vals.items() if y != ymax]
        spikes[k] = (ymax, vals[ymax] - sum(others) / 2)      # amount removed from that year
if not OPTS['spike_screen']:
    spikes = {}
print('spike-screened flows:', len(spikes))
for k, (y, amt) in sorted(spikes.items(), key=lambda x: -x[1][1])[:20]:
    print(f'   {k[0]} {k[1]} {k[2]} {y}: -{amt/1e6:.1f}m')

def line_base(f, c, p):
    """2022-24 mean at CN8-partner level; spike adjustment allocated pro rata within the HS6-partner flow."""
    tot = 0.0
    for y in (2022, 2023, 2024):
        v = line_year.get((f, c, p, y), 0.0)
        sp = spikes.get((f, c[:6], p))
        if sp and sp[0] == y:
            h6v = h6_year[(f, c[:6], p)][y]
            v -= sp[1] * (v / h6v if h6v else 0)
        tot += v
    return tot / 3

def removed(hs, f, year=None):
    t = 0.0
    for (ff, h6, p), (y, amt) in spikes.items():
        if ff == f and h6[:2] == hs and (year is None or y == year):
            t += amt
    return t

def growth(hs, f):
    cap = OPTS['growth_cap']
    a = chap_year.get((hs, f, 2014), 0.0); b = chap_year.get((hs, f, 2024), 0.0) - removed(hs, f, 2024)
    if a > 0 and b > 0:
        g = (b / a) ** (1 / 10) - 1
    elif a == 0 and b > 0:
        g = cap
    elif a > 0 and b <= 0:
        g = -cap
    else:
        g = 0.0
    return max(-cap, min(cap, g))

# ---------------- 2. schedules ----------------
eu_rows = list(csv.DictReader(open(SCHED_EU, encoding='utf-8')))
me_rows = list(csv.DictReader(open(SCHED_ME, encoding='utf-8')))
for r in eu_rows:
    r['code'] = re.sub(r'\D', '', r['CN 2013'])
for r in me_rows:
    r['code'] = re.sub(r'\D', '', r['NCM'])
eu_by_code = collections.defaultdict(list)
for r in eu_rows:
    eu_by_code[r['code']].append(r)
eu_by_h6 = collections.defaultdict(list)
for r in eu_rows:
    eu_by_h6[r['code'][:6]].append(r)
me_by_h6 = collections.defaultdict(list)
for r in me_rows:
    me_by_h6[r['code'][:6]].append(r)

eu_by_h4 = collections.defaultdict(list)
for r in eu_rows:
    eu_by_h4[r['code'][:4]].append(r)
CAT_RATIOS = {'0': 0.0, '4': 0.8, '7': 0.875, '8': 8 / 9, '10': 10 / 11, '15': 15 / 16}
MAIN_PARTNER = {}   # filled in section 3 (largest Mercosur-party supplier per CN8)


def taric_implied_category(c, partner_iso, uv):
    """Infer the staging category from the Mercosur preference TARIC applies today (Year 0 of the agreement, in force
    since 1 May 2026): ratio of preferential to erga-omnes AVE at the Greek unit value -> 0 / 0.8 / 0.875 / 0.889 / 0.909 / 1."""
    m = taric.measures(c, partner_iso, '20260930')
    if not m or not m['goods']:
        return None, 'TARIC 2026 lookup failed'
    base_aves, pref_aves, has_trq = [], [], False
    for g in m['goods']:
        expr = taric.duty_at_unit_value(g, uv, c)
        if expr is None:
            continue
        b = taric.ave(expr, uv, taric.uv_hl(c, uv))
        if b is None:
            continue
        prefs = [e for e in g['measures'] if (e.get('area') or '').lower().startswith('mercosur')]
        if any(e['type'] == 'Preferential tariff quota' for e in prefs):
            has_trq = True
        tp = [e for e in prefs if e['type'] == 'Tariff preference']
        base_aves.append(b)
        if tp:
            e = tp[0]
            d = e['duty'] or taric.duty_at_unit_value({'third_country_duty': None, 'third_country_conditions': e.get('conditions')}, uv, c)
            a = taric.ave(d, uv, taric.uv_hl(c, uv)) if d else None
            pref_aves.append(a if a is not None else b)
        else:
            pref_aves.append(b)
    if not base_aves:
        return None, 'no computable duty'
    if has_trq:
        return None, 'TRQ measure present'
    b = sum(base_aves) / len(base_aves); p = sum(pref_aves) / len(pref_aves)
    if b == 0:
        return '0', 'TARIC 2026: base duty free'
    ratio = p / b
    cat, target = min(CAT_RATIOS.items(), key=lambda kv: abs(kv[1] - ratio))
    if abs(target - ratio) < 0.02:
        return cat, f'TARIC-implied from 2026 Mercosur preference (pref/base = {ratio:.3f})'
    if ratio > 0.98:
        return 'E', f'TARIC-implied: no Year-0 cut (pref/base = {ratio:.3f})'
    return None, f'TARIC ratio {ratio:.3f} matches no category'


def eu_path_for(cn8):
    """Staging path r(y) for a CN8 line: exact CN 2013 row if it exists; otherwise the category implied by the Mercosur
    preference TARIC applies today (handles HS 2017/2022 code changes such as 0307 43 92 = CN 2013 0307 99 11); otherwise
    base-weighted average of HS6, then HS4, rows. Returns (path, trq_code or None, source, cats)."""
    rows_ = eu_by_code.get(cn8) or []
    src = 'CN8 exact'
    if not rows_:
        cat, note = taric_implied_category(cn8, MAIN_PARTNER.get(cn8, 'BR'), UV100.get(cn8))
        if cat is not None:
            return [r_category(cat, y) for y in H], None, note, [cat]
        rows_ = eu_by_h6.get(cn8[:6]) or []
        src = 'HS6 rows (' + note + ')'
        if not rows_:
            rows_ = eu_by_h4.get(cn8[:4]) or []
            src = 'HS4 rows (' + note + ')'
    if not rows_:
        return [1.0] * 11, None, 'no row', []
    cats = [r['Staging category'].replace(' ', '') for r in rows_]
    trqs = [c for c in cats if c in EU_TRQ]
    lin = [r for r in rows_ if r['Staging category'].replace(' ', '') not in EU_TRQ]
    if trqs and len(trqs) >= len(lin):
        return [1.0] * 11, collections.Counter(trqs).most_common(1)[0][0], src, cats
    num, den = [0.0] * 11, 0.0
    for r in lin:
        p = [r_category(r['Staging category'], y) for y in H]
        if p[0] is None:
            continue
        br = r['Base rate'].strip().lower()
        w = 0.0 if br.startswith('free') else (num_rate(r['Base rate']) or 0.05)
        num = [a + w * b for a, b in zip(num, p)]; den += w
    if den > 0:
        return [a / den for a in num], None, src, cats
    if all(r['Base rate'].strip().lower().startswith('free') for r in lin):
        return [0.0] * 11, None, src, cats
    return [1.0] * 11, None, src, cats

# ---------------- 3. EU tariffs: TARIC at CN8 for material lines, WITS HS6 AVE elsewhere ----------------
imp_lines = collections.defaultdict(float)   # (cn8, partner) -> baseline USD
for (f, c, p, y), v in line_year.items():
    if f == 'M' and 2022 <= y <= 2024 and c[:2] in included:
        imp_lines[(c, p)] += 0  # ensure key
for (c, p) in list(imp_lines):
    imp_lines[(c, p)] = line_base('M', c, p)
cn8_base = collections.defaultdict(float)
_pp = collections.defaultdict(dict)
for (c, p), v in imp_lines.items():
    if p != 'BO':
        cn8_base[c] += v; _pp[c][p] = v
for c, d in _pp.items():
    MAIN_PARTNER[c] = max(d, key=d.get)
material = sorted([c for c, v in cn8_base.items() if v >= OPTS['taric_min_usd']], key=lambda c: -cn8_base[c])
print(f'import CN8 lines: {len(cn8_base)}; with baseline >= USD {OPTS["taric_min_usd"]:,}: {len(material)}  (TARIC lookups)')
h6_all = sorted({c[:6] for c in cn8_base})
W_EU = wits.fetch('918', h6_all, 2023, 'aveestimated')
print('WITS EU AVE fetched for', sum(1 for v in W_EU.values() if v is not None), 'of', len(h6_all), 'HS6 lines')

taric_tau, taric_note = {}, {}
for i, c in enumerate(material):
    m = taric.measures(c, '', '20230701')
    if m is None or not m['goods']:
        taric_note[c] = 'TARIC lookup failed'; continue
    uv = UV100.get(c)
    aves, exprs = [], []
    for g in m['goods']:
        expr = taric.duty_at_unit_value(g, uv, c)
        if expr is None:
            continue
        a = taric.ave(expr, uv_eur_per_100kg=uv, uv_eur_per_hl=(taric.uv_hl(c, uv)))
        exprs.append((g['goods_code'], expr, a))
        if a is not None:
            aves.append((g['goods_code'], a))
    if not aves:
        taric_note[c] = 'not computable: ' + '; '.join(e[1] for e in exprs)[:120]; continue
    # if the CN8 splits into 10-digit lines with different duties, keep them all; choose later per line rule
    taric_tau[c] = {'sub': aves, 'exprs': exprs}
    if i % 25 == 0:
        print(f'   TARIC {i+1}/{len(material)} done')
json.dump({'taric_tau': taric_tau, 'taric_note': taric_note}, open('taric_lines_C.json', 'w'), indent=1)

# Documented legal specifics that the inline TARIC duty does not convey
LEGAL_OVERRIDES = {
    '27101951': (0.0, 'CN 2710 19 51 is an end-use subheading (UCC Art. 254; CN footnote): goods can only be declared there under the '
                      'end-use procedure at the autonomous rate Free, so the duty collected is 0%, not the conventional 3.5%'),
}


def eu_tau0(c):
    """tau0 for a CN8 import line, with provenance."""
    if c in LEGAL_OVERRIDES:
        return LEGAL_OVERRIDES[c][0], 'legal override: ' + LEGAL_OVERRIDES[c][1][:60]
    if c in taric_tau:
        subs = taric_tau[c]['sub']
        if c.startswith('2401') and len(subs) > 1:
            # leaf tobacco: the certificate-of-authenticity sub-lines (Virginia / Burley / Maryland) carry 18.4% MIN 22 MAX 24,
            # the residual sub-line 11.2% MIN 22 MAX 56; at Greek unit values the MAX binds, so certified = lowest AVE
            lo = min(subs, key=lambda x: x[1]); hi = max(subs, key=lambda x: x[1])
            if OPTS['tobacco'] == 'certified':
                return lo[1], f'TARIC CN8 (certified leaf sub-line {lo[0][-2:]})'
            if OPTS['tobacco'] == 'uncertified':
                return hi[1], f'TARIC CN8 (uncertified sub-line {hi[0][-2:]})'
        if len(subs) == 1:
            return subs[0][1], 'TARIC CN8'
        return sum(a for _, a in subs) / len(subs), 'TARIC CN8 (mean of 10-digit sub-lines)'
    # TARIC duty not shown inline (entry-price or other condition-dependent duty): use the agreement schedule's own
    # base rate for the exact CN row when it is a plain ad valorem number (e.g. lemons '6,4' with EP)
    if c in eu_by_code:
        brs = [num_rate(r['Base rate']) for r in eu_by_code[c]]
        brs = [b for b in brs if b is not None]
        if brs:
            return sum(brs) / len(brs), 'OJ schedule base rate (CN8 exact)'
        if all(r['Base rate'].strip().lower().startswith('free') for r in eu_by_code[c]):
            return 0.0, 'OJ schedule base rate Free (CN8 exact)'
    w = W_EU.get(c[:6])
    if w is not None:
        return w['rate'] / 100, 'WITS HS6 AVE 2023'
    return None, 'no source'

# ---------------- 4. Mercosur tariffs ----------------
exp_lines = collections.defaultdict(float)
for (f, c, p, y), v in line_year.items():
    if f == 'X' and 2022 <= y <= 2024 and c[:2] in included:
        exp_lines[(c, p)] = 0
for (c, p) in list(exp_lines):
    exp_lines[(c, p)] = line_base('X', c, p)
W_ME = {}
for iso, num in [('BR', '076'), ('AR', '032'), ('UY', '858'), ('PY', '600')]:
    h6s = sorted({c[:6] for (c, p), v in exp_lines.items() if p == iso and v > 0})
    W_ME[iso] = wits.fetch(num, h6s, 2023, 'reported')
    print('WITS', iso, 'MFN fetched for', sum(1 for v in W_ME[iso].values() if v is not None), 'of', len(h6s))
COL = {'AR': 'Base rate of Argentina', 'BR': 'Base rate of Brazil', 'PY': 'Base rate of Paraguay', 'UY': 'Base rate of Uruguay'}

me_by_h4 = collections.defaultdict(list)
for r in me_rows:
    me_by_h4[r['code'][:4]].append(r)


def mer_line(c, p):
    """Appendix 2-A-2 is in NCM 2012; HS 2017/2022 splits (e.g. 1509 20 extra-virgin olive oil, formerly 1509 10) have no
    HS6 row, so fall back to the HS4 rows of the same heading (all 1509 rows: category 15, base 10% / AR 31.5%)."""
    rows_ = me_by_h6.get(c[:6]) or []
    src = 'NCM rows (HS6)' if rows_ else 'no schedule row'
    if not rows_ and me_by_h4.get(c[:4]):
        rows_ = me_by_h4[c[:4]]; src = 'NCM rows (HS4 fallback, HS6 code absent from NCM 2012)'
    bases, num, den = [], [0.0] * 11, 0.0
    for r in rows_:
        b = num_rate(r[COL[p]]); cat = r['Staging category'].replace(' ', '')
        if cat.startswith('TRQ-'):
            pth = [1 - DAIRY_PREF[y] for y in H] if cat != 'TRQ-4' else [1 - (0.30 + 0.70 * min(1.0, y / 7)) for y in H]
        elif cat in ('CH1', 'T1'):
            pth = [r_linear(9, y) for y in H]
        elif cat == 'CH2':
            pth = [r_linear(14, y) for y in H]
        else:
            pth = [r_category(cat, y) for y in H]
            if pth[0] is None:
                pth = [r_linear(18, y) for y in H]
        if b is not None:
            bases.append(b)
        w = b if (b is not None and b > 0) else 0.0
        num = [a + w * q for a, q in zip(num, pth)]; den += w
    base = sum(bases) / len(bases) if bases else None
    path = [a / den for a in num] if den > 0 else [r_linear(10, y) for y in H]
    applied = W_ME.get(p, {}).get(c[:6])
    applied = applied['rate'] / 100 if applied else None
    t0 = applied if base is None else (min(base, applied) if applied is not None else base)
    return t0, prorate0(path), src, base, applied

# ---------------- 5. assemble chapters ----------------
v5 = {c['hs']: c for c in json.load(open(V5, encoding='utf-8'))['chapters']}
NAMES = {'02': 'Meat and edible offal', '03': 'Fish and crustaceans', '07': 'Edible vegetables', '10': 'Cereals', '16': 'Meat and fish preparations',
         '17': 'Sugars and sugar confectionery', '21': 'Miscellaneous edible preparations', '28': 'Inorganic chemicals', '29': 'Organic chemicals',
         '32': 'Tanning and dyeing extracts, paints', '33': 'Essential oils, perfumery, cosmetics', '40': 'Rubber and articles', '44': 'Wood and articles',
         '48': 'Paper and paperboard', '49': 'Printed books and products', '64': 'Footwear', '74': 'Copper and articles', '90': 'Optical, medical instruments',
         '94': 'Furniture, lighting'}
NEW_WEDGE = {'02': (0.01, 'EUDR cattle'), '40': (0.01, 'EUDR rubber'), '44': (0.01, 'EUDR wood'), '48': (0.01, 'EUDR paper'),
             '49': (0.01, 'EUDR printed products'), '07': (0.01, 'Import-standards compliance, perishables')}
chapters, trq_members, diag = [], collections.defaultdict(list), []
for hs in included:
    base_imp = sum(v for (c, p), v in imp_lines.items() if c[:2] == hs)
    base_exp = sum(v for (c, p), v in exp_lines.items() if c[:2] == hs)
    rec = {'hs': hs, 'name': v5[hs]['name'] if hs in v5 else NAMES.get(hs, hs), 'baseline_imp': max(base_imp, 0.0), 'baseline_exp': max(base_exp, 0.0),
           'cagr_imp': growth(hs, 'M'), 'cagr_exp': growth(hs, 'X'), 'imp_elast': -3.5, 'exp_elast': -2.5,
           'wedge': v5[hs]['wedge'] if hs in v5 else NEW_WEDGE.get(hs, (0.0, ''))[0],
           'tier': v5[hs].get('tier', '') if hs in v5 else NEW_WEDGE.get(hs, (0.0, 'None'))[1],
           'reduction_type': 'Full', 'target_eu': 0.0, 'phase_eu': 7, 'phase_mer': 10, 'sensitivity': v5[hs].get('sensitivity', '') if hs in v5 else 'n/a'}
    eu_lines = []
    for (c, p), v in sorted(imp_lines.items(), key=lambda x: -x[1]):
        if c[:2] != hs or v <= 0 or base_imp <= 0:
            continue
        share = v / base_imp
        if p == 'BO':
            t0, tsrc = eu_tau0(c)
            eu_lines.append({'hs6': c[:6], 'cn8': c, 'share': share, 'tau0': t0 or 0.0, 'r_path': [1.0] * 11, 'partner': 'BOL', 'info': {'source': 'BOL not party', 'tau_src': tsrc}})
            diag.append(dict(flow='M', hs=hs, cn8=c, partner=p, baseline_usd=round(v), share=round(share, 4), tau0=t0, tau_src=tsrc, staging='n/a (Bolivia)', path_src='', trq=''))
            continue
        t0, tsrc = eu_tau0(c)
        if t0 is None:
            t0, tsrc = 0.0, 'no source -> 0'
        if OPTS['fish_atq_zero'] and c[:4] in ('0306', '0307', '0304', '0303'):
            t0 = 0.0
        path, trq_code, psrc, cats = eu_path_for(c)
        ln = {'hs6': c[:6], 'cn8': c, 'lid': f'{c}|{p}', 'share': share, 'tau0': t0, 'r_path': prorate0(path) if not trq_code else [1.0] * 11,
              'partner': 'MER4', 'origin': PARTY[p], 'info': {'source': psrc, 'cats': cats[:6], 'tau_src': tsrc}}
        if trq_code:
            ln['trq'] = {'code': trq_code, 'r_in': r_in_path(trq_code, t0)}
            trq_members[trq_code].append((hs, f'{c}|{p}', v))
        eu_lines.append(ln)
        diag.append(dict(flow='M', hs=hs, cn8=c, partner=p, baseline_usd=round(v), share=round(share, 4), tau0=t0, tau_src=tsrc,
                         staging='|'.join(sorted(set(cats)))[:40], path_src=psrc, trq=trq_code or ''))
    rec['eu_lines'] = eu_lines
    mer_lines = []
    for (c, p), v in sorted(exp_lines.items(), key=lambda x: -x[1]):
        if c[:2] != hs or v <= 0 or base_exp <= 0:
            continue
        share = v / base_exp
        if p == 'BO':
            mer_lines.append({'hs6': c[:6], 'cn8': c, 'partner': 'BOL', 'share': share, 'tau0': 0.0, 'r_path': [1.0] * 11, 'info': {'source': 'BOL not party'}})
            diag.append(dict(flow='X', hs=hs, cn8=c, partner=p, baseline_usd=round(v), share=round(share, 4), tau0=0, tau_src='n/a', staging='n/a (Bolivia)', path_src='', trq=''))
            continue
        t0, path, src, base, applied = mer_line(c, p)
        mer_lines.append({'hs6': c[:6], 'cn8': c, 'partner': PARTY[p], 'share': share, 'tau0': t0, 'r_path': path,
                          'info': {'source': src, 'schedule_base': base, 'wits_applied': applied}})
        diag.append(dict(flow='X', hs=hs, cn8=c, partner=p, baseline_usd=round(v), share=round(share, 4), tau0=t0,
                         tau_src=f'min(schedule {base}, WITS {applied})', staging=src, path_src='', trq=''))
    # unknown Mercosur tau0 -> chapter trade-weighted mean of known lines
    known = [(m['tau0'], m['share']) for m in mer_lines if m['tau0'] is not None]
    fb = sum(t * s for t, s in known) / sum(s for _, s in known) if known and sum(s for _, s in known) > 0 else 0.0
    for m in mer_lines:
        if m['tau0'] is None:
            m['tau0'] = fb
    rec['mer_lines'] = mer_lines
    rec['eu_mfn'] = sum(l['share'] * l['tau0'] for l in eu_lines)
    rec['mer_mfn'] = sum(l['share'] * l['tau0'] for l in mer_lines)
    chapters.append(rec)

# pooled notional Greek TRQ slices
for code, members in trq_members.items():
    totv = sum(v for _, _, v in members)
    for hs, lid, v in members:
        ch = next(x for x in chapters if x['hs'] == hs)
        c = lid.split('|')[0]
        for ln in ch['eu_lines']:
            if ln.get('lid') == lid and 'trq' in ln:
                uv = UVUSDKG.get(c, 3.0)
                cwe = EU_TRQ[code][3]
                ln['trq']['slice'] = [GREEK_SHARE * quota_t(code, y) * 1000 / cwe * uv * (v / totv if totv else 1) for y in H]

meta = {'opts': OPTS, 'spikes': {f'{k[0]}|{k[1]}|{k[2]}': {'year': y, 'removed_usd': a} for k, (y, a) in spikes.items()},
        'fx_usd_per_eur': FX, 'sources': {'trade': 'Eurostat COMEXT ds-045409 (CN8), accessed 2026-09-30', 'eu_tariff': 'EU TARIC third-country duty, SimDate 2023-07-01; WITS-TRAINS 2023 AVE fallback',
                                          'mer_tariff': 'OJ L 2026/184 Appendix 2-A-2 base rates capped at WITS-TRAINS 2023 applied MFN', 'staging': 'OJ L 2026/184 Appendices 2-A-1 / 2-A-2'}}
json.dump({'chapters': chapters, 'meta': meta}, open('config_C.json', 'w'), indent=0)
with open('lines_C.csv', 'w', newline='', encoding='utf-8') as f:
    w = csv.DictWriter(f, fieldnames=list(diag[0].keys())); w.writeheader(); w.writerows(diag)
tot_imp = sum(c['baseline_imp'] for c in chapters); tot_exp = sum(c['baseline_exp'] for c in chapters)
free = sum(l['share'] * c['baseline_imp'] for c in chapters for l in c['eu_lines'] if l['tau0'] == 0)
print(f'\nBaseline imports {tot_imp/1e6:.1f}m USD, exports {tot_exp/1e6:.1f}m; share of imports MFN duty-free {free/tot_imp:.1%}')
for c in sorted(chapters, key=lambda c: -c['baseline_imp'])[:15]:
    print(f"  HS {c['hs']} imp {c['baseline_imp']/1e6:7.2f} g {c['cagr_imp']*100:5.1f} | exp {c['baseline_exp']/1e6:6.2f} g {c['cagr_exp']*100:5.1f} | EU t0 {c['eu_mfn']*100:5.2f} Mer t0 {c['mer_mfn']*100:5.1f}")
print('TARIC notes:', {k: v for k, v in list(taric_note.items())[:15]})
