"""Run Analysis C (independent inputs) through the independent engine, cross-check against the v8.1 engine, and
produce the A / B / C comparison plus decomposition hybrids and sensitivities. Writes results_C.json."""
import json, sys, io, os, copy, importlib.util
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import engine_C as E

R = os.environ.get('EUMG_PACKAGE', os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
A = json.load(open(os.path.join(R, 'model/inputs/v5_config.json'), encoding='utf-8'))['chapters']
B = json.load(open(os.path.join(R, 'outputs/v81_config.json'), encoding='utf-8'))['chapters']
C = json.load(open('config_C.json', encoding='utf-8'))['chapters']
sA, sB, sC = E.sign_wedges(A), E.sign_wedges(B), E.sign_wedges(C)
M = 1e6


def summary(chs, label, v81=True):
    h = E.headline(chs, capw=(E.CAPW_V81 if v81 else None), trq_caps=({} if v81 else None))
    top = sorted(h['per'].items(), key=lambda kv: -abs(kv[1]['net']))[:8]
    print(f"{label:34s} dM {h['dM']/M:7.2f}  dX_adj {h['dX_adj']/M:6.2f}  widening {h['widening']/M:7.2f}")
    for hs, v in top:
        print(f"      HS {hs}: net {v['net']/M:7.2f}  (imp {v['dM']/M:6.2f}, exp {v['dX_adj']/M:6.2f})")
    return h


out = {}
print('=== Year 10 widening, million USD ===')
out['A'] = summary(sA, 'A  published v7.5 (Comtrade HS2)', v81=False)
out['B'] = summary(sB, 'B  corrected v8.1 (Comtrade HS6)')
out['C'] = summary(sC, 'C  independent (COMEXT CN8 + TARIC)')

# cross-check: run the v8.1 engine file on config C
spec = importlib.util.spec_from_file_location('pe81', os.path.join(R, 'model/pe_v81.py')); pe81 = importlib.util.module_from_spec(spec); spec.loader.exec_module(pe81)
capw = dict(pe81.CAP_WEIGHTS_V75)
capw.update({'02': 1.0, '07': 1.0, '10': 1.0, '16': 0.8, '17': 0.8, '21': 0.5, '03': 0.0})
hB = pe81.headline(pe81.prepare_chapters(copy.deepcopy(C)), capw, {})
print(f"\ncross-check: v8.1 engine on config C -> widening {hB['widening']/M:.3f}  (engine_C gives {out['C']['widening']/M:.3f})")
out['C_pe81_engine'] = hB['widening']

# ---- decomposition hybrids ----
print('\n=== Decomposition: which input drives A -> C? ===')
# H1: COMEXT baselines/growth for the 25 chapters, but A's chapter-average tariffs and phase-outs
byC = {c['hs']: c for c in C}
h1 = []
for a in A:
    c = byC.get(a['hs'])
    x = copy.deepcopy(a)
    if c:
        x['baseline_imp'], x['baseline_exp'], x['cagr_imp'], x['cagr_exp'] = c['baseline_imp'], c['baseline_exp'], c['cagr_imp'], c['cagr_exp']
    h1.append(x)
out['H1'] = summary(E.sign_wedges(h1), 'H1 COMEXT trade + A chapter tariffs', v81=False)
# H2: A's trade data (Comtrade baselines/growth) with C's line-level tariffs and staging, 25 chapters only
h2 = []
for a in A:
    c = byC.get(a['hs'])
    if not c:
        continue
    x = copy.deepcopy(c)
    x['baseline_imp'], x['baseline_exp'], x['cagr_imp'], x['cagr_exp'] = a['baseline_imp'], a['baseline_exp'], a['cagr_imp'], a['cagr_exp']
    h2.append(x)
out['H2'] = summary(E.sign_wedges(h2), 'H2 A trade + C line tariffs (A chapters)')
# H3: C restricted to the 25 published chapters
h3 = [copy.deepcopy(byC[a['hs']]) for a in A if a['hs'] in byC]
out['H3'] = summary(E.sign_wedges(h3), 'H3 C, A chapters only')
# H4: C with feed and coffee forced to A's chapter-average tariffs (the two disputed inputs alone)
h4 = copy.deepcopy(C)
for x in h4:
    if x['hs'] in ('23', '09'):
        a = next(a for a in A if a['hs'] == x['hs'])
        x.pop('eu_lines', None); x.update({k: a[k] for k in ('eu_mfn', 'phase_eu', 'target_eu', 'reduction_type')})
out['H4'] = summary(E.sign_wedges(h4), 'H4 C + A tariffs on feed & coffee only')

# ---- sensitivities on C ----
print('\n=== Sensitivities on C (Year 10 widening) ===')
sens = {}
for e in [-2.0, -2.5, -3.5, -5.0, -8.0]:
    sens[f'elasticity {e}'] = E.headline(sC, eps_m=e, capw=E.CAPW_V81, trq_caps={})['widening']
sens['entry 2027 (S2)'] = E.headline(sC, agr=2027, capw=E.CAPW_V81, trq_caps={})['widening']
for wm in [0.5, 1.5, 3.0]:
    sens[f'wedge x{wm}'] = E.headline(sC, wedge_mult=wm, capw=E.CAPW_V81, trq_caps={})['widening']
for f2f in [0.07, 0.20]:
    sens[f'f2f {f2f}'] = E.headline(sC, f2f=f2f, capw=E.CAPW_V81, trq_caps={})['widening']
# TRQ bounds: no in-quota response at all vs uncapped in-quota response
c_notrq = copy.deepcopy(sC)
for x in c_notrq:
    for l in x.get('eu_lines', []):
        l.pop('trq', None)
sens['TRQ lines: no response (r=1)'] = E.headline(c_notrq, capw=E.CAPW_V81, trq_caps={})['widening']
c_trqfull = copy.deepcopy(sC)
for x in c_trqfull:
    for l in x.get('eu_lines', []):
        if 'trq' in l:
            l['r_path'] = l['trq']['r_in']; l.pop('trq')
sens['TRQ lines: in-quota rate, uncapped'] = E.headline(c_trqfull, capw=E.CAPW_V81, trq_caps={})['widening']
# growth off
c_g0 = copy.deepcopy(sC)
for x in c_g0:
    x['cagr_imp'] = 0.0; x['cagr_exp'] = 0.0
sens['no counterfactual growth'] = E.headline(c_g0, capw=E.CAPW_V81, trq_caps={})['widening']
for k, v in sens.items():
    print(f'   {k:40s} {v/M:7.2f}')
out['sens_C'] = sens
out['yearly_C'] = {y: E.widening(E.project(sC, trq_caps={}), y, E.CAPW_V81)['widening'] for y in range(11)}
print('\nyearly C:', {y: round(v / M, 1) for y, v in out['yearly_C'].items()})
mc = E.monte_carlo(sC, capw=E.CAPW_V81, trq_caps={})
print('Monte Carlo C (Y10 widening): mean %.2f  [%.2f, %.2f]' % (mc['mean'] / M, mc['lo'] / M, mc['hi'] / M))
out['mc_C'] = mc
json.dump({k: (v if not isinstance(v, dict) or 'per' not in v else {kk: vv for kk, vv in v.items() if kk != 'per'} | {'per': {hs: p for hs, p in v['per'].items()}}) for k, v in out.items()},
          open('results_C.json', 'w'), indent=1, default=float)
