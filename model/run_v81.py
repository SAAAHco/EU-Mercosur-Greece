"""Run the v8.1 specification: central case, sensitivities and the tables the report needs."""
import json, os, copy
HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
import build_v81 as B
import pe_v81 as pe

B.DEFAULTS['growth_rule'] = 'endpoint'   # published rule, one-offs excluded
CAPW = dict(pe.CAP_WEIGHTS_V75)
CAPW.update({'02': 1.0, '07': 1.0, '10': 1.0, '16': 0.8, '17': 0.8, '21': 0.5, '03': 0.0})
GI = {'22': 9, '15': 4, '04': 3, '20': 2, '08': 1, '09': 1, '13': 1}   # Greek names in Annex 13-B (ouzo counted in 22)


def run(opts=None, imp_elast=None, agr=2026):
    cfg = B.build(opts)
    chs = pe.prepare_chapters(cfg['chapters'])
    return pe.headline(chs, CAPW, {}, agr=agr, imp_elast=imp_elast), cfg, chs


if __name__ == '__main__':
    h, cfg, chs = run()
    name = {c['hs']: c['name'] for c in chs}
    out = {'central': {k: v for k, v in h.items() if k != 'results'}}
    print(f"CENTRAL v8.1  Y10: imports {h['add_imp']/1e6:.2f}  exports adj {h['add_exp_adj']/1e6:.2f}  widening {h['widening']/1e6:.2f}  cap loss {h['cap_loss']/1e6:.2f}")
    pos = sorted([c for c in h['chapters'] if c['net'] > 0.05e6], key=lambda c: -c['net'])
    neg = sorted([c for c in h['chapters'] if c['net'] < -0.05e6], key=lambda c: c['net'])
    gp = sum(c['net'] for c in h['chapters'] if c['net'] > 0)
    gn = sum(c['net'] for c in h['chapters'] if c['net'] < 0)
    print(f"gross positive {gp/1e6:.2f} ({sum(1 for c in h['chapters'] if c['net']>0)} ch)  gross favorable {gn/1e6:.2f} ({sum(1 for c in h['chapters'] if c['net']<0)} ch)")
    print(f"{'HS':3s} {'chapter':34s} {'dImp':>7s} {'dExp':>7s} {'net':>7s} {'%gross+':>8s} {'GIs':>4s}")
    for c in pos + neg:
        print(f"{c['hs']:3s} {name[c['hs']][:34]:34s} {c['d_imp']/1e6:7.2f} {c['d_exp_adj']/1e6:7.2f} {c['net']/1e6:7.2f} {100*max(c['net'],0)/gp:8.1f} {GI.get(c['hs'],0):4d}")
    # GI placement
    gi_net = {hs: next((c['net'] for c in h['chapters'] if c['hs'] == hs), None) for hs in GI}
    print('GI chapters net (M):', {k: (round(v / 1e6, 2) if v is not None else 'not modelled') for k, v in gi_net.items()})
    # year profile
    r = h['results']
    yearly = []
    for y in range(11):
        ai = sum(x['years'][y]['proj_imp'] - x['years'][y]['cf_imp'] for x in r)
        ae = pe.adj_exports(r, y, 0.10, 0.07, CAPW)['total_adj']
        yearly.append([y, ai, ae, ai - ae])
    print('widening by year:', [round(w / 1e6, 1) for *_, w in yearly], '| Y5 share of Y10 widening %.0f%%' % (100 * yearly[5][3] / yearly[10][3]))
    out['yearly'] = yearly
    # scenario 2
    h2, *_ = run(agr=2027)
    print('S2 (2027 entry) widening %.2f' % (h2['widening'] / 1e6))
    out['s2'] = h2['widening']
    # grid and MC
    g = pe.grid(chs, CAPW, {})
    mc = pe.monte_carlo(chs, CAPW, {})
    print('grid span %.2f M; MC Y10 widening mean %.2f [%.2f, %.2f] (wedge and capacity parameters only)' % (
        (max(g.values()) - min(g.values())) / 1e6, mc['y10']['wid']['mean'] / 1e6, mc['y10']['wid']['lo'] / 1e6, mc['y10']['wid']['hi'] / 1e6))
    out['grid'] = g
    out['mc'] = mc
    # sensitivities
    sens = {}
    for e in [-2.0, -2.5, -3.5, -5.0, -8.0]:
        he, *_ = run(imp_elast=e)
        sens[f'elasticity {e}'] = he['widening']
    variants = {
        'tobacco uncertified (9.12%)': {'tobacco_tau0': 0.0912},
        'tobacco WITS AVE (5.97%)': {'tobacco_tau0': 0.0597},
        'fish processing quotas: base 0%': {'fish_atq_zero': True},
        'ethanol chemical-use quota added': {'ethanol_chem_quota': True},
        'spike screen off (one-offs and spikes kept)': {'exclude_oneoffs': False},
        'beef Hilton quota duty removal (with Greek slice)': {'hilton_quota': True},
        'growth on 3-year averages': {'growth_rule': 'avg3'},
        'spike screen also drops 2024-only flows': {'drop_new_2024': True},
        'growth capped at +/-10%': {'growth_cap': 0.10},
        'growth capped at +/-5%': {'growth_cap': 0.05},
        'no baseline growth (static)': {'static_growth': True},
    }
    for k, v in variants.items():
        hv, *_ = run(v)
        top = sorted(hv['chapters'], key=lambda c: -c['net'])
        sens[k] = hv['widening']
        print(f"  {k:42s} widening {hv['widening']/1e6:7.2f} | top: " + ', '.join(f"{c['hs']} {c['net']/1e6:.1f}" for c in top[:4]) + ' | bottom: ' + ', '.join(f"{c['hs']} {c['net']/1e6:.1f}" for c in top[-3:]))
    for k in [k for k in sens if k.startswith('elasticity')]:
        print(f"  {k:42s} widening {sens[k]/1e6:7.2f}")
    out['sensitivities'] = sens
    # tariff facts
    tot = sum(c['baseline_imp'] for c in cfg['chapters'])
    free = sum(l['share'] * c['baseline_imp'] for c in cfg['chapters'] for l in c['eu_lines'] if l['tau0'] == 0)
    nopref = sum(l['share'] * c['baseline_imp'] for c in cfg['chapters'] for l in c['eu_lines']
                 if l['tau0'] == 0 or (min(l['r_path']) >= 0.999 and 'trq' not in l))
    print('baseline imports %.1f M, MFN-free %.1f%%, no preference at all %.1f%%' % (tot / 1e6, 100 * free / tot, 100 * nopref / tot))
    out['baseline'] = {'imports': tot, 'mfn_free': free, 'no_pref': nopref, 'oneoffs': cfg['meta']['oneoffs']}
    # TRQ fill status at Y5/Y10
    for c in cfg['chapters']:
        for l in c['eu_lines']:
            if 'trq' in l and l['share'] * c['baseline_imp'] > 1e5:
                cfy = [l['share'] * pe.cf_imports(c, y, 2026, 2023) for y in (0, 5, 10)]
                print(f"  TRQ {l['trq']['code']:4s} {l['hs6']} cf {[round(x/1e6,2) for x in cfy]} slice {[round(l['trq']['slice'][y]/1e6,2) for y in (0,5,10)]}")
    json.dump(out, open('results_v81.json', 'w'), indent=0, default=float)
    json.dump(cfg, open('v81_config.json', 'w'), indent=0)
