"""Independent legal check of the staging used in config_C: for the top EU import lines, read the Mercosur
'Tariff preference' TARIC shows today (Year 0 of the agreement, in force since 1 May 2026) and compare its ad valorem
equivalent with base x r(0) implied by the staging category taken from the OJ schedule. Writes staging_check_C.csv."""
import json, csv, sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
import taric
cfg = json.load(open('config_C.json', encoding='utf-8'))
FX = {int(k): v for k, v in json.load(open('ecb_usd_per_eur.json')).items()}
rows = list(csv.DictReader(open('comext_cn8_tidy.csv', encoding='utf-8')))
uvn, uvd, part = {}, {}, {}
for r in rows:
    if r['flow'] == 'M' and 2022 <= int(r['year']) <= 2024 and r['partner'] != 'BO':
        c = r['cn8']; uvn[c] = uvn.get(c, 0) + float(r['value_eur']); uvd[c] = uvd.get(c, 0) + float(r['qty_100kg'])
        part.setdefault(c, {}); part[c][r['partner']] = part[c].get(r['partner'], 0) + float(r['value_eur'])
lines = []
for ch in cfg['chapters']:
    for l in ch['eu_lines']:
        if l['partner'] == 'MER4':
            lines.append((ch['baseline_imp'] * l['share'], ch['hs'], l))
lines.sort(key=lambda x: -x[0])
out = []
for base_usd, hs, l in lines[:45]:
    c = l['cn8']; uv = uvn[c] / uvd[c] if uvd.get(c) else None
    main_partner = max(part[c], key=part[c].get) if part.get(c) else 'BR'
    m = taric.measures(c, main_partner, '20260930')
    prefs = []
    tcd = None
    if m:
        for g in m['goods']:
            if g['third_country_duty']:
                tcd = g['third_country_duty']
            for e in g['measures']:
                if e['type'] in ('Tariff preference', 'Preferential tariff quota') and (e['area'] or '').lower().startswith('mercosur'):
                    duty = e['duty'] or taric.duty_at_unit_value({'third_country_duty': None, 'third_country_conditions': e.get('conditions')}, uv, c) or ''
                    prefs.append((g['goods_code'], e['type'], duty, e['start'], e['end']))
    # AVE of the preference vs model's year-0 unprorated rate
    pref_ave = None
    for gc, ty, duty, s, e_ in prefs:
        if ty == 'Tariff preference' and duty:
            a = taric.ave(duty, uv_eur_per_100kg=uv, uv_eur_per_hl=(taric.uv_hl(c, uv)))
            if a is not None:
                pref_ave = a if pref_ave is None else min(pref_ave, a)
    r0 = l['r_path'][0]
    r0_unprorated = 1 - (1 - r0) * 12 / 8 if cfg['meta']['opts'].get('year0_prorate') else r0
    model_y0 = l['tau0'] * r0_unprorated
    cat = '|'.join(l['info'].get('cats', [])[:3]) if isinstance(l['info'].get('cats'), list) else l['info'].get('cat', '')
    rec = dict(cn8=c, hs=hs, partner=main_partner, baseline_musd=round(base_usd / 1e6, 2), tau0_model=round(l['tau0'], 4), tau_src=l['info'].get('tau_src', ''),
               staging=cat, trq=l.get('trq', {}).get('code', ''), model_year0_rate=round(model_y0, 4),
               taric_third_country=tcd, taric_mercosur_pref=[p[2] for p in prefs][:3], taric_pref_ave=None if pref_ave is None else round(pref_ave, 4),
               match='' if pref_ave is None else ('OK' if abs(pref_ave - model_y0) < 0.005 else 'CHECK'))
    out.append(rec)
    print(f"{c} HS{hs} {main_partner} {rec['baseline_musd']:7.2f}m  tau0 {rec['tau0_model']:.4f} [{rec['tau_src'][:22]}] cat {cat[:12]:12s} trq {rec['trq']:4s} "
          f"model Y0 {model_y0:.4f} | TARIC TCD {str(tcd)[:45]:45s} | Mercosur pref {str([p[2] for p in prefs][:2])[:60]} -> AVE {rec['taric_pref_ave']} {rec['match']}")
with open('staging_check_C.csv', 'w', newline='', encoding='utf-8') as f:
    w = csv.DictWriter(f, fieldnames=list(out[0].keys())); w.writeheader(); w.writerows(out)
