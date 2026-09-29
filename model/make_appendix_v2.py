"""Generate the Online Appendix markup (tables from the v8.1 results) and supplementary figures."""
import json, os, sys, collections, csv
import re
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import pandas as pd
HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
import pe_v8, pe_v81 as pe
import run_v81 as R

OUTTXT = sys.argv[1]
FIGDIR = sys.argv[2]
os.makedirs(FIGDIR, exist_ok=True)
RED, BLUE, GRAY, INK, INK2, GRID = '#e34948', '#2a78d6', '#8a8983', '#0b0b0b', '#52514e', '#e4e3df'
plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 9, 'axes.spines.top': False, 'axes.spines.right': False,
                     'axes.edgecolor': INK2, 'xtick.color': INK2, 'ytick.color': INK2})

h, cfg, chs = R.run()
res = json.load(open('results_v81.json'))
C = {c['hs']: c for c in cfg['chapters']}
NET = {c['hs']: c for c in h['chapters']}
f1 = lambda x: f"{x/1e6:.1f}".replace('-', '−')
f2 = lambda x: f"{x/1e6:.2f}".replace('-', '−')
pct = lambda x: f"{100*x:.1f}%"
L = []
w = L.append

w("@title Online Appendix")
w("@subtitle Protest without Exposure: Why Greece Backed the EU-Mercosur Agreement")
w("")
w("This appendix documents the data, the tariff treatment of every traded line, the projection and its robustness, the detailed results, the comparison with a chapter-average specification, the location of the Greek protected designations and the documentary record of the Greek position. Appendix A describes the data and baselines, Appendix B the tariff treatment, Appendix C the projection engine and the treatment of tariff-rate quotas, Appendix D the detailed results and sensitivities, Appendix E the chapter-average comparison, Appendix F the protected designations and Appendix G the political chronology and the alternative accounts. Tables are numbered A.1 to A.9 and figures S1 to S4.")
w("")
# ---------------- Appendix A ----------------
w("# Appendix A. Data, chapter set and baselines")
w("")
w("Bilateral trade between Greece and Argentina, Bolivia, Brazil, Paraguay and Uruguay is taken from UN Comtrade at the two-digit level for 2014 to 2024 and at the six-digit level, by partner and year, for 2022 to 2024 (UN Comtrade, 2026). The six-digit records sum exactly to the two-digit totals for 2022 to 2024. The chapter set comprises every Harmonized System chapter in which bilateral trade in either direction exceeded one million US dollars in at least one year between 2014 and 2024 (43 chapters), plus dairy (HS 04), retained for the quota and designation tests; together they cover 44 chapters. Bolivia, which joined Mercosur in July 2024 but is not a party to the agreement, is kept in the baselines and receives no tariff preference.")
w("")
w("Baselines are 2022 to 2024 means. Because single shipments can dominate a three-year window at this resolution, a spike screen is applied to each line-partner flow: where one year exceeds 80% of the flow's 2022 to 2024 sum and one million dollars, that year is replaced by the mean of the other two years, so that a pure one-off drops out. Table A.1 lists the flows affected. Counterfactual growth follows the endpoint rule, the compound annual rate between 2014 and 2024 at chapter level (net of the screened amounts in 2024), bounded at plus or minus 15% a year; chapters with a zero endpoint take the bound in the direction of change. The counterfactual compounds the baseline from 2023 to the projection year. Growth rates computed on three-year averages (2014 to 2016 and 2022 to 2024) are reported as a sensitivity in Appendix D.")
w("")
w("@caption **Table A.1.** Flows adjusted by the spike screen, 2022 to 2024 (million US dollars removed from the flagged year).")
w("@widths 10,12,12,10,56")
w("@align l,l,l,r,l")
w("@table")
w("Flow | HS-6 | Partner | Removed | Description")
DESC = {'030743': 'Squid (Illex), 2024 spike', '230400': 'Soybean oilcake, 2024 (duty-free, no effect)', '271019': 'Petroleum oils, single-year shipment',
        '760200': 'Aluminum scrap, 2024 (duty-free)', '841012': 'Hydraulic turbines, single shipment', '470329': 'Wood pulp, 2022 (duty-free)',
        '270900': 'Crude petroleum, 2023 cargo (duty-free)', '200969': 'Grape must, 2024 only', '300420': 'Antibiotic preparations, 2022',
        '847790': 'Parts of plastics machinery, 2022', '999999': 'Commodities not elsewhere specified, 2022', '730511': 'Line pipe, 2023 project shipment',
        '842230': 'Filling and packaging machinery, 2023', '843810': 'Bakery machinery, 2024'}
for k, v in sorted(cfg['meta']['oneoffs'].items(), key=lambda kv: -sum(kv[1].values())):
    fl, h6, p = k.split('|')
    w(f"{'Imports' if fl == 'M' else 'Exports'} | {h6} | {p} | {sum(v.values())/1e6:.1f} | {DESC.get(h6, '')}")
w("@endtable")
w("@note Source: authors' calculations from UN Comtrade (2026).")
w("")
# ---------------- Appendix B ----------------
w("# Appendix B. Tariff treatment of traded lines")
w("")
w("Base tariffs on the EU side are applied most-favored-nation (MFN) duties, expressed as ad valorem equivalents. For all lines the 2023 ad valorem equivalents of the World Bank's Trade Analysis Information System (TRAINS) are the default (World Bank, 2026). For the 21 lines that carry the result, the eight-digit line of the Combined Nomenclature (CN) actually traded by Greece was identified from Eurostat Comext data for 2022 to 2024 (Eurostat, 2026a), its conventional duty was read from the EU's integrated tariff (TARIC, 2026), and compound or specific duties were valued at Greek import unit values (Table A.2). Where the TRAINS average mixes eight-digit lines that do not apply to Mercosur trade, such as juice rows valid only below a value threshold or wine rows reserved to wine produced in the Union, the traded-row value replaces it.")
w("")
w("Phase-outs are read from the schedules annexed to the Interim Agreement on Trade (European Union, 2026). Under Annex 2-A, the linear categories 4, 7, 8, 10 and 15 eliminate the base duty in equal annual cuts, the first at entry into force, so that the duty is zero from the first of January of the year the category names; category 0 eliminates the duty at entry into force; category E excludes the line; and the tariff-rate-quota categories open in-quota access at the rates and volumes of Section B (EU) and Section C (Mercosur). Table A.3 summarizes the categories that occur in Greek trade. Year 0 runs from 1 May to 31 December 2026, so the year-0 cut is applied to eight twelfths of the year. Where the schedule lists several eight-digit rows under a traded six-digit line with different categories, their paths are averaged with weights equal to their ad valorem base rates.")
w("")
w("On the Mercosur side, whose schedule uses the Mercosur Common Nomenclature (NCM), the base rate is the partner-specific rate in Appendix 2-A-2, or the partner's applied MFN rate where it is lower, as Article 2.4(7) provides. Staging categories are common to the four parties. Exports to Bolivia receive no preference.")
w("")
tc = list(csv.DictReader(open(os.path.join(HERE, 'inputs', 'tariff_check_lines.csv'), encoding='utf-8')))
w("@caption **Table A.2.** Tariff treatment of the traded lines that carry the result (EU imports from Mercosur).")
w("@widths 9,36,14,13,28")
w("@align l,l,r,l,l")
w("@table")
w("HS-6 | Traded eight-digit line and duty | Base used (%) | Category | Note")
NOTES = {'240120': 'certified Virginia/Burley; 9.12% uncertified (sensitivity)', '030617': 'upper bound; 0% processing quotas exist',
         '030743': 'upper bound; 0% processing quotas exist', '030474': 'upper bound; 0% processing quotas exist',
         '020230': 'in-quota base; new quota at 7.5%', '020130': 'Hilton base; removal treated as rent (Greek Hilton slice at 0% in sensitivity, Table A.7)',
         '220720': 'specific duty at Greek prices; all-uses quota at one third', '170199': 'excluded from preferences',
         '080550': 'lemon rate (WITS applies the lime rate)', '200919': 'traded row; value-threshold rows dropped', '200911': 'traded row',
         '200969': 'ad valorem part; entry price not triggered', '230990': 'specific duty at Greek prices', '080830': 'seasonal duty, monthly weights',
         '271019': 'end-use relief for refinery feedstock', '020714': 'specific duty at Greek prices; poultry quota at 0%'}
for r in tc:
    h6 = r['hs6']
    base = r['recommended_tau0_pct'].split(' ')[0]
    cat = r['staging_of_traded_CN8'].replace('EL TRQ', 'EL (quota)').replace('BF1 + Hilton', 'BF1 (quota)')
    cell = f"{r['traded_CN8_share_2022_24']}; {r['conventional_duty_TARIC']}".replace('authorised', 'authorized')
    cell = re.sub(r'\b(BR|AR|UY|PY)\b', lambda m: {'BR': 'BRA', 'AR': 'ARG', 'UY': 'URY', 'PY': 'PRY'}[m.group(1)], cell)
    w(f"{h6} | {cell} | {base} | {cat} | {NOTES.get(h6, '')}")
w("030366 | 03036612 ARG, URY (Argentine hake, frozen whole); 15% | 15.0 | 0 | upper bound; 0% processing quotas exist")
w("020629 | 02062999 ARG (frozen bovine offal, other); Free | 0.0 | 0 | traded row duty-free")
w("@endtable")
w("@note Duties valued at Greek unit values for 2022 to 2024. ATQ: autonomous tariff quota; EIF: entry into force; SIV: standard import value; UV: unit value; MIN and MAX: minimum and maximum specific duty. Source: authors' calculations from Eurostat (2026a) and TARIC (2026); European Union (2026).")
w("")
w("@caption **Table A.3.** Staging categories occurring in Greek-Mercosur trade (Annex 2-A, Section A).")
w("@widths 14,86")
w("@align l,l")
w("@table")
w("Category | Treatment")
for cat, t in [('0', 'Duty eliminated at entry into force'), ('4, 7, 8, 10, 15', 'Base duty eliminated in 5, 8, 9, 11 or 16 equal annual cuts, the first at entry into force; duty-free from 1 January of year 4, 7, 8, 10 or 15'),
               ('50 %', 'Base duty reduced by half over five equal cuts'), ('10/EP, 0/EP', 'Ad valorem component eliminated as for 10 or 0; entry-price specific duty retained'), ('SW/12', 'Sparkling wine: duty-free at entry into force if the customs value is at least 8 US dollars per liter; otherwise the base rate for 12 years'), ('10/OS ≥ 70 %', 'Products with less than 70% sugar eliminated in 11 equal annual cuts; products with 70% sugar or more enter only under the other-sugars quota'),
               ('E', 'Excluded: base rate maintained'), ('BF1, BF2, PY1, PY2, PK, ME, RE, SR, EL, HY, RM, SC, SH2', 'EU tariff-rate quotas (beef, poultry, pigmeat, maize and sorghum, rice, sugar for refining, ethanol, honey, rum, sweetcorn, starch derivatives): in-quota rate for quota volumes; base rate out of quota'),
               ('TRQ-1 to TRQ-4', 'Mercosur tariff-rate quotas for EU dairy and garlic: preference rising from 10% (30% for garlic) to 100%'), ('CH2', 'Mercosur tariff-rate quotas for EU chocolate and white chocolate: in-quota duty reduced in stages, duty-free from year 14')]:
    w(f"{cat} | {t}")
w("@endtable")
w("@note Source: European Union (2026), Annex 2-A.")
w("")
# ---------------- Appendix C ----------------
w("# Appendix C. Projection engine, tariff-rate quotas and robustness")
w("")
w("For chapter *c* in projection year *y*, imports are projected as")
w("")
w("@eq *M*(*c*, *y*) = *M*^{cf}(*c*, *y*) × [1 + *ε*_{i} × *π*(*c*, *y*) × *W*(*c*, *y*)] ||| (C.1)")
w("")
w("where *M*^{cf} is the counterfactual, *ε*_{i} = −3.5 the Armington import elasticity (Armington, 1969; Hertel et al., 2007), and *π*(*c*, *y*) the chapter's proportional change in the tariff-inclusive price, computed as the baseline-share-weighted sum of line-level changes, *π*(*c*, *y*) = Σ_{l} *s*_{l} [(1 + *τ*_{l}(*y*))/(1 + *τ*_{l0}) − 1]. Lines that are duty-free at base therefore contribute zero. *W* is the compliance-cost multiplier: exporter-borne wedges of −1.0 percentage points for products covered by the EU Deforestation Regulation, −0.5 to −3.0 points for products covered by the Carbon Border Adjustment Mechanism, and −1.0 point for import-standards compliance on perishables, calibrated on Rijk and Kuepper (2025) and Bonnet et al. (2026). Because the wedge multiplies the tariff response, it has no effect on lines without a tariff cut, and across the whole projection it moves the Year 10 result by less than 1.5 million dollars. Exports are projected analogously on the Mercosur schedules with *ε*_{x} = −2.5 and no wedge, and agricultural exports carry a capacity adjustment for Farm to Fork compliance and seasonal labor (10% and 7% at Year 10, weighted by agricultural content), which reduces Year 10 exports by 3.3 million dollars.")
w("")
w("Tariff-rate quotas are treated at the margin. Greece is assigned a notional slice of each EU quota equal to its 2.4% population share, converted from carcass weight where the quota is so defined (130% for boneless beef, 140% for boneless poultry), pooled across the lines the quota covers, valued at Greek import unit values and pro-rated for 2026. If projected counterfactual imports of a line exceed its slice, the marginal unit pays the out-of-quota duty and the line shows no quantity response; below the slice, the in-quota price change applies, capped so that projected imports do not exceed the slice. Table A.4 reports the slices for the quota lines Greece trades. The existing Hilton beef quotas, whose 20% in-quota duty the agreement removes at entry into force, are treated as a transfer of quota rent in the central case, since those quotas are filled at Union level; a sensitivity gives Greece a notional 2.4% slice of the Hilton volumes at 0%, filled before the new beef quotas. Where a line carries two quotas, the cheaper fills first.")
w("")
cfy = lambda c, share, y: share * pe.cf_imports(c, y, 2026, 2023)
w("@caption **Table A.4.** Notional Greek quota slices against projected counterfactual Greek imports, quota lines (million US dollars).")
w("@widths 10,10,26,18,18,18")
w("@align l,l,l,r,r,l")
w("@table")
w("Quota | HS-6 | Product | Slice, Year 5 | Counterfactual imports, Year 5 | Status at Year 10")
PROD = {'020130': 'Fresh boneless beef', '020230': 'Frozen boneless beef', '020714': 'Frozen boneless chicken', '100590': 'Maize', '100630': 'Rice',
        '220720': 'Denatured ethanol', '220710': 'Undenatured ethanol', '040900': 'Natural honey', '160249': 'Pigmeat preparations', '100620': 'Husked rice', '170114': 'Raw cane sugar for refining'}
for c in cfg['chapters']:
    for l in c['eu_lines']:
        if 'trq' in l:
            s5 = l['trq']['slice'][5]; m5 = cfy(c, l['share'], 5); m10 = cfy(c, l['share'], 10); s10 = l['trq']['slice'][10]
            w(f"{l['trq']['code']} | {l['hs6']} | {PROD.get(l['hs6'], l['hs6'])} | {s5/1e6:.2f} | {m5/1e6:.2f} | {'binding' if m10 >= s10 else 'not binding'}")
w("@endtable")
w("@note Slices are 2.4% of the EU quota volume in product weight, valued at Greek import unit values and pooled across the lines covered. Source: authors' calculations; European Union (2026), Annex 2-A, Section B.")
w("")
w("The robustness exercises reported here perturb only the wedge multiplier and the capacity drags. A deterministic grid over four wedge multipliers and three Farm to Fork drags spans " + f2(max(res['grid'].values()) - min(res['grid'].values())) + " million dollars at Year 10, and a 1,000-draw Monte Carlo simulation with triangular priors gives a mean of " + f2(res['mc']['y10']['wid']['mean']) + " million with a 95% interval of " + f2(res['mc']['y10']['wid']['lo']) + " to " + f2(res['mc']['y10']['wid']['hi']) + " million. These intervals exclude the dominant uncertainties, which are the elasticities, the counterfactual growth rule and the customs regime of particular lines; those are reported as sensitivities in Table A.7.")
w("")
# ---------------- Appendix D ----------------
w("# Appendix D. Detailed results and sensitivities")
w("")
w("@caption **Table A.5.** Year-by-year projection, central case (million US dollars).")
w("@widths 16,28,28,28")
w("@align c,r,r,r")
w("@table")
w("Year | Additional imports | Additional exports, capacity-adjusted | Widening")
for y, ai, ae, wd in res['yearly']:
    w(f"{int(y)} ({2026+int(y)}) | {f1(ai)} | {f1(ae)} | {f1(wd)}")
w("@endtable")
w(f"@note Entry into force in 2027 instead of 2026 gives a Year 10 widening of {f1(res['s2'])} million dollars. Source: authors' calculations.")
w("")
w("@caption **Table A.6.** Year 10 change in the bilateral balance by chapter, central case (million US dollars).")
w("@widths 8,38,14,14,14,12")
w("@align l,l,r,r,r,r")
w("@table")
w("HS | Chapter | Additional imports | Additional exports | Net widening | Baseline imports")
rows = sorted(h['chapters'], key=lambda c: -c['net'])
for c in rows:
    if abs(c['d_imp']) < 5e3 and abs(c['d_exp_adj']) < 5e3:
        continue
    w(f"{c['hs']} | {C[c['hs']]['name']} | {f2(c['d_imp'])} | {f2(c['d_exp_adj'])} | {f2(c['net'])} | {f1(C[c['hs']]['baseline_imp'])}")
rest = [c for c in rows if abs(c['d_imp']) < 5e3 and abs(c['d_exp_adj']) < 5e3]
w(f"Other | {len(rest)} chapters with no tariff change or negligible trade | {f2(sum(c['d_imp'] for c in rest))} | {f2(sum(c['d_exp_adj'] for c in rest))} | {f2(sum(c['net'] for c in rest))} | {f1(sum(C[c['hs']]['baseline_imp'] for c in rest))}")
w(f"**Total** | **44 chapters** | **{f2(h['add_imp'])}** | **{f2(h['add_exp_adj'])}** | **{f2(h['widening'])}** | **{f1(sum(c['baseline_imp'] for c in cfg['chapters']))}**")
w("@endtable")
w("@note Negative values indicate additional exports exceeding additional imports. Source: authors' calculations.")
w("")
w("@caption **Table A.7.** Year 10 widening under alternative assumptions (million US dollars).")
w("@widths 70,30")
w("@align l,r")
w("@table")
w(f"Assumption | Year 10 widening")
w(f"Central case | {f1(res['central']['widening'])}")
LAB = {'elasticity -2.0': 'Import elasticity −2.0', 'elasticity -2.5': 'Import elasticity −2.5', 'elasticity -3.5': None, 'elasticity -5.0': 'Import elasticity −5.0',
       'elasticity -8.0': 'Import elasticity −8.0', 'tobacco uncertified (9.12%)': 'Tobacco leaf at the uncertified duty (9.12%)',
       'tobacco WITS AVE (5.97%)': 'Tobacco leaf at the TRAINS average (5.97%)', 'fish processing quotas: base 0%': 'Seafood entering under 0% processing quotas',
       'ethanol chemical-use quota added': 'Ethanol chemical-use quota (0%) added', 'spike screen off (one-offs and spikes kept)': 'Spike screen off',
       'beef Hilton quota duty removal (with Greek slice)': 'Hilton beef quota duty removed, with a Greek slice', 'spike screen also drops 2024-only flows': 'Spike screen also drops flows first seen in 2024', 'growth on 3-year averages': 'Growth on three-year averages',
       'growth capped at +/-10%': 'Growth bounded at ±10% a year', 'growth capped at +/-5%': 'Growth bounded at ±5% a year', 'no baseline growth (static)': 'No counterfactual growth'}
for k, v in res['sensitivities'].items():
    if LAB.get(k):
        w(f"{LAB[k]} | {f1(v)}")
w("@endtable")
w("@note Source: authors' calculations.")
w("")
# ---------------- Appendix E ----------------
v5 = pe_v8.prepare_chapters(json.load(open(os.path.join(HERE, 'inputs', 'v5_config.json')))['chapters'])
h75 = pe_v8.headline(v5, pe_v8.CAP_WEIGHTS_V75, pe_v8.TRQ_CAPS_V75)
n75 = {c['hs']: c['net'] for c in h75['chapters']}
w("# Appendix E. Comparison with a chapter-average specification")
w("")
w(f"A chapter-average specification assigns each of 25 chapters a single EU tariff and phase-out equal to the simple average over the chapter's tariff lines, together with a single Mercosur tariff and phase-out. It gives a Year 10 widening of {f1(h75['widening'])} million US dollars, with {f1(h75['add_imp'])} million of additional imports and {f1(h75['add_exp_adj'])} million of adjusted exports. Animal feed contributes {f1(n75['23'])} million, beverages {f1(n75['22'])} million and coffee {f1(n75['09'])} million; the beverages figure arises because Brazilian ethanol, which the agreement liberalizes only within a quota, is assigned the chapter's 14% duty on a ten-year phase-out. Both specifications run through the same engine, so the difference between the two results is due entirely to the inputs.")
w("")
w("Three features of the chapter-average inputs produce the difference. First, duty-free lines receive the chapter's average duty: Greek feed imports from Mercosur are soybean oilcake (CN 2304) and Greek coffee imports green coffee (CN 0901 11), both duty-free at MFN and listed with a base of 'Free' in the EU schedule, yet they are assigned 5.0% and 3.0%. Trade weighting would remove most of this first error, since soybean oilcake and green coffee dominate their chapters; the second and third are properties of aggregation as such. Second, the 25-chapter set omits nineteen chapters that meet the inclusion rule used here, including meat, fish, cereals and sugar, and records no tobacco imports. Third, phase-outs are assigned by chapter: Greek olive oil exports, for example, are treated as liberalized over six years, whereas Mercosur's schedule phases out the duty on olive oil over fifteen. Figure 1 of the article compares the chapter results of the two specifications.")
w("")
# ---------------- Appendix F ----------------
w("# Appendix F. Greek protected designations and the projected balance")
w("")
w("@caption **Table A.8.** Greek geographical indications protected by the agreement, by chapter, and the chapter's projected Year 10 balance.")
w("@widths 10,60,15,15")
w("@align l,l,r,r")
w("@table")
w("HS | Protected names | Number | Year 10 net (million US dollars)")
GIS = [('22', 'Amyntaio, Mantineia, Naoussa, Nemea, Retsina of Attiki, Samos, Santorini, Tsipouro, Ouzo (with Cyprus)', 9),
       ('15', 'Kalamata, Kolymvari Chanion Kritis, Lygourio Asklipiou, Sitia Lasithiou Kritis (olive oils)', 4),
       ('04', 'Feta, Kefalograviera, Manouri', 3), ('20', 'Elia Kalamatas, Konservolia Amfissis (table olives)', 2),
       ('08', 'Korinthiaki Stafida Vostitsa (currant)', 1), ('09', 'Krokos Kozanis (saffron)', 1), ('13', 'Masticha Chiou (mastic)', 1)]
for hs, names, n in GIS:
    net = f2(NET[hs]['net']) if hs in NET else 'not modeled'
    w(f"{hs} | {names} | {n} | {net}")
w("@endtable")
w("@note Chapter 13 falls below the inclusion threshold. Greek table olives (NCM 2005.70) are excluded from Mercosur's tariff preferences. Source: agreement annex on geographical indications as reported in iefimerida (2026b); authors' calculations.")
w("")
open(OUTTXT, 'w', encoding='utf-8').write('\n'.join(L) + '\n')

# ---------------- Figures ----------------
raw = pd.read_excel(os.path.join(HERE, 'inputs', 'Greece__Mar_Total_TradeData_Balance_Sheet.xlsx'), 'Sheet1')
g = raw.groupby(['refYear', 'flowCode'])['primaryValue'].sum().unstack() / 1e6
fig, ax = plt.subplots(figsize=(6.4, 3.4))
ax.plot(g.index, g['M'], color=RED, lw=2, marker='o', ms=4, label='Greek imports from Mercosur')
ax.plot(g.index, g['X'], color=BLUE, lw=2, marker='o', ms=4, label='Greek exports to Mercosur')
ax.yaxis.grid(True, color=GRID, lw=0.6); ax.set_axisbelow(True)
ax.set_ylabel('million US dollars'); ax.legend(frameon=False, fontsize=8)
fig.savefig(os.path.join(FIGDIR, 'Figure_S1.png'), dpi=300, bbox_inches='tight', facecolor='white'); fig.savefig(os.path.join(FIGDIR, 'Figure_S1.eps'), bbox_inches='tight'); plt.close(fig)
ys = [r[0] for r in res['yearly']]
fig, ax = plt.subplots(figsize=(6.4, 3.4))
ax.plot(ys, [r[1] / 1e6 for r in res['yearly']], color=RED, lw=2, label='Additional imports')
ax.plot(ys, [r[2] / 1e6 for r in res['yearly']], color=BLUE, lw=2, label='Additional exports, capacity-adjusted')
ax.plot(ys, [r[3] / 1e6 for r in res['yearly']], color=INK, lw=2, ls='--', label='Widening of the bilateral balance')
ax.yaxis.grid(True, color=GRID, lw=0.6); ax.set_axisbelow(True)
ax.set_xlabel('Year after entry into force'); ax.set_ylabel('million US dollars'); ax.legend(frameon=False, fontsize=8)
fig.savefig(os.path.join(FIGDIR, 'Figure_S4.png'), dpi=300, bbox_inches='tight', facecolor='white'); fig.savefig(os.path.join(FIGDIR, 'Figure_S4.eps'), bbox_inches='tight'); plt.close(fig)
# S3: duty-free share of baseline imports, largest import chapters
top = sorted(cfg['chapters'], key=lambda c: -c['baseline_imp'])[:12][::-1]
fig, ax = plt.subplots(figsize=(6.4, 3.8))
for i, c in enumerate(top):
    free = sum(l['share'] for l in c['eu_lines'] if l['tau0'] == 0 or l['partner'] == 'BOL') * c['baseline_imp'] / 1e6
    dut = c['baseline_imp'] / 1e6 - free
    ax.barh(i, free, color=GRAY, height=0.6, label='Duty-free at MFN or no preference' if i == 0 else None)
    ax.barh(i, dut, left=free + 0.8, color=RED, height=0.6, label='Dutiable' if i == 0 else None)
ax.set_yticks(range(len(top))); ax.set_yticklabels([f"HS {c['hs']} {c['name'][:30]}" for c in top], fontsize=8)
ax.xaxis.grid(True, color=GRID, lw=0.6); ax.set_axisbelow(True)
ax.set_xlabel('Baseline Greek imports from Mercosur, 2022 to 2024 mean, million US dollars'); ax.legend(frameon=False, fontsize=8, loc='lower right')
fig.savefig(os.path.join(FIGDIR, 'Figure_S2.png'), dpi=300, bbox_inches='tight', facecolor='white'); fig.savefig(os.path.join(FIGDIR, 'Figure_S2.eps'), bbox_inches='tight'); plt.close(fig)
# S4: quota slices vs counterfactual, beef and ethanol lines
fig, axs = plt.subplots(1, 3, figsize=(6.6, 2.6), sharey=False)
for ax, h6 in zip(axs, ['020130', '020230', '220720']):
    c = next(c for c in cfg['chapters'] if any(l['hs6'] == h6 and 'trq' in l for l in c['eu_lines']))
    l = next(l for l in c['eu_lines'] if l['hs6'] == h6 and 'trq' in l)
    yy = list(range(11))
    ax.plot(yy, [l['trq']['slice'][y] / 1e6 for y in yy], color=BLUE, lw=2, label='Notional Greek slice')
    ax.plot(yy, [cfy(c, l['share'], y) / 1e6 for y in yy], color=RED, lw=2, label='Counterfactual imports')
    ax.set_title(PROD[h6], fontsize=8, color=INK); ax.yaxis.grid(True, color=GRID, lw=0.6); ax.set_axisbelow(True); ax.set_xlabel('Year', fontsize=8)
axs[0].set_ylabel('million US dollars'); axs[0].legend(frameon=False, fontsize=7, loc='lower right')
fig.tight_layout()
fig.savefig(os.path.join(FIGDIR, 'Figure_S3.png'), dpi=300, bbox_inches='tight', facecolor='white'); fig.savefig(os.path.join(FIGDIR, 'Figure_S3.eps'), bbox_inches='tight'); plt.close(fig)
print('appendix and figures written')
