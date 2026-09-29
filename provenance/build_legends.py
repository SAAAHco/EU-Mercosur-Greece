"""Build staging_category_legend.csv and trq_summary.csv from the parsed OJ L 2026/184 schedules
and the Annex 2-A Section A-C text (OJ_L_202600184_annex2A_sectionsA-E_staging_TRQ_notes.txt)."""
import csv, collections, os
os.chdir(os.path.dirname(os.path.abspath(__file__)))
eu = list(csv.reader(open('appendix_2-A-1_EU_schedule_from_OJ_xhtml.csv', encoding='utf-8')))[1:]
me = list(csv.reader(open('appendix_2-A-2_MERCOSUR_schedule_from_OJ_xhtml.csv', encoding='utf-8')))[1:]
ceu = collections.Counter(r[3] for r in eu)
cme = collections.Counter(r[6] for r in me)
R = 'Annex 2-A Sec A para 6'
GE = '≥'  # >=
leg = [
    ('0', 'EU+MERCOSUR', 'linear elimination', 1, '0', '100', 'Duty eliminated at entry into force (EIF); duty-free from EIF', R + '(a)'),
    ('4', 'EU+MERCOSUR', 'linear elimination', 5, '4', '100', '5 equal annual cuts (20% per step, first at EIF); duty-free 1 Jan year 4', R + '(b)'),
    ('7', 'EU+MERCOSUR', 'linear elimination', 8, '7', '100', '8 equal annual cuts (12.5%); duty-free 1 Jan year 7', R + '(c)'),
    ('8', 'EU+MERCOSUR', 'linear elimination', 9, '8', '100', '9 equal annual cuts (11.1%); duty-free 1 Jan year 8', R + '(d)'),
    ('10', 'EU+MERCOSUR', 'linear elimination', 11, '10', '100', '11 equal annual cuts (9.1%); duty-free 1 Jan year 10', R + '(e)'),
    ('15', 'EU+MERCOSUR (used only in Appendix 2-A-2)', 'linear elimination', 16, '15', '100', '16 equal annual cuts (6.25%); duty-free 1 Jan year 15', R + '(g)'),
    ('15V', 'MERCOSUR', 'back-loaded elimination + TRQ', 9, '15', '100',
     'Vehicles (14 NCM-2022 items in 8701/8703/8704): base rate to end of year 6; cumulative cuts 19/38.1/57.1/64.3/71.4/78.6/85.7/92.9/100% in years 7-15; '
     'plus 50% cut of base within a 50,000-unit annual quota from EIF to end of year 8 (AR 15,500; BR 32,000; PY 750; UY 1,750 units)', R + '(h) + Chronogram table'),
    ('See comments', 'MERCOSUR', 'special vehicle regime', '', '18 / 25 / 30', '100',
     'NCM 87039000: EV and hybrids (HS2022 8703.40-.80 excl. H2) 28.6% cut at EIF to end year 5 (AR/BR 25%, UY 16.4%, PY 14.3%), duty-free 1 Jan year 18; '
     'hydrogen fuel-cell (ex 8703.80) base to year 6, 28.6% cut years 7-12, duty-free year 25; HS2022 8703.90 base to year 6, 28.6% cut years 7-17, duty-free year 30',
     R + '(g),(i),(j),(k)'),
    ('SW/12', 'EU+MERCOSUR', 'conditional', 1, '0 or 12', '100', 'Sparkling wine: duty-free at EIF if customs value >= 8 USD FOB/litre; otherwise base rate for 12 years, duty-free 1 Jan year 12', R + '(f)'),
    ('4-EG', 'EU', 'linear elimination (conditional)', 5, '4', '100', 'Shell eggs 04072100, 04079010: 5 equal cuts, duty-free year 4; requires animal-welfare certificate (Dir. 1999/74/EC or equivalent)', R + '(o)'),
    ('FP30 %', 'EU (MERCOSUR spells it FP 30 %)', 'partial reduction', 1, 'never', '30', 'Fixed 30% reduction of base rate at EIF; stays at 70% of base (butter)', R + '(p)'),
    ('FP50 %', 'EU (MERCOSUR spells it FP 50 %)', 'partial reduction', 1, 'never', '50', 'Fixed 50% reduction of base rate at EIF; stays at 50% of base (yoghurt)', R + '(q)'),
    ('50 %', 'EU', 'partial reduction', 5, 'never', '50', '50% reduction phased in over 5 equal annual stages; 50% of base rate from 1 Jan year 4', R + '(r)'),
    ('0/EP', 'EU', 'entry-price line', 1, '0', 'ad valorem only', 'Ad valorem component eliminated at EIF; entry-price specific duty (triggered when import price < entry price) maintained', R + '(s)'),
    ('7/EP', 'EU', 'entry-price line', 8, '7', 'ad valorem only', 'Ad valorem component eliminated in 8 equal stages; entry-price specific duty maintained', R + '(t)'),
    ('10/EP', 'EU', 'entry-price line', 11, '10', 'ad valorem only', 'Ad valorem component eliminated in 11 equal stages; entry-price specific duty maintained', R + '(u)'),
    ('E', 'EU+MERCOSUR', 'exclusion', '', 'never', '0', 'Excluded from tariff preferences; remains at base rate', R + '(v)'),
    ('BA', 'EU', 'fixed specific rate', 1, 'never', 'to 75 EUR/t', 'Bananas 08039010: 75 EUR/t from EIF (base 136 EUR/t); Notes column "75EUR/t at EIF"', R + '(w)'),
    ('0 + EA/10; OS ' + GE + ' 70 %', 'EU', 'mixed (ad valorem + agricultural component EA)', 11, '10', '100',
     'Ad valorem eliminated at EIF; EA for products <70% sugar cut in 11 equal stages, duty-free year 10; products >=70% sugar only via TRQ-OS (Annex text writes "0 + 10 EA / OS >= 70 %")', R + '(x)'),
    ('10/OS ' + GE + ' 70 %', 'EU', 'mixed', 11, '10', '100', 'Products <70% sugar: 11 equal stages, duty-free year 10; >=70% sugar: TRQ-OS only', R + '(y)'),
    ('CE/E', 'EU', 'TRQ + exclusion', '', '', '', '04061020 fresh cheese: inside TRQ-CE, but mozzarella excluded (Notes: "E for mozzarella")', 'Appendix 2-A-1 Notes column'),
]
trq_eu = [('BF1', 'Sec B para 1', 'fresh/chilled beef'), ('BF2', 'Sec B para 3', 'frozen beef incl. for processing'),
          ('PK', 'Sec B para 4', 'pigmeat'), ('PY1', 'Sec B para 5', 'boneless poultry and preparations'),
          ('PY2', 'Sec B para 6', 'bone-in poultry'), ('MP', 'Sec B para 7', 'milk powders'), ('CE', 'Sec B para 8', 'cheese'),
          ('IF', 'Sec B para 9', 'infant formula'), ('ME', 'Sec B para 10', 'maize and sorghum'), ('RE', 'Sec B para 11', 'rice'),
          ('SR', 'Sec B para 12', 'raw cane sugar for refining'), ('OS', 'Sec B para 13', 'other sugars'), ('EG1', 'Sec B para 14', 'egg products'),
          ('EG2', 'Sec B para 15', 'egg albumins'), ('HY', 'Sec B para 16', 'honey'), ('RM', 'Sec B para 17', 'rum/cane spirits'),
          ('SC', 'Sec B para 18', 'sweetcorn'), ('SH1', 'Sec B para 19', 'maize and manioc starch'), ('SH2', 'Sec B para 20', 'starch derivatives'),
          ('EL', 'Sec B para 21', 'ethanol'), ('GC', 'Sec B para 22', 'garlic')]
for k, ref, name in trq_eu:
    leg.append((k, 'EU', 'TRQ', '', '', '', 'TRQ-' + k + ' (' + name + '); in-quota rate and volumes in Annex 2-A ' + ref + '; out-of-quota = base rate (no liberalisation)',
                'Annex 2-A Sec A para 8; ' + ref))
for k, ref, name in [('TRQ-1', 'Sec C para 1', 'milk powders'), ('TRQ-2', 'Sec C para 2', 'cheese'), ('TRQ-3', 'Sec C para 3', 'infant formula'), ('TRQ-4', 'Sec C para 4', 'garlic')]:
    leg.append((k, 'MERCOSUR', 'TRQ', '', '', '', k + ' (' + name + '); in-quota preference and volumes in Annex 2-A ' + ref + '; out-of-quota = base rate', 'Annex 2-A Sec A para 8; ' + ref))
leg += [
    ('CH1', 'MERCOSUR', 'TRQ with declining in-quota rate', '', '9', '100', 'Chocolate NCM 1806.20 / 1806.90: in-quota rate falls yearly and quota rises; quota and out-of-quota duty end in year 9 (duty-free); PY out-of-quota 2% on 1806.90', R + '(l)'),
    ('CH2', 'MERCOSUR', 'TRQ with declining in-quota rate', '', '14', '100', 'NCM 1704.90.10, 1806.10, 1806.31, 1806.32: declining in-quota rates, duty-free year 14', R + '(m)'),
    ('T1', 'MERCOSUR', 'TRQ with declining in-quota rate', '', '9', '100', 'Prepared tomatoes NCM 2002.10: 7,500 t quota; in-quota 12.6% (year 0) to 1.4% (year 8); duty-free year 9; out-of-quota 14% until then', R + '(n)'),
]
with open('staging_category_legend.csv', 'w', newline='', encoding='utf-8') as f:
    w = csv.writer(f)
    w.writerow(['code_as_in_schedule', 'applies_to', 'type', 'n_stages', 'duty_free_from_year', 'final_cut_pct_of_base', 'meaning', 'reference', 'n_lines_EU_2-A-1', 'n_lines_MERCOSUR_2-A-2'])
    for row in leg:
        code = row[0]
        me_code = {'FP30 %': 'FP 30 %', 'FP50 %': 'FP 50 %'}.get(code, code)
        w.writerow(list(row) + [ceu.get(code, 0), cme.get(me_code, 0)])
known = set(r[0] for r in leg)
print('EU codes not in legend:', set(ceu) - known)
print('MERCOSUR codes not in legend:', set(cme) - known - {'FP 30 %', 'FP 50 %'})

# TRQ summary (volumes verified against the OJ text of Annex 2-A Sections A-C)
trq = [
 # opening_party, notation, product, in_quota_rate, year0_qty, final_qty, final_year, unit, eligible_origin, reference
 ('EU', 'BF1', 'fresh/chilled beef', '7.5%', 9075, 54450, 5, 't cwe', 'all MERCOSUR', 'Sec B para 1'),
 ('EU', 'BF2', 'frozen beef incl. for processing', '7.5%', 7425, 44550, 5, 't cwe', 'all MERCOSUR', 'Sec B para 3'),
 ('EU', '(existing WTO Hilton-type quotas 09.4450/4452/4453/4455)', 'high-quality beef', '0% (in-quota duty 20% eliminated at EIF)', '', '', 0, '', 'all MERCOSUR', 'Sec B para 2'),
 ('EU', 'PK', 'pigmeat', '83 EUR/t', 4167, 25000, 5, 't cwe', 'all MERCOSUR (+1,500 t duty-free for Paraguay only)', 'Sec B para 4'),
 ('EU', 'PY1', 'boneless poultry and preparations', '0%', 15000, 90000, 5, 't cwe', 'all MERCOSUR', 'Sec B para 5'),
 ('EU', 'PY2', 'bone-in poultry', '0%', 15000, 90000, 5, 't cwe', 'all MERCOSUR', 'Sec B para 6'),
 ('EU', 'MP', 'milk powders', 'preference 10% of base rising to 100% (year 10)', 1000, 10000, 10, 't', 'all MERCOSUR', 'Sec B para 7'),
 ('EU', 'CE', 'cheese', 'preference 10% rising to 100% (year 10)', 3000, 30000, 10, 't', 'all MERCOSUR', 'Sec B para 8'),
 ('EU', 'IF', 'infant formula', 'preference 10% rising to 100% (year 10)', 500, 5000, 10, 't', 'all MERCOSUR', 'Sec B para 9'),
 ('EU', 'ME', 'maize and sorghum', '0%', 166667, 1000000, 5, 't', 'all MERCOSUR', 'Sec B para 10'),
 ('EU', 'RE', 'rice', '0%', 10000, 60000, 5, 't', 'all MERCOSUR', 'Sec B para 11'),
 ('EU', 'SR', 'raw cane sugar for refining', '0% (inside EU WTO quota 09.4318, Brazil)', 180000, 180000, 0, 't', 'Brazil 180,000 t; Paraguay 10,000 t duty-free; AR/UY base rate', 'Sec B para 12'),
 ('EU', 'OS', 'other sugars', '50% preference on base rate', 2000, 2000, 0, 't', 'all MERCOSUR', 'Sec B para 13'),
 ('EU', 'EG1', 'egg products', '0%', 500, 3000, 5, 't shell-egg equivalent', 'all MERCOSUR', 'Sec B para 14'),
 ('EU', 'EG2', 'egg albumins', '0%', 500, 3000, 5, 't shell-egg equivalent', 'all MERCOSUR', 'Sec B para 15'),
 ('EU', 'HY', 'honey', '0%', 7500, 45000, 5, 't', 'all MERCOSUR', 'Sec B para 16'),
 ('EU', 'RM', 'rum / cane spirits', '0%', 400, 2400, 5, 't pure alcohol', 'all MERCOSUR', 'Sec B para 17'),
 ('EU', 'SC', 'sweetcorn', '0%', 1000, 1000, 0, 't', 'all MERCOSUR', 'Sec B para 18'),
 ('EU', 'SH1', 'maize and manioc starch', '50% of base rate', 1500, 1500, 0, 't', 'all MERCOSUR', 'Sec B para 19'),
 ('EU', 'SH2', 'starch derivatives', '0%', 100, 600, 5, 't', 'all MERCOSUR', 'Sec B para 20'),
 ('EU', 'EL', 'ethanol - all uses', '6.4 EUR/hl undenatured; 3.4 EUR/hl denatured', 33333, 200000, 5, 't', 'all MERCOSUR', 'Sec B para 21'),
 ('EU', 'EL', 'ethanol - chemical industry use', '0%', 75000, 450000, 5, 't', 'all MERCOSUR', 'Sec B para 21'),
 ('EU', 'GC', 'garlic', 'preference 30% rising to 100% (year 7)', 1875, 15000, 7, 't', 'all MERCOSUR', 'Sec B para 22'),
 ('EU', 'BD (not shown in staging column; lines 38260010/90 are staged 10)', 'biodiesel', '0%', 50000, 50000, 0, 't', 'Paraguay only; others follow staging 10', 'Sec B para 23'),
 ('MERCOSUR', 'TRQ-1', 'milk powders', 'preference 10% rising to 100% (year 10)', 1000, 10000, 10, 't', 'EU', 'Sec C para 1'),
 ('MERCOSUR', 'TRQ-2', 'cheese', 'preference 10% rising to 100% (year 10)', 3000, 30000, 10, 't', 'EU', 'Sec C para 2'),
 ('MERCOSUR', 'TRQ-3', 'infant formula', 'preference 10% rising to 100% (year 10)', 500, 5000, 10, 't', 'EU', 'Sec C para 3'),
 ('MERCOSUR', 'TRQ-4', 'garlic', 'preference 30% rising to 100% (year 7)', 1875, 15000, 7, 't', 'EU', 'Sec C para 4'),
 ('MERCOSUR', 'CH1', 'chocolate NCM 1806.20', '16.2% falling to 1.8% (yr 8); free yr 9', 1710, 4760, 8, 't', 'EU', 'Sec A para 6(l)'),
 ('MERCOSUR', 'CH1', 'chocolate NCM 1806.90', '18.0% falling to 2.0% (yr 8); free yr 9', 6320, 17640, 8, 't', 'EU', 'Sec A para 6(l)'),
 ('MERCOSUR', 'CH2', 'white chocolate NCM 1704.90.10', '18.7% falling to 1.3% (yr 13); free yr 14', 771, 2030, 13, 't', 'EU', 'Sec A para 6(m)'),
 ('MERCOSUR', 'CH2', 'cocoa powder NCM 1806.10', '16.8% falling to 1.2% (yr 13); free yr 14', 90, 150, 13, 't', 'EU', 'Sec A para 6(m)'),
 ('MERCOSUR', 'CH2', 'chocolate NCM 1806.31', '18.7% falling to 1.3% (yr 13); free yr 14', 1890, 4380, 13, 't', 'EU', 'Sec A para 6(m)'),
 ('MERCOSUR', 'CH2', 'chocolate NCM 1806.32', '18.7% falling to 1.3% (yr 13); free yr 14', 1800, 5200, 13, 't', 'EU', 'Sec A para 6(m)'),
 ('MERCOSUR', 'T1', 'prepared tomatoes NCM 2002.10', '12.6% falling to 1.4% (yr 8); free yr 9', 7500, 7500, 8, 't', 'EU', 'Sec A para 6(n)'),
 ('MERCOSUR', '15V', 'vehicles (8701/8703/8704)', '50% of base rate to end of year 8', 50000, 50000, 8, 'units', 'EU (allocated by importing state AR 15,500 / BR 32,000 / PY 750 / UY 1,750)', 'Sec A para 6(h)'),
]
with open('trq_summary.csv', 'w', newline='', encoding='utf-8') as f:
    w = csv.writer(f)
    w.writerow(['opening_party', 'notation', 'product', 'in_quota_rate', 'year0_quantity', 'final_quantity', 'final_quantity_reached_in_year', 'unit', 'eligible_origin_or_allocation', 'reference_in_Annex_2-A'])
    w.writerows(trq)
print('wrote', len(leg), 'legend rows and', len(trq), 'TRQ rows')
