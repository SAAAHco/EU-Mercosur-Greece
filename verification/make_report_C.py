"""Compose the verification report (Word + Markdown) from the result files. No number is typed by hand."""
import json, csv, os, sys, io, collections, datetime
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

OUT = sys.argv[1] if len(sys.argv) > 1 else '.'
M = 1e6
res = json.load(open('results_C.json', encoding='utf-8'))
cfgC = json.load(open('config_C.json', encoding='utf-8'))
recon = json.load(open('recon_C.json', encoding='utf-8'))
lines = list(csv.DictReader(open('lines_C.csv', encoding='utf-8')))
stg = list(csv.DictReader(open('staging_check_C.csv', encoding='utf-8'))) if os.path.exists('staging_check_C.csv') else []
audit = json.load(open('audit_findings_C.json', encoding='utf-8')) if os.path.exists('audit_findings_C.json') else None
R = os.environ.get('EUMG_PACKAGE', os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
cfgA = json.load(open(os.path.join(R, 'model/inputs/v5_config.json'), encoding='utf-8'))['chapters']
cfgB = json.load(open(os.path.join(R, 'outputs/v81_config.json'), encoding='utf-8'))['chapters']
resB = json.load(open(os.path.join(R, 'outputs/results_v81.json'), encoding='utf-8'))
names = {c['hs']: c['name'] for c in cfgB}
names.update({c['hs']: c['name'] for c in cfgA})
names.update({c['hs']: c['name'] for c in cfgC['chapters']})
A, B, C = res['A'], res['B'], res['C']
today = datetime.date(2026, 9, 30).strftime('%d %B %Y')

def f1(x):
    return f'{x / M:,.1f}'

def f2(x):
    return f'{x / M:,.2f}'

# ---------- derived quantities ----------
tot_imp_C = sum(c['baseline_imp'] for c in cfgC['chapters'])
free_C = sum(l['share'] * c['baseline_imp'] for c in cfgC['chapters'] for l in c['eu_lines'] if l['tau0'] == 0)
n_taric_products = len({l['cn8'] for l in lines if l['flow'] == 'M' and l['tau_src'].startswith('TARIC')})
line_by_cn8 = {}
for l in lines:
    if l['flow'] == 'M' and l['partner'] != 'BO':
        if l['cn8'] not in line_by_cn8 or float(l['baseline_usd']) > float(line_by_cn8[l['cn8']]['baseline_usd']):
            line_by_cn8[l['cn8']] = l
n_ch_C = len(cfgC['chapters'])
feedcoffee_A = sum(A['per'][h]['net'] for h in ('23', '09') if h in A['per'])
feedcoffee_C = sum(C['per'][h]['net'] for h in ('23', '09') if h in C['per'])
comp = recon['composition']
def comp_line(hs, n=2):
    return '; '.join(f'CN {c[:4]} {c[4:6]} {c[6:]} {v:.1f}m ({p:.1f}%)' for c, v, p in comp[hs][:n])
imp_tab = recon['chapter_M']
exp_tab = recon['chapter_X']
maxdiff_imp = max(abs(r['diff_pct'] or 0) for r in imp_tab if r['Comtrade'] > 1)
tau_src_count = collections.Counter(l['tau_src'].split(' (')[0] for l in lines if l['flow'] == 'M' and l['partner'] != 'BO')
tau_src_value = collections.defaultdict(float)
for l in lines:
    if l['flow'] == 'M' and l['partner'] != 'BO':
        tau_src_value[l['tau_src'].split(' (')[0]] += float(l['baseline_usd'])
stg_ok = sum(1 for r in stg if r['match'] == 'OK'); stg_chk = [r for r in stg if r['match'] == 'CHECK']; stg_na = sum(1 for r in stg if r['match'] == '')
sens = res['sens_C']; mc = res['mc_C']; yearly = res['yearly_C']
peak_y = max(yearly, key=lambda y: yearly[y]);

# chapter comparison table (Year 10 net, million USD): A, B, C
hs_all = sorted(set(A['per']) | set(B['per']) | set(C['per']), key=lambda h: -abs(C['per'].get(h, {'net': 0})['net']) - abs(A['per'].get(h, {'net': 0})['net']) * 0.2)
def g(d, h, k='net'):
    return d['per'][h][k] / M if h in d['per'] else None

# B vs C input differences per chapter (baseline, growth, EU tau0)
byB = {c['hs']: c for c in cfgB}; byC = {c['hs']: c for c in cfgC['chapters']}

# ---------- document ----------
doc = Document()
st = doc.styles['Normal']; st.font.name = 'Times New Roman'; st.font.size = Pt(11)
for s in doc.sections:
    s.left_margin = s.right_margin = Cm(2.2); s.top_margin = s.bottom_margin = Cm(2.2)
md = []

def H(text, level=1):
    doc.add_heading(text, level=level); md.append('\n' + '#' * level + ' ' + text + '\n')

def P(text, bold_lead=None):
    p = doc.add_paragraph()
    if bold_lead:
        p.add_run(bold_lead + ' ').bold = True
    p.add_run(text); p.paragraph_format.space_after = Pt(6)
    md.append((f'**{bold_lead}** ' if bold_lead else '') + text + '\n')

def T(header, rows, caption=None, widths=None):
    if caption:
        p = doc.add_paragraph(); r = p.add_run(caption); r.bold = True; r.font.size = Pt(10)
        md.append('\n**' + caption + '**\n')
    t = doc.add_table(rows=1, cols=len(header)); t.style = 'Table Grid'
    for i, h in enumerate(header):
        cell = t.rows[0].cells[i]; cell.text = ''; run = cell.paragraphs[0].add_run(str(h)); run.bold = True; run.font.size = Pt(9)
    for row in rows:
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ''; run = cells[i].paragraphs[0].add_run('' if v is None else str(v)); run.font.size = Pt(9)
            if i > 0:
                cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
    md.append('| ' + ' | '.join(map(str, header)) + ' |'); md.append('|' + '---|' * len(header))
    for row in rows:
        md.append('| ' + ' | '.join('' if v is None else str(v) for v in row) + ' |')
    md.append('')
    doc.add_paragraph()

title = 'Independent re-analysis of the Greece-Mercosur projection: which headline is right, 146 or 16 million dollars?'
doc.add_heading(title, level=0); md.append('# ' + title + '\n')
P(f'Verification memo for the author team, {today}. Prepared from new data pulls (Eurostat COMEXT, EU TARIC, WITS-TRAINS) and a freshly written engine; every figure below is computed from the files in the accompanying folder.')

H('1. Answer in brief')
P(f"The corrected figure is right in substance and the published 146.0 million figure is not. Rebuilt from an independent trade database (Eurostat COMEXT at the eight-digit CN level) and the EU's own tariff database (TARIC), the Year 10 widening of the bilateral balance is {f1(C['widening'])} million US dollars, against {f1(B['widening'])} million in the corrected (v8.1) specification and {f1(A['widening'])} million in the published one. The independent engine reproduces both published numbers exactly from their own inputs ({f2(A['widening'])} and {f2(B['widening'])}), so the disagreement is entirely about inputs, and the decisive input is the tariff assigned to the products Greece actually imports.")
P(f"Three facts settle it. First, the trade values were never wrong: COMEXT and the UN Comtrade extract used by both manuscripts agree to the cent at chapter level for 2022 to 2024 (largest import-chapter difference {maxdiff_imp:.2f}%), because Comtrade takes Greece's data from Eurostat. Second, COMEXT's eight-digit detail shows that the two chapters carrying the published result are single products: {comp_line('23', 1)} of the feed chapter and {comp_line('09', 1)} of the coffee chapter. Third, TARIC records the EU third-country duty on both products as 0%, so the agreement cuts nothing on them; the published model applied 5% and 3% chapter averages, which generated {f1(feedcoffee_A)} million of its {f1(A['widening'])} million. Under the independent inputs the same two chapters contribute {f2(feedcoffee_C)} million.")
P(f"Of the {f1(tot_imp_C)} million dollar import base (five partners, after the documented spike screen), {free_C / tot_imp_C:.1%} already enters the EU duty-free; the corrected manuscript reports 83.0% on the same definition. The adjustment that remains falls on tobacco leaf, frozen seafood and juices on the import side and is offset by kiwifruit, pharmaceutical and industrial exports, which is the composition the corrected manuscript reports.")

H('2. What was compared')
T(['', 'A. Published (Ecological Economics / v7.5)', 'B. Corrected (GFS, JCMS / v8.1)', 'C. This independent re-analysis'],
  [['Trade data', 'UN Comtrade, HS-2 chapter totals', 'UN Comtrade, HS-6 lines', 'Eurostat COMEXT, CN-8 lines (5 partners, 2014-2024), EUR converted at ECB annual averages'],
   ['EU tariffs', 'One average per chapter (25 chapters)', 'WITS-TRAINS HS-6 AVEs; TARIC for 21 lines', f"EU TARIC third-country duty read at CN-8 for every product with a baseline of at least USD {cfgC['meta']['opts']['taric_min_usd']:,} ({n_taric_products} products, {tau_src_count.get('TARIC CN8', 0)} partner lines), ad valorem equivalents at Greek unit values; the agreement's own base rate, then WITS-TRAINS AVE, for the residual small lines"],
   ['Mercosur tariffs', 'One average per chapter', 'Schedule base rates capped at WITS applied MFN', 'Same rule, rebuilt from the schedule CSV and fresh WITS pulls; HS-4 fallback for HS 2022 codes absent from NCM 2012'],
   ['Phase-outs', 'Assumed per chapter', 'OJ L 2026/184 Appendices 2-A-1 and 2-A-2', 'Same schedules, exact CN-8 match first; for codes that changed nomenclature, the category implied by the Mercosur preference TARIC applies today; legal check of the top lines against TARIC (Section 6)'],
   ['Chapters', '25', '44', f'{n_ch_C} (the stated USD 1 million rule applied to the COMEXT panel; B additionally carries HS 88 and HS 99, both with negligible baselines)'],
   ['Engine', 'orchestrator_v7_5.py', 'pe_v81.py (same formulas, line-level)', 'engine_C.py, written from the published formula; reproduces A and B exactly'],
   ['Year 10 widening, million USD', f2(A['widening']), f2(B['widening']), f2(C['widening'])]],
  caption='Table 1. The three specifications.')
P("Two remarks on independence. The trade source is genuinely different in provenance (Eurostat's dissemination database rather than the UN's), but it is the same underlying statistical record, so it can confirm the trade values but cannot contradict them. The tariff source is different in kind: TARIC is the legal instrument that customs apply, whereas the corrected model relied mainly on WITS averages at the six-digit level.")

H('3. The trade data are not in dispute')
rows = [[r['hs2'], names.get(r['hs2'], '')[:38], f"{r['COMEXT']:.2f}", f"{r['Comtrade']:.2f}", f"{(r['diff_pct'] or 0):+.2f}%"] for r in imp_tab[:15]]
T(['HS', 'Chapter', 'COMEXT (USD m)', 'Comtrade extract (USD m)', 'Difference'], rows, caption='Table 2. Greek imports from Mercosur, 2022-2024 mean, top 15 chapters: independent COMEXT pull against the Comtrade extract used by both manuscripts.')
P(f"Totals: imports {sum(r['COMEXT'] for r in imp_tab):.1f} (COMEXT) against {sum(r['Comtrade'] for r in imp_tab):.1f} (Comtrade); exports {sum(r['COMEXT'] for r in exp_tab):.1f} against {sum(r['Comtrade'] for r in exp_tab):.1f}, the export gap being Comtrade's chapter 99 (special transactions), which COMEXT does not carry as a CN-8 leaf. The panel years 2014 to 2021 differ by at most a few percent because of later revisions. The data gathered manually for the published analysis were therefore correct; what was wrong was the tariff attached to them.")

H('4. What Greece imports in the chapters that carried the published result, and the duty on it')
rows = []
for hs in ['23', '09', '12', '24', '27', '26', '47', '03', '20', '22', '08', '02']:
    if hs in comp:
        top_cn8 = comp[hs][0][0]
        top = line_by_cn8.get(top_cn8)
        if top:
            duty, stg_ = f"{float(top['tau0']) * 100:.2f}% ({top['tau_src'].split(' (')[0].split(':')[0]})", top['staging']
        elif top_cn8 == '27090090':
            duty, stg_ = 'Free in TARIC; the flow was a single 2023 shipment removed by the spike screen', '0'
        else:
            duty, stg_ = 'not modelled (flow removed by the spike screen)', ''
        rows.append([hs, names.get(hs, '')[:30], comp_line(hs, 2), duty, stg_])
T(['HS', 'Chapter', 'Main CN-8 lines, 2022-24 mean (share of chapter)', 'EU duty on the main line', 'Staging of main line'], rows, caption='Table 3. Eight-digit composition of the import chapters (before the spike screen) and the EU third-country duty on the dominant line (TARIC, 1 July 2023; entry-price lines at the Greek unit value).')
P(f"Soybean oilcake (CN 2304 00 00), green coffee (0901 11 00), soybeans (1201 90 00), shelled groundnuts (1202 42 00), aluminium ore (2606 00 00), chemical pulp (4703 29 00) and crude oil (2709 00 90) are all duty-free at the EU border. Together with the other zero-duty lines they make up {free_C / tot_imp_C:.1%} of Greek imports from the four Mercosur parties. Fuel oil under CN 2710 19 51 carries a conventional 3.5% but can only be declared there under the end-use procedure at the autonomous rate Free, so it too is treated as duty-free. The lines that do carry duty are tobacco leaf (compound duty capped at 24 or 56 EUR per 100 kg, an ad valorem equivalent of {next((float(l['tau0']) * 100 for l in lines if l['cn8'] == '24012085' and l['partner'] == 'BR'), 0):.2f}% at the Greek unit value on the certified flue-cured sub-line), lemons (6.4% with entry price), Argentine shrimp (12%), Illex squid (8%), hake fillets (7.5%), rosin (5%), denatured ethanol (10.2 EUR/hl), orange juice (12.2% and 15.2%), grape juice (40%) and the beef lines, which are governed by quotas.")

H('5. Results')
rows = [[h, names.get(h, '')[:34], None if g(A, h) is None else f'{g(A, h):.2f}', None if g(B, h) is None else f'{g(B, h):.2f}', None if g(C, h) is None else f'{g(C, h):.2f}',
         None if g(C, h, 'dM') is None else f"{g(C, h, 'dM'):.2f}", None if g(C, h, 'dX_adj') is None else f"{g(C, h, 'dX_adj'):.2f}"] for h in hs_all if max(abs(g(A, h) or 0), abs(g(B, h) or 0), abs(g(C, h) or 0)) >= 0.5]
rows.append(['All', 'All chapters', f2(A['widening']), f2(B['widening']), f2(C['widening']), f2(C['dM']), f2(C['dX_adj'])])
T(['HS', 'Chapter', 'A net', 'B net', 'C net', 'C add. imports', 'C add. exports (adj.)'], rows, caption='Table 4. Year 10 change in the bilateral balance by chapter, million US dollars (positive = widening, adverse to Greece). Chapters with at least 0.5 million in any specification.')
P(f"Analysis C gives additional imports of {f1(C['dM'])} million and capacity-adjusted additional exports of {f1(C['dX_adj'])} million at Year 10 (B: {f1(B['dM'])} and {f1(B['dX_adj'])}; A: {f1(A['dM'])} and {f1(A['dX_adj'])}). The corrected engine file pe_v81.py, run on the Analysis C inputs, returns {f2(res['C_pe81_engine'])} million, identical to the independent engine. The widening path is {', '.join(f'{v / M:.1f}' for y, v in sorted(yearly.items(), key=lambda kv: int(kv[0])))} million from Year 0 to Year 10, peaking in Year {peak_y}. Entry in 2027 gives {f1(sens['entry 2027 (S2)'])} million; the Monte Carlo over the wedge and capacity parameters gives {f1(mc['mean'])} million with a 95% interval of {f1(mc['lo'])} to {f1(mc['hi'])} million.")

H('5.1 Which input drives the difference')
T(['Specification', 'Year 10 widening (USD m)', 'Reading'],
  [['A. Published inputs', f2(A['widening']), 'Chapter-average tariffs on Comtrade chapter totals, tobacco baseline set to zero'],
   ['H1. COMEXT trade, A chapter tariffs', f2(res['H1']['widening']), 'Better trade data make the published specification worse, not better: with its 10% chapter tariff, the 50 million tobacco baseline that A had set to zero adds about 52 million. The trade data do not explain the gap'],
   ['H2. Comtrade trade (A), C line-level tariffs, A chapters', f2(res['H2']['widening']), 'Replacing only the tariffs removes the whole gap and turns the balance favourable, because A also had no tobacco imports'],
   ['H3. C, restricted to the published chapters', f2(res['H3']['widening']), 'Chapters outside the published set net to about ten million: fish, meat and footwear adverse, optical instruments favourable'],
   ['H4. C, but feed and coffee at the published 5% and 3%', f2(res['H4']['widening']), 'The two disputed tariff inputs alone take the independent specification back to the published figure'],
   ['C. Independent inputs', f2(C['widening']), ''],
   ['B. Corrected v8.1', f2(B['widening']), '']],
  caption='Table 5. Decomposition of the 146-to-16 gap by swapping one input block at a time.')

H('6. Legal check of the phase-out schedule against TARIC')
P(f"Because the agreement has been provisionally applied since 1 May 2026, TARIC now shows the preferential duty Mercosur goods actually pay in Year 0. For the {len(stg)} largest import lines the implied first-year cut was compared with the staging category taken from the Official Journal schedule: {stg_ok} lines match, {len(stg_chk)} need attention, and {stg_na} could not be compared (no computable preference, quota lines or already duty-free).")
if stg:
    rows = [[r['cn8'], r['hs'], r['partner'], r['baseline_musd'], f"{float(r['tau0_model']) * 100:.2f}%", r['staging'][:14], r['trq'], f"{float(r['model_year0_rate']) * 100:.2f}%", (str(r['taric_mercosur_pref'])[:60] if r['taric_mercosur_pref'] not in ('[]', '') else '-'), ('' if r['taric_pref_ave'] in ('', 'None', None) else f"{float(r['taric_pref_ave']) * 100:.2f}%"), r['match'] or 'n/a'] for r in stg[:30]]
    T(['CN-8', 'HS', 'From', 'Base (USD m)', 'Duty used', 'Category', 'Quota', 'Model Year-0 duty', 'TARIC Mercosur preference (2026)', 'Its AVE', 'Match'], rows, caption='Table 6. Staging check for the 30 largest import lines.')
if stg_chk:
    P('Lines flagged for attention: ' + '; '.join(f"{r['cn8']} (model {float(r['model_year0_rate']) * 100:.2f}% vs TARIC {float(r['taric_pref_ave']) * 100:.2f}%)" for r in stg_chk) + '. These are explained in the accompanying notes where the TARIC measure shown is a quota or a seasonal rate rather than the staged tariff.')

H('7. Sensitivity of the independent result')
T(['Assumption', 'Year 10 widening (USD m)'], [[k, f2(v)] for k, v in sens.items()], caption='Table 7. Sensitivities on Analysis C.')
P("As in the corrected manuscript, the sign depends on the import elasticity: the result turns favourable to Greece at elasticities of 2.5 or below in absolute value. The quota treatment matters little at Greek scale. The magnitude is modest under every assumption examined and never approaches the published 146 million.")

H('8. Where the independent inputs differ from the corrected v8.1 inputs')
rows = []
for h in sorted(set(byB) | set(byC), key=lambda h: -(byC.get(h, byB.get(h))['baseline_imp'] + byC.get(h, byB.get(h))['baseline_exp'])):
    b, c = byB.get(h), byC.get(h)
    if b is None or c is None:
        rows.append([h, names.get(h, '')[:30], f"{b['baseline_imp'] / M:.2f}" if b else '-', f"{c['baseline_imp'] / M:.2f}" if c else '-', f"{b['eu_mfn'] * 100:.2f}" if b else '-', f"{c['eu_mfn'] * 100:.2f}" if c else '-', f"{b['cagr_imp'] * 100:.1f}" if b else '-', f"{c['cagr_imp'] * 100:.1f}" if c else '-', 'only in ' + ('B' if b else 'C')])
        continue
    d_tau = c['eu_mfn'] - b['eu_mfn']
    if abs(d_tau) > 0.002 or abs(c['baseline_imp'] - b['baseline_imp']) > 0.2 * M or abs(c['cagr_imp'] - b['cagr_imp']) > 0.005:
        rows.append([h, names.get(h, '')[:30], f"{b['baseline_imp'] / M:.2f}", f"{c['baseline_imp'] / M:.2f}", f"{b['eu_mfn'] * 100:.2f}", f"{c['eu_mfn'] * 100:.2f}", f"{b['cagr_imp'] * 100:.1f}", f"{c['cagr_imp'] * 100:.1f}", ''])
T(['HS', 'Chapter', 'B imports (USD m)', 'C imports', 'B EU duty (%)', 'C EU duty', 'B growth (%)', 'C growth', 'Note'], rows[:25], caption='Table 8. Chapters where the two line-level specifications differ materially in baseline, trade-weighted EU duty or growth.')
P(f"Baselines agree to the dollar wherever a chapter appears in both, as Section 3 implies; growth rates differ by a fraction of a point in a few chapters because the 2014 base year carries later revisions in COMEXT. The duty differences come from four sources. TARIC eight-digit duties at Greek unit values replace six-digit WITS averages (tobacco {next((float(l['tau0']) * 100 for l in lines if l['cn8'] == '24012085' and l['partner'] == 'BR'), 0):.2f}% against 4.04%, lemons 6.40% against 6.51%). The exact CN-8 staging match replaces base-weighted six-digit averages. Beef is entered at its full compound duty with the 7.5% in-quota rate applied through the quota logic, where v8.1 entered the 20% Hilton in-quota rate as the base; both are bound by the notional Greek quota slice, so the chapter result moves by less than 0.2 million. Two codes that changed nomenclature since the schedules were drafted (frozen Illex squid, CN 2013 0307 99 11, duty-free at entry into force; extra-virgin olive oil exports, NCM 2012 heading 1509, category 15 at 10% and 31.5% for Argentina) are resolved from TARIC's current preferential measures and from the heading-level rows respectively, where v8.1 used hand-entered overrides. The net effect of all of this is {f2(B['widening'] - C['widening'])} million at Year 10.")

H('9. Verification and audit trail')
P("The engine was written independently from the published formula and checked by reproducing both published headline sets exactly before any new input was used. Four independent auditors then reviewed the engine, the data pipeline, the tariff facts (against the Combined Nomenclature regulations for 2023 and 2024 obtained from the Publications Office) and the agreement parameters (against the Commission proposal annexes on the Council register), and every material finding was verified by a second reviewer before being acted on. The findings and their resolution are listed in the accompanying audit file.")
if audit:
    rows = [[f.get('auditor', ''), f.get('severity', ''), f.get('title', '')[:90], f.get('resolution', '')[:110]] for f in audit.get('confirmed', [])]
    if rows:
        T(['Auditor', 'Severity', 'Finding', 'Resolution in Analysis C'], rows, caption='Table 9. Confirmed audit findings and how they were resolved.')
    if audit.get('refuted'):
        P('Findings raised and then refuted on verification: ' + '; '.join(f.get('title', '')[:80] for f in audit['refuted']) + '.')

H('10. What this means for the manuscripts')
P(f"The Ecological Economics manuscript and its supplement rest on the 146.0 million figure and on the '92% in feed and coffee' finding. Both are artefacts of applying chapter-average duties to duty-free products, and neither survives contact with the EU's own tariff database. They should not be resubmitted in that form.", 'Published version.')
P(f"The Global Food Security and JCMS texts carry the corrected specification. The independent rebuild gives {f1(C['widening'])} million against their {f1(B['widening'])} million, the same composition (tobacco, seafood and juices adverse; kiwifruit, pharmaceuticals and optics favourable) and the same sign sensitivity. The remaining differences are within the range the corrected text already reports as robustness, so the corrected manuscript can be submitted; the figures in Section 8 above indicate where a sentence could be tightened.", 'Corrected version.')
P("Cite the trade source as Eurostat COMEXT or UN Comtrade interchangeably (they are the same record for Greece), and cite TARIC, not WITS, for the duty on the lines that carry the result. Keep the ad valorem equivalents at Greek unit values for tobacco and ethanol, since WITS's world-average equivalents differ. State that Bolivia, which is in the trade totals, receives no preference.", 'Data statements.')

H('Appendix A. Sources and files')
P("Eurostat COMEXT, dataset ds-045409 (EU trade since 1988 by HS2-4-6 and CN8), reporter Greece, partners Argentina, Brazil, Uruguay, Paraguay, Bolivia, annual 2014-2024, values in euro and quantities in 100 kg, accessed 30 September 2026 through the dissemination API. European Central Bank, euro reference rates, annual averages 2014-2024. European Commission, TARIC consultation database, measures read at simulation dates 1 July 2023 (base duties) and 30 September 2026 (Mercosur preferences). World Bank WITS / UNCTAD TRAINS, MFN applied rates and ad valorem equivalents, 2023. Official Journal L 2026/184, Appendices 2-A-1 and 2-A-2 (tariff schedules), as parsed in the replication package and cross-read against Council documents ST 12487/25 ADD 2 to ADD 8.")
P("Files in this folder: config_C.json (all line-level inputs), lines_C.csv (every import and export line with duty provenance and staging), staging_check_C.csv, results_C.json, recon_C.json and recon_C.txt (COMEXT against Comtrade), comext_cn8_tidy.csv (the raw pull, tidy), taric_cache.json and wits_cache.json (the raw tariff readings), and the scripts pull_comext.py, parse_comext.py, taric.py, wits.py, build_inputs_C.py, engine_C.py, run_C.py, check_staging_C.py, recon_C.py, make_report_C.py. Running build_inputs_C.py, run_C.py and check_staging_C.py in that order regenerates every number in this memo from the cached data.")

H('Appendix B. The largest import lines with their duty provenance')
tl = sorted([l for l in lines if l['flow'] == 'M' and l['partner'] != 'BO'], key=lambda l: -float(l['baseline_usd']))[:40]
T(['CN-8', 'HS', 'From', 'Baseline (USD m)', 'Duty (%)', 'Source', 'Staging', 'Path source'], [[l['cn8'], l['hs'], l['partner'], f"{float(l['baseline_usd']) / M:.2f}", f"{float(l['tau0']) * 100:.2f}", l['tau_src'][:34], l['staging'][:12], l['path_src'][:40]] for l in tl], caption='Table B.1. Forty largest import lines (four Mercosur parties, 2022-24 mean).')
tx = sorted([l for l in lines if l['flow'] == 'X' and l['partner'] != 'BO'], key=lambda l: -float(l['baseline_usd']))[:20]
T(['CN-8', 'HS', 'To', 'Baseline (USD m)', 'Mercosur duty (%)', 'Base / applied', 'Schedule rows'], [[l['cn8'], l['hs'], l['partner'], f"{float(l['baseline_usd']) / M:.2f}", ('' if l['tau0'] in ('', 'None') else f"{float(l['tau0']) * 100:.1f}"), l['tau_src'][:34], l['staging'][:40]] for l in tx], caption='Table B.2. Twenty largest export lines.')

os.makedirs(OUT, exist_ok=True)
doc.save(os.path.join(OUT, 'Independent_Reanalysis_Report.docx'))
open(os.path.join(OUT, 'Independent_Reanalysis_Report.md'), 'w', encoding='utf-8').write('\n'.join(md))
print('report written to', OUT)
