# -*- coding: utf-8 -*-
"""
make_appendix_figs_v8.py
========================
Regenerates the supplementary figures under the CORRECTED tariff-response
specification (see make_figures_v8.py for the diagnosis of the superseded
term).

Covers S1-S18 except S12, which draws on external FADN and Eurostat series
not held in the analytical workbook.

PALETTE
-------
Okabe-Ito, colour-blind safe. No greys anywhere: every plotted element is a
saturated hue or near-black, so nothing washes out in print, in greyscale
photocopy, or on projection.

Outputs PNG (200 dpi) and EPS (vector, fonttype 42) for each figure.

USAGE
-----
    python3 make_appendix_figs_v8.py [--engine PATH] [--out DIR]
"""
import os, sys, argparse, importlib.util, random, statistics
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Rectangle

ap = argparse.ArgumentParser()
ap.add_argument('--engine', default='/home/claude/work/make_figures_v8.py')
ap.add_argument('--out', default='/home/claude/work/figs_v8/appendix')
A = ap.parse_args()
os.makedirs(A.out, exist_ok=True)

_argv = sys.argv
sys.argv = ['engine', '--out', '/tmp/_engine_junk']
spec = importlib.util.spec_from_file_location('engine', A.engine)
E = importlib.util.module_from_spec(spec)
spec.loader.exec_module(E)
sys.argv = _argv
CH, project, WEDGE = E.CH, E.project, E.WEDGE

C = dict(blue='#0072B2', orange='#E69F00', green='#009E73', verm='#D55E00',
         sky='#56B4E9', purple='#AA4499', navy='#1F3B73', teal='#117777',
         gold='#B8860B', ink='#111111')
plt.rcParams.update(matplotlib.rcParamsDefault)
plt.rcParams.update({
    'font.family': 'DejaVu Sans', 'font.size': 8.5, 'axes.titlesize': 9,
    'axes.labelsize': 8.5, 'axes.edgecolor': C['ink'], 'axes.linewidth': 0.8,
    'axes.spines.top': False, 'axes.spines.right': False,
    'xtick.color': C['ink'], 'ytick.color': C['ink'],
    'grid.color': '#9FB6CD', 'grid.linewidth': 0.6,
    'figure.dpi': 200, 'savefig.dpi': 200, 'savefig.facecolor': 'white',
    'legend.frameon': False, 'legend.fontsize': 7.6,
    'ps.fonttype': 42, 'pdf.fonttype': 42,
})

def save(fig, name):
    fig.savefig(f'{A.out}/{name}.png', bbox_inches='tight')
    fig.savefig(f'{A.out}/{name}.eps', bbox_inches='tight', format='eps')
    plt.close(fig)
    print(f'  {name}.png + {name}.eps')

def panel(ax, letter):
    ax.set_title(letter, loc='left', fontweight='bold', fontsize=10)

def box(ax, x, y, w, h, text, fc='#EAF1F8', ec=None, fs=7.3, lw=1.0):
    ec = ec or C['navy']
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle='round,pad=0.012',
                                fc=fc, ec=ec, lw=lw, mutation_aspect=1.4))
    ln = text.split('\n')
    ax.text(x + w/2, y + h*0.74, ln[0], ha='center', va='center',
            fontsize=fs+0.5, fontweight='bold')
    if len(ln) > 1:
        ax.text(x + w/2, y + h*0.33, '\n'.join(ln[1:]), ha='center', va='center', fontsize=fs-0.4)

def arrow(ax, x1, y1, x2, y2, col=None):
    ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle='-|>',
                                 mutation_scale=9, lw=1.0, color=col or C['navy']))

def totals(y, **kw):
    d = project(CH, y, **kw)
    i = sum(v[0] for v in d.values()); x = sum(v[1] for v in d.values())
    return i/1e6, x/1e6, (i-x)/1e6

I10, X10, W10 = totals(10)
NET = {h: (v[0]-v[1])/1e6 for h, v in project(CH, 10).items()}

# ================= S1  analytical pipeline =================
fig, ax = plt.subplots(figsize=(7.3, 4.3)); ax.axis('off')
ax.set_xlim(0, 10); ax.set_ylim(0, 10)
ins = [('Trade panel\nUN Comtrade 2014-2024\nCOMEXT CN-8 check', 8.45),
       ('Agreement annexes\nMFN, phase-ins,\nTRQs, 357 GIs', 6.35),
       ('Greek structure\nELSTAT Census 2020\nFADN 2019-2023', 4.25),
       ('Cost anchors\nEUDR, CBAM, GAP,\nenv-tax (signed)', 2.0)]
for t, y in ins: box(ax, 0.15, y, 2.5, 1.55, t)
box(ax, 3.5, 5.6, 3.0, 2.6, 'Partial-equilibrium engine\nprice-change tariff term\n'
    'incidence-signed wedges\nHS 04 TRQ cap\ncapacity ladder T1-T3', fc='#DCE9F5')
box(ax, 3.5, 1.7, 3.0, 1.9, 'XGBoost companion\ntrain 2014-2023\n2024 holdout',
    fc='#FCEBD5', ec=C['orange'])
outs = [('Policy effect\nScenarios S1 and S2', 8.3, '#EAF1F8', 1.65, C['navy']),
        ('Robustness\ngrid, Monte Carlo', 6.0, '#EAF1F8', 1.65, C['navy']),
        ('Counterfactual baseline\nvalidated out of sample', 3.9, '#FCEBD5', 1.65, C['orange']),
        ('Specification uncertainty\ncarried by the import\nelasticity, not by the\nML companion', 1.0, '#E2F3EA', 2.1, C['green'])]
for t, y, fc, h, ec in outs: box(ax, 7.25, y, 2.6, h, t, fc=fc, ec=ec)
for _, y in ins:
    arrow(ax, 2.68, y+0.75, 3.47, 6.9); arrow(ax, 2.68, y+0.75, 3.47, 2.6, C['orange'])
arrow(ax, 6.53, 7.4, 7.22, 9.1); arrow(ax, 6.53, 6.6, 7.22, 6.8)
arrow(ax, 6.53, 2.6, 7.22, 4.7, C['orange']); arrow(ax, 8.55, 5.97, 8.55, 3.15)
ax.text(8.55, 0.55, 'no tariff variation in sample:\nML cannot identify the policy effect',
        ha='center', fontsize=7.0, color=C['verm'], style='italic')
save(fig, 's1')

# ================= S2  phase-out schedules =================
fig, (a1, a2) = plt.subplots(1, 2, figsize=(7.3, 2.9), sharey=True)
eug, merg = {}, {}
for c in CH:
    if c['t0'] > 0: eug.setdefault(round(c['phi']), []).append(c['hs'])
    if c['m0'] > 0: merg.setdefault(round(c['phx']), []).append(c['hs'])
pal = [C['sky'], C['green'], C['ink'], C['orange'], C['purple'], C['teal']]
for i, (ph, hs) in enumerate(sorted(eug.items())):
    yy = np.linspace(0, 10, 200)
    a1.plot(yy, [100*max(0, 1-y/ph) for y in yy], lw=1.9, color=pal[i % len(pal)],
            label=f'{ph} y: HS {", ".join(sorted(hs)[:4])}')
for i, (ph, hs) in enumerate(sorted(merg.items())):
    yy = np.linspace(0, 15, 220)
    a2.plot(yy, [100*max(0, 1-y/ph) for y in yy], lw=1.9, color=pal[i % len(pal)],
            label=f'{ph} y: HS {", ".join(sorted(hs)[:4])}')
a2.axvspan(10, 15, color=C['orange'], alpha=.16)
a2.text(12.4, 62, 'beyond model\nhorizon', fontsize=7.2, ha='center', color=C['gold'])
a1.set_title('EU-side phase-out on Mercosur imports', fontsize=8.4)
a2.set_title('Mercosur-side phase-out on Greek exports', fontsize=8.4)
for a in (a1, a2):
    a.set_xlabel('Years after entry into force'); a.legend(fontsize=6.5); a.grid(alpha=.35)
a1.set_ylabel('Applied tariff, % of initial MFN')
panel(a1, 'a'); panel(a2, 'b')
fig.tight_layout(); save(fig, 's2')

# ================= S3  wedge composition =================
items = sorted(((v, k) for k, v in WEDGE.items()))
labs = [f'HS {h}' for _, h in items]; vals = [v*100 for v, _ in items]
cols = [C['verm'] if h in ('73','76','68','38','25','27') else
        (C['green'] if h in ('23','12','09','47') else C['blue']) for _, h in items]
fig, ax = plt.subplots(figsize=(7.0, 3.6))
ax.barh(labs, vals, color=cols, height=.65)
for i, v in enumerate(vals):
    ax.text(v - .10, i, f'{v:+.1f}', va='center', ha='right', fontsize=7.4,
            fontweight='bold', color=C['ink'])
ax.axvline(0, color=C['ink'], lw=1.0)
ax.set_xlim(-3.75, 0.25)
ax.set_xlabel('signed wedge, percentage points of per-unit supply cost (negative dampens imports)')
ax.set_title('green: EUDR    red: CBAM    blue: good agricultural practice\n'
             'HS 23 nets the \u22121.0 pt EUDR channel against the +0.5 pt Greek env-tax addon',
             fontsize=7.4, loc='left', pad=8)
ax.grid(axis='x', alpha=.35)
save(fig, 's3')

# ================= S4  capacity weights =================
cw = sorted(((c['capw'], c['hs']) for c in CH if c['capw'] > 0), reverse=True)
anch = {'04':'10.87 M sheep and goats','08':'6.2% of UAA','15':'olives for oil, 18.5% of UAA',
        '20':'table olives, 2.2% of UAA','24':'full CAP support','22':'non-table vineyards, 1.2% of UAA',
        '23':'cereals 20.8% of UAA, partial','12':'limited Greek oilseed base','09':'aromatic herbs, no census table'}
fig, ax = plt.subplots(figsize=(7.2, 3.4))
ys = np.arange(len(cw))[::-1]
ax.hlines(ys, 0, [v for v, _ in cw], color=C['teal'], lw=3.6)
ax.plot([v for v, _ in cw], ys, 'o', ms=7.5, color=C['teal'])
ax.set_yticks(ys); ax.set_yticklabels([f'{h}' for _, h in cw])
for y, (v, h) in zip(ys, cw):
    ax.text(1.07, y, f'{v:.2f}', fontsize=7.8, fontweight='bold', va='center')
    ax.text(1.20, y, anch.get(h, ''), fontsize=7.0, va='center', color=C['navy'])
ax.set_xlim(0, 2.15); ax.set_xticks([0, .25, .5, .75, 1.0])
ax.set_xlabel('capacity weight \u03ba(c)')
ax.text(1.07, len(cw)-0.35, '\u03ba       census validation anchor', fontsize=7.4,
        fontweight='bold', color=C['navy'])
ax.grid(axis='x', alpha=.35)
save(fig, 's4')

# ================= S5  stochastic priors =================
fig, axs = plt.subplots(1, 4, figsize=(7.3, 2.4))
tri = [('Wedge multiplier', 0.5, 1.0, 3.0, C['blue']),
       ('Farm-to-Fork drag', 0.05, 0.10, 0.20, C['green']),
       ('Labor drag', 0.03, 0.07, 0.12, C['orange'])]
for ax, (nm, lo, mo, hi, col) in zip(axs[:3], tri):
    xs = np.linspace(lo, hi, 300)
    pdf = np.where(xs < mo, 2*(xs-lo)/((hi-lo)*(mo-lo)), 2*(hi-xs)/((hi-lo)*(hi-mo)))
    ax.fill_between(xs, pdf, color=col, alpha=.50)
    ax.plot(xs, pdf, color=col, lw=1.7)
    ax.axvline(mo, color=C['ink'], ls='--', lw=1.0)
    ax.set_title(nm, fontsize=8.0); ax.set_xticks([lo, mo, hi]); ax.set_yticks([])
    ax.text(.5, -.40, f'Tri({lo}, {mo}, {hi})', transform=ax.transAxes, ha='center', fontsize=7.0)
axs[0].set_ylabel('probability density')
a4 = axs[3]
a4.barh([1, 0], [2, 0.4], left=[1, 0.8], color=[C['purple'], C['teal']], height=.42)
a4.set_yticks([1, 0]); a4.set_yticklabels(['GDP growth %', 'phase speed'], fontsize=7.4)
a4.set_xlim(0, 3.4); a4.set_title('ML uniform priors', fontsize=8.0)
a4.text(2.0, 1.33, '1 to 3', fontsize=7.2, ha='center')
a4.text(1.0, .33, '0.8 to 1.2', fontsize=7.2, ha='center')
a4.set_xlabel('uniform draw range', fontsize=7.4)
fig.tight_layout(); save(fig, 's5')

# ================= S6  capacity ladder =================
TXU = AGRI = WTD = 0.0
for c in CH:
    if c['m0'] > 0:
        my = c['m0'] * max(0.0, 1 - min(10, c['phx'])/c['phx']) if c['phx'] else 0.0
        u = c['cx0'] * (1+c['gx'])**10 * E.EPS_EXP * ((1+my)/(1+c['m0']) - 1)
        TXU += u; WTD += c['capw']*u
        if c['capw'] > 0: AGRI += u
bench, t1, t2, t3 = TXU*0.183*0.10/1e6, AGRI*0.10/1e6, AGRI*0.17/1e6, WTD*0.17/1e6
fig, ax = plt.subplots(figsize=(7.3, 3.3))
steps = ['Economy-wide\nbenchmark', 'Tier 1\n+ F2F drag', 'Tier 2\n+ labor drag', 'Tier 3\ncensus weights']
vals = [bench, t1, t2, t3]
bars = ax.bar(steps, vals, color=[C['purple'], C['sky'], C['sky'], C['blue']], width=.6)
for b, v in zip(bars, vals):
    ax.text(b.get_x()+b.get_width()/2, v+.09, f'{v:,.2f}', ha='center', fontsize=8.4, fontweight='bold')
for i in range(3):
    ax.annotate(f'{vals[i+1]-vals[i]:+,.2f}', xy=(i+.5, max(vals[i], vals[i+1])*.55),
                ha='center', fontsize=7.9, color=C['verm'], fontweight='bold')
ax.set_ylabel('Year 10 export-capacity loss, $ million'); ax.set_ylim(0, max(vals)*1.42)
ax.text(.02, .95, f'bilateral exports are {100*AGRI/TXU:,.1f}% agricultural against an 18.3% economy-wide share;\n'
                  f'the benchmark understates the Tier 3 loss by {100*(t3-bench)/t3:,.0f}%',
        transform=ax.transAxes, fontsize=7.4, va='top')
ax.grid(axis='y', alpha=.35)
save(fig, 's6')

# ================= S7  TRQ construction =================
QUOTA = [('Cheese', 30_000, 5_000), ('Milk powder', 10_000, 3_000), ('Honey', 45_000, 4_000)]
rows = list(E.WB['4-Scenario S1 (2026)'].iter_rows(values_only=True))[5:30]
cf04 = [r[10] for r in rows if str(r[1]).zfill(2) == '04'][0] or 0
proj04 = (cf04 + project(CH, 10)['04'][0])/1e6
fig, (a1, a2) = plt.subplots(1, 2, figsize=(7.3, 3.0))
nm = [q[0] for q in QUOTA]; eu = [q[1]*q[2]/1e6 for q in QUOTA]; gr = [v*0.024 for v in eu]
xp = np.arange(3); w = .36
a1.bar(xp-w/2, eu, w, color=C['navy'], label='EU-wide quota value')
a1.bar(xp+w/2, gr, w, color=C['orange'], label='notional Greek slice (2.4%)')
a1.set_yscale('log'); a1.set_xticks(xp); a1.set_xticklabels(nm, fontsize=7.6)
a1.set_ylabel('$ million, log scale'); a1.legend(fontsize=7.0)
for i, (e_, g) in enumerate(zip(eu, gr)):
    a1.text(i-w/2, e_*1.16, f'{e_:,.0f}', ha='center', fontsize=7.2)
    a1.text(i+w/2, g*1.16, f'{g:,.2f}', ha='center', fontsize=7.2)
a1.grid(axis='y', alpha=.35, which='both'); panel(a1, 'a')
head = [g/proj04 for g in gr]; be = [100*proj04/(g/0.024) for g in gr]
a2.bar(nm, head, color=C['teal'], width=.55)
for i, (h, b) in enumerate(zip(head, be)):
    a2.text(i, h+4, f'{h:,.0f}\u00d7', ha='center', fontsize=8.6, fontweight='bold')
    a2.text(i, h*.42, f'break-even\n{b:.3f}%', ha='center', fontsize=7.0,
            color='white', fontweight='bold')
a2.set_ylabel('structural headroom (\u00d7 projected Greek imports)')
a2.set_ylim(0, max(head)*1.22); a2.grid(axis='y', alpha=.35); panel(a2, 'b')
fig.tight_layout(); save(fig, 's7')

# ================= S8  delay bracket =================
def shifted_total():
    ti = tx = 0.0
    for c in CH:
        if c['t0'] > 0:
            ti += c['ci0']*(1+c['gi'])**11 * E.EPS_IMP * (1/(1+c['t0'])-1) * (1+c['w'])
        if c['m0'] > 0:
            my = c['m0']*max(0.0, 1-min(10, c['phx'])/c['phx']) if c['phx'] else 0.0
            tx += c['cx0']*(1+c['gx'])**11 * E.EPS_EXP * ((1+my)/(1+c['m0'])-1) * (1-c['capw']*0.17)
    return (ti-tx)/1e6
S2v = shifted_total()
fig, ax = plt.subplots(figsize=(7.3, 3.0))
ax.barh(['Wait-and-see bound\n(zero delay-year growth)', 'Compound-growth bound\n(full baseline growth)'],
        [W10, S2v], color=[C['green'], C['verm']], height=.5)
ax.axvline(W10, color=C['navy'], ls='--', lw=1.1)
for y, v in zip([0, 1], [W10, S2v]):
    ax.text(v+2.0, y, f'{v:,.1f}', va='center', fontsize=8.6, fontweight='bold')
ax.annotate('', xy=(S2v, 1.42), xytext=(W10, 1.42),
            arrowprops=dict(arrowstyle='<->', color=C['ink'], lw=1.3))
ax.text((W10+S2v)/2, 1.54, f'delay bracket: 0 to {100*(S2v/W10-1):,.1f}%',
        ha='center', fontsize=8.4, fontweight='bold')
ax.set_xlim(0, S2v*1.24); ax.set_ylim(-.6, 1.9)
ax.set_xlabel('Year 10 bilateral widening, $ million'); ax.grid(axis='x', alpha=.35)
save(fig, 's8')

# ================= S9  cadence and HS 15 reversal =================
yrs = np.arange(0, 11)
imp = [totals(y)[0] for y in yrs]; wid = [totals(y)[2] for y in yrs]; exp = [totals(y)[1] for y in yrs]
fig, (a1, a2) = plt.subplots(1, 2, figsize=(7.3, 3.05))
a1.plot(yrs, imp, '-o', ms=3.8, lw=1.9, color=C['verm'], label='additional imports')
a1.plot(yrs, wid, '-o', ms=3.8, lw=1.9, color=C['blue'], label='net widening')
a1.plot(yrs, exp, '-o', ms=3.8, lw=1.9, color=C['orange'], label='adjusted exports')
for y in (3, 5):
    a1.annotate(f'{100*imp[y]/imp[10]:,.0f}%', xy=(y, imp[y]), textcoords='offset points',
                xytext=(-5, 10), fontsize=7.6, color=C['verm'], fontweight='bold')
a1.set_xlabel('Year after entry into force'); a1.set_ylabel('$ million')
a1.set_xticks(yrs); a1.legend(loc='upper left'); a1.grid(alpha=.35); panel(a1, 'a')
net15 = [(project(CH, y)['15'][0]-project(CH, y)['15'][1])/1e6 for y in yrs]
a2.plot(yrs, net15, '-o', ms=3.8, lw=2.0, color=C['purple'])
a2.axhline(0, color=C['ink'], lw=1.0)
cross = next((y for y in yrs if net15[y] > 0), None)
if cross:
    a2.axvline(cross, color=C['navy'], ls='--', lw=1.1)
    a2.text(cross+.18, min(net15)*.55, f'sign reversal\nYear {cross}', fontsize=7.4, color=C['navy'])
a2.set_xlabel('Year after entry into force'); a2.set_ylabel('HS 15 net balance, $ million')
a2.set_xticks(yrs); a2.grid(alpha=.35); panel(a2, 'b')
fig.tight_layout(); save(fig, 's9')

# ================= S10  sensitivity grid =================
MULT = [0.5, 1.0, 1.5, 3.0]; DRAG = [0.07, 0.10, 0.20]
G = np.array([[totals(10, wmult=m, f2f=f)[2] for f in DRAG] for m in MULT])
fig, ax = plt.subplots(figsize=(6.0, 3.5))
im = ax.imshow(G, cmap='RdYlBu_r', aspect='auto')
ax.set_xticks(range(3)); ax.set_xticklabels([f'{int(d*100)}%' for d in DRAG])
ax.set_yticks(range(4)); ax.set_yticklabels([f'{m}\u00d7' for m in MULT])
ax.set_xlabel('Farm-to-Fork compliance drag')
ax.set_ylabel('wedge multiplier (de facto enforcement)')
for i in range(4):
    for j in range(3):
        ax.text(j, i, f'{G[i, j]:,.1f}', ha='center', va='center', fontsize=8.8,
                fontweight='bold' if (i == 1 and j == 1) else 'normal', color=C['ink'])
ax.add_patch(Rectangle((0.5, 0.5), 1, 1, fill=False, edgecolor=C['ink'], lw=2.2))
cb = fig.colorbar(im, ax=ax, shrink=.85); cb.set_label('Year 10 widening, $ million', fontsize=8)
ax.text(-0.45, 3.95, f'full grid span {G.max()-G.min():,.1f} M '
                     f'({100*(G.max()-G.min())/W10:,.1f}% of central); central cell boxed', fontsize=7.4)
save(fig, 's10')

# ================= S11  Monte Carlo =================
rng = random.Random(42); W, I, X, W5 = [], [], [], []
for _ in range(1000):
    a = rng.triangular(0.5, 3.0, 1.0); b = rng.triangular(0.05, 0.20, 0.10); c_ = rng.triangular(0.03, 0.12, 0.07)
    r = totals(10, wmult=a, f2f=b, lab=c_); W.append(r[2]); I.append(r[0]); X.append(r[1])
    W5.append(totals(5, wmult=a, f2f=b, lab=c_)[2])
Ws = sorted(W); lo, hi = Ws[24], Ws[974]
fig, (a1, a2) = plt.subplots(1, 2, figsize=(7.3, 3.05))
a1.hist(W, bins=38, color=C['sky'], edgecolor='white', linewidth=.4)
a1.axvspan(lo, hi, color=C['blue'], alpha=.24)
a1.axvline(statistics.mean(W), color=C['verm'], lw=1.7)
a1.set_xlabel('Year 10 widening, $ million'); a1.set_ylabel('draws')
a1.text(.03, .95, f'mean {statistics.mean(W):,.1f}\n95% CI [{lo:,.1f}, {hi:,.1f}]\n1,000 draws, seed 42',
        transform=a1.transAxes, va='top', fontsize=7.6)
panel(a1, 'a')
names = ['Y5 widening', 'Y10 widening', 'Y10 imports', 'Y10 adj exports']
series = [W5, W, I, X]
mns = [statistics.mean(s) for s in series]
los = [sorted(s)[24] for s in series]; his = [sorted(s)[974] for s in series]
yp = np.arange(4)[::-1]
a2.errorbar(mns, yp, xerr=[np.array(mns)-np.array(los), np.array(his)-np.array(mns)],
            fmt='o', ms=6, color=C['blue'], ecolor=C['navy'], capsize=3.5, lw=1.5)
for y, mn, l, h in zip(yp, mns, los, his):
    a2.text(mn, y+.24, f'{mn:,.1f}  [{l:,.1f}, {h:,.1f}]', ha='center', fontsize=7.2)
a2.set_yticks(yp); a2.set_yticklabels(names); a2.set_xlabel('$ million')
a2.set_xlim(0, max(his)*1.28); a2.grid(axis='x', alpha=.35); panel(a2, 'b')
fig.tight_layout(); save(fig, 's11')

# ================= S13  environmental-tax asymmetry =================
fig, (a1, a2) = plt.subplots(1, 2, figsize=(7.3, 2.95))
a1.bar(['Greece', 'LAC proxy'], [3.6, 1.2], color=[C['verm'], C['teal']], width=.5)
a1.annotate('', xy=(0, 3.6), xytext=(0, 1.2), arrowprops=dict(arrowstyle='<->', color=C['ink'], lw=1.4))
a1.text(0.13, 2.4, 'gap 2.4 pp', fontsize=8.4, fontweight='bold')
for i, v in enumerate([3.6, 1.2]):
    a1.text(i, v+.12, f'{v}%', ha='center', fontsize=8.4, fontweight='bold')
a1.set_ylabel('environmental-tax revenue, % of GDP'); a1.set_ylim(0, 4.4)
a1.grid(axis='y', alpha=.35); panel(a1, 'a')
a2.bar(['EUDR\n(exporter-borne)', 'Env-tax addon\n(Greek-borne)', 'Net HS 23 wedge'],
       [-1.0, 0.5, -0.5], color=[C['green'], C['orange'], C['blue']], width=.55)
a2.axhline(0, color=C['ink'], lw=1.0)
for i, v in enumerate([-1.0, 0.5, -0.5]):
    a2.text(i, v + (.07 if v > 0 else -.17), f'{v:+.1f}', ha='center', fontsize=8.4, fontweight='bold')
a2.set_ylabel('percentage points of per-unit cost'); a2.set_ylim(-1.35, .82)
a2.text(.5, .05, '20% pass-through of the 2.4 pp gap', transform=a2.transAxes,
        ha='center', fontsize=7.2)
a2.grid(axis='y', alpha=.35); panel(a2, 'b')
fig.tight_layout(); save(fig, 's13')

# ================= S14  five readings =================
fig, ax = plt.subplots(figsize=(7.3, 2.5)); ax.set_ylim(0, 3); ax.set_xlim(-.6, 4.6); ax.axis('off')
reads = [('Gohin &\nMatthews', 'muted effects', 0.15),
         ('EP assessment\n(Hagemejer)', 'aggregate positive,\nuneven', 1.35),
         ('Campos et al.', 'import-side\nconcentration', 2.55),
         ('Fert\u0151', 'protective\nasymmetry', 3.7)]
ax.annotate('', xy=(4.4, 1.6), xytext=(-.35, 1.6), arrowprops=dict(arrowstyle='->', color=C['ink'], lw=1.5))
ax.text(-.35, 1.26, 'muted effect', fontsize=7.8, color=C['teal'], fontweight='bold')
ax.text(4.4, 1.26, 'binding asymmetry', fontsize=7.8, ha='right', color=C['verm'], fontweight='bold')
for nm_, sub, x in reads:
    ax.plot([x], [1.6], 'o', ms=9.5, color=C['navy'])
    ax.text(x, 2.02, nm_, ha='center', fontsize=7.6, fontweight='bold')
    ax.text(x, 0.92, sub, ha='center', fontsize=7.0, color=C['navy'])
ax.plot([3.12], [1.6], '*', ms=19, color=C['orange'])
ax.text(3.12, 0.36, 'Greek evidence\n(this article)', ha='center', fontsize=7.8,
        fontweight='bold', color=C['orange'])
save(fig, 's14')

# ================= S15  exposure typology =================
fig, ax = plt.subplots(figsize=(7.0, 3.3)); ax.set_xlim(0, 3.1); ax.set_ylim(0, 3); ax.axis('off')
cells = [('1. Tree crops, GIs,\nsmall ruminants',
          'exposure in imported inputs,\norganisation in final goods\n\u2192 weak effective opposition', 0.05, '#FCEBD5', C['orange']),
         ('2. Industrial\nexport advantage',
          'favourable balance offsets\nagricultural exposure\n\u2192 support', 1.07, '#E2F3EA', C['green']),
         ('3. Feed-deficit livestock +\norganised horticulture',
          'exposure and capacity overlap\n\u2192 quotas bind, resistance', 2.09, '#EAF1F8', C['navy'])]
for t, s_, x, fc, ec in cells:
    box(ax, x, 1.35, 0.94, 1.25, t + '\n' + s_, fc=fc, ec=ec, fs=6.7)
ax.text(0.52, 1.12, 'Greece, Spain, Italy', ha='center', fontsize=7.6,
        fontweight='bold', color=C['orange'])
ax.text(0.52, 0.80, 'Greece adds a wheat-dominated cereal area\nand the highest OECD environmental-tax burden',
        ha='center', fontsize=6.9, color=C['orange'])
ax.text(1.55, 2.82, 'exposure-capacity configurations across member states',
        ha='center', fontsize=9, fontweight='bold')
save(fig, 's15')

# ================= S16  sector scoreboard =================
sect = [('23  Animal feed', NET['23'], 'no designation'),
        ('22  Beverages, wine', NET['22'], 'PDO/PGI on 26% of area'),
        ('09  Coffee', NET['09'], 'no Greek production'),
        ('20  Prepared vegetables', NET['20'], 'table-olive designations'),
        ('15  Fats and oils', NET['15'], 'Kalamata PDO'),
        ('08  Fruit and nuts', NET['08'], 'Korinthiaki PDO'),
        ('04  Dairy (Feta)', NET['04'], 'Feta PDO, 7-y transition')]
fig, ax = plt.subplots(figsize=(7.3, 3.5))
ys = np.arange(len(sect))[::-1]; vals = [v for _, v, _ in sect]
ax.barh(ys, vals, color=[C['verm'] if v > 0 else C['green'] for v in vals], height=.6)
ax.set_yticks(ys); ax.set_yticklabels([n for n, _, _ in sect], fontsize=7.6)
ax.axvline(0, color=C['ink'], lw=1.0)
for y, (n, v, a_) in zip(ys, sect):
    ax.text(v + (3 if v > 0 else -4), y, f'{v:,.1f}', va='center',
            ha='left' if v > 0 else 'right', fontsize=7.8, fontweight='bold')
    ax.text(max(vals)*1.28, y, a_, va='center', fontsize=6.9, color=C['navy'])
ax.set_xlabel('Year 10 net position, $ million')
ax.set_xlim(min(vals)*2.2, max(vals)*2.05); ax.grid(axis='x', alpha=.35)
save(fig, 's16')

# ================= S17  ratification pathway =================
fig, ax = plt.subplots(figsize=(7.3, 2.7)); ax.set_ylim(0, 4); ax.set_xlim(2023.5, 2042.5); ax.axis('off')
ax.annotate('', xy=(2042.2, 2.0), xytext=(2023.8, 2.0), arrowprops=dict(arrowstyle='->', color=C['ink'], lw=1.5))
marks = [(2024.9, 'political\nconclusion', C['navy']), (2026.0, 'S1 entry', C['green']),
         (2027.0, f'S2 entry\n(+{100*(S2v/W10-1):.1f}%)', C['verm']),
         (2033.0, 'modal EU-side\nphase complete', C['blue']),
         (2034.0, 'HS 15 sign\nreversal', C['purple']),
         (2041.0, 'longest Mercosur\nschedule (15 y)', C['orange'])]
for x, t, col in marks:
    ax.plot([x], [2.0], 'o', ms=8.5, color=col)
    ax.plot([x, x], [2.0, 2.45], color=col, lw=1.1)
    ax.text(x, 2.56, t, ha='center', fontsize=6.9, color=col, fontweight='bold')
ax.axvspan(2028, 2034, ymin=.24, ymax=.36, color=C['teal'], alpha=.32)
ax.text(2031, 1.24, 'MFF 2028-2034: the adjustment window', ha='center',
        fontsize=7.8, fontweight='bold', color=C['teal'])
for yr in range(2024, 2043, 3):
    ax.text(yr, 1.74, str(yr), ha='center', fontsize=7.0)
save(fig, 's17')

# ================= S18  policy directions =================
pol = [('Feed-grain reorientation\n100-200 kha under the MFF', NET['23'], C['verm']),
       ('Designation enforcement\nand extension', abs(NET['08'])+abs(NET['04']), C['green']),
       ('Wine sector: PDO share\nand export promotion', NET['22'], C['orange']),
       ('Table-olive processing\ncapacity', NET['20'], C['purple']),
       ('Industrial export support\n(delay costs)', abs(NET['73'])+abs(NET['30']), C['blue'])]
fig, ax = plt.subplots(figsize=(7.3, 3.2))
ys = np.arange(len(pol))[::-1]
ax.barh(ys, [v for _, v, _ in pol], color=[c for _, _, c in pol], height=.6)
ax.set_yticks(ys); ax.set_yticklabels([n for n, _, _ in pol], fontsize=7.3)
ax.set_xscale('log')
for y, (n, v, c) in zip(ys, pol):
    ax.text(v*1.12, y, f'{v:,.1f}', va='center', fontsize=7.8, fontweight='bold')
ax.set_xlabel('bilateral flow addressed, $ million, log scale')
ax.grid(axis='x', alpha=.35, which='both')
save(fig, 's18')

print(f'\nSupplementary figures written to {A.out} as PNG and EPS.')
print('S12 omitted: it draws on external FADN and Eurostat series not held in the workbook.')
