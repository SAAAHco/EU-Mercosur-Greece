"""Main-text figures for the corrected (v8.1) specification. Writes PNG (300 dpi) and EPS.
Colors: adverse (red #e34948), favorable (blue #2a78d6), published/neutral (gray #8a8983);
text in ink, never in series colors."""
import json, os, sys
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
HERE = os.path.dirname(os.path.abspath(__file__))
os.chdir(HERE)
import pe_v8
import run_v81 as R

OUT = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, 'figs_v81')
os.makedirs(OUT, exist_ok=True)
RED, BLUE, GRAY, INK, INK2, GRID = '#e34948', '#2a78d6', '#8a8983', '#0b0b0b', '#52514e', '#e4e3df'
plt.rcParams.update({'font.family': 'DejaVu Sans', 'font.size': 9, 'axes.edgecolor': INK2, 'axes.labelcolor': INK,
                     'xtick.color': INK2, 'ytick.color': INK2, 'axes.spines.top': False, 'axes.spines.right': False})
SHORT = {'03': 'Fish and crustaceans', '24': 'Tobacco', '38': 'Chemical products (rosin)', '02': 'Meat',
         '20': 'Fruit and vegetable preparations', '64': 'Footwear', '44': 'Wood', '85': 'Electrical machinery',
         '23': 'Animal feed', '32': 'Paints and dyes', '29': 'Organic chemicals', '40': 'Rubber', '68': 'Stone articles', '09': 'Coffee, tea, spices', '22': 'Beverages', '08': 'Fruit and nuts', '30': 'Pharmaceuticals',
         '90': 'Optical instruments', '84': 'Machinery', '27': 'Mineral fuels', '82': 'Tools', '73': 'Iron and steel articles',
         '15': 'Fats and oils', '04': 'Dairy', '28': 'Inorganic chemicals', '76': 'Aluminum', '12': 'Oilseeds', '10': 'Cereals'}


def save(fig, name):
    fig.savefig(os.path.join(OUT, name + '.png'), dpi=300, bbox_inches='tight', facecolor='white')
    fig.savefig(os.path.join(OUT, name + '.eps'), bbox_inches='tight', facecolor='white')
    plt.close(fig)


# ---------- data ----------
h81, cfg81, chs81 = R.run()
net81 = {c['hs']: c['net'] / 1e6 for c in h81['chapters']}
v5 = pe_v8.prepare_chapters(json.load(open(os.path.join(HERE, 'inputs', 'v5_config.json')))['chapters'])
h75 = pe_v8.headline(v5, pe_v8.CAP_WEIGHTS_V75, pe_v8.TRQ_CAPS_V75)
net75 = {c['hs']: c['net'] / 1e6 for c in h75['chapters']}

# ---------- Figure 1: published versus corrected, by chapter ----------
keys = sorted(set(net75) | set(net81), key=lambda k: -max(abs(net75.get(k, 0)), abs(net81.get(k, 0))))[:12]
keys = sorted(keys, key=lambda k: net75.get(k, 0))
fig, ax = plt.subplots(figsize=(6.6, 4.4))
y = range(len(keys))
ax.barh([i + 0.2 for i in y], [net75.get(k, 0) for k in keys], height=0.38, color=GRAY, label='Chapter-average specification (25 chapters)')
ax.barh([i - 0.2 for i in y], [net81.get(k, 0) for k in keys], height=0.38, color=BLUE, label='Line-level specification (44 chapters)')
ax.set_yticks(list(y))
ax.set_yticklabels([f"{SHORT.get(k, k)} (HS {k})" for k in keys])
ax.axvline(0, color=INK2, lw=0.8)
ax.xaxis.grid(True, color=GRID, lw=0.6)
ax.set_axisbelow(True)
ax.set_xlabel('Year 10 widening of the bilateral balance, million US dollars\n(negative values are favorable to Greece)')
ax.legend(frameon=False, loc='lower right', fontsize=8)
for i, k in enumerate(keys):
    for off, v in ((0.2, net75.get(k, 0)), (-0.2, net81.get(k, 0))):
        if abs(v) >= 5:
            ax.text(v + (1.5 if v > 0 else -1.5), i + off, f"{v:.1f}".replace('-', '−'), va='center', ha='left' if v > 0 else 'right', fontsize=7, color=INK)
ax.set_xlim(-32, max(net75.values()) + 18)
save(fig, 'Figure_1')

# ---------- Figure 3: where the corrected adjustment falls ----------
GI = R.GI
sel = [k for k, v in net81.items() if abs(v) >= 0.2 or k in GI]
sel = sorted(sel, key=lambda k: net81[k])
fig, ax = plt.subplots(figsize=(6.6, 5.0))
vals = [net81[k] for k in sel]
ax.barh(range(len(sel)), vals, height=0.62, color=[RED if v > 0 else BLUE for v in vals])
ax.set_yticks(range(len(sel)))
ax.set_yticklabels([f"{SHORT.get(k, k)} (HS {k})" + (f"  [{GI[k]} PDO/PGI]" if k in GI else '') for k in sel], fontsize=8)
ax.axvline(0, color=INK2, lw=0.8)
ax.xaxis.grid(True, color=GRID, lw=0.6)
ax.set_axisbelow(True)
for i, v in enumerate(vals):
    if abs(v) >= 1:
        ax.text(v + (0.6 if v > 0 else -0.6), i, f"{v:.1f}".replace('-', '−'), va='center', ha='left' if v > 0 else 'right', fontsize=7, color=INK)
ax.set_xlim(min(vals) - 6, max(vals) + 5)
ax.set_xlabel('Year 10 widening of the bilateral balance, million US dollars')
ax.text(0.99, 0.02, 'red: additional imports exceed additional exports\nblue: additional exports exceed additional imports',
        transform=ax.transAxes, ha='right', va='bottom', fontsize=7, color=INK2)
save(fig, 'Figure_3')

# ---------- Figure 2: sensitivity of the Year 10 result ----------
res = json.load(open('results_v81.json'))
central = res['central']['widening'] / 1e6
sens = dict(res['sensitivities'])
order = [('elasticity -2.0', 'Import elasticity −2.0'), ('elasticity -2.5', 'Import elasticity −2.5'),
         ('elasticity -5.0', 'Import elasticity −5.0'), ('elasticity -8.0', 'Import elasticity −8.0'),
         ('no baseline growth (static)', 'No counterfactual growth'), ('growth capped at +/-5%', 'Growth bounded at ±5% a year'),
         ('growth capped at +/-10%', 'Growth bounded at ±10% a year'), ('growth on 3-year averages', 'Growth on three-year averages'),
         ('spike screen off (one-offs and spikes kept)', 'Spike screen off'),
         ('fish processing quotas: base 0%', 'Seafood entering under 0% processing quotas'),
         ('tobacco WITS AVE (5.97%)', 'Tobacco leaf at the TRAINS average (5.97%)'), ('tobacco uncertified (9.12%)', 'Tobacco leaf at the uncertified duty (9.12%)'),
         ('beef Hilton quota duty removal (with Greek slice)', 'Hilton quota duty removed, with a Greek slice'), ('spike screen also drops 2024-only flows', 'Spike screen also drops flows first seen in 2024')]
order = sorted(order, key=lambda t: sens[t[0]])
fig, ax = plt.subplots(figsize=(6.6, 4.2))
for i, (k, lab) in enumerate(order):
    v = sens[k] / 1e6
    ax.plot([central, v], [i, i], color=GRAY, lw=1.5)
    ax.plot(v, i, 'o', ms=6, color=RED if v > 0 else BLUE, mec='white', mew=1.2)
    ax.text(v + (2 if v >= central else -2), i, f"{v:.1f}".replace('-', '−'), va='center', ha='left' if v >= central else 'right', fontsize=7, color=INK)
ax.axvline(central, color=INK, lw=1.0)
ax.axvline(0, color=INK2, lw=0.8, ls=':')
ax.text(central + 1.5, len(order) - 0.45, f'central case {central:.1f}', ha='left', va='bottom', fontsize=7.5, color=INK)
ax.set_yticks(range(len(order)))
ax.set_yticklabels([lab for _, lab in order], fontsize=8)
ax.xaxis.grid(True, color=GRID, lw=0.6)
ax.set_axisbelow(True)
ax.set_xlabel('Year 10 widening of the bilateral balance, million US dollars')
ax.set_ylim(-0.7, len(order) + 0.3)
ax.set_xlim(min(sens[k] for k, _ in order) / 1e6 - 12, max(sens[k] for k, _ in order) / 1e6 + 12)
save(fig, 'Figure_2')
print('figures written to', OUT, '| central', round(central, 2))
