"""Regenerate Figure 4 of the manuscript.

Panel (a) reproduces the published cereal-area composition unchanged.
Panel (b) replots the verified HS 23 additional-import series and states the
Year 1 to Year 7 compound rate on the figure itself.

Outputs: vector EPS and PDF for submission, PNG for review, and an
uncompressed 300 dpi TIFF matching the dimensions of the embedded original.
"""
import numpy as np, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyArrowPatch
from PIL import Image

NAVY, GOLD, VERM = '#1F3B73', '#E69F00', '#D55E00'
GRID, INK, AXIS = '#9FB6CD', '#111111', '#4D4D4D'
# EPS has no alpha channel, so the phase-in band uses a pre-blended solid tint
# equal to VERM at 7 per cent over white.
BAND = '#FCF4ED'

plt.rcParams.update({
    'font.family': 'DejaVu Sans', 'font.size': 9,
    'text.color': INK, 'axes.labelcolor': INK,
    'xtick.color': AXIS, 'ytick.color': AXIS,
    'axes.edgecolor': AXIS, 'axes.linewidth': 0.9,
    'savefig.facecolor': 'white', 'figure.facecolor': 'white',
    'ps.fonttype': 42, 'pdf.fonttype': 42,   # embed TrueType, keep text as text
})

cats = ['Wheat', 'Barley', 'Maize', 'Other\ncereals']
area = [369, 94, 43, 80]
cols = [NAVY, GOLD, GOLD, NAVY]

years = np.arange(0, 11)
imp = np.array([0.000, 7.943, 17.022, 27.359, 39.086, 52.351,
                67.312, 84.145, 90.225, 96.745, 103.735])
cagr = (imp[7] / imp[1]) ** (1 / 6) - 1

fig, (axa, axb) = plt.subplots(1, 2, figsize=(7.18, 3.193))
fig.subplots_adjust(left=0.088, right=0.985, top=0.905, bottom=0.175, wspace=0.28)

# ---------------- panel (a) ----------------
axa.set_axisbelow(True)
axa.yaxis.grid(True, color=GRID, lw=0.8)
bars = axa.bar(cats, area, color=cols, width=0.60)
for b, v in zip(bars, area):
    axa.text(b.get_x() + b.get_width() / 2, v + 7, f'{v}',
             ha='center', va='bottom', fontsize=9)
axa.set_ylabel('thousand hectares', fontsize=9)
axa.set_ylim(0, 430); axa.set_yticks(range(0, 401, 50))
axa.tick_params(labelsize=9, length=3.5)
for s in ('top', 'right'): axa.spines[s].set_visible(False)
axa.text(0.28, 0.70, 'feed-relevant: 137.0 kha\nagainst 1.87 M livestock units',
         transform=axa.transAxes, color=VERM, fontsize=8.4, linespacing=1.4)
axa.text(0.04, 1.03, 'a', transform=axa.transAxes, fontsize=12,
         fontweight='bold', va='bottom')

# ---------------- panel (b) ----------------
axb.set_axisbelow(True)
axb.grid(True, color=GRID, lw=0.8)
axb.axvspan(1, 7, facecolor=BAND, edgecolor='none', lw=0, zorder=0)
axb.plot(years, imp, color=VERM, lw=2.0, marker='o', markersize=4.6,
         markerfacecolor=VERM, markeredgecolor=VERM, zorder=4)
axb.axvline(7, color=NAVY, lw=1.5, ls=(0, (5, 4)), zorder=3)

axb.add_patch(FancyArrowPatch((1, 104), (7, 104),
                              arrowstyle='|-|,widthA=2.6,widthB=2.6',
                              color=VERM, lw=1.1, mutation_scale=1, zorder=5))
axb.annotate('', xy=(7, 104), xytext=(1, 104),
             arrowprops=dict(arrowstyle='-', color=VERM, lw=1.1), zorder=5)
axb.text(4.0, 112, f'compounds at {100 * cagr:.0f}% annually',
         color=VERM, fontsize=8.4, ha='center', va='center')
axb.text(7.3, 26, 'phase-out\ncomplete\nYear 7', color=NAVY,
         fontsize=8.4, linespacing=1.4, va='center')

axb.set_xlabel('Year after entry into force', fontsize=9)
axb.set_ylabel('additional HS 23 imports, $ million', fontsize=9)
axb.set_xticks(years); axb.set_xlim(-0.45, 10.45)
axb.set_ylim(-6, 122); axb.set_yticks(range(0, 101, 20))
axb.tick_params(labelsize=9, length=3.5)
for s in ('top', 'right'): axb.spines[s].set_visible(False)
axb.text(0.04, 1.03, 'b', transform=axb.transAxes, fontsize=12,
         fontweight='bold', va='bottom')

fig.savefig('Figure4.eps', format='eps')
fig.savefig('Figure4.pdf', format='pdf')
fig.savefig('fig4_new.png', dpi=300)
print(f'Year 1 to Year 7 compound rate: {100 * cagr:.2f}%')

im = Image.open('fig4_new.png').convert('RGB').resize((2154, 958), Image.LANCZOS)
im.save('image4.tiff', format='TIFF', compression='raw', dpi=(300, 300))
print('wrote Figure4.eps, Figure4.pdf, fig4_new.png, image4.tiff')
