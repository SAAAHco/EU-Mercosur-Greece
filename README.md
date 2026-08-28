# Trade-agreement exposure toolkit: Greece and the EU-Mercosur Agreement

Konstantinos Pappas, Zainab Ashkanani, Rashed Albatayneh
Texas A&M University

Code accompanying the manuscript *"Protection without exposure, exposure without protection:
Greece and the EU-Mercosur Agreement"* (*Journal of European Public Policy*).

This repository contains the **full analysis code** together with **blank input templates**,
so that the method can be applied to any bilateral trade relationship: a partial-equilibrium
(PE) engine that projects bilateral trade under a trade agreement's tariff phase-out, an
**XGBoost machine-learning companion** that validates the counterfactual baseline, and the
robustness and figure pipeline. Users supply their own trade data and parameters through the
templates in `templates/` and run the scripts below.

---

## What the code does

| Component | Script | Method |
|---|---|---|
| PE projection engine + ML companion + robustness | `code/orchestrator_v7_5.py` | Armington import/export projections per HS-2 chapter over an 11-year phase-in, M = M_cf × [1 + ε·π·W] with the tariff shock expressed as the proportional change in the tariff-inclusive price π = (1+τ₁)/(1+τ₀) − 1 (WITS/SMART form); incidence-signed non-tariff cost wedges; per-chapter export-capacity adjustment X·[1 − κ(c)(δ_f2f + δ_lab·y/10)]; TRQ caps; two entry-year scenarios; wedge × drag sensitivity grid; 1,000-draw Monte Carlo with triangular priors; **XGBoost companion** (gradient-boosted regression, fixed hyperparameters, holdout-year evaluation, direct horizon-specific prediction) with its own Monte Carlo |
| Workbook verification | `code/build_workbook_v7_5.py` | Recomputes projection cells of an existing analysis workbook and prints a verification table |
| Main figures | `code/make_figures_v8.py` | Six publication figures (PNG + EPS) drawn from the analysis workbook |
| Supplementary figures | `code/make_appendix_figs_v8.py` | Seventeen appendix figures; imports `make_figures_v8.py` as its engine |
| Single-figure regeneration example | `code/make_figure4.py` | Standalone two-panel figure (EPS/PDF/TIFF) |
| Elasticity sensitivity | `code/elasticity_sensitivity.py` | Re-runs the Year-10 projection across alternative Armington import elasticities (−2.0 … −8.0) |

## Repository structure

```
├── code/                                  All analysis scripts (see table above)
├── templates/
│   ├── TradeData_template.xlsx            Blank bilateral trade-data input (record level)
│   ├── v5_config_template.json            Blank chapter-parameter register (tariffs, phase-ins,
│   │                                      baselines, growth rates, elasticities, wedges)
│   ├── Analysis_Workbook_template.xlsx    Blank analysis-workbook structure (16 sheets) read by
│   │                                      the figure and verification scripts
│   └── README.md                          Field-by-field instructions for every template
├── requirements.txt
├── LICENSE                                MIT
└── CITATION.cff
```

## Installation

Python ≥ 3.10, then:

```
pip install -r requirements.txt
```

## How to use

**Step 1 — fill the two input templates** (full field documentation in `templates/README.md`):

1. `templates/TradeData_template.xlsx` — one row per (year × flow × partner × HS-2 chapter)
   trade record, e.g. from UN Comtrade. Save it as
   `Greece__Mar_Total_TradeData_Balance_Sheet.xlsx` next to `code/orchestrator_v7_5.py`.
2. `templates/v5_config_template.json` — one entry per HS-2 chapter: MFN tariffs on both
   sides, phase-out lengths, baseline flows, counterfactual growth rates, elasticities and
   non-tariff wedge magnitudes, each with its source citation. Save it as `v5_config.json`
   next to `code/orchestrator_v7_5.py`.

**Step 2 — run the full pipeline:**

```
python code/orchestrator_v7_5.py
```

The orchestrator prints a six-stage progress log — (1) PE Armington projections,
(2) sensitivity grid and PE Monte Carlo, (3) XGBoost training, scenario prediction and ML
Monte Carlo, (4) Excel workbook writer, (5) diagnostic charts, (6) audit log — and writes
to an `outputs/` folder next to the script:

- `EU_Mercosur_Greece_Analysis_v7.xlsx` — the complete 15-sheet analysis workbook
  (raw data, parameters, both scenarios, summary, TRQ analysis, sensitivity grid,
  Monte Carlo intervals, year-by-year detail, ML-vs-PE comparison, capacity adjustments)
- `charts/01…07_*.png` — diagnostic charts (feature importance, sensitivity heatmap,
  ML vs PE, Monte Carlo distributions, decomposition, yearly trajectory)
- `orchestrator_run.json` — full configuration and per-stage audit log

**Step 3 — draw the figures** from the workbook produced in Step 2:

```
python code/make_figures_v8.py --workbook outputs/EU_Mercosur_Greece_Analysis_v7.xlsx --out figures_main
python code/make_appendix_figs_v8.py --engine code/make_figures_v8.py --out figures_supp
```

**Step 4 — run the elasticity sensitivity:**

```
python code/elasticity_sensitivity.py --workbook outputs/EU_Mercosur_Greece_Analysis_v7.xlsx
```

**Alternative entry point:** if you prefer to parameterize the model in Excel rather than
run the orchestrator, fill the input sheets of `templates/Analysis_Workbook_template.xlsx`
(sheets 2, 3, 4, 5 and 13 — the template's note sheet explains which) and point the Step 3
and Step 4 scripts at that file.

## The machine-learning companion

Module 3 of the orchestrator trains a gradient-boosted regression (XGBoost) on the
bilateral panel: 11 features (one- and two-year lagged trade values, lagged growth,
log distance, log GDPs of both partners, a tariff proxy, direction indicator, year trend,
and chapter/partner identifiers), fixed hyperparameters (200 trees, depth 5, learning
rate 0.08, fixed seed), the final panel year held out for evaluation (RMSE reported),
and direct horizon-specific prediction with policy and macro features shifted to the
target year. Its role is to supply a persistence baseline that validates the PE
counterfactual — where tariffs do not vary within chapters over the estimation window,
the ML model cannot identify a tariff response, so uncertainty about the policy effect
is carried by the import-elasticity sensitivity (Step 4) instead. Gain importance,
holdout metrics and Monte Carlo intervals are written to the workbook's Monte Carlo and
ML-vs-PE sheets and to the diagnostic charts.

## Methodological notes

- **Tariff shock specification.** The engine expresses the tariff shock as the
  proportional change in the tariff-inclusive price, π = (1+τ₁)/(1+τ₀) − 1 (the standard
  WITS/SMART form). Expressing it as the proportional change in the tariff *rate*
  (−Δτ/τ₀) overstates the import response — by roughly 21× at a 5% MFN — and must not be
  used; the code headers document this correction.
- **Incidence-signed wedges.** Non-tariff cost channels are signed by who bears the cost:
  exporter-borne compliance costs dampen the import response (negative), importer-side
  producer burdens amplify it (positive). Enter wedge magnitudes as positive decimals in
  the config; the code signs them at load time.
- **Determinism.** Both Monte Carlos and the XGBoost model run on a fixed seed; results
  are reproducible conditional on the seed and the configuration.
- Some figure scripts embed observed series and label constants from the published
  analysis as plotting literals (annotations, benchmark panels); when applying the
  toolkit to your own data, replace those literals with your own values — they are marked
  by comments in the scripts.

## Data sources used in the published analysis

UN Comtrade (bilateral HS-2 flows; https://comtradeplus.un.org/), Eurostat COMEXT
(cross-validation), the EU-Mercosur Agreement technical annexes (tariff schedules,
phase-ins, TRQs), ELSTAT 2020 Agricultural–Livestock Census, FADN Farm Economy Focus,
and OECD environmental-tax indicators. Full parameter provenance is reported in the
article's supplementary appendix (Table A2).

## Created by
Dr. Zainab Ashkanani
Email: Ashkanani@saaah.co

## License and citation

Code is released under the MIT License (see `LICENSE`). If you use this toolkit, please
cite the article (see `CITATION.cff`).
