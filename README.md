# Trade-agreement exposure toolkit: Greece and the EU-Mercosur Agreement

Konstantinos Pappas, Zainab Ashkanani, Rashed Albatayneh
Texas A&M University

Code, data and templates accompanying the manuscript *"Protection without exposure, exposure
without protection: the distributive and environmental incidence of the EU-Mercosur Agreement
in Greek agriculture"* (submitted to *Ecological Economics*).

The repository serves two purposes:

1. **Replication.** `data/` holds the analytical workbook that is simultaneously the seed
   dataset, the parameter register and the full results ledger of the published analysis, the
   audit register of the tariff-specification correction, and the published chapter-parameter
   file. Every number in the article can be verified from these files with the scripts in
   `code/` (see *Verifying the published results*).
2. **Re-use.** `templates/` holds blank input files so the method can be applied to another
   bilateral trade relationship: a partial-equilibrium (PE) engine that projects bilateral
   trade under a trade agreement's tariff phase-out, an XGBoost machine-learning companion
   that validates the counterfactual baseline, and the robustness and figure pipeline. Some
   country-specific constants live in the code rather than the templates; they are listed in
   *Applying the toolkit elsewhere*.

---

## What the code does

| Component | Script | Method |
|---|---|---|
| PE projection engine + ML companion + robustness | `code/orchestrator_v7_5.py` | Armington import/export projections per HS-2 chapter over an 11-year phase-in, M = M_cf × [1 + ε·π·W] with the tariff shock expressed as the proportional change in the tariff-inclusive price π = (1+τ₁)/(1+τ₀) − 1 (WITS/SMART form); incidence-signed non-tariff cost wedges; per-chapter export-capacity adjustment X·[1 − κ(c)(δ_f2f + δ_lab·y/10)]; TRQ caps; two entry-year scenarios; 4 × 3 wedge × drag sensitivity grid; 1,000-draw Monte Carlo with triangular priors; XGBoost companion (gradient-boosted regression, fixed hyperparameters, holdout-year evaluation, direct horizon-specific prediction) with its own Monte Carlo |
| Workbook verification | `code/build_workbook_v7_5.py` | Recomputes the projection cells of an analysis workbook and prints a 22-line verification table against the published values (Greek analysis only) |
| Main figures | `code/make_figures_v8.py` | Six publication figures (PNG + EPS) drawn from the analysis workbook |
| Supplementary figures | `code/make_appendix_figs_v8.py` | Seventeen appendix figures (S1 to S11, S13 to S18); imports `make_figures_v8.py` as its engine |
| Single-figure regeneration example | `code/make_figure4.py` | Standalone two-panel figure (EPS/PDF/TIFF) |
| Elasticity sensitivity | `code/elasticity_sensitivity.py` | Re-runs the Year-10 projection across alternative Armington import elasticities (−2.0 … −8.0) |

## Repository structure

```
├── code/                                  All analysis scripts (see table above)
├── data/
│   ├── EU_Mercosur_Greece_Analysis_v7_5.xlsx   Analytical workbook: seed data, parameter
│   │                                           register and results ledger (16 sheets)
│   ├── Tariff_Specification_Correction.xlsx    Audit register of the specification correction
│   ├── v5_config.json                          Published 25-chapter parameter file (Greece)
│   └── README.md                               Sheet guide and notes on the raw inputs
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

`requirements.txt` pins **XGBoost to the 2.x series** (`xgboost>=2.0,<3`). The
machine-learning values reported in the article (holdout RMSE, gain importances, the
persistence projection and its Monte Carlo interval) were generated with XGBoost 2.1 and are
reproduced exactly by versions 2.0.3 and 2.1.4. XGBoost 3.x changes tree-construction defaults
and produces different values; the partial-equilibrium results do not depend on XGBoost.

## Verifying the published results

All commands are run from the repository root.

**1. Headline estimates and the Appendix F elasticity sensitivity** (seconds; needs only
`openpyxl`):

```
python code/elasticity_sensitivity.py --workbook data/EU_Mercosur_Greece_Analysis_v7_5.xlsx
```

Expected output — the central row is the manuscript headline, the other rows are Appendix F:

| ε_i | Δ imports ($M) | Δ exports, adj. ($M) | Widening ($M) |
|------|------|------|------|
| −2.0 | 112.6 | 51.1 | 61.5 |
| −2.5 | 140.8 | 51.1 | 89.7 |
| **−3.5 (central)** | **197.1** | **51.1** | **146.0** |
| −5.0 | 281.5 | 51.1 | 230.4 |
| −8.0 | 450.4 | 51.1 | 399.3 |

**2. Full verification table** (recomputes every projection cell and checks 22 quantities
against the manuscript, ending in `ALL CHECKS PASSED`):

```
python code/build_workbook_v7_5.py --src data/EU_Mercosur_Greece_Analysis_v7_5.xlsx --out verified_copy.xlsx
```

**3. Figures** from the deposited workbook:

```
python code/make_figures_v8.py --workbook data/EU_Mercosur_Greece_Analysis_v7_5.xlsx --out figures_main
python code/make_appendix_figs_v8.py --engine code/make_figures_v8.py --workbook data/EU_Mercosur_Greece_Analysis_v7_5.xlsx --out figures_supp
python code/make_figure4.py
```

**4. Full end-to-end re-run** (about one minute; re-estimates everything, including the
XGBoost companion and both Monte Carlos). Two files must sit next to `code/orchestrator_v7_5.py`:
the record-level UN Comtrade extract `Greece__Mar_Total_TradeData_Balance_Sheet.xlsx`
(rebuildable from https://comtradeplus.un.org/ as described in `data/README.md`; the
aggregated series it yields are registered in workbook Sheet 1) and a copy of
`data/v5_config.json`. Then:

```
python code/orchestrator_v7_5.py
```

The orchestrator prints a six-stage progress log — (1) PE Armington projections,
(2) sensitivity grid and PE Monte Carlo, (3) XGBoost training, scenario prediction and ML
Monte Carlo, (4) Excel workbook writer, (5) diagnostic charts, (6) audit log — and writes to
an `outputs/` folder next to the script:

- `EU_Mercosur_Greece_Analysis_v7.xlsx` — the complete 15-sheet analysis workbook
- `charts/01…07_*.png` — diagnostic charts
- `orchestrator_run.json` — full configuration and per-stage audit log

Step 2 run on that output must again end in `ALL CHECKS PASSED`.

## Applying the toolkit to another trade relationship

**Step 1 — fill the two input templates** (full field documentation in `templates/README.md`):

1. `templates/TradeData_template.xlsx` — one row per (year × flow × partner × HS-2 chapter)
   trade record, e.g. from UN Comtrade. Save it as
   `Greece__Mar_Total_TradeData_Balance_Sheet.xlsx` next to `code/orchestrator_v7_5.py`
   (the filename the script searches for; rename it in `CONFIG["globals"]["trade_data_path"]`
   if you prefer).
2. `templates/v5_config_template.json` — one entry per HS-2 chapter: MFN tariffs on both
   sides, phase-out lengths, baseline flows, counterfactual growth rates, elasticities and
   non-tariff wedge magnitudes, each with its source citation. Save it as `v5_config.json`
   next to `code/orchestrator_v7_5.py`. Every field, including `sensitivity`, must be present
   for every chapter, and every baseline, growth and capacity value must be filled: the
   workbook writer divides by the projected totals and does not run on an all-zero file.

**Step 2 — edit the country-specific constants in `code/orchestrator_v7_5.py`.** The
templates carry the trade panel and the chapter parameters; the following are hard-coded and
must be changed for another country pair:

- `CONFIG["globals"]`: agreement entry years (2026 and 2027), baseline year (2023), the
  economy-wide agri share (0.183), the Farm-to-Fork and labor drags and their Monte Carlo
  priors.
- `OECD_ENV_TAX`: the importer-side environmental-tax addon (+0.5 pp on HS 23 from a 2.4 pp
  burden gap at 20% pass-through). The config `wedge` field expresses exporter-borne
  magnitudes only; an importer-side burden in any other chapter has to be added here.
- `CONFIG["cap_weights"]`: the per-chapter capacity weights κ(c).
- `CONFIG["trq_caps"]`: the quota caps at the notional 2.4% national share.
- `CONFIG["ml"]`: the reporting country's and the partners' GDP series, partner distances,
  the EU MFN proxy table and phase years; the `AGRI_HS` set and the capacity factors 0.95 and
  0.90 in `run_xgboost_pipeline`.
- Fixed years in the ML pipeline: the holdout year is 2024 (`train_xgboost`), the trend
  origin is 2014 (`build_features`) and predictions start from the 2024 cross-section
  (`run_xgboost_pipeline`). A panel ending in another year needs these changed, otherwise the
  holdout set is empty.
- Sheets 12, 14 and 15 of the workbook writer (literature anchors, Greek structural
  indicators, discussion) contain Greek narrative content; delete or replace them.

The figure scripts also embed values from the published analysis as plotting literals:
the gain-importance bars in Fig. 2 panel (b), the cereal areas in Fig. 4, the capacity
classes and designation counts in Fig. 6, the quota volumes and prices in Fig. 5 and Fig. S7,
and most annotation text. Search the scripts for `vals =`, `area =`, `CAP =`, `GI =` and
`QUOTA =` and replace them with your own values.

**Step 3 — run the pipeline and draw the figures** exactly as in *Verifying the published
results*, steps 4 then 3, pointing the figure and sensitivity scripts at
`code/outputs/EU_Mercosur_Greece_Analysis_v7.xlsx`. `build_workbook_v7_5.py` verifies against
the Greek published values and will report failures on any other analysis; use it only as an
example of a verification harness.

**Alternative entry point:** if you prefer to parameterize the model in Excel rather than run
the orchestrator, fill the input sheets of `templates/Analysis_Workbook_template.xlsx`
(sheets 2, 3, 4, 5 and 13 — the template's note sheet explains which) and point the figure
and sensitivity scripts at that file.

## The machine-learning companion

Module 3 of the orchestrator trains a gradient-boosted regression (XGBoost 2.x) on the
bilateral panel: 11 features (one- and two-year lagged trade values, lagged growth,
log distance, log GDPs of both partners, a tariff proxy, direction indicator, year trend,
and chapter/partner identifiers), fixed hyperparameters (200 trees, depth 5, learning
rate 0.08, fixed seed), the year 2024 held out for evaluation (RMSE reported), and direct
horizon-specific prediction with policy and macro features shifted to the target year. Its
role is to supply a persistence baseline that validates the PE counterfactual — where tariffs
do not vary within chapters over the estimation window, the ML model cannot identify a tariff
response, so uncertainty about the policy effect is carried by the import-elasticity
sensitivity instead. Gain importance, holdout metrics and Monte Carlo intervals are written to
the workbook's Monte Carlo and ML-vs-PE sheets and to the diagnostic charts.

## Methodological notes

- **Tariff shock specification.** The engine expresses the tariff shock as the
  proportional change in the tariff-inclusive price, π = (1+τ₁)/(1+τ₀) − 1 (the standard
  WITS/SMART form). Expressing it as the proportional change in the tariff *rate*
  (−Δτ/τ₀) overstates the import response — by roughly 21× at a 5% MFN — and must not be
  used; the code headers and `data/Tariff_Specification_Correction.xlsx` document this
  correction.
- **Incidence-signed wedges.** Non-tariff cost channels are signed by who bears the cost:
  exporter-borne compliance costs dampen the import response (negative), importer-side
  producer burdens amplify it (positive). Enter wedge magnitudes as positive decimals in
  the config; the code signs them at load time.
- **Determinism.** Both Monte Carlos and the XGBoost model run on a fixed seed; results
  are reproducible conditional on the seed, the configuration and the XGBoost 2.x series.
  The orchestrator's PE Monte Carlo (NumPy generator) and the verification script's
  (Python `random`) give identical 95% intervals and means that differ in the first decimal
  (145.8 versus 145.7 million); the article reports the latter.

## Data sources used in the published analysis

UN Comtrade (bilateral HS-2 flows; https://comtradeplus.un.org/), Eurostat COMEXT
(cross-validation), the EU-Mercosur Agreement technical annexes (tariff schedules,
phase-ins, TRQs), ELSTAT 2020 Agricultural–Livestock Census, FADN Farm Economy Focus,
and OECD environmental-tax indicators. Full parameter provenance is reported in the
article's supplementary appendix (Table A.2) and in workbook Sheet 12.

## Created by
Dr. Zainab Ashkanani
Email: Ashkanani@saaah.co

## License and citation

Code is released under the MIT License (see `LICENSE`). If you use this toolkit, please
cite the article (see `CITATION.cff`).
