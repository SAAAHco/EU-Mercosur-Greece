# Input templates

Three blank templates. Fill the first two to run the full pipeline
(`code/orchestrator_v7_5.py`); the third is an alternative entry point for users who
prefer to parameterize in Excel and run only the figure/sensitivity scripts.

---

## 1. `TradeData_template.xlsx` — bilateral trade records

One row per **(year × flow × partner × HS-2 chapter)** record on `Sheet1`:

| Column | Meaning |
|---|---|
| `refYear` | Calendar year (a continuous annual panel; the published analysis uses 2014–2024) |
| `flowCode` | `M` = imports of the reporting country, `X` = its exports |
| `partnerISO` | Partner ISO3 code (e.g. `ARG`, `BOL`, `BRA`, `PRY`, `URY`) |
| `partnerDesc` | Partner name |
| `cmdCode` | Two-digit HS chapter code |
| `cmdDesc` | Chapter description |
| `primaryValue` | Trade value, current US dollars |

Typical source: UN Comtrade annual HS-2 bilateral flows. The machine-learning companion
trains on all years but the last and evaluates on the last (holdout) year, so keep the
panel continuous.

**Save as:** `Greece__Mar_Total_TradeData_Balance_Sheet.xlsx`, placed next to
`orchestrator_v7_5.py` (that filename is what the script searches for; it looks in the
script's folder, `./data`, `./input`, the parent folder, and the working directory).

## 2. `v5_config_template.json` — chapter parameter register

The master parameter table driving the PE engine: one JSON object per HS-2 chapter under
the `"chapters"` key. The template ships with the 25 chapter identities used in the
published analysis (all values blank) and a `_field_guide` explaining every field:

| Field | Meaning |
|---|---|
| `hs`, `name` | Chapter identity |
| `eu_mfn`, `target_eu`, `phase_eu`, `reduction_type` | Import-side tariff: initial MFN rate (decimal), end-of-phase rate, phase-out years, `"Linear"` or `"None"` (excluded) |
| `mer_mfn`, `phase_mer` | Partner-side tariff on your exports and its phase-out length (0 = no reduction) |
| `baseline_imp`, `baseline_exp` | Baseline annual flows in USD (e.g. mean of the last three observed years) |
| `cagr_imp`, `cagr_exp` | Counterfactual growth rates per year (winsorize thin-base chapters) |
| `imp_elast`, `exp_elast` | Armington elasticities (published central values −3.5 / −2.5 are pre-filled) |
| `wedge` | Non-tariff cost-wedge magnitude as a decimal (1.0 pp → `0.010`), entered **positive**; the code signs it by incidence at load time |
| `tier`, `citation` | Provenance notes: calibration source for the wedge and for the chapter parameterization |

Document a source for every calibrated value — the workbook writer carries these
citations into the output.

**Save as:** `v5_config.json`, next to `orchestrator_v7_5.py`.

## 3. `Analysis_Workbook_template.xlsx` — analysis-workbook structure

The 16-sheet workbook layout that `make_figures_v8.py`, `elasticity_sensitivity.py` and
`build_workbook_v7_5.py` read. **You normally do not fill this by hand** — running the
orchestrator generates the complete workbook. Fill it manually only if you want to skip
the orchestrator: enter your parameters and counterfactual/projected flows in sheets
`2-Tariff Parameters`, `3-Product Baselines`, `4-Scenario S1 (2026)`,
`5-Scenario S2 (2027)` and `13-Per-Chapter Adjustments`, keeping sheet names and
row/column positions exactly as shipped (the scripts read fixed positions; chapter rows
sit at rows 6–30 on the scenario sheets). The first sheet of the template explains the
same in place. Labels, headers and source-citation examples are kept so you can see how
each sheet is meant to be documented; every value cell is blank.
