# Replication package, version 2.0.0

**Article:** *Protest without Exposure: Why Greece Backed the EU-Mercosur Agreement* (submitted to *JCMS: Journal of Common Market Studies*).

**Authors:** Konstantinos Pappas, Zainab Ashkanani and Rashed Albatayneh (Texas A&M University).

**License:** MIT.

This package reproduces every number, table and figure in the article and its Online Appendix. It replaces version 1.0.0, which implemented a chapter-average tariff specification. That specification applied average chapter duties to lines that enter the EU duty-free, and it projected a Year 10 widening of 146.0 million US dollars. Online Appendix E of the article documents the difference, and `pe_v8.py` reproduces the version 1.0.0 result exactly from `inputs/v5_config.json`.

## What the model does

The model is a partial-equilibrium Armington projection of Greek trade with the four Mercosur parties and Bolivia over the 2026 to 2036 phase-in of the EU-Mercosur Interim Trade Agreement. It covers 44 HS-2 chapters: the 43 chapters in which bilateral trade exceeded one million US dollars in some year from 2014 to 2024, plus dairy.

Each traded HS-6 line receives its own tariff treatment:

- **Tariff paths.** The base rate and staging category come from the agreement's schedules (OJ L 2026/184, Appendices 2-A-1 and 2-A-2). Cuts start at entry into force, and Year 0 runs from 1 May 2026 and is pro-rated.
- **Traded-row duties.** For the 21 lines that carry the result, the duty on the traded CN-8 row is taken from TARIC and valued at Greek unit values (`inputs/tariff_check_lines.csv`).
- **Tariff-rate quotas.** These are treated at the margin against a notional Greek slice of 2.4% of each EU quota.

Chapter price changes are baseline-weighted sums of the line-level changes. Imports respond with an elasticity of −3.5 and exports with −2.5.

## Layout

| Path | Contents |
|---|---|
| `model/build_v81.py` | Builds the line-level configuration from the inputs (spike screen, growth rule, tariff paths, quota slices) |
| `model/pe_v81.py` | Projection engine with marginal quota logic |
| `model/pe_v8.py` | The earlier engine, used for the chapter-average comparison |
| `model/run_v81.py` | Central case, year profile, sensitivities, grid and Monte Carlo; writes `results_v81.json` and `v81_config.json` |
| `model/make_figs_v81.py` | Figures 1 to 3 of the article |
| `model/make_appendix_v2.py` | Online Appendix Tables A.1 to A.8 (markup) and Figures S1 to S4 |
| `model/inputs/comtrade_hs6/` | UN Comtrade HS-6 records, Greece with each partner, 2022 to 2024, imports (M) and exports (X) |
| `model/inputs/Greece__Mar_Total_TradeData_Balance_Sheet.xlsx` | UN Comtrade HS-2 panel, 2014 to 2024 |
| `model/inputs/wits_rates.json` | WITS TRAINS MFN ad valorem equivalents, 2023 |
| `model/inputs/schedules/` | The EU and Mercosur schedules parsed from the Official Journal, the staging legend and the quota summary |
| `model/inputs/tariff_check_lines.csv` | Traded CN-8 duties checked against TARIC |
| `model/inputs/lines44.json` | HS-6 line list for the 44 chapters |
| `model/inputs/v5_config.json` | Version 1.0.0 chapter-average inputs, for the comparison in Online Appendix E |
| `outputs/` | Archived results (`results_v81.json`, `v81_config.json`) and figures (PNG at 300 dpi, EPS) |
| `provenance/` | Scripts used to retrieve and parse the inputs (Comtrade and WITS pulls, schedule parsing); they keep the paths used during construction and are not needed to reproduce the results |

## How to run

Use Python 3.11 or later with the packages in `requirements.txt`.

```
cd model
python run_v81.py
python make_figs_v81.py ../outputs/figures
python make_appendix_v2.py appendix_tables.txt ../outputs/figures
```

`run_v81.py` prints the central case (Year 10 widening of 16.30 million US dollars: additional imports 59.72 million, capacity-adjusted additional exports 43.42 million) and the sensitivities reported in Online Appendix Table A.7. The Monte Carlo uses a fixed seed (42), so a rerun reproduces `outputs/results_v81.json` exactly; this was checked on 29 September 2026. A full run takes several minutes.

## Key results

| Quantity | Value |
|---|---|
| Baseline Greek imports from Mercosur, 2022 to 2024 mean | 657.7 million US dollars |
| Share already duty-free at MFN | 83.0% |
| Year 10 widening of the bilateral balance, central case | 16.3 million US dollars |
| Peak widening | 18.0 million (Year 7) |
| Import elasticity −2.0 / −8.0 | −9.3 / 91.0 million |
| Largest adverse chapters | Tobacco 23.2, fish 14.2, fruit juices 7.7 |
| Largest favorable chapters | Fruit (kiwifruit) −13.7, pharmaceuticals −9.8, optics −7.4 |

## Data sources

- UN Comtrade, public API (reporter 300, partners 32, 68, 76, 600 and 858).
- World Bank WITS TRAINS.
- Official Journal of the European Union, L 2026/184.
- TARIC.

All are public. Accessed 28 and 29 September 2026.
