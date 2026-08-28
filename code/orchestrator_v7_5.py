#!/usr/bin/env python3
"""
EU-Mercosur Agreement Impact on Greek Agriculture
ORCHESTRATOR: Single-source-of-truth Python pipeline (v7.5).

v7.5 CHANGES vs v7.4 (CRITICAL SPECIFICATION FIX; all PE results change):
  - proj_imports() and proj_exports() previously computed the policy shock as the
    proportional change in the TARIFF RATE, -(tau0-tau1)/tau0, which reads the
    removal of a 5% tariff as a 100% buyer-price cut and, at elasticity -3.5,
    overstates the import response by roughly 21x per chapter (23.8x on the
    headline). Both functions now use the proportional change in the
    TARIFF-INCLUSIVE PRICE, (1+tau1)/(1+tau0)-1, the standard WITS/SMART form.
  - Year 10 headline widening moves from 3,513.4M to 146.0M dollars; additional
    imports from 4,011.5M to 197.1M; capacity-adjusted exports from 498.1M to
    51.1M. Chapter ranking and the 92% feed+coffee concentration are unchanged.
  - ML companion, counterfactual baselines, wedge calibration, capacity weights
    and all input sheets are unaffected.

Runs the full analysis end-to-end and writes a fresh Excel workbook:
  1. Read Greek bilateral trade data (UN Comtrade)
  2. PE Armington projections (25 HS chapters, 2 scenarios, 11 horizons)
  3. TRQ caps applied via MIN() logic for HS 04 dairy
  4. XGBoost ML model with F2F-adjusted exports
  5. Monte Carlo 95% CIs (1,000 draws, GDP and phase-out perturbations)
  6. 3x3 sensitivity grid (wedge multiplier x F2F drag)
  7. Excel writer producing 15-sheet workbook

v7 CHANGES vs v6:
  v7.3 patch (specification ladder, headline unchanged):
  - Sheet 13 comparison block rebuilt as a one-change-per-step ladder with an exact
    bridge: economy-wide benchmark (bias check) -> Tier 1 (bilateral agricultural
    base, F2F only) -> Tier 2 (+ labour drag) -> Tier 3 (census weights), with the
    order-interaction term reported. Tier 3 and every downstream number unchanged.
  v7.2 patch (wedge respecification, PE results change by about -2.2%):
  - Wedge channels signed by cost incidence: exporter-borne EUDR/CBAM/EC-2014
    channels dampen imports; the Greek env-tax addon keeps its amplifying sign.
    Headline Tier 3 Y10 widening moves from 3,591.3M to 3,513.4M dollars.
  - Sensitivity-grid labels updated to the signed semantics.
  v7.1 patch (audit fixes, no change to PE results):
  - ML Monte Carlo now projects partner GDP at 2%/yr, matching the deterministic
    scenario, so the MC brackets the point estimate. Requires ML re-run.
  - Sheet 9 header corrected from "Tier 2" to "Tier 3" (code already ran Tier 3).
  - labor_drag central raised from 0.05 to 0.07, anchored on ELSTAT 2020 Census Table 43
    (seasonal workers = 12.8% of total workdays, 57% of non-family headcount).
  - labor_drag Monte Carlo distribution widened from (0.00, 0.05, 0.10) to (0.03, 0.07, 0.12).
  - Six new literature anchors added (ELSTAT employment, ELSTAT holdings/UAA, ELSTAT
    livestock, FADN 2019-2023, OECD environmental tax, EEA Greece country profile 2025).
  - Sheet 10 number format bug fixed (was "$#,##0," which hid three trailing zeros).
  - Sheet 13 gains ELSTAT UAA-share and validation columns (J, K).
  - New Sheet 14: Greek Structural Indicators with 9 sections (A-I): ELSTAT Census,
    FADN time series, labor_drag justification, OECD env-tax burden, EEA 2025 indicators.
  - New Sheet 15: Discussion (effects, contributions, current situation, 6 future directions),
    synthesizes model output with Greek structural data.

USAGE:
  python3 orchestrator.py
  Outputs: /mnt/user-data/outputs/EU_Mercosur_Greece_Analysis_v7.xlsx
           /mnt/user-data/outputs/orchestrator_run.json (audit log)

DEPENDENCIES:
  pandas, numpy, openpyxl, xgboost, scikit-learn

DATA SOURCES:
  /mnt/project/Greece__Mar_Total_TradeData_Balance_Sheet.xlsx (UN Comtrade)
  ELSTAT 2020 Agricultural-Livestock Census (Tables 43, B01, 7, 9, 10, 17, 22, 27, E05, 35A)
  FADN Farm Economy Focus, Greece 2019-2023 (DG AGRI dashboard)
  OECD Environmental tax indicator, % of GDP, 2022
  EEA Europe's Environment 2025: Greece country profile (28 September 2025)
  Tariff parameters and wedge calibration: documented inline in CONFIG.
"""

import os
import sys
import json
import copy
import warnings
from datetime import datetime
from pathlib import Path

import numpy as np
import pandas as pd
import openpyxl
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

import xgboost as xgb
from sklearn.metrics import mean_squared_error, mean_absolute_percentage_error

warnings.filterwarnings("ignore")


# ====================================================================
# v7 ADDITIONS: ELSTAT 2020 Census + FADN 2019-2023 data
# These constants teach the model about Greece. They are referenced
# from CONFIG (labor_drag derivation) and rendered into Sheets 14, 15.
# ====================================================================

ELSTAT_2020 = {
    "source": "Hellenic Statistical Authority (ELSTAT) 2020 Agricultural-Livestock Census",
    # Table 43: Employment
    "employment": {
        "total_persons": 2_106_534,
        "total_holdings": 463_320,
        "total_workdays": 97_173_709,
        "holders_household": {"holdings": 462_428, "persons": 722_973, "workdays": 76_353_132},
        "permanent_workers": {"holdings": 14_286, "persons": 27_350, "workdays": 5_710_760},
        "seasonal_workers": {"holdings": 215_948, "persons": 794_811, "workdays": 12_443_440},
        "mutual_help":      {"holdings": 66_122,  "persons": 141_915, "workdays": 1_747_066},
        "other_workers":    {"holdings": 203_418, "persons": 419_485, "workdays": 919_311},
        # Derived shares (denominators: total persons = 2,106,534;
        # non-family headcount = perm+seas+mutual+other = 1,383,561; total workdays = 97,173,709)
        "seasonal_pct_of_paid_headcount": 0.574,  # 794,811 / 1,383,561 (non-family denominator)
        "seasonal_pct_of_total_workdays": 0.128,  # 12,443,440 / 97,173,709
        "family_pct_of_total_workdays":   0.786,  # 76,353,132 / 97,173,709
    },
    # Table B01: Holdings & UAA
    "holdings_uaa": {
        "total_holdings": 530_679,
        "holdings_with_uaa": 525_284,
        "uaa_hectares": 2_822_886,
    },
    # Tables 7, 9, 10, 17, 22, 27: UAA by crop category (hectares)
    # Note: Per-crop areas come from Tables 7/9/10/17/22/27 which are above-threshold-only
    # (per EU Reg 2018/1091). Total UAA = 2,714,266 ha on that basis (vs 2,822,886 in B01).
    # Section C uses B01 denominator (2,822,886) for headline shares, which slightly understates
    # crop shares vs the within-table-7 basis but stays consistent with the official UAA total.
    "uaa_by_crop_ha": {
        "arable_total":              1_308_939,
        "cereals_total":             585_989,
        "wheat_durum":               231_372,
        "wheat_soft":                88_892,
        "barley":                    93_662,
        "maize":                     43_344,
        "dry_pulses":                23_387,
        "potatoes":                  10_800,
        "tree_crops_total":          730_447,   # Table 7 basis
        "tree_crops_total_t17":      734_533,   # Table 17 basis (slightly different scope)
        "olives_total":              584_242,
        "olives_for_oil":            522_069,
        "olives_table":              62_173,
        "fruit_trees_temperate":     64_205,   # Table 17 c11/12 (renamed from generic "fruit_trees")
        "fruit_trees_tropical":      13_942,   # Table 17 c13/14 (previously mislabeled as "other_tree_crops")
        "nuts":                      33_097,   # Table 17 c15/16 (almonds, etc.) — NEW v7 addition
        "citrus":                    34_559,   # Table 17 c17/18 (oranges etc.) — NEW v7 addition
        "tree_crops_other_perennials": 4_488,  # Table 17 c19/20 — NEW v7 addition
        "vineyards_total":           58_754,
        "vineyards_pdo":             9_310,
        "vineyards_pgi":             6_070,
        "vineyards_other_wine":      19_358,
        "vineyards_table_grapes":    10_859,
        "vineyards_raisins":         13_157,
        "irrigated_total":           924_103,
    },
    # Tables E05, 35A: Livestock
    "livestock": {
        "bovine": {"holdings": 10_865, "animals": 624_397},
        "sheep":  {"holdings": 56_761, "animals": 7_721_799},
        "goats":  {"holdings": 36_978, "animals": 3_149_008},
        "pigs":   {"holdings": 5_906,  "animals": 742_963},
    },
}

FADN_2019_2023 = {
    "source": "EU Farm Accountancy Data Network (FADN), Farm Economy Focus, Greece, IFS 2020 reference, all farm types. Data updated 2026-01-28.",
    # Per representative farm, EUR (corrected from high-resolution FADN dashboard read)
    "per_farm_eur": [
        # year, total_output, intermediate_cons, balance_subs, gross_farm_income, farm_net_va, total_subs_excl_inv, total_direct_pay
        (2019, 23425, 12834, 6749, 17340, 14030, 7069, 5522),
        (2020, 23927, 12661, 7061, 18327, 15019, 7367, 5135),
        (2021, 27647, 14330, 6083, 19400, 16090, 6425, 4938),
        (2022, 33229, 16268, 5723, 22684, 19314, 6209, 4585),
        (2023, 32929, 16910, 4957, 20976, 17803, 5461, 4376),
    ],
    # Population-level totals
    "scale": [
        # year, farms_represented, pop_farms, uaa_ha, livestock_units, awu, std_output_eur_bn
        (2019, 268322, 525000, 2810000, 1960000, 568400, 7.74),
        (2020, 268036, 525000, 2810000, 1960000, 568400, 7.74),
        (2021, 266970, 525000, 2810000, 1960000, 568400, 7.74),
        (2022, 265348, 525000, 2810000, 1870000, 568400, 7.74),
        (2023, 266966, 523900, 2800000, 1870000, 565400, 7.60),
    ],
}

OECD_ENV_TAX = {
    "source": "OECD Environmental tax indicator (% of GDP, 2022), https://www.oecd.org/en/data/indicators/environmental-tax.html",
    "greece_pct_gdp_2022": 0.036,                # Greece highest on the OECD dashboard
    # NOTE: The OECD dashboard PDF shows a highlighted "OECD" bar at ~1.2%, but per Web of Science
    # follow-up: the highlighted bar reflects the AVERAGE OF THE SELECTED (filtered) countries on
    # the dashboard view, not the full OECD-wide average. The full OECD-wide environmentally-related
    # tax revenue average is ~2.1-2.2% of GDP per OECD/CIAT/ECLAC/IDB Revenue Statistics LAC 2021
    # and per OECD WKP(2023)23. EU27 average was 2.2% of GDP in 2020 (IMF/OECD).
    "oecd_full_avg_pct_gdp_recent": 0.021,       # 2.1% - real OECD-wide average per literature, not the dashboard "OECD" bar
    "eu27_avg_pct_gdp_2020": 0.022,              # 2.2% of GDP in 2020 (IMF and OECD WKP-2023-23)
    "greece_pct_gdp_2020_via_wkp": 0.038,        # Greece 3.8% in 2020 per OECD WKP(2023)23 (highest in EU)
    "ireland_pct_gdp_2020_via_wkp": 0.012,       # Ireland 1.2% in 2020 (lowest in EU) - reference low
    "greece_rank_in_oecd": 1,                    # Highest in OECD per dashboard
    # LAC regional benchmark for Mercosur proxy. Mercosur countries are not OECD members
    # so are not on the OECD dashboard. The most defensible Mercosur-region proxy is the
    # LAC environmentally-related tax revenue average reported by OECD/CIAT/ECLAC/IDB
    # Revenue Statistics LAC 2021.
    "lac_regional_avg_pct_gdp_2019": 0.012,      # 1.2% of GDP, 25 LAC countries with data
    "lac_source": ("OECD/CIAT/ECLAC/IDB Revenue Statistics in Latin America and the Caribbean 2021. "
                   "Environmentally-related tax revenues averaged 1.2% of GDP in 2019 across the 25 LAC countries "
                   "with data, below the OECD average of 2.1%. For Brazil specifically, the OECD Environmental "
                   "Performance Review of Brazil (2015) reports green taxes 'less than 2% of total tax revenue', "
                   "which translates to roughly 0.6% of GDP at Brazil's ~32% tax-to-GDP ratio."),
    "lac_avg_caveat": ("LAC average (1.2%) is the proxy for the Mercosur cost-asymmetry calculation. Country-specific "
                       "values for Brazil, Argentina, Uruguay, Paraguay are not on the OECD dashboard and were not "
                       "found in primary sources during this build. The 1.2% figure is from the official LAC Revenue "
                       "Statistics 2021 publication."),
    # Implied wedge: Greek producer environmental-tax cost burden minus LAC regional average
    # 3.6% of GDP for Greece vs 1.2% for LAC mean = 2.4 percentage points of GDP-equivalent burden
    # Greek-vs-OECD-full gap is only 3.6 - 2.1 = 1.5pp (smaller); Greek-vs-LAC gap is 2.4pp.
    # The LAC gap is the relevant one for Mercosur cost asymmetry.
    "implied_greece_vs_lac_gap_pp_gdp": 0.024,    # 3.6% - 1.2% = 2.4 pp
    "implied_greece_vs_oecd_gap_pp_gdp": 0.015,   # 3.6% - 2.1% = 1.5 pp (for reference, smaller)
    "applies_to_chapters": ["23"],                # HS 23 (animal feed processing). HS 12 (oilseeds) was considered but has eu_mfn=0, so the wedge multiplies tariff_change=0 and has no calculated effect. HS 12 omitted to avoid spurious "ANCHORS" claims.
    "wedge_addon_per_chapter": 0.005,             # Conservative: 0.5pp added to chap['wedge'] for HS 23. Partial pass-through (~20%) of the 2.4pp env-tax gap. Full pass-through would imply 0.7-1.2pp.
    "note": ("Greek producers face a structural environmental-tax burden 2.4 pp of GDP higher than the LAC regional "
             "average (most relevant Mercosur proxy). For energy- and transport-intensive Greek agri processing, this "
             "translates to an additional 0.5 pp cost wedge after partial pass-through. The orchestrator adds this at "
             "config-load time to chap['wedge'] for HS 23 (animal feed). HS 12 (oilseeds) was considered but has eu_mfn=0, "
             "so the wedge channel is mathematically inactive there; HS 12 is omitted from the applies_to list. Future "
             "candidates for v8: HS 20 (canning, eu_mfn=0.18, real channel) and HS 24 (tobacco drying, eu_mfn=0.10), "
             "both pending evidence on energy-cost pass-through ratios."),
}

EEA_GREECE_2025 = {
    "source": "European Environment Agency. 'Greece - Country profiles - Europe's environment 2025'. Published 28 September 2025. https://www.eea.europa.eu/en/europe-environment-2025/countries/greece",
    "key_indicators": {
        "area_under_organic_farming_pct":     0.172,   # 17.2% of UAA, "Acceleration needed"
        "organic_farming_status":             "Acceleration needed",
        "circular_material_use_rate_pct":     0.052,   # 5.2%, "Acceleration needed"
        "climate_economic_losses_eur_cap":    60.0,    # EUR per capita, "Wrong direction"
        "designated_terrestrial_protected_pct": 0.347, # 34.7%, "Achieved"
        "final_energy_consumption_idx_2005":  74.8,    # Index 2005=100, "On track"
        "renewable_share_final_energy_pct":   0.253,   # 25.3%, "Acceleration needed"
        "ghg_emissions_idx_1990_no_lulucf":   69.2,    # 30.8% reduction vs 1990 baseline
        "fossil_fuel_subsidies_pct_gdp_2023": 0.009,   # 0.9% of GDP, vs EU avg 0.7%
        "fossil_fuel_subsidies_eu_avg_pct_gdp_2023": 0.007,
        "good_ecological_surface_water_pct":  0.638,   # 63.8% of surface waters in good ecological condition
        "good_chemical_surface_water_pct":    0.886,   # 88.6%
        "eco_innovation_index_status":        "Below EU average but highest rate of increase in EU 2013-2022",
        "energy_poverty_status":              "Much higher than EU average; many households cannot keep homes adequately warm",
        "marine_protected_areas_pct":         0.183,   # 18.3% (target 30%)
    },
    "agri_specific_findings": {
        "primary_sector_value_yield_eur_per_1000m2": 110,  # EUR 110 / 1000 m² = EUR 1,100/ha average value yield
        "structural_constraints": [
            "Small farms predominate",
            "Segmented geography (mainland + small islands) limits scale",
            "Innovation deployment to farmers is very limited",
            "Limited farmer education schemes",
            "Knowledge transfer happens mainly through company-driven contract farming",
        ],
        "consumption_trends": [
            "Olive oil consumption declining, replaced by cheaper seed oils",
            "Increase in fast food and processed food consumption",
            "Plant-based diets emerging but still niche",
            "Reliance on processed/cheap food has risen with inflation",
        ],
        "policy_context": [
            "CAP renegotiation in early 2025",
            "Circular economy action plan 2021-2025 adopted",
            "Waste prevention programme 2021-2030 enacted",
            "Front-of-pack nutrition labelling under EU review",
        ],
    },
    "environmental_tax_trend": ("Share of environmental taxes in total Greek tax revenue has SIGNIFICANTLY increased "
                                "since 2013, contrary to the slightly negative trend in the EU average. This is the "
                                "EEA's narrative confirmation of the OECD indicator showing Greece at #1 OECD position."),
    "eco_innovation_note":     ("Greek Eco-Innovation Index score is still lower than the EU average, but the increase "
                                "between 2013 and 2022 is the highest in the whole EU. Combined with new laws making "
                                "green investment more likely and increasing employment in environmental goods and services, "
                                "this is a meaningful structural tailwind."),
}


# ====================================================================
# PATH RESOLUTION: find input files relative to this script's location.
# This makes the script portable. Just put orchestrator.py, v5_config.json,
# and the trade data file in the same folder, and run.
# ====================================================================

SCRIPT_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()


def find_file(filename, search_dirs=None):
    """Look for a file in the script directory, then in common subdirectories."""
    if search_dirs is None:
        search_dirs = [SCRIPT_DIR, SCRIPT_DIR / "data", SCRIPT_DIR / "input",
                       SCRIPT_DIR.parent, Path.cwd()]
    for d in search_dirs:
        candidate = d / filename
        if candidate.exists():
            return str(candidate)
    # Last resort: cloud paths (only valid when running in the original environment)
    cloud_paths = {
        "Greece__Mar_Total_TradeData_Balance_Sheet.xlsx": "/mnt/project/Greece__Mar_Total_TradeData_Balance_Sheet.xlsx",
        "v5_config.json": "/home/claude/v5_config.json",
    }
    if filename in cloud_paths and Path(cloud_paths[filename]).exists():
        return cloud_paths[filename]
    raise FileNotFoundError(
        f"Could not find '{filename}'. Put it next to orchestrator.py, "
        f"or in a 'data' subfolder. Searched: {[str(d) for d in search_dirs]}"
    )


def resolve_output_dir():
    """Output directory: 'outputs' next to the script, created if missing."""
    # Prefer the cloud location if it exists (original environment); else local
    cloud = Path("/mnt/user-data/outputs")
    if cloud.parent.exists():
        cloud.mkdir(exist_ok=True)
        return str(cloud)
    local = SCRIPT_DIR / "outputs"
    local.mkdir(exist_ok=True)
    return str(local)


# ====================================================================
# CONFIG: All tunable parameters live here.
# ====================================================================

CONFIG = {
    "globals": {
        "greek_agri_share": 0.183,      # Greek agri share of total exports (KEPE 2025/57). Used as fallback only; per-chapter cap_weights now drive the calculation.
        "f2f_drag": 0.10,                # Central F2F capacity reduction (USDA + Wageningen + Mandanas)
        "labor_drag": 0.07,              # Tier 3 (v7 update): Greek agri labor-shortage drag at Y10. Anchored on ELSTAT 2020 Agricultural-Livestock Census, Table 43: seasonal workers are 794,811 persons, 57% of non-family headcount and 12.8% of total workdays (Greece country total). Combined with Tzouramani 2020 (avg farm manager age 51.7 yrs) and Maro et al. 2025 (post-COVID labor decline). Raised from v6 value of 0.05 (Labrianidis & Sykas 2009 anchor) to 0.07. Applied to chapters with cap_weight > 0.
        "agreement_year_s1": 2026,       # S1: agreement enters 2026
        "agreement_year_s2": 2027,       # S2: agreement enters 2027
        "baseline_year": 2023,           # Baseline year for CAGR projection
        "import_elasticity": -3.5,       # Hertel et al. 2007 GTAP norm
        "export_elasticity": -2.5,       # Hertel et al. 2007 GTAP norm
        "mc_n_draws": 1000,              # Monte Carlo simulations
        "mc_seed": 42,                   # Reproducibility seed
        "trade_data_path": find_file("Greece__Mar_Total_TradeData_Balance_Sheet.xlsx"),
    },
    # Tier 3: Per-chapter Greek capacity weights. Applied to F2F drag and labor drag.
    # weight = 1.0 means chapter is fully agricultural and fully exposed to CAP eco-scheme conditionality.
    # weight = 0.0 means chapter is non-agricultural and not subject to F2F or Greek labor constraints.
    # Sourced from Kostas's email mapping HS chapters to CAP intervention codes (may_5_email_data.docx)
    # plus interpretation of which chapters represent Greek agri-food versus industrial output.
    "cap_weights": {
        "23": 0.80,  # Animal feed; agri-adjacent, partly processed
        "12": 0.70,  # Oil seeds; primary agri but limited Greek production
        "09": 0.40,  # Coffee/tea/spices; Greek produces aromatic herbs (CAP-supported) but not coffee
        "26": 0.00,  # Ores; non-agri
        "47": 0.00,  # Wood pulp; CAP forestry only, marginal
        "27": 0.00,  # Mineral fuels; non-agri
        "24": 1.00,  # Tobacco; full agri, CAP coupled support
        "08": 1.00,  # Fruits and nuts; full agri, GI-relevant
        "30": 0.00,  # Pharmaceuticals; non-agri
        "20": 1.00,  # Prepared vegetables/fruits/olives; full agri, GI Kalamata olives
        "38": 0.00,  # Chemicals; non-agri
        "84": 0.00,  # Machinery; non-agri
        "39": 0.00,  # Plastics; non-agri
        "76": 0.00,  # Aluminium; non-agri
        "85": 0.00,  # Electrical machinery; non-agri
        "88": 0.00,  # Aircraft; non-agri
        "82": 0.00,  # Tools/cutlery; non-agri
        "73": 0.00,  # Iron/steel articles; non-agri
        "99": 0.00,  # Commodities n.e.c.; ambiguous, set to 0
        "68": 0.00,  # Stone/cement articles; non-agri
        "25": 0.00,  # Salt/stone; non-agri
        "89": 0.00,  # Ships; non-agri
        "22": 0.90,  # Beverages/wine; nearly full agri, GI Samos/Nemea wines
        "04": 1.00,  # Dairy/eggs/honey; full agri, GI Feta/Graviera/Manouri
        "15": 1.00,  # Animal/vegetable fats and oils; full agri, GI Kalamata olive oil
    },
    "trq_caps": {
        # Greek-share TRQ caps in USD; applied=True engages MIN() in projection
        "04": {"cap_usd": 8_640_000, "coverage": "Cheese + Milk powder + Honey",
               "eu_combined_m": 360.0, "applied": True,
               "calc_note": "(30k cheese x $5k + 10k milk x $3k + 45k honey x $4k) tonnes x 2.4% Greek share"},
        "22": {"cap_usd": 9_360_000, "coverage": "Ethanol fuel + chemical",
               "eu_combined_m": 390.0, "applied": False,
               "reason": "TRQ only covers ethanol; HS 22 mostly wine/spirits at bilateral level"},
        "02": {"cap_usd": 22_680_000, "coverage": "Beef + Poultry",
               "eu_combined_m": 945.0, "applied": False,
               "reason": "HS 02 not in projection model; bilateral trade ~$8M"},
        "17": {"cap_usd": 1_728_000, "coverage": "Sugar refining",
               "eu_combined_m": 72.0, "applied": False,
               "reason": "HS 17 not in projection model; bilateral trade ~$1.7M"},
    },
    "sensitivity_grid": {
        "wedge_multipliers": [0.5, 1.0, 1.5, 3.0],
        "f2f_drags": [0.07, 0.10, 0.20],
        "note": ("Wedge multipliers expanded in Tier 2 to include 3.0x (= 3% wedge), "
                 "the empirical upper bound implied by KEPE 2025/57 observed 5.5pp Greek-EU "
                 "asymmetric cost burden divergence over 2019-2023, projected forward 10 years."),
    },
    "pe_monte_carlo": {
        "n_draws": 1000,
        "seed": 42,
        "wedge_mult_dist": "triangular",
        "wedge_mult_low": 0.5,    # compliance-study lower bound
        "wedge_mult_mode": 1.0,   # compliance-study central (Profundo, IISD, EC)
        "wedge_mult_high": 3.0,   # KEPE-implied empirical upper bound
        "f2f_drag_dist": "triangular",
        "f2f_drag_low": 0.05,     # compliance studies lower bound
        "f2f_drag_mode": 0.10,    # USDA/Wageningen/Mandanas central
        "f2f_drag_high": 0.20,    # Wageningen high scenario
        "labor_drag_dist": "triangular",  # Tier 3 (v7 update): separate labor constraint perturbation
        "labor_drag_low": 0.03,    # v7: non-zero lower bound. Structural seasonal-labor share (12.8% of workdays per ELSTAT 2020 Census) is empirically observed, not hypothetical.
        "labor_drag_mode": 0.07,   # v7: Central anchored on ELSTAT 2020 Census Table 43. Was 0.05 in v6.
        "labor_drag_high": 0.12,   # v7: Widened from 0.10 to reflect compounding effect of farm-manager aging (Tzouramani 2020: 51.7 yrs) on top of seasonal-labor vulnerability.
        "note": ("Monte Carlo over PE Armington parameters. Tier 3 adds labor_drag as a third "
                 "independent perturbation. Wedge multiplier ~ Triangular(0.5, 1.0, 3.0). F2F drag "
                 "~ Triangular(0.05, 0.10, 0.20). Labor drag ~ Triangular(0.03, 0.07, 0.12) anchored "
                 "on ELSTAT 2020 Agricultural-Livestock Census (v7 update from v6 0.00/0.05/0.10).")
    },
    "ml": {
        "distances_km": {"ARG": 11686, "BRA": 9800, "PRY": 10700, "URY": 11500, "BOL": 11200},
        "gdp_greece": {2014: 237, 2015: 196, 2016: 192, 2017: 203, 2018: 218, 2019: 207,
                       2020: 191, 2021: 218, 2022: 219, 2023: 244, 2024: 257},
        "gdp_partners": {
            "ARG": {2014: 526, 2015: 594, 2016: 557, 2017: 643, 2018: 519, 2019: 450,
                    2020: 386, 2021: 487, 2022: 631, 2023: 646, 2024: 633},
            "BRA": {2014: 2456, 2015: 1802, 2016: 1796, 2017: 2063, 2018: 1886, 2019: 1878,
                    2020: 1476, 2021: 1649, 2022: 1920, 2023: 2127, 2024: 2180},
            "PRY": {2014: 40, 2015: 36, 2016: 36, 2017: 39, 2018: 41, 2019: 38,
                    2020: 35, 2021: 39, 2022: 41, 2023: 43, 2024: 44},
            "URY": {2014: 57, 2015: 53, 2016: 52, 2017: 59, 2018: 60, 2019: 56,
                    2020: 53, 2021: 59, 2022: 71, 2023: 78, 2024: 81},
            "BOL": {2014: 33, 2015: 33, 2016: 34, 2017: 38, 2018: 40, 2019: 41,
                    2020: 37, 2021: 40, 2022: 44, 2023: 46, 2024: 48},
        },
        "eu_tariffs": {
            '02':0.59,'03':0.12,'04':0.45,'05':0.0,'07':0.12,'08':0.10,'09':0.03,'10':0.25,
            '11':0.15,'12':0.0,'13':0.05,'14':0.0,'15':0.08,'16':0.16,'17':0.35,'18':0.10,
            '19':0.12,'20':0.18,'21':0.10,'22':0.14,'23':0.05,'24':0.10,'25':0.02,'26':0.0,
            '27':0.03,'28':0.04,'29':0.05,'30':0.04,'31':0.03,'32':0.05,'33':0.04,'34':0.04,
            '35':0.06,'36':0.06,'37':0.04,'38':0.05,'39':0.06,'40':0.04,'41':0.0,'42':0.05,
            '43':0.03,'44':0.03,'46':0.03,'47':0.0,'48':0.04,'49':0.0,'51':0.03,'52':0.05,
            '53':0.03,'54':0.06,'55':0.06,'56':0.06,'57':0.08,'58':0.08,'59':0.06,'60':0.08,
            '61':0.12,'62':0.12,'63':0.10,'64':0.08,'65':0.05,'66':0.05,'68':0.03,'69':0.04,
            '70':0.05,'71':0.02,'72':0.02,'73':0.03,'74':0.04,'76':0.06,'78':0.03,'79':0.03,
            '81':0.03,'82':0.03,'83':0.03,'84':0.02,'85':0.03,'86':0.03,'87':0.06,'88':0.03,
            '89':0.02,'90':0.03,'91':0.04,'92':0.03,'93':0.03,'94':0.04,'95':0.04,'96':0.04,
            '97':0.0,'99':0.0,
        },
        "phase_years": {
            '02':6,'04':6,'07':10,'08':10,'09':4,'10':10,'12':0,'15':10,'16':6,'17':6,
            '20':10,'22':10,'23':7,'24':10,'25':4,'26':0,'27':7,'30':7,'38':7,'39':7,
            '47':0,'68':7,'73':7,'76':7,'82':7,'84':7,'85':7,'87':10,'88':7,'89':7,'99':0,
        },
    },
}

# Chapter parameters with sourced wedge calibration; loaded from extracted v5 config.
# This is the master parameter table that drives the PE Armington model.
CHAPTERS = json.load(open(find_file("v5_config.json")))["chapters"]


# ====================================================================
# v7 ADDITION: Apply env-tax wedge addon to affected chapters at load time
# ====================================================================
# For chapters in OECD_ENV_TAX['applies_to_chapters'] (HS 23, HS 12), add the
# wedge_addon_per_chapter (0.5pp) directly to chap['wedge']. This propagates
# automatically into PE Armington, sensitivity grid, and Monte Carlo without
# touching any downstream code.
def _apply_env_tax_wedge_addon():
    addon = OECD_ENV_TAX["wedge_addon_per_chapter"]
    affected = set(OECD_ENV_TAX["applies_to_chapters"])
    applied = []
    for chap in CHAPTERS:
        if chap["hs"] in affected:
            old_wedge = chap["wedge"]
            chap["wedge"] = old_wedge + addon
            # Update the tier source so Sheet 2 reflects the combined provenance
            old_tier = chap.get("tier", "")
            chap["tier"] = (old_tier + " + OECD env-tax v7").strip(" +")
            applied.append((chap["hs"], chap["name"], old_wedge, chap["wedge"]))
    return applied

_ENV_TAX_WEDGE_APPLIED = _apply_env_tax_wedge_addon()


# ====================================================================
# v7.2 ADDITION: incidence-consistent wedge signs
# ====================================================================
# The EUDR, CBAM and EC-2014 compliance channels are calibrated on sources that
# measure costs borne by Mercosur exporters (Profundo 2025; IISD and Frontier
# Economics 2024; EC 2014). Costs on the import supply chain dampen the import
# response, so these channels enter with a NEGATIVE sign. The environmental-tax
# addon measures a Greek-producer burden and keeps its POSITIVE (amplifying)
# sign. Net chapter wedge = addon_component - exporter_component; for HS 23 this
# is +0.005 - 0.010 = -0.005. The wedge multiplier now scales the magnitude of
# the net wedge, so 3.0x is a full-enforcement stress on the compliance
# channels, not an upper bound on a Greek-side burden. NOTE: the manuscript
# 3.0x anchor text (KEPE/Reziti Greek input-cost asymmetry) must be re-anchored
# to compliance-cost uncertainty (Profundo/IISD ranges); see revision notes.
def _apply_wedge_incidence_signs():
    addon = OECD_ENV_TAX["wedge_addon_per_chapter"]
    affected = set(OECD_ENV_TAX["applies_to_chapters"])
    applied = []
    for chap in CHAPTERS:
        if chap.get("wedge", 0) == 0:
            continue
        addon_part = addon if chap["hs"] in affected else 0.0
        exporter_part = chap["wedge"] - addon_part
        old = chap["wedge"]
        chap["wedge"] = addon_part - exporter_part
        chap["tier"] = (chap.get("tier", "") + " | signed by incidence v7.2").strip(" |")
        applied.append((chap["hs"], old, chap["wedge"]))
    return applied

_WEDGE_SIGNS_APPLIED = _apply_wedge_incidence_signs()
if _WEDGE_SIGNS_APPLIED:
    print("[v7.2] Signed wedges by incidence: " +
          ", ".join(f"HS {hs} ({o:+.3f}->{n:+.3f})" for hs, o, n in _WEDGE_SIGNS_APPLIED))
if _ENV_TAX_WEDGE_APPLIED:
    print(f"[v7] Applied env-tax wedge addon of +{OECD_ENV_TAX['wedge_addon_per_chapter']:.3f} "
          f"to {len(_ENV_TAX_WEDGE_APPLIED)} chapter(s): " +
          ", ".join(f"HS {hs} ({old:.3f}->{new:.3f})" for hs, _, old, new in _ENV_TAX_WEDGE_APPLIED))


# ====================================================================
# MODULE 1: PE ARMINGTON
# ====================================================================

def eu_tariff(chap, year_offset):
    """EU tariff phase-out, linear schedule."""
    if chap["reduction_type"] == "None":
        return chap["eu_mfn"]
    if year_offset >= chap["phase_eu"]:
        return chap["target_eu"]
    return chap["eu_mfn"] - (chap["eu_mfn"] - chap["target_eu"]) * year_offset / chap["phase_eu"]


def mer_tariff(chap, year_offset):
    """Mercosur tariff phase-out, linear schedule."""
    if chap["phase_mer"] == 0:
        return chap["mer_mfn"]
    if year_offset >= chap["phase_mer"]:
        return 0
    return chap["mer_mfn"] - chap["mer_mfn"] * year_offset / chap["phase_mer"]


def cf_imports(chap, year_offset, agr_yr, base_yr):
    return chap["baseline_imp"] * (1 + chap["cagr_imp"]) ** (agr_yr + year_offset - base_yr)


def cf_exports(chap, year_offset, agr_yr, base_yr):
    return chap["baseline_exp"] * (1 + chap["cagr_exp"]) ** (agr_yr + year_offset - base_yr)


def proj_imports(chap, year_offset, agr_yr, base_yr, trq_cap=None, wedge_mult=1.0):
    """Projected imports with cost-wedge and optional TRQ cap.

    Wedge ramps linearly with year_offset / phase_eu, saturating at 1. This matches
    the Excel formula: (1 + L * MIN(1, year/MAX(E, 1))). The wedge represents
    adoption costs that scale with cumulative trade displaced under the new regime.
    """
    cf = cf_imports(chap, year_offset, agr_yr, base_yr)
    cur = eu_tariff(chap, year_offset)
    initial = chap["eu_mfn"]
    tariff_change = (1 + cur) / (1 + initial) - 1  # v7.5 FIX: proportional change in the tariff-INCLUSIVE price (WITS/SMART form dM/M = eps*dtau/(1+tau0)); previous form -(initial-cur)/initial read a 5% removal as a 100% price cut
    effective_wedge = chap["wedge"] * wedge_mult
    wedge_factor = 1 + effective_wedge * min(1, year_offset / max(chap["phase_eu"], 1))
    proj = cf * (1 + chap["imp_elast"] * tariff_change * wedge_factor)
    if trq_cap is not None:
        proj = min(proj, trq_cap)
    return proj


def proj_exports(chap, year_offset, agr_yr, base_yr):
    cf = cf_exports(chap, year_offset, agr_yr, base_yr)
    cur = mer_tariff(chap, year_offset)
    initial = chap["mer_mfn"]
    tariff_change = (1 + cur) / (1 + initial) - 1  # v7.5 FIX: proportional change in the tariff-INCLUSIVE price (WITS/SMART form dM/M = eps*dtau/(1+tau0)); previous form -(initial-cur)/initial read a 5% removal as a 100% price cut
    return cf * (1 + chap["exp_elast"] * tariff_change)


def project_chapter_yearly(chap, agr_yr, base_yr, trq_caps, wedge_mult=1.0):
    """Compute projection for a single chapter across all years 0-10."""
    cap = None
    if chap["hs"] in trq_caps and trq_caps[chap["hs"]]["applied"]:
        cap = trq_caps[chap["hs"]]["cap_usd"]
    out = {"hs": chap["hs"], "name": chap["name"], "years": {}}
    for y in range(0, 11):
        out["years"][y] = {
            "cf_imp": cf_imports(chap, y, agr_yr, base_yr),
            "cf_exp": cf_exports(chap, y, agr_yr, base_yr),
            "proj_imp": proj_imports(chap, y, agr_yr, base_yr, cap, wedge_mult),
            "proj_exp": proj_exports(chap, y, agr_yr, base_yr),
            "eu_tariff": eu_tariff(chap, y),
            "mer_tariff": mer_tariff(chap, y),
        }
    return out


def run_pe_armington(agr_yr, base_yr, trq_caps, wedge_mult=1.0):
    """Full PE Armington pass for all 25 chapters."""
    return [project_chapter_yearly(c, agr_yr, base_yr, trq_caps, wedge_mult) for c in CHAPTERS]


# ====================================================================
# MODULE 2: SENSITIVITY GRID
# ====================================================================

def adj_exports_per_chapter(results, year_idx, f2f_drag, labor_drag, cap_weights):
    """Tier 3: Apply per-chapter Greek capacity adjustment to additional exports.

    For each chapter, adjusted additional exports = unadjusted * (1 - cap_weight * (f2f_drag + labor_drag * year_idx/10)).
    Labor drag ramps linearly from 0 at Y0 to full at Y10. F2F drag is applied at full strength
    (already reflects the by-Y10 magnitude per Beckman/Wageningen).
    Returns per-chapter adjustment data and total.
    """
    chapter_adjustments = []
    total_unadj = 0
    total_adj = 0
    for r in results:
        chap_export = r["years"][year_idx]["proj_exp"] - r["years"][year_idx]["cf_exp"]
        w = cap_weights.get(r["hs"], 0.0)
        labor_factor = labor_drag * (year_idx / 10.0) if year_idx > 0 else 0.0
        total_drag = w * (f2f_drag + labor_factor)
        chap_adj = chap_export * (1 - total_drag)
        chapter_adjustments.append({
            "hs": r["hs"], "name": r["name"], "cap_weight": w,
            "unadj_export": chap_export, "adj_export": chap_adj,
            "drag_pct": total_drag, "loss": chap_export - chap_adj,
        })
        total_unadj += chap_export
        total_adj += chap_adj
    return {
        "chapters": chapter_adjustments,
        "total_unadj": total_unadj,
        "total_adj": total_adj,
        "total_loss": total_unadj - total_adj,
        "effective_share": (total_unadj - total_adj) / max(total_unadj * (f2f_drag + labor_drag), 1e-9),
    }


def run_sensitivity_grid(agr_yr, base_yr, trq_caps, f2f_drags, wedge_mults, greek_agri_share,
                          cap_weights=None, labor_drag=0.0):
    """4x3 grid of wedge multiplier x F2F drag. Tier 3: now uses per-chapter cap_weights when provided."""
    grid = {}
    for w in wedge_mults:
        results = run_pe_armington(agr_yr, base_yr, trq_caps, wedge_mult=w)
        # Aggregate Y10 totals
        add_imp = sum(c["years"][10]["proj_imp"] - c["years"][10]["cf_imp"] for c in results)
        add_exp = sum(c["years"][10]["proj_exp"] - c["years"][10]["cf_exp"] for c in results)
        for f in f2f_drags:
            if cap_weights:
                # Tier 3: per-chapter adjustment
                adj_data = adj_exports_per_chapter(results, 10, f, labor_drag, cap_weights)
                adj_exp = adj_data["total_adj"]
            else:
                # Pre-Tier 3: uniform agri share adjustment
                adj_exp = add_exp * (1 - greek_agri_share * f)
            widening = add_imp - adj_exp
            grid[(w, f)] = {
                "wedge_mult": w,
                "f2f_drag": f,
                "add_imp_y10": add_imp,
                "add_exp_y10_unadj": add_exp,
                "add_exp_y10_adj": adj_exp,
                "bilateral_widening": widening,
            }
    return grid


def run_pe_monte_carlo(agr_yr, base_yr, trq_caps, mc_cfg, greek_agri_share,
                        cap_weights=None):
    """PE Armington Monte Carlo. Tier 3: now perturbs wedge, F2F, AND labor drag,
    with per-chapter cap_weights when provided."""
    rng = np.random.default_rng(mc_cfg["seed"])
    n = mc_cfg["n_draws"]

    y5 = {"add_imp": [], "add_exp_adj": [], "widening": []}
    y10 = {"add_imp": [], "add_exp_adj": [], "widening": []}
    draws_log = []

    has_labor = "labor_drag_low" in mc_cfg

    for i in range(n):
        wedge_mult = rng.triangular(mc_cfg["wedge_mult_low"], mc_cfg["wedge_mult_mode"],
                                     mc_cfg["wedge_mult_high"])
        f2f_drag = rng.triangular(mc_cfg["f2f_drag_low"], mc_cfg["f2f_drag_mode"],
                                   mc_cfg["f2f_drag_high"])
        if has_labor:
            labor_drag = rng.triangular(mc_cfg["labor_drag_low"], mc_cfg["labor_drag_mode"],
                                         mc_cfg["labor_drag_high"])
        else:
            labor_drag = 0.0

        results = run_pe_armington(agr_yr, base_yr, trq_caps, wedge_mult=wedge_mult)

        for yh, store in [(5, y5), (10, y10)]:
            ai = sum(r["years"][yh]["proj_imp"] - r["years"][yh]["cf_imp"] for r in results)
            if cap_weights:
                # Tier 3: per-chapter adjustment with both F2F and labor drag
                adj_data = adj_exports_per_chapter(results, yh, f2f_drag, labor_drag, cap_weights)
                ae_adj = adj_data["total_adj"]
            else:
                ae = sum(r["years"][yh]["proj_exp"] - r["years"][yh]["cf_exp"] for r in results)
                ae_adj = ae * (1 - greek_agri_share * f2f_drag * (yh / 10))
            widening = ai - ae_adj
            store["add_imp"].append(ai)
            store["add_exp_adj"].append(ae_adj)
            store["widening"].append(widening)

        if i < 50:
            draws_log.append({"i": i, "wedge_mult": wedge_mult, "f2f_drag": f2f_drag,
                              "labor_drag": labor_drag if has_labor else None})

    def _stats(arr):
        return {
            "mean": float(np.mean(arr)),
            "median": float(np.median(arr)),
            "std": float(np.std(arr)),
            "lo": float(np.percentile(arr, 2.5)),
            "hi": float(np.percentile(arr, 97.5)),
            "p10": float(np.percentile(arr, 10)),
            "p90": float(np.percentile(arr, 90)),
        }

    return {
        "n_draws": n,
        "seed": mc_cfg["seed"],
        "y5": {key: _stats(arr) for key, arr in y5.items()},
        "y10": {key: _stats(arr) for key, arr in y10.items()},
        "y10_widening_draws": [float(x) for x in y10["widening"]],
        "y10_add_imp_draws": [float(x) for x in y10["add_imp"]],
        "draws_sample": draws_log,
    }


# ====================================================================
# MODULE 3: XGBoost + Monte Carlo
# ====================================================================

def load_trade_data(path):
    """Load Greek bilateral trade data from UN Comtrade Excel."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Sheet1"]
    headers = [c.value for c in ws[1]]
    raw = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        rec = dict(zip(headers, row))
        val = rec.get("primaryValue") or 0
        raw.append({
            "year": int(rec["refYear"]),
            "flow": rec["flowCode"],
            "partner": rec["partnerISO"],
            "hs": str(rec["cmdCode"]).zfill(2),
            "value": float(val) if val else 0,
        })
    wb.close()
    df = pd.DataFrame(raw)
    df = df[df["value"] > 0].copy()
    return df


def build_features(df, gdp_greece, gdp_partners, distances):
    """Engineer features for XGBoost."""
    df = df.copy()
    df["ln_distance"] = df["partner"].map(distances).apply(lambda x: np.log(x) if x else 0)
    df["ln_gdp_greece"] = df["year"].map(gdp_greece).apply(lambda x: np.log(x) if x else 0)

    def _gdp_partner(row):
        partner = row["partner"]
        year = row["year"]
        if partner in gdp_partners and year in gdp_partners[partner]:
            return np.log(gdp_partners[partner][year])
        return 0
    df["ln_gdp_partner"] = df.apply(_gdp_partner, axis=1)
    df["hs_num"] = df["hs"].astype(int)
    df["partner_code"] = df["partner"].astype("category").cat.codes
    df["is_import"] = (df["flow"] == "M").astype(int)
    df["year_trend"] = df["year"] - 2014
    # Per-HS EU MFN tariff rates (matches phase3_ml_pipeline_v2.py)
    EU_TARIFFS = {
        '02':0.59,'03':0.12,'04':0.45,'05':0.0,'07':0.12,'08':0.10,'09':0.03,'10':0.25,
        '11':0.15,'12':0.0,'13':0.05,'14':0.0,'15':0.08,'16':0.16,'17':0.35,'18':0.10,
        '19':0.12,'20':0.18,'21':0.10,'22':0.14,'23':0.05,'24':0.10,'25':0.02,'26':0.0,
        '27':0.03,'28':0.04,'29':0.05,'30':0.04,'31':0.03,'32':0.05,'33':0.04,'34':0.04,
        '35':0.06,'36':0.06,'37':0.04,'38':0.05,'39':0.06,'40':0.04,'41':0.0,'42':0.05,
        '43':0.03,'44':0.03,'46':0.03,'47':0.0,'48':0.04,'49':0.0,'51':0.03,'52':0.05,
        '53':0.03,'54':0.06,'55':0.06,'56':0.06,'57':0.08,'58':0.08,'59':0.06,'60':0.08,
        '61':0.12,'62':0.12,'63':0.10,'64':0.08,'65':0.05,'66':0.05,'68':0.03,'69':0.04,
        '70':0.05,'71':0.02,'72':0.02,'73':0.03,'74':0.04,'76':0.06,'78':0.03,'79':0.03,
        '81':0.03,'82':0.03,'83':0.03,'84':0.02,'85':0.03,'86':0.03,'87':0.06,'88':0.03,
        '89':0.02,'90':0.03,'91':0.04,'92':0.03,'93':0.03,'94':0.04,'95':0.04,'96':0.04,
        '97':0.0,'99':0.0,
    }
    df["tariff"] = df["hs"].apply(lambda h: EU_TARIFFS.get(h, 0.03))

    # Lagged values (per partner-hs-flow group)
    df = df.sort_values(["partner", "hs", "flow", "year"])
    df["value_lag1"] = df.groupby(["partner", "hs", "flow"])["value"].shift(1).fillna(0)
    df["value_lag2"] = df.groupby(["partner", "hs", "flow"])["value"].shift(2).fillna(0)
    df["growth_lag1"] = (
        (df["value_lag1"] - df.groupby(["partner", "hs", "flow"])["value"].shift(2).fillna(0))
        / df.groupby(["partner", "hs", "flow"])["value"].shift(2).fillna(1).replace(0, 1)
    ).fillna(0).clip(-2, 5)
    return df


def train_xgboost(df_features, feature_cols, target_col="value"):
    """Train XGBoost with year-based holdout."""
    train = df_features[df_features["year"] < 2024].copy()
    test = df_features[df_features["year"] == 2024].copy()
    X_train = train[feature_cols]
    y_train = train[target_col]
    X_test = test[feature_cols]
    y_test = test[target_col]

    model = xgb.XGBRegressor(
        n_estimators=200, max_depth=5, learning_rate=0.08,
        random_state=42, n_jobs=2, verbosity=0,
    )
    model.fit(X_train, y_train)
    pred = model.predict(X_test)
    rmse = float(np.sqrt(mean_squared_error(y_test, pred)))
    mape = float(mean_absolute_percentage_error(y_test, pred) * 100)

    importance = model.feature_importances_
    feat_imp = sorted(
        [(feature_cols[i], float(importance[i])) for i in range(len(feature_cols))],
        key=lambda x: -x[1],
    )
    return model, {"rmse": rmse, "mape": mape, "feature_importance": feat_imp}


def predict_scenario(model, df_features, feature_cols, year_horizon, gdp_greece, gdp_partners,
                     phase_speed=1.0, agri_hs_for_f2f=None, capacity_factor=1.0,
                     ml_eu_tariffs=None, ml_phase_years=None):
    """Predict bilateral trade for a specific year horizon by adjusting features."""
    base_year = 2024
    target_year = base_year + year_horizon
    df_pred = df_features[df_features["year"] == base_year].copy()
    df_pred["year"] = target_year
    df_pred["year_trend"] = target_year - 2014

    # Project GDP forward (2% growth)
    df_pred["ln_gdp_greece"] = np.log(gdp_greece[base_year] * (1.02 ** year_horizon))
    df_pred["ln_gdp_partner"] = df_pred["partner"].apply(
        lambda p: np.log(gdp_partners.get(p, {}).get(base_year, 100) * (1.02 ** year_horizon))
    )

    # Project per-HS tariff phase-out
    def _tariff(h):
        c = (ml_eu_tariffs or {}).get(h, 0.03)
        p = (ml_phase_years or {}).get(h, 7)
        if c == 0 or p == 0:
            return c
        return c * max(0, 1 - min(year_horizon * phase_speed / p, 1.0))
    df_pred["tariff"] = df_pred["hs"].apply(_tariff)

    # Apply F2F capacity factor to Greek exports of agri chapters
    df_pred["_f2f_factor"] = 1.0
    if agri_hs_for_f2f and capacity_factor != 1.0:
        agri_export_mask = (df_pred["flow"] == "X") & (df_pred["hs"].isin(agri_hs_for_f2f))
        df_pred.loc[agri_export_mask, "_f2f_factor"] = capacity_factor

    pred = model.predict(df_pred[feature_cols]) * df_pred["_f2f_factor"].values
    return pred.sum(), df_pred


def run_xgboost_pipeline(trade_data_path, ml_cfg, mc_seed=42, mc_n_draws=1000):
    """Full XGBoost + Monte Carlo pipeline."""
    df = load_trade_data(trade_data_path)
    distances = ml_cfg["distances_km"]
    gdp_g = ml_cfg["gdp_greece"]
    gdp_p = ml_cfg["gdp_partners"]
    eu_tariffs = ml_cfg["eu_tariffs"]
    phase_years = ml_cfg["phase_years"]
    df_feat = build_features(df, gdp_g, gdp_p, distances)
    feature_cols = [
        "ln_distance", "value_lag2", "value_lag1", "year_trend", "growth_lag1",
        "is_import", "ln_gdp_partner", "hs_num", "tariff", "ln_gdp_greece", "partner_code",
    ]
    model, model_stats = train_xgboost(df_feat, feature_cols)

    AGRI_HS = {"01", "02", "04", "07", "08", "09", "10", "11", "12", "15",
               "16", "17", "18", "19", "20", "21", "22", "23", "24"}

    base_imp = df[(df["year"] == 2024) & (df["flow"] == "M")]["value"].sum()
    base_exp = df[(df["year"] == 2024) & (df["flow"] == "X")]["value"].sum()

    F2F_AT_YEAR = {0: 1.0, 5: 0.95, 10: 0.9}
    scenarios = []

    def _project_features(yh, phase_speed):
        """Build the prediction feature matrix for year horizon yh."""
        df_pred = df_feat[df_feat["year"] == 2024].copy()
        df_pred["year"] = 2024 + yh
        df_pred["year_trend"] = (2024 + yh) - 2014
        df_pred["ln_gdp_greece"] = np.log(gdp_g[2024] * (1.02 ** yh))
        df_pred["ln_gdp_partner"] = df_pred["partner"].apply(
            lambda p: np.log(gdp_p.get(p, {}).get(2024, 100) * (1.02 ** yh))
        )
        def _t(h):
            c = eu_tariffs.get(h, 0.03)
            p = phase_years.get(h, 7)
            if c == 0 or p == 0 or phase_speed == 0:
                return c
            return c * max(0, 1 - min(yh * phase_speed / p, 1.0))
        df_pred["tariff"] = df_pred["hs"].apply(_t)
        return df_pred

    for yh in [0, 5, 10]:
        # CF: phase_speed=0 (tariffs stay at MFN)
        df_cf = _project_features(yh, phase_speed=0.0)
        cf_imp = float(model.predict(df_cf[df_cf["flow"] == "M"][feature_cols]).sum())
        cf_exp = float(model.predict(df_cf[df_cf["flow"] == "X"][feature_cols]).sum())

        # AG: phase_speed=1.0 (full phase-out per agreement)
        df_ag = _project_features(yh, phase_speed=1.0)
        ag_imp = float(model.predict(df_ag[df_ag["flow"] == "M"][feature_cols]).sum())

        # AG exports unadjusted vs F2F-adjusted
        capf = F2F_AT_YEAR[yh]
        df_ag_exp = df_ag[df_ag["flow"] == "X"].copy()
        ag_exp_unadj = float(model.predict(df_ag_exp[feature_cols]).sum())
        f2f_factor = np.where(df_ag_exp["hs"].isin(AGRI_HS), capf, 1.0)
        ag_exp_adj = float((model.predict(df_ag_exp[feature_cols]) * f2f_factor).sum())

        # CF exports F2F-adjusted (for symmetric comparison)
        df_cf_exp = df_cf[df_cf["flow"] == "X"].copy()
        cf_factor = np.where(df_cf_exp["hs"].isin(AGRI_HS), capf, 1.0)
        cf_exp_adj = float((model.predict(df_cf_exp[feature_cols]) * cf_factor).sum())

        scenarios.append({
            "horizon": yh,
            "cf_imports": cf_imp,
            "ag_imports": ag_imp,
            "import_change": ag_imp - cf_imp,
            "cf_exports_unadj": cf_exp,
            "cf_exports_adj": cf_exp_adj,
            "ag_exports_unadj": ag_exp_unadj,
            "ag_exports_adj": ag_exp_adj,
            "export_change_unadj": ag_exp_unadj - cf_exp,
            "export_change_adj": ag_exp_adj - cf_exp_adj,
            "capacity_factor": capf,
            "f2f_export_loss": ag_exp_adj - ag_exp_unadj,
        })

    # Monte Carlo
    rng = np.random.default_rng(mc_seed)
    mc_y5 = []
    mc_y10 = []
    for _ in range(mc_n_draws):
        gdp_g_pert = rng.uniform(0.01, 0.03)
        phase_pert = rng.uniform(0.8, 1.2)
        for yh in [5, 10]:
            df_pred = df_feat[(df_feat["year"] == 2024) & (df_feat["flow"] == "M")].copy()
            df_pred["year"] = 2024 + yh
            df_pred["year_trend"] = (2024 + yh) - 2014
            df_pred["ln_gdp_greece"] = np.log(gdp_g[2024] * ((1 + gdp_g_pert) ** yh))
            # v7.1 fix: project partner GDP at the fixed 2% used by the deterministic
            # scenario, so the MC is centered on the point estimate. Previously partner
            # GDP was left frozen at 2024, which shifted the MC distribution and left the
            # Y10 point estimate outside its own 95% CI (Sheet 9 symptom).
            df_pred["ln_gdp_partner"] = df_pred["partner"].apply(
                lambda p: np.log(gdp_p.get(p, {}).get(2024, 100) * (1.02 ** yh))
            )

            def _t_ag(h):
                c = eu_tariffs.get(h, 0.03)
                p = phase_years.get(h, 7)
                if c == 0 or p == 0:
                    return c
                return c * max(0, 1 - min(yh * phase_pert / p, 1.0))
            df_pred["tariff"] = df_pred["hs"].apply(_t_ag)
            ag = float(model.predict(df_pred[feature_cols]).sum())

            df_cf = df_pred.copy()
            df_cf["tariff"] = df_cf["hs"].apply(lambda h: eu_tariffs.get(h, 0.03))
            cf = float(model.predict(df_cf[feature_cols]).sum())
            change = ag - cf
            if yh == 5:
                mc_y5.append(change)
            else:
                mc_y10.append(change)

    mc_results = {
        "y5_mean": float(np.mean(mc_y5)),
        "y5_lo": float(np.percentile(mc_y5, 2.5)),
        "y5_hi": float(np.percentile(mc_y5, 97.5)),
        "y10_mean": float(np.mean(mc_y10)),
        "y10_lo": float(np.percentile(mc_y10, 2.5)),
        "y10_hi": float(np.percentile(mc_y10, 97.5)),
        "n_draws": mc_n_draws,
        "seed": mc_seed,
    }

    return {
        "model_stats": model_stats,
        "scenarios": scenarios,
        "monte_carlo": mc_results,
        "baselines": {"imports_2024": float(base_imp), "exports_2024": float(base_exp)},
    }


# ====================================================================
# MODULE 4: EXCEL WRITER
# ====================================================================

# Style constants
STYLE = {
    "FONT_HEADER": Font(name="Arial", size=11, bold=True, color="FFFFFF"),
    "FONT_TITLE": Font(name="Arial", size=14, bold=True),
    "FONT_SUBTITLE": Font(name="Arial", size=11, bold=True, italic=True),
    "FONT_NOTE": Font(name="Arial", size=9, italic=True, color="555555"),
    "FONT_TOTAL": Font(name="Arial", size=10, bold=True),
    "FILL_HEADER": PatternFill(start_color="305496", end_color="305496", fill_type="solid"),
    "FILL_SECTION": PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid"),
    "FILL_TOTAL": PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"),
    "FILL_NEW": PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"),
    "FILL_CENTRAL": PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid"),
    "FILL_PARAM": PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid"),
    "BORDER": Border(left=Side(border_style="thin", color="999999"),
                     right=Side(border_style="thin", color="999999"),
                     top=Side(border_style="thin", color="999999"),
                     bottom=Side(border_style="thin", color="999999")),
    "ALIGN_CENTER": Alignment(horizontal="center", vertical="center", wrap_text=True),
    "ALIGN_LEFT": Alignment(horizontal="left", vertical="center", wrap_text=True),
}


def style_cell(cell, font=None, fill=None, alignment=None, border=True, number_format=None):
    if font: cell.font = font
    if fill: cell.fill = fill
    if alignment: cell.alignment = alignment
    if border: cell.border = STYLE["BORDER"]
    if number_format: cell.number_format = number_format


def write_sheet1_raw_trade(wb, trade_data_path):
    """Sheet 1: Raw Greek-Mercosur bilateral trade aggregates."""
    df = load_trade_data(trade_data_path)
    yearly = df.groupby(["year", "flow"])["value"].sum().unstack(fill_value=0)
    yearly = yearly.rename(columns={"M": "imports", "X": "exports"})
    yearly["balance"] = yearly["exports"] - yearly["imports"]
    yearly["imp_yoy"] = yearly["imports"].pct_change()
    yearly["exp_yoy"] = yearly["exports"].pct_change()

    ws = wb.create_sheet("1-Raw Trade Data")
    ws["A1"] = "Greece-Mercosur Bilateral Trade, 2014-2024 (UN Comtrade)"
    style_cell(ws["A1"], font=STYLE["FONT_TITLE"], border=False)
    ws.merge_cells("A1:F1")

    headers = ["Year", "Total Imports (USD)", "Total Exports (USD)", "Trade Balance",
               "Import YoY Growth", "Export YoY Growth"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(4, i)
        c.value = h
        style_cell(c, font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"],
                   alignment=STYLE["ALIGN_CENTER"])

    for i, year in enumerate(sorted(yearly.index)):
        r = 5 + i
        ws.cell(r, 1).value = year
        ws.cell(r, 2).value = float(yearly.loc[year, "imports"])
        ws.cell(r, 3).value = float(yearly.loc[year, "exports"])
        ws.cell(r, 4).value = float(yearly.loc[year, "balance"])
        if not pd.isna(yearly.loc[year, "imp_yoy"]):
            ws.cell(r, 5).value = float(yearly.loc[year, "imp_yoy"])
            ws.cell(r, 5).number_format = "0.0%"
        if not pd.isna(yearly.loc[year, "exp_yoy"]):
            ws.cell(r, 6).value = float(yearly.loc[year, "exp_yoy"])
            ws.cell(r, 6).number_format = "0.0%"
        for c in range(2, 5):
            ws.cell(r, c).number_format = "$#,##0"
        for c in range(1, 7):
            ws.cell(r, c).border = STYLE["BORDER"]

    # CAGR row
    cagr_r = 5 + len(yearly) + 1
    first_year = sorted(yearly.index)[0]
    last_year = sorted(yearly.index)[-1]
    n_years = last_year - first_year
    cagr_imp = (yearly.loc[last_year, "imports"] / yearly.loc[first_year, "imports"]) ** (1/n_years) - 1
    cagr_exp = (yearly.loc[last_year, "exports"] / yearly.loc[first_year, "exports"]) ** (1/n_years) - 1
    ws.cell(cagr_r, 1).value = f"CAGR ({first_year}-{last_year})"
    ws.cell(cagr_r, 2).value = float(cagr_imp); ws.cell(cagr_r, 2).number_format = "0.0%"
    ws.cell(cagr_r, 3).value = float(cagr_exp); ws.cell(cagr_r, 3).number_format = "0.0%"
    style_cell(ws.cell(cagr_r, 1), font=STYLE["FONT_TOTAL"], fill=STYLE["FILL_TOTAL"])
    style_cell(ws.cell(cagr_r, 2), font=STYLE["FONT_TOTAL"], fill=STYLE["FILL_TOTAL"])
    style_cell(ws.cell(cagr_r, 3), font=STYLE["FONT_TOTAL"], fill=STYLE["FILL_TOTAL"])

    for col, w in [(1, 12), (2, 22), (3, 22), (4, 18), (5, 16), (6, 16)]:
        ws.column_dimensions[get_column_letter(col)].width = w


def write_sheet2_tariff_params(wb):
    """Sheet 2: Tariff Parameters with sourced wedge."""
    ws = wb.create_sheet("2-Tariff Parameters")
    ws["A1"] = "Tariff Parameters and Cost Wedge Calibration"
    style_cell(ws["A1"], font=STYLE["FONT_TITLE"], border=False)
    ws.merge_cells("A1:M1")

    headers = ["Product", "HS Code", "EU MFN", "Reduction", "Phase EU", "Target EU",
               "Mercosur MFN", "Phase MS", "Sensitivity", "Imp Elast", "Exp Elast",
               "Wedge", "Tier (Source)"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(4, i)
        c.value = h
        style_cell(c, font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"],
                   alignment=STYLE["ALIGN_CENTER"])

    for i, chap in enumerate(CHAPTERS):
        r = 5 + i
        ws.cell(r, 1).value = chap["name"]
        ws.cell(r, 2).value = chap["hs"]
        ws.cell(r, 3).value = chap["eu_mfn"]; ws.cell(r, 3).number_format = "0.0%"
        ws.cell(r, 4).value = chap["reduction_type"]
        ws.cell(r, 5).value = chap["phase_eu"]
        ws.cell(r, 6).value = chap["target_eu"]; ws.cell(r, 6).number_format = "0.0%"
        ws.cell(r, 7).value = chap["mer_mfn"]; ws.cell(r, 7).number_format = "0.0%"
        ws.cell(r, 8).value = chap["phase_mer"]
        ws.cell(r, 9).value = chap["sensitivity"]
        ws.cell(r, 10).value = chap["imp_elast"]
        ws.cell(r, 11).value = chap["exp_elast"]
        ws.cell(r, 12).value = chap["wedge"]
        ws.cell(r, 12).number_format = "0.0%"
        ws.cell(r, 12).font = Font(name="Arial", color="0000FF")
        if chap.get("citation"):
            cmt = Comment(chap["citation"], "Sourced calibration")
            cmt.width = 380
            cmt.height = 160
            ws.cell(r, 12).comment = cmt
        ws.cell(r, 13).value = chap["tier"]
        for c in range(1, 14):
            ws.cell(r, c).border = STYLE["BORDER"]

    for col, w in [(1, 32), (2, 8), (3, 9), (4, 10), (5, 9), (6, 10), (7, 11), (8, 9),
                   (9, 11), (10, 10), (11, 10), (12, 9), (13, 22)]:
        ws.column_dimensions[get_column_letter(col)].width = w


def write_sheet3_baselines(wb):
    """Sheet 3: Product baselines and CAGRs."""
    ws = wb.create_sheet("3-Product Baselines")
    ws["A1"] = "Product Baselines (avg 2022-2024) and Compound Growth Rates"
    style_cell(ws["A1"], font=STYLE["FONT_TITLE"], border=False)
    ws.merge_cells("A1:F1")

    headers = ["Product Category", "HS Code", "Avg Annual Imports (2022-24)",
               "Avg Annual Exports (2022-24)", "Import CAGR", "Export CAGR"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(4, i)
        c.value = h
        style_cell(c, font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"],
                   alignment=STYLE["ALIGN_CENTER"])

    for i, chap in enumerate(CHAPTERS):
        r = 5 + i
        ws.cell(r, 1).value = chap["name"]
        ws.cell(r, 2).value = chap["hs"]
        ws.cell(r, 3).value = chap["baseline_imp"]; ws.cell(r, 3).number_format = "$#,##0"
        ws.cell(r, 4).value = chap["baseline_exp"]; ws.cell(r, 4).number_format = "$#,##0"
        ws.cell(r, 5).value = chap["cagr_imp"]; ws.cell(r, 5).number_format = "0.00%"
        ws.cell(r, 6).value = chap["cagr_exp"]; ws.cell(r, 6).number_format = "0.00%"
        for c in range(1, 7):
            ws.cell(r, c).border = STYLE["BORDER"]

    for col, w in [(1, 32), (2, 8), (3, 22), (4, 22), (5, 12), (6, 12)]:
        ws.column_dimensions[get_column_letter(col)].width = w


def write_scenario_sheet(wb, sheet_name, results, agreement_year):
    """Sheet 4 or 5: Scenario projection table."""
    ws = wb.create_sheet(sheet_name)
    ws["A1"] = f"{sheet_name}: Armington PE Projections (entry year {agreement_year})"
    style_cell(ws["A1"], font=STYLE["FONT_TITLE"], border=False)
    ws.merge_cells("A1:T1")

    ws["A3"] = "Agreement entry year"; ws["B3"] = agreement_year
    ws["B3"].font = STYLE["FONT_TOTAL"]; ws["B3"].fill = STYLE["FILL_PARAM"]

    # Header rows
    headers = ["Product", "HS", "EU Tariff Y0", "EU Tariff Y5", "EU Tariff Y10",
               "MS Tariff Y0", "MS Tariff Y5", "MS Tariff Y10",
               "CF Imp Y0", "CF Imp Y5", "CF Imp Y10",
               "CF Exp Y0", "CF Exp Y5", "CF Exp Y10",
               "Proj Imp Y0", "Proj Imp Y5", "Proj Imp Y10",
               "Proj Exp Y0", "Proj Exp Y5", "Proj Exp Y10"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(5, i)
        c.value = h
        style_cell(c, font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"],
                   alignment=STYLE["ALIGN_CENTER"])

    for i, r_data in enumerate(results):
        r = 6 + i
        ws.cell(r, 1).value = r_data["name"]
        ws.cell(r, 2).value = r_data["hs"]
        ws.cell(r, 3).value = r_data["years"][0]["eu_tariff"]
        ws.cell(r, 4).value = r_data["years"][5]["eu_tariff"]
        ws.cell(r, 5).value = r_data["years"][10]["eu_tariff"]
        ws.cell(r, 6).value = r_data["years"][0]["mer_tariff"]
        ws.cell(r, 7).value = r_data["years"][5]["mer_tariff"]
        ws.cell(r, 8).value = r_data["years"][10]["mer_tariff"]
        ws.cell(r, 9).value = r_data["years"][0]["cf_imp"]
        ws.cell(r, 10).value = r_data["years"][5]["cf_imp"]
        ws.cell(r, 11).value = r_data["years"][10]["cf_imp"]
        ws.cell(r, 12).value = r_data["years"][0]["cf_exp"]
        ws.cell(r, 13).value = r_data["years"][5]["cf_exp"]
        ws.cell(r, 14).value = r_data["years"][10]["cf_exp"]
        ws.cell(r, 15).value = r_data["years"][0]["proj_imp"]
        ws.cell(r, 16).value = r_data["years"][5]["proj_imp"]
        ws.cell(r, 17).value = r_data["years"][10]["proj_imp"]
        ws.cell(r, 18).value = r_data["years"][0]["proj_exp"]
        ws.cell(r, 19).value = r_data["years"][5]["proj_exp"]
        ws.cell(r, 20).value = r_data["years"][10]["proj_exp"]
        for c in range(3, 9):
            ws.cell(r, c).number_format = "0.0%"
        for c in range(9, 21):
            ws.cell(r, c).number_format = "$#,##0"
        for c in range(1, 21):
            ws.cell(r, c).border = STYLE["BORDER"]
        if r_data["hs"] in ["04", "15", "22"]:
            for c in range(1, 21):
                ws.cell(r, c).fill = STYLE["FILL_NEW"]

    # TOTAL row
    total_r = 6 + len(results)
    ws.cell(total_r, 1).value = "TOTAL"
    style_cell(ws.cell(total_r, 1), font=STYLE["FONT_TOTAL"], fill=STYLE["FILL_TOTAL"])
    for c in range(9, 21):
        col_letter = get_column_letter(c)
        ws.cell(total_r, c).value = f"=SUM({col_letter}6:{col_letter}{total_r-1})"
        ws.cell(total_r, c).number_format = "$#,##0"
        style_cell(ws.cell(total_r, c), font=STYLE["FONT_TOTAL"], fill=STYLE["FILL_TOTAL"])

    # Import change section
    sec_r = total_r + 2
    ws.cell(sec_r, 1).value = "IMPORT/EXPORT CHANGE vs COUNTERFACTUAL"
    style_cell(ws.cell(sec_r, 1), font=STYLE["FONT_SUBTITLE"], fill=STYLE["FILL_SECTION"])
    ws.merge_cells(start_row=sec_r, start_column=1, end_row=sec_r, end_column=14)

    hdr2_r = sec_r + 1
    h2 = ["Product", "HS", "Δ Imp Y0", "Δ Imp Y5", "Δ Imp Y10", "Δ Imp % Y0", "Δ Imp % Y5",
          "Δ Imp % Y10", "Δ Exp Y0", "Δ Exp Y5", "Δ Exp Y10", "Δ Exp % Y0", "Δ Exp % Y5", "Δ Exp % Y10"]
    for i, h in enumerate(h2, 1):
        c = ws.cell(hdr2_r, i)
        c.value = h
        style_cell(c, font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"],
                   alignment=STYLE["ALIGN_CENTER"])

    for i, r_data in enumerate(results):
        r = hdr2_r + 1 + i
        ws.cell(r, 1).value = r_data["name"]
        ws.cell(r, 2).value = r_data["hs"]
        for j, y in enumerate([0, 5, 10]):
            d_imp = r_data["years"][y]["proj_imp"] - r_data["years"][y]["cf_imp"]
            d_exp = r_data["years"][y]["proj_exp"] - r_data["years"][y]["cf_exp"]
            ws.cell(r, 3+j).value = d_imp
            ws.cell(r, 3+j).number_format = "$#,##0"
            ws.cell(r, 6+j).value = d_imp / r_data["years"][y]["cf_imp"] if r_data["years"][y]["cf_imp"] else 0
            ws.cell(r, 6+j).number_format = "0.0%"
            ws.cell(r, 9+j).value = d_exp
            ws.cell(r, 9+j).number_format = "$#,##0"
            ws.cell(r, 12+j).value = d_exp / r_data["years"][y]["cf_exp"] if r_data["years"][y]["cf_exp"] else 0
            ws.cell(r, 12+j).number_format = "0.0%"
        for c in range(1, 15):
            ws.cell(r, c).border = STYLE["BORDER"]

    # Final TOTAL row for change section
    final_r = hdr2_r + 1 + len(results)
    ws.cell(final_r, 1).value = "TOTAL"
    style_cell(ws.cell(final_r, 1), font=STYLE["FONT_TOTAL"], fill=STYLE["FILL_TOTAL"])
    for c in [3, 4, 5, 9, 10, 11]:
        col_letter = get_column_letter(c)
        ws.cell(final_r, c).value = f"=SUM({col_letter}{hdr2_r+1}:{col_letter}{final_r-1})"
        ws.cell(final_r, c).number_format = "$#,##0"
        style_cell(ws.cell(final_r, c), font=STYLE["FONT_TOTAL"], fill=STYLE["FILL_TOTAL"])

    ws.column_dimensions["A"].width = 32
    for col in range(2, 21):
        ws.column_dimensions[get_column_letter(col)].width = 14
    return total_r, final_r  # Return TOTAL row positions for cross-references


def write_sheet6_summary(wb, s1_results, s2_results):
    """Sheet 6: Cross-sheet summary."""
    ws = wb.create_sheet("6-Summary")
    ws["A1"] = "Summary Comparison: S1 vs S2 — All Headlines"
    style_cell(ws["A1"], font=STYLE["FONT_TITLE"], border=False)
    ws.merge_cells("A1:F1")

    headers = ["Metric", "S1 Year 0", "S1 Year 5", "S1 Year 10", "S2 Year 5", "S2 Year 10"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(4, i)
        c.value = h
        style_cell(c, font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"],
                   alignment=STYLE["ALIGN_CENTER"])

    def agg(results, y, key):
        return sum(r["years"][y][key] for r in results)

    # Compute aggregates directly (rather than cross-sheet refs since values are already computed)
    metrics = [
        ("Total Counterfactual Imports", "cf_imp"),
        ("Total Projected Imports", "proj_imp"),
        ("Additional Imports (Proj-CF)", None),
        ("Total Counterfactual Exports", "cf_exp"),
        ("Total Projected Exports", "proj_exp"),
        ("Additional Exports (Proj-CF)", None),
        ("Net Balance Change", None),
    ]
    for i, (label, key) in enumerate(metrics):
        r = 5 + i
        ws.cell(r, 1).value = label
        if key:
            ws.cell(r, 2).value = agg(s1_results, 0, key)
            ws.cell(r, 3).value = agg(s1_results, 5, key)
            ws.cell(r, 4).value = agg(s1_results, 10, key)
            ws.cell(r, 5).value = agg(s2_results, 5, key)
            ws.cell(r, 6).value = agg(s2_results, 10, key)
        elif "Additional Imports" in label:
            ws.cell(r, 2).value = agg(s1_results, 0, "proj_imp") - agg(s1_results, 0, "cf_imp")
            ws.cell(r, 3).value = agg(s1_results, 5, "proj_imp") - agg(s1_results, 5, "cf_imp")
            ws.cell(r, 4).value = agg(s1_results, 10, "proj_imp") - agg(s1_results, 10, "cf_imp")
            ws.cell(r, 5).value = agg(s2_results, 5, "proj_imp") - agg(s2_results, 5, "cf_imp")
            ws.cell(r, 6).value = agg(s2_results, 10, "proj_imp") - agg(s2_results, 10, "cf_imp")
        elif "Additional Exports" in label:
            ws.cell(r, 2).value = agg(s1_results, 0, "proj_exp") - agg(s1_results, 0, "cf_exp")
            ws.cell(r, 3).value = agg(s1_results, 5, "proj_exp") - agg(s1_results, 5, "cf_exp")
            ws.cell(r, 4).value = agg(s1_results, 10, "proj_exp") - agg(s1_results, 10, "cf_exp")
            ws.cell(r, 5).value = agg(s2_results, 5, "proj_exp") - agg(s2_results, 5, "cf_exp")
            ws.cell(r, 6).value = agg(s2_results, 10, "proj_exp") - agg(s2_results, 10, "cf_exp")
        elif "Net Balance" in label:
            for col, (results, y) in enumerate([(s1_results, 0), (s1_results, 5), (s1_results, 10),
                                                 (s2_results, 5), (s2_results, 10)], 2):
                ai = agg(results, y, "proj_imp") - agg(results, y, "cf_imp")
                ae = agg(results, y, "proj_exp") - agg(results, y, "cf_exp")
                ws.cell(r, col).value = ai - ae

        for c in range(2, 7):
            ws.cell(r, c).number_format = "$#,##0"
            ws.cell(r, c).border = STYLE["BORDER"]
        ws.cell(r, 1).border = STYLE["BORDER"]

    ws.column_dimensions["A"].width = 35
    for col in range(2, 7):
        ws.column_dimensions[get_column_letter(col)].width = 17


def write_sheet7_trq(wb, trq_caps, s1_results):
    """Sheet 7: TRQ Caps Analysis."""
    ws = wb.create_sheet("7-TRQ Caps Analysis")
    ws["A1"] = "Tariff Rate Quotas (TRQs): Greek-Bilateral Binding Analysis"
    style_cell(ws["A1"], font=STYLE["FONT_TITLE"], border=False)
    ws.merge_cells("A1:H1")

    ws["A2"] = ("Eight TRQs in the EU-Mercosur agreement, translated to Greek-share allocations "
                "via 2.4% population share. Comparison to v6 projected Y10 imports shows none bind "
                "at the Greek bilateral level. Caps remain documented for defensive robustness.")
    ws["A2"].font = STYLE["FONT_NOTE"]
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:H2")
    ws.row_dimensions[2].height = 50

    headers = ["Product", "HS Code", "EU Volume (tonnes/yr)", "World Price ($/tonne)",
               "EU TRQ Value ($M)", "Greek Share ($M)", "v6 Y10 Greek Proj Imp ($M)",
               "Binding for Greece?"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(5, i)
        c.value = h
        style_cell(c, font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"],
                   alignment=STYLE["ALIGN_CENTER"])
    ws.row_dimensions[5].height = 30

    GSHARE = 0.024
    trqs = [
        ("Cheese", "0406", 30000, 5000, "04"),
        ("Milk powder", "0402", 10000, 3000, "04"),
        ("Honey", "0409", 45000, 4000, "04"),
        ("Beef", "0201/02", 99000, 5000, "02"),
        ("Poultry", "0207", 180000, 2500, "02"),
        ("Sugar refining", "1701", 180000, 400, "17"),
        ("Ethanol fuel", "2207.20", 200000, 600, "22"),
        ("Ethanol chemical", "2207.10", 450000, 600, "22"),
    ]
    proj_y10_by_hs = {r["hs"]: r["years"][10]["proj_imp"]/1e6 for r in s1_results}

    for i, (name, hs_code, vol, price, chapter) in enumerate(trqs):
        r = 6 + i
        ws.cell(r, 1).value = name
        ws.cell(r, 2).value = hs_code
        ws.cell(r, 3).value = vol; ws.cell(r, 3).number_format = "#,##0"
        ws.cell(r, 4).value = price; ws.cell(r, 4).number_format = "$#,##0"
        eu_val = vol * price / 1e6
        ws.cell(r, 5).value = eu_val; ws.cell(r, 5).number_format = "$#,##0.0"
        greek_cap = eu_val * GSHARE
        ws.cell(r, 6).value = greek_cap; ws.cell(r, 6).number_format = "$#,##0.00"
        if chapter in proj_y10_by_hs:
            proj = proj_y10_by_hs[chapter]
            ws.cell(r, 7).value = proj; ws.cell(r, 7).number_format = "$#,##0.0"
            if chapter == "22":
                ws.cell(r, 8).value = "Ethanol subset only"
            elif proj > greek_cap:
                ws.cell(r, 8).value = "BINDING"
            elif proj > 0:
                ws.cell(r, 8).value = f"Not binding ({greek_cap/proj:.0f}x headroom)"
            else:
                ws.cell(r, 8).value = "Not binding (proj~0)"
        else:
            ws.cell(r, 7).value = "NOT in v6"
            ws.cell(r, 8).value = "n/a (chapter not in model)"
        for c in range(1, 9):
            ws.cell(r, c).border = STYLE["BORDER"]

    # TRQ cap parameters section
    cap_start = 6 + len(trqs) + 3
    ws.cell(cap_start, 1).value = "TRQ CAP PARAMETERS (applied to projection formulas)"
    style_cell(ws.cell(cap_start, 1), font=STYLE["FONT_SUBTITLE"], fill=STYLE["FILL_SECTION"])
    ws.merge_cells(start_row=cap_start, start_column=1, end_row=cap_start, end_column=8)

    headers2 = ["HS Chapter", "Greek Share Cap (USD)", "Coverage", "Combined EU TRQ ($M)",
                "Status", "Applied?", "Cell Reference", ""]
    for i, h in enumerate(headers2, 1):
        c = ws.cell(cap_start + 2, i)
        c.value = h
        style_cell(c, font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"],
                   alignment=STYLE["ALIGN_CENTER"])

    for i, (hs, cap_data) in enumerate(trq_caps.items()):
        r = cap_start + 3 + i
        ws.cell(r, 1).value = hs
        ws.cell(r, 2).value = cap_data["cap_usd"]; ws.cell(r, 2).number_format = '"$"#,##0'
        ws.cell(r, 2).fill = STYLE["FILL_PARAM"]
        ws.cell(r, 3).value = cap_data["coverage"]
        ws.cell(r, 4).value = cap_data["eu_combined_m"]; ws.cell(r, 4).number_format = '"$"#,##0.0"M"'
        ws.cell(r, 5).value = "ACTIVE" if cap_data["applied"] else "Documented only"
        ws.cell(r, 6).value = "Yes" if cap_data["applied"] else f"No ({cap_data.get('reason', '')})"
        ws.cell(r, 7).value = f"$B${r}"
        for c in range(1, 8):
            ws.cell(r, c).border = STYLE["BORDER"]

    for col, w in [(1, 22), (2, 14), (3, 14), (4, 14), (5, 14), (6, 14), (7, 22), (8, 28)]:
        ws.column_dimensions[get_column_letter(col)].width = w


def write_sheet8_sensitivity(wb, sens_grid_s1, sens_grid_s2, greek_agri_share, wedge_mults, f2f_drags):
    """Sheet 8: Sensitivity Grid (4 wedge multipliers x 3 F2F drag levels in Tier 2)."""
    ws = wb.create_sheet("8-Sensitivity Grid")
    ws["A1"] = "Sensitivity Grid: Wedge Multiplier x F2F Drag (Year 10)"
    style_cell(ws["A1"], font=STYLE["FONT_TITLE"], border=False)
    ws.merge_cells("A1:G1")

    ws["A2"] = (f"{len(wedge_mults)} wedge multipliers x {len(f2f_drags)} F2F drag levels = "
                f"{len(wedge_mults)*len(f2f_drags)} combinations. Bilateral widening = additional "
                f"imports minus F2F-adjusted exports at Y10. The 3.0x wedge level is the empirical "
                f"upper bound implied by KEPE 2025/57 (5.5pp Greek-EU asymmetric cost burden, "
                f"projected forward at half pace).")
    ws["A2"].font = STYLE["FONT_NOTE"]
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:G2")
    ws.row_dimensions[2].height = 50

    ws["A4"] = "INPUTS"; ws["A4"].font = STYLE["FONT_SUBTITLE"]; ws["A4"].fill = STYLE["FILL_SECTION"]
    ws.merge_cells("A4:G4")
    ws["A5"] = "Greek agri share"; ws["B5"] = greek_agri_share; ws["B5"].number_format = "0.0%"

    # Wedge level labels including KEPE source
    wedge_labels = {
        0.5: "0.5x (weak de facto compliance)",
        1.0: "1.0x (central, Profundo + IISD + EC)",
        1.5: "1.5x (compliance-cost upper)",
        3.0: "3.0x (full-enforcement stress)",
    }

    row = 7
    for sname, grid in [("S1: Agreement enters 2026", sens_grid_s1),
                         ("S2: Agreement enters 2027", sens_grid_s2)]:
        ws.cell(row, 1).value = sname
        style_cell(ws.cell(row, 1), font=STYLE["FONT_SUBTITLE"], fill=STYLE["FILL_SECTION"])
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
        row += 1

        # Headers
        ws.cell(row, 1).value = "Wedge / F2F"
        f2f_labels = [f"F2F {int(f*100)}%" for f in f2f_drags]
        for j, label in enumerate(f2f_labels):
            ws.cell(row, 2+j).value = label
        ws.cell(row, 2+len(f2f_drags)).value = "Add Imp Y10"
        ws.cell(row, 3+len(f2f_drags)).value = "Add Exp Y10 (unadj)"
        for c in range(1, 4 + len(f2f_drags)):
            style_cell(ws.cell(row, c), font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"],
                       alignment=STYLE["ALIGN_CENTER"])
        row += 1

        for w in wedge_mults:
            label = wedge_labels.get(w, f"{w}x")
            ws.cell(row, 1).value = label
            for j, f in enumerate(f2f_drags):
                cell_data = grid[(w, f)]
                ws.cell(row, 2+j).value = cell_data["bilateral_widening"]
                ws.cell(row, 2+j).number_format = "$#,##0"
                if w == 1.0 and f == 0.10:
                    ws.cell(row, 2+j).fill = STYLE["FILL_CENTRAL"]
                    ws.cell(row, 2+j).font = STYLE["FONT_TOTAL"]
                # Highlight 3.0x row in different color (KEPE)
                if w == 3.0:
                    ws.cell(row, 2+j).fill = PatternFill(
                        start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
            ws.cell(row, 2+len(f2f_drags)).value = grid[(w, 0.10)]["add_imp_y10"]
            ws.cell(row, 2+len(f2f_drags)).number_format = "$#,##0"
            ws.cell(row, 3+len(f2f_drags)).value = grid[(w, 0.10)]["add_exp_y10_unadj"]
            ws.cell(row, 3+len(f2f_drags)).number_format = "$#,##0"
            for c in range(1, 4 + len(f2f_drags)):
                ws.cell(row, c).border = STYLE["BORDER"]
            row += 1
        row += 2

    # Range bounds
    s1_widenings = [g["bilateral_widening"] for g in sens_grid_s1.values()]
    ws.cell(row, 1).value = "S1 GRID RANGE (Y10 Bilateral Widening)"
    style_cell(ws.cell(row, 1), font=STYLE["FONT_SUBTITLE"], fill=STYLE["FILL_SECTION"])
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    row += 1
    ws.cell(row, 1).value = "Minimum"; ws.cell(row, 2).value = min(s1_widenings)
    row += 1
    ws.cell(row, 1).value = "Central (1.0x x 10%)"
    ws.cell(row, 2).value = sens_grid_s1[(1.0, 0.10)]["bilateral_widening"]
    style_cell(ws.cell(row, 1), fill=STYLE["FILL_TOTAL"])
    style_cell(ws.cell(row, 2), font=STYLE["FONT_TOTAL"], fill=STYLE["FILL_TOTAL"])
    row += 1
    ws.cell(row, 1).value = "Maximum (KEPE upper)"; ws.cell(row, 2).value = max(s1_widenings)
    row += 1
    ws.cell(row, 1).value = "Range"; ws.cell(row, 2).value = max(s1_widenings) - min(s1_widenings)
    for r_set in range(row-3, row+1):
        ws.cell(r_set, 2).number_format = "$#,##0"

    for col, w in [(1, 36), (2, 16), (3, 16), (4, 16), (5, 18), (6, 18), (7, 22)]:
        ws.column_dimensions[get_column_letter(col)].width = w


def write_sheet9_monte_carlo(wb, ml_results, pe_mc):
    """Sheet 9: Monte Carlo CI (XGBoost ML + PE Armington)."""
    ws = wb.create_sheet("9-Monte Carlo CI")
    ws["A1"] = "Monte Carlo 95% Confidence Intervals: ML and PE Armington"
    style_cell(ws["A1"], font=STYLE["FONT_TITLE"], border=False)
    ws.merge_cells("A1:H1")

    mc = ml_results["monte_carlo"]
    ws["A2"] = (f"Two Monte Carlo simulations. ML side: XGBoost trained on UN Comtrade 2014-2023, "
                f"perturbing GDP growth and phase-out speed. PE side: Armington model perturbing "
                f"wedge multiplier (triangular 0.5-1.0-3.0, KEPE-anchored upper) and F2F drag "
                f"(triangular 0.05-0.10-0.20). Both use {mc['n_draws']} draws, seed {mc['seed']}.")
    ws["A2"].font = STYLE["FONT_NOTE"]
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:H2")
    ws.row_dimensions[2].height = 70

    ws["A4"] = "ML MODEL STATS AND MONTE CARLO RESULTS"
    style_cell(ws["A4"], font=STYLE["FONT_SUBTITLE"], fill=STYLE["FILL_SECTION"])
    ws.merge_cells("A4:H4")
    ws["A5"] = "RMSE (test set)"; ws["B5"] = ml_results["model_stats"]["rmse"]
    ws["B5"].number_format = "$#,##0"
    ws["A6"] = "MAPE (test set)"; ws["B6"] = ml_results["model_stats"]["mape"] / 100
    ws["B6"].number_format = "0.0%"
    ws["A7"] = "MC draws"; ws["B7"] = mc["n_draws"]
    ws["A8"] = "MC seed"; ws["B8"] = mc["seed"]

    headers = ["Year", "CF Imports ($)", "Adj Imports ($)", "Δ Imports ($)",
               "MC Mean ($)", "95% CI Low ($)", "95% CI High ($)", "CI Width ($)"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(10, i)
        c.value = h
        style_cell(c, font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"],
                   alignment=STYLE["ALIGN_CENTER"])

    for i, sc in enumerate(ml_results["scenarios"]):
        r = 11 + i
        ws.cell(r, 1).value = f"Year {sc['horizon']}"
        ws.cell(r, 2).value = sc["cf_imports"]
        ws.cell(r, 3).value = sc["ag_imports"]
        ws.cell(r, 4).value = sc["import_change"]
        if sc["horizon"] == 5:
            ws.cell(r, 5).value = mc["y5_mean"]
            ws.cell(r, 6).value = mc["y5_lo"]
            ws.cell(r, 7).value = mc["y5_hi"]
            ws.cell(r, 8).value = mc["y5_hi"] - mc["y5_lo"]
        elif sc["horizon"] == 10:
            ws.cell(r, 5).value = mc["y10_mean"]
            ws.cell(r, 6).value = mc["y10_lo"]
            ws.cell(r, 7).value = mc["y10_hi"]
            ws.cell(r, 8).value = mc["y10_hi"] - mc["y10_lo"]
        else:
            for c in range(5, 9):
                ws.cell(r, c).value = 0
        for c in range(2, 9):
            ws.cell(r, c).number_format = "$#,##0"
            ws.cell(r, c).border = STYLE["BORDER"]
        ws.cell(r, 1).border = STYLE["BORDER"]

    # ----- PE Armington Monte Carlo Section -----
    pe_r = 16
    ws.cell(pe_r, 1).value = "PE ARMINGTON MONTE CARLO RESULTS (Tier 3, literature-anchored)"
    style_cell(ws.cell(pe_r, 1), font=STYLE["FONT_SUBTITLE"], fill=STYLE["FILL_SECTION"])
    ws.merge_cells(start_row=pe_r, start_column=1, end_row=pe_r, end_column=8)

    pe_r += 1
    ws.cell(pe_r, 1).value = ("Wedge multiplier ~ Triangular(0.5, 1.0, 3.0). Upper bound 3.0 "
                              "anchored on KEPE 2025/57 (5.5pp Greek-EU asymmetric burden over "
                              "5 years, projected forward). F2F drag ~ Triangular(0.05, 0.10, 0.20).")
    ws.cell(pe_r, 1).font = STYLE["FONT_NOTE"]
    ws.cell(pe_r, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=pe_r, start_column=1, end_row=pe_r, end_column=8)
    ws.row_dimensions[pe_r].height = 36

    pe_r += 1
    pe_headers = ["Metric", "Year 5 Mean ($)", "Y5 95% CI Low", "Y5 95% CI High",
                   "Year 10 Mean ($)", "Y10 95% CI Low", "Y10 95% CI High", "Std Dev Y10"]
    for i, h in enumerate(pe_headers, 1):
        c = ws.cell(pe_r, i)
        c.value = h
        style_cell(c, font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"],
                   alignment=STYLE["ALIGN_CENTER"])

    pe_r += 1
    rows_data = [
        ("Additional Imports", "add_imp"),
        ("F2F-Adjusted Add Exports", "add_exp_adj"),
        ("Bilateral Widening", "widening"),
    ]
    for i, (label, key) in enumerate(rows_data):
        r = pe_r + i
        ws.cell(r, 1).value = label
        ws.cell(r, 2).value = pe_mc["y5"][key]["mean"]
        ws.cell(r, 3).value = pe_mc["y5"][key]["lo"]
        ws.cell(r, 4).value = pe_mc["y5"][key]["hi"]
        ws.cell(r, 5).value = pe_mc["y10"][key]["mean"]
        ws.cell(r, 6).value = pe_mc["y10"][key]["lo"]
        ws.cell(r, 7).value = pe_mc["y10"][key]["hi"]
        ws.cell(r, 8).value = pe_mc["y10"][key]["std"]
        for c in range(2, 9):
            ws.cell(r, c).number_format = "$#,##0"
            ws.cell(r, c).border = STYLE["BORDER"]
        ws.cell(r, 1).border = STYLE["BORDER"]
        if "Widening" in label:
            for c in range(1, 9):
                ws.cell(r, c).fill = STYLE["FILL_TOTAL"]
                ws.cell(r, c).font = STYLE["FONT_TOTAL"]

    # Feature importance section moved down
    fi_r = pe_r + len(rows_data) + 2
    ws.cell(fi_r, 1).value = "ML FEATURE IMPORTANCE"
    style_cell(ws.cell(fi_r, 1), font=STYLE["FONT_SUBTITLE"], fill=STYLE["FILL_SECTION"])
    ws.merge_cells(start_row=fi_r, start_column=1, end_row=fi_r, end_column=8)

    ws.cell(fi_r+1, 1).value = "Rank"
    ws.cell(fi_r+1, 2).value = "Feature"
    ws.cell(fi_r+1, 3).value = "Gain Importance"
    for c in [1, 2, 3]:
        style_cell(ws.cell(fi_r+1, c), font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"],
                   alignment=STYLE["ALIGN_CENTER"])

    for i, (feat, imp) in enumerate(ml_results["model_stats"]["feature_importance"]):
        r = fi_r + 2 + i
        ws.cell(r, 1).value = i + 1
        ws.cell(r, 2).value = feat
        ws.cell(r, 3).value = imp; ws.cell(r, 3).number_format = "0.0000"
        for c in range(1, 4):
            ws.cell(r, c).border = STYLE["BORDER"]

    for col, w in [(1, 22), (2, 18), (3, 16), (4, 16), (5, 18), (6, 16), (7, 16), (8, 16)]:
        ws.column_dimensions[get_column_letter(col)].width = w


def write_sheet10_yearly(wb, s1_results):
    """Sheet 10: Year-by-Year Annual Detail."""
    ws = wb.create_sheet("10-Year-by-Year")
    ws["A1"] = "Year-by-Year Annual Detail (S1, 2026-2036)"
    style_cell(ws["A1"], font=STYLE["FONT_TITLE"], border=False)
    ws.merge_cells("A1:N1")

    year_headers = [f"Y{y} ({2026+y})" for y in range(11)]

    def section(start_row, title, key_proj, key_cf, signed_pos=True):
        ws.cell(start_row, 1).value = title
        style_cell(ws.cell(start_row, 1), font=STYLE["FONT_SUBTITLE"], fill=STYLE["FILL_SECTION"])
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=14)

        hr = start_row + 1
        ws.cell(hr, 1).value = "Product"; ws.cell(hr, 2).value = "HS"
        for j, yh in enumerate(year_headers):
            ws.cell(hr, 3+j).value = yh
        for c in range(1, 14):
            style_cell(ws.cell(hr, c), font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"],
                       alignment=STYLE["ALIGN_CENTER"])

        for i, r_data in enumerate(s1_results):
            r = hr + 1 + i
            ws.cell(r, 1).value = r_data["name"]
            ws.cell(r, 2).value = r_data["hs"]
            for j in range(11):
                if signed_pos:
                    val = r_data["years"][j][key_proj] - r_data["years"][j][key_cf]
                else:  # net = exp_change - imp_change
                    de = r_data["years"][j]["proj_exp"] - r_data["years"][j]["cf_exp"]
                    di = r_data["years"][j]["proj_imp"] - r_data["years"][j]["cf_imp"]
                    val = de - di
                ws.cell(r, 3+j).value = val
                ws.cell(r, 3+j).number_format = "$#,##0;($#,##0);-"
                ws.cell(r, 3+j).border = STYLE["BORDER"]
            for c in [1, 2]:
                ws.cell(r, c).border = STYLE["BORDER"]
            if r_data["hs"] in ["04", "15", "22"]:
                for c in range(1, 14):
                    ws.cell(r, c).fill = STYLE["FILL_NEW"]

        total_r = hr + 1 + len(s1_results)
        ws.cell(total_r, 1).value = "TOTAL"
        for j in range(11):
            col = get_column_letter(3+j)
            ws.cell(total_r, 3+j).value = f"=SUM({col}{hr+1}:{col}{total_r-1})"
            ws.cell(total_r, 3+j).number_format = "$#,##0;($#,##0);-"
            style_cell(ws.cell(total_r, 3+j), font=STYLE["FONT_TOTAL"], fill=STYLE["FILL_TOTAL"])
        style_cell(ws.cell(total_r, 1), font=STYLE["FONT_TOTAL"], fill=STYLE["FILL_TOTAL"])
        return total_r + 2

    next_r = 4
    next_r = section(next_r, "SECTION A: Additional Imports by Year ($USD)", "proj_imp", "cf_imp", True)
    next_r = section(next_r, "SECTION B: Additional Exports by Year ($USD)", "proj_exp", "cf_exp", True)
    next_r = section(next_r, "SECTION C: Net Impact (Δ Exp - Δ Imp) by Year ($USD)", None, None, False)

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 6
    for col in range(3, 14):
        ws.column_dimensions[get_column_letter(col)].width = 18


def write_sheet11_ml_vs_pe(wb, s1_results, ml_results, greek_agri_share, f2f_drag,
                            cap_weights=None, labor_drag=0.0):
    """Sheet 11: ML vs PE Comparison."""
    ws = wb.create_sheet("11-ML vs PE Comparison")
    ws["A1"] = "Direct Comparison: PE Armington vs XGBoost ML"
    style_cell(ws["A1"], font=STYLE["FONT_TITLE"], border=False)
    ws.merge_cells("A1:F1")

    ws["A2"] = ("Two-orders-of-magnitude divergence is itself the central empirical finding. "
                "PE Armington assumes -3.5 import elasticity (Hertel et al. 2007 GTAP norm); "
                "XGBoost learns from 11 years of data and finds tariff is a low-importance feature.")
    ws["A2"].font = STYLE["FONT_NOTE"]
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:F2")
    ws.row_dimensions[2].height = 50

    ws["A4"] = "HEADLINE COMPARISON (Year 10, S1)"
    style_cell(ws["A4"], font=STYLE["FONT_SUBTITLE"], fill=STYLE["FILL_SECTION"])
    ws.merge_cells("A4:F4")

    headers = ["Metric", "PE Armington ($)", "XGBoost ML ($)", "Difference ($)", "Ratio (PE/ML)", "Interpretation"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(5, i)
        c.value = h
        style_cell(c, font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"],
                   alignment=STYLE["ALIGN_CENTER"])

    pe_imp_y5 = sum(c["years"][5]["proj_imp"] - c["years"][5]["cf_imp"] for c in s1_results)
    pe_imp_y10 = sum(c["years"][10]["proj_imp"] - c["years"][10]["cf_imp"] for c in s1_results)
    pe_exp_y10 = sum(c["years"][10]["proj_exp"] - c["years"][10]["cf_exp"] for c in s1_results)

    if cap_weights:
        # Tier 3: per-chapter adjustment
        adj_data = adj_exports_per_chapter(s1_results, 10, f2f_drag, labor_drag, cap_weights)
        pe_exp_y10_adj = adj_data["total_adj"]
    else:
        pe_exp_y10_adj = pe_exp_y10 * (1 - greek_agri_share * f2f_drag)
    pe_widening = pe_imp_y10 - pe_exp_y10_adj

    ml_y5 = ml_results["scenarios"][1]["import_change"]
    ml_y10 = ml_results["scenarios"][2]["import_change"]
    ml_exp_y10 = ml_results["scenarios"][2]["export_change_unadj"]
    ml_exp_y10_adj = ml_results["scenarios"][2]["export_change_adj"]
    ml_widening = ml_y10 - ml_exp_y10_adj

    rows = [
        ("Δ Imports Y5", pe_imp_y5, ml_y5, "Empirical model: tariff has tiny predictive power"),
        ("Δ Imports Y10", pe_imp_y10, ml_y10, "PE: large; ML: small contraction"),
        ("Δ Exports Y10 (unadj)", pe_exp_y10, ml_exp_y10, "Both models project export change; magnitudes differ"),
        ("Δ Exports Y10 (capacity-adj)", pe_exp_y10_adj, ml_exp_y10_adj,
         "PE uses per-chapter Greek capacity weights; ML uses uniform F2F"),
        ("Bilateral Widening Y10", pe_widening, ml_widening, "PE = $3.6B widening; ML = effectively zero"),
    ]
    for i, (label, pe, ml, interp) in enumerate(rows):
        r = 6 + i
        ws.cell(r, 1).value = label
        ws.cell(r, 2).value = pe; ws.cell(r, 2).number_format = "$#,##0"
        ws.cell(r, 3).value = ml; ws.cell(r, 3).number_format = "$#,##0"
        ws.cell(r, 4).value = pe - ml; ws.cell(r, 4).number_format = "$#,##0"
        if abs(ml) > 1:
            ws.cell(r, 5).value = pe / ml; ws.cell(r, 5).number_format = '#,##0"x"'
        else:
            ws.cell(r, 5).value = "n/a"
        ws.cell(r, 6).value = interp
        for c in range(1, 7):
            ws.cell(r, c).border = STYLE["BORDER"]
        ws.row_dimensions[r].height = 28

    for col, w in [(1, 28), (2, 18), (3, 16), (4, 16), (5, 12), (6, 40)]:
        ws.column_dimensions[get_column_letter(col)].width = w


# ====================================================================
# MODULE 5: CHART GENERATION
# ====================================================================

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

# Chart style constants. Paper-quality, consistent palette.
CHART_DPI = 200
CHART_PALETTE = {
    "pe": "#1F4E79",        # Dark blue, PE Armington
    "ml": "#C00000",        # Dark red, XGBoost ML
    "f2f": "#7F6000",       # Olive, F2F drag
    "central": "#FF8C42",   # Orange, central case highlight
    "imports": "#2E75B6",   # Light blue, imports
    "exports": "#548235",   # Green, exports
    "widening": "#7030A0",  # Purple, widening
    "ci_band": "#BDD7EE",   # Light blue band
    "neutral": "#404040",   # Dark gray
}
plt.rcParams.update({
    "font.family": "sans-serif",
    "font.sans-serif": ["Arial", "DejaVu Sans"],
    "font.size": 10,
    "axes.titlesize": 12,
    "axes.labelsize": 10,
    "axes.spines.top": False,
    "axes.spines.right": False,
    "axes.grid": True,
    "grid.alpha": 0.3,
    "grid.linestyle": "--",
    "figure.dpi": CHART_DPI,
})


def chart_feature_importance(ml_results, output_path):
    """Horizontal bar chart of XGBoost feature importance."""
    fi = ml_results["model_stats"]["feature_importance"]
    feats = [f for f, _ in fi]
    imps = [i for _, i in fi]
    # Friendly labels
    label_map = {
        "value_lag1": "Lag 1 trade value",
        "ln_distance": "Log distance",
        "ln_gdp_greece": "Log Greek GDP",
        "value_lag2": "Lag 2 trade value",
        "is_import": "Import flag",
        "ln_gdp_partner": "Log partner GDP",
        "year_trend": "Year trend",
        "growth_lag1": "Growth lag 1",
        "tariff": "EU tariff",
        "hs_num": "HS chapter (numeric)",
        "partner_code": "Partner (categorical)",
    }
    labels = [label_map.get(f, f) for f in feats]

    fig, ax = plt.subplots(figsize=(8, 5))
    y = np.arange(len(feats))[::-1]  # Reverse so most important is at top
    bars = ax.barh(y, imps, color=CHART_PALETTE["pe"], alpha=0.85)
    # Highlight tariff in red
    for i, f in enumerate(feats):
        if f == "tariff":
            bars[i].set_color(CHART_PALETTE["ml"])
    ax.set_yticks(y)
    ax.set_yticklabels(labels)
    ax.set_xlabel("Gain importance")
    ax.set_title("XGBoost Feature Importance, Greek-Mercosur Bilateral Trade Model",
                 fontweight="bold", pad=12)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"{x:.0%}"))
    # Annotate values
    for i, v in enumerate(imps):
        ax.text(v + 0.005, y[i], f"{v:.1%}", va="center", fontsize=9)
    ax.set_xlim(0, max(imps) * 1.15)
    fig.tight_layout()
    fig.savefig(output_path, bbox_inches="tight")
    plt.close(fig)


def chart_sensitivity_heatmap(sens_grid_s1, output_path):
    """4x3 heatmap of bilateral widening at Y10 for S1 (Tier 2)."""
    wedge_mults = [0.5, 1.0, 1.5, 3.0]
    f2f_drags = [0.07, 0.10, 0.20]
    grid = np.array([[sens_grid_s1[(w, f)]["bilateral_widening"] / 1e6 for f in f2f_drags]
                     for w in wedge_mults])

    fig, ax = plt.subplots(figsize=(8, 5.5))
    im = ax.imshow(grid, cmap="YlOrRd", aspect="auto")
    ax.set_xticks(np.arange(len(f2f_drags)))
    ax.set_yticks(np.arange(len(wedge_mults)))
    ax.set_xticklabels([f"{int(f*100)}%" for f in f2f_drags])
    wedge_labels = ["0.5x\n(low)", "1.0x\n(central)", "1.5x\n(high)", "3.0x\n(KEPE upper)"]
    ax.set_yticklabels(wedge_labels)
    ax.set_xlabel("F2F capacity drag (Greek agri export reduction by Y10)")
    ax.set_ylabel("Wedge multiplier")
    ax.set_title("Bilateral Trade Widening at Year 10, S1 (entry 2026)\nin USD millions",
                 fontweight="bold", pad=12)

    # Annotate cells
    for i in range(len(wedge_mults)):
        for j in range(len(f2f_drags)):
            val = grid[i, j]
            text_color = "white" if val > grid.mean() else "black"
            weight = "bold" if (wedge_mults[i] == 1.0 and f2f_drags[j] == 0.10) else "normal"
            ax.text(j, i, f"${val:,.0f}M", ha="center", va="center",
                    color=text_color, fontweight=weight, fontsize=11)

    cbar = fig.colorbar(im, ax=ax, shrink=0.8)
    cbar.set_label("USD millions", rotation=270, labelpad=15)
    # Highlight central case with a box
    ax.add_patch(plt.Rectangle((0.5, 0.5), 1.0, 1.0, fill=False,
                                edgecolor=CHART_PALETTE["central"], lw=3))
    fig.text(0.5, 0.02,
             "Orange box: central case (1.0x x 10%). Bottom row (3.0x) is KEPE-anchored empirical upper bound.",
             ha="center", fontsize=8, style="italic", color=CHART_PALETTE["central"])

    fig.tight_layout(rect=[0, 0.05, 1, 1])
    fig.savefig(output_path, bbox_inches="tight")
    plt.close(fig)


def chart_pe_monte_carlo(pe_mc, output_path):
    """Histogram of PE Armington Monte Carlo Y10 widening distribution."""
    draws = np.array(pe_mc["y10_widening_draws"]) / 1e6  # Convert to millions
    mean = np.mean(draws)
    lo, hi = np.percentile(draws, [2.5, 97.5])
    median = np.median(draws)

    fig, ax = plt.subplots(figsize=(9, 5.5))
    n_bins = 40
    counts, bins, patches = ax.hist(draws, bins=n_bins, color=CHART_PALETTE["pe"],
                                     alpha=0.75, edgecolor="white", linewidth=0.5)

    # Reference lines
    ax.axvline(mean, color=CHART_PALETTE["central"], linewidth=2.5,
               label=f"Mean: ${mean:,.0f}M")
    ax.axvline(lo, color=CHART_PALETTE["ml"], linewidth=1.5, linestyle="--",
               label=f"95% CI Low: ${lo:,.0f}M")
    ax.axvline(hi, color=CHART_PALETTE["ml"], linewidth=1.5, linestyle="--",
               label=f"95% CI High: ${hi:,.0f}M")

    # Shade the 95% CI
    in_ci = (bins[:-1] >= lo) & (bins[:-1] <= hi)
    for patch, inside in zip(patches, in_ci):
        if inside:
            patch.set_alpha(0.95)

    ax.set_xlabel("Bilateral Widening at Year 10 (USD millions)")
    ax.set_ylabel("Number of draws")
    ax.set_title(f"PE Armington Monte Carlo Distribution, S1 Y10 Bilateral Widening\n"
                 f"({pe_mc['n_draws']} draws, wedge ~ Triangular(0.5, 1.0, 3.0) anchored on KEPE; "
                 f"F2F drag ~ Triangular(0.05, 0.10, 0.20))",
                 fontweight="bold", pad=12)
    ax.xaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:,.0f}M"))
    ax.legend(loc="upper right", frameon=True, fontsize=10)

    # Annotation
    text = (f"N = {pe_mc['n_draws']} draws. Mean ${mean:,.0f}M, Median ${median:,.0f}M, "
            f"Std ${np.std(draws):,.0f}M.\n"
            f"95% CI: [${lo:,.0f}M, ${hi:,.0f}M], width ${hi-lo:,.0f}M.")
    ax.text(0.02, 0.98, text, transform=ax.transAxes, fontsize=9,
            va="top", bbox=dict(boxstyle="round,pad=0.4", facecolor="white",
                                 edgecolor="gray", alpha=0.9))

    fig.tight_layout()
    fig.savefig(output_path, bbox_inches="tight")
    plt.close(fig)


def chart_ml_vs_pe(s1_results, ml_results, greek_agri_share, f2f_drag, output_path):
    """Side-by-side bar chart at Y10."""
    pe_imp = sum(c["years"][10]["proj_imp"] - c["years"][10]["cf_imp"] for c in s1_results) / 1e6
    pe_exp = sum(c["years"][10]["proj_exp"] - c["years"][10]["cf_exp"] for c in s1_results) / 1e6
    pe_exp_adj = pe_exp * (1 - greek_agri_share * f2f_drag)
    pe_widening = pe_imp - pe_exp_adj

    ml = ml_results["scenarios"][2]
    ml_imp = ml["import_change"] / 1e6
    ml_exp_adj = ml["export_change_adj"] / 1e6
    ml_widening = ml_imp - ml_exp_adj

    metrics = ["Δ Imports", "Δ Exports\n(F2F-adjusted)", "Bilateral\nWidening"]
    pe_vals = [pe_imp, pe_exp_adj, pe_widening]
    ml_vals = [ml_imp, ml_exp_adj, ml_widening]

    fig, ax = plt.subplots(figsize=(8, 5))
    x = np.arange(len(metrics))
    w = 0.35
    bars1 = ax.bar(x - w/2, pe_vals, w, label="PE Armington", color=CHART_PALETTE["pe"], alpha=0.9)
    bars2 = ax.bar(x + w/2, ml_vals, w, label="XGBoost ML", color=CHART_PALETTE["ml"], alpha=0.9)
    ax.set_xticks(x)
    ax.set_xticklabels(metrics)
    ax.set_ylabel("USD millions")
    ax.set_title("PE Armington vs XGBoost ML: Year 10 Projections, S1 (entry 2026)",
                 fontweight="bold", pad=12)
    ax.legend(loc="center right", frameon=True)
    ax.axhline(0, color="black", linewidth=0.5)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:,.0f}M"))
    # Add headroom for labels
    ymax = max(max(pe_vals), max(ml_vals)) * 1.15
    ax.set_ylim(0, ymax)
    # Annotate
    for bars in [bars1, bars2]:
        for b in bars:
            v = b.get_height()
            label = f"${v:,.0f}M"
            yoff = ymax * 0.015 if v >= 0 else -ymax * 0.04
            ax.text(b.get_x() + b.get_width()/2, v + yoff, label,
                    ha="center", va="bottom" if v >= 0 else "top", fontsize=9)

    fig.tight_layout()
    fig.savefig(output_path, bbox_inches="tight")
    plt.close(fig)


def chart_mc_distribution(ml_results, output_path):
    """Error bar plot of Monte Carlo CIs at Y5 and Y10."""
    mc = ml_results["monte_carlo"]
    horizons = ["Year 5", "Year 10"]
    means = [mc["y5_mean"] / 1e6, mc["y10_mean"] / 1e6]
    los = [mc["y5_lo"] / 1e6, mc["y10_lo"] / 1e6]
    his = [mc["y5_hi"] / 1e6, mc["y10_hi"] / 1e6]
    err_lo = [m - l for m, l in zip(means, los)]
    err_hi = [h - m for h, m in zip(his, means)]

    fig, ax = plt.subplots(figsize=(7, 5))
    x = np.arange(len(horizons))
    ax.errorbar(x, means, yerr=[err_lo, err_hi], fmt="o", markersize=12,
                color=CHART_PALETTE["ml"], ecolor=CHART_PALETTE["ml"], capsize=8,
                capthick=2, elinewidth=2, label="MC mean ± 95% CI")
    # Band
    for i, (m, l, h) in enumerate(zip(means, los, his)):
        ax.fill_between([i - 0.15, i + 0.15], [l, l], [h, h],
                        color=CHART_PALETTE["ci_band"], alpha=0.5)
        ax.text(i + 0.18, m, f"${m:,.1f}M", va="center", fontsize=10, fontweight="bold")
        ax.text(i + 0.18, l, f"${l:,.1f}M", va="center", fontsize=8, color="gray")
        ax.text(i + 0.18, h, f"${h:,.1f}M", va="center", fontsize=8, color="gray")
    ax.set_xticks(x)
    ax.set_xticklabels(horizons)
    ax.set_ylabel("Δ Imports (USD millions)")
    ax.set_title(f"XGBoost Monte Carlo 95% Confidence Intervals\n"
                 f"({mc['n_draws']} draws, perturbing GDP growth and phase-out speed)",
                 fontweight="bold", pad=12)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:,.0f}M"))
    ax.legend(loc="upper left", frameon=True)
    ax.set_xlim(-0.5, len(horizons) - 0.3)

    fig.tight_layout()
    fig.savefig(output_path, bbox_inches="tight")
    plt.close(fig)


def chart_decomposition(s1_results, greek_agri_share, f2f_drag, output_path):
    """Waterfall: Add Imp, then F2F-adjusted Add Exp, then Widening (Y10 S1)."""
    add_imp = sum(c["years"][10]["proj_imp"] - c["years"][10]["cf_imp"] for c in s1_results) / 1e6
    add_exp_unadj = sum(c["years"][10]["proj_exp"] - c["years"][10]["cf_exp"] for c in s1_results) / 1e6
    f2f_loss = add_exp_unadj * greek_agri_share * f2f_drag
    add_exp_adj = add_exp_unadj - f2f_loss
    widening = add_imp - add_exp_adj

    labels = ["Additional\nImports", "Additional\nExports\n(unadjusted)",
              "F2F\ncapacity loss", "Adjusted\nExports", "Bilateral\nWidening"]
    values = [add_imp, add_exp_unadj, -f2f_loss, add_exp_adj, widening]
    colors = [CHART_PALETTE["imports"], CHART_PALETTE["exports"],
              CHART_PALETTE["f2f"], CHART_PALETTE["exports"], CHART_PALETTE["widening"]]

    fig, ax = plt.subplots(figsize=(9, 5))
    x = np.arange(len(labels))
    bars = ax.bar(x, values, color=colors, alpha=0.85, edgecolor="black", linewidth=0.5)
    # Make negative bar (F2F loss) hatched
    bars[2].set_hatch("///")

    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=9)
    ax.set_ylabel("USD millions")
    ax.set_title("Year 10 Decomposition: How Bilateral Widening Builds Up (Central Case, S1)",
                 fontweight="bold", pad=12)
    ax.axhline(0, color="black", linewidth=0.5)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:,.0f}M"))

    for i, v in enumerate(values):
        label = f"${v:,.0f}M"
        if v >= 0:
            ax.text(i, v + max(values) * 0.02, label, ha="center", va="bottom",
                    fontsize=10, fontweight="bold")
        else:
            ax.text(i, v - max(values) * 0.04, label, ha="center", va="top",
                    fontsize=10, fontweight="bold", color=CHART_PALETTE["f2f"])

    # Add an arrow from imports to widening
    ax.annotate("", xy=(4, widening), xytext=(0, add_imp * 0.9),
                arrowprops=dict(arrowstyle="->", color="gray", alpha=0.4, lw=1.5))
    ax.text(2, add_imp * 1.05,
            f"Imports (${add_imp:,.0f}M) minus Adjusted Exports (${add_exp_adj:,.0f}M) = "
            f"${widening:,.0f}M widening",
            ha="center", fontsize=9, style="italic", color="gray")

    fig.tight_layout()
    fig.savefig(output_path, bbox_inches="tight")
    plt.close(fig)


def chart_yearly_trajectory(s1_results, greek_agri_share, f2f_drag, output_path):
    """Time series of Add Imp, Adj Add Exp, and Widening over 11 years."""
    years = list(range(11))
    add_imp = [sum(c["years"][y]["proj_imp"] - c["years"][y]["cf_imp"] for c in s1_results) / 1e6
               for y in years]
    add_exp_unadj = [sum(c["years"][y]["proj_exp"] - c["years"][y]["cf_exp"] for c in s1_results) / 1e6
                     for y in years]
    # F2F ramps linearly from 0 at Y0 to f2f_drag * greek_agri_share at Y10
    add_exp_adj = [e * (1 - greek_agri_share * f2f_drag * (y / 10))
                   for y, e in zip(years, add_exp_unadj)]
    widening = [i - e for i, e in zip(add_imp, add_exp_adj)]

    calendar_years = [2026 + y for y in years]

    fig, ax = plt.subplots(figsize=(9, 5))
    ax.plot(calendar_years, add_imp, marker="o", color=CHART_PALETTE["imports"],
            linewidth=2.5, label="Additional Imports", markersize=7)
    ax.plot(calendar_years, add_exp_adj, marker="s", color=CHART_PALETTE["exports"],
            linewidth=2.5, label="Additional Exports (F2F-adjusted)", markersize=7)
    ax.plot(calendar_years, widening, marker="^", color=CHART_PALETTE["widening"],
            linewidth=2.5, label="Bilateral Widening (Imp - Adj Exp)", markersize=7)
    ax.fill_between(calendar_years, add_imp, add_exp_adj,
                    color=CHART_PALETTE["widening"], alpha=0.15, label="_widening_band")

    ax.set_xlabel("Year")
    ax.set_ylabel("USD millions")
    ax.set_title("PE Armington Annual Trajectory, S1 (entry 2026)",
                 fontweight="bold", pad=12)
    ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:,.0f}M"))
    ax.legend(loc="upper left", frameon=True)
    ax.set_xticks(calendar_years)
    ax.set_xticklabels([str(y) for y in calendar_years], rotation=0)

    # Annotate Y10 endpoint
    ax.text(calendar_years[-1] + 0.1, widening[-1], f"${widening[-1]:,.0f}M",
            color=CHART_PALETTE["widening"], fontweight="bold", va="center", fontsize=10)

    fig.tight_layout()
    fig.savefig(output_path, bbox_inches="tight")
    plt.close(fig)


def write_all_charts(s1_results, sens_grid_s1, ml_results, pe_mc, greek_agri_share, f2f_drag, output_dir):
    """Generate all charts to output_dir."""
    charts_dir = os.path.join(output_dir, "charts")
    os.makedirs(charts_dir, exist_ok=True)
    paths = {}
    chart_feature_importance(ml_results, os.path.join(charts_dir, "01_feature_importance.png"))
    paths["feature_importance"] = "charts/01_feature_importance.png"
    chart_sensitivity_heatmap(sens_grid_s1, os.path.join(charts_dir, "02_sensitivity_heatmap.png"))
    paths["sensitivity_heatmap"] = "charts/02_sensitivity_heatmap.png"
    chart_ml_vs_pe(s1_results, ml_results, greek_agri_share, f2f_drag,
                   os.path.join(charts_dir, "03_ml_vs_pe.png"))
    paths["ml_vs_pe"] = "charts/03_ml_vs_pe.png"
    chart_mc_distribution(ml_results, os.path.join(charts_dir, "04_mc_distribution.png"))
    paths["mc_distribution"] = "charts/04_mc_distribution.png"
    chart_decomposition(s1_results, greek_agri_share, f2f_drag,
                        os.path.join(charts_dir, "05_decomposition.png"))
    paths["decomposition"] = "charts/05_decomposition.png"
    chart_yearly_trajectory(s1_results, greek_agri_share, f2f_drag,
                            os.path.join(charts_dir, "06_yearly_trajectory.png"))
    paths["yearly_trajectory"] = "charts/06_yearly_trajectory.png"
    chart_pe_monte_carlo(pe_mc, os.path.join(charts_dir, "07_pe_monte_carlo.png"))
    paths["pe_monte_carlo"] = "charts/07_pe_monte_carlo.png"
    return paths


# ====================================================================
# MODULE 6: LITERATURE ANCHORS SHEET
# Documents which paper provides the empirical anchor for each model parameter.
# ====================================================================

LITERATURE_ANCHORS = [
    # (paper_short, year, citation, finding, number, maps_to, status, notes)
    {
        "paper": "KEPE Greek Economic Outlook",
        "year": "2025",
        "citation": "Reziti, I. (2025). The impact of external crises on Greece's agricultural economy: A production and income analysis for the period 2019-2023. KEPE, Greek Economic Outlook 2025/57, pp. 68-76.",
        "finding": "Greek input cost share of output rose from 48.8% (2019) to 55.3% (2023), an increase of 6.5pp. EU-27 same-period rise was only 1.0pp (57.3% to 58.3%). Net asymmetric burden on Greek producers: 5.5pp.",
        "number": "5.5pp asymmetric cost burden -> 3.0x wedge upper bound",
        "maps_to": "Sensitivity grid wedge multiplier upper bound (Sheet 8) AND PE Monte Carlo wedge distribution upper bound (Sheet 9). The 5.5pp burden divergence over 5 years projects to ~10pp over 10 years; we conservatively set 3.0x (3% wedge) as the upper bound, which is roughly half the projected pace.",
        "status": "ANCHORS sensitivity upper bound and PE MC distribution",
        "notes": "Tier 2: KEPE 5.5pp now drives a concrete model parameter (3.0x wedge cap) instead of being purely documentary. Reduces the gap between KEPE empirical evidence and the model's wedge specification.",
    },
    {
        "paper": "KEPE Greek Economic Outlook",
        "year": "2025",
        "citation": "Reziti, I. (2025). KEPE 2025/57, Table 4.2.5.",
        "finding": "Greek agricultural productivity collapsed from 0.98 (2019) to 0.66 (2023), a 33% decline. EU-27 productivity remained more stable. Volume index of GVA fell from 102.76 to 69.99 (2015 = 100).",
        "number": "33% productivity decline 2019-2023",
        "maps_to": "F2F drag (Greek agri export capacity factor). The observed 33% productivity decline across 5 years gives an upper bound; our 10% drag by Y10 is conservative within this range.",
        "status": "VALIDATES the F2F drag",
        "notes": "Our 10% F2F drag is roughly 1/3 of the observed 5-year productivity decline, defensibly conservative.",
    },
    {
        "paper": "KEPE Greek Economic Outlook",
        "year": "2025",
        "citation": "Reziti, I. (2025). KEPE 2025/57, section 4.2.4.",
        "finding": "Greek 2023 input cost shocks: labor wages +9.4%, capital costs +4.5%, machinery rental +3.9%, loan interest +5.3%, land rents +3.5%.",
        "number": "Annual cost shock magnitudes -> drives PE MC width",
        "maps_to": "PE Armington Monte Carlo (Sheet 9, Tier 2). The KEPE-observed annual variability informs the width of the F2F drag and wedge perturbation distributions.",
        "status": "ANCHORS PE Monte Carlo distribution width",
        "notes": "Tier 2: previously AVAILABLE, now used. The triangular(0.5, 1.0, 3.0) wedge distribution upper bound and triangular(0.05, 0.10, 0.20) F2F distribution width reflect the empirical volatility KEPE documents.",
    },
    {
        "paper": "Mandanas et al. (Proceedings MDPI)",
        "year": "2026",
        "citation": "Mandanas, Z.C., Petropoulos, D.P., Apostolopoulos, N. (2026). The Impact of CAP Investment Subsidies on Agricultural Productivity in Greece: A Time-Series Analysis. Proceedings 2026, 134, 6.",
        "finding": "VAR(1) elasticity of Greek agricultural productivity to CAP investment subsidies = 0.35 (p = 0.002). Granger-causal at p = 0.020. Productivity persistence coefficient = 0.62. Estimated on Greek time-series 2000-2023.",
        "number": "0.35 elasticity x 30% eco-scheme conditional share = 10.5% F2F drag",
        "maps_to": "F2F drag (Greek agri export capacity factor). Derivation: 32% of Greek CAP CSP budget is conditional on eco-scheme compliance (Kostas's email). If farmers cannot fully transition, they lose access to ~30% of their effective subsidy. Mandanas elasticity 0.35 then implies productivity drop of 0.35 x 30% = 10.5%, matching our 10% central F2F drag.",
        "status": "ANCHORS the F2F drag (with explicit derivation)",
        "notes": "Tier 2: now provides primary empirical anchor (replaces VALIDATES status). The 10% F2F drag is no longer just a borrow from Beckman/Bremmer EU-wide projections; it is independently derived from Greek-specific time-series econometrics.",
    },
    {
        "paper": "Sustainability (MDPI)",
        "year": "2025",
        "citation": "MDPI Sustainability 2025, 17, 11135. (Comparative analysis of Mercosur and EU agricultural production factors.)",
        "finding": "Mercosur agricultural labor cost: Brazilian agri wages are 25% of Polish, 15% of Italian, <10% of Belgian/Danish/French/German/Dutch wages. Brazilian diesel: ~33% lower than Polish, ~50% lower than other EU. Mercosur agricultural value added per worker: 67.5% lower than EU (2023).",
        "number": "Mercosur cost advantage: labor 75-90% lower, diesel 33-50% lower",
        "maps_to": "Cost wedge (asymmetric production cost variable). Direct empirical justification that Mercosur producers operate at materially lower production cost than EU/Greek competitors, which is what the wedge captures.",
        "status": "ANCHORS the wedge",
        "notes": "This is the most direct empirical answer to Kostas's request for an asymmetric production cost variable. Mercosur cost advantage is empirically very large.",
    },
    {
        "paper": "AGRI Committee CASP study",
        "year": "2026",
        "citation": "European Parliament, Committee on Agriculture and Rural Development. (2026). Support measures for farmers' income in different Member States in the context of inflation and rising production costs. CASP_STU(2026)759349.",
        "finding": "Of 188 billion EUR direct payments under CAP 2023-2027 (37.5 billion EUR per annum), 65% allocated to income support (SO1), 35% to other objectives (sustainability, competitiveness). Eco-scheme payments mainly cover opportunity costs of environmental actions, providing only limited net income support.",
        "number": "65/35 income support split; eco-schemes are opportunity-cost only",
        "maps_to": "F2F drag interpretation. Justifies why eco-scheme conditionality reduces effective Greek production capacity rather than fully offsetting compliance costs.",
        "status": "CONTEXT for F2F interpretation",
        "notes": "Confirms that even with CAP subsidies, eco-compliant Greek producers do not fully recover the cost of environmental compliance, supporting the asymmetric burden hypothesis.",
    },
    {
        "paper": "ELSTAT Environmental Tax Statistics",
        "year": "2024",
        "citation": "Hellenic Statistical Authority (ELSTAT). Environmental Taxes by Sector (2022, 2023). https://www.statistics.gr",
        "finding": "Greek agriculture, forestry, and fishing paid 410.9 million EUR in environmental taxes in 2023 (4.4% of total Greek environmental tax revenue). 81.5% of total Greek environmental taxes are energy taxes.",
        "number": "410.9M EUR Greek agri environmental tax burden (2023)",
        "maps_to": "Cost wedge (asymmetric production cost variable). The ELSTAT figure is the symptom; the wedge captures the mechanism (Greek producers pay an environmental tax burden Mercosur producers do not).",
        "status": "VALIDATES the wedge mechanism",
        "notes": "Kostas's specific motivating evidence. The wedge is calibrated from EUDR/CBAM/EC compliance studies that capture the SAME mechanism as a percentage of trade rather than absolute euros.",
    },
    {
        "paper": "Profundo (compliance cost study)",
        "year": "2025",
        "citation": "Profundo (Feb 2025). EUDR Compliance Costs: Economic analysis.",
        "finding": "EUDR compliance adds approximately 1% to per-unit production cost for affected commodities (animal feed, soy derivatives, palm oil, beef, coffee, rubber, wood products).",
        "number": "1% compliance cost premium",
        "maps_to": "Cost wedge for HS 23 (animal feed/oilseed cake), HS 12 (oilseeds), HS 09 (coffee), HS 47 (pulp).",
        "status": "ANCHORS wedge for HS 23, 12, 09, 47",
        "notes": "Already used in v6 model.",
    },
    {
        "paper": "IISD / Frontier Economics (CBAM study)",
        "year": "2026",
        "citation": "IISD and Frontier Economics (Jan 2026). CBAM impact assessment for affected sectors.",
        "finding": "CBAM transition compliance adds approximately 1% production cost for steel, aluminum, fertilizer, and downstream products.",
        "number": "1% compliance cost premium",
        "maps_to": "Cost wedge for HS 76 (aluminum), HS 73 (iron/steel), HS 38 (chemicals).",
        "status": "ANCHORS wedge for HS 76, 73, 38",
        "notes": "Already used in v6 model.",
    },
    {
        "paper": "European Commission (farmer compliance)",
        "year": "2014",
        "citation": "European Commission (2014). Assessment of the cost of farmer compliance with EU regulations.",
        "finding": "EU producer baseline compliance cost burden (good agricultural practices, traceability, food safety): approximately 1% of value of production.",
        "number": "1% baseline compliance cost",
        "maps_to": "Cost wedge for remaining HS chapters not covered by EUDR or CBAM.",
        "status": "ANCHORS wedge for remaining chapters",
        "notes": "Already used in v6 model. Conservative baseline for chapters where no specific compliance regime applies.",
    },
    {
        "paper": "USDA ERS (Beckman et al.)",
        "year": "2020",
        "citation": "Beckman, J., Ivanic, M., Jelliffe, J., Baquedano, F., Scott, S. (2020). Economic and Food Security Impacts of Agricultural Input Reduction Under the European Union Green Deal's Farm to Fork and Biodiversity Strategies. USDA Economic Research Service.",
        "finding": "EU agricultural production projected to decline 7-12% by 2030 across three F2F adoption scenarios.",
        "number": "7-12% EU agri output decline (central 9.5%)",
        "maps_to": "F2F drag (Greek agri export capacity factor). Used central case of 10% by Y10.",
        "status": "ANCHORS the F2F drag",
        "notes": "Already used in v6 model.",
    },
    {
        "paper": "Wageningen (Bremmer et al.)",
        "year": "2021",
        "citation": "Bremmer, J., Gonzalez-Martinez, A., Jongeneel, R., Huiting, H., Stokkers, R. (2021). Wageningen University AGMEMOD assessment.",
        "finding": "F2F adoption: 10-20% production decline expected, with up to 25% yield loss for olives and 2% for maize. EUR 92 billion EU production value reduction by 2030.",
        "number": "10-20% production decline",
        "maps_to": "F2F drag, sensitivity grid upper bound (20% F2F level).",
        "status": "ANCHORS the F2F drag and sensitivity range",
        "notes": "Already used in v6 model. The 20% upper bound in the sensitivity grid is the Wageningen high-impact scenario.",
    },
    {
        "paper": "Labrianidis & Sykas",
        "year": "2009",
        "citation": "Labrianidis, L., Sykas, T. (2009). Geographical proximity and immigrant labour in agriculture: Albanian immigrants in the Greek countryside. Sociologia Ruralis 49(4): 394-414.",
        "finding": "Albanian migrants accounted for approximately 20% of Greek agricultural labor force in the 2000s, rising to 33% in intensive crops (fruit production, greenhouse vegetables, olive harvests). This labor source has steadily declined as migrants have moved to other sectors or returned home.",
        "number": "20-33% of Greek agri labor was migrant-sourced",
        "maps_to": "Labor drag (Tier 3, separate from F2F drag). Central case 5% capacity drag at Y10 reflects partial loss of migrant labor and unfilled replacement; sensitivity range 0-10% in PE Monte Carlo.",
        "status": "ANCHORS the labor drag",
        "notes": "Tier 3: Provides the empirical basis for the labor drag parameter Kostas explicitly requested ('Labor Constraint Index'). Greek agriculture's historical dependence on migrant labor and the documented decline of that supply justifies treating labor as a separate constraint.",
    },
    {
        "paper": "Maro et al. systematic review",
        "year": "2025",
        "citation": "Maro, Z.M., Borda, A., Balogh, J. (2025). Challenges and potential solutions to employment issues in the agri-food sector of developed countries. Sustainable Futures 10: 100895.",
        "finding": "Systematic review of 128 studies on agri-food employment in developed countries. EU agriculture is heavily dependent on migrant and seasonal labor; supply has declined post-COVID and remains constrained by visa restrictions, demographic shifts, and competition from other sectors. Mechanization is partial substitute but slow.",
        "number": "Confirmed declining migrant labor supply trend",
        "maps_to": "Labor drag (Tier 3). Validates that the Labrianidis & Sykas 2009 baseline number is not just historical but continues to decline through the projection horizon.",
        "status": "VALIDATES the labor drag direction",
        "notes": "Tier 3: Recent (post-COVID) confirmation that the labor constraint is structural, not transient. Combined with Labrianidis & Sykas, supports the triangular(0.00, 0.05, 0.10) labor drag distribution in PE Monte Carlo.",
    },
    {
        "paper": "may_5_email_data.docx (CAP-HS mapping)",
        "year": "2025",
        "citation": "Internal team analysis. Mapping of Greek CAP Strategic Plan 2023-2027 intervention codes to HS chapters.",
        "finding": "Per-HS-chapter analysis of which Greek agricultural products are subject to CAP Strategic Plan eco-scheme conditionality. HS 04 (dairy/honey, GI Feta) and HS 15 (fats/oils, GI Kalamata olive oil) and HS 20 (prep olives) are fully agri and CAP-conditional (cap_weight = 1.0). HS 22 (beverages, GI wines) is 0.9. HS 09 (coffee/tea/spices) is 0.4 because Greek production is mostly aromatic herbs not coffee. Industrial chapters (HS 27, 73, 76, 84, 85, etc.) are 0.0.",
        "number": "Per-chapter cap_weight (0 to 1) for all 25 HS chapters",
        "maps_to": "Per-chapter Greek capacity adjustment (Sheet 13). Replaces uniform 18.3% Greek agri share with chapter-specific weights. Total Y10 loss with per-chapter approach: ~$32M vs ~$10M with uniform. The difference reflects that bilateral Greek-Mercosur exports are more concentrated in agri chapters than the economy-wide 18.3% would imply.",
        "status": "ANCHORS the per-chapter cap_weights",
        "notes": "Tier 3: Replaces the Tier 1/2 uniform agri-share assumption with chapter-specific exposure to CAP eco-scheme conditionality, F2F drag, and Greek labor constraints. Addresses Kostas's request that the model treat each product category according to its specific CAP eligibility profile.",
    },
    {
        "paper": "Tzouramani et al.",
        "year": "2020",
        "citation": "Tzouramani, I., Mantziaris, S., Karanikolas, P. (2020). Assessing Sustainability Performance at the Farm Level: Examples from Greek Agricultural Systems. Sustainability 12(7): 2929.",
        "finding": "Average age of Greek farm managers: 51.7 years across all studied farm types, ranging 50.2 (arable crops) to 53.6 (livestock). Female managers more common in olive (42%) and permanent crop (37%) farms than arable (16%) or livestock (16%). Three typical Mediterranean farming systems (permanent crops, olive trees, sheep) rated more sustainable than arable crops via AHP analysis.",
        "number": "Average Greek farm manager age 51.7 years (range 50.2-53.6)",
        "maps_to": "Labor drag (validates the demographic context). Average age above 50 across all Greek farm types confirms the aging-workforce mechanism behind the labor drag parameter. Differential sustainability ranking (olives, permanent crops, sheep > arable) validates differential cap_weight treatment (HS 15 olive oil w=1.0, HS 04 dairy w=1.0, HS 20 olives w=1.0 vs lower weights for arable-derived chapters).",
        "status": "VALIDATES labor drag demographic context",
        "notes": "Tier 3: Greek-specific demographic data confirming the aging-workforce premise behind the labor drag parameter. The study's finding that Mediterranean farming systems (olives, permanent crops, sheep) are more sustainable than arable also supports the model's emphasis on GI-protected products as Greece's resilient export channel.",
    },
    {
        "paper": "Barnes (Journal of Agricultural Economics)",
        "year": "2023",
        "citation": "Barnes, A.P. (2023). The role of family life-cycle events on persistent and transient inefficiencies in less favoured areas. Journal of Agricultural Economics 74: 295-315.",
        "finding": "Multi-step stochastic frontier model on Scottish LFA cattle/sheep farms (2003-2020). Average family age coefficient: 0.026 (NOT significant). Succession plan in place: -0.324 on persistent inefficiency (highly significant). Higher agricultural education: -0.380 on persistent inefficiency. Off-farm revenue share: +0.394 on persistent inefficiency. Subsidy share: +0.683 on persistent inefficiency.",
        "number": "Family age effect insignificant (coef 0.026, ns)",
        "maps_to": "Demographics handling (model design decision). Barnes finds that raw farmer age is NOT a significant determinant of farm efficiency once family structure variables are controlled for. This justifies the model's choice to handle demographics narratively rather than as a quantitative variable: the simple aging-decline narrative is not supported by the strongest available efficiency study.",
        "status": "CONTEXT for demographics design decision",
        "notes": "Tier 3: Provides scholarly basis for NOT adding a raw demographic variable to the model. Barnes shows farmer age proxies are mixed and often insignificant; succession planning and education matter more, but those are not in our trade dataset. Strengthens the manuscript's narrative-only treatment of demographics.",
    },
    {
        "paper": "Borda, Sarvari, and Balogh",
        "year": "2023",
        "citation": "Borda, A.J., Sarvari, B., Balogh, J.M. (2023). Generation Change in Agriculture: A Systematic Review of the Literature. Economies 11: 129. https://doi.org/10.3390/economies11050129",
        "finding": "Systematic review of 55 articles on generational change in agriculture (1990-2022). EU average farmer age rose from 49.2 (2004) to 51.4 (2014), an increase of 2.2 years per decade. In 2016, 35% of EU farmers were >= 65 years old; only 11% were < 40 years old. US farmer average age 57.21 years (2017 Census), up 1.2 years from 2012. Identifies main barriers to generational renewal: low income, limited land access, administrative burdens, weak bargaining position, and climate uncertainty. Notes that the Young Farmers Scheme in Greek agriculture (Gkatsikos et al. 2022) stimulates regional output and supports generational renewal.",
        "number": "EU average farmer age 51.4 years (2014), growing 2.2 years per decade",
        "maps_to": "Labor drag (Tier 3). Provides EU-wide comparative context for Tzouramani's Greek-specific number (51.7 years). Confirms the aging trend is structural, EU-wide, and accelerating. Validates the labor drag distribution Triangular(0.00, 0.05, 0.10) as a forward-looking projection of an existing trend rather than a hypothetical scenario.",
        "status": "VALIDATES labor drag (EU-wide demographic context)",
        "notes": "Tier 3: Closes the last gap from Kostas's 7-paper list. The systematic review's identification of barriers to generational renewal (low income, land access, administrative burdens) all reinforce the structural nature of Greek agri labor decline that the labor drag captures.",
    },
    # ===== v7 additions: ELSTAT 2020 Census and FADN 2019-2023 =====
    {
        "paper": "ELSTAT 2020 Agricultural-Livestock Census (Employment, Table 43)",
        "year": "2020",
        "citation": "Hellenic Statistical Authority (ELSTAT). 2020 Agricultural-Livestock Census. Table 43: Number of persons employed in agricultural-livestock holdings, by region. Athens, ELSTAT, 2022. Greece total: 2,106,534 persons; 463,320 holdings; 97,173,709 workdays. Seasonal workers: 794,811 persons; permanent workers: 27,350; family labor: 722,973.",
        "finding": "Greek agri labor force structure as of 2020: family labor 78.6% of workdays, permanent paid 5.9%, seasonal paid 12.8%, mutual help 1.8%, other 0.9%. Seasonal workers are 57% of non-family headcount and structurally vulnerable to migration policy, demographic shift, and post-COVID supply.",
        "number": "Seasonal workers = 12.8% of workdays, 57% of non-family headcount",
        "maps_to": "labor_drag (Tier 3, v7 update). Replaces Labrianidis & Sykas 2009 as the primary anchor with current, Greek-specific Census data. Census 2020 supports raising central labor_drag from 0.05 to 0.07 and widening Monte Carlo distribution from Triangular(0.00, 0.05, 0.10) to Triangular(0.03, 0.07, 0.12).",
        "status": "ANCHORS the labor drag (v7 primary anchor)",
        "notes": "v7 update: This is the most defensible empirical foundation for the labor_drag parameter. The 2009 Labrianidis & Sykas paper remains in the model as historical context, but the 2020 Census is the new primary anchor.",
    },
    {
        "paper": "ELSTAT 2020 Agricultural Census (Holdings, UAA, and crops, Tables B01, 7, 9, 10, 17, 22, 27)",
        "year": "2020",
        "citation": "Hellenic Statistical Authority (ELSTAT). 2020 Agricultural-Livestock Census. Tables B01 (holdings and UAA by region), 7 (land use), 9 (annual crops), 10 (cereals), 17 (tree crops), 22 (vineyards), 27 (irrigated areas). Athens, ELSTAT, 2022. Greece total: 525,284 holdings with UAA covering 2,822,886 hectares.",
        "finding": "Greek UAA composition 2020 (denominator: 2,822,886 ha total UAA from Table B01): 46.4% arable (incl. cereals 20.8% of UAA), 25.9% permanent tree crops (incl. olives 20.7% of UAA), 2.1% vineyards, 25.6% other (fallow, meadows, kitchen gardens). Olives for oil 522k ha (18.5% of UAA); table olives 62k ha (2.2%); fruit and nuts categories combined for HS 08 mapping: temperate fruit 64k + tropical 14k + nuts 33k + citrus 35k + other perennials 4k + table grapes 11k + raisins 13k = 174k ha (6.2% of UAA); PDO+PGI vineyards 15k ha (0.5%); non-table wine vineyards 35k ha (1.2%).",
        "number": "Olives 20.7%, fruit-and-nuts 6.2%, vineyards 2.1% of UAA",
        "maps_to": "cap_weights for HS 04, 08, 15, 20, 22. UAA shares empirically validate the v6 cap_weights: HS 15 (1.0) anchored on 522k ha olives for oil (largest tree-crop contributor, 18.5% of UAA); HS 20 (1.0) on 62k ha table olives; HS 22 (0.9) on 35k ha non-table-grape vineyards; HS 08 (1.0) on 174k ha across temperate fruit, tropical fruit, nuts, citrus, other perennials, table grapes, and raisins.",
        "status": "ANCHORS the cap_weights (v7 empirical validation)",
        "notes": "v7 update: Replaces judgment-based cap_weights with UAA-share-validated weights. The v6 weights largely stand because the empirical UAA shares align with the qualitative assignments.",
    },
    {
        "paper": "ELSTAT 2020 Livestock Census (Tables E05, 35A)",
        "year": "2020",
        "citation": "Hellenic Statistical Authority (ELSTAT). 2020 Agricultural-Livestock Census. Tables E05 (livestock by region) and 35A (animals by kind). Athens, ELSTAT, 2022. Greece total: 624,397 bovine animals; 7,721,799 sheep; 3,149,008 goats; 742,963 pigs.",
        "finding": "Greek livestock is overwhelmingly small-ruminant: 10.87M sheep+goats across 93,739 holdings, vs 624k bovines and 743k pigs. Sheep/goat dominance is the structural basis for the Feta PDO regulation (Greek sheep/goat milk only).",
        "number": "Sheep + Goats = 10.87M animals (94% of livestock by head)",
        "maps_to": "cap_weight = 1.00 for HS 04 (Dairy, eggs, honey). Greek-specific livestock structure validates the maximum cap_weight given Feta PDO regulatory tie to Greek sheep/goat milk.",
        "status": "ANCHORS the HS 04 cap_weight (v7 addition)",
        "notes": "v7 update: Explicit empirical basis for the HS 04 cap_weight maximum. Replaces 'GI Feta argument' with hard census numbers.",
    },
    {
        "paper": "FADN Farm Economy Focus, Greece 2019-2023",
        "year": "2019-2023",
        "citation": "European Commission, DG AGRI. Farm Accountancy Data Network (FADN) Farm Economy Focus dashboard. Greece, member-state level, all farm types, IFS 2020 reference, 5-year time series 2019-2023. Data updated 2026-01-28.",
        "finding": "Per representative Greek farm 2023: total output EUR 32,929; intermediate consumption EUR 16,910; balance of current subsidies EUR 4,957; gross farm income EUR 20,976; farm net value added EUR 17,803; total subsidies excluding investment EUR 5,461; direct payments EUR 4,376. Population-level: 523,900 farms; 2.80M ha UAA; 1.87M LSU; 565,400 AWU; EUR 7.60bn standard output. UAA cross-validates ELSTAT to within 0.5%.",
        "number": "EUR 7.60bn Greek standard output (2023); EUR 4.4k direct payments per farm",
        "maps_to": "Validation of Greek agri scale and cap_weight calibration. Per-farm-type splits (Fieldcrops, Horticulture, Wine) confirm the chapter mapping. Direct-payment decline 2019-2023 from EUR 5,522 to EUR 4,376 reflects the 2023-2027 CAP reform.",
        "status": "VALIDATES Greek agri scale (v7 addition)",
        "notes": "v7 update: FADN does not replace the EU-vs-Mercosur wedge sources (it measures Greek-only quantities), but it provides an independent view of Greek agri size and trajectory that sits alongside the trade-data view driving the model.",
    },
    {
        "paper": "OECD Environmental Tax indicator (Greece vs OECD vs LAC)",
        "year": "2022 (Greece/OECD); 2019 (LAC benchmark)",
        "citation": "OECD. Environmental tax indicator: % of GDP, 2022. https://www.oecd.org/en/data/indicators/environmental-tax.html. LAC regional benchmark from OECD/CIAT/ECLAC/IDB. Revenue Statistics in Latin America and the Caribbean 2021.",
        "finding": "Greece's environmental tax burden is the HIGHEST in the OECD at 3.6% of GDP (2022, OECD dashboard) and 3.8% (2020, OECD WKP-2023-23, independent cross-check). The OECD-wide average is approximately 2.1% of GDP per OECD/CIAT 2021 Revenue Statistics LAC (the highlighted 'OECD' bar on the dashboard at ~1.2% reflects only selected filter countries, not the full OECD). EU27 average was 2.2% of GDP in 2020 per IMF/OECD. Mercosur countries are not OECD members and country-specific environmentally-related tax revenue data was not located in primary sources during this build. The most defensible Mercosur proxy is the LAC regional average of 1.2% of GDP (2019, 25 LAC countries, OECD/CIAT/ECLAC/IDB Revenue Statistics LAC 2021). Implied structural environmental-tax gap between Greek producers and Mercosur producers is approximately 2.4 percentage points of GDP-equivalent burden.",
        "number": "Greece env tax 3.6% of GDP (rank #1 OECD); LAC regional avg 1.2%; gap = 2.4pp",
        "maps_to": "Cost wedge component (new in v7). Adds 0.5 percentage points to chap['wedge'] for HS 23 (animal feed), the energy-intensive Greek agri processing chapter where a tariff channel exists (eu_mfn=5%). Reflects partial pass-through (~20%) of the 2.4pp env-tax gap into producer costs. Wired into the orchestrator's CONFIG loader so it propagates into PE Armington, sensitivity grid, and Monte Carlo. HS 12 (oilseeds) was considered but has eu_mfn=0, so the wedge channel is mathematically inactive (tariff_change*wedge=0); HS 12 was omitted.",
        "status": "ANCHORS new env-tax wedge component for HS 23 (v7 addition, parameter applied)",
        "notes": "v7 update: This is a previously unmodeled component of the EU-vs-Mercosur cost asymmetry. The 0.5pp pass-through is conservative; full pass-through would imply 0.7-1.2pp depending on energy intensity. EEA Greece profile 2025 confirms the Greek env-tax share of total tax revenue is RISING since 2013 while the EU average is slightly declining, so this wedge is expected to widen over the 2026-2036 horizon.",
    },
    {
        "paper": "EEA Greece Country Profile, Europe's Environment 2025",
        "year": "2025",
        "citation": "European Environment Agency. Greece - Country profiles - Europe's environment 2025. Published 28 September 2025. https://www.eea.europa.eu/en/europe-environment-2025/countries/greece",
        "finding": "Greek agri-environmental profile 2025: area under organic farming 17.2% of UAA (acceleration needed for 2030); designated terrestrial protected areas 34.7% (achieved); climate-related economic losses EUR 60/capita (wrong direction); fossil fuel subsidies 0.9% of GDP (vs EU avg 0.7%); GHG emissions index 69.2 vs 1990 (no LULUCF). Eco-Innovation Index: below EU average but highest rate of increase in EU 2013-2022. Greek primary-sector value yield approximately EUR 110 per 1000 m² (EUR 1,100/ha). Olive oil consumption declining domestically, replaced by cheaper seed oils. Food innovation deployment to farmers very limited.",
        "number": "Organic farming 17.2% of UAA; Greek primary-sector yield EUR 1,100/ha",
        "maps_to": "Multiple. (1) Organic 17.2% modulates F2F_drag exposure: organic farms have lower marginal F2F burden because they already comply. Effective F2F drag for the organic share is approximately 50% of central case. (2) The EUR 1,100/ha yield is a Greek productivity floor for any farm-level scaling. (3) Climate-loss trend supports the asymmetric wedge framing in Sheet 8 sensitivity.",
        "status": "CONTEXT for F2F drag and structural framing (v7 addition)",
        "notes": "v7 update: The EEA profile is mostly narrative context but contains several quantitative anchors. The 17.2% organic share is the most actionable: it implies F2F_drag effective central case is 10% x 0.828 (non-organic share) + 5% x 0.172 (organic, half-burden) = 9.1%, a slight reduction from the 10% central. The model does NOT yet apply this adjustment; it is documented here for the team to discuss. The yield figure EUR 1,100/ha is consistent with the FADN per-farm output figures.",
    },
]


def write_sheet12_literature_anchors(wb):
    """Sheet 12: Literature anchors documenting each paper's contribution to the model."""
    ws = wb.create_sheet("12-Literature Anchors")
    ws["A1"] = "Literature Anchors: Empirical Sources for Each Model Parameter"
    style_cell(ws["A1"], font=STYLE["FONT_TITLE"], border=False)
    ws.merge_cells("A1:H1")

    ws["A2"] = ("This sheet documents which empirical paper anchors each parameter in the model. "
                "It addresses Kostas's request to ground the model in specific literature. "
                "Status column indicates whether the paper ANCHORS a current parameter, "
                "VALIDATES it from another angle, provides CONTEXT, or is AVAILABLE for future expansion.")
    ws["A2"].font = STYLE["FONT_NOTE"]
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:H2")
    ws.row_dimensions[2].height = 56

    # Header
    headers = ["Paper", "Year", "Specific Finding", "Number Extracted", "Maps to Model Parameter",
               "Status", "Notes", "Citation"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(4, i)
        c.value = h
        style_cell(c, font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"],
                   alignment=STYLE["ALIGN_CENTER"])
    ws.row_dimensions[4].height = 30

    # Color-code by status
    status_fills = {
        "ANCHORS": PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
        "VALIDATES": PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid"),
        "CONTEXT": PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid"),
        "AVAILABLE": PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"),
    }

    for i, anchor in enumerate(LITERATURE_ANCHORS):
        r = 5 + i
        ws.cell(r, 1).value = anchor["paper"]
        ws.cell(r, 2).value = anchor["year"]
        ws.cell(r, 3).value = anchor["finding"]
        ws.cell(r, 4).value = anchor["number"]
        ws.cell(r, 5).value = anchor["maps_to"]
        ws.cell(r, 6).value = anchor["status"]
        ws.cell(r, 7).value = anchor["notes"]
        ws.cell(r, 8).value = anchor["citation"]

        # Determine status fill
        status_key = anchor["status"].split()[0]
        fill = status_fills.get(status_key)
        if fill:
            for c in range(1, 9):
                ws.cell(r, c).fill = fill

        for c in range(1, 9):
            ws.cell(r, c).border = STYLE["BORDER"]
            ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")
            ws.cell(r, c).font = Font(name="Arial", size=9)
        ws.cell(r, 1).font = Font(name="Arial", size=9, bold=True)
        ws.cell(r, 6).font = Font(name="Arial", size=9, bold=True)
        ws.row_dimensions[r].height = 110

    # Status legend below the table
    legend_r = 5 + len(LITERATURE_ANCHORS) + 2
    ws.cell(legend_r, 1).value = "STATUS LEGEND"
    style_cell(ws.cell(legend_r, 1), font=STYLE["FONT_SUBTITLE"], fill=STYLE["FILL_SECTION"])
    ws.merge_cells(start_row=legend_r, start_column=1, end_row=legend_r, end_column=8)

    legend_items = [
        ("ANCHORS", "C6EFCE", "Paper is the primary empirical source for this parameter. Direct calibration."),
        ("VALIDATES", "FFEB9C", "Paper independently confirms a parameter set elsewhere; cross-validation."),
        ("CONTEXT", "DDEBF7", "Paper provides interpretation or framing but not a calibration number."),
        ("AVAILABLE", "F2F2F2", "Paper provides numbers we have not yet incorporated; Tier 2 candidate."),
    ]
    for i, (label, color, desc) in enumerate(legend_items):
        r = legend_r + 1 + i
        c1 = ws.cell(r, 1)
        c1.value = label
        c1.fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
        c1.font = Font(name="Arial", size=9, bold=True)
        c1.border = STYLE["BORDER"]
        ws.cell(r, 2).value = desc
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        ws.cell(r, 2).font = Font(name="Arial", size=9)
        ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="center")

    # Summary row
    summary_r = legend_r + len(legend_items) + 2
    ws.cell(summary_r, 1).value = "SUMMARY: Coverage of Kostas's email asks"
    style_cell(ws.cell(summary_r, 1), font=STYLE["FONT_SUBTITLE"], fill=STYLE["FILL_SECTION"])
    ws.merge_cells(start_row=summary_r, start_column=1, end_row=summary_r, end_column=8)

    summary_items = [
        ("Capacity constraint discount factor (F2F drag)",
         "ANCHORED by USDA Beckman 2020 (7-12%), Wageningen Bremmer 2021 (10-20%), Mandanas 2026 (0.35 x 30% = 10.5%)",
         "DONE"),
        ("Asymmetric production cost variable (cost wedge)",
         "ANCHORED by Profundo 2025 (EUDR 1%), IISD 2026 (CBAM 1%), EC 2014 (1% baseline). Sensitivity grid 3.0x upper anchored by KEPE 2025 (5.5pp burden differential). PE Monte Carlo distribution anchored on KEPE volatility data.",
         "DONE (Tier 2)"),
        ("Per-product CAP eligibility weighting",
         "ANCHORED by may_5_email_data.docx (HS-to-CAP intervention mapping). 25 chapters assigned cap_weight 0-1 based on agri intensity and eco-scheme exposure. See Sheet 13.",
         "DONE (Tier 3)"),
        ("Labor constraint variable (Labor Constraint Index)",
         "ANCHORED by Labrianidis & Sykas 2009 (Albanian migrants 20-33% of Greek agri labor). VALIDATED by Maro et al. 2025 systematic review (post-COVID decline). Implemented as separate labor_drag parameter (5% central, 0-10% in PE MC).",
         "DONE (Tier 3)"),
        ("Use the listed papers to inform the model",
         "All 7 papers from Kostas's email addressed: Mandanas (ANCHORS F2F), Tzouramani 2020 (VALIDATES labor drag, Greek context), Borda 2023 (VALIDATES labor drag, EU-wide trend), Maro et al. (VALIDATES labor drag direction), Barnes 2023 (CONTEXT for demographics), KEPE (ANCHORS sensitivity upper bound + PE MC), AGRI CASP (CONTEXT). All 7 of 7 papers now in Sheet 12.",
         "DONE (7 of 7)"),
        ("Demographics (median farm age, succession)",
         "Not in quantitative model. Discussed narratively in manuscript. Avoided to prevent double-counting with XGBoost year_trend and Mandanas-derived F2F drag.",
         "Narrative only"),
        ("FADN production cost time series",
         "Not in quantitative model. FADN measures Greek-only production costs; we need EU-vs-Mercosur differential, which compliance studies provide more directly.",
         "Not used"),
        ("Olive oil $7.80/L benchmark",
         "Not in quantitative model. Aggregation level mismatch (we work at HS 2-digit; this is a sub-HS-15 product).",
         "Not used"),
    ]
    for i, (ask, coverage, status) in enumerate(summary_items):
        r = summary_r + 1 + i
        ws.cell(r, 1).value = ask
        ws.cell(r, 1).font = Font(name="Arial", size=10, bold=True)
        ws.cell(r, 2).value = coverage
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        ws.cell(r, 8).value = status
        ws.cell(r, 8).font = Font(name="Arial", size=10, bold=True)
        for c in range(1, 9):
            ws.cell(r, c).border = STYLE["BORDER"]
            ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 50

    # Column widths
    for col, w in [(1, 26), (2, 8), (3, 35), (4, 24), (5, 30), (6, 16), (7, 30), (8, 35)]:
        ws.column_dimensions[get_column_letter(col)].width = w


def write_sheet13_per_chapter_adjustments(wb, s1_results, cap_weights, f2f_drag, labor_drag):
    """Sheet 13 (Tier 3): Per-Chapter Greek Capacity Adjustments.

    Documents the cap_weight assigned to each HS chapter and the resulting Greek
    capacity adjustment applied to projected exports. This addresses Kostas's
    request for per-product CAP eligibility weighting.
    """
    ws = wb.create_sheet("13-Per-Chapter Adjustments")
    ws["A1"] = "Per-Chapter Greek Capacity Adjustments (Tier 3)"
    style_cell(ws["A1"], font=STYLE["FONT_TITLE"], border=False)
    ws.merge_cells("A1:I1")

    ws["A2"] = ("Each HS chapter receives a cap_weight from 0 to 1 reflecting its exposure to "
                "Greek agri-food production constraints (CAP eco-scheme conditionality, F2F input "
                "restrictions, migrant labor shortage). Per-chapter capacity adjustment replaces "
                "the previous uniform 18.3% Greek agri-share assumption. Total drag at Y10 = "
                f"cap_weight x (F2F drag {f2f_drag*100:.0f}% + labor drag {labor_drag*100:.0f}%).")
    ws["A2"].font = STYLE["FONT_NOTE"]
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:I2")
    ws.row_dimensions[2].height = 60

    ws["A4"] = "INPUTS"
    style_cell(ws["A4"], font=STYLE["FONT_SUBTITLE"], fill=STYLE["FILL_SECTION"])
    ws.merge_cells("A4:K4")
    ws["A5"] = "F2F drag at Y10"
    ws["B5"] = f2f_drag; ws["B5"].number_format = "0.0%"
    ws["A6"] = "Labor drag at Y10"
    ws["B6"] = labor_drag; ws["B6"].number_format = "0.0%"
    ws["C5"] = "Source"
    ws["D5"] = "USDA Beckman 2020, Wageningen Bremmer 2021, Mandanas 2026"
    ws["C6"] = "Source"
    ws["D6"] = "ELSTAT 2020 Census Table 43 (v7); Tzouramani 2020; Maro 2025"

    # Header (v7: added ELSTAT UAA share + validation columns at J, K)
    headers = ["HS", "Chapter Name", "cap_weight", "Rationale", "Y10 Unadj Export ($)",
               "Total Drag", "Y10 Adj Export ($)", "Y10 Loss ($)", "Loss %",
               "ELSTAT 2020 UAA share", "ELSTAT cap_weight validation"]
    for i, h in enumerate(headers, 1):
        c = ws.cell(8, i)
        c.value = h
        style_cell(c, font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"],
                   alignment=STYLE["ALIGN_CENTER"])
    ws.row_dimensions[8].height = 30

    # Per-chapter rationales
    rationales = {
        "23": "Animal feed; agri-adjacent processed; partial CAP exposure",
        "12": "Oil seeds; primary agri but limited Greek production",
        "09": "Coffee/tea/spices; aromatic herbs (CAP-supported), no coffee",
        "26": "Ores; not agricultural",
        "47": "Wood pulp; CAP forestry only, marginal",
        "27": "Mineral fuels; not agricultural",
        "24": "Tobacco; full agri, CAP coupled support",
        "08": "Fruits and nuts; full agri, GI-relevant",
        "30": "Pharmaceuticals; not agricultural",
        "20": "Prepared vegetables/olives; full agri, GI Kalamata olives",
        "38": "Chemicals; not agricultural",
        "84": "Machinery; not agricultural",
        "39": "Plastics; not agricultural",
        "76": "Aluminium; not agricultural",
        "85": "Electrical machinery; not agricultural",
        "88": "Aircraft; not agricultural",
        "82": "Tools/cutlery; not agricultural",
        "73": "Iron/steel; not agricultural",
        "99": "Commodities n.e.c.; ambiguous classification",
        "68": "Stone/cement articles; not agricultural",
        "25": "Salt/stone; not agricultural",
        "89": "Ships; not agricultural",
        "22": "Beverages/wine; nearly full agri, GI Samos/Nemea",
        "04": "Dairy/eggs/honey; full agri, GI Feta/Graviera/Manouri",
        "15": "Fats/oils; full agri, GI Kalamata olive oil",
    }

    # ELSTAT 2020 UAA share and validation commentary per HS chapter (v7)
    elstat_validation = {
        "23": (None, "Cereals 20.8% of UAA support animal feed processing; partial empirical basis for cap_weight"),
        "12": (None, "Limited Greek oilseed production; cap_weight unchanged"),
        "09": (None, "Aromatic herb area not in census tables; cap_weight unchanged"),
        "26": (None, "Non-agri"),
        "47": (None, "Forestry not in UAA tables; non-agri for our purposes"),
        "27": (None, "Non-agri"),
        "24": (None, "Tobacco area not in tables 9-17; cap_weight unchanged"),
        "08": (0.062, "All fruit & nut categories: temperate fruit 64k + tropical 14k + nuts 33k + citrus 35k + other perennials 4k + table grapes 11k + raisins 13k = 174k ha = 6.2% of UAA; strongly validates cap_weight=1.00"),
        "30": (None, "Non-agri"),
        "20": (0.022, "Table olives 62k ha = 2.2% of UAA; validates cap_weight=1.00"),
        "38": (None, "Non-agri"),
        "84": (None, "Non-agri"),
        "39": (None, "Non-agri"),
        "76": (None, "Non-agri"),
        "85": (None, "Non-agri"),
        "88": (None, "Non-agri"),
        "82": (None, "Non-agri"),
        "73": (None, "Non-agri"),
        "99": (None, "Non-agri"),
        "68": (None, "Non-agri"),
        "25": (None, "Non-agri"),
        "89": (None, "Non-agri"),
        "22": (0.012, "Vineyards (non-table) 35k ha = 1.2% of UAA. PDO+PGI 15k ha (44% of wine vineyards); validates cap_weight=0.90"),
        "04": (None, "Sheep+Goats = 10.87M animals (Feta PDO basis); validates cap_weight=1.00"),
        "15": (0.185, "Olives for oil 522k ha = 18.5% of UAA; largest tree-crop contributor; validates cap_weight=1.00"),
    }

    # Compute per-chapter adjustments at Y10
    adj_data = adj_exports_per_chapter(s1_results, 10, f2f_drag, labor_drag, cap_weights)

    total_unadj = 0
    total_adj = 0
    for i, chap_data in enumerate(adj_data["chapters"]):
        r = 9 + i
        hs = chap_data["hs"]
        ws.cell(r, 1).value = hs
        ws.cell(r, 2).value = chap_data["name"]
        ws.cell(r, 3).value = chap_data["cap_weight"]
        ws.cell(r, 3).number_format = "0%"
        ws.cell(r, 4).value = rationales.get(hs, "")
        ws.cell(r, 5).value = chap_data["unadj_export"]
        ws.cell(r, 5).number_format = "$#,##0"
        ws.cell(r, 6).value = chap_data["drag_pct"]
        ws.cell(r, 6).number_format = "0.0%"
        ws.cell(r, 7).value = chap_data["adj_export"]
        ws.cell(r, 7).number_format = "$#,##0"
        ws.cell(r, 8).value = chap_data["loss"]
        ws.cell(r, 8).number_format = "$#,##0"
        ws.cell(r, 9).value = (chap_data["loss"] / chap_data["unadj_export"]
                                if chap_data["unadj_export"] else 0)
        ws.cell(r, 9).number_format = "0.0%"

        # v7: ELSTAT validation columns J (UAA share) and K (validation commentary)
        elstat_share, elstat_note = elstat_validation.get(hs, (None, ""))
        if elstat_share is not None:
            ws.cell(r, 10).value = elstat_share
            ws.cell(r, 10).number_format = "0.0%"
        ws.cell(r, 11).value = elstat_note

        # Highlight chapters with non-zero cap_weight
        if chap_data["cap_weight"] > 0:
            for c in range(1, 12):
                ws.cell(r, c).fill = STYLE["FILL_NEW"]
        for c in range(1, 12):
            ws.cell(r, c).border = STYLE["BORDER"]
            ws.cell(r, c).alignment = Alignment(wrap_text=True, vertical="top")
        total_unadj += chap_data["unadj_export"]
        total_adj += chap_data["adj_export"]

    # Total row
    total_r = 9 + len(adj_data["chapters"])
    ws.cell(total_r, 1).value = "TOTAL"
    style_cell(ws.cell(total_r, 1), font=STYLE["FONT_TOTAL"], fill=STYLE["FILL_TOTAL"])
    ws.cell(total_r, 5).value = total_unadj
    ws.cell(total_r, 5).number_format = "$#,##0"
    ws.cell(total_r, 7).value = total_adj
    ws.cell(total_r, 7).number_format = "$#,##0"
    ws.cell(total_r, 8).value = total_unadj - total_adj
    ws.cell(total_r, 8).number_format = "$#,##0"
    ws.cell(total_r, 9).value = (total_unadj - total_adj) / total_unadj if total_unadj else 0
    ws.cell(total_r, 9).number_format = "0.0%"
    for c in [5, 7, 8, 9]:
        style_cell(ws.cell(total_r, c), font=STYLE["FONT_TOTAL"], fill=STYLE["FILL_TOTAL"])

    # Comparison section
    cmp_r = total_r + 3
    ws.cell(cmp_r, 1).value = "COMPARISON: Per-Chapter (Tier 3) vs Uniform (Pre-Tier 3)"
    style_cell(ws.cell(cmp_r, 1), font=STYLE["FONT_SUBTITLE"], fill=STYLE["FILL_SECTION"])
    ws.merge_cells(start_row=cmp_r, start_column=1, end_row=cmp_r, end_column=11)

    GAS = 0.183  # Greek economy-wide agri share, from KEPE; retained as bias check only (v7.3 ladder)
    uniform_loss = total_unadj * GAS * f2f_drag
    agri_unadj = sum(c["unadj_export"] for c in adj_data["chapters"] if c["cap_weight"] > 0)
    tier1 = agri_unadj * f2f_drag                    # bilateral agricultural base, F2F only
    tier2 = agri_unadj * (f2f_drag + labor_drag)     # + labour drag, flat scaling
    tier3 = total_unadj - total_adj                  # census weights, combined drag
    interaction = (tier3 - tier2) - (tier3 * f2f_drag / (f2f_drag + labor_drag) - tier1)

    cmp_rows = [
        ("Method", "Effective drag formula", "Total Y10 Loss ($M)"),
        ("Economy-wide benchmark (bias check)",
         f"unadj * 0.183 * F2F; understates by {1 - uniform_loss / tier3:.0%} (aggregation bias)",
         uniform_loss),
        ("Tier 1 (bilateral agricultural base)",
         f"agri_unadj ({agri_unadj/1e6:.1f}M = {agri_unadj/total_unadj:.1%} of total) * F2F",
         tier1),
        ("Tier 2 (+ labour drag)",
         "agri_unadj * (F2F + labor), flat scaling",
         tier2),
        ("Per-chapter (Tier 3)",
         "sum(unadj_chap * cap_weight_chap * (F2F + labor))",
         tier3),
        ("Bridge",
         f"base corr {(tier1-uniform_loss)/1e6:+.2f} | labour {(tier2-tier1)/1e6:+.2f} | weights {(tier3-tier2)/1e6:+.2f}; order interaction {interaction/1e6:+.2f}M",
         tier3 - uniform_loss),
    ]
    for i, (m, formula, val) in enumerate(cmp_rows):
        r = cmp_r + 1 + i
        ws.cell(r, 1).value = m
        ws.cell(r, 2).value = formula
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        if isinstance(val, (int, float)):
            ws.cell(r, 9).value = val / 1e6
            ws.cell(r, 9).number_format = '"$"#,##0.0"M"'
        else:
            ws.cell(r, 9).value = val
        if i == 0:
            for c in [1, 9]:
                style_cell(ws.cell(r, c), font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"])
            style_cell(ws.cell(r, 2), font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"])
        else:
            for c in [1, 2, 9]:
                ws.cell(r, c).border = STYLE["BORDER"]
                if i == 3:
                    ws.cell(r, c).fill = STYLE["FILL_TOTAL"]
                    ws.cell(r, c).font = STYLE["FONT_TOTAL"]

    # Column widths (v7: extended to include J, K for ELSTAT validation)
    for col, w in [(1, 6), (2, 32), (3, 11), (4, 38), (5, 18), (6, 11), (7, 18), (8, 16), (9, 10),
                   (10, 14), (11, 50)]:
        ws.column_dimensions[get_column_letter(col)].width = w


# ====================================================================
# v7 ADDITIONS: Sheet 14 (Greek Structural Indicators) and Sheet 15
# (Discussion: Effects, Contributions, Current Situation, Future Directions)
# ====================================================================

def write_sheet14_greek_structural_indicators(wb):
    """Sheet 14 (v7): ELSTAT 2020 Census and FADN 2019-2023 structural data.

    This is the empirical record of the Greek-specific facts that anchor
    labor_drag and cap_weights. Anyone reading this sheet can see exactly
    what Greek data the model is using and trace each parameter back to
    a primary source.
    """
    ws = wb.create_sheet("14-Greek Structural Indicators")
    ws["A1"] = "Greek Structural Indicators 2020-2023: ELSTAT Census and FADN"
    style_cell(ws["A1"], font=STYLE["FONT_TITLE"], border=False)
    ws.merge_cells("A1:H1")
    ws.row_dimensions[1].height = 28

    ws["A2"] = ("This sheet records the Greek-specific structural data anchoring the labor_drag and cap_weight parameters. "
                "All ELSTAT figures are Greece-total values from the 2020 Agricultural-Livestock Census. "
                "FADN figures are per-representative-farm averages and population-level totals from the EU Farm Economy Focus "
                "dashboard for Greece, 2019-2023.")
    ws["A2"].font = STYLE["FONT_NOTE"]
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:H2")
    ws.row_dimensions[2].height = 40

    FILL_HIGHLIGHT = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    FILL_INFO = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")

    row = 4

    def section_header(r, title):
        c = ws.cell(r, 1)
        c.value = title
        style_cell(c, font=STYLE["FONT_SUBTITLE"], fill=STYLE["FILL_SECTION"])
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        ws.row_dimensions[r].height = 22
        return r + 1

    def col_headers(r, names):
        for i, h in enumerate(names, 1):
            c = ws.cell(r, i)
            c.value = h
            style_cell(c, font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"],
                       alignment=STYLE["ALIGN_CENTER"])
        ws.row_dimensions[r].height = 30
        return r + 1

    # ----- A. Employment -----
    row = section_header(row, "A. Agricultural Employment, Greece 2020 (ELSTAT Census, Table 43)")
    row = col_headers(row, ["Labor category", "Holdings using", "Persons", "% of non-family headcount",
                            "Workdays", "% of workdays", "Source / interpretation", ""])
    emp = ELSTAT_2020["employment"]
    # Non-family-headcount denominator = perm + seasonal + mutual + other = 1,383,561
    # Note: Mutual help is non-monetized peer labor (technically not paid), but is grouped under
    # "non-family headcount" since the denominator is everyone outside the
    # holders + household members category.
    PAID = 1_383_561
    emp_rows = [
        ("Holders + household members", emp["holders_household"]["holdings"], emp["holders_household"]["persons"],
         None, emp["holders_household"]["workdays"], emp["family_pct_of_total_workdays"],
         "Family labor; structurally tied to the holding. Largest workdays share."),
        ("Permanent workers", emp["permanent_workers"]["holdings"], emp["permanent_workers"]["persons"],
         emp["permanent_workers"]["persons"]/PAID, emp["permanent_workers"]["workdays"], 0.059,
         "Full-time hired labor. Small, declining over decades."),
        ("Seasonal workers", emp["seasonal_workers"]["holdings"], emp["seasonal_workers"]["persons"],
         emp["seasonal_pct_of_paid_headcount"], emp["seasonal_workers"]["workdays"], emp["seasonal_pct_of_total_workdays"],
         "Largest non-family headcount. Structurally vulnerable layer (migrant + casual). Replaces 2009 paper as labor_drag anchor."),
        ("Mutual help", emp["mutual_help"]["holdings"], emp["mutual_help"]["persons"],
         emp["mutual_help"]["persons"]/PAID, emp["mutual_help"]["workdays"], 0.018,
         "Non-monetized peer labor exchange (not paid). Declining with rural depopulation."),
        ("Other workers", emp["other_workers"]["holdings"], emp["other_workers"]["persons"],
         emp["other_workers"]["persons"]/PAID, emp["other_workers"]["workdays"], 0.009,
         "Day-labor on non-standard terms. Counts in headcount but tiny workdays share."),
        ("TOTAL (non-family headcount = 1,383,561; total persons = 2,106,534)", emp["total_holdings"], emp["total_persons"],
         1.000, emp["total_workdays"], 1.000,
         "Greece country total, 2020 Census."),
    ]
    for i, r_data in enumerate(emp_rows):
        is_total = "TOTAL" in r_data[0]
        for c_idx, v in enumerate(r_data, 1):
            cell = ws.cell(row, c_idx)
            cell.value = v
            cell.font = STYLE["FONT_TOTAL"] if is_total else Font(name="Arial", size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = STYLE["BORDER"]
            if is_total:
                cell.fill = FILL_HIGHLIGHT
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                if c_idx in (4, 6):  # percent columns
                    cell.number_format = "0.0%"
                else:
                    cell.number_format = "#,##0"
        row += 1

    # Key takeaway
    ws.cell(row, 1).value = ("KEY TAKEAWAY for labor_drag: Seasonal workers are 57% of non-family headcount and "
                             "12.8% of total workdays. The seasonal layer is the structurally vulnerable part of the "
                             "labor pool (migrant + casual day-labor). Family labor accounts for 78.6% of total workdays but "
                             "is itself ageing (Tzouramani 2020: avg farm manager age 51.7 yrs). "
                             "This anchors v7 central labor_drag at 7%, range 3-12% in Monte Carlo, "
                             "replacing the 2009 Labrianidis & Sykas anchor.")
    ws.cell(row, 1).font = Font(name="Arial", size=10, bold=True, color="1F4E79")
    ws.cell(row, 1).fill = FILL_INFO
    ws.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row, 1).border = STYLE["BORDER"]
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.row_dimensions[row].height = 70
    row += 2

    # ----- B. Holdings and UAA -----
    row = section_header(row, "B. Holdings and Utilised Agricultural Area, Greece 2020 (ELSTAT Census, Table B01)")
    row = col_headers(row, ["Indicator", "Value", "Unit", "Source", "Notes", "", "", ""])
    hu = ELSTAT_2020["holdings_uaa"]
    hold_rows = [
        ("Total holdings", hu["total_holdings"], "count", "ELSTAT B01",
         "Includes holdings with and without UAA"),
        ("Holdings with UAA", hu["holdings_with_uaa"], "count", "ELSTAT B01",
         "Used as denominator for cap_weight derivation"),
        ("Utilised agricultural area (UAA)", hu["uaa_hectares"], "hectares", "ELSTAT B01",
         "28,228.86 thousand stremmas converted at 1 stremma = 0.1 ha"),
        ("UAA cross-check (FADN 2020)", 2_810_000, "hectares", "FADN 2020",
         "Independent measurement confirms ELSTAT to within 0.5%"),
    ]
    for i, r_data in enumerate(hold_rows):
        for c_idx, v in enumerate(r_data, 1):
            cell = ws.cell(row, c_idx)
            cell.value = v
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = STYLE["BORDER"]
            if isinstance(v, (int, float)) and not isinstance(v, bool):
                cell.number_format = "#,##0"
        row += 1
    row += 1

    # ----- C. UAA by crop -----
    row = section_header(row, "C. UAA by Crop Category, Greece 2020 (ELSTAT Census, Tables 7, 9, 10, 17, 22, 27)")
    row = col_headers(row, ["Crop category", "Area (ha)", "Share of total UAA", "Source", "Maps to HS chapter", "", "", ""])
    # Use the official total UAA from Table B01 (2,822,886 ha) as the denominator.
    # Table 7 reports a "holdings with UAA" land-use breakdown summing to ~2,714,266 ha,
    # the difference (~108,620 ha) is fallow/uncultivated land counted in B01 but not in T7's
    # crop-category breakdown. Using B01 makes shares comparable to the headline UAA figure
    # shown in Section B.
    total_uaa = ELSTAT_2020["holdings_uaa"]["uaa_hectares"]
    uaa = ELSTAT_2020["uaa_by_crop_ha"]
    crop_rows = [
        ("Arable crops (total)", uaa["arable_total"], "Table 7", "23 (animal feed via cereals)"),
        ("  Cereals (total)",    uaa["cereals_total"], "Table 9", "23 (animal feed input)"),
        ("    Durum wheat",      uaa["wheat_durum"], "Table 10", "Not in chapter set"),
        ("    Soft wheat",       uaa["wheat_soft"], "Table 10", "Not in chapter set"),
        ("    Barley",           uaa["barley"], "Table 10", "23 (animal feed)"),
        ("    Maize",            uaa["maize"], "Table 10", "23 (animal feed)"),
        ("Permanent tree crops (total)", uaa["tree_crops_total"], "Table 7", "08, 15, 20"),
        ("  Olives total",       uaa["olives_total"], "Table 17", "15 + 20"),
        ("    Olives for oil",   uaa["olives_for_oil"], "Table 17", "15 (Animal/vegetable fats and oils, GI Kalamata)"),
        ("    Olives for table", uaa["olives_table"], "Table 17", "20 (Prepared vegetables/fruits/olives)"),
        ("  Temperate fruit trees",  uaa["fruit_trees_temperate"], "Table 17 c11-12", "08 (Fruit and nuts, temperate: apples, pears, peaches)"),
        ("  Tropical/subtropical fruit", uaa["fruit_trees_tropical"], "Table 17 c13-14", "08 (Fruit and nuts, subtropical)"),
        ("  Nuts (almonds etc.)", uaa["nuts"], "Table 17 c15-16", "08 (Fruit and nuts, nuts category)"),
        ("  Citrus",             uaa["citrus"], "Table 17 c17-18", "08 (Fruit and nuts, citrus: oranges etc.)"),
        ("  Other perennials",   uaa["tree_crops_other_perennials"], "Table 17 c19-20", "08 (residual)"),
        ("Vineyards (total)",    uaa["vineyards_total"], "Table 22", "22 + 08"),
        ("  PDO wine",           uaa["vineyards_pdo"], "Table 22", "22 (GI wines)"),
        ("  PGI wine",           uaa["vineyards_pgi"], "Table 22", "22 (GI wines)"),
        ("  Other wine",         uaa["vineyards_other_wine"], "Table 22", "22 (non-GI)"),
        ("  Table grapes",       uaa["vineyards_table_grapes"], "Table 22", "08"),
        ("  Raisins/sultanas",   uaa["vineyards_raisins"], "Table 22", "08 (GI Korinthiaki/Sultana)"),
        ("Irrigated area",       uaa["irrigated_total"], "Table 27", "Cross-cuts above; 32.7% of UAA"),
    ]
    for i, r_data in enumerate(crop_rows):
        name, area, src, hs = r_data
        share = area / total_uaa
        for c_idx, v in enumerate([name, area, share, src, hs], 1):
            cell = ws.cell(row, c_idx)
            cell.value = v
            cell.font = Font(name="Arial", size=10, italic=name.startswith("  "))
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = STYLE["BORDER"]
            if c_idx == 2:
                cell.number_format = "#,##0"
            if c_idx == 3:
                cell.number_format = "0.0%"
        row += 1
    row += 1

    # ----- D. Livestock -----
    row = section_header(row, "D. Livestock, Greece 2020 (ELSTAT Census, Tables E05, 35A)")
    row = col_headers(row, ["Species", "Holdings", "Number of animals", "% of livestock holdings",
                            "", "", "Maps to HS chapter", ""])
    ls = ELSTAT_2020["livestock"]
    ls_total_holdings = sum(v["holdings"] for v in ls.values())
    ls_rows = [
        ("Bovine animals", ls["bovine"]["holdings"], ls["bovine"]["animals"], "04 (dairy); 02 not in chapter set"),
        ("Sheep",          ls["sheep"]["holdings"],  ls["sheep"]["animals"],  "04 (sheep milk for Feta PDO)"),
        ("Goats",          ls["goats"]["holdings"],  ls["goats"]["animals"],  "04 (goat milk; Feta is mixed sheep/goat by PDO spec)"),
        ("Pigs",           ls["pigs"]["holdings"],   ls["pigs"]["animals"],   "02 (not in chapter set)"),
    ]
    for r_data in ls_rows:
        name, holdings, animals, hs = r_data
        pct = holdings / ls_total_holdings if ls_total_holdings else 0
        vals = [name, holdings, animals, pct, "", "", hs, ""]
        for c_idx, v in enumerate(vals, 1):
            cell = ws.cell(row, c_idx)
            cell.value = v
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = STYLE["BORDER"]
            if c_idx in (2, 3) and isinstance(v, (int, float)):
                cell.number_format = "#,##0"
            if c_idx == 4:
                cell.number_format = "0.0%"
        row += 1

    ws.cell(row, 1).value = ("Greek livestock is overwhelmingly small ruminants. Sheep + Goat = 10.87M animals across "
                             "93,739 holdings, vs only 624k bovines. This concentration is the empirical basis for "
                             "cap_weight = 1.00 on HS 04 (dairy), since Feta PDO requires Greek sheep/goat milk by regulation.")
    ws.cell(row, 1).font = Font(name="Arial", size=9, italic=True, color="595959")
    ws.cell(row, 1).fill = FILL_INFO
    ws.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row, 1).border = STYLE["BORDER"]
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.row_dimensions[row].height = 48
    row += 2

    # ----- E. FADN per-farm time series -----
    row = section_header(row, "E. FADN Farm Economy Focus, Greece 2019-2023 (Per representative farm, EUR)")
    row = col_headers(row, ["Year", "Total output", "Intermediate cons.", "Balance current subs.",
                            "Gross Farm Income", "Farm Net Value Added", "Total subs. (ex. invest.)",
                            "Total direct payments"])
    for r_data in FADN_2019_2023["per_farm_eur"]:
        for c_idx, v in enumerate(r_data, 1):
            cell = ws.cell(row, c_idx)
            cell.value = v
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.border = STYLE["BORDER"]
            if c_idx == 1:
                cell.number_format = "0"
            else:
                cell.number_format = "#,##0"
        row += 1

    ws.cell(row, 1).value = ("TREND: Total output per farm rose 40.6% from 2019 to 2023 (EUR 23.4k -> EUR 32.9k), "
                             "reflecting price inflation more than volume. Direct payments per farm declined 20.8% "
                             "(EUR 5,522 -> EUR 4,376) under the 2023-2027 CAP reform. Farm Net Value Added grew 26.9%. "
                             "FADN measures Greek-only quantities and validates the Greek agri scale but cannot replace "
                             "the EU-vs-Mercosur wedge sources.")
    ws.cell(row, 1).font = Font(name="Arial", size=9, italic=True, color="595959")
    ws.cell(row, 1).fill = FILL_INFO
    ws.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row, 1).border = STYLE["BORDER"]
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.row_dimensions[row].height = 56
    row += 2

    # ----- F. FADN scale -----
    row = section_header(row, "F. FADN Greek Agriculture Scale, 2019-2023 (Population-level totals)")
    row = col_headers(row, ["Year", "Farms represented", "Population farms", "UAA (ha)", "Livestock (LSU)",
                            "Labour force (AWU)", "Std Output (EUR bn)", ""])
    for r_data in FADN_2019_2023["scale"]:
        for c_idx, v in enumerate(r_data, 1):
            cell = ws.cell(row, c_idx)
            cell.value = v
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.border = STYLE["BORDER"]
            if c_idx == 1:
                cell.number_format = "0"
            elif c_idx == 7:
                cell.number_format = "0.00"
            else:
                cell.number_format = "#,##0"
        row += 1
    row += 2

    # ----- G. Updated labor_drag parameters -----
    row = section_header(row, "G. Updated labor_drag Parameter (v7 anchoring)")
    row = col_headers(row, ["Parameter", "v6 value", "v7 value", "Rationale", "Source", "", "", ""])
    ld_rows = [
        ("labor_drag central case", "5.0%", "7.0%",
         "Census shows seasonal workers are 57% of non-family headcount and 12.8% of total workdays, much higher than the 20% migrant share from 2009. Raises central from 5% to 7%.",
         "ELSTAT 2020 Census, Table 43"),
        ("labor_drag MC low", "0.0%", "3.0%",
         "Lower bound non-zero because the structural seasonal-labor share is empirically observed, not hypothetical.",
         "ELSTAT 2020 Census, Table 43"),
        ("labor_drag MC mode", "5.0%", "7.0%",
         "Same as central.", "ELSTAT 2020 Census, Table 43"),
        ("labor_drag MC high", "10.0%", "12.0%",
         "Upper bound widened to reflect compounding effect of farm-manager aging (Tzouramani 2020: 51.7 yrs) on top of seasonal-labor vulnerability.",
         "Tzouramani 2020 + ELSTAT 2020 Census"),
    ]
    for r_data in ld_rows:
        for c_idx, v in enumerate(r_data, 1):
            cell = ws.cell(row, c_idx)
            cell.value = v
            cell.font = Font(name="Arial", size=10, bold=(c_idx == 3))
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = STYLE["BORDER"]
            if c_idx == 3:
                cell.fill = FILL_HIGHLIGHT
        row += 1
    row += 1

    # ----- H. OECD Environmental Tax indicator -----
    row = section_header(row, "H. Environmental Tax Burden, Greece vs OECD vs LAC region (OECD Indicator 2022; OECD Rev Stats LAC 2021)")
    row = col_headers(row, ["Country / region", "Env tax (% of GDP)", "Source", "Year", "Implication for model", "", "", ""])
    et = OECD_ENV_TAX
    et_rows = [
        ("Greece",                 et["greece_pct_gdp_2022"],          "OECD dashboard", "2022", "Highest env-tax burden in the OECD; rising since 2013 per EEA"),
        ("Greece (cross-check)",   et["greece_pct_gdp_2020_via_wkp"],   "OECD WKP-2023-23",  "2020", "Independent literature value, confirms ~3.6-3.8% range"),
        ("OECD full average",      et["oecd_full_avg_pct_gdp_recent"],  "OECD/CIAT/ECLAC/IDB Rev Stats LAC 2021", "2019",
         "True OECD-wide average; the dashboard 'OECD' bar at ~1.2% reflected only the SELECTED filter countries, not all members"),
        ("EU27 average",           et["eu27_avg_pct_gdp_2020"],         "IMF/OECD 2022",  "2020", "EU27 average (Greece is well above)"),
        ("Ireland (EU low)",       et["ireland_pct_gdp_2020_via_wkp"],  "OECD WKP-2023-23", "2020", "Lowest EU value for reference (1.2%)"),
        ("LAC regional average",   et["lac_regional_avg_pct_gdp_2019"], "OECD/CIAT/ECLAC/IDB Rev Stats LAC 2021", "2019",
         "Best available proxy for Mercosur (which are not OECD members)"),
        ("GREECE vs LAC gap",      et["implied_greece_vs_lac_gap_pp_gdp"], "Calculated (3.6% - 1.2%)", "—",
         f"Implied gap = {et['implied_greece_vs_lac_gap_pp_gdp']*100:.1f} pp; drives wedge addon"),
        ("GREECE vs OECD gap",     et["implied_greece_vs_oecd_gap_pp_gdp"], "Calculated (3.6% - 2.1%)", "—",
         f"Reference: smaller {et['implied_greece_vs_oecd_gap_pp_gdp']*100:.1f} pp gap for context"),
    ]
    for i, r_data in enumerate(et_rows):
        is_total = "GREECE vs LAC" in str(r_data[0])
        for c_idx, v in enumerate(r_data, 1):
            cell = ws.cell(row, c_idx)
            cell.value = v
            cell.font = STYLE["FONT_TOTAL"] if is_total else Font(name="Arial", size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = STYLE["BORDER"]
            if is_total:
                cell.fill = FILL_HIGHLIGHT
            if c_idx == 2 and isinstance(v, (int, float)):
                cell.number_format = "0.0%"
        row += 1

    # Caveat note about country-specific Mercosur data and dashboard "OECD" interpretation
    ws.cell(row, 1).value = ("CAVEATS: (1) The OECD dashboard PDF shows a highlighted 'OECD' bar at ~1.2%, but per "
                             "OECD/IMF literature this represents the average of the SELECTED-FILTER countries on that "
                             "dashboard view, not the full OECD-wide average. The true OECD-wide average is approximately "
                             "2.1% per OECD/CIAT 2021 Revenue Statistics LAC. EU27 average is 2.2% per IMF/OECD 2022. "
                             "(2) Mercosur countries (Brazil, Argentina, Uruguay, Paraguay) are not OECD members and are "
                             "not on the OECD dashboard. Country-specific Mercosur figures were not located in primary "
                             "sources during this build. The 1.2% LAC regional average from OECD/CIAT/ECLAC/IDB Revenue "
                             "Statistics LAC 2021 is used as the most defensible Mercosur proxy. Brazil specifically is "
                             "reported at less than 2% of total tax revenue (roughly 0.6% of GDP) per the 2015 OECD EPR.")
    ws.cell(row, 1).font = Font(name="Arial", size=9, italic=True, color="595959")
    ws.cell(row, 1).fill = FILL_INFO
    ws.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row, 1).border = STYLE["BORDER"]
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.row_dimensions[row].height = 90
    row += 1

    ws.cell(row, 1).value = ("KEY TAKEAWAY for cost wedge: Greek producers carry an environmental-tax burden 2.4 percentage "
                             "points of GDP higher than the LAC regional average (3.6% vs 1.2%). For HS 23 (animal feed, the "
                             "energy-intensive Greek agri processing chapter where a tariff channel exists at eu_mfn=5%), "
                             "this translates to approximately 0.5 pp of additional cost wedge after partial pass-through. "
                             "HS 12 was considered but omitted: eu_mfn=0 means the wedge multiplies tariff_change=0 and has "
                             "no calculated effect. EEA confirms this gap is WIDENING: Greek env-tax share has risen since "
                             "2013 while the EU average has slightly declined. The model adds this as a chapter wedge "
                             "component in v7, applied at config-load time to chap['wedge'] for HS 23 only.")
    ws.cell(row, 1).font = Font(name="Arial", size=10, bold=True, color="1F4E79")
    ws.cell(row, 1).fill = FILL_INFO
    ws.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row, 1).border = STYLE["BORDER"]
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.row_dimensions[row].height = 70
    row += 2

    # ----- I. EEA Greece Country Profile 2025 indicators -----
    row = section_header(row, "I. Greek Environmental and Agri Indicators (EEA Europe's Environment 2025)")
    row = col_headers(row, ["Indicator", "Value", "Unit", "Status / EU comparison", "Relevance for model", "", "", ""])
    eea = EEA_GREECE_2025["key_indicators"]
    eea_rows = [
        ("Area under organic farming",      eea["area_under_organic_farming_pct"], "% of UAA",  "Acceleration needed for 2030 target",
         "Modulates F2F_drag exposure; organic farms already compliant"),
        ("Designated terrestrial protected areas", eea["designated_terrestrial_protected_pct"], "% of land", "Achieved (above 2030 target)",
         "Constrains some land-use expansion; minor for trade model"),
        ("Climate-related economic losses", eea["climate_economic_losses_eur_cap"], "EUR per capita", "Wrong direction",
         "Supports asymmetric-risk framing for wedge upper bound"),
        ("Fossil fuel subsidies (Greece)",  eea["fossil_fuel_subsidies_pct_gdp_2023"], "% of GDP",  "Above EU avg of 0.7%",
         "Higher input-cost subsidies partially offset Greek env-tax burden"),
        ("Renewable share of final energy", eea["renewable_share_final_energy_pct"], "%",         "Acceleration needed for 2030",
         "Energy transition cost burden flows into agri processing prices"),
        ("Final energy consumption",        eea["final_energy_consumption_idx_2005"], "Index 2005=100", "On track for 2030 target",
         "Reduced energy intensity supports lower wedge for energy-light chapters"),
        ("GHG emissions (no LULUCF)",       eea["ghg_emissions_idx_1990_no_lulucf"], "Index 1990=100", "Not assessed in EEA card",
         "30.8% reduction vs 1990 baseline; substantive decoupling"),
        ("Circular material use rate",      eea["circular_material_use_rate_pct"], "%",          "Acceleration needed",
         "Below EU avg; affects materials cost trajectory"),
        ("Marine protected areas",          eea["marine_protected_areas_pct"], "% of marine",     "Target 30%",
         "Affects fisheries (HS 03 not in chapter set)"),
        ("Surface water in good ecological condition", eea["good_ecological_surface_water_pct"], "%", "Improving",
         "Limits agri water-pollution liabilities"),
    ]
    for i, r_data in enumerate(eea_rows):
        for c_idx, v in enumerate(r_data, 1):
            cell = ws.cell(row, c_idx)
            cell.value = v
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = STYLE["BORDER"]
            if c_idx == 2 and isinstance(v, (int, float)):
                if v <= 1:
                    cell.number_format = "0.0%"
                else:
                    cell.number_format = "#,##0.0"
        row += 1

    # Greek agri-specific findings
    ws.cell(row, 1).value = ("GREEK AGRI STRUCTURAL FINDINGS (EEA 2025 narrative): Primary-sector average value yield is "
                             "EUR 1,100 per hectare (EUR 110 per 1,000 m²); small farms predominate; segmented geography "
                             "limits scale; innovation deployment to farmers is very limited; knowledge transfer happens "
                             "mainly through company-driven contract farming. Olive oil consumption is declining domestically "
                             "and being replaced by cheaper seed oils. Plant-based diets emerging but still niche. "
                             "CAP renegotiation took place early 2025.")
    ws.cell(row, 1).font = Font(name="Arial", size=9, italic=True, color="595959")
    ws.cell(row, 1).fill = FILL_INFO
    ws.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row, 1).border = STYLE["BORDER"]
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.row_dimensions[row].height = 70
    row += 1

    ws.cell(row, 1).value = ("ECO-INNOVATION SIGNAL: Greece's Eco-Innovation Index score is below the EU average but its "
                             "rate of increase 2013-2022 is the HIGHEST in the entire EU. Combined with rising environmental "
                             "tax share, falling fossil-fuel subsidies trajectory, and new laws on green investment, "
                             "this is a structural tailwind. The model treats this as a downward bias on the wedge over the "
                             "10-year horizon, but does not currently parametrize it.")
    ws.cell(row, 1).font = Font(name="Arial", size=9, italic=True, color="595959")
    ws.cell(row, 1).fill = FILL_INFO
    ws.cell(row, 1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row, 1).border = STYLE["BORDER"]
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws.row_dimensions[row].height = 60

    # Column widths
    widths = [42, 14, 16, 18, 12, 16, 36, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"


def write_sheet15_discussion(wb, s1_results, sens_grid_s1, pe_mc, cap_weights, f2f_drag, labor_drag):
    """Sheet 15 (v7): Discussion. Effects, contributions, current situation, future directions.

    This sheet synthesizes the quantitative model output with the Greek-specific
    structural data on Sheet 14. It is the place where the model 'speaks' about
    Greece, not just computes numbers.
    """
    ws = wb.create_sheet("15-Discussion")
    ws["A1"] = "Discussion: Effects, Contributions, Current Situation, Future Directions"
    style_cell(ws["A1"], font=STYLE["FONT_TITLE"], border=False)
    ws.merge_cells("A1:F1")
    ws.row_dimensions[1].height = 28

    ws["A2"] = ("This sheet synthesizes the model output (Sheets 4-13) with Greek structural data (Sheet 14). "
                "It is the analytical bridge between numerical projections and policy interpretation, "
                "addressed to readers who want to understand WHY the numbers look the way they do for Greece specifically.")
    ws["A2"].font = STYLE["FONT_NOTE"]
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:F2")
    ws.row_dimensions[2].height = 40

    FILL_HIGHLIGHT = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    FILL_INFO = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
    FILL_WARN = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    row = 4

    def section(r, title):
        c = ws.cell(r, 1)
        c.value = title
        style_cell(c, font=STYLE["FONT_SUBTITLE"], fill=STYLE["FILL_SECTION"])
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 22
        return r + 1

    def para(r, text, fill=None, height=None):
        c = ws.cell(r, 1)
        c.value = text
        c.font = Font(name="Arial", size=10)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        c.border = STYLE["BORDER"]
        if fill: c.fill = fill
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=6)
        if height: ws.row_dimensions[r].height = height
        else: ws.row_dimensions[r].height = max(15, 14 * (len(text) // 110 + 2))
        return r + 1

    # Compute key numbers from inputs
    central = sens_grid_s1.get((1.0, 0.10), {})
    central_widening = central.get("bilateral_widening", 0) / 1e6 if central else 0
    pe_y10 = pe_mc.get("y10", {}).get("widening", {})
    pe_mean = pe_y10.get("mean", 0) / 1e6
    pe_lo = pe_y10.get("lo", 0) / 1e6
    pe_hi = pe_y10.get("hi", 0) / 1e6

    # Find Y10 imports and exports
    y10_imp = sum(c["years"][10]["proj_imp"] - c["years"][10]["cf_imp"] for c in s1_results) / 1e6
    y10_exp_unadj = sum(c["years"][10]["proj_exp"] - c["years"][10]["cf_exp"] for c in s1_results) / 1e6
    # Compute Tier 3 adjusted exports from the adj_exports_per_chapter helper
    _adj_block = adj_exports_per_chapter(s1_results, 10, f2f_drag, labor_drag, cap_weights)
    y10_exp_adj = sum(c["adj_export"] for c in _adj_block["chapters"]) / 1e6 \
                  - sum(c["years"][10]["cf_exp"] for c in s1_results) / 1e6
    # Note: adj_export in the helper is the absolute adjusted level, so subtract counterfactual to get the increment
    # Actually, easier: total Greek capacity loss = sum of c['loss']
    y10_capacity_loss = sum(c["loss"] for c in _adj_block["chapters"]) / 1e6

    # ----- 1. Headline result -----
    row = section(row, "1. Headline result")
    row = para(row, (f"Under Scenario 1 (agreement enters force 2026), the model projects bilateral widening of "
                     f"approximately ${central_widening:,.0f}M at Year 10 (Tier 3 deterministic central case with "
                     f"per-chapter cap_weights). The PE Armington Monte Carlo yields a mean of ${pe_mean:,.0f}M with a "
                     f"95% confidence interval of [${pe_lo:,.0f}M, ${pe_hi:,.0f}M]. Additional imports at Y10: "
                     f"${y10_imp:,.0f}M. Additional exports at Y10: ${y10_exp_unadj:,.0f}M unadjusted, reduced to "
                     f"${y10_exp_unadj - y10_capacity_loss:,.0f}M after applying Greek per-chapter capacity adjustment "
                     f"(F2F drag + labor drag, weighted by cap_weight); Greek capacity loss is ${y10_capacity_loss:,.0f}M. "
                     f"The widening is driven almost entirely by the import side; Greek exports rise modestly but "
                     f"are constrained by structural capacity factors."),
               fill=FILL_HIGHLIGHT, height=100)

    # ----- 2. Effects per HS chapter -----
    row = section(row, "2. Effects per HS chapter (Y10, Tier 3 with per-chapter Greek capacity adjustment)")

    # Compute per-chapter Y10 widening contribution (Tier 3, using adjusted exports)
    # _adj_block has 'chapters' list with unadj_export, adj_export, cap_weight, loss per chapter
    adj_by_hs = {c["hs"]: c for c in _adj_block["chapters"]}
    chapter_effects = []
    for r_data in s1_results:
        hs = r_data["hs"]
        name = r_data["name"]
        d_imp = (r_data["years"][10]["proj_imp"] - r_data["years"][10]["cf_imp"]) / 1e6
        d_exp_unadj = (r_data["years"][10]["proj_exp"] - r_data["years"][10]["cf_exp"]) / 1e6
        # Tier 3: adjusted Delta-exports = unadj - chapter capacity loss
        chap_loss = adj_by_hs.get(hs, {}).get("loss", 0) / 1e6
        d_exp_adj = d_exp_unadj - chap_loss
        cw = cap_weights.get(hs, 0.0) if cap_weights else 0.0
        widening_t3 = d_imp - d_exp_adj
        chapter_effects.append((hs, name, d_imp, d_exp_adj, widening_t3, cw))
    chapter_effects.sort(key=lambda x: -x[4])  # largest widening first

    # Header
    for c_idx, h in enumerate(["HS", "Chapter", "Δ Imports Y10 ($M)", "Δ Exports (T3 adj, $M)",
                                "Net widening T3 ($M)", "cap_weight"], 1):
        cell = ws.cell(row, c_idx)
        cell.value = h
        style_cell(cell, font=STYLE["FONT_HEADER"], fill=STYLE["FILL_HEADER"],
                   alignment=STYLE["ALIGN_CENTER"])
    ws.row_dimensions[row].height = 28
    row += 1

    for hs, name, d_imp, d_exp, net, cw in chapter_effects[:10]:  # top 10
        for c_idx, v in enumerate([hs, name, d_imp, d_exp, net, cw], 1):
            cell = ws.cell(row, c_idx)
            cell.value = v
            cell.font = Font(name="Arial", size=10)
            cell.alignment = Alignment(horizontal="right" if c_idx >= 3 else "left", vertical="top",
                                       wrap_text=True)
            cell.border = STYLE["BORDER"]
            if c_idx in (3, 4, 5):
                cell.number_format = '$#,##0.0"M";($#,##0.0"M");-'
            elif c_idx == 6:
                cell.number_format = "0%"
        row += 1

    # Sum top-10 vs rest
    top10_d_imp = sum(c[2] for c in chapter_effects[:10])
    top10_d_exp = sum(c[3] for c in chapter_effects[:10])
    top10_net = sum(c[4] for c in chapter_effects[:10])
    rest_d_imp = sum(c[2] for c in chapter_effects[10:])
    rest_d_exp = sum(c[3] for c in chapter_effects[10:])
    rest_net = sum(c[4] for c in chapter_effects[10:])
    total_d_imp = top10_d_imp + rest_d_imp
    total_d_exp = top10_d_exp + rest_d_exp
    total_net = top10_net + rest_net

    # Rest-of-chapters row (Greek exports rise more than imports in some chapters, producing negative widening)
    for c_idx, v in enumerate(["", "Other 15 chapters (net)", rest_d_imp, rest_d_exp, rest_net, ""], 1):
        cell = ws.cell(row, c_idx)
        cell.value = v
        cell.font = Font(name="Arial", size=10, italic=True, color="595959")
        cell.alignment = Alignment(horizontal="right" if c_idx >= 3 else "left", vertical="top", wrap_text=True)
        cell.border = STYLE["BORDER"]
        if c_idx in (3, 4, 5):
            cell.number_format = '$#,##0.0"M";($#,##0.0"M");-'
    row += 1

    # TOTAL row
    for c_idx, v in enumerate(["", "TOTAL (all 25 chapters)", total_d_imp, total_d_exp, total_net, ""], 1):
        cell = ws.cell(row, c_idx)
        cell.value = v
        cell.font = Font(name="Arial", size=10, bold=True)
        cell.alignment = Alignment(horizontal="right" if c_idx >= 3 else "left", vertical="top", wrap_text=True)
        cell.border = STYLE["BORDER"]
        cell.fill = FILL_HIGHLIGHT
        if c_idx in (3, 4, 5):
            cell.number_format = '$#,##0.0"M";($#,##0.0"M");-'
    row += 1
    row += 1

    # ----- 3. Contributions: which Greek sectors are most exposed and why -----
    row = section(row, "3. Contributions: Why these chapters dominate the result")
    row = para(row, ("The widening concentrates in animal feed (HS 23), coffee/tea/spices (HS 09), and to a lesser extent "
                     "wine (HS 22) and prepared olives/vegetables (HS 20). "
                     "The reasons are visible directly in Greek structural data on Sheet 14:"))

    row = para(row, ("HS 23 (Animal feed): Greece imports large volumes of soy and feed grains from Mercosur. "
                     "Greek cereals occupy 20.8% of UAA but are not currently price-competitive with imported feed. "
                     "Tariff elimination amplifies the existing structural deficit. Greek production cannot quickly fill "
                     "the gap because the 585,989 ha of cereal area is largely committed to wheat (durum 231k ha, soft 89k ha) "
                     "rather than feed grains. Additionally, Greek producers face an environmental-tax burden 2.4 pp of GDP "
                     "higher than the LAC regional average (OECD 2022: Greece 3.6%, LAC avg 1.2%, used as Mercosur proxy), "
                     "passing through to feed processing costs."),
              fill=FILL_INFO)

    row = para(row, ("HS 22 (Wine and beverages): Greek vineyards cover 58,754 ha total, of which 15,380 ha are "
                     "PDO+PGI protected (26% of vineyard area, 44% of wine-only vineyards excluding table grapes "
                     "and raisins). The remaining 74% is non-GI wine, table grapes, and "
                     "raisins, which compete on price with Mercosur exporters. The cap_weight of 0.90 reflects this: "
                     "GI wines are protected, but the majority of Greek vineyard area is not."), fill=FILL_INFO)

    row = para(row, ("HS 04 (Dairy, including Feta PDO): cap_weight 1.00 protects Feta because the PDO regulation "
                     "requires Greek sheep/goat milk. The 10.87 million sheep+goats across 93,739 holdings (ELSTAT 2020) "
                     "give Greece exclusive supply for Feta. The TRQ cap of $8.64M Greek share on dairy is binding in the "
                     "model and limits import substitution from Mercosur for this category."), fill=FILL_INFO)

    row = para(row, ("HS 15 (Olive oil, GI Kalamata): cap_weight 1.00 anchored on 522,069 ha of olive oil plantations "
                     "(18.5% of UAA, the largest tree-crop contributor). Greek olive oil is exported to high-value EU markets "
                     "and does not face direct Mercosur competition in those segments. The cap_weight protects this "
                     "structural advantage in the model."), fill=FILL_INFO)
    row += 1

    # ----- 4. Current situation -----
    row = section(row, "4. Current situation (2020-2024 baseline)")
    row = para(row, ("Greek agriculture in the 2020-2024 baseline period shows three structural patterns visible in the data:"))

    row = para(row, ("Pattern A: Labor pyramid. ELSTAT 2020 Census shows 78.6% of agricultural workdays are family labor, "
                     "12.8% seasonal paid, 5.9% permanent paid. Average farm manager age is 51.7 years (Tzouramani 2020). "
                     "The seasonal layer is the immediate structural risk: 794,811 persons across 215,948 holdings depend "
                     "on labor supply that has been declining since the 2009 financial crisis and again post-COVID."))

    row = para(row, ("Pattern B: Land use concentration. Of 2.82 million hectares of UAA, the arable category covers "
                     "46.4% (cereals alone 20.8% of UAA, the largest single crop category) and permanent tree crops "
                     "cover 25.9% (dominated by olives at 20.7% of UAA). Greek competitive advantages are concentrated "
                     "in tree crops and small-ruminant livestock, both of which are slower to adapt to demand shifts "
                     "than arable rotations."))

    row = para(row, ("Pattern C: Subsidy dependence and reform. FADN 2019-2023 shows direct payments per farm declined "
                     "from EUR 5,522 to EUR 4,376 (a 20.8% drop) under the 2023-2027 CAP reform, while total output rose "
                     "from EUR 23,425 to EUR 32,929 per farm (+40.6%, mostly inflation). The balance of current subsidies "
                     "fell from EUR 6,749 to EUR 4,957 (-26.6%), shifting more weight onto market revenues."))

    row += 1

    # ----- 5. Future directions / policy implications -----
    row = section(row, "5. Future directions and policy implications")

    row = para(row, ("Future direction 1: Protect what already works. The model treats GI-protected products (Feta, Kalamata "
                     "olive oil, PDO/PGI wines) with cap_weight 1.00 because the regulatory ties to Greek production are "
                     "strong. The data shows these categories cover meaningful land area (olives for oil 522k ha, GI wines "
                     "15k ha) and are unlikely to be displaced by Mercosur imports if the GI regulations are enforced. "
                     "Policy priority: maintain and expand GI protection."))

    row = para(row, ("Future direction 2: Address structural feed-grain deficit. HS 23 (animal feed) is the largest single "
                     "contributor to bilateral widening. Greek cereals at 586k ha are predominantly wheat, not feed-grade. "
                     "Policy options include either accepting the import dependence (status quo) or shifting cereal area "
                     "toward feed crops via CAP coupled support adjustments. The model does not endorse either choice; "
                     "it quantifies the cost of the status quo at approximately $2.2B by Y10."))

    row = para(row, ("Future direction 3: Stabilize the labor base. The structural seasonal-labor share (12.8% of workdays, "
                     "57% of non-family headcount) is the immediate risk factor behind the 7% labor_drag. Policy options "
                     "include expanding agricultural visa programs, supporting mechanization investment for small-ruminant "
                     "operations, or expanding the Young Farmers Scheme. The Borda 2023 review found Young Farmers Scheme "
                     "programs in Greece have already stimulated regional output."))

    row = para(row, ("Future direction 4: Use the 2026-2036 phase-in period actively. The model's Y0-Y10 trajectory shows "
                     "the widening accumulates gradually. Greek policy can use this decade to invest in productivity, "
                     "reduce input-cost asymmetry (KEPE 2025/57: Greek input share rose 6.5pp 2019-2023 vs EU 1.0pp), and "
                     "negotiate stronger safeguards in the agreement's TRQ administration. Each percentage point reduction "
                     "in the input-cost wedge translates to approximately $40-50M less widening at Y10 in the sensitivity grid."),
               fill=FILL_HIGHLIGHT)

    row = para(row, ("Future direction 5: Acknowledge and address the environmental-tax wedge. Greek producers carry the "
                     "highest environmental-tax burden in the OECD (3.6% of GDP, 2022) versus a LAC regional average of "
                     "1.2% (most defensible Mercosur proxy, since Mercosur countries are not OECD members and country-"
                     "specific data was not located in primary sources). This gap is structural and widening: EEA 2025 "
                     "confirms Greek env-tax share has risen since 2013 while the EU average has slightly declined. Policy "
                     "options include either border adjustments (CBAM-style) for goods from countries with much lower "
                     "env-tax burdens, or domestic offsets to neutralize the gap for export-exposed agri processing. The "
                     "model adds a 0.5pp env-tax wedge component for HS 23 (animal feed, the energy-intensive Greek agri "
                     "processing chapter where a tariff channel exists). HS 12 (oilseeds) was considered but omitted because "
                     "eu_mfn is already zero, so the wedge has no mathematical effect there. Full pass-through would imply "
                     "0.7-1.2pp depending on energy intensity."),
               fill=FILL_INFO)

    row = para(row, ("Future direction 6: Leverage the eco-innovation tailwind. The EEA notes Greece's Eco-Innovation Index "
                     "score is still below the EU average, but the increase 2013-2022 is the highest in the entire EU. "
                     "Combined with the 17.2% organic-farming share (the highest single point of comparative advantage in "
                     "the European Green Deal context), Greek agri has a structural opportunity to differentiate on "
                     "sustainability rather than compete on price. The model does not quantify this opportunity, but it sits "
                     "in the same time window as the Mercosur agreement phase-in and is a credible policy lever."), fill=FILL_INFO)

    row += 1

    # ----- 6. Caveats -----
    row = section(row, "6. Caveats and what the model does not say")
    row = para(row, ("This model is partial-equilibrium Armington at HS 2-digit, calibrated to Greek bilateral trade with "
                     "Mercosur 2014-2024. It does NOT capture: (a) third-country substitution effects (e.g. EU producers "
                     "outside Greece displacing Greek share), (b) consumer welfare gains from cheaper imports, (c) general-"
                     "equilibrium feedback through wages and exchange rates, (d) climate transition impacts on either side, "
                     "(e) non-tariff measures beyond the EUDR/CBAM compliance wedges already included. The XGBoost ML "
                     "complement (Sheet 11) finds that distance, lagged trade, and GDP collectively explain 93.6% of "
                     "bilateral trade variation while tariffs explain 0.2%, suggesting the PE Armington headline may "
                     "overstate the trade-policy effect. Treat the $3.5B range as an upper bound on the trade-policy "
                     "channel, not a forecast."), fill=FILL_WARN)

    # Column widths
    widths = [14, 38, 16, 16, 16, 14]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"


# ====================================================================
# MAIN ORCHESTRATOR
# ====================================================================

def main(output_dir=None):
    if output_dir is None:
        output_dir = resolve_output_dir()
    print("=" * 80)
    print("EU-MERCOSUR GREEK AGRICULTURE ANALYSIS, ORCHESTRATOR")
    print(f"Started: {datetime.now().isoformat()}")
    print(f"Output directory: {output_dir}")
    print("=" * 80)

    g = CONFIG["globals"]
    audit_log = {"started": datetime.now().isoformat(), "config": CONFIG, "stages": {}}

    # ----- STAGE 1: PE Armington -----
    print("\n[1/6] Running PE Armington...")
    s1_results = run_pe_armington(g["agreement_year_s1"], g["baseline_year"], CONFIG["trq_caps"])
    s2_results = run_pe_armington(g["agreement_year_s2"], g["baseline_year"], CONFIG["trq_caps"])
    pe_imp_y10 = sum(c["years"][10]["proj_imp"] - c["years"][10]["cf_imp"] for c in s1_results)
    pe_exp_y10 = sum(c["years"][10]["proj_exp"] - c["years"][10]["cf_exp"] for c in s1_results)
    print(f"      S1 Y10: Add Imp ${pe_imp_y10/1e6:,.1f}M | Add Exp ${pe_exp_y10/1e6:,.1f}M")
    audit_log["stages"]["pe_armington"] = {
        "s1_y10_add_imp": pe_imp_y10, "s1_y10_add_exp": pe_exp_y10,
    }

    # ----- STAGE 2: Sensitivity Grid -----
    print("\n[2/6] Running 3x3 Sensitivity Grid...")
    sg = CONFIG["sensitivity_grid"]
    cap_weights = CONFIG.get("cap_weights")  # Tier 3
    labor_drag = g.get("labor_drag", 0.0)
    sens_grid_s1 = run_sensitivity_grid(g["agreement_year_s1"], g["baseline_year"],
                                         CONFIG["trq_caps"], sg["f2f_drags"], sg["wedge_multipliers"],
                                         g["greek_agri_share"], cap_weights=cap_weights,
                                         labor_drag=labor_drag)
    sens_grid_s2 = run_sensitivity_grid(g["agreement_year_s2"], g["baseline_year"],
                                         CONFIG["trq_caps"], sg["f2f_drags"], sg["wedge_multipliers"],
                                         g["greek_agri_share"], cap_weights=cap_weights,
                                         labor_drag=labor_drag)
    central = sens_grid_s1[(1.0, 0.10)]
    print(f"      S1 central widening Y10 (Tier 3 per-chapter): ${central['bilateral_widening']/1e6:,.1f}M")
    audit_log["stages"]["sensitivity_grid"] = {
        "s1_central_widening": central["bilateral_widening"],
    }

    # ----- STAGE 2b: PE Armington Monte Carlo -----
    print("\n      Running PE Armington Monte Carlo (literature-anchored perturbations)...")
    pe_mc = run_pe_monte_carlo(g["agreement_year_s1"], g["baseline_year"], CONFIG["trq_caps"],
                                CONFIG["pe_monte_carlo"], g["greek_agri_share"],
                                cap_weights=cap_weights)
    print(f"      PE MC Y10 widening: mean ${pe_mc['y10']['widening']['mean']/1e6:,.1f}M, "
          f"95% CI [${pe_mc['y10']['widening']['lo']/1e6:,.1f}M, "
          f"${pe_mc['y10']['widening']['hi']/1e6:,.1f}M]")
    audit_log["stages"]["pe_monte_carlo"] = {
        "y10_widening_mean": pe_mc["y10"]["widening"]["mean"],
        "y10_widening_lo": pe_mc["y10"]["widening"]["lo"],
        "y10_widening_hi": pe_mc["y10"]["widening"]["hi"],
    }

    # ----- STAGE 3: XGBoost + Monte Carlo -----
    print("\n[3/6] Running XGBoost + Monte Carlo...")
    ml_results = run_xgboost_pipeline(g["trade_data_path"], CONFIG["ml"],
                                       mc_seed=g["mc_seed"], mc_n_draws=g["mc_n_draws"])
    print(f"      RMSE: ${ml_results['model_stats']['rmse']:,.0f}")
    print(f"      Y10 mean ΔImp: ${ml_results['monte_carlo']['y10_mean']/1e6:,.2f}M  "
          f"95% CI [${ml_results['monte_carlo']['y10_lo']/1e6:,.2f}M, "
          f"${ml_results['monte_carlo']['y10_hi']/1e6:,.2f}M]")
    audit_log["stages"]["xgboost_mc"] = ml_results["monte_carlo"]

    # ----- STAGE 4: Excel Writer -----
    print("\n[4/6] Writing Excel workbook...")
    os.makedirs(output_dir, exist_ok=True)
    wb = Workbook()
    wb.remove(wb.active)  # Remove default sheet

    write_sheet1_raw_trade(wb, g["trade_data_path"])
    write_sheet2_tariff_params(wb)
    write_sheet3_baselines(wb)
    write_scenario_sheet(wb, "4-Scenario S1 (2026)", s1_results, g["agreement_year_s1"])
    write_scenario_sheet(wb, "5-Scenario S2 (2027)", s2_results, g["agreement_year_s2"])
    write_sheet6_summary(wb, s1_results, s2_results)
    write_sheet7_trq(wb, CONFIG["trq_caps"], s1_results)
    write_sheet8_sensitivity(wb, sens_grid_s1, sens_grid_s2, g["greek_agri_share"],
                              sg["wedge_multipliers"], sg["f2f_drags"])
    write_sheet9_monte_carlo(wb, ml_results, pe_mc)
    write_sheet10_yearly(wb, s1_results)
    write_sheet11_ml_vs_pe(wb, s1_results, ml_results, g["greek_agri_share"], g["f2f_drag"],
                            cap_weights=cap_weights, labor_drag=labor_drag)
    write_sheet12_literature_anchors(wb)
    write_sheet13_per_chapter_adjustments(wb, s1_results, CONFIG.get("cap_weights", {}),
                                           g["f2f_drag"], labor_drag)
    # v7 additions
    write_sheet14_greek_structural_indicators(wb)
    write_sheet15_discussion(wb, s1_results, sens_grid_s1, pe_mc,
                              CONFIG.get("cap_weights", {}), g["f2f_drag"], labor_drag)

    output_path = os.path.join(output_dir, "EU_Mercosur_Greece_Analysis_v7.xlsx")
    wb.save(output_path)
    print(f"      Saved: {output_path}")

    # ----- STAGE 5: Charts -----
    print("\n[5/6] Generating charts...")
    chart_paths = write_all_charts(s1_results, sens_grid_s1, ml_results, pe_mc,
                                    g["greek_agri_share"], g["f2f_drag"], output_dir)
    for name, path in chart_paths.items():
        print(f"      {name}: {path}")
    audit_log["stages"]["charts"] = chart_paths

    # ----- STAGE 6: Audit Log -----
    print("\n[6/6] Writing audit log...")
    audit_log["completed"] = datetime.now().isoformat()
    audit_log["output_file"] = output_path
    audit_log_path = os.path.join(output_dir, "orchestrator_run.json")
    # Make audit log JSON-serializable
    def _convert(obj):
        if isinstance(obj, dict):
            return {str(k): _convert(v) for k, v in obj.items()}
        if isinstance(obj, list):
            return [_convert(v) for v in obj]
        if isinstance(obj, (np.integer, np.floating)):
            return float(obj)
        if isinstance(obj, tuple):
            return list(_convert(v) for v in obj)
        return obj
    audit_log_clean = _convert(audit_log)
    # Replace tuple keys in sensitivity grid
    if "sensitivity_grid" in audit_log_clean.get("config", {}):
        pass
    with open(audit_log_path, "w") as f:
        json.dump(audit_log_clean, f, indent=2, default=str)
    print(f"      Saved: {audit_log_path}")

    print("\n" + "=" * 80)
    print("ORCHESTRATOR COMPLETE")
    print("=" * 80)
    return output_path, audit_log_path


if __name__ == "__main__":
    main()
