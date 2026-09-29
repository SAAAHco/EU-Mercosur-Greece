"""Partial-equilibrium engine, v8.1 (marginal-price TRQ logic added): the v7.5 engine generalized to line-level tariffs.

Faithful to orchestrator_v7_5.py (Module 1 and Module 2): counterfactual compounding,
price-change tariff term, incidence-signed wedge multiplier, Tier 3 export-capacity
adjustment, sensitivity grid and PE Monte Carlo. The only structural change is that a
chapter may carry `eu_lines` / `mer_lines`; the chapter tariff shock is then the
trade-weighted average of line-level price changes, so MFN-free lines contribute zero.
Chapters without line lists fall back to the v7.5 chapter-level treatment, which lets
the module reproduce the published v7.5 results from v5_config.json.
"""
import json, copy, math
import numpy as np

ENV_TAX_ADDON = 0.005
ENV_TAX_CHAPTERS = {"23"}

CAP_WEIGHTS_V75 = {"23": 0.80, "12": 0.70, "09": 0.40, "26": 0.0, "47": 0.0, "27": 0.0, "24": 1.00,
                   "08": 1.00, "30": 0.0, "20": 1.00, "38": 0.0, "84": 0.0, "39": 0.0, "76": 0.0,
                   "85": 0.0, "88": 0.0, "82": 0.0, "73": 0.0, "99": 0.0, "68": 0.0, "25": 0.0,
                   "89": 0.0, "22": 0.90, "04": 1.00, "15": 1.00}
TRQ_CAPS_V75 = {"04": {"cap_usd": 8_640_000, "applied": True}}
GLOBALS = {"f2f_drag": 0.10, "labor_drag": 0.07, "agreement_year_s1": 2026, "agreement_year_s2": 2027,
           "baseline_year": 2023, "greek_agri_share": 0.183}
GRID = {"wedge_multipliers": [0.5, 1.0, 1.5, 3.0], "f2f_drags": [0.07, 0.10, 0.20]}
MC = {"n_draws": 1000, "seed": 42, "wedge_mult_low": 0.5, "wedge_mult_mode": 1.0, "wedge_mult_high": 3.0,
      "f2f_drag_low": 0.05, "f2f_drag_mode": 0.10, "f2f_drag_high": 0.20,
      "labor_drag_low": 0.03, "labor_drag_mode": 0.07, "labor_drag_high": 0.12}


def prepare_chapters(chapters):
    """Apply the v7 env-tax addon and the v7.2 incidence signs, exactly as the orchestrator does."""
    chs = copy.deepcopy(chapters)
    for c in chs:
        if c["hs"] in ENV_TAX_CHAPTERS:
            c["wedge"] = c["wedge"] + ENV_TAX_ADDON
    for c in chs:
        if c.get("wedge", 0) == 0:
            continue
        addon = ENV_TAX_ADDON if c["hs"] in ENV_TAX_CHAPTERS else 0.0
        exporter = c["wedge"] - addon
        c["wedge"] = addon - exporter
    return chs


# ---------------- tariffs ----------------
def _line_tau(line, y):
    """Line tariff in projection year y. phase < 0: no change; phase 0: target at entry."""
    if "r_path" in line:
        return line["tau0"] * line["r_path"][min(y, len(line["r_path"]) - 1)]
    t0, tgt, ph = line["tau0"], line.get("target", 0.0), line["phase"]
    if ph < 0:
        return t0
    if ph == 0 or y >= ph:
        return tgt
    return t0 - (t0 - tgt) * y / ph


def _chapter_tau_v75(tau0, target, phase, reduction_type, y):
    if reduction_type == "None":
        return tau0
    if y >= phase:
        return target
    return tau0 - (tau0 - target) * y / phase


def _trq_pc(l, y, m_line, eps):
    """Marginal-price TRQ response with one or two quotas filled cheapest first. No response
    once cumulative slices bind; otherwise the in-quota price change of the tier the marginal
    unit falls in, capped so projected imports do not exceed that tier's cumulative slice.
    Year 0 runs from 1 May: compare eight months of imports with the pro-rated slice and
    scale the price change by 8/12."""
    t0 = l["tau0"]
    tiers = [(l["trq"]["r_in"], l["trq"]["slice"])]
    if "slice2" in l["trq"]:
        tiers.append((l["trq"]["r_in2"], l["trq"]["slice2"]))
    tiers.sort(key=lambda t: t[0][y])
    frac = 8 / 12 if y == 0 else 1.0
    m = m_line * frac
    if m <= 0:
        return 0.0
    cum = 0.0
    for r_in, sl in tiers:
        cum += sl[y]
        if m < cum:
            pc_in = (1 + t0 * r_in[y]) / (1 + t0) - 1
            cap = (cum / m - 1) / eps
            return max(pc_in, cap) * frac
    return 0.0


def eu_price_change(c, y, cf_imp=None, eps=-3.5):
    if "eu_lines" in c:
        tot = 0.0
        for l in c["eu_lines"]:
            if "trq" in l:
                pc = _trq_pc(l, y, l["share"] * (cf_imp or 0.0), eps)
            else:
                pc = (1 + _line_tau(l, y)) / (1 + l["tau0"]) - 1
            tot += l["share"] * pc
        return tot
    cur = _chapter_tau_v75(c["eu_mfn"], c["target_eu"], c["phase_eu"], c["reduction_type"], y)
    return (1 + cur) / (1 + c["eu_mfn"]) - 1


def mer_price_change(c, y):
    if "mer_lines" in c:
        return sum(l["share"] * ((1 + _line_tau(l, y)) / (1 + l["tau0"]) - 1) for l in c["mer_lines"])
    if c["phase_mer"] == 0:
        cur = c["mer_mfn"]
    elif y >= c["phase_mer"]:
        cur = 0
    else:
        cur = c["mer_mfn"] - c["mer_mfn"] * y / c["phase_mer"]
    return (1 + cur) / (1 + c["mer_mfn"]) - 1


def eu_tariff_display(c, y):
    if "eu_lines" in c:
        return sum(l["share"] * _line_tau(l, y) for l in c["eu_lines"])
    return _chapter_tau_v75(c["eu_mfn"], c["target_eu"], c["phase_eu"], c["reduction_type"], y)


# ---------------- projections ----------------
def cf_imports(c, y, agr, base):
    return c["baseline_imp"] * (1 + c["cagr_imp"]) ** (agr + y - base)


def cf_exports(c, y, agr, base):
    return c["baseline_exp"] * (1 + c["cagr_exp"]) ** (agr + y - base)


def proj_imports(c, y, agr, base, trq_cap=None, wedge_mult=1.0, imp_elast=None):
    cf = cf_imports(c, y, agr, base)
    e = c["imp_elast"] if imp_elast is None else imp_elast
    tc = eu_price_change(c, y, cf, e)
    wf = 1 + c["wedge"] * wedge_mult * min(1, y / max(c["phase_eu"], 1))
    p = cf * (1 + e * tc * wf)
    if trq_cap is not None:
        p = min(p, trq_cap)
    return p


def proj_exports(c, y, agr, base):
    cf = cf_exports(c, y, agr, base)
    return cf * (1 + c["exp_elast"] * mer_price_change(c, y))


def run_pe(chs, agr, base, trq_caps, wedge_mult=1.0, imp_elast=None):
    res = []
    for c in chs:
        cap = trq_caps[c["hs"]]["cap_usd"] if c["hs"] in trq_caps and trq_caps[c["hs"]]["applied"] else None
        yrs = {}
        for y in range(11):
            yrs[y] = {"cf_imp": cf_imports(c, y, agr, base), "cf_exp": cf_exports(c, y, agr, base),
                      "proj_imp": proj_imports(c, y, agr, base, cap, wedge_mult, imp_elast),
                      "proj_exp": proj_exports(c, y, agr, base), "eu_tariff": eu_tariff_display(c, y)}
        res.append({"hs": c["hs"], "name": c["name"], "years": yrs})
    return res


def adj_exports(results, yi, f2f, labor, capw):
    out = {"chapters": [], "total_unadj": 0.0, "total_adj": 0.0}
    for r in results:
        ex = r["years"][yi]["proj_exp"] - r["years"][yi]["cf_exp"]
        w = capw.get(r["hs"], 0.0)
        lf = labor * (yi / 10.0) if yi > 0 else 0.0
        adj = ex * (1 - w * (f2f + lf))
        out["chapters"].append({"hs": r["hs"], "unadj": ex, "adj": adj, "loss": ex - adj})
        out["total_unadj"] += ex
        out["total_adj"] += adj
    out["total_loss"] = out["total_unadj"] - out["total_adj"]
    return out


def headline(chs, capw, trq_caps, agr=2026, base=2023, wedge_mult=1.0, f2f=0.10, labor=0.07, yi=10, imp_elast=None):
    r = run_pe(chs, agr, base, trq_caps, wedge_mult, imp_elast)
    ai = sum(x["years"][yi]["proj_imp"] - x["years"][yi]["cf_imp"] for x in r)
    a = adj_exports(r, yi, f2f, labor, capw)
    per = []
    for x, ce in zip(r, a["chapters"]):
        di = x["years"][yi]["proj_imp"] - x["years"][yi]["cf_imp"]
        per.append({"hs": x["hs"], "name": x["name"], "d_imp": di, "d_exp_adj": ce["adj"], "d_exp_unadj": ce["unadj"],
                    "net": di - ce["adj"], "loss": ce["loss"]})
    return {"add_imp": ai, "add_exp_adj": a["total_adj"], "add_exp_unadj": a["total_unadj"],
            "widening": ai - a["total_adj"], "cap_loss": a["total_loss"], "chapters": per, "results": r}


def grid(chs, capw, trq_caps, agr=2026, base=2023, labor=0.07):
    g = {}
    for w in GRID["wedge_multipliers"]:
        for f in GRID["f2f_drags"]:
            h = headline(chs, capw, trq_caps, agr, base, w, f, labor)
            g[f"{w}x|{f}"] = h["widening"]
    return g


def monte_carlo(chs, capw, trq_caps, agr=2026, base=2023):
    rng = np.random.default_rng(MC["seed"])
    y5, y10 = {"imp": [], "exp": [], "wid": []}, {"imp": [], "exp": [], "wid": []}
    for _ in range(MC["n_draws"]):
        w = rng.triangular(MC["wedge_mult_low"], MC["wedge_mult_mode"], MC["wedge_mult_high"])
        f = rng.triangular(MC["f2f_drag_low"], MC["f2f_drag_mode"], MC["f2f_drag_high"])
        lab = rng.triangular(MC["labor_drag_low"], MC["labor_drag_mode"], MC["labor_drag_high"])
        r = run_pe(chs, agr, base, trq_caps, w)
        for yh, st in [(5, y5), (10, y10)]:
            ai = sum(x["years"][yh]["proj_imp"] - x["years"][yh]["cf_imp"] for x in r)
            ae = adj_exports(r, yh, f, lab, capw)["total_adj"]
            st["imp"].append(ai); st["exp"].append(ae); st["wid"].append(ai - ae)
    s = lambda a: {"mean": float(np.mean(a)), "lo": float(np.percentile(a, 2.5)), "hi": float(np.percentile(a, 97.5))}
    return {"y5": {k: s(v) for k, v in y5.items()}, "y10": {k: s(v) for k, v in y10.items()}}


if __name__ == "__main__":
    CFG = __import__('os').path.join(__import__('os').path.dirname(__import__('os').path.abspath(__file__)), 'inputs', 'v5_config.json')
    chs = prepare_chapters(json.load(open(CFG))["chapters"])
    h = headline(chs, CAP_WEIGHTS_V75, TRQ_CAPS_V75)
    print(f"v7.5 replication: imports {h['add_imp']/1e6:.1f}  exports adj {h['add_exp_adj']/1e6:.1f}  widening {h['widening']/1e6:.1f}  cap loss {h['cap_loss']/1e6:.2f}")
    for c in sorted(h["chapters"], key=lambda c: -abs(c["net"]))[:8]:
        print(f"  HS {c['hs']}: net {c['net']/1e6:7.1f}  imp {c['d_imp']/1e6:6.1f}  exp {c['d_exp_adj']/1e6:6.1f}")
    s2 = headline(chs, CAP_WEIGHTS_V75, TRQ_CAPS_V75, agr=2027)
    print(f"S2 widening {s2['widening']/1e6:.1f}")
    g = grid(chs, CAP_WEIGHTS_V75, TRQ_CAPS_V75)
    print('grid span', round((max(g.values()) - min(g.values())) / 1e6, 2))
    for e in [-2.0, -2.5, -3.5, -5.0, -8.0]:
        he = headline(chs, CAP_WEIGHTS_V75, TRQ_CAPS_V75, imp_elast=e)
        print(f"  eps {e}: imp {he['add_imp']/1e6:.1f} widening {he['widening']/1e6:.1f}")
    mc = monte_carlo(chs, CAP_WEIGHTS_V75, TRQ_CAPS_V75)
    print('MC y10 widening', {k: round(v / 1e6, 1) for k, v in mc['y10']['wid'].items()}, 'y5', {k: round(v / 1e6, 1) for k, v in mc['y5']['wid'].items()})
