"""Independent re-implementation of the paper's partial-equilibrium projection (Analysis C engine).

Written from the published formula (manuscript Section 4 / README), not copied from orchestrator_v7_5.py or pe_v81.py:
  counterfactual  M_cf(y) = B * (1+g)^(agr + y - base)
  price term      pi(y)   = (1 + tau(y)) / (1 + tau0) - 1     (trade-weighted over lines when lines are given)
  wedge factor    W(y)    = 1 + w * mult * min(1, y / max(phase_eu, 1))
  imports         M(y)    = M_cf(y) * (1 + eps_m * pi_EU(y) * W(y))       [optionally capped by a chapter TRQ cap]
  exports         X(y)    = X_cf(y) * (1 + eps_x * pi_MER(y))
  capacity adj.   dX_adj  = dX * (1 - kappa_c * (f2f + labor * y/10))
  widening(y)     = sum_c dM_c(y) - sum_c dX_adj_c(y)
Line-level TRQ logic (marginal price, pooled notional Greek slice) is re-implemented from the v8.1 description.
"""
import os, json, math, copy
import numpy as np

CAPW_V75 = {"23": 0.80, "12": 0.70, "09": 0.40, "24": 1.00, "08": 1.00, "20": 1.00, "22": 0.90, "04": 1.00, "15": 1.00}
# v8.1 (run_v81.py) extends the capacity weights to the added agri-food chapters and drops the v7.5 HS 04 chapter cap,
# because the line-level TRQ logic supersedes it. Analysis C uses the same convention for B and C runs.
CAPW_V81 = dict(CAPW_V75, **{"02": 1.0, "07": 1.0, "10": 1.0, "16": 0.8, "17": 0.8, "21": 0.5, "03": 0.0})
TRQ_CAP_V75 = {"04": 8_640_000}
ENV_ADDON, ENV_CH = 0.005, {"23"}


def sign_wedges(chapters):
    """Reproduce the orchestrator's load-time treatment: +0.005 env-tax addon on HS 23, then incidence signing
    (exporter-borne magnitude enters negative, importer-side addon positive)."""
    out = copy.deepcopy(chapters)
    for c in out:
        w = c.get("wedge", 0.0)
        addon = ENV_ADDON if c["hs"] in ENV_CH else 0.0
        if w == 0.0 and addon == 0.0:
            continue
        c["wedge"] = addon - w  # w is the exporter-borne magnitude stored positive in the config
    return out


def tau_line(line, y):
    if "r_path" in line:
        return line["tau0"] * line["r_path"][min(y, len(line["r_path"]) - 1)]
    t0, tgt, ph = line["tau0"], line.get("target", 0.0), line["phase"]
    if ph < 0:
        return t0
    if ph == 0 or y >= ph:
        return tgt
    return t0 - (t0 - tgt) * y / ph


def tau_chapter(t0, tgt, phase, reduction_type, y):
    if reduction_type == "None" or phase == 0 and tgt == t0:
        return t0
    if phase == 0 or y >= phase:
        return tgt
    return t0 - (t0 - tgt) * y / phase


def trq_price_change(line, y, m_line, eps):
    """Marginal-price TRQ: tiers filled cheapest first; if the marginal unit falls inside a tier the in-quota
    price change applies but the quantity response is capped at that tier's cumulative slice; once all slices
    bind there is no response. Year 0 covers eight months (pro-rated)."""
    t0 = line["tau0"]
    tiers = [(line["trq"]["r_in"], line["trq"]["slice"])]
    if "slice2" in line["trq"]:
        tiers.append((line["trq"]["r_in2"], line["trq"]["slice2"]))
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


def pi_eu(c, y, m_cf, eps):
    if "eu_lines" in c:
        tot = 0.0
        for l in c["eu_lines"]:
            if "trq" in l:
                tot += l["share"] * trq_price_change(l, y, l["share"] * m_cf, eps)
            else:
                tot += l["share"] * ((1 + tau_line(l, y)) / (1 + l["tau0"]) - 1)
        return tot
    cur = tau_chapter(c["eu_mfn"], c["target_eu"], c["phase_eu"], c["reduction_type"], y)
    return (1 + cur) / (1 + c["eu_mfn"]) - 1


def pi_mer(c, y):
    if "mer_lines" in c:
        return sum(l["share"] * ((1 + tau_line(l, y)) / (1 + l["tau0"]) - 1) for l in c["mer_lines"])
    if c["phase_mer"] == 0:
        return 0.0
    cur = 0.0 if y >= c["phase_mer"] else c["mer_mfn"] * (1 - y / c["phase_mer"])
    return (1 + cur) / (1 + c["mer_mfn"]) - 1


def project(chapters, agr=2026, base=2023, wedge_mult=1.0, eps_m=None, trq_caps=None, years=range(11)):
    trq_caps = TRQ_CAP_V75 if trq_caps is None else trq_caps
    out = {}
    for c in chapters:
        e = c["imp_elast"] if eps_m is None else eps_m
        rows = {}
        for y in years:
            m_cf = c["baseline_imp"] * (1 + c["cagr_imp"]) ** (agr + y - base)
            x_cf = c["baseline_exp"] * (1 + c["cagr_exp"]) ** (agr + y - base)
            W = 1 + c["wedge"] * wedge_mult * min(1, y / max(c["phase_eu"], 1))
            m = m_cf * (1 + e * pi_eu(c, y, m_cf, e) * W)
            if c["hs"] in trq_caps:
                m = min(m, trq_caps[c["hs"]])
            x = x_cf * (1 + c["exp_elast"] * pi_mer(c, y))
            rows[y] = {"m_cf": m_cf, "x_cf": x_cf, "m": m, "x": x, "dM": m - m_cf, "dX": x - x_cf}
        out[c["hs"]] = rows
    return out


def widening(proj, y, capw=None, f2f=0.10, labor=0.07):
    capw = CAPW_V75 if capw is None else capw
    dM = sum(r[y]["dM"] for r in proj.values())
    dX_adj = 0.0
    per = {}
    for hs, r in proj.items():
        k = capw.get(hs, 0.0)
        adj = r[y]["dX"] * (1 - k * (f2f + (labor * y / 10 if y > 0 else 0.0)))
        dX_adj += adj
        per[hs] = {"dM": r[y]["dM"], "dX_adj": adj, "net": r[y]["dM"] - adj}
    return {"dM": dM, "dX_adj": dX_adj, "widening": dM - dX_adj, "per": per}


def headline(chapters, **kw):
    y = kw.pop("y", 10)
    capw = kw.pop("capw", None); f2f = kw.pop("f2f", 0.10); labor = kw.pop("labor", 0.07)
    return widening(project(chapters, **kw), y, capw, f2f, labor)


def monte_carlo(chapters, n=1000, seed=42, capw=None, trq_caps=None):
    rng = np.random.default_rng(seed)
    w10 = []
    for _ in range(n):
        wm = rng.triangular(0.5, 1.0, 3.0); f = rng.triangular(0.05, 0.10, 0.20); lab = rng.triangular(0.03, 0.07, 0.12)
        p = project(chapters, wedge_mult=wm, years=[5, 10], trq_caps=trq_caps)
        w10.append(widening(p, 10, capw, f, lab)["widening"])
    return {"mean": float(np.mean(w10)), "lo": float(np.percentile(w10, 2.5)), "hi": float(np.percentile(w10, 97.5))}


if __name__ == "__main__":
    import sys, io, os
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    R = os.environ.get('EUMG_PACKAGE', os.path.join(os.path.dirname(os.path.abspath(__file__)), '..'))
    # 1) published v7.5 inputs -> must give 197.1 / 51.1 / 146.0
    v5 = sign_wedges(json.load(open(os.path.join(R, "model/inputs/v5_config.json"), encoding="utf-8"))["chapters"])
    h = headline(v5)
    print(f"A  v7.5 inputs : dM {h['dM']/1e6:7.2f}  dX_adj {h['dX_adj']/1e6:6.2f}  widening {h['widening']/1e6:7.2f}   (published 197.1 / 51.1 / 146.0)")
    # 2) corrected v8.1 inputs (already signed in file? check) -> must give 59.7 / 43.4 / 16.3
    v81 = json.load(open(os.path.join(R, "outputs/v81_config.json"), encoding="utf-8"))["chapters"]
    h2 = headline(sign_wedges(v81))
    print(f"B  v8.1 inputs : dM {h2['dM']/1e6:7.2f}  dX_adj {h2['dX_adj']/1e6:6.2f}  widening {h2['widening']/1e6:7.2f}   (published 59.7 / 43.4 / 16.3)")
    for e in [-2.0, -2.5, -5.0, -8.0]:
        print(f"     eps {e}: widening {headline(sign_wedges(v81), eps_m=e)['widening']/1e6:.1f}")
    print("   S2 (2027):", round(headline(sign_wedges(v81), agr=2027)['widening']/1e6, 1))
