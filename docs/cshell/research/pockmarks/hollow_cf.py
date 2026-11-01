"""Defect 2, part 2 (throwaway): is the hollow design the optimum of the formulation, or a trap?
Phase-0 model with the run's final D_comp = 8 mm, settings at the end of phase 0 (frac = 1).
  1. J of the three phase-0 finals (check: the run logged J ≈ +0.506 for vshell at phase 0 iteration 59).
  2. Local KKT check at the vshell final: per-volume efficiency of closing (∂J/∂ξᵃ ÷ ∂V/∂ξᵃ) vs the multiplier λ,
     split by r_MA quantile (thick regions = joints/nodes) and by term (stiffness/strength/buckling).
  3. Counterfactuals "solid joints": vertices with r_MA ≥ the q-quantile (q = 0.9, 0.75) and the 3 mm neighbourhood
     of the mounts forced closed (ξᵃ = 1, ξᵈ unchanged), mass restored, J evaluated; then MMA re-optimised for N
     iterations with those vertices held closed, versus a free re-optimisation control of the same length.
Usage: python hollow_cf.py <run_dir> <out_dir> [n_reopt]
"""
import json
import os
import sys
import time

import numpy as np
from scipy.spatial import cKDTree

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../../.."))
sys.path.insert(0, ROOT)
from cshell import config as ccfg  # noqa: E402
from cshell.mma import MMA, restore_mass  # noqa: E402
from cshell.objective import capacity, evaluate  # noqa: E402
from cshell.optimise import Context, PhaseModel  # noqa: E402

run_dir, out = sys.argv[1], sys.argv[2]
n_reopt = int(sys.argv[3]) if len(sys.argv) > 3 else 40
os.makedirs(out, exist_ok=True)
cfg = json.load(open(os.path.join(run_dir, "config_used.json")))
cfg["step_file"] = os.path.join(ROOT, "examples/parts/Solid MBB.step")
log_f = open(os.path.join(out, "hollow_cf_log.txt"), "w")


def log(*a):
    s = " ".join(str(x) for x in a)
    print(s, flush=True)
    log_f.write(s + "\n")
    log_f.flush()


t0 = time.time()
ctx = Context(cfg, log=log)
nv = ctx.surf.n_vertices
model = PhaseModel(ctx, 0, 8.0, log=log)
dm, disc, ds = model.dm, model.disc, model.ds
actv = ~ds.passive
act = np.concatenate([actv, actv])
A = ctx.surf.A
opt = cfg["optimiser"]
base = model.baseline(model.settings(0.0))
FRAC = 1.0


def vol(x):
    return dm.volume(dm.forward(x[:nv], x[nv:], as_built=False))


def vol_grad(x):
    gd, ga = dm.volume_grad(dm.forward(x[:nv], x[nv:], as_built=False))
    return np.concatenate([gd, ga])


def fe_parts(rho, st, base_c, c_norm, prev_modes=None):
    off = model.geometry_device == "gpu" and model.geometry_offload
    if off:
        disc.release()
    try:
        if model.ensure_active(rho):
            prev_modes = None
        g = model.grid_fe
        r = g.gather_elems(rho)
        ev = evaluate(model.fe, disc.rho_min + (1 - disc.rho_min) * r, r,
                      g.gather_elems(model.stress_mask_dense).astype(float), st, base_c, c_norm, prev_modes)
    finally:
        if off:
            model.t_upload = disc.upload()
    return ev


def pull(sab, comp):
    cx = dm.cxp
    g = cx.zeros(disc.n_elem)
    g[cx.asarray(model.grid_fe.elem_dense)] = cx.asarray(comp)
    gd, ga = dm.pull_rho_hat(sab, g)
    return np.asarray(gd), np.asarray(ga)


def ev_xi(xi, c_norm, frac=FRAC, prev=None):
    st = model.settings(frac)
    sab = dm.forward(xi[:nv], xi[nv:], as_built=True)
    rho = dm.rho(sab)
    base_c = dict(base)
    lam_s0 = cfg["material"]["strength_MPa"] / (c_norm * base["sigma_pn0"])
    base_c["lam0"] = capacity(lam_s0, base["mu0"], st.knockdown, st.beta)[0]
    return st, sab, fe_parts(rho, st, base_c, c_norm, prev)


def closed_share(xi, sel=None):
    _, xt_a, _, a, _, _ = dm.vertex_fields(xi[:nv], xi[nv:])
    w = A * actv * (1 if sel is None else sel)
    return float((w * (a >= ctx.r_ma)).sum() / w.sum())


def summary(name, xi, c_norm, extra=None):
    st, sab, ev = ev_xi(xi, c_norm)
    r = dict(case=name, J=ev.J, K_K0=base["C0"] / ev.C, lam_s=ev.lam_s, lam_b=ev.lam_b, gov=ev.governing,
             closed=closed_share(xi), mass_err=vol(xi) / model.V0 - 1, **(extra or {}))
    log(json.dumps(r))
    return r, st, sab, ev


def reopt(xi, fixed, n, c_norm):
    """MMA as optimise.Run (phase 0, frac 1 settings, w_s = 0) with `fixed` vertices held at ξᵃ = 1."""
    a_ = act.copy()
    a_[nv:][fixed] = False
    a_[:nv][fixed] = False
    x = xi.copy()
    x[nv:][fixed] = 1.0
    x, _, _ = restore_mass(x, vol, vol_grad, model.V0, a_, tol=opt["mass_tol"])
    mma = MMA(int(a_.sum()), 2, move=opt["move"], con_scale=1.0)
    s_J, tau, prev, Js = None, 5e-4, None, []
    for it in range(n):
        st, sab, ev = ev_xi(x, c_norm, FRAC, prev)
        Js.append(ev.J)
        gd, ga = pull(sab, np.asarray(ev.dJ))
        snom = dm.forward(x[:nv], x[nv:], as_built=False)
        V = dm.volume(snom)
        gVd, gVa = dm.volume_grad(snom)
        gJ = np.concatenate([gd, ga])[a_]
        gV = np.concatenate([np.asarray(gVd), np.asarray(gVa)])[a_] / model.V0
        if s_J is None:
            s_J = 10.0 / max(np.abs(gJ).max() * opt["move"] * 10, 1e-12)
        e = V / model.V0 - 1.0
        xn = x.copy()
        xn[a_] = mma.step(x[a_], -s_J * ev.J, -s_J * gJ, np.array([e - tau, -e - tau]) / tau,
                          np.vstack([gV, -gV]) / tau)
        x, _, _ = restore_mass(xn, vol, vol_grad, model.V0, a_, tol=opt["mass_tol"])
        prev = ev.parts["buckling"].modes if ev.parts["buckling"] is not None else None
    return x, Js


# 1. finals
c_norm = {}
res = {}
xf = {}
for s in ("vshell", "uniform", "neutral"):
    x = np.load(os.path.join(run_dir, f"xi_{s}_phase0.npy"))
    x[:nv][ds.passive] = 0.0
    x[nv:][ds.passive] = 1.0
    xf[s] = x
    res[s], st, sab, ev = summary(f"final_phase0_{s}", x, base["c"])
x = xf["vshell"]
cn = base["c"]

# 2. local efficiency at the vshell final by r_MA quantile and term
st, sab, ev = ev_xi(x, cn)
wsum = st.w_K + st.w_P
w = ev.parts["shares"]
terms = dict(total=np.asarray(ev.dJ), stiff=-st.w_K * np.asarray(ev.parts["dC"]) / ev.C / wsum,
             strength=-st.w_P * w[0] * np.asarray(ev.parts["stress"].grad) / ev.parts["stress"].value / wsum)
if ev.parts["buckling"] is not None and ev.mu_tilde > 0:
    terms["buckling"] = -st.w_P * w[1] * np.asarray(ev.parts["buckling"].dmu_tilde) / ev.mu_tilde / wsum
snom = dm.forward(x[:nv], x[nv:], as_built=False)
gVd, gVa = [np.asarray(g) / model.V0 for g in dm.volume_grad(snom)]
gt = {k: pull(sab, v) for k, v in terms.items()}
gJall = np.concatenate(gt["total"])[act]
gVall = np.concatenate([gVd, gVa])[act]
lam = float(gJall @ gVall / (gVall @ gVall))
# multiplier from the free (interior) ξᵈ variables, which are not at bounds: median of gJ/gV there
xt_d = np.clip(ds.filt.apply(x[:nv]), 0, 1)
free_d = actv & (xt_d > 0.05) & (xt_d < 0.95) & (np.abs(gVd) > 1e-12)
lam_d = float(np.median(gt["total"][0][free_d] / gVd[free_d]))
rq = ctx.r_ma[actv]
loc = dict(lambda_ls=lam, lambda_free_d_median=lam_d, shares=w)
for q in (0.0, 0.5, 0.75, 0.9, 0.97):
    sel = actv & (ctx.r_ma >= np.quantile(rq, q)) & (np.abs(gVa) > 1e-12)
    row = {k: float(np.median(gt[k][1][sel] / gVa[sel])) for k in gt}
    row["frac_closing_favoured"] = float((gt["total"][1][sel] / gVa[sel] > lam_d).mean())
    row["n"] = int(sel.sum())
    row["r_ma_min"] = float(np.quantile(rq, q))
    loc[f"rma_q{q}"] = row
log("local efficiency of ξᵃ (∂J/∂ξᵃ ÷ ∂(V/V0)/∂ξᵃ, median) vs λ:", json.dumps(loc))
res["local"] = loc

# 3. counterfactuals
mount_d, _ = cKDTree(ctx.surf.V[ctx.mount]).query(ctx.surf.V) if ctx.mount.any() else (np.full(nv, np.inf), None)
sets = {"rma_q0.9": actv & (ctx.r_ma >= np.quantile(rq, 0.9)),
        "rma_q0.75": actv & (ctx.r_ma >= np.quantile(rq, 0.75)),
        "near_mounts_3mm": actv & (mount_d <= 3.0)}
for name, sel in sets.items():
    y = x.copy()
    y[nv:][sel] = 1.0
    a_ = act.copy()
    a_[nv:][sel] = False
    a_[:nv][sel] = False
    y, _, _ = restore_mass(y, vol, vol_grad, model.V0, a_, tol=opt["mass_tol"])
    res[f"forced_{name}"], *_ = summary(f"forced_closed_{name}", y, cn,
                                         dict(area_share=float((A * sel).sum() / (A * actv).sum()),
                                              closed_in_set=closed_share(y, sel)))
    xr, Js = reopt(y, sel, n_reopt, cn)
    res[f"reopt_{name}"], *_ = summary(f"reopt_{n_reopt}_{name}", xr, cn,
                                        dict(J_hist=[round(j, 4) for j in Js[::5]], closed_in_set=closed_share(xr, sel)))
xr, Js = reopt(x, np.zeros(nv, bool), n_reopt, cn)
res["reopt_free"], *_ = summary(f"reopt_{n_reopt}_free_control", xr, cn, dict(J_hist=[round(j, 4) for j in Js[::5]]))
json.dump(res, open(os.path.join(out, "hollow_cf.json"), "w"), indent=1, default=float)
log(f"done {time.time() - t0:.0f} s")
