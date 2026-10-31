"""Defect 2 diagnosis (throwaway): why are the cshell designs hollow? Phase 0 (h_e 0.8, h_g 0.2) of the MBB only.

1. Build Context + PhaseModel(phase 0) exactly as optimise.Run does, with the dshell (vshell) start.
2. Re-run the phase-0 MMA loop for N iterations (same code path: as-built forward, evaluate, pull back, restore mass)
   and log per iteration: closed share (a >= r_MA), snap share (xit_a >= 0.98), xit_a histogram, and for the set S0
   of vertices closed at the start: share still closed, mean xit_a.
3. At selected iterations split dJ/dxi_a into stiffness, strength, buckling terms (and volume gradient) and report
   their sums and per-volume efficiencies over S0 versus the open vertices.
4. Counterfactual: the converged phase-0 vshell design (xi_vshell_phase0.npy) vs the same design with S0 forced
   closed (mass restored on the rest by the optimiser's own restore_mass): J of each.
Usage: python hollow_trace.py <run_dir> <out_dir> [n_iter]
"""
import json
import os
import sys
import time

import numpy as np

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../../.."))
sys.path.insert(0, ROOT)
from cshell import config as ccfg  # noqa: E402
from cshell.collect import dshell_start_xi  # noqa: E402
from cshell.mma import MMA, restore_mass  # noqa: E402
from cshell.objective import capacity, evaluate  # noqa: E402
from cshell.optimise import Context, PhaseModel  # noqa: E402

run_dir, out = sys.argv[1], sys.argv[2]
n_iter = int(sys.argv[3]) if len(sys.argv) > 3 else 25
os.makedirs(out, exist_ok=True)
cfg = json.load(open(os.path.join(run_dir, "config_used.json")))
cfg["step_file"] = os.path.join(ROOT, "examples/parts/Solid MBB.step")
log_f = open(os.path.join(out, "hollow_trace_log.txt"), "w")


def log(*a):
    s = " ".join(str(x) for x in a)
    print(s, flush=True)
    log_f.write(s + "\n")
    log_f.flush()


t0 = time.time()
ctx = Context(cfg, log=log)
nv = ctx.surf.n_vertices
xi_start = dshell_start_xi(ctx, cfg, log=log)
model = PhaseModel(ctx, 0, cfg["design"]["outward_extent_initial_mm"], log=log)
dm, disc, ds = model.dm, model.disc, model.ds
act = np.concatenate([~ds.passive, ~ds.passive])
actv = ~ds.passive
A = ctx.surf.A
opt = cfg["optimiser"]
st0 = model.settings(0.0)
base = model.baseline(st0)
log(f"setup {time.time() - t0:.0f} s, base {json.dumps({k: float(v) for k, v in base.items() if isinstance(v, (int, float))})}")


def vol(x):
    return dm.volume(dm.forward(x[:nv], x[nv:], as_built=False))


def vol_grad(x):
    gd, ga = dm.volume_grad(dm.forward(x[:nv], x[nv:], as_built=False))
    return np.concatenate([gd, ga])


def fe_parts(rho, st, base_c, c_norm, prev_modes):
    off = model.geometry_device == "gpu" and model.geometry_offload
    if off:
        disc.release()
    try:
        if model.ensure_active(rho):
            prev_modes = None
        g = model.grid_fe
        r = g.gather_elems(rho)
        rho_hat = disc.rho_min + (1 - disc.rho_min) * r
        m = g.gather_elems(model.stress_mask_dense).astype(float)
        ev = evaluate(model.fe, rho_hat, r, m, st, base_c, c_norm, prev_modes)
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


def stats(xi, S0=None):
    xt_d, xt_a, d, a, _, _ = dm.vertex_fields(xi[:nv], xi[nv:])
    Aa = A * actv
    At = Aa.sum()
    closed = a >= ctx.r_ma
    o = dict(closed=float((Aa * closed).sum() / At), snap=float((Aa * (xt_a >= 0.98)).sum() / At),
             hist_xta=np.histogram(xt_a[actv], bins=[0, 1e-9, 0.1, 0.5, 0.9, 0.98, 1 - 1e-9, 1.0001], weights=Aa[actv])[0]
             .round(1).tolist(),
             d_mean=float((Aa * d).sum() / At))
    if S0 is not None:
        o["S0_closed"] = float((Aa * closed * S0).sum() / (Aa * S0).sum())
        o["S0_xta_mean"] = float((Aa * xt_a * S0).sum() / (Aa * S0).sum())
        o["S0_gap_mm_median"] = float(np.median((ctx.r_ma - a)[S0]))
    return o, closed, xt_a


def evaluate_xi(xi, c_norm, it_frac=0.0, prev_modes=None):
    st = model.settings(it_frac)
    sab = dm.forward(xi[:nv], xi[nv:], as_built=True)
    rho = dm.rho(sab)
    base_c = dict(base)
    lam_s0 = cfg["material"]["strength_MPa"] / (c_norm * base["sigma_pn0"])
    base_c["lam0"] = capacity(lam_s0, base["mu0"], st.knockdown, st.beta)[0]
    ev = fe_parts(rho, st, base_c, c_norm, prev_modes)
    return st, sab, ev


def decompose(st, sab, ev, S0, open_v, xi):
    """dJ = w_K(-dC/C) + w_P(-w_s dσ/σ - w_b dμ/μ), / (w_K + w_P); pulled back to ξᵃ and ξᵈ."""
    wsum = st.w_K + st.w_P
    sp, bk, w = ev.parts["stress"], ev.parts["buckling"], ev.parts["shares"]
    comps = dict(stiff=-st.w_K * np.asarray(ev.parts["dC"]) / ev.C / wsum,
                 strength=-st.w_P * w[0] * np.asarray(sp.grad) / sp.value / wsum)
    if bk is not None and ev.mu_tilde > 0:
        comps["buckling"] = -st.w_P * w[1] * np.asarray(bk.dmu_tilde) / ev.mu_tilde / wsum
    snom = dm.forward(xi[:nv], xi[nv:], as_built=False)
    gVd, gVa = dm.volume_grad(snom)
    gVa = np.asarray(gVa) / model.V0
    gVd = np.asarray(gVd) / model.V0
    res = dict(shares=w, governing=ev.governing)
    # Lagrange multiplier estimate: least squares of gJ ≈ λ gV over all active design variables
    gJd_t, gJa_t = pull(sab, np.asarray(ev.dJ))
    gJ = np.concatenate([gJd_t, gJa_t])[act]
    gV = np.concatenate([gVd, gVa])[act]
    lam = float(gJ @ gV / (gV @ gV))
    res["lambda_ls"] = lam
    for k, c in comps.items():
        gd, ga = pull(sab, c)
        res[k] = dict(S0_sum_a=float(ga[S0].sum()), open_sum_a=float(ga[open_v].sum()),
                      S0_sum_d=float(gd[S0].sum()), open_sum_d=float(gd[open_v].sum()))
    res["volume"] = dict(S0_sum_a=float(gVa[S0].sum()), open_sum_a=float(gVa[open_v].sum()))
    # per-variable net KKT residual r = gJ - λ gV on ξᵃ: > 0 means "wants more ξᵃ (closing)" at price λ
    ra = gJa_t - lam * gVa
    res["net_a"] = dict(S0_frac_positive=float((ra[S0] > 0).mean()), S0_median=float(np.median(ra[S0])),
                        open_median=float(np.median(ra[open_v])),
                        S0_eff_median=float(np.median(gJa_t[S0] / np.where(np.abs(gVa[S0]) > 0, gVa[S0], np.nan))
                                            if S0.any() else np.nan),
                        open_eff_median=float(np.nanmedian(gJa_t[open_v] / np.where(np.abs(gVa[open_v]) > 0, gVa[open_v],
                                                                                    np.nan))))
    res["dab_da_S0_median"] = float(np.median(np.asarray(sab.dab_da)[S0]))
    return res


# ------------------------------------------------------------------------------------------------- trajectory
xi = xi_start.copy()
xi[:nv][ds.passive] = 0.0
xi[nv:][ds.passive] = 1.0
xi, eta, nres = restore_mass(xi, vol, vol_grad, model.V0, act, tol=opt["mass_tol"])
s, closed0, xta0 = stats(xi)
S0 = actv & (xta0 >= 0.98)
log(f"start (after restore): {json.dumps(s)}; S0 active vertices {int(S0.sum())}, area share "
    f"{float((A * S0).sum() / (A * actv).sum()):.4f}; closed by a>=r_MA {int((closed0 & actv).sum())}")
np.save(os.path.join(out, "S0_mask.npy"), S0)
optimiser = MMA(int(act.sum()), 2, move=opt["move"], con_scale=1.0)
c_norm, s_J, tau, prev_modes = base["c"], None, 5e-4, None
max_it = ccfg.phase_value(opt["max_iter"], 0)
rows = []
for it in range(n_iter):
    frac = it / max(max_it - 1, 1)
    st, sab, ev = evaluate_xi(xi, c_norm, frac, prev_modes)
    s, closed, xta = stats(xi, S0)
    open_v = actv & ~closed & ~S0
    row = dict(it=it, J=ev.J, C_C0=ev.C / base["C0"], gov=ev.governing, **s)
    if it in (0, 1, 2, 4, 6, 9, 14, 19, n_iter - 1):
        row["decomp"] = decompose(st, sab, ev, S0, open_v, xi)
    log(json.dumps(row))
    rows.append(row)
    gJd, gJa = pull(sab, np.asarray(ev.dJ))
    snom = dm.forward(xi[:nv], xi[nv:], as_built=False)
    V = dm.volume(snom)
    gVd, gVa = dm.volume_grad(snom)
    w_s = ccfg.phase_value(opt["sliver_weight"], 0, frac)
    Gs, _ = dm.sliver(snom)
    gGd, gGa = dm.sliver_grad(snom)
    gJ = np.concatenate([gJd, gJa])[act]
    gG = np.concatenate([np.asarray(gGd), np.asarray(gGa)])[act]
    gV = np.concatenate([np.asarray(gVd), np.asarray(gVa)])[act] / model.V0
    if s_J is None:
        s_J = 10.0 / max(np.abs(gJ).max() * opt["move"] * 10, 1e-12)
    f0 = s_J * (-ev.J + w_s * Gs)
    df0 = s_J * (-gJ + w_s * gG)
    e = V / model.V0 - 1.0
    g = np.array([e - tau, -e - tau]) / tau
    dg = np.vstack([gV, -gV]) / tau
    xn = xi.copy()
    xn[act] = optimiser.step(xi[act], f0, df0, g, dg)
    xi, _, _ = restore_mass(xn, vol, vol_grad, model.V0, act, tol=opt["mass_tol"])
    prev_modes = ev.parts["buckling"].modes if ev.parts["buckling"] is not None else None
    if st.w_P > 0 and np.isfinite(ev.parts["stress"].quantile):
        c_norm = float(np.sqrt(c_norm) * np.sqrt(ev.parts["stress"].quantile / ev.sigma_pn))
np.save(os.path.join(out, f"xi_trace_it{n_iter}.npy"), xi)

# ------------------------------------------------------------------------------------------------- counterfactual
xf = np.load(os.path.join(run_dir, "xi_vshell_phase0.npy"))
for name, x in [("final_phase0", xf.copy())]:
    x[:nv][ds.passive] = 0.0
    x[nv:][ds.passive] = 1.0
    st, sab, ev = evaluate_xi(x, base["c"])
    s, _, _ = stats(x, S0)
    log(json.dumps(dict(case=name, J=ev.J, C_C0=ev.C / base["C0"], lam_s=ev.lam_s, lam_b=ev.lam_b, gov=ev.governing, **s)))
    # force S0 closed: raw ξᵃ = 1 on S0 and its 1-ring dilation by filter (ξ̃ᵃ≥0.98 needs ξ = 1 over ~r_f)
    for grow in (0.0, 2.0):
        y = x.copy()
        if grow > 0:
            from scipy.spatial import cKDTree
            tr = cKDTree(ctx.surf.V[S0])
            dd, _ = tr.query(ctx.surf.V)
            sel = actv & (dd <= grow)
        else:
            sel = S0
        y[nv:][sel] = 1.0
        y[:nv][sel] = 0.0
        act_rest = act.copy()
        act_rest[nv:][sel] = False
        act_rest[:nv][sel] = False
        y, _, _ = restore_mass(y, vol, vol_grad, model.V0, act_rest, tol=opt["mass_tol"])
        st, sab, ev = evaluate_xi(y, base["c"])
        s, _, _ = stats(y, S0)
        log(json.dumps(dict(case=f"S0_forced_closed_grow{grow}", J=ev.J, C_C0=ev.C / base["C0"], lam_s=ev.lam_s,
                            lam_b=ev.lam_b, gov=ev.governing, mass_err=vol(y) / model.V0 - 1, **s)))
json.dump(rows, open(os.path.join(out, "hollow_trace.json"), "w"), indent=1, default=float)
log(f"done {time.time() - t0:.0f} s")
