"""Defect 2, part 3 (throwaway): "solid joints" imposed AFTER the filter (ξ̃ᵃ := 1 on a vertex set, as the passive
rule does for mounts), so the set is exactly closed (a = a_hi = r_MA + γ h_g) while ξᵈ stays free; J before and after
40 MMA iterations on the rest, versus a free 40-iteration control. Reuses the set-up of hollow_cf.py (sections before
"# 1. finals"). Usage: python hollow_cf2.py <run_dir> <out_dir> [n_reopt]
"""
import os

_here = os.path.dirname(os.path.abspath(__file__))
_src = open(os.path.join(_here, "hollow_cf.py")).read().split("# 1. finals")[0]
_src = _src.replace("hollow_cf_log.txt", "hollow_cf2_log.txt").replace("__file__", repr(os.path.join(_here, "hollow_cf.py")))
exec(compile(_src, "hollow_cf_setup", "exec"))

from cshell.design import DesignMap  # noqa: E402

FORCE = np.zeros(nv, bool)  # noqa: F821


def vf_forced(xi_d, xi_a):
    xt_d, xt_a, d, a, a_lo, da_lo = DesignMap.vertex_fields(dm, xi_d, xi_a)  # noqa: F821
    xt_a = np.where(FORCE, 1.0, xt_a)
    a = a_lo + (ds.a_hi - a_lo) * xt_a  # noqa: F821
    return xt_d, xt_a, d, a, a_lo, da_lo


def pv_forced(st, g_d, g_aused):
    g_a = g_aused * st.dab_da
    g_d = g_d + g_a * st.da_lo * (1.0 - st.xit_a)
    g_xt_d = np.where(ds.passive, 0.0, g_d * ds.d_hi)  # noqa: F821
    g_xt_a = np.where(ds.passive | FORCE, 0.0, g_a * (ds.a_hi - st.a_lo))  # noqa: F821
    return ds.filt.adjoint(g_xt_d), ds.filt.adjoint(g_xt_a)  # noqa: F821


dm.vertex_fields = vf_forced  # noqa: F821
dm.pull_vertex = pv_forced  # noqa: F821
cn = base["c"]  # noqa: F821
x = np.load(os.path.join(run_dir, "xi_vshell_phase0.npy"))  # noqa: F821
x[:nv][ds.passive] = 0.0  # noqa: F821
x[nv:][ds.passive] = 1.0  # noqa: F821
res = {}
res["final"], *_ = summary("final_phase0_vshell", x, cn)  # noqa: F821
rq = ctx.r_ma[actv]  # noqa: F821
mount_d, _ = cKDTree(ctx.surf.V[ctx.mount]).query(ctx.surf.V)  # noqa: F821
sets = {"rma_q0.9": actv & (ctx.r_ma >= np.quantile(rq, 0.9)),  # noqa: F821
        "rma_q0.75": actv & (ctx.r_ma >= np.quantile(rq, 0.75)),  # noqa: F821
        "near_mounts_5mm": actv & (mount_d <= 5.0)}  # noqa: F821
xr, Js = reopt(x, np.zeros(nv, bool), n_reopt, cn)  # noqa: F821
res["reopt_free"], *_ = summary(f"reopt_{n_reopt}_free_control", xr, cn,  # noqa: F821
                                dict(J_hist=[round(j, 4) for j in Js[::5]]))
for name, sel in sets.items():
    FORCE[:] = sel
    a_ = act.copy()  # noqa: F821
    a_[nv:][sel] = False
    y, _, _ = restore_mass(x.copy(), vol, vol_grad, model.V0, a_, tol=opt["mass_tol"])  # noqa: F821
    ex = dict(area_share=float((A * sel).sum() / (A * actv).sum()), closed_in_set=closed_share(y, sel))  # noqa: F821
    res[f"forced_{name}"], *_ = summary(f"post_filter_closed_{name}", y, cn, ex)  # noqa: F821
    xr, Js = reopt(y, np.zeros(nv, bool), n_reopt, cn)  # noqa: F821  (set's ξᵃ has zero gradient: inert)
    res[f"reopt_{name}"], *_ = summary(f"reopt_{n_reopt}_post_filter_closed_{name}", xr, cn,  # noqa: F821
                                        dict(J_hist=[round(j, 4) for j in Js[::5]],
                                             closed_in_set=closed_share(xr, sel), area_share=ex["area_share"]))  # noqa: F821
    FORCE[:] = False
json.dump(res, open(os.path.join(out, "hollow_cf2.json"), "w"), indent=1, default=float)  # noqa: F821
log(f"done {time.time() - t0:.0f} s")  # noqa: F821
