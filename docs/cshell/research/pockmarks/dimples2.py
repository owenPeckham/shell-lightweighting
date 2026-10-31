"""Defect 1, part 2 (throwaway): attribute the vertex-scale roughness of d and a on Γ_h.

1. Rebuild the phase-1 filter (r_f = 3 mm) on Γ_h; check d == d_hi clip(F ξᵈ) and a == a_lo + (a_hi − a_lo) clip(F ξᵃ).
2. Raw ξ statistics (share at the MMA bounds, vertex-scale std) and the filter's point response: F e_i at i vs its
   1-ring (a 2-D Helmholtz kernel ~ K0(ρ/r_f) is log-singular, so isolated raw outliers survive as cones).
3. High-pass (r = 0.375 mm) RMS of d and a on interior active vertices (≥ 3 mm from passive) for variants:
     d:  as run | d_hi clip(F² ξᵈ) (filter applied twice)
     a:  as run | r_MA smoothed (F_1mm) in a_hi | F² ξᵃ | both
4. Patch maps of the field model: P-interpolated heights (d outer, a inner) sampled on Γ_h triangles at 0.05 mm,
   detrended (1 mm Gaussian), as run vs the variants, for the four patches of dimples.py.
Usage: python dimples2.py <run_dir> <out_dir> <dimples_report.json>
"""
import json
import os
import sys

import numpy as np
from scipy import ndimage, sparse as sp
from scipy.interpolate import griddata
from scipy.spatial import cKDTree

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../../.."))
sys.path.insert(0, ROOT)
from cshell.surface import HelmholtzFilter, mesh_operators  # noqa: E402
from cshell.design import softplus  # noqa: E402
import pyvista as pv  # noqa: E402
import matplotlib  # noqa: E402
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402

run_dir, out, rep_path = sys.argv[1], sys.argv[2], sys.argv[3]
os.makedirs(out, exist_ok=True)
cfg = json.load(open(os.path.join(run_dir, "config_used.json")))
R1 = json.load(open(rep_path))
m = pv.read(os.path.join(run_dir, "fields.vtp"))
V = np.asarray(m.points, np.float64)
T = np.asarray(m.faces).reshape(-1, 4)[:, 1:].astype(np.int64)
f = {k: np.asarray(m.point_data[k], np.float64) for k in m.point_data.keys()}
passive = f["passive"].astype(bool)
act = ~passive
d_hi, r_ma = f["d_hi"], f["r_MA"]
M, L, A, N = mesh_operators(V, T)
F3 = HelmholtzFilter(M, L, cfg["filter"]["radius_mm"][1])
F1 = HelmholtzFilter(M, L, 1.0)
Fhp = HelmholtzFilter(M, L, 0.375)
gam = cfg["design"]["closure_gamma"] * cfg["grid"]["phases"][-1]["h_g"]
t_min, w_s = cfg["manufacturing"]["t_min_mm"], cfg["design"]["inset_soft_mm"]
REP = {}


def fields(xd, xa, rma):
    xtd = np.where(passive, 0.0, np.clip(F3.apply(xd), 0, 1))
    xta = np.where(passive, 1.0, np.clip(F3.apply(xa), 0, 1))
    d = d_hi * xtd
    a_lo, _ = softplus(np, t_min - d, w_s)
    a = a_lo + (rma + gam - a_lo) * xta
    return d, a, xtd, xta


xd, xa = f["xi_d"], f["xi_a"]
d0, a0, xtd0, xta0 = fields(xd, xa, r_ma)
REP["check"] = dict(d_maxdiff=float(np.abs(d0 - f["d"]).max()), a_maxdiff=float(np.abs(a0 - f["a"]).max()),
                    xta_maxdiff=float(np.abs(xta0 - f["xit_a"]).max()))
# raw ξ statistics
E = np.concatenate([T[:, [0, 1]], T[:, [1, 2]], T[:, [2, 0]]])
Adj = sp.coo_matrix((np.ones(len(E)), (E[:, 0], E[:, 1])), shape=(len(V),) * 2).tocsr()
Adj = ((Adj + Adj.T) > 0).astype(float)
deg = np.asarray(Adj.sum(1)).ravel()
umb = lambda x: x - (Adj @ x) / np.maximum(deg, 1)  # noqa: E731
dpas, _ = cKDTree(V[passive]).query(V)
inner = act & (dpas > 3.0)
REP["n_interior"] = int(inner.sum())
for nm, x in (("xi_d", xd), ("xi_a", xa)):
    REP[nm] = dict(at0=float((x[inner] <= 1e-6).mean()), at1=float((x[inner] >= 1 - 1e-6).mean()),
                   umbrella_rms=float(np.sqrt(np.mean(umb(x)[inner] ** 2))), mean=float(x[inner].mean()))
# point response of F: centre value vs 1-ring mean, for F and F²
rng = np.random.default_rng(0)
idx = rng.choice(np.flatnonzero(inner), 50, replace=False)
ratios, ratios2, peak = [], [], []
for i in idx:
    e = np.zeros(len(V))
    e[i] = 1.0
    g = F3.apply(e)
    g2 = F3.apply(g)
    nb = Adj[i].indices
    ratios.append(g[i] / g[nb].mean())
    ratios2.append(g2[i] / g2[nb].mean())
    peak.append(g[i] - g[nb].mean())
REP["filter_point_response"] = dict(F_centre_over_ring_median=float(np.median(ratios)),
                                    F2_centre_over_ring_median=float(np.median(ratios2)),
                                    F_cone_height_per_unit_xi_median=float(np.median(peak)),
                                    cone_in_d_mm_for_unit_flip_median=float(np.median(peak) * np.median(d_hi[inner])))
# variants
rma_s = F1.apply(r_ma)
var = {"as_run": (d0, a0),
       "F2_xi": fields(F3.apply(xd), F3.apply(xa), r_ma)[:2],
       "rMA_smooth": fields(xd, xa, rma_s)[:2],
       "F2_xi+rMA_smooth": fields(F3.apply(xd), F3.apply(xa), rma_s)[:2]}


def hp(x):
    y = x - Fhp.apply(x)
    return float(np.sqrt(np.sum((A * y * y)[inner]) / np.sum(A[inner])))


REP["hp_rms_r0.375_interior"] = {k: dict(d=hp(v[0]), a=hp(v[1])) for k, v in var.items()}
REP["hp_rms_r0.375_interior"]["r_MA"] = hp(r_ma)
REP["hp_rms_r0.375_interior"]["xit_d_times_dhi"] = hp(xtd0 * np.median(d_hi))
REP["mean_shift_by_variant_mm"] = {k: dict(d=float(np.mean((v[0] - d0)[inner])), a=float(np.mean((v[1] - a0)[inner])))
                                   for k, v in var.items()}
print(json.dumps(REP, indent=1), flush=True)

# patch maps on Γ_h (barycentric lattice, 0.05 mm)
k = 12
bb = np.array([(i, j) for i in range(k + 1) for j in range(k + 1 - i)], float) / k
B = np.c_[1 - bb.sum(1), bb]                              # (nb, 3)
for pk, c in R1["centres"].items():
    i0 = c["vertex"]
    ctr, n0 = V[i0], N[i0]
    t1 = np.cross(n0, [1.0, 0, 0] if abs(n0[0]) < 0.9 else [0, 1.0, 0])
    t1 /= np.linalg.norm(t1)
    t2 = np.cross(n0, t1)
    tri = np.flatnonzero(np.all(np.abs(V[T].mean(1) - ctr) < 5.0, 1) & (np.abs((V[T].mean(1) - ctr) @ n0) < 1.0))
    P = np.einsum("bk,tkd->tbd", B, V[T[tri]]).reshape(-1, 3)
    u, v = (P - ctr) @ t1, (P - ctr) @ t2
    gu = np.arange(-4, 4.0001, 0.025)
    GU, GV = np.meshgrid(gu, gu, indexing="xy")
    fig, axs = plt.subplots(2, len(var), figsize=(4.4 * len(var), 8.4))
    prow = {}
    for ci, (vn, (dv, av)) in enumerate(var.items()):
        for ri, (side, fv) in enumerate((("outer d", dv), ("inner a", av))):
            val = np.einsum("bk,tk->tb", B, fv[T[tri]]).ravel()
            Z = griddata(np.c_[u, v], val, (GU, GV), method="linear")
            w = np.isfinite(Z).astype(float)
            Zf = np.nan_to_num(Z)
            Dz = np.where(w > 0, Z - ndimage.gaussian_filter(Zf, 40) / np.maximum(ndimage.gaussian_filter(w, 40), 1e-9),
                          np.nan)
            prow[f"{vn}|{side}"] = float(np.nanstd(Dz))
            im = axs[ri, ci].pcolormesh(GU, GV, Dz, cmap="RdBu_r", vmin=-0.06, vmax=0.06, shading="auto")
            axs[ri, ci].set_title(f"{vn}: {side}, rms {np.nanstd(Dz) * 1000:.0f} µm", fontsize=9)
            axs[ri, ci].set_aspect("equal")
    plt.colorbar(im, ax=axs, shrink=0.6, label="P-interpolated height − 1 mm Gaussian (mm)")
    fig.suptitle(f"field model on Γ_h, patch {pk}")
    fig.savefig(os.path.join(out, f"variants_{pk}.png"), dpi=100)
    plt.close(fig)
    REP[f"patch_{pk}_detrended_rms"] = prow
    print(pk, json.dumps(prow), flush=True)
json.dump(REP, open(os.path.join(out, "dimples2_report.json"), "w"), indent=1)
