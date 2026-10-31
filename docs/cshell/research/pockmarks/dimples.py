"""Defect 1 diagnosis (throwaway): where do the small dimples on final_design.stl come from?

A. Γ_h fields (fields.vtp): roughness of d, a, d_hi, r_MA, ξ̃ vs wavelength (high-pass by surface Helmholtz filters of
   radius r: RMS of f − F_r f), vertex-scale umbrella roughness, clip saturation shares.
B. Patches (boxes of half-width HB mm around chosen Γ_h vertices): crop final_design.stl and Γ₀ (original.stl);
   re-evaluate ψ on the export grid (h = 0.05 mm) stage by stage with the export's own code:
     raw   : ψ = max(φ − P(d+δ), −φ − P a)                (stage A, nominal, with the closure rule)
     open  : void cells removed by the v_min opening set to −h/2 (stage B as in _psi_parts)
     amb   : + resolve_ambiguous_faces (stage E1)
   Marching cubes of each; height maps of the outer (φ) and void (−φ) surfaces in the patch tangent plane for the
   STL and each stage, detrended (Gaussian 1 mm) to show dimples; dimple count/depth/spacing; residual φ − P d.
Usage: python dimples.py <run_dir> <out_dir>
"""
import json
import os
import sys
import time

import numpy as np
from scipy import ndimage, sparse as sp
from scipy.sparse import linalg as spla
from scipy.spatial import cKDTree

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../../.."))
sys.path.insert(0, ROOT)
from cshell import export as cex  # noqa: E402
from cshell.surface import mesh_operators  # noqa: E402

run_dir, out = sys.argv[1], sys.argv[2]
os.makedirs(out, exist_ok=True)
H = 0.05
HB = 6.0
REP = {}
SIGNCHK = {}
import faulthandler
faulthandler.enable()
t00 = time.time()


def log(*a):
    print(*a, flush=True)


# ------------------------------------------------------------------------------------------------ A. fields
import pyvista as pv  # noqa: E402

m = pv.read(os.path.join(run_dir, "fields.vtp"))
V = np.asarray(m.points, np.float64)
T = np.asarray(m.faces).reshape(-1, 4)[:, 1:].astype(np.int64)
pdn = {k: np.asarray(m.point_data[k]) for k in m.point_data.keys()}
log("arrays:", {k: v.shape for k, v in pdn.items()})
d, a, r_ma, d_hi = pdn["d"], pdn["a"], pdn["r_MA"], pdn["d_hi"]
xta = pdn["xit_a"]
passive = pdn["passive"].astype(bool)
act = ~passive
xtd = np.where(d_hi > 1e-9, d / np.maximum(d_hi, 1e-9), 0.0)
M, L, Av, N = mesh_operators(V, T)
E = np.concatenate([T[:, [0, 1]], T[:, [1, 2]], T[:, [2, 0]]])
el = np.linalg.norm(V[E[:, 0]] - V[E[:, 1]], axis=1)
REP["gamma_h"] = dict(nv=len(V), nt=len(T), edge_mean=float(el.mean()), edge_p5=float(np.percentile(el, 5)),
                      edge_p95=float(np.percentile(el, 95)))
Aa = Av * act
At = Aa.sum()
closed = xta >= 0.98
REP["saturation"] = dict(
    xtd_eq0=float((Aa * (xtd <= 1e-9)).sum() / At), xtd_eq1=float((Aa * (xtd >= 1 - 1e-9)).sum() / At),
    xta_eq0=float((Aa * (xta <= 1e-9)).sum() / At), xta_eq1=float((Aa * (xta >= 1 - 1e-9)).sum() / At),
    closed_snap=float((Aa * closed).sum() / At), closed_rma=float((Aa * (a >= r_ma)).sum() / At),
    xta_hist=np.histogram(xta[act], bins=[0, 1e-9, .1, .3, .5, .7, .9, .98, 1 - 1e-9, 1.0001], weights=Aa[act])[0]
    .tolist())
# umbrella (graph Laplacian with uniform weights) and Helmholtz high-pass
Adj = sp.coo_matrix((np.ones(len(E)), (E[:, 0], E[:, 1])), shape=(len(V),) * 2).tocsr()
Adj = ((Adj + Adj.T) > 0).astype(float)
deg = np.asarray(Adj.sum(1)).ravel()
mdiag = np.asarray(M.diagonal()).ravel()
filt = {}
for r in (0.375, 0.75, 1.5, 3.0):
    filt[r] = spla.splu(sp.csc_matrix(sp.diags(mdiag) + r * r * L))
fields = dict(d=d, a=a, d_hi=d_hi, r_MA=r_ma, xit_d=xtd, xit_a=xta, w=np.where(a < r_ma, a + d, np.nan))
rough = {}
umb = {}
for k, f in fields.items():
    f = np.nan_to_num(f, nan=0.0)
    u = f - (Adj @ f) / np.maximum(deg, 1)
    umb[k] = u
    row = dict(umbrella_rms=float(np.sqrt((Aa * u * u).sum() / At)), umbrella_p99=float(np.percentile(np.abs(u[act]), 99)))
    for r, lu in filt.items():
        hp = f - lu.solve(mdiag * f)
        row[f"hp_rms_r{r}"] = float(np.sqrt((Aa * hp * hp).sum() / At))
    rough[k] = row
REP["roughness"] = rough
log(json.dumps(REP, indent=1))

# ------------------------------------------------------------------------------------------------ B. patches
summ = json.load(open(os.path.join(run_dir, "summary.json")))
delta = float(summ["export"]["best"]["mass_correction"]["delta"])
REP["delta"] = delta
openv = act & ~closed & (a < r_ma)
rng = np.random.default_rng(1)
cent = {}
cent["max_umb_d"] = int(np.flatnonzero(openv)[np.argmax(np.abs(umb["d"][openv]))])
cent["max_umb_a"] = int(np.flatnonzero(openv)[np.argmax(np.abs(umb["a"][openv]))])
# typical flat open vertex: normal close to its 3 mm neighbourhood mean
tr = cKDTree(V)
cand = rng.choice(np.flatnonzero(openv & (d > 0.5)), 400, replace=False)
best = []
for i in cand:
    nb = tr.query_ball_point(V[i], 4.0)
    best.append((float(np.abs(N[nb] @ N[i]).min()), i))
best.sort(reverse=True)
cent["typical_flat_1"] = int(best[0][1])
cent["typical_flat_2"] = int(best[5][1])
REP["centres"] = {k: dict(vertex=v, xyz=V[v].tolist(), d=float(d[v]), a=float(a[v]), r_ma=float(r_ma[v]),
                          d_hi=float(d_hi[v]), xta=float(xta[v])) for k, v in cent.items()}
log(json.dumps(REP["centres"], indent=1))
boxes = {k: (V[v] - HB, V[v] + HB) for k, v in cent.items()}


def crop_stl(path, boxes, margin):
    n = int(np.fromfile(path, dtype="<u4", count=1, offset=80)[0])
    dt = np.dtype([("n", "<f4", 3), ("v", "<f4", (3, 3)), ("a", "<u2")])
    mm = np.memmap(path, dtype=dt, mode="r", offset=84, shape=(n,))
    outs = {k: [] for k in boxes}
    for s in range(0, n, 4_000_000):
        tv = np.asarray(mm["v"][s:s + 4_000_000], np.float64)
        c = tv.mean(1)
        for k, (lo, hi) in boxes.items():
            sel = np.all((c >= lo - margin) & (c <= hi + margin), 1)
            if sel.any():
                outs[k].append(tv[sel])
    res = {}
    for k, L_ in outs.items():
        tv = np.concatenate(L_) if L_ else np.zeros((0, 3, 3))
        P = tv.reshape(-1, 3)
        key = np.round(P / 1e-5).astype(np.int64)
        _, idx, inv = np.unique(key, axis=0, return_index=True, return_inverse=True)
        G = inv.reshape(-1, 3).astype(np.int64)
        G = G[(G[:, 0] != G[:, 1]) & (G[:, 1] != G[:, 2]) & (G[:, 0] != G[:, 2])]
        _, ui = np.unique(np.sort(G, 1), axis=0, return_index=True)
        res[k] = (np.ascontiguousarray(P[idx]), np.ascontiguousarray(G[np.sort(ui)]))
    return res


t0 = time.time()
g0 = crop_stl(os.path.join(run_dir, "original.stl"), boxes, 10.0)
gS = crop_stl(os.path.join(run_dir, "final_design.stl"), boxes, 0.0)
log(f"cropped STLs {time.time() - t0:.0f} s: " + str({k: (len(g0[k][1]), len(gS[k][1])) for k in boxes}))

from skimage.measure import marching_cubes  # noqa: E402
import matplotlib  # noqa: E402
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.colors import LightSource  # noqa: E402
from scipy.interpolate import griddata  # noqa: E402


def frame(n):
    t1 = np.cross(n, [1.0, 0, 0] if abs(n[0]) < 0.9 else [0, 1.0, 0])
    t1 /= np.linalg.norm(t1)
    return t1, np.cross(n, t1)


def height_map(P, val, ctr, t1, t2, half=4.0, hs=0.025):
    u, v = (P - ctr) @ t1, (P - ctr) @ t2
    sel = (np.abs(u) < half + 0.3) & (np.abs(v) < half + 0.3)
    gu = np.arange(-half, half + 1e-9, hs)
    GU, GV = np.meshgrid(gu, gu, indexing="xy")
    if sel.sum() < 10:
        return GU, GV, np.full(GU.shape, np.nan)
    Z = griddata(np.c_[u[sel], v[sel]], val[sel], (GU, GV), method="linear")
    return GU, GV, Z


def detrend(Z, sig_mm=1.0, hs=0.025):
    w = np.isfinite(Z).astype(float)
    Zf = np.nan_to_num(Z)
    s = sig_mm / hs
    num = ndimage.gaussian_filter(Zf, s)
    den = ndimage.gaussian_filter(w, s)
    return np.where(w > 0, Z - num / np.maximum(den, 1e-9), np.nan)


def dimples(D, hs=0.025, thr=0.01, win_mm=0.3):
    Dn = np.where(np.isfinite(D), D, 0.0)
    mn = ndimage.minimum_filter(Dn, size=int(round(win_mm / hs)) | 1)
    pk = (Dn == mn) & (Dn < -thr) & np.isfinite(D)
    ij = np.argwhere(pk)
    o = dict(n=int(len(ij)), area_mm2=float(np.isfinite(D).sum() * hs * hs), rms=float(np.nanstd(D)),
             p1=float(np.nanpercentile(D, 1)) if np.isfinite(D).any() else None)
    if len(ij) > 1:
        dd, _ = cKDTree(ij * hs).query(ij * hs, k=2)
        o.update(depth_median=float(-np.median(Dn[pk])), depth_max=float(-Dn[pk].min()),
                 nn_spacing_median=float(np.median(dd[:, 1])))
    return o, ij * hs


for k, (lo, hi) in boxes.items():
    t0 = time.time()
    i0 = cent[k]
    log(k, "start patch", len(g0[k][0]), len(g0[k][1]))
    V0c, T0c = g0[k]
    dz = cex.Design(V0c, T0c, V, T, d, a, d_hi, active=act, closed=closed, sign="pseudonormal")
    n = int(round(2 * HB / H))
    ax = [lo[j] + (np.arange(n) + 0.5) * H for j in range(3)]
    X = np.stack(np.meshgrid(*ax, indexing="ij"), -1).reshape(-1, 3)
    phi, C = dz.phi_closest(X)
    tri, b1, b2 = dz.locate(C)
    dat = dz.d_at(tri, b1, b2, delta)
    aat = dz.a_at(tri, b1, b2)
    o = (phi - dat).reshape(n, n, n)
    pin = (-phi - aat).reshape(n, n, n)
    log(k, "phi/locate done", f"{time.time() - t0:.0f} s")
    SIGNCHK[k] = (V0c, T0c, X[::7].copy(), phi[::7].copy())
    sign_mismatch = None
    psi_raw = np.maximum(o, pin)
    void = pin > 0
    R2 = cex.ball_r2(0.5 * 3.0, H)
    opened = cex.morph_open(void, R2)
    removed = void & ~opened
    pin_o = np.where(removed, -0.5 * H, pin)
    psi_open = np.maximum(o, pin_o)
    eps = np.float32(1e-3 * H)
    psi_amb = psi_open.astype(np.float32)
    psi_amb = np.where(np.abs(psi_amb) < eps, np.where(psi_amb < 0, -eps, eps), psi_amb).astype(np.float32)
    psi_amb_in = psi_amb.copy()
    nflip, nit = cex.resolve_ambiguous_faces(psi_amb, eps)
    flipped = (psi_amb_in < 0) & (psi_amb > 0)
    flipped_raw = None
    pr = psi_raw.astype(np.float32)
    pr = np.where(np.abs(pr) < eps, np.where(pr < 0, -eps, eps), pr).astype(np.float32)
    nflip_raw, _ = cex.resolve_ambiguous_faces(pr.copy(), eps)
    meshes = {}
    for name, psi in (("raw", psi_raw), ("open", psi_open), ("amb", psi_amb)):
        vv, ff, _, _ = marching_cubes(psi, 0.0, allow_degenerate=False)
        meshes[name] = (lo + (vv + 0.5) * H, ff)
    meshes["STL"] = gS[k]
    # classify mesh vertices, heights and residuals vs the field model
    nrm = N[i0]
    t1, t2 = frame(nrm)
    ctr = V[i0]
    hm = {}
    pr_rep = dict(n_cells=n ** 3, sign_mismatch_wn_vs_pn_sub7=sign_mismatch, opened_removed_cells=int(removed.sum()),
                  amb_flipped_nodes=int(nflip), amb_flipped_nodes_raw=int(nflip_raw),
                  flipped_near_outer=int((flipped & (np.abs(o) < 2 * H)).sum()),
                  flipped_near_inner=int((flipped & (np.abs(pin_o) < 2 * H)).sum()))
    for name, (P, F) in meshes.items():
        if len(P) == 0:
            continue
        ph, Cm = dz.phi_closest(P)
        trm, c1, c2 = dz.locate(Cm)
        om = ph - dz.d_at(trm, c1, c2, delta)
        im = -ph - dz.a_at(trm, c1, c2)
        outer = np.abs(om) <= np.abs(im)
        inb = np.all(np.abs(P - ctr) < HB - 1.6, 1)
        r = dict(n_vertices=int(len(P)),
                 outer_resid_rms=float(np.sqrt(np.mean(om[outer & inb] ** 2))) if (outer & inb).any() else None,
                 outer_resid_p01=float(np.percentile(om[outer & inb], 1)) if (outer & inb).any() else None,
                 outer_resid_p99=float(np.percentile(om[outer & inb], 99)) if (outer & inb).any() else None,
                 inner_resid_rms=float(np.sqrt(np.mean(im[~outer & inb] ** 2))) if (~outer & inb).any() else None,
                 inner_resid_p01=float(np.percentile(im[~outer & inb], 1)) if (~outer & inb).any() else None,
                 inner_resid_p99=float(np.percentile(im[~outer & inb], 99)) if (~outer & inb).any() else None)
        # height maps: outer uses φ (offset above Γ₀), inner uses −φ (depth); only the sheet nearest the centre
        for side, sel, val, mod in (("outer", outer, ph, ph - om), ("inner", ~outer, -ph, -ph - im)):
            s2 = sel & (np.abs((P - ctr) @ nrm) < HB)
            GU, GV, Z = height_map(P[s2], val[s2], ctr, t1, t2)
            _, _, Zm = height_map(P[s2], mod[s2], ctr, t1, t2)
            Dz = detrend(Z)
            dr, pts = dimples(Dz)
            r[f"{side}_dimples"] = dr
            r[f"{side}_field_rough_rms"] = float(np.nanstd(detrend(Zm)))
            hm[(name, side)] = (GU, GV, Z, Dz, Zm, pts)
        pr_rep[name] = r
    # Γ_h vertices projected into the patch (for spacing comparison)
    nb = np.flatnonzero(np.all(np.abs(V - ctr) < 4.5, 1) & (np.abs((V - ctr) @ nrm) < 1.0))
    gh_uv = np.c_[(V[nb] - ctr) @ t1, (V[nb] - ctr) @ t2]
    for side in ("outer", "inner"):
        if ("STL", side) in hm and len(hm[("STL", side)][5]) > 3 and len(gh_uv):
            pts = hm[("STL", side)][5] - 4.0
            dv, _ = cKDTree(gh_uv).query(pts)
            rnd = rng.uniform(-4, 4, (2000, 2))
            dr, _ = cKDTree(gh_uv).query(rnd)
            pr_rep[f"STL_{side}_dimple_to_gh_vertex_median"] = float(np.median(dv))
            pr_rep[f"random_to_gh_vertex_median"] = float(np.median(dr))
    REP[f"patch_{k}"] = pr_rep
    log(k, json.dumps(pr_rep, indent=1), f"{time.time() - t0:.0f} s")
    # figure: rows = STL, raw, open, amb; columns = outer detrended, inner detrended, outer hillshade
    names = [x for x in ("STL", "raw", "open", "amb")]
    fig, axs = plt.subplots(len(names), 4, figsize=(17, 4.1 * len(names)))
    ls = LightSource(azdeg=315, altdeg=35)
    for ri, name in enumerate(names):
        for ci, side in enumerate(("outer", "inner")):
            if (name, side) not in hm:
                continue
            GU, GV, Z, Dz, Zm, pts = hm[(name, side)]
            im_ = axs[ri, ci].pcolormesh(GU, GV, Dz, cmap="RdBu_r", vmin=-0.06, vmax=0.06, shading="auto")
            plt.colorbar(im_, ax=axs[ri, ci], label="height − 1 mm Gaussian (mm)")
            axs[ri, ci].plot(gh_uv[:, 0], gh_uv[:, 1], "k.", ms=1.5)
            axs[ri, ci].set_title(f"{name}: {side} surface detrended")
            axs[ri, ci].set_aspect("equal")
            Zf = np.nan_to_num(Z, nan=np.nanmean(Z) if np.isfinite(Z).any() else 0)
            axs[ri, 2 + ci].imshow(ls.hillshade(Zf, vert_exag=10, dx=0.025, dy=0.025), cmap="gray", origin="lower",
                                   extent=[-4, 4, -4, 4])
            axs[ri, 2 + ci].set_title(f"{name}: {side} hillshade (10x)")
    fig.suptitle(f"patch {k} at {np.round(ctr, 2).tolist()}  d={d[i0]:.2f} a={a[i0]:.2f} r_MA={r_ma[i0]:.2f}; "
                 "black dots: Γ_h vertices")
    fig.tight_layout()
    fig.savefig(os.path.join(out, f"patch_{k}.png"), dpi=110)
    plt.close(fig)
    # field model heights (P d, P a) vs STL
    fig, axs = plt.subplots(1, 2, figsize=(10, 4.3))
    for ci, side in enumerate(("outer", "inner")):
        if ("STL", side) in hm:
            GU, GV, Z, Dz, Zm, pts = hm[("STL", side)]
            im_ = axs[ci].pcolormesh(GU, GV, detrend(Zm), cmap="RdBu_r", vmin=-0.06, vmax=0.06, shading="auto")
            plt.colorbar(im_, ax=axs[ci])
            axs[ci].plot(gh_uv[:, 0], gh_uv[:, 1], "k.", ms=1.5)
            axs[ci].set_title(f"model P{'(d+δ)' if side == 'outer' else 'a'} detrended ({side})")
            axs[ci].set_aspect("equal")
    fig.tight_layout()
    fig.savefig(os.path.join(out, f"patch_{k}_fieldmodel.png"), dpi=110)
    plt.close(fig)
    np.savez_compressed(os.path.join(out, f"patch_{k}.npz"), **{f"{nm}_{sd}_Z": hm[(nm, sd)][2] for (nm, sd) in hm},
                        **{f"{nm}_{sd}_Zmodel": hm[(nm, sd)][4] for (nm, sd) in hm})
    # 3D render of the STL crop (offscreen) if VTK can render here
    try:
        pv.OFF_SCREEN = True
        P, F = gS[k]
        sm = pv.PolyData(P, np.c_[np.full(len(F), 3), F].ravel())
        pl = pv.Plotter(off_screen=True, window_size=(1400, 1000))
        pl.add_mesh(sm, color="lightgrey", smooth_shading=False, specular=0.4)
        pl.camera_position = [tuple(ctr + 25 * nrm + 8 * t1), tuple(ctr), tuple(t2)]
        pl.screenshot(os.path.join(out, f"patch_{k}_3d_outer.png"))
        pl.camera_position = [tuple(ctr - 25 * nrm + 8 * t1), tuple(ctr), tuple(t2)]
        pl.screenshot(os.path.join(out, f"patch_{k}_3d_inner_side.png"))
        pl.close()
    except Exception as e:  # noqa: BLE001
        log("3D render failed:", repr(e)[:200])

json.dump(REP, open(os.path.join(out, "dimples_report.json"), "w"), indent=1, default=float)
# (iv) sign check last (C++ abort risk): generalised winding number vs pseudonormal sign
import igl  # noqa: E402
for k, (V0c, T0c, Xs, phs) in SIGNCHK.items():
    w = igl.fast_winding_number(V0c, T0c, Xs)
    REP[f"patch_{k}"]["sign_mismatch_wn_vs_pn"] = int(((w > 0.5) != (phs < 0)).sum())
    REP[f"patch_{k}"]["sign_checked"] = int(len(Xs))
    log(k, "sign mismatches", REP[f"patch_{k}"]["sign_mismatch_wn_vs_pn"], "of", len(Xs))
json.dump(REP, open(os.path.join(out, "dimples_report.json"), "w"), indent=1, default=float)
log(f"done {time.time() - t00:.0f} s")
