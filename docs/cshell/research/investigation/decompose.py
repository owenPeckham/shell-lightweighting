"""Per-member strain-energy decomposition, buckling participation and hollowing of a cshell run (investigation).

    python docs/cshell/research/investigation/decompose.py <run_dir> <dshell_run_dir> <out_dir> [--phase 1]

Read-only use of the cshell library. For the SOLID and the run's FINAL design (phase `--phase` grid):
  * exact element energies U_e = ½ u_eᵀK_e u_e and material centre stresses σ_e (q1_im, exact at the centre);
  * member labels from dshell's segmentation of the same STEP (arrays.npz seg_xyz/seg_label, 1 mm grid, nearest);
  * beam-section decomposition: every element belongs to its nearest skeleton station (same region); per station a
    weighted least-squares fit σ_tt = a + b·y + c·z over the section (t = local skeleton tangent, y, z centred section
    coordinates, weight ρ̂ h³): a = axial (sign = tension/compression), b·y + c·z = bending, rest = residual (warping,
    local); τ_tn = transverse shear + torsion. Energies use the uniaxial form σ²/2E, τ²/2G; "other" = centre-stress
    energy − the parts (transverse normal stresses, in-section shear, Poisson coupling);
  * buckling: per retained mode φ_j (K-normalised) the modal strain-energy share φ_eᵀK_eφ_e and the destabilising
    geometric term c_e = −detJ Σ_g σ_g·Γ_g(φ) (Σ_e c_e = μ_j = 1/λ_j, checked) per member;
  * wall-normal chord split on the solid: σ along x − s n, s ∈ [0, 2 r_MA] (outer surface to the opposite surface
    through the medial axis), in-plane stresses fitted linear in s: mean = membrane, slope = bending; the optimiser's
    U_flex (half depth, membrane+bending lumped), U_shear and R_v (PhaseModel.shear_flexure) for comparison;
  * per member: hollow-allowed (H) and open (a < r_MA) area fractions of the final design, mean outward offset d.
Outputs: members_<design>.csv, modes_<design>.csv, members_surface.csv, summary.json in out_dir.
"""
from __future__ import annotations

import argparse
import csv
import dataclasses
import json
import os
import sys
import time

import numpy as np
from scipy.spatial import cKDTree

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(HERE, "../../../..")))

from cshell import config as ccfg                                       # noqa: E402
from cshell.design import solid_start_xi_a                              # noqa: E402
from cshell.fe_sparse import _gamma_voigt, _host, _ramp, unit_D         # noqa: E402
from cshell.objective import ElementStressSampler, wall_frames          # noqa: E402
from cshell.optimise import Context, PhaseModel, load_design, vertex_fields   # noqa: E402
from dshell.segment import _local_tangents                              # noqa: E402


def log(*a):
    print(*a, flush=True)


def read_regions(path):
    with open(path) as f:
        return {int(r["id"]): r for r in csv.DictReader(f)}


def perp_frame(t):
    t = t / np.maximum(np.linalg.norm(t, axis=1, keepdims=True), 1e-300)
    ax = np.eye(3)[np.argmin(np.abs(t), axis=1)]
    e1 = np.cross(t, ax)
    e1 /= np.linalg.norm(e1, axis=1, keepdims=True)
    return t, e1, np.cross(t, e1)


class Segm:
    def __init__(self, dshell_dir):
        z = np.load(os.path.join(dshell_dir, "arrays.npz"))
        self.seg_xyz, self.seg_label = z["seg_xyz"].astype(float), z["seg_label"].astype(int)
        self.skel, self.skel_reg = z["skel_pts"].astype(float), z["skel_region"].astype(int)
        self.tan = _local_tangents(self.skel, 3.0)
        self.regions = read_regions(os.path.join(dshell_dir, "regions.csv"))
        self.kd = cKDTree(self.seg_xyz)
        self.kd_skel = {r: (cKDTree(self.skel[self.skel_reg == r]), np.flatnonzero(self.skel_reg == r))
                        for r in np.unique(self.skel_reg)}
        self.kd_all = cKDTree(self.skel)

    def label(self, X):
        d, i = self.kd.query(X, workers=-1)
        return self.seg_label[i], d

    def station(self, X, lab):
        st = np.empty(len(X), np.int64)
        for r in np.unique(lab):
            sel = np.flatnonzero(lab == r)
            if r in self.kd_skel:
                kd, idx = self.kd_skel[r]
                st[sel] = idx[kd.query(X[sel], workers=-1)[1]]
            else:
                st[sel] = self.kd_all.query(X[sel], workers=-1)[1]
        return st


def element_centres(model):
    g = model.grid_fe
    ijk = np.stack(np.unravel_index(np.asarray(g.elem_dense), model.elem_shape), 1)
    return np.asarray(model.origin, float) + (ijk + 0.5) * model.h_e


def section_decomposition(X, sig, w, st, seg, E, nu):
    """Per element: U_ax, U_bend, U_res, U_shear (uniaxial forms) and station axial stress a; w = ρ̂ h³."""
    G = E / (2 * (1 + nu))
    ns = len(seg.skel)
    t, e1, e2 = perp_frame(seg.tan[st])
    T = np.zeros((len(X), 3, 3))
    T[:, 0, 0], T[:, 1, 1], T[:, 2, 2] = sig[:, 0], sig[:, 1], sig[:, 2]
    T[:, 1, 2] = T[:, 2, 1] = sig[:, 3]
    T[:, 0, 2] = T[:, 2, 0] = sig[:, 4]
    T[:, 0, 1] = T[:, 1, 0] = sig[:, 5]
    Tt = np.einsum("eij,ej->ei", T, t)
    s_tt = (Tt * t).sum(1)
    tau = (Tt * e1).sum(1) ** 2 + (Tt * e2).sum(1) ** 2
    q = X - seg.skel[st]
    y, z = (q * e1).sum(1), (q * e2).sum(1)
    bc = lambda v: np.bincount(st, v, minlength=ns)                    # noqa: E731
    W = np.maximum(bc(w), 1e-300)
    y0, z0 = bc(w * y) / W, bc(w * z) / W
    y, z = y - y0[st], z - z0[st]
    a = bc(w * s_tt) / W
    Syy, Syz, Szz = bc(w * y * y), bc(w * y * z), bc(w * z * z)
    Sy, Sz = bc(w * s_tt * y), bc(w * s_tt * z)
    M = np.stack([np.stack([Syy, Syz], -1), np.stack([Syz, Szz], -1)], -2) + 1e-12 * np.eye(2)
    bcoef = np.linalg.solve(M, np.stack([Sy, Sz], -1)[..., None])[..., 0]
    s_ax = a[st]
    s_b = bcoef[st, 0] * y + bcoef[st, 1] * z
    s_r = s_tt - s_ax - s_b
    return dict(U_ax=w * s_ax ** 2 / (2 * E), U_bend=w * s_b ** 2 / (2 * E), U_res=w * s_r ** 2 / (2 * E),
                U_shear=w * tau / (2 * G), a_station=a, s_ax=s_ax)


def geometric_terms(fe, u, phi):
    """c_e = −detJ Σ_g σ_g·Γ_g(φ) with the buckling() stress (ramped, k_σ mask); Σ c_e = φᵀGφ = μ."""
    xp = fe.xp
    rho = fe.rho_hat
    r, _ = _ramp(rho, fe._min_density_inv)
    km = getattr(fe, "ksig_mask", None)
    if km is not None:
        r = r * xp.asarray(km, dtype=xp.float64)
    u_e = u[fe._dofs]
    sig = fe.E * xp.einsum("gij,ej->egi", fe._DB, u_e) * (rho * r)[:, None, None]
    phi_e = phi[fe._dofs].reshape(-1, 8, 3)
    grad = xp.einsum("gia,eak->egki", fe._Bg, phi_e)
    Gam = _gamma_voigt(xp, grad)
    return _host(-fe._detJ * xp.einsum("egi,egi->e", sig, Gam))


def analyse(model, rho, tag, seg, cfg, out_dir):
    t0 = time.time()
    st = dataclasses.replace(model.settings(1.0), objective="compliance", constraints=("buckling", "strength"),
                             buckling=True)
    res = model.fe_eval(rho, st, None, want_all=True)
    fe = model.fe
    E, nu = cfg["material"]["E_MPa"], cfg["material"]["nu"]
    pr = res["primal"]
    U_exact = np.asarray(pr.U_e)
    rh = _host(fe.rho_hat)
    w = rh * model.h_e ** 3
    sig = _host(fe.centre_stress(pr.u, solid=True))
    Dinv = np.linalg.inv(unit_D(nu)) / E
    U_c = 0.5 * w * np.einsum("ei,ij,ej->e", sig, Dinv, sig)
    X = element_centres(model)
    lab, dist = seg.label(X)
    stn = seg.station(X, lab)
    dec = section_decomposition(X, sig, w, stn, seg, E, nu)
    # buckling modes
    fe._min_density_inv = cfg["capacity"]["min_density_geometric"]
    modes = res["modes"]
    nm = 0 if modes is None else int(modes.shape[1])
    e_modes, c_modes, mus = [], [], []
    for j in range(nm):
        phi = modes[:, j]
        e_modes.append(fe.element_bilinear(phi, phi, solid=False))
        c = geometric_terms(fe, pr.u, phi)
        c_modes.append(c)
        mus.append(float(c.sum()))
    p = cfg["capacity"]["buckling_p"]
    mu_arr = np.array(mus)
    mu_tilde_chk = float(mu_arr.max() * np.sum((mu_arr / mu_arr.max()) ** p) ** (1 / p)) if nm else 0.0
    ids = sorted(seg.regions)
    Ut = U_exact.sum()
    rows = []
    for r in ids:
        m = lab == r
        if not m.any():
            continue
        st_in = np.unique(stn[m])
        a_st = dec["a_station"]
        # axial energy split by the sign of the station's mean axial stress
        pos = m & (dec["s_ax"] > 0)
        neg = m & (dec["s_ax"] < 0)
        row = dict(id=r, kind=seg.regions[r]["kind"], dshell_final=seg.regions[r]["final"],
                   vol_mm3=float(w[m].sum()), U_share=float(U_exact[m].sum() / Ut),
                   f_ax_tension=float(dec["U_ax"][pos].sum() / U_c[m].sum()),
                   f_ax_compression=float(dec["U_ax"][neg].sum() / U_c[m].sum()),
                   f_bend=float(dec["U_bend"][m].sum() / U_c[m].sum()),
                   f_shear=float(dec["U_shear"][m].sum() / U_c[m].sum()),
                   f_res=float(dec["U_res"][m].sum() / U_c[m].sum()),
                   f_other=float(1 - (dec["U_ax"][m].sum() + dec["U_bend"][m].sum() + dec["U_shear"][m].sum()
                                      + dec["U_res"][m].sum()) / U_c[m].sum()),
                   N_sign=float(np.sign(np.median(a_st[st_in]))) if st_in.size else 0.0,
                   sigma_ax_med_MPa=float(np.median(a_st[st_in])) if st_in.size else np.nan)
        for j in range(min(nm, 6)):
            row[f"mode{j + 1}_E"] = float(e_modes[j][m].sum() / e_modes[j].sum())
            row[f"mode{j + 1}_G"] = float(c_modes[j][m].sum() / c_modes[j].sum())
        rows.append(row)
    path = os.path.join(out_dir, f"members_{tag}.csv")
    with open(path, "w", newline="") as f:
        wr = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        wr.writeheader()
        wr.writerows(rows)
    glob = dict(C=float(pr.compliance), U_exact=float(Ut), U_centre=float(U_c.sum()),
                centre_over_exact=float(U_c.sum() / Ut), half_C_over_U=float(0.5 * pr.compliance / Ut),
                f_ax=float(dec["U_ax"].sum() / U_c.sum()), f_bend=float(dec["U_bend"].sum() / U_c.sum()),
                f_shear=float(dec["U_shear"].sum() / U_c.sum()), f_res=float(dec["U_res"].sum() / U_c.sum()),
                mu=mus, lam_raw=[1 / m_ if m_ > 0 else np.inf for m_ in mus], mu_tilde_check=mu_tilde_chk,
                mu_tilde_optimiser=float(res["mu_tilde"]), lam_B=float(res["lam_b"]),
                sigma_pn=float(res["sigma_pn"]), n_elem=int(len(w)), label_dist_p99_mm=float(np.percentile(dist, 99)),
                mode_energy_sums=[float(e.sum()) for e in e_modes], seconds=time.time() - t0)
    log(f"[decomp] {tag}: {json.dumps(glob, default=float)[:900]}")
    return glob, rows, res


def chord_split(model, res, ctx, seg, n=24):
    """Solid only: wall-normal chord from the outer surface through the medial axis to depth 2 r_MA."""
    fe = model.fe
    xp = fe.xp
    sig = fe.centre_stress(res["primal"].u, solid=True)
    smp = ElementStressSampler(sig, res["rho_active"], model.grid_fe.elem_dense, model.elem_shape, model.origin,
                               model.h_e, xp=xp)
    E, nu = ctx.cfg["material"]["E_MPa"], ctx.cfg["material"]["nu"]
    G = E / (2 * (1 + nu))
    t1, t2, nn = wall_frames(np.asarray(ctx.N, float))
    X = np.asarray(ctx.surf.V, float)
    L = 2.0 * np.asarray(ctx.r_ma, float)
    nv = len(X)
    frac = (np.arange(n) + 0.5) / n
    out = {k: np.zeros(nv) for k in ("U_mem", "U_bend", "U_res", "U_sh", "s_mem_t")}
    per = 20000
    for a0 in range(0, nv, per):
        b0 = min(nv, a0 + per)
        s = -frac[None, :] * L[a0:b0, None]
        P = X[a0:b0, None, :] + s[..., None] * nn[a0:b0, None, :]
        sg, _ = smp(xp.asarray(P.reshape(-1, 3)))
        sg = _host(sg).reshape(b0 - a0, n, 6)
        T = np.zeros(sg.shape[:2] + (3, 3))
        T[..., 0, 0], T[..., 1, 1], T[..., 2, 2] = sg[..., 0], sg[..., 1], sg[..., 2]
        T[..., 1, 2] = T[..., 2, 1] = sg[..., 3]
        T[..., 0, 2] = T[..., 2, 0] = sg[..., 4]
        T[..., 0, 1] = T[..., 1, 0] = sg[..., 5]
        e1, e2, en = (v[a0:b0, None, :] for v in (t1, t2, nn))
        c = lambda u_, v_: np.einsum("mki,mkij,mkj->mk", np.broadcast_to(u_, T.shape[:3]), T,   # noqa: E731
                                     np.broadcast_to(v_, T.shape[:3]))
        comps = [c(e1, e1), c(e2, e2), c(e1, e2)]
        wgt = (L[a0:b0] / n)[:, None]
        sc = frac - frac.mean()
        mem = bend = resid = 0.0
        for k, comp in enumerate(comps):
            mean = comp.mean(1, keepdims=True)
            slope = (comp * sc).sum(1, keepdims=True) / (sc ** 2).sum()
            lin = slope * sc
            r_ = comp - mean - lin
            mod = 2 * G if k == 2 else 2 * E
            mem = mem + (mean ** 2 * wgt).sum(1) / mod
            bend = bend + (lin ** 2 * wgt).sum(1) / mod
            resid = resid + (r_ ** 2 * wgt).sum(1) / mod
        sh = ((c(e1, en) ** 2 + c(e2, en) ** 2) * wgt).sum(1) / (2 * G)
        out["U_mem"][a0:b0], out["U_bend"][a0:b0], out["U_res"][a0:b0], out["U_sh"][a0:b0] = mem, bend, resid, sh
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("run_dir")
    ap.add_argument("dshell_dir")
    ap.add_argument("out_dir")
    ap.add_argument("--phase", type=int, default=1)
    a = ap.parse_args()
    os.makedirs(a.out_dir, exist_ok=True)
    cfg = ccfg.normalise_starts(json.load(open(os.path.join(a.run_dir, "config_used.json"))))
    ctx = Context(cfg, log=log)
    seg = Segm(a.dshell_dir)
    Vb = np.asarray(ctx.surf.V)
    log(f"[decomp] bbox cshell {Vb.min(0).round(1)} {Vb.max(0).round(1)}; dshell seg "
        f"{seg.seg_xyz.min(0).round(1)} {seg.seg_xyz.max(0).round(1)}")
    nv = ctx.surf.n_vertices
    vlab, vdist = seg.label(Vb)
    A = np.asarray(ctx.surf.A)
    S = dict(run_dir=a.run_dir, phase=a.phase, vertex_label_dist_p99_mm=float(np.percentile(vdist, 99)))

    # ---- phase-0 solid: the optimiser's R_v as used at the first growth boundary ---------------------------------
    d0 = load_design(a.run_dir, phase=0)
    m0 = PhaseModel(ctx, 0, d0["D_comp"], log, hollow=None)
    st0 = dataclasses.replace(m0.settings(1.0), objective="compliance", constraints=(), buckling=False)
    r0 = m0.fe_eval(m0.rho0, st0, None)
    xs = np.zeros(2 * nv)
    xs[nv:] = solid_start_xi_a(m0.ds)
    R0 = m0.shear_flexure(r0, xs, int(cfg["growth"].get("samples", 16)))
    pool0 = ~m0.ds.passive
    m0.release()
    del m0

    # ---- phase-p model: solid and final design ---------------------------------------------------------------------
    des = load_design(a.run_dir, phase=a.phase)
    model = PhaseModel(ctx, a.phase, des["D_comp"], log, hollow=des["hollow"])
    g_s, rows_s, res_s = analyse(model, model.rho0, "solid", seg, cfg, a.out_dir)
    # chord split on the solid (the phase-p grid) and the optimiser's half-depth R_v at phase p
    cs = chord_split(model, res_s, ctx, seg)
    Rp = model.shear_flexure(res_s, xs, int(cfg["growth"].get("samples", 16)))
    xi = np.asarray(des["xi"], float)
    sab = model.dm.forward(xi[:nv], xi[nv:], as_built=True)
    g_f, rows_f, _ = analyse(model, model.dm.rho(sab), "final", seg, cfg, a.out_dir)
    vf = vertex_fields(ctx, xi, phase=a.phase, D_comp=des["D_comp"], hollow=des["hollow"])
    model.release()

    # ---- surface table per member ------------------------------------------------------------------------------------
    act = ~vf["passive"]
    order = np.argsort(R0[pool0])
    Ap = A[pool0][order]
    cdf = np.empty(pool0.sum())
    cdf[order] = np.cumsum(Ap) / Ap.sum()
    qrank = np.full(nv, np.nan)
    qrank[pool0] = cdf                                     # growth quantile at which the vertex enters H
    srows = []
    for r in sorted(seg.regions):
        m = vlab == r
        if not m.any():
            continue
        ma = m & act
        Am = A[m].sum()
        Ama = max(A[ma].sum(), 1e-300)
        U = {k: float((A[m] * v[m]).sum()) for k, v in cs.items() if k != "s_mem_t"}
        Uflex_c = U["U_mem"] + U["U_bend"] + U["U_res"]
        srows.append(dict(
            id=r, kind=seg.regions[r]["kind"], area_mm2=float(Am), active_frac=float(A[ma].sum() / Am),
            q_entry_median=float(np.nanmedian(qrank[m & pool0])) if (m & pool0).any() else np.nan,
            R0_area_mean=float((A[m & pool0] * R0[m & pool0]).sum() / max(A[m & pool0].sum(), 1e-300)),
            chord_f_membrane=U["U_mem"] / max(Uflex_c, 1e-300), chord_f_bending=U["U_bend"] / max(Uflex_c, 1e-300),
            R_lumped_chord=U["U_sh"] / max(Uflex_c, 1e-300), R_bend_only_chord=U["U_sh"] / max(U["U_bend"], 1e-300),
            Rp_area_mean=float((A[ma] * Rp[ma]).sum() / Ama),
            H_frac=float((A[ma] * vf["hollow_allowed"][ma]).sum() / Ama),
            open_frac=float((A[ma] * (~vf["closed"][ma])).sum() / Ama),
            d_mean_mm=float((A[ma] * vf["d"][ma]).sum() / Ama),
            wall_open_med_mm=float(np.nanmedian(vf["w"][ma])) if np.isfinite(vf["w"][ma]).any() else np.nan))
    with open(os.path.join(a.out_dir, "members_surface.csv"), "w", newline="") as f:
        wr = csv.DictWriter(f, fieldnames=list(srows[0].keys()))
        wr.writeheader()
        wr.writerows(srows)
    S.update(solid=g_s, final=g_f)
    json.dump(S, open(os.path.join(a.out_dir, "summary.json"), "w"), indent=1, default=float)
    np.savez_compressed(os.path.join(a.out_dir, "vertex_fields.npz"), label=vlab.astype(np.int16), R0=R0, Rp=Rp,
                        qrank=qrank, closed=vf["closed"], hollow=vf["hollow_allowed"], d=vf["d"], passive=vf["passive"],
                        **{k: v for k, v in cs.items()})
    log("[decomp] done")


if __name__ == "__main__":
    main()
