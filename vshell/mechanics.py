"""
Performance metrics of one analysed design.

Stiffness   K = |F|^2 / C,  C = f . u (compliance) -- the slope of the load-deflection line along F.
Strength    lambda_s = sigma_allow / q-quantile of von Mises (material stress, volume weighted), over the
            design regions, excluding cells within `exclude_bc_mm` of fixed or loaded nodes. In hollow
            regions stresses are scaled by s / (s - bias) for the as-built wall.
Euler       per strut (see strut_section): axial force N = (1/L) sum sigma_tt^hom h^3 along the strut's own
            skeleton tangents (compression < 0); I_min from cell offsets to a smoothed centroid line,
            L = length actually covered by the strut's cells (as-built factor (s - bias)/s for hollow struts);
            P_cr = pi^2 E I_min / (K_eff L_m)^2,  L_m = L_cells + 2 joint_k r;  lambda_E = P_cr / |N|.
Local       hollow struts: thin cylinder, sigma_cr = gamma E t / (r_m sqrt(3 (1 - nu^2))), with the
buckling    NASA SP-8007 knock-down gamma = 1 - 0.901 (1 - exp(-sqrt(r_m / t) / 16)).
            hollow joints: simply supported plate, sigma_cr = 4 pi^2 E / (12 (1 - nu^2)) (t / b)^2, b = 2 r.
            t = s - bias (as built). Demand: q-quantile of the compressive principal stress in the region.
Capacity    lambda = min(lambda_s, lambda_E, lambda_L),  P = lambda |F|.
Shear/bend  R_i = U_shear / U_bend with sigma_tt = t.sigma.t and tau_tn on the two normals of the local
            skeleton tangent t: U_bend = sum sigma_tt^2 / (2E) V, U_shear = sum (tau_1^2 + tau_2^2) / (2G) V.
"""
from __future__ import annotations

import numpy as np
from scipy.spatial import cKDTree

from . import fe


def weighted_quantile(values, weights, q):
    if len(values) == 0:
        return float("nan")
    o = np.argsort(values)
    v, w = values[o], weights[o]
    cw = np.cumsum(w)
    return float(np.interp(q * cw[-1], cw, v))


def local_frames(cell_xyz, seg):
    """Nearest skeleton point, its tangent t and two unit normals for each cell centre."""
    kt = cKDTree(seg.skel_pts)
    _, j = kt.query(cell_xyz)
    t = seg.skel_tan[j]
    tmp = np.where(np.abs(t[:, :1]) < 0.9, np.array([[1.0, 0, 0]]), np.array([[0, 1.0, 0]]))
    n1 = np.cross(t, tmp)
    n1 /= np.linalg.norm(n1, axis=1, keepdims=True)
    n2 = np.cross(t, n1)
    return j, t, n1, n2


def nasa_knockdown(r_over_t):
    phi = np.sqrt(np.maximum(r_over_t, 1e-9)) / 16.0
    return 1.0 - 0.901 * (1.0 - np.exp(-phi))


def strut_section(xyz, w, sp, st, smooth=3.0):
    """Length-averaged minimum second moment of area of one strut from its cells.

    Cells are assigned to their nearest skeleton station; station centroids are averaged over +-`smooth` mm
    along the strut to give a smooth centroid line (a single station's Voronoi cell can be lopsided);
    offsets r from that line, perpendicular to the local tangent, give the section tensor
    (1/L) sum w r r^T, whose middle eigenvalue is I_min. L is the length the cells actually cover
    (sum of station spacings), so I and the mean area A = (sum w) / L do not depend on how far the
    skeleton path runs into neighbouring regions. Returns (I_min, A, L, station index per cell).
    """
    kt = cKDTree(sp)
    _, jj = kt.query(xyz)
    ns = len(sp)
    W = np.bincount(jj, w, ns)
    C = np.stack([np.bincount(jj, w * xyz[:, k], ns) for k in range(3)], 1)
    has = W > 0
    cbar = np.array(sp, float)
    for j in np.where(has)[0]:
        nb = [k for k in kt.query_ball_point(sp[j], smooth) if has[k]]
        cbar[j] = C[nb].sum(0) / W[nb].sum()
    tc = st[jj]
    dv = xyz - cbar[jj]
    r = dv - np.einsum("ni,ni->n", dv, tc)[:, None] * tc
    d, _ = kt.query(sp, k=min(3, ns))
    dlt = np.clip(0.5 * (d[:, 1] + d[:, -1]), 0.5, None) if ns > 1 else np.ones(1)
    L = float(dlt[has].sum())
    M = np.einsum("n,ni,nj->ij", w, r, r) / L
    return float(np.sort(np.linalg.eigvalsh(M))[1]), float(w.sum() / L), L, jj


def evaluate(mesh, sol, cell_region, seg, cfg, F_vec, fixed_xyz, load_xyz, design):
    """All metrics of one design. `design` carries: hollow (set of region ids), s (outer offset)."""
    mat, cap, man = cfg["material"], cfg["capacity"], cfg["manufacturing"]
    E, nu = mat["E_MPa"], mat["nu"]
    G = E / (2 * (1 + nu))
    eps, s_mat, s_hom = fe.cell_stresses(mesh, sol["u"], E, nu)
    vm = fe.von_mises(s_mat)
    xyz = mesh.cell_centres
    w = mesh.vf * mesh.h ** 3
    R = seg.by_id()
    kind = np.array([R[int(i)]["kind"] if int(i) in R else "mount" for i in cell_region])
    hollow = set(design.get("hollow", ()))
    s_off = float(design.get("s", 0.0) or 0.0)
    bias = man["wall_bias_mm"]
    amp = np.ones(len(vm))
    if hollow and s_off > bias:
        amp[np.isin(cell_region, list(hollow))] = s_off / (s_off - bias)
    # exclusion zone around boundary conditions
    near = np.zeros(len(vm), bool)
    if cap["exclude_bc_mm"] > 0:
        bc = np.vstack([fixed_xyz, load_xyz])
        d, _ = cKDTree(bc).query(xyz, distance_upper_bound=cap["exclude_bc_mm"] + 1e-9)
        near = np.isfinite(d)
    design_cells = (kind != "mount") | bool(cap["include_mounts"])
    use = design_cells & ~near
    vm_eff = vm * amp
    sigma_q = weighted_quantile(vm_eff[use], w[use], cap["stress_quantile"])
    lam_s = mat["strength_MPa"] / sigma_q if sigma_q > 0 else np.inf
    mount_cells = (kind == "mount") & ~near
    sigma_mount = weighted_quantile(vm[mount_cells], w[mount_cells], cap["stress_quantile"]) if mount_cells.any() else float("nan")

    # local frames
    _, t, n1, n2 = local_frames(xyz, seg)
    T = fe.tensor(s_mat)
    Tt = np.einsum("nij,nj->ni", T, t)
    s_tt = np.einsum("ni,ni->n", Tt, t)
    t1 = np.einsum("ni,ni->n", Tt, n1)
    t2 = np.einsum("ni,ni->n", Tt, n2)
    s_tt_hom = s_tt * mesh.vf
    # principal stresses (compressive demand)
    smin = np.linalg.eigvalsh(T)[:, 0]
    comp = np.maximum(-smin, 0.0) * amp

    per_region = []
    lam_E, lam_L = np.inf, np.inf
    gov_E, gov_L = None, None
    for r in seg.regions:
        rid = r["id"]
        c = cell_region == rid
        if not c.any():
            continue
        Ub = float(np.sum(s_tt[c] ** 2 / (2 * E) * w[c]))
        Us = float(np.sum((t1[c] ** 2 + t2[c] ** 2) / (2 * G) * w[c]))
        U = float(np.sum(fe.strain_energy_density(s_hom[c], eps[c]) * mesh.h ** 3))
        row = dict(id=rid, kind=r["kind"], U=U, U_bend=Ub, U_shear=Us, R=Us / Ub if Ub > 0 else np.nan,
                   vm_q=weighted_quantile(vm_eff[c & ~near], w[c & ~near], cap["stress_quantile"]) if (c & ~near).any() else np.nan)
        if r["kind"] == "strut" and r.get("L_mm"):
            own = seg.skel_region == rid
            sp = seg.skel_pts[own] if own.sum() >= 3 else seg.skel_pts
            st = seg.skel_tan[own] if own.sum() >= 3 else seg.skel_tan
            I_min, A_mean, L_ext, jj = strut_section(xyz[c], w[c], sp, st)
            tc = st[jj]                                     # the strut's own tangents (not a nearby junction's)
            stt_hom = np.einsum("ni,nij,nj->n", tc, T[c], tc) * mesh.vf[c]
            N = float(np.sum(stt_hom) * mesh.h ** 3 / L_ext)  # mean axial force (compression < 0)
            if rid in hollow and s_off > bias:
                I_min *= (s_off - bias) / s_off             # as-built thin wall: I ~ t
            L_m = L_ext + 2 * cfg["segmentation"]["joint_k"] * (r["r_mm"] or 0)   # free length + joint halves
            Pcr = np.pi ** 2 * E * I_min / (cap["K_eff"] * L_m) ** 2
            row.update(N=N, I_min=I_min, A=A_mean, L_cells=L_ext, L_member=L_m, P_cr=Pcr)
            if N < 0:
                lamE = Pcr / abs(N)
                row["lam_E"] = lamE
                if lamE < lam_E:
                    lam_E, gov_E = lamE, rid
        if cap["local_buckling"] and rid in hollow and s_off > 0:
            t_w = max(s_off - bias, 1e-3)
            if r["kind"] == "strut":
                r_m = (r["r_mm"] or 1.0) + s_off / 2
                scr = nasa_knockdown(r_m / t_w) * E * t_w / (r_m * np.sqrt(3 * (1 - nu ** 2)))
            else:
                b = 2.0 * (r["r_mm"] or 1.0) + s_off
                scr = 4 * np.pi ** 2 * E / (12 * (1 - nu ** 2)) * (t_w / b) ** 2
            cc = c & ~near
            demand = weighted_quantile(comp[cc], w[cc], cap["stress_quantile"]) if cc.any() else 0.0
            lamL = scr / demand if demand > 0 else np.inf
            row.update(sigma_cr_local=float(scr), sigma_c=float(demand), lam_L=float(lamL), t_asbuilt=t_w)
            if lamL < lam_L:
                lam_L, gov_L = lamL, rid
        per_region.append(row)

    Fmag = float(np.linalg.norm(F_vec))
    C = sol["compliance"]
    K = Fmag ** 2 / C if (C > 0 and Fmag > 0) else np.nan
    mode, lam, gov_region = min([("strength", lam_s, None), ("Euler buckling", lam_E, gov_E),
                                 ("local buckling", lam_L, gov_L)], key=lambda m: m[1])
    return dict(compliance=C, K=K, lam_s=lam_s, lam_E=lam_E, lam_L=lam_L, lam=lam, P=lam * Fmag,
                mode=mode, governing_region=gov_region, sigma_q=sigma_q, sigma_mount_q=sigma_mount,
                vm_max=float(vm[use].max()) if use.any() else np.nan, per_region=per_region,
                cell=dict(vm=vm.astype(np.float32), region=cell_region.astype(np.int32), xyz=xyz.astype(np.float32),
                          vf=mesh.vf.astype(np.float32), excluded=near))
