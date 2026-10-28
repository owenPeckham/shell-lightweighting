"""Design surface Gamma_h, surface Helmholtz filter, per-vertex depth/clearance/keep-out and passive rules.

THEORY §3.4 (design surface, filter, passive vertices) and §3.5 (local depth, clearance, keep-outs); §3.8 for the
passive thin rule. Units mm. Float64 throughout.

Contents
--------
- `MeshQuery`: persistent AABB tree on a triangle mesh (exact point-to-mesh distance, rays), threaded.
- `remesh_isotropic`: pymeshlab isotropic remesh to edge length ell, then curvature-adaptive edge splitting until the
  sampled two-sided Hausdorff distance to the input is <= tol (or the vertex budget is spent), then Delaunay flips so
  the cotangent weights are non-negative (discrete maximum principle of the filter, §3.4).
- `DesignSurface`: V, T, B-rep face id per vertex, lumped Voronoi mass M, PSD cotangent stiffness L, vertex areas A,
  area-weighted outward normals n.
- `HelmholtzFilter`: F = (M + r_f^2 L)^-1 M, its adjoint M (M + r_f^2 L)^-1 and the inverse M^-1 (M + r_f^2 L).
- `medial_depth`, `outward_clearance`, `keepout_distance`: §3.5.
- `mount_face_ids`, `mount_vertices`, `thin_passive`, `passive_vertices`: §3.4 passive rule.

Usage
-----
    V, T, TF = vshell.geometry.tessellate(shape, 0.02, 0.2)
    ds = DesignSurface.build(V, T, TF, ell=0.75, hausdorff_tol=0.02, tess_tol=0.02)   # ds.info["hausdorff"], ["hausdorff_ok"]
    F = ds.filter(3.0);  xt = F.apply(xi);  g = F.adjoint(dJ_dxt);  xi = F.invert(xt)
    r_ma, info = ds.depth(h_g=0.1, return_info=True);  r_out = ds.clearance()
    t_K = keepout_distance(ds.V, ds.projected_normals(), [KeepOut("sphere", center=[0, 0, 9], radius=3)], c_K=0.5)
    ids = mount_face_ids(analytic_ids, support_load_ids, ds.V0, ds.T0, ds.TF0)
    passive, parts = passive_vertices(ds, ids, a_hi=r_ma + 0.75 * eps, a_lo=1.5, v_min=3.0, dilate_mm=1.5)
Cantilever (6.6e4-vertex tessellation): ~35 s remesh (Gamma_h: 8.98e4 vertices), 0.2 s CHOLMOD factorisation, ~21 s r_MA.

Deviations from the letter of the spec (all recorded in the docstrings where they apply):
 * `medial_depth` detects the first t with D(t) < t - delta by *binary search on the h_g/2 grid* (e(t) = t - D(t) is
   monotone for exact distance, since |D'| <= 1) instead of a linear march: 12 queries per vertex instead of ~200.
   `check_march` compares against the literal march. delta defaults to max(1e-3, 1.5 * tess_tol, 0.05 h_g): a tessellation with
   chord deflection s makes e(t) as large as s, so an absolute 1e-3 would stop the search at t ~ delta R / s.
"""
from __future__ import annotations

import os
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from typing import Callable, Iterable, Sequence

import numpy as np
import scipy.sparse as sp
import scipy.sparse.linalg as spla

_NTHREADS = max(1, min(int(os.environ.get("CSHELL_THREADS", os.cpu_count() or 1)), 64))
ANALYTIC_TYPES = ("Plane", "Cylinder", "Cone", "Sphere", "Torus")   # = vshell.geometry.ANALYTIC


# ----------------------------------------------------------------------------------------------------------------
# mesh utilities
# ----------------------------------------------------------------------------------------------------------------
def _f64(a) -> np.ndarray:
    return np.ascontiguousarray(a, dtype=np.float64)


def _i64(a) -> np.ndarray:
    return np.ascontiguousarray(a, dtype=np.int64)


def signed_volume(V: np.ndarray, T: np.ndarray) -> float:
    """Signed volume (mm^3) of a closed triangle mesh; positive for outward-facing (counter-clockwise) triangles."""
    a, b, c = V[T[:, 0]], V[T[:, 1]], V[T[:, 2]]
    return float(np.einsum("ij,ij->i", a, np.cross(b, c)).sum() / 6.0)


def orient_outward(V: np.ndarray, T: np.ndarray) -> np.ndarray:
    """Return T with consistent orientation such that the signed volume is positive (outward normals)."""
    T = _i64(T)
    return T[:, ::-1].copy() if signed_volume(V, T) < 0 else T


def _chunks(n: int, nchunk: int):
    edges = np.linspace(0, n, nchunk + 1).astype(int)
    return [(edges[i], edges[i + 1]) for i in range(nchunk) if edges[i + 1] > edges[i]]


class MeshQuery:
    """Exact point-to-mesh distance and ray queries on a triangle mesh (igl AABB, chunked over threads).

    `closest(P)` -> (distance, triangle id, closest point). `ray_first(O, D)` -> (hit distance t or inf, triangle).
    The mesh arrays are kept by reference (V float64, T int64, no copy).
    """

    def __init__(self, V: np.ndarray, T: np.ndarray, nthreads: int | None = None):
        import igl
        self.V, self.T = _f64(V), _i64(T)
        self.tree = igl.AABB()
        self.tree.init(self.V, self.T)
        self.nthreads = _NTHREADS if nthreads is None else nthreads
        self.n_queries = 0

    def _map(self, fn: Callable, n: int):
        if n < 4096 or self.nthreads == 1:
            return [fn(0, n)]
        with ThreadPoolExecutor(self.nthreads) as ex:
            return list(ex.map(lambda ab: fn(*ab), _chunks(n, 4 * self.nthreads)))

    def closest(self, P: np.ndarray):
        P = _f64(P).reshape(-1, 3)
        self.n_queries += len(P)

        def fn(a, b):
            return self.tree.squared_distance(self.V, self.T, P[a:b])
        out = self._map(fn, len(P))
        d2 = np.concatenate([o[0] for o in out])
        return np.sqrt(np.maximum(d2, 0.0)), np.concatenate([o[1] for o in out]), np.concatenate([o[2] for o in out])

    def distance(self, P: np.ndarray) -> np.ndarray:
        return self.closest(P)[0]

    def ray_first(self, O: np.ndarray, D: np.ndarray):
        """First hit of rays O + t D (t > 0, |D| = 1) with the mesh: (t, tri); t = inf and tri = -1 on a miss."""
        O, D = _f64(O).reshape(-1, 3), _f64(D).reshape(-1, 3)

        def fn(a, b):
            return self.tree.intersect_ray_first(self.V, self.T, O[a:b], D[a:b])
        out = self._map(fn, len(O))
        tri = np.concatenate([o[0] for o in out])
        t = np.concatenate([o[1] for o in out]).astype(np.float64)
        hit = (tri >= 0) & np.isfinite(t)
        return np.where(hit, t, np.inf), np.where(hit, tri, -1)


# ----------------------------------------------------------------------------------------------------------------
# operators on the design surface
# ----------------------------------------------------------------------------------------------------------------
def mesh_operators(V: np.ndarray, T: np.ndarray):
    """Lumped Voronoi mass M (diagonal csr), PSD cotangent stiffness L = -igl.cotmatrix, areas A, outward normals n.

    M_ii = Voronoi (mixed) area of vertex i, L_ij = -(cot a + cot b)/2 for edges, L 1 = 0 (§3.4). n_i is the
    area-weighted mean of the incident face normals. T must be outward oriented.
    """
    import igl
    V, T = _f64(V), _i64(T)
    M = sp.csr_matrix(igl.massmatrix(V, T, igl.MASSMATRIX_TYPE_VORONOI))
    L = sp.csr_matrix(-igl.cotmatrix(V, T))
    A = np.asarray(M.diagonal()).ravel().copy()
    N = igl.per_vertex_normals(V, T, igl.PER_VERTEX_NORMALS_WEIGHTING_TYPE_AREA)
    N = N / np.maximum(np.linalg.norm(N, axis=1, keepdims=True), 1e-300)
    return M, L, A, N


def sampled_hausdorff(V: np.ndarray, T: np.ndarray, qref: MeshQuery, tol: float | None = None) -> dict:
    """Two-sided Hausdorff distance between (V, T) and the mesh behind `qref`, from dense samples.

    Samples on (V, T): vertices, edge midpoints, face centroids -> distance to the reference mesh (`d_out_in`).
    Samples on the reference: vertices and face centroids -> distance to (V, T) (`d_in_out`). The maximum is a lower
    bound of the true Hausdorff distance, tight to O(ell^2 / R) at the sample spacing used here.
    Returns max/99.9 percentile/fraction above `tol`, and the per-face deviation of (V, T) (for refinement).
    """
    V, T = _f64(V), _i64(T)
    nf = len(T)
    E = np.concatenate([V[T[:, [0, 1]]].mean(1), V[T[:, [1, 2]]].mean(1), V[T[:, [2, 0]]].mean(1), V[T].mean(1)])
    d_oi = qref.distance(E).reshape(4, nf)
    d_v = qref.distance(V)
    q_self = MeshQuery(V, T, qref.nthreads)
    P_in = np.concatenate([qref.V, qref.V[qref.T].mean(1)])
    d_io, tri_io, _ = q_self.closest(P_in)
    face_dev = d_oi.max(0)
    np.maximum.at(face_dev, tri_io, d_io)
    out = dict(
        d_out_in=float(max(d_oi.max(), d_v.max())),
        d_in_out=float(d_io.max()),
        p999=float(np.percentile(np.concatenate([d_oi.ravel(), d_io]), 99.9)),
        face_dev=face_dev,
    )
    out["hausdorff"] = max(out["d_out_in"], out["d_in_out"])
    if tol is not None:
        out["frac_above_tol"] = float((np.concatenate([d_oi.ravel(), d_io]) > tol).mean())
    return out


# ---- edge splitting (curvature-adaptive refinement) and Delaunay flips (vectorised) --------------------------------
def _edge_table(T: np.ndarray, nv: int):
    """Unique edges. Returns (E (ne,2) sorted, face-edge id (nf,3) for edges (0-1, 1-2, 2-0))."""
    e = np.stack([T[:, [0, 1]], T[:, [1, 2]], T[:, [2, 0]]], 1)          # (nf, 3, 2)
    key = e.min(2).astype(np.int64) * nv + e.max(2)
    uk, inv = np.unique(key.ravel(), return_inverse=True)
    E = np.stack([uk // nv, uk % nv], 1)
    return E, inv.reshape(-1, 3)


def _split_edges(V: np.ndarray, T: np.ndarray, mark_edge: np.ndarray, E: np.ndarray, fe: np.ndarray):
    """Split marked edges at their midpoints with the standard 1/2/3-mark templates (conforming). Returns (V', T')."""
    nv = len(V)
    new_id = -np.ones(len(E), np.int64)
    sel = np.where(mark_edge)[0]
    new_id[sel] = nv + np.arange(len(sel))
    Vn = np.concatenate([V, 0.5 * (V[E[sel, 0]] + V[E[sel, 1]])])
    mid = new_id[fe]                                                     # (nf, 3): midpoint id of edge k or -1
    nm = (mid >= 0).sum(1)
    out = [T[nm == 0]]
    # one mark: roll so the marked edge is edge 0 (v0-v1)
    for k in range(3):
        f = np.where((nm == 1) & (mid[:, k] >= 0))[0]
        if len(f):
            v = np.roll(T[f], -k, axis=1)
            m = mid[f, k]
            out += [np.stack([v[:, 0], m, v[:, 2]], 1), np.stack([m, v[:, 1], v[:, 2]], 1)]
    # two marks: roll so the unmarked edge is edge 2 (v2-v0): marked edges 0 (m01) and 1 (m12)
    for k in range(3):                                                    # k = unmarked edge
        f = np.where((nm == 2) & (mid[:, k] < 0))[0]
        if len(f):
            r = (k + 1) % 3
            v = np.roll(T[f], -r, axis=1)
            m01, m12 = mid[f, r], mid[f, (r + 1) % 3]
            v0, v1, v2 = v[:, 0], v[:, 1], v[:, 2]
            d_a = np.linalg.norm(Vn[v0] - Vn[m12], axis=1)               # diagonal v0-m12
            d_b = np.linalg.norm(Vn[m01] - Vn[v2], axis=1)               # diagonal m01-v2
            a = d_a <= d_b
            out += [np.stack([m01, v1, m12], 1)]
            out += [np.where(a[:, None], np.stack([v0, m01, m12], 1), np.stack([v0, m01, v2], 1))]
            out += [np.where(a[:, None], np.stack([v0, m12, v2], 1), np.stack([m01, m12, v2], 1))]
    f = np.where(nm == 3)[0]
    if len(f):
        v = T[f]
        m01, m12, m20 = mid[f, 0], mid[f, 1], mid[f, 2]
        out += [np.stack([v[:, 0], m01, m20], 1), np.stack([m01, v[:, 1], m12], 1),
                np.stack([m20, m12, v[:, 2]], 1), np.stack([m01, m12, m20], 1)]
    return Vn, np.concatenate(out), nv


def _cot_flip_pass(V: np.ndarray, T: np.ndarray, max_dihedral_deg: float, protect: np.ndarray | None = None,
                   qref: "MeshQuery | None" = None, dev_max: float = 0.0):
    """One batch of Delaunay edge flips (edges with cot a + cot b < 0), independent set, normal-continuity guarded."""
    nv = len(V)
    E, fe = _edge_table(T, nv)
    ne = len(E)
    # adjacent faces per edge
    fidx = np.repeat(np.arange(len(T)), 3)
    order = np.argsort(fe.ravel(), kind="stable")
    se = fe.ravel()[order]
    first = np.searchsorted(se, np.arange(ne))
    cnt = np.bincount(se, minlength=ne)
    inner = np.where(cnt == 2)[0]
    f1, f2 = fidx[order[first[inner]]], fidx[order[first[inner] + 1]]
    a, b = E[inner, 0], E[inner, 1]
    # opposite vertices
    def opp(f):
        t = T[f]
        return np.where((t[:, 0] != a) & (t[:, 0] != b), t[:, 0], np.where((t[:, 1] != a) & (t[:, 1] != b), t[:, 1], t[:, 2]))
    c, d = opp(f1), opp(f2)
    def cot(p, q, r):                                                     # angle at r between p-r and q-r
        u, w = V[p] - V[r], V[q] - V[r]
        return np.einsum("ij,ij->i", u, w) / np.maximum(np.linalg.norm(np.cross(u, w), axis=1), 1e-300)
    w = cot(a, b, c) + cot(a, b, d)
    cand = w < -1e-12
    # flipped faces keep the winding of the old ones; f1 = (a,b,c) -> (a,d,c),(d,b,c); f1 = (b,a,c) -> (b,d,c),(d,a,c)
    def oriented(f, p, q):                                                # True if face f contains p->q in its winding
        t = T[f]
        return ((t[:, 0] == p) & (t[:, 1] == q)) | ((t[:, 1] == p) & (t[:, 2] == q)) | ((t[:, 2] == p) & (t[:, 0] == q))
    ab1 = oriented(f1, a, b)
    A1 = np.where(ab1[:, None], np.stack([a, d, c], 1), np.stack([b, d, c], 1))
    B1 = np.where(ab1[:, None], np.stack([d, b, c], 1), np.stack([d, a, c], 1))
    nA = np.cross(V[A1[:, 1]] - V[A1[:, 0]], V[A1[:, 2]] - V[A1[:, 0]])
    nB = np.cross(V[B1[:, 1]] - V[B1[:, 0]], V[B1[:, 2]] - V[B1[:, 0]])
    no1 = np.cross(V[T[f1, 1]] - V[T[f1, 0]], V[T[f1, 2]] - V[T[f1, 0]])
    no2 = np.cross(V[T[f2, 1]] - V[T[f2, 0]], V[T[f2, 2]] - V[T[f2, 0]])
    cosmin = np.cos(np.radians(max_dihedral_deg))
    def cosang(x, y):
        return np.einsum("ij,ij->i", x, y) / np.maximum(np.linalg.norm(x, axis=1) * np.linalg.norm(y, axis=1), 1e-300)
    ok = cand & (cosang(nA, no1) > cosmin) & (cosang(nB, no1) > cosmin) & (cosang(nA, no2) > cosmin) & (cosang(nB, no2) > cosmin)
    ok &= cosang(no1, no2) > cosmin                                       # do not flip across creases
    ok &= (np.linalg.norm(nA, axis=1) > 1e-12) & (np.linalg.norm(nB, axis=1) > 1e-12)
    if protect is not None:
        ok &= ~protect[inner]
    if qref is not None:                                                  # new diagonal and faces must stay on the input
        ic = np.where(ok)[0]
        if len(ic):
            pts = np.concatenate([0.5 * (V[c[ic]] + V[d[ic]]), V[A1[ic]].mean(1), V[B1[ic]].mean(1)])
            dv = qref.distance(pts).reshape(3, len(ic)).max(0)
            ok[ic[dv > dev_max]] = False
    # no duplicate edge: c-d must not exist already
    key_all = E[:, 0] * nv + E[:, 1]
    key_cd = np.minimum(c, d).astype(np.int64) * nv + np.maximum(c, d)
    ok &= ~np.isin(key_cd, key_all)
    ok &= c != d
    sel = np.where(ok)[0]
    if len(sel) == 0:
        return T, 0
    # independent set: each face claimed by the edge with the most negative weight
    prio = np.argsort(w[sel])
    sel = sel[prio]
    rank = np.arange(len(sel))
    best = np.full(len(T), np.iinfo(np.int64).max)
    np.minimum.at(best, f1[sel], rank)
    np.minimum.at(best, f2[sel], rank)
    keep = (best[f1[sel]] == rank) & (best[f2[sel]] == rank)
    sel = sel[keep]
    # also forbid two flips creating the same diagonal key
    _, uniq_idx = np.unique(key_cd[sel], return_index=True)
    sel = sel[np.sort(uniq_idx)]
    Tn = T.copy()
    Tn[f1[sel]] = A1[sel]
    Tn[f2[sel]] = B1[sel]
    return Tn, len(sel)


def delaunay_flips(V: np.ndarray, T: np.ndarray, max_passes: int = 40, max_dihedral_deg: float = 10.0,
                   qref: "MeshQuery | None" = None, dev_max: float = 0.0) -> tuple[np.ndarray, int]:
    """Flip edges with cot a + cot b < 0 (non-Delaunay) while the new faces stay within `max_dihedral_deg` of the old
    ones and (if `qref` is given) the new diagonal midpoint and face centroids stay within `dev_max` of the reference
    mesh. Makes the cotangent weights non-negative wherever geometry allows (discrete maximum principle, §3.4)."""
    total = 0
    for _ in range(max_passes):
        T, n = _cot_flip_pass(V, T, max_dihedral_deg, None, qref, dev_max)
        total += n
        if n == 0:
            break
    return T, total


def remesh_isotropic(V: np.ndarray, T: np.ndarray, ell: float = 0.75, tol: float = 0.02, adaptive: bool = True,
                     max_growth: float = 1.7, iterations: int = 10, feature_deg: float = 30.0,
                     max_rounds: int = 14, flips: bool = True, min_edge_frac: float = 0.08, verbose: bool = False):
    """Isotropic remesh of (V, T) to edge length `ell` (§3.4) with Hausdorff control against the input.

    1. pymeshlab explicit isotropic remeshing (split/collapse/flip/relax with reprojection onto the input,
       dihedral feature angle `feature_deg` preserved). gpytoolbox has no aarch64 wheel, pymeshlab does.
    2. If `adaptive`: while the sampled Hausdorff distance exceeds `tol` and the vertex count is below
       `max_growth` x the uniform count, split every edge whose midpoint (or incident-face centroid) lies more than
       0.8 tol from the input and put the new vertex on the input surface (closest point).
    3. Delaunay flips (guarded by `max_dihedral_deg`).
    Returns (V', T', info). info: n_uniform, n_final, hausdorff (dict of `sampled_hausdorff`), rounds, seconds.
    The uniform edge length cannot satisfy tol on tight fillets (chord sag ell^2/8R): that is why step 2 exists.
    """
    import pymeshlab as pm
    t0 = time.time()
    V, T = _f64(V), orient_outward(V, T)
    qref = MeshQuery(V, T)
    ms = pm.MeshSet()
    ms.add_mesh(pm.Mesh(V, T))
    ms.meshing_isotropic_explicit_remeshing(targetlen=pm.PureValue(float(ell)), iterations=int(iterations),
                                            featuredeg=float(feature_deg), checksurfdist=False)
    m = ms.current_mesh()
    V1, T1 = _f64(m.vertex_matrix()), _i64(m.face_matrix())
    T1 = orient_outward(V1, T1)
    n_uniform = len(V1)
    t_pml = time.time() - t0
    hist = []
    for rnd in range(max_rounds + 1):
        hd = sampled_hausdorff(V1, T1, qref, tol)
        hist.append((len(V1), hd["hausdorff"]))
        if verbose:
            print(f"  remesh round {rnd}: nv={len(V1)} hausdorff={hd['hausdorff']:.4f} (out->in {hd['d_out_in']:.4f}, "
                  f"in->out {hd['d_in_out']:.4f}) frac>tol={hd.get('frac_above_tol', 0):.4f}", flush=True)
        if not adaptive or hd["hausdorff"] <= tol or rnd == max_rounds or len(V1) > max_growth * n_uniform:
            break
        E, fe = _edge_table(T1, len(V1))
        mid = 0.5 * (V1[E[:, 0]] + V1[E[:, 1]])
        d_mid = qref.distance(mid)
        fbad = hd["face_dev"] > 0.8 * tol
        mark = d_mid > 0.8 * tol
        # faces that are bad through their centroid / in->out samples: split their longest edge
        fl = np.where(fbad)[0]
        elen = np.linalg.norm(V1[E[:, 0]] - V1[E[:, 1]], axis=1)
        le = fe[fl][np.arange(len(fl)), np.argmax(elen[fe[fl]], axis=1)]
        mark[le] = True
        mark &= elen > min_edge_frac * ell                                # never split below min_edge_frac * ell
        if not mark.any():
            break
        Vn, Tn, nv0 = _split_edges(V1, T1, mark, E, fe)
        _, _, Cn = qref.closest(Vn[nv0:])
        Vn[nv0:] = Cn
        V1, T1 = Vn, orient_outward(Vn, Tn)
    n_flips = 0
    if flips:
        T1, n_flips = delaunay_flips(V1, T1, qref=qref, dev_max=0.6 * tol)
        hd = sampled_hausdorff(V1, T1, qref, tol)
    info = dict(n_uniform=n_uniform, n_final=len(V1), hausdorff=hd, history=hist, flips=n_flips,
                seconds=time.time() - t0, seconds_pymeshlab=t_pml)
    return V1, T1, info


# ----------------------------------------------------------------------------------------------------------------
# DesignSurface
# ----------------------------------------------------------------------------------------------------------------
@dataclass
class DesignSurface:
    """Gamma_h (§3.4): mesh, B-rep face id per vertex, and the operators the filter needs.

    V (nv,3), T (nt,3) outward oriented; face_id (nv,) B-rep face of the nearest input triangle (-1 if unknown);
    M lumped Voronoi mass (diag csr), L PSD cotangent stiffness (csr), A vertex areas (M diagonal), N area-weighted
    outward vertex normals. `V0, T0, TF0` is the input tessellation Gamma_0 (reference for all exact queries).
    """
    V: np.ndarray
    T: np.ndarray
    face_id: np.ndarray
    M: sp.csr_matrix
    L: sp.csr_matrix
    A: np.ndarray
    N: np.ndarray
    V0: np.ndarray | None = None
    T0: np.ndarray | None = None
    TF0: np.ndarray | None = None
    ell: float = 0.75
    tess_tol: float = 0.02
    info: dict = field(default_factory=dict)
    _q0: MeshQuery | None = field(default=None, repr=False)

    @property
    def n_vertices(self) -> int:
        return len(self.V)

    @property
    def query0(self) -> MeshQuery:
        """Persistent exact-distance structure on the input tessellation Gamma_0."""
        if self._q0 is None:
            if self.V0 is None:
                raise ValueError("DesignSurface has no reference tessellation (V0, T0)")
            self._q0 = MeshQuery(self.V0, self.T0)
        return self._q0

    @classmethod
    def from_mesh(cls, V: np.ndarray, T: np.ndarray, TF: np.ndarray | None = None, ell: float = 0.75,
                  tess_tol: float = 0.02, ref: tuple | None = None) -> "DesignSurface":
        """Wrap an existing mesh (no remeshing). `ref=(V0, T0, TF0)` is the reference tessellation (default: itself)."""
        V, T = _f64(V), orient_outward(V, T)
        V0, T0, TF0 = (V, T, TF) if ref is None else (_f64(ref[0]), orient_outward(ref[0], ref[1]), ref[2])
        M, L, A, N = mesh_operators(V, T)
        ds = cls(V=V, T=T, face_id=-np.ones(len(V), np.int64), M=M, L=L, A=A, N=N, V0=V0, T0=T0,
                 TF0=None if TF0 is None else np.asarray(TF0), ell=ell, tess_tol=tess_tol)
        if TF0 is not None:
            ds.assign_face_ids()
        return ds

    @classmethod
    def build(cls, V: np.ndarray, T: np.ndarray, TF: np.ndarray | None = None, ell: float = 0.75,
              hausdorff_tol: float = 0.02, tess_tol: float = 0.02, adaptive: bool = True, max_growth: float = 1.7,
              verbose: bool = False, **kw) -> "DesignSurface":
        """Isotropic remesh of the tessellation (V, T, TF) to edge ell, Hausdorff check, face ids, operators (§3.4).

        TF (nt,) B-rep face id per input triangle (from `vshell.geometry.tessellate`). The achieved Hausdorff
        distance is in `info["hausdorff"]` (and `info["hausdorff_ok"]`); it is reported, not asserted, because a
        tight fillet can make the bound unreachable within the vertex budget `max_growth`.
        """
        V = _f64(V)
        T = orient_outward(V, T)
        V1, T1, rinfo = remesh_isotropic(V, T, ell=ell, tol=hausdorff_tol, adaptive=adaptive, max_growth=max_growth,
                                         verbose=verbose, **kw)
        M, L, A, N = mesh_operators(V1, T1)
        ds = cls(V=V1, T=T1, face_id=-np.ones(len(V1), np.int64), M=M, L=L, A=A, N=N, V0=V, T0=T,
                 TF0=None if TF is None else np.asarray(TF), ell=ell, tess_tol=tess_tol)
        if TF is not None:
            ds.assign_face_ids()
        ds.info.update(rinfo)
        ds.info["hausdorff_ok"] = bool(rinfo["hausdorff"]["hausdorff"] <= hausdorff_tol)
        ds.info["hausdorff_tol"] = hausdorff_tol
        w = -ds.L.tocoo()
        off = w.row != w.col
        ds.info["n_negative_cot_edges"] = int(((w.data < -1e-12) & off).sum() // 2)
        ds.info["edge_length"] = edge_length_stats(V1, T1)
        return ds

    def assign_face_ids(self) -> np.ndarray:
        """B-rep face id of the nearest input triangle for each vertex (§3.4)."""
        _, tri, _ = self.query0.closest(self.V)
        self.face_id = np.asarray(self.TF0)[tri].astype(np.int64)
        return self.face_id

    def projected_normals(self, feature_deg: float = 30.0) -> np.ndarray:
        """Unit outward normals taken from the input tessellation at the closest point: barycentric interpolation of
        igl per-corner normals (face normals averaged over neighbours within `feature_deg`, so flat faces stay flat
        and creases stay sharp). Smoother and more accurate than `N` (the design mesh's own area-weighted normals,
        the spec's default), which matters for ray marching: a tilt beta in the inward direction makes the marched
        e(t) grow like t^2 beta^2 / 2(R - t) (see `medial_depth`)."""
        import igl
        q = self.query0
        _, tri, cp = q.closest(self.V)
        CN = igl.per_corner_normals(q.V, q.T, float(feature_deg)).reshape(len(q.T), 3, 3)
        a, b, c = q.V[q.T[tri, 0]], q.V[q.T[tri, 1]], q.V[q.T[tri, 2]]
        v0, v1, v2 = b - a, c - a, cp - a
        d00, d01, d11 = (v0 * v0).sum(1), (v0 * v1).sum(1), (v1 * v1).sum(1)
        d20, d21 = (v2 * v0).sum(1), (v2 * v1).sum(1)
        den = np.maximum(d00 * d11 - d01 * d01, 1e-300)
        bv, bw = (d11 * d20 - d01 * d21) / den, (d00 * d21 - d01 * d20) / den
        bu = 1 - bv - bw
        n = bu[:, None] * CN[tri, 0] + bv[:, None] * CN[tri, 1] + bw[:, None] * CN[tri, 2]
        return n / np.maximum(np.linalg.norm(n, axis=1, keepdims=True), 1e-300)

    def filter(self, r_f: float) -> "HelmholtzFilter":
        return HelmholtzFilter(self.M, self.L, r_f)

    def depth(self, h_g: float, **kw):
        """`medial_depth` for the vertices of this surface (inward marching along `projected_normals`)."""
        return medial_depth(self.V, self.projected_normals(), self.query0, h_g=h_g, tess_tol=self.tess_tol, **kw)

    def clearance(self, **kw) -> np.ndarray:
        return outward_clearance(self.V, self.projected_normals(), self.query0, tess_tol=self.tess_tol, **kw)


def edge_length_stats(V: np.ndarray, T: np.ndarray) -> dict:
    e = np.concatenate([V[T[:, 1]] - V[T[:, 0]], V[T[:, 2]] - V[T[:, 1]], V[T[:, 0]] - V[T[:, 2]]])
    l = np.linalg.norm(e, axis=1)
    return dict(mean=float(l.mean()), min=float(l.min()), max=float(l.max()), p05=float(np.percentile(l, 5)),
                p95=float(np.percentile(l, 95)))


# ----------------------------------------------------------------------------------------------------------------
# Helmholtz filter
# ----------------------------------------------------------------------------------------------------------------
class HelmholtzFilter:
    """Surface PDE filter (Lazarov & Sigmund 2011; vertex morphing), THEORY §3.4.

    (M + r_f^2 L) x~ = M x, so F = (M + r_f^2 L)^-1 M. One sparse LU of the SPD matrix A_f = M + r_f^2 L
    (CHOLMOD via scikit-sparse; scipy `splu` fallback); the factors serve all of `apply`, `adjoint`, `invert`.
      apply(x)   = F x                 = A_f^-1 (M x)
      adjoint(g) = F^T g               = M A_f^-1 g           (A_f symmetric)
      invert(xt) = F^-1 xt             = M^-1 (A_f xt)        (§3.5 step 2, D_comp re-mapping)
    Accepts vectors (n,) or column blocks (n, k). M x is the lumped mass-weighted average; with non-negative
    cotangent weights F is a stochastic matrix and the maximum principle holds: min x <= F x <= max x.
    """

    def __init__(self, M: sp.spmatrix, L: sp.spmatrix, r_f: float):
        self.r_f = float(r_f)
        self.m = np.asarray(M.diagonal()).ravel().astype(np.float64)
        if not np.all(self.m > 0):
            raise ValueError("mass matrix has non-positive entries")
        self.M = sp.diags(self.m, format="csc")
        self.L = sp.csc_matrix(L)
        self.A = (self.M + self.r_f ** 2 * self.L).tocsc()
        try:                                                          # CHOLMOD (scikit-sparse): ~1 s at 9e4 vertices
            from sksparse.cholmod import cho_factor
            self._chol = cho_factor(self.A)
            self._solve = self._chol.solve
            self.backend = "cholmod"
        except ImportError:                                           # SuperLU: ~3 min at 9e4 vertices
            self._lu = spla.splu(self.A, permc_spec="COLAMD")
            self._solve = self._lu.solve
            self.backend = "superlu"

    def _mul(self, v: np.ndarray, x: np.ndarray) -> np.ndarray:
        return v * x if x.ndim == 1 else v[:, None] * x

    def apply(self, x: np.ndarray) -> np.ndarray:
        x = np.asarray(x, dtype=np.float64)
        return self._solve(self._mul(self.m, x))

    def adjoint(self, g: np.ndarray) -> np.ndarray:
        g = np.asarray(g, dtype=np.float64)
        return self._mul(self.m, self._solve(g))

    def invert(self, x_tilde: np.ndarray) -> np.ndarray:
        x_tilde = np.asarray(x_tilde, dtype=np.float64)
        return self._mul(1.0 / self.m, self.A @ x_tilde)


# ----------------------------------------------------------------------------------------------------------------
# §3.5 local depth, clearance, keep-outs
# ----------------------------------------------------------------------------------------------------------------
def _as_query(mesh) -> MeshQuery:
    return mesh if isinstance(mesh, MeshQuery) else MeshQuery(*mesh)


def medial_depth(P: np.ndarray, N: np.ndarray, mesh, h_g: float = 0.1, delta: float | None = None,
                 tess_tol: float = 0.0, tol: float = 0.01, bias_correct: bool = True, t_max: float | None = None,
                 check_march: int = 0, return_info: bool = False):
    """Local inradius r_MA at points P with inward direction -N (THEORY §3.5).

    D(t) = dist(P - t N, mesh) (exact point-to-mesh distance, no grid interpolation) equals t until the medial axis,
    then falls below it. e(t) = t - D(t) is non-decreasing (|D'| <= 1). The first t with e(t) > delta is bracketed on
    the h_g/2 grid (binary search, §3.5 step 1) and then bisected to `tol` = 0.01 mm (step 2).

    delta: default max(1e-3, 1.5 tess_tol, 0.05 h_g). On a tessellation with chord deflection s the mesh distance differs from
    the true one by up to s, so e(t) reaches s on a smooth axis-free surface; an absolute 1e-3 would stop at
    t ~ 1e-3 R / s. The 0.05 h_g floor limits the sensitivity to normal tilt (r shifts by ~ R beta^2 / 2 delta; beta =
    1e-3 rad, R = 5 gives 0.0025 mm at delta = 1e-3... and 5e-4 at 0.005). `bias_correct`: for a planar medial axis e(t) = 2 (t - r) so the crossing is at r + delta/2;
    subtract delta/2. N must be unit and accurate (a tilt beta makes e(t) ~ t^2 beta^2 / 2(R - t)): use
    `DesignSurface.projected_normals`.
    The ray-chord upper bound (half the chord to the opposite surface) caps the search and is returned as
    `info["chord_half"]`; `info["n_violations"]` counts r_MA > chord/2 + delta (should be 0).
    `check_march` = k > 0 compares k random points against the literal linear march and reports mismatches > h_g/2.
    """
    q = _as_query(mesh)
    P, N = _f64(P).reshape(-1, 3), _f64(N).reshape(-1, 3)
    n = len(P)
    step = 0.5 * h_g
    if delta is None:
        delta = max(1e-3, 1.5 * tess_tol, 0.05 * h_g)
    eta = max(1e-4, 0.25 * delta)
    # chord upper bound: ray from p - eta n along -n
    tc, _ = q.ray_first(P - eta * N, -N)
    chord = tc + eta
    bbox = np.linalg.norm(q.V.max(0) - q.V.min(0))
    cap = np.where(np.isfinite(chord), 0.5 * (chord + delta), bbox)
    if t_max is not None:
        cap = np.minimum(cap, t_max)

    def e_of(idx: np.ndarray, t: np.ndarray) -> np.ndarray:
        return t - q.distance(P[idx] - t[:, None] * N[idx])

    idx_all = np.arange(n)
    kcap = np.maximum(np.ceil(cap / step).astype(np.int64), 1)
    lo = np.zeros(n, np.int64)                       # predicate e(t_k) > delta is False at lo (t_0 = 0: e = 0)
    hi = kcap.copy()                                 # ... and True at hi (guaranteed by the chord bound)
    # verify the cap really is True; if not, the point is not bracketed (open mesh or noise): return cap
    ok_hi = e_of(idx_all, hi * step) > delta
    active = ok_hi & (hi - lo > 1)
    while active.any():
        ia = np.where(active)[0]
        mid = (lo[ia] + hi[ia]) // 2
        pred = e_of(ia, mid * step) > delta
        hi[ia] = np.where(pred, mid, hi[ia])
        lo[ia] = np.where(pred, lo[ia], mid)
        active = (hi - lo > 1) & ok_hi
    # bisect inside [lo step, hi step] to `tol`
    a, b = lo * step, np.minimum(hi * step, np.where(ok_hi, hi * step, 0.0))
    nb = int(np.ceil(np.log2(max(2.0 * step / tol, 1.0))))        # final bracket <= tol / 2
    for _ in range(nb):
        mid = 0.5 * (a + b)
        pred = e_of(idx_all, mid) > delta
        b = np.where(pred, mid, b)
        a = np.where(pred, a, mid)
    t_star = 0.5 * (a + b)
    r = t_star - (0.5 * delta if bias_correct else 0.0)
    r = np.where(ok_hi, r, cap)
    r = np.maximum(r, 0.0)
    if not return_info:
        return r
    info = dict(delta=delta, step=step, chord_half=0.5 * chord, n_unbracketed=int((~ok_hi).sum()),
                n_violations=int((np.isfinite(chord) & (r > 0.5 * chord + delta + tol)).sum()), n_queries=q.n_queries)
    if check_march:
        rng = np.random.default_rng(0)
        sub = rng.choice(n, size=min(check_march, n), replace=False)
        r_m = np.empty(len(sub))
        for j, i in enumerate(sub):
            k = np.arange(1, int(kcap[i]) + 1)
            e = (k * step) - q.distance(P[i] - (k * step)[:, None] * N[i])
            hit = np.where(e > delta)[0]
            r_m[j] = (k[hit[0]] * step if len(hit) else cap[i])
        info["march_mismatch_cells"] = int((np.abs(r_m - (t_star[sub] + 0.0)) > step).sum())
        info["march_checked"] = len(sub)
    return r, info


def outward_clearance(P: np.ndarray, N: np.ndarray, mesh, tess_tol: float = 0.0, eta: float | None = None) -> np.ndarray:
    """r_out = g / 2 with g the distance from P along +N to the first hit on the mesh; +inf when there is none (§3.5).

    The ray starts at P + eta N, eta = max(1e-3, 2.5 tess_tol), to step over the facets adjacent to P; g includes eta.
    """
    q = _as_query(mesh)
    P, N = _f64(P).reshape(-1, 3), _f64(N).reshape(-1, 3)
    if eta is None:
        eta = max(1e-3, 2.5 * tess_tol)
    t, _ = q.ray_first(P + eta * N, N)
    return 0.5 * (t + eta)


# ---- keep-out signed distances ---------------------------------------------------------------------------------
@dataclass
class KeepOut:
    """Keep-out body with an exact signed distance phi_K (negative inside). kind: 'sphere' (center, radius),
    'box' (lo, hi; axis aligned), 'cylinder' (p0, p1 end centres, radius; finite, capped), 'mesh' (V, T closed)."""
    kind: str
    center: Sequence[float] | None = None
    radius: float | None = None
    lo: Sequence[float] | None = None
    hi: Sequence[float] | None = None
    p0: Sequence[float] | None = None
    p1: Sequence[float] | None = None
    V: np.ndarray | None = None
    T: np.ndarray | None = None

    def phi(self, X: np.ndarray) -> np.ndarray:
        X = _f64(X).reshape(-1, 3)
        if self.kind == "sphere":
            return np.linalg.norm(X - np.asarray(self.center, float), axis=1) - float(self.radius)
        if self.kind == "box":
            c, h = 0.5 * (np.asarray(self.lo, float) + np.asarray(self.hi, float)), 0.5 * (np.asarray(self.hi, float) - np.asarray(self.lo, float))
            q = np.abs(X - c) - h
            return np.linalg.norm(np.maximum(q, 0.0), axis=1) + np.minimum(q.max(1), 0.0)
        if self.kind == "cylinder":
            p0, p1 = np.asarray(self.p0, float), np.asarray(self.p1, float)
            ax = p1 - p0
            L = np.linalg.norm(ax)
            ax /= L
            rel = X - 0.5 * (p0 + p1)
            s = rel @ ax
            rad = np.linalg.norm(rel - s[:, None] * ax, axis=1)
            q = np.stack([rad - float(self.radius), np.abs(s) - 0.5 * L], 1)
            return np.linalg.norm(np.maximum(q, 0.0), axis=1) + np.minimum(q.max(1), 0.0)
        if self.kind == "mesh":
            import igl
            return igl.signed_distance(X, _f64(self.V), _i64(self.T))[0]
        raise ValueError(f"unknown keep-out kind {self.kind!r}")


def keepout_distance(P: np.ndarray, N: np.ndarray, keepouts: Iterable[KeepOut], c_K: float = 0.5,
                     t_max: float = 50.0, tol: float = 1e-4, max_iter: int = 2000) -> np.ndarray:
    """Distance t from P along +N to the first point where phi_K < c_K, minimum over the keep-outs (§3.5).

    Sphere tracing: phi_K is 1-Lipschitz, so t can advance by phi_K - c_K without crossing the c_K level set; stop
    when within `tol` of it (returns the crossing distance). +inf if not reached within t_max. 0 if P is already
    within c_K of a keep-out. This t is itself the bound on d^hi (it already contains the clearance c_K): design.py
    must not subtract c_K again.
    """
    P, N = _f64(P).reshape(-1, 3), _f64(N).reshape(-1, 3)
    out = np.full(len(P), np.inf)
    for K in keepouts:
        t = np.zeros(len(P))
        alive = np.ones(len(P), bool)
        res = np.full(len(P), np.inf)
        for _ in range(max_iter):
            ia = np.where(alive)[0]
            if len(ia) == 0:
                break
            g = K.phi(P[ia] + t[ia, None] * N[ia]) - c_K
            hit = g <= tol
            res[ia[hit]] = t[ia[hit]]
            alive[ia[hit]] = False
            t[ia] = t[ia] + np.maximum(g, 0.0)
            far = t[ia] > t_max
            alive[ia[far & ~hit]] = False
        out = np.minimum(out, res)
    return out


# ----------------------------------------------------------------------------------------------------------------
# §3.4 passive vertices
# ----------------------------------------------------------------------------------------------------------------
def mount_face_ids(analytic_face_ids: Iterable[int], forced_face_ids: Iterable[int], V0: np.ndarray, T0: np.ndarray,
                   TF0: np.ndarray, auto_analytic: bool = True, min_blob_fraction: float = 0.01) -> list[int]:
    """B-rep faces that make up the mounts, reusing vshell's rule (vshell/segment.py, mounts part only).

    vshell: analytic faces (plane, cylinder, ...) form `mount` blobs, kept only if the blob is larger than
    `min_blob_fraction` of the part (volume there; here the blob's share of the surface *area*, since this module
    has no volume grid); faces carrying supports or loads (`forced_face_ids`, incl. `mounts.extra_faces`) always
    count. A blob is a connected set of analytic faces (faces sharing an input-mesh vertex are connected).
    """
    forced = set(int(i) for i in forced_face_ids)
    out = set(forced)
    if auto_analytic:
        analytic = set(int(i) for i in analytic_face_ids)
        a2 = igl_doublearea(V0, T0) / 2.0
        total = a2.sum()
        face_area = {f: float(a2[TF0 == f].sum()) for f in analytic}
        # connectivity between analytic faces via shared vertices
        vert_faces: dict[int, set[int]] = {}
        for f in analytic:
            for v in np.unique(T0[TF0 == f]):
                vert_faces.setdefault(int(v), set()).add(f)
        parent = {f: f for f in analytic}

        def find(x):
            while parent[x] != x:
                parent[x] = parent[parent[x]]
                x = parent[x]
            return x
        for fs in vert_faces.values():
            fs = list(fs)
            for g in fs[1:]:
                parent[find(g)] = find(fs[0])
        blobs: dict[int, list[int]] = {}
        for f in analytic:
            blobs.setdefault(find(f), []).append(f)
        for members in blobs.values():
            if sum(face_area[f] for f in members) > min_blob_fraction * total or any(f in forced for f in members):
                out.update(members)
    return sorted(out)


def igl_doublearea(V: np.ndarray, T: np.ndarray) -> np.ndarray:
    return np.linalg.norm(np.cross(V[T[:, 1]] - V[T[:, 0]], V[T[:, 2]] - V[T[:, 0]]), axis=1)


def mount_vertices(P: np.ndarray, V0: np.ndarray, T0: np.ndarray, TF0: np.ndarray, mount_ids: Iterable[int],
                   dilate_mm: float = 1.5) -> np.ndarray:
    """Bool mask: points within `dilate_mm` (exact distance) of the input triangles of the mount faces (§3.4)."""
    ids = list(mount_ids)
    if not ids:
        return np.zeros(len(P), bool)
    sel = np.isin(TF0, ids)
    qm = MeshQuery(V0, _i64(T0[sel]))
    return qm.distance(P) <= dilate_mm


def thin_passive(a_hi: np.ndarray, a_lo: float | np.ndarray, v_min: float = 3.0) -> np.ndarray:
    """Thin-zone rule (§3.4/§3.8): passive where a^hi <= a^lo + v_min/2, i.e. no void >= v_min fits in a wall >= a^lo.

    Pure function of the arrays so design.py can apply the D1 variant (e.g. a^lo = t_min per vertex, or a^hi from a
    modified r_MA) without touching surface geometry."""
    return np.asarray(a_hi) <= np.asarray(a_lo) + 0.5 * v_min


def passive_vertices(surf: DesignSurface, mount_ids: Iterable[int] | None = None, a_hi: np.ndarray | None = None,
                     a_lo: float = 1.5, v_min: float = 3.0, dilate_mm: float = 1.5) -> tuple[np.ndarray, dict]:
    """Passive mask = mount zone (within `dilate_mm` of mount faces) OR thin zone. Returns (mask, parts dict).

    Mount ids come from `mount_face_ids`. Pass a_hi = r_MA + gamma_c eps to include the thin rule."""
    mount = (mount_vertices(surf.V, surf.V0, surf.T0, surf.TF0, mount_ids, dilate_mm)
             if mount_ids is not None else np.zeros(surf.n_vertices, bool))
    thin = thin_passive(a_hi, a_lo, v_min) if a_hi is not None else np.zeros(surf.n_vertices, bool)
    return mount | thin, dict(mount=mount, thin=thin)
