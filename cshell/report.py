"""cshell report.pdf (PLAN_single_run.md section 5, contract C4), built from the run directory only.

    python -m cshell report <run_dir>      ->  cshell.report.write_report(run_dir) -> <run_dir>/report.pdf

Eleven pages, in this order (a page whose inputs are missing becomes a "not available" placeholder listing what is
missing, so numbering is stable): summary, geometry and BCs, 3D renders of the final design, sections through the
thickest regions, von Mises solid vs final, capacity, wall-thickness histogram, convergence, voids and escape holes,
resolution check, run time.  There is no comparison to other optimisers.

Inputs (C2/C3 artefacts + config_used.json, all optional):
  summary.json, history.csv, config_used.json, bc_faces.npz, fields.vtp, final_design.stl, wall_thickness.csv,
  escape_holes.csv.
summary.json keys read (C3): name, decision{accepted,reason}, mass{target_g,final_g}, opt{J,K_K0,lamB_ratio,sig_ratio},
  shift{...,feasible}, export{h_mm,faces,watertight,volume_mm3,holes,voids,wall{min,p05,median,max}},
  closed_area_frac, hollow_area_frac, timing{phase_s,finalise_s,export_s}, peak{gpu_gib,host_gib}.
history.csv columns (C2): it, phase, growth_step, R_k, J, K_K0, lamB_ratio, sig_ratio, mass_err, closed_frac,
  hollow_allowed_frac.
bc_faces.npz (C2): V, T (original surface triangles), support_faces, load_faces, force_N.  Accepted layouts: the faces
  as a per-triangle bool mask / int label (-1 = none, i = load i) or as a list of triangle indices.
fields.vtp point data: w (NaN where not hollow), passive (mounts), closed, stress, stress_solid  (+ d, a, r_MA, R, ...).
  Needed beyond C2: stress and stress_solid (von Mises, page 5) and w (page 3, 4).

Adapter for Stage 2 outputs (examples/cshell/runs/*): summary keys (best, resolutions, checks, export.best, timings_s,
resources), history columns (start, iter, C_C0, lam_s_ratio, lam_b_ratio) and bc_faces (support, load) are mapped onto
the C2/C3 names; the final STL is final_design.stl, else final_design_preview.stl.  `mesh_dir` points the STL lookup
elsewhere (e.g. the multi-GB runs_cshell/FULL/<part>); STLs are read by streaming (memory-mapped), never loaded whole.
"""
from __future__ import annotations

import json
import os
import textwrap
import traceback
from timeit import default_timer as timer

import numpy as np

from dshell import report as vr
from dshell import util

PAGES = ["summary", "geometry_bcs", "renders", "sections", "stress", "capacity", "wall_hist", "convergence", "voids",
         "resolution", "runtime"]
C_SOLID, C_MOUNT, C_SUP, C_LOAD = (0.82, 0.82, 0.82), (0.33, 0.33, 0.33), (0.12, 0.47, 0.71), (1.0, 0.5, 0.05)
VIEWS4 = [(25, -60, "isometric"), (90, -90, "top (+z)"), (0, -90, "front (-y)"), (20, 30, "rear isometric")]
LN102 = float(np.log(1.02))


class NA(Exception):
    """A page input is missing; the page becomes a 'not available' page."""


def _need(*conds_msgs):
    miss = [m for ok, m in conds_msgs if not ok]
    if miss:
        raise NA(", ".join(miss))


# ---------------------------------------------------------------------------------------------------------------------
# geometry helpers: BC patches and arrows (testable without any plotting)
# ---------------------------------------------------------------------------------------------------------------------
def _labels(x, n_tri, kind):
    """Per-triangle labels from a mask / label array / index list: bool (support) or int (load, -1 = none)."""
    x = np.asarray(x)
    if kind == "support":
        if x.dtype == bool and x.size == n_tri:
            return x.copy()
        m = np.zeros(n_tri, bool)
        if x.size and x.dtype != bool:
            m[x.astype(int).ravel()] = True
        return m
    if x.size == n_tri and x.dtype != bool and x.size and x.min() >= -1:
        return x.astype(int)
    lab = np.full(n_tri, -1, int)
    if x.dtype == bool and x.size == n_tri:
        lab[x] = 0
    elif x.size:
        lab[x.astype(int).ravel()] = 0
    return lab


def load_bc(z, cfg=None):
    """Normalise bc_faces.npz (dict-like) -> dict(V, T, support(bool per tri), load(int per tri), force(n_load,3))."""
    keys = set(z.keys())
    if "V" not in keys or "T" not in keys:
        raise NA("bc_faces.npz: V, T")
    V, T = np.asarray(z["V"], float), np.asarray(z["T"], int)
    sk = "support_faces" if "support_faces" in keys else ("support" if "support" in keys else None)
    lk = "load_faces" if "load_faces" in keys else ("load" if "load" in keys else None)
    if lk is None:
        raise NA("bc_faces.npz: load_faces")
    sup = _labels(z[sk], len(T), "support") if sk else np.zeros(len(T), bool)
    load = _labels(z[lk], len(T), "load")
    nl = int(load.max()) + 1
    F = None
    if "force_N" in keys:
        F = np.asarray(z["force_N"], float).reshape(-1, 3)
    elif cfg and cfg.get("loads"):
        F = np.array([l.get("force_N", [0, 0, 0]) for l in cfg["loads"]], float).reshape(-1, 3)
    if F is None:
        raise NA("force_N (bc_faces.npz or config_used.json loads)")
    if len(F) < nl:
        F = np.vstack([F, np.tile(F[-1:], (nl - len(F), 1))])
    return {"V": V, "T": T, "support": sup, "load": load, "force": F[:max(nl, 1)]}


def tri_areas(V, T):
    t = V[T]
    return np.linalg.norm(np.cross(t[:, 1] - t[:, 0], t[:, 2] - t[:, 0]), axis=1) / 2


def patches(T, mask):
    """Connected components (shared vertices) of the triangles in `mask`: list of triangle-index arrays."""
    from scipy.sparse import coo_matrix
    from scipy.sparse.csgraph import connected_components
    idx = np.where(mask)[0]
    if idx.size == 0:
        return []
    u, inv = np.unique(T[idx], return_inverse=True)
    inv = inv.reshape(-1, 3)
    r, c = np.concatenate([inv[:, 0], inv[:, 1]]), np.concatenate([inv[:, 1], inv[:, 2]])
    k, lab = connected_components(coo_matrix((np.ones(len(r)), (r, c)), shape=(len(u), len(u))), directed=False)
    tl = lab[inv[:, 0]]
    return [idx[tl == j] for j in range(k)]


def closest_on_triangles(p, tri):
    """Closest point to p on any of the triangles tri (n,3,3) (Ericson, vectorised); returns (point, distance)."""
    a, b, c = tri[:, 0], tri[:, 1], tri[:, 2]
    ab, ac, ap = b - a, c - a, p - a
    d1, d2 = (ab * ap).sum(1), (ac * ap).sum(1)
    bp = p - b
    d3, d4 = (ab * bp).sum(1), (ac * bp).sum(1)
    cp = p - c
    d5, d6 = (ab * cp).sum(1), (ac * cp).sum(1)
    vc, vb, va = d1 * d4 - d3 * d2, d5 * d2 - d1 * d6, d3 * d6 - d5 * d4
    with np.errstate(all="ignore"):
        q = np.empty_like(a)
        den = va + vb + vc
        v, w = vb / den, vc / den
        q[:] = a + ab * v[:, None] + ac * w[:, None]
        m = (vc <= 0) & (d1 >= 0) & (d3 <= 0)
        q[m] = (a + ab * (d1 / (d1 - d3))[:, None])[m]
        m = (vb <= 0) & (d2 >= 0) & (d6 <= 0)
        q[m] = (a + ac * (d2 / (d2 - d6))[:, None])[m]
        m = (va <= 0) & ((d4 - d3) >= 0) & ((d5 - d6) >= 0)
        q[m] = (b + (c - b) * ((d4 - d3) / ((d4 - d3) + (d5 - d6)))[:, None])[m]
        m = (d1 <= 0) & (d2 <= 0)
        q[m] = a[m]
        m = (d3 >= 0) & (d4 <= d3)
        q[m] = b[m]
        m = (d6 >= 0) & (d5 <= d6)
        q[m] = c[m]
    dist = np.linalg.norm(q - p, axis=1)
    j = int(np.nanargmin(dist))
    return q[j], float(dist[j])


def load_arrows(bc):
    """One arrow per connected loaded patch: base = area centroid snapped onto the patch (it is on the surface already
    for a planar patch), vec = F_patch (along F, |F_patch| = |F_i| * patch area / area of load i), area, load id."""
    V, T = bc["V"], bc["T"]
    A = tri_areas(V, T)
    out = []
    for i in range(bc["force"].shape[0]):
        P = patches(T, bc["load"] == i)
        tot = sum(A[p].sum() for p in P)
        for p in P:
            tri = V[T[p]]
            cen = (tri.mean(1) * A[p, None]).sum(0) / A[p].sum()
            base, dist = closest_on_triangles(cen, tri)
            out.append({"load": i, "base": base, "vec": bc["force"][i] * A[p].sum() / tot, "area": float(A[p].sum()),
                        "centroid": cen, "snap_mm": dist, "tris": p})
    return out


# ---------------------------------------------------------------------------------------------------------------------
# streamed STL: vertex-cluster decimation + plane sections in one pass over a (possibly multi-GB) binary STL
# ---------------------------------------------------------------------------------------------------------------------
def _stl_records(path):
    n = os.path.getsize(path)
    with open(path, "rb") as f:
        f.seek(80)
        cnt = int(np.frombuffer(f.read(4), "<u4")[0])
    if n != 84 + 50 * cnt:
        raise ValueError("not a binary STL")
    dt = np.dtype([("n", "<f4", 3), ("v", "<f4", (3, 3)), ("a", "<u2")])
    return np.memmap(path, dt, "r", offset=84, shape=(cnt,))


class _Cluster:
    """Vertex clustering on a grid of cell h: accumulates mean vertex positions and the deduplicated triangles."""

    def __init__(self, h, lo):
        self.h, self.lo, self.pk, self.ps, self.pc, self.tk = h, lo, [], [], [], []
        self.n = 0

    def add(self, tri):
        q = np.floor((tri.reshape(-1, 3) - self.lo) / self.h).astype(np.int64)
        key = (q[:, 0] << 42) | (q[:, 1] << 21) | q[:, 2]
        u, inv = np.unique(key, return_inverse=True)
        s = np.stack([np.bincount(inv, tri.reshape(-1, 3)[:, k].astype(float), len(u)) for k in range(3)], 1)
        self.pk.append(u), self.ps.append(s), self.pc.append(np.bincount(inv, minlength=len(u)))
        k3 = key.reshape(-1, 3)
        ok = (k3[:, 0] != k3[:, 1]) & (k3[:, 1] != k3[:, 2]) & (k3[:, 0] != k3[:, 2])
        t = k3[ok]
        # canonical rotation (smallest key first) so that duplicates collapse, orientation preserved
        r = np.argmin(t, axis=1)
        t = np.take_along_axis(t, (r[:, None] + np.arange(3)) % 3, 1)
        self.tk.append(np.unique(t, axis=0))
        self.n += 1
        if self.n % 8 == 0:
            self._reduce()

    def _reduce(self):
        if not self.pk:
            return
        k = np.concatenate(self.pk)
        u, inv = np.unique(k, return_inverse=True)
        s = np.stack([np.bincount(inv, np.concatenate(self.ps)[:, j], len(u)) for j in range(3)], 1)
        c = np.bincount(inv, np.concatenate(self.pc), len(u))
        self.pk, self.ps, self.pc = [u], [s], [c]
        self.tk = [np.unique(np.concatenate(self.tk), axis=0)] if self.tk else []

    def result(self):
        self._reduce()
        if not self.tk or not len(self.tk[0]):
            return np.zeros((0, 3)), np.zeros((0, 3), int)
        u, s, c = self.pk[0], self.ps[0], self.pc[0]
        V = s / c[:, None]
        T = np.searchsorted(u, self.tk[0])
        return V, T


def _plane_segments(tri, axis, c):
    s = tri[:, :, axis] - c
    out = []
    for i, j in ((0, 1), (1, 2), (2, 0)):
        m = (s[:, i] < 0) != (s[:, j] < 0)
        t = np.where(m, s[:, i] / np.where(m, s[:, i] - s[:, j], 1), np.nan)
        out.append(tri[:, i] + (tri[:, j] - tri[:, i]) * t[:, None])
    P = np.stack(out, 1)                                  # (n, 3 edges, 3)
    valid = ~np.isnan(P[:, :, 0])
    sel = valid.sum(1) == 2
    P, valid = P[sel], valid[sel]
    order = np.argsort(~valid, axis=1, kind="stable")[:, :2]
    seg = np.take_along_axis(P, order[:, :, None], 1)
    return seg


def decimate_mesh(V, T, target):
    """Vertex-cluster decimation of an in-memory mesh to roughly `target` faces."""
    if len(T) <= target:
        return V, T
    A = tri_areas(V, T).sum()
    h = np.sqrt(A / (target / 2.0) * 1.15)
    cl = _Cluster(h, V.min(0) - h)
    cl.add(V[T])
    return cl.result()


def stream_stl(path, target=150000, planes=(), chunk=2_000_000):
    """(V, T, full_faces, sections) from a binary STL by streaming.  planes = [(axis, c), ...] -> sections[i] = (m,2,3)."""
    try:
        rec = _stl_records(path)
    except ValueError:
        import trimesh
        tm = trimesh.load(path, force="mesh")
        V, T = np.asarray(tm.vertices, float), np.asarray(tm.faces)
        secs = []
        for ax, c in planes:
            secs.append(_plane_segments(V[T], ax, c))
        return (*decimate_mesh(V, T, target), len(T), secs)
    N = len(rec)
    smp = np.asarray(rec["v"][::max(1, N // 20000)], float)
    area = np.linalg.norm(np.cross(smp[:, 1] - smp[:, 0], smp[:, 2] - smp[:, 0]), axis=1).mean() / 2 * N
    lo = np.full(3, np.inf)
    hi = np.full(3, -np.inf)
    for i in range(0, N, chunk * 4):
        v = np.asarray(rec["v"][i:i + chunk * 4:max(1, N // 2_000_000)], float).reshape(-1, 3)
        lo, hi = np.minimum(lo, v.min(0)), np.maximum(hi, v.max(0))
    h = np.sqrt(area / (target / 2.0) * 1.15) if N > target else 0.0
    cl = _Cluster(h, lo - max(h, 1.0)) if h > 0 else None
    secs = [[] for _ in planes]
    allT = []
    for i in range(0, N, chunk):
        tri = np.asarray(rec["v"][i:i + chunk], np.float64)
        if cl is not None:
            cl.add(tri)
        else:
            allT.append(tri)
        for k, (ax, c) in enumerate(planes):
            m = (tri[:, :, ax].min(1) < c) & (tri[:, :, ax].max(1) >= c)
            if m.any():
                secs[k].append(_plane_segments(tri[m], ax, c))
    if cl is not None:
        V, T = cl.result()
    else:
        tri = np.concatenate(allT)
        V = tri.reshape(-1, 3)
        T = np.arange(len(V)).reshape(-1, 3)
        u, inv = np.unique(np.round(V, 5), axis=0, return_inverse=True)
        V, T = u, inv.reshape(-1, 3)
    secs = [np.concatenate(s) if s else np.zeros((0, 2, 3)) for s in secs]
    return V, T, N, secs


# ---------------------------------------------------------------------------------------------------------------------
# run access and the Stage 2 adapter
# ---------------------------------------------------------------------------------------------------------------------
def _num(v):
    try:
        v = float(v)
        return v if np.isfinite(v) else None
    except (TypeError, ValueError):
        return None


def _g(d, *path, default=None):
    for k in path:
        if not isinstance(d, dict) or k not in d:
            return default
        d = d[k]
    return d


class Run:
    def __init__(self, d, out=None, mesh_dir=None, dpi=110, stl_target=150000):
        self.dir, self.out, self.dpi, self.target = d, out or d, dpi, stl_target
        self.mesh_dir = mesh_dir or d
        self.plot_dir = os.path.join(self.out, "plots")
        os.makedirs(self.plot_dir, exist_ok=True)
        self.S0 = self._json("summary.json")
        self.cfg = self._json("config_used.json")
        self.log = []
        self.name = self.S0.get("name") or self.cfg.get("name") or os.path.basename(os.path.normpath(d))
        self.S = self._adapt(self.S0)
        self.hist = self._history()
        self._f = self._mesh = self._bc = None

    def p(self, f):
        return os.path.join(self.dir, f)

    def has(self, f):
        return os.path.isfile(self.p(f))

    def _json(self, f):
        try:
            with open(self.p(f), encoding="utf-8") as fh:
                x = json.load(fh)
            return x if isinstance(x, dict) else {}
        except Exception:
            return {}

    def csv(self, f):
        try:
            return util.read_csv(self.p(f))
        except Exception:
            return []

    # -- summary: C3 keys, with Stage 2 fallbacks ---------------------------------------------------------------
    def _adapt(self, S):
        o = dict(S)
        dec = S.get("decision")
        if not isinstance(dec, dict):
            dec = {"accepted": S.get("accepted"), "reason": S.get("reason") or S.get("outcome")}
        o["decision"] = dec
        m = S.get("mass") or {}
        o["mass"] = {"target_g": m.get("target_g", m.get("target")), "final_g": m.get("final_g", m.get("final")),
                     "error": m.get("error")}
        rows = S.get("resolutions")
        if isinstance(rows, list) and rows and "opt" not in S:
            def pick(res):
                sol = [r for r in rows if r.get("res") == res and r.get("design") == "solid"]
                des = [r for r in rows if r.get("res") == res and r.get("design") not in
                       ("solid", "uniform_shell", "dshell", "neutral", "uniform")]
                des = des[-1:] or [r for r in rows if r.get("res") == res][-1:]
                if not des:
                    return None
                r, s0 = des[0], (sol[0] if sol else {})
                r_ = lambda a, b: (_num(a) / _num(b)) if _num(a) is not None and _num(b) else None
                return {"J": r.get("J"), "K_K0": r.get("K_K0"), "lamB_ratio": r_(r.get("lam_b"), s0.get("lam_b")),
                        "sig_ratio": r_(r.get("sigma_pn"), s0.get("sigma_pn")) or
                        (1 / r["P_P0"] if _num(r.get("P_P0")) else None),
                        "P_P0": r.get("P_P0"), "resolution": res}
            chk = {c.get("res"): c for c in S.get("checks") or []}
            for res, key in (("opt", "opt"), ("shift", "shift"), ("coarse", "coarse")):
                v = pick(res)
                if v:
                    c = chk.get(res, {})
                    v["feasible"] = bool(c.get("pass_J", True) and c.get("pass_K", True) and c.get("pass_P", True)) \
                        if c else None
                    o[key] = v
        ex = S.get("export")
        if isinstance(ex, dict) and "best" in ex and isinstance(ex["best"], dict):
            b = ex["best"]
            w = b.get("wall") or {}
            o["export"] = {"h_mm": _g(b, "mesh", "h_mm") or _g(S, "files_h"), "faces": _g(b, "mesh", "faces"),
                           "watertight": _g(b, "mesh", "watertight"),
                           "volume_mm3": _g(b, "mesh", "volume_mm3"), "holes": b.get("holes"),
                           "voids": b.get("voids"),
                           "wall": {"min": w.get("min_nominal"), "p05": w.get("p1_nominal"),
                                    "median": w.get("median_nominal"), "max": w.get("max")}}
        t = S.get("timings_s")
        if isinstance(t, dict) and "timing" not in S:
            o["timing"] = {"stages_s": {k: v for k, v in t.items() if k != "total"}, "total_s": t.get("total")}
        r = S.get("resources")
        if isinstance(r, dict) and "peak" not in S:
            hr = r.get("host_rss_GiB") or {}
            o["peak"] = {"gpu_gib": r.get("peak_gpu_collect_GiB", r.get("peak_gpu_GiB")),
                         "host_gib": hr.get("peak_rss_GiB") if isinstance(hr, dict) else hr}
        return o

    def _history(self):
        rows = self.csv("history.csv")
        if rows and "start" in rows[0]:
            best = _g(self.S0, "best", "start")
            sel = [r for r in rows if r.get("start") == best] or rows
            rows = sel
        for r in rows:                                    # Stage 2 names -> C2 names
            if "lamB_ratio" not in r and "lam_b_ratio" in r:
                r["lamB_ratio"] = r["lam_b_ratio"]
            if "K_K0" not in r and _num(r.get("C_C0")):
                r["K_K0"] = 1 / r["C_C0"]
            if "sig_ratio" not in r and _num(r.get("lam_s_ratio")):
                r["sig_ratio"] = 1 / r["lam_s_ratio"]
            if "it" not in r and "iter" in r:
                r["it"] = r["iter"]
        return rows

    # -- fields -------------------------------------------------------------------------------------------------
    def fields(self):
        """dict of point arrays + 'points' from fields.vtp, or None."""
        if self._f is None:
            self._f = False
            if self.has("fields.vtp"):
                import pyvista as pv
                m = pv.read(self.p("fields.vtp"))
                F = {k: np.asarray(m.point_data[k], float) for k in m.point_data.keys()
                     if np.asarray(m.point_data[k]).ndim == 1}
                F["points"] = np.asarray(m.points, float)
                self._f = F
        return self._f or None

    def field(self, *names):
        f = self.fields()
        if f is None:
            raise NA("fields.vtp")
        miss = [n for n in names if n not in f]
        if miss:
            raise NA("fields.vtp arrays: " + ", ".join(miss))
        return [f[n] for n in names]

    def classify(self, X):
        """Per query point: nearest fields.vtp point's class -> (kind: 0 hollow, 1 solid, 2 mount; w)."""
        from scipy.spatial import cKDTree
        F = self.fields()
        if F is None or "w" not in F:
            raise NA("fields.vtp arrays: w")
        j = cKDTree(F["points"]).query(X)[1]
        w = F["w"][j]
        kind = np.zeros(len(X), int)
        passive = F.get("passive")
        if passive is not None:
            kind[passive[j] > 0.5] = 2
        closed = F.get("closed")
        solid = (closed[j] > 0.5) if closed is not None else np.isnan(w)
        kind[(kind == 0) & solid] = 1
        return kind, np.where(kind == 0, w, np.nan)

    def thick_points(self):
        """Two thickest-wall locations (>= 40 mm apart) as (point, w)."""
        F = self.fields()
        if F is None or "w" not in F:
            raise NA("fields.vtp arrays: w")
        w = np.where(np.isfinite(F["w"]), F["w"], -1)
        if "passive" in F:
            w = np.where(F["passive"] > 0.5, -1, w)
        P = F["points"]
        if w.max() <= 0:
            raise NA("hollow points in fields.vtp (w)")
        i = int(np.argmax(w))
        out = [(P[i], float(w[i]))]
        far = np.linalg.norm(P - P[i], axis=1) >= 40
        if (far & (w > 0)).any():
            j = int(np.argmax(np.where(far, w, -1)))
            out.append((P[j], float(w[j])))
        return out

    def bc(self):
        if self._bc is None:
            if not self.has("bc_faces.npz"):
                raise NA("bc_faces.npz")
            self._bc = load_bc(np.load(self.p("bc_faces.npz"), allow_pickle=True), self.cfg)
        return self._bc

    def stl_path(self):
        for d in dict.fromkeys([self.mesh_dir, self.dir]):
            for f in ("final_design.stl", "final_design_preview.stl"):
                if os.path.isfile(os.path.join(d, f)):
                    return os.path.join(d, f)
        raise NA("final_design.stl")

    def mesh(self):
        """(V, T, full_faces, sections list) of the decimated final design; sections for self.plane_list()."""
        if self._mesh is None:
            planes = []
            try:
                for pt, _ in self.thick_points():
                    planes += [(a, float(pt[a])) for a in range(3)]
            except NA:
                pass
            self._planes = planes
            self._mesh = stream_stl(self.stl_path(), self.target, planes)
        return self._mesh


# ---------------------------------------------------------------------------------------------------------------------
# drawing helpers
# ---------------------------------------------------------------------------------------------------------------------
def _fmt(v, nd=4):
    return vr._fmt(v, nd)


def _na_page(title, why):
    fig = vr.new_figure(8.27, 11.69)
    fig.text(0.5, 0.6, f"{title}\n\nnot available", ha="center", fontsize=14, weight="bold", color="0.4")
    fig.text(0.5, 0.52, f"missing input: {why}", ha="center", fontsize=9, color="0.4")
    return fig


def _text(fig, x, y, lines, size=7.4):
    fig.text(x, y, "\n".join(lines), family="monospace", fontsize=size, va="top")


def _shade(V, T, rgb, view):
    """Face colours * Lambert-style shade; back faces culled for the given (elev, azim) view."""
    tri = V[T]
    n = np.cross(tri[:, 1] - tri[:, 0], tri[:, 2] - tri[:, 0])
    n /= np.linalg.norm(n, axis=1, keepdims=True) + 1e-30
    el, az = np.radians(view[0]), np.radians(view[1])
    toward = np.array([np.cos(el) * np.cos(az), np.cos(el) * np.sin(az), np.sin(el)])   # direction to the camera
    front = (n @ toward) > -0.05                          # slightly lenient; sorting handles the rest
    light = 0.35 + 0.65 * np.clip(np.abs(n @ toward) * 0.8 + 0.2 * np.clip(n @ np.array([0.3, -0.4, 0.85]), 0, 1), 0, 1)
    return tri, np.clip(rgb * light[:, None], 0, 1), front


def _draw(ax, V, T, rgb, view, equal=None):
    from mpl_toolkits.mplot3d.art3d import Poly3DCollection
    tri, col, front = _shade(V, T, np.asarray(rgb)[:, :3], view)
    pc = Poly3DCollection(tri[front], facecolors=col[front], edgecolor="none", linewidths=0)
    pc.set_rasterized(True)
    ax.add_collection3d(pc)
    P = V if equal is None else equal
    lo, hi = P.min(0), P.max(0)
    pad = 0.03 * (hi - lo).max()
    ax.set_xlim(lo[0] - pad, hi[0] + pad), ax.set_ylim(lo[1] - pad, hi[1] + pad), ax.set_zlim(lo[2] - pad, hi[2] + pad)
    ax.set_box_aspect(np.maximum(hi - lo + 2 * pad, 1e-6))            # equal scale on all axes, box fitted to the part
    ax.set_xlabel("x", fontsize=7), ax.set_ylabel("y", fontsize=7), ax.set_zlabel("z", fontsize=7)
    ax.tick_params(labelsize=6)
    ax.view_init(view[0], view[1])


def _cbar(fig, cmap, vmin, vmax, label, rect):
    """A single colourbar in its own axes (rect = [l, b, w, h] in figure coords): never overlaps the plots."""
    from matplotlib import cm, colors
    cax = fig.add_axes(rect)
    cb = fig.colorbar(cm.ScalarMappable(colors.Normalize(vmin, vmax), cmap), cax=cax)
    cb.set_label(label, fontsize=8)
    cb.ax.tick_params(labelsize=7)
    return cb


def _vals_to_rgb(vals, cmap, vmin, vmax):
    import matplotlib
    from matplotlib import colors
    return matplotlib.colormaps[cmap](colors.Normalize(vmin, vmax)(np.where(np.isfinite(vals), vals, vmin)))[:, :3]


def _face_mean(T, v):
    import warnings
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", RuntimeWarning)
        return np.nanmean(np.where(np.isfinite(v[T]), v[T], np.nan), axis=1) if np.isfinite(v).any() else np.full(len(T), np.nan)


# ---------------------------------------------------------------------------------------------------------------------
# pages
# ---------------------------------------------------------------------------------------------------------------------
def p_summary(R):
    S = R.S
    _need((bool(R.S0), "summary.json"))
    dec = S["decision"]
    acc = dec.get("accepted")
    banner = {True: "HOLLOW DESIGN ACCEPTED", False: "ORIGINAL SOLID KEPT"}.get(acc, "DECISION UNKNOWN")
    fig = vr.new_figure(8.27, 11.69)
    fig.text(0.04, 0.975, f"{R.name}: {banner}", fontsize=13, weight="bold", color=vr.C_OK if acc else vr.C_SOLID, va="top")
    lines = []
    for ln in str(dec.get("reason") or "").split("\n"):
        lines += textwrap.wrap(ln, 105) or [""]
    m = S["mass"]
    lines += ["", f"mass      target {_fmt(m.get('target_g'))} g   final {_fmt(m.get('final_g'))} g   error {_fmt(m.get('error'))}"]
    for key, lab in (("opt", "optimiser grid"), ("shift", "shifted grid")):
        v = S.get(key)
        if isinstance(v, dict):
            lines.append(f"{lab:15s} J {_fmt(v.get('J'))}  K/K0 {_fmt(v.get('K_K0'))}  lamB/lamB0 {_fmt(v.get('lamB_ratio'))}"
                         f"  sigPN/sigPN0 {_fmt(v.get('sig_ratio'))}  feasible {v.get('feasible')}   (accept: J > ln 1.02 = {LN102:.4f},"
                         " lamB ratio >= 1, sigma ratio <= 1)")
    ex = S.get("export")
    if isinstance(ex, dict):
        w = ex.get("wall") or {}
        lines += ["", f"export    h {_fmt(ex.get('h_mm'))} mm  faces {_fmt(ex.get('faces'))}  watertight {ex.get('watertight')}"
                  f"  volume {_fmt(ex.get('volume_mm3'))} mm3  voids {_fmt(_g(ex, 'voids', 'total'))}  escape holes {_fmt(ex.get('holes'))}",
                  f"wall [mm] min {_fmt(w.get('min'))}  p05 {_fmt(w.get('p05'))}  median {_fmt(w.get('median'))}  max {_fmt(w.get('max'))}"]
    for k in ("closed_area_frac", "hollow_area_frac"):
        if k in S:
            lines.append(f"{k:18s} {_fmt(S[k])}")
    o = R.cfg.get("optimiser", {})
    if o:
        lines += ["", "optimiser: " + _fmt(o), "grid: " + _fmt(R.cfg.get("grid")), "growth: " + _fmt(R.cfg.get("growth")),
                  "objective: " + _fmt(R.cfg.get("objective")), "export cfg: " + _fmt(R.cfg.get("export"))]
    shown = {"name", "decision", "mass", "opt", "shift", "export", "closed_area_frac", "hollow_area_frac", "timing", "peak",
             "accepted", "reason", "outcome"}
    wrapped = []
    for ln in lines:
        wrapped += textwrap.wrap(ln, 125, subsequent_indent="    ") or [""]
    _text(fig, 0.04, 0.945, wrapped[:95])
    return fig


def p_geometry(R):
    """Original part with the faces the analysis uses: supports blue, loaded faces orange, one arrow per loaded patch."""
    bc = R.bc()
    V, T = bc["V"], bc["T"]
    rgb = np.tile(C_SOLID, (len(T), 1))
    rgb[bc["support"]] = C_SUP
    rgb[bc["load"] >= 0] = C_LOAD
    arrows = load_arrows(bc)
    span = np.ptp(V, 0).max()
    fig = vr.new_figure(15, 8.2)
    Ftot = np.linalg.norm(bc["force"].sum(0))
    for j, view in enumerate(VIEWS4):
        ax = fig.add_subplot(2, 2, j + 1, projection="3d")
        _draw(ax, V, T, rgb, view)
        for a in arrows:
            u = a["vec"] / (np.linalg.norm(bc["force"].sum(0)) + 1e-12)       # share of |F| sets the length
            d = u * 0.20 * span
            ax.quiver(*(a["base"] - d), *d, color=vr.C_BAD, lw=2.2, arrow_length_ratio=0.22)
            if j == 0:
                ax.text(*(a["base"] - d * 1.08), f"{np.linalg.norm(a['vec']):.0f} N", fontsize=7, color=vr.C_BAD)
        ax.set_title(view[2], fontsize=9)
    ns = len(patches(T, bc["support"]))
    fig.suptitle(f"{R.name}: supports (blue, all DOF fixed; {ns} patch{'es' if ns != 1 else ''}) and loads (orange faces; "
                 f"{len(arrows)} red arrow{'s' if len(arrows) != 1 else ''}, one per loaded patch at its area centroid, "
                 f"length and label = its share of the total {Ftot:.0f} N). Faces read from bc_faces.npz (as analysed).",
                 fontsize=9)
    fig.subplots_adjust(left=0.01, right=0.99, top=0.93, bottom=0.01, wspace=0.0, hspace=0.05)
    return fig


def p_renders(R):
    V, T, nfull, _ = R.mesh()
    kind, w = R.classify(V)
    wf = _face_mean(T, w)
    # face class by majority of its vertices
    km = (kind[T] == 2).sum(1) >= 2
    ks = ((kind[T] == 1).sum(1) >= 2) & ~km
    hollow = ~km & ~ks & np.isfinite(wf)
    ks |= ~km & ~ks & ~hollow
    hv = w[np.isfinite(w)]
    lo, hi = (float(np.nanpercentile(hv, 1)), float(np.nanpercentile(hv, 99.5))) if hv.size else (0, 1)
    if hi <= lo:
        hi = lo + 1
    rgb = np.tile(C_SOLID, (len(T), 1))
    rgb[km] = C_MOUNT
    rgb[hollow] = _vals_to_rgb(wf[hollow], "viridis", lo, hi)
    fig = vr.new_figure(12, 11.5)
    for j, view in enumerate(VIEWS4):
        ax = fig.add_subplot(2, 2, j + 1, projection="3d")
        _draw(ax, V, T, rgb, view)
        ax.set_title(view[2], fontsize=9)
    fig.subplots_adjust(left=0.01, right=0.88, top=0.93, bottom=0.04, wspace=0.0, hspace=0.05)
    _cbar(fig, "viridis", lo, hi, "wall thickness w [mm] (hollow regions)", [0.91, 0.25, 0.018, 0.5])
    fig.text(0.5, 0.012, "light grey = solid (closed) regions; dark grey = mounts; the colourbar covers the hollow regions only",
             ha="center", fontsize=8)
    fig.suptitle(f"{R.name}: final design (STL {nfull:,} faces, shown decimated to {len(T):,}; wall thickness from fields.vtp)",
                 fontsize=10)
    return fig


def p_sections(R):
    R.mesh()
    pts = R.thick_points()
    secs = R._mesh[3]
    _need((bool(secs), "fields.vtp w (section locations)"))
    from matplotlib.collections import LineCollection
    fig = vr.new_figure(15, 5.2 * len(pts))
    win = 30.0
    for r, (pt, w) in enumerate(pts):
        for a in range(3):
            ax = fig.add_subplot(len(pts), 3, 3 * r + a + 1)
            seg = secs[3 * r + a]
            ij = [i for i in range(3) if i != a]
            if len(seg):
                lc = LineCollection(seg[:, :, ij], colors="k", linewidths=0.7)
                ax.add_collection(lc)
            ax.plot(*pt[ij], "r+", ms=10)
            ax.set_xlim(pt[ij[0]] - win, pt[ij[0]] + win)
            ax.set_ylim(pt[ij[1]] - win, pt[ij[1]] + win)
            ax.set_aspect("equal")
            ax.set_xlabel("xyz"[ij[0]] + " [mm]", fontsize=8)
            ax.set_ylabel("xyz"[ij[1]] + " [mm]", fontsize=8)
            ax.tick_params(labelsize=7)
            ax.set_title(f"region {r + 1} (w = {w:.2f} mm): cut normal {'xyz'[a]} = {pt[a]:.1f} mm", fontsize=8)
    fig.suptitle(f"{R.name}: sections of the final design through the two thickest-wall regions (red +); +-{win:.0f} mm window",
                 fontsize=10, y=0.995)
    fig.tight_layout(rect=(0, 0, 1, 0.95 if len(pts) == 1 else 0.97), h_pad=3.5)
    return fig


def p_stress(R):
    sv, sf = R.field("stress_solid", "stress")
    F = R.fields()
    # final design: decimated STL surface, stress from the nearest fields point; solid: the original surface triangles
    V, T, nfull, _ = R.mesh()
    from scipy.spatial import cKDTree
    j = cKDTree(F["points"]).query(V)[1]
    vmax = float(np.nanpercentile(np.concatenate([sv, sf]), 99.5))
    bc = R.bc()
    Vs, Ts = decimate_mesh(bc["V"], bc["T"], R.target)
    js = cKDTree(F["points"]).query(Vs)[1]
    fig = vr.new_figure(15, 8.6)
    views = VIEWS4[:2]
    for r, (lab, Vv, Tt, val) in enumerate((("solid", Vs, Ts, sv[js]), ("final design", V, T, sf[j]))):
        rgb = _vals_to_rgb(_face_mean(Tt, val), "inferno", 0, vmax)
        for c, view in enumerate(views):
            ax = fig.add_subplot(2, 2, 2 * r + c + 1, projection="3d")
            _draw(ax, Vv, Tt, rgb, view)
            ax.set_title(f"{lab}: {view[2]}", fontsize=9)
    fig.subplots_adjust(left=0.01, right=0.88, top=0.93, bottom=0.02, wspace=0.0, hspace=0.08)
    _cbar(fig, "inferno", 0, vmax, "von Mises [MPa] (same scale, clipped at the 99.5th percentile)", [0.91, 0.25, 0.018, 0.5])
    fig.suptitle(f"{R.name}: von Mises stress, original solid vs final design (peak solid {np.nanmax(sv):.1f}, final {np.nanmax(sf):.1f} MPa; "
                 "values from fields.vtp, nearest point)", fontsize=10)
    return fig


def _opt_shift(R):
    rows = [(lab, R.S.get(k)) for k, lab in (("opt", "optimiser grid"), ("shift", "shifted grid")) if isinstance(R.S.get(k), dict)]
    return rows


def p_capacity(R):
    rows = _opt_shift(R)
    _need((bool(rows), "summary.json: opt / shift"))
    fig = vr.new_figure(10, 7)
    ax = fig.add_subplot(1, 1, 1)
    names = ["buckling  lamB / lamB0", "strength  sigPN0 / sigPN", "stiffness  K / K0"]
    x = np.arange(3)
    wd = 0.8 / len(rows)
    cols = [vr.C_HOLLOW, vr.C_ACC]
    for i, (lab, v) in enumerate(rows):
        s = _num(v.get("sig_ratio"))
        y = [_num(v.get("lamB_ratio")), (1 / s) if s else None, _num(v.get("K_K0"))]
        ys = [np.nan if q is None else q for q in y]
        b = ax.bar(x + (i - (len(rows) - 1) / 2) * wd, ys, wd * 0.95, color=cols[i % 2], label=lab)
        for xx, q in zip(x + (i - (len(rows) - 1) / 2) * wd, ys):
            ax.text(xx, 0 if np.isnan(q) else q, "n/a" if np.isnan(q) else f"{q:.3f}", ha="center", va="bottom", fontsize=8)
    ax.axhline(1.0, color=vr.C_SOLID, lw=1.4, label="original solid / constraint (= 1)")
    ax.set_xticks(x, names)
    ax.set_ylabel("ratio to the original solid (higher is better)")
    ax.legend(fontsize=8)
    ax.set_title(f"{R.name}: capacity by mode against the solid's constraints. The buckling and strength constraints require "
                 "both bars >= 1.", fontsize=9)
    fig.tight_layout()
    return fig


def p_wall_hist(R):
    rows = R.csv("wall_thickness.csv")
    _need((bool(rows), "wall_thickness.csv"))
    cols = list(rows[0])
    fig = vr.new_figure(10, 6.5)
    ax = fig.add_subplot(1, 1, 1)
    tmin = _g(R.cfg, "manufacturing", "t_min_mm")
    if "bin_lo_mm" in cols and "bin_hi_mm" in cols:
        lo = np.array([_num(r["bin_lo_mm"]) for r in rows], float)
        hi = np.array([_num(r["bin_hi_mm"]) for r in rows], float)
        cnt = [c for c in cols if c.startswith("count")]
        for i, c in enumerate(cnt):
            y = np.array([_num(r[c]) or 0 for r in rows], float)
            if i == 0:
                ax.bar(lo, y, hi - lo, align="edge", color=vr.C_HOLLOW, alpha=0.55, label=c.replace("count_", ""))
            else:
                ax.stairs(y, np.append(lo, hi[-1]), color=vr.C_ACC, lw=1.4, label=c.replace("count_", ""))
        ax.set_xlim(0, hi[-1])
        ax.set_ylabel("count of wall samples")
    else:                                                  # raw samples: any column named like a thickness in mm
        k = next((c for c in cols if "mm" in c.lower() or c.lower().startswith("w")), cols[0])
        v = np.array([_num(r[k]) for r in rows], float)
        ax.hist(v[np.isfinite(v)], bins=60, color=vr.C_HOLLOW, label=k)
        ax.set_ylabel("count")
    if tmin:
        ax.axvline(tmin, color=vr.C_BAD, ls="--", label=f"t_min = {tmin} mm")
    ax.set_xlabel("wall thickness [mm]")
    ax.legend(fontsize=8)
    w = _g(R.S, "export", "wall") or {}
    ax.set_title(f"{R.name}: wall thickness (min {_fmt(w.get('min'), 3)}, p05 {_fmt(w.get('p05'), 3)}, median {_fmt(w.get('median'), 3)}, "
                 f"max {_fmt(w.get('max'), 3)} mm)", fontsize=9)
    fig.tight_layout()
    return fig


def p_convergence(R):
    H = R.hist
    _need((bool(H), "history.csv"))
    col = lambda k: np.array([np.nan if _num(r.get(k)) is None else float(r[k]) for r in H])
    x = np.arange(len(H))
    keys = [("J", "J = ln K/K0 (accept above ln 1.02)", LN102), ("K_K0", "K / K0", None), ("lamB_ratio", "lamB / lamB0 (>= 1)", 1.0),
            ("sig_ratio", "sigPN / sigPN0 (<= 1)", 1.0), ("mass_err", "mass error V/V0 - 1", 0.0),
            ("R_k", "R_k (hollow allowed where R <= R_k)", None), ("closed_frac", "closed fraction", None),
            ("hollow_allowed_frac", "hollow-allowed fraction", None)]
    gs, ph = col("growth_step"), col("phase")
    fig = vr.new_figure(15, 10.5)
    for i, (k, lab, ref) in enumerate(keys):
        ax = fig.add_subplot(3, 3, i + 1)
        y = col(k)
        if np.isfinite(y).any():
            ax.plot(x, y, lw=1.0, color=vr.C_HOLLOW)
        else:
            ax.text(0.5, 0.5, f"no '{k}' column", transform=ax.transAxes, ha="center", color="0.5", fontsize=8)
        if ref is not None:
            ax.axhline(ref, color=vr.C_BAD, lw=0.7, ls="--")
        if len(ph) > 1:
            for b in np.where(np.diff(np.nan_to_num(ph)) != 0)[0] + 1:
                ax.axvline(b, color="0.5", lw=0.8)
        if np.isfinite(gs).any():
            for b in np.where(np.diff(np.nan_to_num(gs, nan=-1)) != 0)[0] + 1:
                ax.axvline(b, color=vr.C_ACC, lw=0.5, alpha=0.6)
        ax.set_title(lab, fontsize=8)
        ax.set_xlim(0, max(len(H) - 1, 1))
        ax.tick_params(labelsize=7)
        if i >= 5:
            ax.set_xlabel("iteration", fontsize=8)
    ax = fig.add_subplot(3, 3, 9)
    ax.axis("off")
    n_gs = int(np.nanmax(gs)) + 1 if np.isfinite(gs).any() else 0
    ax.text(0, 1, "grey vertical: phase boundary\norange vertical: growth step (R_k rise)\n"
            f"iterations {len(H)}; growth steps {n_gs}\nred dashed: constraint / acceptance level", va="top", fontsize=8)
    fig.suptitle(f"{R.name}: convergence history", fontsize=11)
    fig.tight_layout(rect=(0, 0, 1, 0.96))
    return fig


def p_voids(R):
    holes = R.csv("escape_holes.csv")
    voids = _g(R.S, "export", "voids") or {}
    _need((bool(holes) or bool(voids), "escape_holes.csv, summary.json export.voids"))
    fig = vr.new_figure(8.27, 11.69)
    nv = voids.get("total", len({h.get("void_id", h.get("void")) for h in holes}))
    lines = [f"enclosed voids: {_fmt(nv)};  escape holes: {len(holes)};  voids without a hole: {_fmt(voids.get('without_hole'))}"]
    for v in (voids.get("per_void") or [])[:30]:
        lines.append(f"  void {v.get('id')}: volume {_fmt(v.get('volume_mm3'), 5)} mm3, holes {v.get('holes')}")
    if holes:
        hc = ["void_id", "hole", "x", "y", "z", "wall_mm", "score_kind"]
        hc = [c for c in hc if c in holes[0]] or list(holes[0])[:7]
        f = lambda v: str(int(v)) if isinstance(v, float) and v.is_integer() else (f"{v:.2f}" if isinstance(v, float) else str(v))
        lines += ["", "  ".join(f"{c:>10s}" for c in hc)] + ["  ".join(f"{f(r.get(c))[:10]:>10s}" for c in hc) for r in holes[:24]]
    lines += ["", "Escape holes sit at the lowest-stress wall points of each void (outward direction in escape_holes.csv)."]
    _text(fig, 0.04, 0.96, lines)
    if holes and all(k in holes[0] for k in "xyz"):
        P = np.array([[h["x"], h["y"], h["z"]] for h in holes], float)
        vid = np.array([int(h.get("void_id", h.get("void", 0))) for h in holes])
        try:
            V, T, _, _ = R.mesh()
        except NA:
            V = T = None
        from matplotlib.collections import PolyCollection
        for i, (a, b) in enumerate(((0, 1), (0, 2))):
            ax = fig.add_axes([0.08, 0.30 - 0.25 * i, 0.84, 0.19 if i == 0 else 0.15])
            if V is not None:
                ax.add_collection(PolyCollection(V[T][:, :, [a, b]], facecolors="0.88", edgecolors="none", rasterized=True))
            ax.scatter(P[:, a], P[:, b], c=vid, cmap="tab10", s=24, edgecolors="k", linewidths=0.4, zorder=3)
            if V is not None:
                ax.set_xlim(V[:, a].min() - 5, V[:, a].max() + 5)
                ax.set_ylim(V[:, b].min() - 5, V[:, b].max() + 5)
            ax.set_aspect("equal")
            ax.set_xlabel("xyz"[a] + " [mm]", fontsize=8)
            ax.set_ylabel("xyz"[b] + " [mm]", fontsize=8)
            ax.tick_params(labelsize=7)
            ax.set_title("escape hole positions (colour = void id) over the part silhouette", fontsize=8)
    return fig


def p_resolution(R):
    rows = _opt_shift(R)
    _need((len(rows) >= 1, "summary.json: opt / shift"))
    fig = vr.new_figure(12, 6.5)
    panels = [("J", "J", LN102), ("K_K0", "K / K0", None), ("lamB_ratio", "lamB / lamB0", 1.0), ("sig_ratio", "sigPN / sigPN0", 1.0)]
    for i, (k, lab, ref) in enumerate(panels):
        ax = fig.add_subplot(1, 4, i + 1)
        y = [np.nan if _num(v.get(k)) is None else float(v[k]) for _, v in rows]
        ax.bar(range(len(rows)), y, color=[vr.C_HOLLOW, vr.C_ACC][:len(rows)], width=0.6)
        for xx, q in enumerate(y):
            ax.text(xx, 0 if np.isnan(q) else q, "n/a" if np.isnan(q) else f"{q:.4f}", ha="center", va="bottom", fontsize=8)
        if ref is not None:
            ax.axhline(ref, color=vr.C_BAD, ls="--", lw=0.8)
        ax.set_xticks(range(len(rows)), [r[0].replace(" grid", "") for r in rows], fontsize=8)
        ax.set_title(lab + ("  (red: ln 1.02)" if k == "J" else ("  (red: limit 1)" if ref else "")), fontsize=8)
        ax.tick_params(labelsize=7)
    feas = {lab: v.get("feasible") for lab, v in rows}
    fig.suptitle(f"{R.name}: resolution check, optimiser grid vs shifted grid; feasible: {feas}", fontsize=10)
    fig.tight_layout(rect=(0, 0, 1, 0.92))
    return fig


def p_runtime(R):
    t = R.S.get("timing") or {}
    peak = R.S.get("peak") or {}
    _need((bool(t) or bool(peak), "summary.json: timing, peak"))
    labels, vals = [], []
    for i, s in enumerate(t.get("phase_s") or []):
        labels.append(f"phase {i}"), vals.append(s)
    for k in ("finalise_s", "export_s"):
        if _num(t.get(k)) is not None:
            labels.append(k[:-2]), vals.append(t[k])
    for k, v in (t.get("stages_s") or {}).items():
        labels.append(k), vals.append(v)
    fig = vr.new_figure(9, 5.2)
    ax = fig.add_subplot(1, 1, 1)
    if vals:
        ax.barh(labels, vals, color=vr.C_HOLLOW)
        ax.invert_yaxis()
    ax.set_xlabel("seconds")
    tot = t.get("total_s") or (sum(t.get("phase_s") or []) + (_num(t.get("finalise_s")) or 0))
    ax.set_title(f"{R.name}: run time by stage (total {tot:.0f} s = {tot / 3600:.2f} h); peak GPU {_fmt(peak.get('gpu_gib'), 3)} GiB, "
                 f"host {_fmt(peak.get('host_gib'), 3)} GiB", fontsize=9)
    fig.tight_layout()
    return fig


PAGE_FUNCS = [("01_summary", p_summary), ("02_geometry_bcs", p_geometry), ("03_renders", p_renders), ("04_sections", p_sections),
              ("05_stress", p_stress), ("06_capacity", p_capacity), ("07_wall_hist", p_wall_hist), ("08_convergence", p_convergence),
              ("09_voids", p_voids), ("10_resolution", p_resolution), ("11_runtime", p_runtime)]


def write_report(run_dir: str, out_dir: str | None = None, mesh_dir: str | None = None, dpi: int = 110,
                 verbose: bool = True) -> str:
    """Write <out_dir or run_dir>/report.pdf (11 pages) and plots/*.png from the run artefacts; return the path."""
    from matplotlib.backends.backend_pdf import PdfPages
    t0 = timer()
    R = Run(run_dir, out_dir, mesh_dir, dpi)
    log_path = os.path.join(R.plot_dir, "report_log.txt")
    if os.path.exists(log_path):
        os.remove(log_path)
    pdf_path = os.path.join(R.out, "report.pdf")
    n = 0
    with PdfPages(pdf_path) as pdf:
        for (name, fn), title in zip(PAGE_FUNCS, PAGES):
            try:
                fig = fn(R)
            except NA as e:
                fig = _na_page(f"{name}: {title}", str(e))
            except Exception as e:
                R.log.append(f"page '{name}' failed: {type(e).__name__}: {e}")
                with open(log_path, "a") as f:
                    f.write(f"--- {name} ---\n{traceback.format_exc()}\n")
                fig = _na_page(f"{name}: {title}", f"error {type(e).__name__}: {e} (see plots/report_log.txt)")
            try:
                fig.savefig(os.path.join(R.plot_dir, f"{name}.png"), dpi=dpi)
                pdf.savefig(fig, dpi=min(dpi, 150))
                n += 1
            except Exception as e:
                R.log.append(f"saving '{name}' failed: {e}")
        pdf.infodict()["Title"] = f"cshell report: {R.name}"
    if R.log:
        with open(log_path, "a") as f:
            f.write("\n".join(R.log) + "\n")
    if verbose:
        print(f"  [report] {R.name}: {n} pages -> {pdf_path} ({timer() - t0:.1f} s)"
              + (f"; {len(R.log)} issue(s), see plots/report_log.txt" if R.log else ""))
    return pdf_path


def build(run_dir: str, dpi: int = 110, verbose: bool = True, **kw) -> str:       # name used by cli.py / older callers
    return write_report(run_dir, dpi=dpi, verbose=verbose, **kw)
