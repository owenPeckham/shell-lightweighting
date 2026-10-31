"""Sparse signed-distance band on the geometry grid (THEORY §3.1, §3.6, §3.7, §3.10, §4.2).

The band is  B = {x : lo <= φ(x) <= D_comp + h_g}  with  lo = -r_in - h_g  (r_in = inf: the whole interior),
φ = exact signed distance to the watertight tessellation (negative inside, sign by generalised winding number).
It is stored as a sorted list of active cells (flat index into the bounding grid, int64), with per cell
φ (float32, mm), the closest triangle of the fine tessellation (int32) and the barycentrics of the closest
point (3 x float32). The dense box is never materialised: a coarse-to-fine search (32^3 then 8^3 blocks,
pruned with the 1-Lipschitz property of φ) selects blocks that can meet the band, and the exact query runs
only on their cells, in chunks, one i-slab at a time, so peak memory is the final arrays plus one slab.

Grid: cell (i, j, k) has centre origin + h_g (i + ½, j + ½, k + ½); flat = (i·ny + j)·nz + k. The lattice is
anchored at the mesh bounding-box minimum, so `Band.extend` (§3.5, D_comp doubling) only shifts indices.

Interpolation (§3.6): d_grid = P d with P (n_cells x N_v, 3 non-zeros per row) from the closest point on the
fine tessellation to a triangle of the design surface Γ_h. Volume (§3.7, §3.10): V = h_g³ Σ H_ε(-ψ_c) in
float64, ε = h_g, quintic Heaviside.

Backends: NVIDIA Warp (`wp.Mesh`, `mesh_query_point_sign_winding_number`; GPU, or Warp's CPU device) and
libigl (`igl.signed_distance`, CPU fallback). Units mm.
"""
from __future__ import annotations

import hashlib
import math
import os
import time
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import scipy.sparse as sp

ROOT = Path(__file__).resolve().parents[1]
LEVELS = (32, 8)                 # coarse-to-fine block sizes (cells); LEVELS[0] is also the i-slab thickness
CHUNK = 1 << 23                  # query points per backend call (8.4 M)
_FAR = 1.0e8


# --------------------------------------------------------------------------------------------------------------
# Query backends: signed / unsigned distance, closest triangle and barycentrics for float64 points
# --------------------------------------------------------------------------------------------------------------
def _barycentric(P: np.ndarray, tri_v: np.ndarray) -> np.ndarray:
    """Barycentric coordinates (float64, (n, 3)) of the projection of P onto the planes of triangles tri_v (n, 3, 3)."""
    a, b, c = tri_v[:, 0], tri_v[:, 1], tri_v[:, 2]
    v0, v1, v2 = b - a, c - a, P - a
    d00, d01, d11 = (v0 * v0).sum(1), (v0 * v1).sum(1), (v1 * v1).sum(1)
    d20, d21 = (v2 * v0).sum(1), (v2 * v1).sum(1)
    den = d00 * d11 - d01 * d01
    den = np.where(np.abs(den) < 1e-300, 1e-300, den)
    v = (d11 * d20 - d01 * d21) / den
    w = (d00 * d21 - d01 * d20) / den
    return np.stack([1.0 - v - w, v, w], axis=1)


class _IglBackend:
    """libigl on the CPU: exact closest point, sign by winding number (default) or pseudonormal."""

    name = "igl"

    def __init__(self, V: np.ndarray, T: np.ndarray, sign: str = "winding_number"):
        import igl
        self.igl = igl
        self.V = np.ascontiguousarray(V, np.float64)
        self.T = np.ascontiguousarray(T, np.int64)
        self.sign_type = {"winding_number": igl.SIGNED_DISTANCE_TYPE_WINDING_NUMBER,
                          "pseudonormal": igl.SIGNED_DISTANCE_TYPE_PSEUDONORMAL,
                          "fast_winding_number": igl.SIGNED_DISTANCE_TYPE_FAST_WINDING_NUMBER}[sign]

    def _one(self, p: np.ndarray, signed: bool):
        st = self.sign_type if signed else self.igl.SIGNED_DISTANCE_TYPE_UNSIGNED
        S, I, C, _ = self.igl.signed_distance(np.ascontiguousarray(p, np.float64), self.V, self.T, st)
        return S.astype(np.float32), I.astype(np.int32), _barycentric(C, self.V[self.T[I]]).astype(np.float32)

    def query(self, P: np.ndarray, signed: bool = True, chunk: int = 1 << 16):
        """libigl holds the GIL, so chunks run in forked worker processes ($CSHELL_IGL_WORKERS, default all cores)."""
        global _IGL_SELF
        n = len(P)
        nw = int(os.environ.get("CSHELL_IGL_WORKERS", len(os.sched_getaffinity(0))))
        parts = [(s, min(n, s + chunk)) for s in range(0, n, chunk)]
        if nw <= 1 or len(parts) <= 1:
            res = [self._one(P[a:b], signed) for a, b in parts]
        else:
            import multiprocessing as mp
            from concurrent.futures import ProcessPoolExecutor
            _IGL_SELF = self
            with ProcessPoolExecutor(min(nw, len(parts)), mp_context=mp.get_context("fork")) as ex:
                res = list(ex.map(_igl_job, [(np.ascontiguousarray(P[a:b]), signed) for a, b in parts]))
        if not res:
            return np.empty(0, np.float32), np.empty(0, np.int32), np.empty((0, 3), np.float32)
        return (np.concatenate([r[0] for r in res]), np.concatenate([r[1] for r in res]),
                np.concatenate([r[2] for r in res]))


_IGL_SELF = None


def _igl_job(args):
    return _IGL_SELF._one(*args)


_WARP = {}


def _warp_kernels():
    """Define the Warp kernel lazily (import cost, and Warp needs the source in a file, which this module is)."""
    if _WARP:
        return _WARP
    import warp as wp
    wp.config.quiet = True
    if "SCRATCH" in os.environ and "WARP_CACHE_PATH" not in os.environ:
        wp.config.kernel_cache_dir = os.path.join(os.environ["SCRATCH"], "warp_cache")

    @wp.kernel
    def query_signed(mesh: wp.uint64, pts: wp.array(dtype=wp.vec3f), phi: wp.array(dtype=wp.float32),
                     face: wp.array(dtype=wp.int32), bary: wp.array(dtype=wp.vec3f)):
        t = wp.tid()
        p = pts[t]
        q = wp.mesh_query_point_sign_winding_number(mesh, p, 1.0e8, 2.0, 0.5)
        cp = wp.mesh_eval_position(mesh, q.face, q.u, q.v)
        phi[t] = q.sign * wp.length(p - cp)
        face[t] = q.face
        bary[t] = wp.vec3f(q.u, q.v, 1.0 - q.u - q.v)

    @wp.kernel
    def query_unsigned(mesh: wp.uint64, pts: wp.array(dtype=wp.vec3f), phi: wp.array(dtype=wp.float32),
                       face: wp.array(dtype=wp.int32), bary: wp.array(dtype=wp.vec3f)):
        t = wp.tid()
        p = pts[t]
        q = wp.mesh_query_point_no_sign(mesh, p, 1.0e8)
        cp = wp.mesh_eval_position(mesh, q.face, q.u, q.v)
        phi[t] = wp.length(p - cp)
        face[t] = q.face
        bary[t] = wp.vec3f(q.u, q.v, 1.0 - q.u - q.v)

    _WARP.update(wp=wp, signed=query_signed, unsigned=query_unsigned)
    return _WARP


class _WarpBackend:
    """NVIDIA Warp BVH queries (GPU by default). Mesh coordinates in float32 (callers pass a centred frame)."""

    name = "warp"

    def __init__(self, V: np.ndarray, T: np.ndarray, device: str | None = None):
        k = _warp_kernels()
        wp = self.wp = k["wp"]
        self.k = k
        self.device = device or os.environ.get("CSHELL_WARP_DEVICE") or ("cuda:0" if wp.get_cuda_device_count() else "cpu")
        self.points = wp.array(np.ascontiguousarray(V, np.float32), dtype=wp.vec3f, device=self.device)
        self.idx = wp.array(np.ascontiguousarray(T, np.int32).ravel(), dtype=wp.int32, device=self.device)
        self.mesh = wp.Mesh(self.points, self.idx, support_winding_number=True)

    def query(self, P: np.ndarray, signed: bool = True, chunk: int = CHUNK):
        wp, dev = self.wp, self.device
        n = len(P)
        phi = np.empty(n, np.float32)
        face = np.empty(n, np.int32)
        bary = np.empty((n, 3), np.float32)
        kern = self.k["signed" if signed else "unsigned"]
        for s in range(0, n, chunk):
            m = min(chunk, n - s)
            pts = wp.array(np.ascontiguousarray(P[s:s + m], np.float32), dtype=wp.vec3f, device=dev)
            o_phi = wp.empty(m, dtype=wp.float32, device=dev)
            o_face = wp.empty(m, dtype=wp.int32, device=dev)
            o_bary = wp.empty(m, dtype=wp.vec3f, device=dev)
            wp.launch(kern, dim=m, inputs=[self.mesh.id, pts, o_phi, o_face, o_bary], device=dev)
            phi[s:s + m] = o_phi.numpy()
            face[s:s + m] = o_face.numpy()
            bary[s:s + m] = o_bary.numpy()
        return phi, face, bary


def make_backend(V: np.ndarray, T: np.ndarray, backend: str = "auto", device: str | None = None):
    """'warp' (GPU if present, else Warp CPU), 'igl', or 'auto' (warp on CUDA, igl otherwise)."""
    if backend == "auto":
        try:
            backend = "warp" if _has_cuda() else "igl"
        except Exception:
            backend = "igl"
    if backend == "warp":
        return _WarpBackend(V, T, device)
    if backend in ("igl", "igl_pseudonormal"):
        return _IglBackend(V, T, "pseudonormal" if backend.endswith("pseudonormal") else "winding_number")
    raise ValueError(backend)


def _has_cuda() -> bool:
    import warp as wp
    wp.config.quiet = True
    wp.init()
    return wp.get_cuda_device_count() > 0


# --------------------------------------------------------------------------------------------------------------
# Grid geometry and the coarse-to-fine band search
# --------------------------------------------------------------------------------------------------------------
@dataclass(frozen=True)
class Grid:
    """Bounding grid: cell (i,j,k) centre = origin + h (i+½, j+½, k+½) (global mm); `n_pad` cells of margin."""
    origin: np.ndarray
    h: float
    shape: tuple
    n_pad: int

    @property
    def nyz(self) -> int:
        return self.shape[1] * self.shape[2]

    def flat(self, i, j, k):
        return (np.asarray(i, np.int64) * self.shape[1] + j) * self.shape[2] + k

    def unravel(self, f):
        f = np.asarray(f, np.int64)
        k = f % self.shape[2]
        ij = f // self.shape[2]
        return ij // self.shape[1], ij % self.shape[1], k


def _pad(h: float, D_comp: float) -> int:
    return int(math.ceil((D_comp + h) / h)) + 1


def make_grid(V: np.ndarray, h: float, D_comp: float) -> Grid:
    bmin, bmax = V.min(0), V.max(0)
    npad = _pad(h, D_comp)
    shape = tuple(int(x) for x in np.ceil((bmax - bmin) / h).astype(np.int64) + 2 * npad + 1)
    return Grid(bmin - npad * h, float(h), shape, npad)


def _in_grid_count(corners: np.ndarray, B: int, shape) -> int:
    """Number of grid cells covered by the blocks [c, c+B)³ (clipped to the grid)."""
    ext = np.clip(np.asarray(shape)[None, :] - corners, 0, B).astype(np.int64)
    return int(ext.prod(1).sum())


def _block_centres(corners: np.ndarray, B: int, grid: Grid, c0: np.ndarray) -> np.ndarray:
    """Block-centre points (local frame) of blocks of B cells starting at index `corners`."""
    return grid.origin - c0 + grid.h * (corners + 0.5 * B)


def _fill(be, grid: Grid, c0: np.ndarray, lo: float, hi: float, lo_open: bool, count_deep: bool,
          levels=LEVELS, chunk: int = CHUNK, log=None):
    """Active cells with lo <= φ <= hi (lo < φ if lo_open), sorted by flat index.

    Returns (flat int64, phi f32, tri i32, bary f32 (n,3), n_deep) where n_deep counts the in-grid cells with
    φ < lo that are not stored (their material fraction is 1 for ψ = max(φ - d, -φ - a) with a <= r_in, §3.10).
    Pruning: |φ(x) - φ(c)| <= |x - c| <= (√3/2) B h for any x in a block of B cells with centre c.
    """
    B1, B0 = levels
    h, shape = grid.h, np.asarray(grid.shape, np.int64)
    n_deep = 0
    R = B1 // B0
    off0 = np.stack(np.meshgrid(*[np.arange(B0)] * 3, indexing="ij"), -1).reshape(-1, 3)
    offR = np.stack(np.meshgrid(*[np.arange(R)] * 3, indexing="ij"), -1).reshape(-1, 3) * B0

    def classify(corners, B):
        pc = _block_centres(corners, B, grid, c0)
        phi_c, _, _ = be.query(pc, signed=True)
        rb = 0.5 * math.sqrt(3.0) * B * h + 1e-3 * h
        keep = (phi_c.astype(np.float64) + rb >= lo) & (phi_c.astype(np.float64) - rb <= hi)
        deep = phi_c.astype(np.float64) + rb < lo
        return keep, deep

    nb = -(-shape // B1)
    # top level: all coarse blocks, ordered by block-i first
    bi = np.arange(nb[0], dtype=np.int64)
    slabs = []
    t0 = time.time()
    # classify per i-slab to bound memory for huge grids
    for b in bi:
        bj, bk = np.meshgrid(np.arange(nb[1]), np.arange(nb[2]), indexing="ij")
        c1 = np.stack([np.full(bj.size, b), bj.ravel(), bk.ravel()], 1).astype(np.int64) * B1
        keep, deep = classify(c1, B1)
        if count_deep and deep.any():
            n_deep += _in_grid_count(c1[deep], B1, shape)
        c1 = c1[keep]
        if len(c1) == 0:
            continue
        ch = (c1[:, None, :] + offR[None]).reshape(-1, 3)
        ch = ch[(ch < shape[None, :]).all(1)]
        keep0, deep0 = classify(ch, B0)
        if count_deep and deep0.any():
            n_deep += _in_grid_count(ch[deep0], B0, shape)
        ch = ch[keep0]
        if len(ch) == 0:
            continue
        per = max(1, chunk // B0 ** 3)
        fl, ph, tr, ba = [], [], [], []
        for s in range(0, len(ch), per):
            cc = ch[s:s + per]
            ijk = (cc[:, None, :] + off0[None]).reshape(-1, 3)
            ok = (ijk < shape[None, :]).all(1)
            ijk = ijk[ok]
            P = grid.origin - c0 + h * (ijk + 0.5)
            phi, face, bary = be.query(P, signed=True)
            p64 = phi.astype(np.float64)
            if count_deep:
                n_deep += int((p64 < lo).sum())
            sel = ((p64 > lo) if lo_open else (p64 >= lo)) & (p64 <= hi)
            ijk = ijk[sel]
            fl.append((ijk[:, 0] * shape[1] + ijk[:, 1]) * shape[2] + ijk[:, 2])
            ph.append(phi[sel]); tr.append(face[sel]); ba.append(bary[sel])
        if not fl:
            continue
        f = np.concatenate(fl)
        if len(f) == 0:
            continue
        o = np.argsort(f, kind="stable")
        slabs.append((f[o], np.concatenate(ph)[o], np.concatenate(tr)[o], np.concatenate(ba)[o]))
        if log:
            log(f"  slab {b + 1}/{nb[0]}: {len(f)} cells  ({time.time() - t0:.1f} s)")
    n = sum(len(s[0]) for s in slabs)
    flat = np.empty(n, np.int64); phi = np.empty(n, np.float32)
    tri = np.empty(n, np.int32); bary = np.empty((n, 3), np.float32)
    p = 0
    while slabs:                      # copy and free slab by slab: peak = final arrays + one slab
        f, ph, tr, ba = slabs.pop(0)
        m = len(f)
        flat[p:p + m], phi[p:p + m], tri[p:p + m], bary[p:p + m] = f, ph, tr, ba
        p += m
    return flat, phi, tri, bary, n_deep


# --------------------------------------------------------------------------------------------------------------
# Smoothed Heaviside and the Band
# --------------------------------------------------------------------------------------------------------------
def heaviside_quintic(s: np.ndarray, eps: float) -> np.ndarray:
    """C² quintic H_ε(s) = ½ + 15/16 (s/ε) - 5/8 (s/ε)³ + 3/16 (s/ε)⁵ on |s| < ε, else 0/1 (§3.7). float64."""
    x = np.clip(np.asarray(s, np.float64) / eps, -1.0, 1.0)
    return 0.5 + x * (15.0 / 16.0 + x * x * (-5.0 / 8.0 + 3.0 / 16.0 * x * x))


def P_from_compact(face: np.ndarray, bary: np.ndarray, Th: np.ndarray, n_v: int, dtype=np.float64) -> sp.csr_matrix:
    """CSR P from (closest Γ_h triangle, raw barycentrics): β ← max(β, 0) / Σ max(β, 0) in float64 (§3.6)."""
    n = len(face)
    b = np.clip(bary.astype(np.float64), 0.0, None)
    b /= b.sum(1, keepdims=True)
    indptr = np.arange(0, 3 * n + 1, 3, dtype=np.int64 if 3 * n >= 2**31 - 1 else np.int32)
    return sp.csr_matrix((b.astype(dtype, copy=False).ravel(), Th[face].ravel(), indptr), shape=(n, n_v))


@dataclass
class Band:
    """Sparse φ band (§3.1). Arrays are indexed by band cell (sorted by `flat`)."""
    grid: Grid
    D_comp: float
    r_in: float
    c0: np.ndarray                 # centring offset of the mesh used for Warp float32 queries (global mm)
    V: np.ndarray                  # fine tessellation vertices (global mm, float64)
    T: np.ndarray                  # fine tessellation triangles (int32)
    flat: np.ndarray               # (n,) int64
    phi: np.ndarray                # (n,) float32, mm, negative inside
    tri: np.ndarray                # (n,) int32 closest triangle of (V, T)
    bary: np.ndarray               # (n, 3) float32 barycentrics of the closest point
    n_deep: int = 0                # in-grid cells deeper than the band (fully material), for V (§3.10)
    info: dict = field(default_factory=dict)
    _be: object = field(default=None, repr=False, compare=False)
    levels: tuple = LEVELS

    # ---- construction ----------------------------------------------------------------------------------------
    @classmethod
    def build(cls, V: np.ndarray, T: np.ndarray, h_g: float, D_comp: float = 4.0, r_in: float = math.inf,
              backend: str = "auto", device: str | None = None, chunk: int = CHUNK, levels=LEVELS,
              verbose: bool = False) -> "Band":
        """Build the band for the watertight mesh (V, T) (mm). r_in = inf keeps the whole interior (§3.1)."""
        V = np.ascontiguousarray(V, np.float64)
        T = np.ascontiguousarray(T, np.int32)
        c0 = np.round(0.5 * (V.min(0) + V.max(0)), 3)
        be = make_backend(V - c0, T, backend, device)
        grid = make_grid(V, h_g, D_comp)
        lo = -math.inf if math.isinf(r_in) else -r_in - h_g
        t0 = time.time()
        flat, phi, tri, bary, n_deep = _fill(be, grid, c0, lo, D_comp + h_g, False, not math.isinf(lo), levels=levels,
                                             chunk=chunk, log=print if verbose else None)
        b = cls(grid, float(D_comp), float(r_in), c0, V, T, flat, phi, tri, bary, n_deep, _be=be)
        b.levels = tuple(levels)
        b.info = dict(seconds=time.time() - t0, backend=be.name, n_cells=len(flat))
        return b

    def backend(self, backend: str = "auto", device: str | None = None):
        if self._be is None:
            self._be = make_backend(self.V - self.c0, self.T, backend, device)
        return self._be

    # ---- basic accessors -------------------------------------------------------------------------------------
    @property
    def n(self) -> int:
        return len(self.flat)

    @property
    def h(self) -> float:
        return self.grid.h

    @property
    def nbytes(self) -> int:
        return self.flat.nbytes + self.phi.nbytes + self.tri.nbytes + self.bary.nbytes

    def ijk(self, idx=slice(None)) -> np.ndarray:
        i, j, k = self.grid.unravel(self.flat[idx])
        return np.stack([i, j, k], 1)

    def centres(self, idx=slice(None)) -> np.ndarray:
        """Cell-centre coordinates (global mm, float64, (m, 3))."""
        return self.grid.origin + self.h * (self.ijk(idx) + 0.5)

    def closest_points(self, idx=slice(None)) -> np.ndarray:
        """π(x): closest point on the fine tessellation, from the stored triangle and barycentrics (float64)."""
        b = self.bary[idx].astype(np.float64)
        b = np.clip(b, 0.0, None)
        b /= b.sum(1, keepdims=True)
        return np.einsum("ni,nij->nj", b, self.V[self.T[self.tri[idx]]])

    def to_dense(self, values: np.ndarray, fill: float = np.nan) -> np.ndarray:
        """Scatter band values to a dense array (tests / small grids only)."""
        out = np.full(self.grid.shape, fill, dtype=np.asarray(values).dtype)
        out.reshape(-1)[self.flat] = values
        return out

    # ---- §3.6 interpolation matrix --------------------------------------------------------------------------
    def P(self, Vh: np.ndarray, Th: np.ndarray, dtype=np.float64, chunk: int = CHUNK, backend: str = "auto",
          device: str | None = None) -> sp.csr_matrix:
        """Sparse (n_cells x N_v) matrix, 3 non-zeros per row: x_c -> π(x_c) on the fine tessellation -> closest
        triangle t' of Γ_h and barycentrics β' (§3.6). Rows sum to 1; a linear vertex field is reproduced exactly
        at the located point. Vh (N_v, 3) mm, Th (N_t, 3)."""
        face, bary = self.P_compact(Vh, Th, chunk, backend, device)
        return P_from_compact(face, bary, np.ascontiguousarray(Th, np.int32), len(Vh), dtype)

    def P_compact(self, Vh: np.ndarray, Th: np.ndarray, chunk: int = CHUNK, backend: str = "auto",
                  device: str | None = None):
        """Compact form of P (§3.6): closest triangle of Γ_h per cell (int32) and the raw float32 barycentrics of the
        located point. `P_from_compact` (CPU) and design.DeviceDisc (GPU) clip and normalise them identically."""
        Vh = np.ascontiguousarray(Vh, np.float64)
        Th = np.ascontiguousarray(Th, np.int32)
        be = make_backend(Vh - self.c0, Th, backend, device)
        n = self.n
        face = np.empty(n, np.int32)
        bary = np.empty((n, 3), np.float32)
        for s in range(0, n, chunk):
            sl = slice(s, min(n, s + chunk))
            pi = self.closest_points(sl) - self.c0
            _, face[sl], bary[sl] = be.query(pi, signed=False)
        return face, bary

    # ---- §3.5 D_comp doubling -------------------------------------------------------------------------------
    def extend(self, D_new: float, chunk: int = CHUNK, verbose: bool = False) -> "Band":
        """Band with outer extent D_new > D_comp: existing cells kept, shell cells D_comp+h < φ <= D_new+h added.
        The grid grows by whole cells (indices are re-based, order preserved). Returns a new Band."""
        if D_new <= self.D_comp:
            return self
        h = self.h
        npad = _pad(h, D_new)
        m = npad - self.grid.n_pad
        g = self.grid
        grid = Grid(g.origin - m * h, h, tuple(s + 2 * m for s in g.shape), npad)
        i, j, k = g.unravel(self.flat)
        old_flat = grid.flat(i + m, j + m, k + m)
        t0 = time.time()
        f, ph, tr, ba, _ = _fill(self.backend(), grid, self.c0, self.D_comp + h, D_new + h, True, False, levels=self.levels,
                                 chunk=chunk, log=print if verbose else None)
        # merge two sorted index lists without sorting: position = own rank + rank in the other list
        pos_old = np.arange(len(old_flat)) + np.searchsorted(f, old_flat)
        pos_new = np.arange(len(f)) + np.searchsorted(old_flat, f)
        n = len(old_flat) + len(f)
        flat = np.empty(n, np.int64); phi = np.empty(n, np.float32)
        tri = np.empty(n, np.int32); bary = np.empty((n, 3), np.float32)
        for dst, a, b in ((flat, old_flat, f), (phi, self.phi, ph), (tri, self.tri, tr), (bary, self.bary, ba)):
            dst[pos_old] = a
            dst[pos_new] = b
        out = Band(grid, float(D_new), self.r_in, self.c0, self.V, self.T, flat, phi, tri, bary, self.n_deep,
                   _be=self._be, levels=self.levels)
        out.info = dict(seconds=time.time() - t0, backend=self.info.get("backend"), n_cells=n, extended_from=self.D_comp)
        return out

    # ---- §3.7, §3.10 volume ---------------------------------------------------------------------------------
    def chi(self, psi: np.ndarray, eps: float | None = None) -> np.ndarray:
        """χ_c = H_ε(-ψ_c), ε = h_g (§3.7); float64."""
        return heaviside_quintic(-np.asarray(psi, np.float64), self.h if eps is None else eps)

    def volume(self, psi: np.ndarray, eps: float | None = None) -> float:
        """V = h_g³ (Σ_band χ + n_deep) in float64 (§3.10). n_deep = 0 when r_in = inf."""
        return float(self.h ** 3 * (math.fsum(self.chi(psi, eps)) + self.n_deep))

    def V0(self, eps: float | None = None) -> float:
        """Volume of the solid on the geometry grid: ψ = φ, ε = h_g (§3.7, mass target M₀ = ρ V₀)."""
        return self.volume(self.phi, eps)

    # ---- cache (.npz) ---------------------------------------------------------------------------------------
    def save(self, path, key: dict | None = None) -> None:
        g = self.grid
        meta = dict(key or {}, h=g.h, D_comp=self.D_comp, r_in=self.r_in, n_pad=g.n_pad, n_deep=self.n_deep)
        tmp = str(path) + ".tmp.npz"
        np.savez(tmp, flat=self.flat, phi=self.phi, tri=self.tri, bary=self.bary, V=self.V, T=self.T,
                 origin=g.origin, shape=np.asarray(g.shape, np.int64), c0=self.c0, meta=np.array(repr(meta)))
        os.replace(tmp, path)

    @classmethod
    def load(cls, path) -> "Band":
        z = np.load(path, allow_pickle=False)
        meta = eval(str(z["meta"]), {"inf": math.inf, "nan": math.nan})   # repr of a dict of floats/ints/str
        grid = Grid(z["origin"], float(meta["h"]), tuple(int(s) for s in z["shape"]), int(meta["n_pad"]))
        return cls(grid, float(meta["D_comp"]), float(meta["r_in"]), z["c0"], z["V"], z["T"], z["flat"], z["phi"],
                   z["tri"], z["bary"], int(meta["n_deep"]))


# --------------------------------------------------------------------------------------------------------------
# Cache keyed by STEP SHA-256 and grid parameters; STEP entry point
# --------------------------------------------------------------------------------------------------------------
def cache_dir() -> Path:
    """$CSHELL_CACHE, default <repo>/.cache_cshell (a symlink to projects storage)."""
    return Path(os.environ.get("CSHELL_CACHE", ROOT / ".cache_cshell"))


def step_sha256(path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for blk in iter(lambda: f.read(1 << 20), b""):
            h.update(blk)
    return h.hexdigest()


def cache_path(sha: str, h_g: float, D_comp: float, tol: float, r_in: float = math.inf) -> Path:
    return cache_dir() / f"band_{sha[:16]}_h{h_g:g}_D{D_comp:g}_tol{tol:g}_r{r_in:g}.npz"


def build_from_step(step_path, h_g: float, D_comp: float = 4.0, tol: float = 0.02, r_in: float = math.inf,
                    use_cache: bool = True, **kw) -> Band:
    """STEP -> tessellation (`dshell.geometry`, chordal tol mm) -> Band, cached to .npz (§3.1)."""
    sha = step_sha256(step_path)
    path = cache_path(sha, h_g, D_comp, tol, r_in)
    if use_cache and path.exists():
        b = Band.load(path)
        b.info = dict(cached=True, n_cells=b.n)
        return b
    from dshell.geometry import load_step, tessellate
    V, T, _ = tessellate(load_step(str(step_path)), lin=tol)
    b = Band.build(V, T, h_g, D_comp, r_in, **kw)
    if use_cache:
        path.parent.mkdir(parents=True, exist_ok=True)
        b.save(path, dict(sha256=sha, tol=tol))
    return b
