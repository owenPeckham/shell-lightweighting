"""Export of the final design (THEORY §8.1-8.5): tiled nominal ψ on an export grid, clean-up, escape holes, mass
correction, marching cubes, mesh checks and wall-thickness statistics.

Material set (nominal, §3.2):  ψ(x) = max(φ(x) − d(π(x)), −φ(x) − a(π(x))),  Ω = {ψ ≤ 0}, on cell centres of a grid
of spacing h_x. φ is the exact signed distance to Γ₀ (igl, pseudonormal sign), π the closest point on Γ₀ and d, a the
vertex fields of Γ_h interpolated at the closest point of Γ_h to π(x) (§3.6). The full grid is never held as floats:
the grid is cut into tiles of n³ cells (default n = 128) processed by forked workers.

Per tile and per cell the pipeline keeps a uint8 code (0 material, 1 void, 2 outside, 3 dead material island) and,
only for cells near a boundary (|φ − d| ≤ Δ + 2h_x or |ψ_in| ≤ 2h_x, with ψ_in = −φ − a the void branch), the table
S = (φ, Γ_h triangle, barycentrics). Everything after the first pass (opening, labelling, escape holes, mass
correction, marching cubes) works from the codes and S, so the closest-point queries are done once.

Stages (`Exporter.run`)
  A  geometry pass: codes, S, V₀ (= solid volume on the export grid, linear-ramp fractions);
  B  clean-up (§8.2): void mask opened with a ball of diameter v_min (EDT, halo 2r), enclosed voids < void_min_mm³
     filled, largest material component kept; the opening rule is re-checked (re-opening changes 0 cells);
  C  escape holes (§8.3): sites = vertices of Γ_h whose void wall point p = x − a n lies on an enclosed void with
     nominal wall a + d ≥ t_min + b_v and > mount_clear_mm from mounts; lowest 30 % stress (or lowest 30 % a, a proxy
     for "depth" when no stress is given), k sites maximising separation along the void's principal axis (PCA of the
     void-wall probe points); cylinder ψ_hole = min(D/2 − ‖(x−p)×n‖, s − s₀, s₁ − s), ψ ← max(ψ, ψ_hole);
  D  mass correction (§8.4): δ by Brent so V(δ) = V₀ with d ← clip(d + δ, 0, d^hi) on active vertices; the void
     mask does not depend on d (void ⊂ φ < 0), so only the outer shell is re-evaluated;
  E  marching cubes per tile (skimage), welded by lattice-edge keys, own watertight / manifold / orientation checks
     (trimesh cross-check for small meshes), volume, STL; wall-thickness histogram by inward ray casting.

Volumes use the linear ramp f = clip(½ − ψ/h_x, 0, 1) (the C⁰ case of §3.7), which is volume-exact to O(h²κ) and
agrees with the marching-cubes mesh volume. Units mm.
"""
from __future__ import annotations

import csv
import json
import math
import multiprocessing as mp
import os
import resource
import time
from dataclasses import dataclass, field, asdict

import numpy as np
from scipy import ndimage, sparse
from scipy.optimize import brentq
from scipy.sparse.csgraph import connected_components
from scipy.spatial import cKDTree

A_CLOSED = 1.0e4            # inset used for exactly closed vertices (a = ∞, §8.2 step 1); interpolation keeps it closed
SIGN_TYPES = ("pseudonormal", "winding_number", "fast_winding_number")


# =====================================================================================================================
# configuration
# =====================================================================================================================
@dataclass
class ExportConfig:
    h_x: float = 0.05                 # export grid spacing (mm), grid.export_h_mm
    tile: int = 128                   # tile edge in cells (multiple of `block`)
    block: int = 16                   # pruning block edge in cells
    v_min: float = 3.0                # void_min_mm: opening ball diameter
    void_min_mm3: float = 30.0        # enclosed voids below this are filled
    t_min: float = 1.5
    b_v: float = 0.2
    holes_per_void: int = 2
    hole_d: float = 3.0
    hole_stress_frac: float = 0.30
    mount_clear: float = 3.0
    closed_thr: float = 0.98          # ξ̃ᵃ ≥ closed_thr: exactly closed
    delta_shell: float = 0.15         # |δ| covered by the cached outer shell; retried once at delta_retry
    delta_retry: float = 0.35
    delta_warn: float = 0.10
    delta_fail: float = 0.30
    mass_tol: float = 1.0e-5          # Brent tolerance on |V/V₀ − 1|
    vol_tol: float = 2.0e-3           # R1
    sign: str = "pseudonormal"
    workers: int | None = None
    open_workers: int = 24            # EDT blocks are large: cap their parallelism (memory)
    clean: bool = True
    holes: bool = True
    check_opening: bool = True
    wall_hist: bool = True
    wall_samples: int = 20000
    wall_max_faces: int = 4_000_000   # ray-cast mesh is extracted at stride ≥ 1 so that it stays below this
    write_stl: bool = True
    trimesh_check_max_faces: int = 2_000_000
    log: object = print

    @classmethod
    def from_run_cfg(cls, cfg: dict, **kw) -> "ExportConfig":
        m = cfg["manufacturing"]
        base = dict(h_x=cfg["grid"]["export_h_mm"], v_min=m["void_min_mm"], void_min_mm3=m["void_min_mm3"],
                    t_min=m["t_min_mm"], b_v=m["wall_bias_mm"], holes_per_void=m["escape_holes_per_void"],
                    hole_d=m["escape_hole_diameter_mm"])
        base.update(kw)
        return cls(**base)


def ramp(s, h):
    """Volume fraction clip(½ + s/h, 0, 1) of a cell of edge h whose centre has signed depth s inside (s = −ψ)."""
    return np.clip(0.5 + s / h, 0.0, 1.0)


# =====================================================================================================================
# morphological opening (§8.2)
# =====================================================================================================================
def ball_r2(r_mm: float, h: float) -> float:
    """Squared radius in cells of the structuring ball (lattice points z with |z|² ≤ R2)."""
    return math.floor((r_mm / h) ** 2 + 1e-9)


def _d2(mask_zero_is_bg: np.ndarray) -> np.ndarray:
    d = ndimage.distance_transform_edt(mask_zero_is_bg)
    return np.rint(d * d)


def morph_open(mask: np.ndarray, R2: float) -> np.ndarray:
    """Opening of a boolean mask by the lattice ball {|z|² ≤ R2}: union of all balls contained in the mask.

    E = {c : every non-mask cell is farther than r}; O = {x : some c ∈ E within r}. O ⊂ mask, and opening is
    idempotent (an exact property of the lattice ball, so the §8.2 re-opening check must give 0 changed cells).
    Cells outside the array count as non-mask (callers supply the halo).
    """
    if not mask.any():
        return mask.copy()
    E = _d2(mask) > R2
    if not E.any():
        return np.zeros_like(mask)
    return _d2(~E) <= R2


# =====================================================================================================================
# geometry source
# =====================================================================================================================
def _barycentric(P, tri_v):
    a, b, c = tri_v[:, 0], tri_v[:, 1], tri_v[:, 2]
    v0, v1, v2 = b - a, c - a, P - a
    d00, d01, d11 = (v0 * v0).sum(1), (v0 * v1).sum(1), (v1 * v1).sum(1)
    d20, d21 = (v2 * v0).sum(1), (v2 * v1).sum(1)
    den = d00 * d11 - d01 * d01
    den = np.where(np.abs(den) < 1e-300, 1e-300, den)
    v = (d11 * d20 - d01 * d21) / den
    w = (d00 * d21 - d01 * d20) / den
    bary = np.clip(np.stack([1.0 - v - w, v, w], axis=1), 0.0, None)
    return bary / bary.sum(1, keepdims=True)


class Design:
    """Nominal two-field design on Γ₀ (V0, T0) and Γ_h (Vh, Th), or a constant (d, a) design when Vh is None.

    d, a (N_v,) mm; closed (bool): exactly closed vertices (a := A_CLOSED); d_hi (N_v,); active (bool): vertices
    that carry the mass correction δ. Constant design: d_const, a_const (a_const ≥ A_CLOSED/2 means no void) and
    δ shifts d for every cell.
    """

    def __init__(self, V0, T0, Vh=None, Th=None, d=None, a=None, d_hi=None, active=None, closed=None,
                 d_const: float = 0.0, a_const: float = A_CLOSED, sign: str = "pseudonormal", d_extent: float | None = None):
        import igl
        self._igl = igl
        self.V0 = np.ascontiguousarray(V0, np.float64)
        self.T0 = np.ascontiguousarray(T0, np.int64)
        self.sign_type = {"pseudonormal": igl.SIGNED_DISTANCE_TYPE_PSEUDONORMAL,
                          "winding_number": igl.SIGNED_DISTANCE_TYPE_WINDING_NUMBER,
                          "fast_winding_number": igl.SIGNED_DISTANCE_TYPE_FAST_WINDING_NUMBER}[sign]
        self.Vh = None if Vh is None else np.ascontiguousarray(Vh, np.float64)
        self.Th = None if Th is None else np.ascontiguousarray(Th, np.int64)
        self.d_const, self.a_const = float(d_const), float(a_const)
        if self.Vh is None:
            self.d, self.a, self.d_hi, self.active = None, None, None, None
            self.d_out_max = self.d_const
            self.has_void = self.a_const < A_CLOSED / 2
            self.a_open_max = self.a_const
            return
        nv = len(self.Vh)
        closed = np.zeros(nv, bool) if closed is None else np.asarray(closed, bool)
        self.closed = closed
        self.d = np.asarray(d, np.float64).copy()
        self.a = np.where(closed, A_CLOSED, np.asarray(a, np.float64))
        self.d_hi = np.asarray(d_hi, np.float64)
        self.active = np.ones(nv, bool) if active is None else np.asarray(active, bool)
        self.d_out_max = float(np.max(np.where(self.active, self.d_hi, self.d)) if nv else 0.0)
        if d_extent is not None:
            self.d_out_max = float(d_extent)
        self.has_void = bool((~closed).any())
        self.a_open_max = float(self.a[~closed].max()) if self.has_void else 0.0
        e = np.concatenate([self.Vh[self.Th[:, 1]] - self.Vh[self.Th[:, 0]], self.Vh[self.Th[:, 2]] - self.Vh[self.Th[:, 1]],
                            self.Vh[self.Th[:, 0]] - self.Vh[self.Th[:, 2]]])
        self.reach = 3.0 * float(np.linalg.norm(e, axis=1).max()) + 0.05
        if not self.has_void and not np.any(self.d > 0):      # ψ = φ exactly: no Γ_h needed
            self.Vh = self.Th = self.d = self.a = self.d_hi = self.active = None
            self.d_const, self.a_const, self.d_out_max = 0.0, A_CLOSED, 0.0
            return
        self._tree_open = cKDTree(self.Vh[~closed]) if self.has_void else None
        self._tree_closed = cKDTree(self.Vh[closed]) if closed.any() else None
        self._aabb_h = None

    # -- queries ----------------------------------------------------------------------------------------------------
    @property
    def trivial_solid(self) -> bool:
        """No void anywhere and d ≡ 0 with no vertex able to move: ψ = φ exactly (V = V₀ with δ = 0)."""
        if self.Vh is None:
            return (not self.has_void) and self.d_const == 0.0
        return (not self.has_void) and not np.any(self.d > 0)

    def phi(self, X: np.ndarray) -> np.ndarray:
        return self._igl.signed_distance(np.ascontiguousarray(X, np.float64), self.V0, self.T0, self.sign_type)[0]

    def phi_closest(self, X: np.ndarray):
        S, _, C, _ = self._igl.signed_distance(np.ascontiguousarray(X, np.float64), self.V0, self.T0, self.sign_type)
        return S, C

    def locate(self, C: np.ndarray):
        """Closest triangle of Γ_h and barycentrics for points C on Γ₀ (§3.6)."""
        if self._aabb_h is None:
            self._aabb_h = self._igl.AABB()
            self._aabb_h.init(self.Vh, self.Th)
        _, tri, Cp = self._aabb_h.squared_distance(self.Vh, self.Th, np.ascontiguousarray(C, np.float64))
        b = _barycentric(Cp, self.Vh[self.Th[tri]])
        return tri.astype(np.int32), b[:, 1].astype(np.float32), b[:, 2].astype(np.float32)

    def a_at(self, tri, b1, b2):
        if self.Vh is None:
            return np.full(len(b1), self.a_const)
        v = self.Th[tri]
        b1, b2 = b1.astype(np.float64), b2.astype(np.float64)
        return (1.0 - b1 - b2) * self.a[v[:, 0]] + b1 * self.a[v[:, 1]] + b2 * self.a[v[:, 2]]

    def d_at(self, tri, b1, b2, delta: float = 0.0):
        if self.Vh is None:
            return np.full(len(b1), max(self.d_const + delta, 0.0))
        v = self.Th[tri]
        dv = np.clip(self.d + delta * self.active, 0.0, np.maximum(self.d_hi, self.d))
        b1, b2 = b1.astype(np.float64), b2.astype(np.float64)
        return (1.0 - b1 - b2) * dv[v[:, 0]] + b1 * dv[v[:, 1]] + b2 * dv[v[:, 2]]

    def void_class(self, centres, phi_c, R, h):
        """Block classification for deep blocks: 0 no void possible, 1 certainly void, 2 unknown."""
        cls = np.full(len(centres), 2, np.int8)
        if not self.has_void:
            cls[:] = 0
            return cls
        if self.Vh is None:
            cls[(-phi_c - R - self.a_const) > 2.0 * h] = 1
            return cls
        rho = np.abs(phi_c) + R + self.reach
        dO = self._tree_open.query(centres)[0]
        cls[dO > rho] = 0
        if self._tree_closed is None:
            dC = np.full(len(centres), np.inf)
        else:
            dC = self._tree_closed.query(centres)[0]
        cls[(dC > rho) & ((-phi_c - R - self.a_open_max) > 2.0 * h) & (cls != 0)] = 1
        return cls


# =====================================================================================================================
# grid and tile store
# =====================================================================================================================
class Grid:
    def __init__(self, lo, hi, h, n):
        self.h, self.n = float(h), int(n)
        ext = np.ceil((np.asarray(hi) - np.asarray(lo)) / h).astype(int) + 2
        self.nt = np.ceil(ext / n).astype(int)
        self.shape = self.nt * n
        # generic sub-cell offset: CAD faces at 'nice' coordinates must not coincide with cell-centre planes (ψ = 0 there)
        self.origin = np.asarray(lo, np.float64) - 0.5 * h * (self.shape - ext) + h * np.array([0.37, 0.41, 0.29])

    def tile_ids(self):
        return [(i, j, k) for i in range(self.nt[0]) for j in range(self.nt[1]) for k in range(self.nt[2])]

    def tile_lo(self, tid):
        return self.origin + self.n * self.h * np.asarray(tid)

    def centres(self, tid, flat_idx=None):
        n, h = self.n, self.h
        lo = self.tile_lo(tid)
        if flat_idx is None:
            ax = [lo[k] + (np.arange(n) + 0.5) * h for k in range(3)]
            return np.stack(np.meshgrid(*ax, indexing="ij"), -1).reshape(-1, 3)
        i = np.unravel_index(flat_idx, (n, n, n))
        return lo + (np.stack(i, -1) + 0.5) * h

    def cell_of(self, X):
        return np.floor((np.asarray(X) - self.origin) / self.h).astype(np.int64)


class TileStore:
    """Per-tile uint8 codes; uniform tiles are kept as Python ints."""

    def __init__(self, grid: Grid):
        self.grid, self.codes, self.S = grid, {}, {}

    def get(self, tid):
        c = self.codes[tid]
        if isinstance(c, (int, np.integer)):
            return np.full((self.grid.n,) * 3, c, np.uint8)
        return c

    def scalar(self, tid):
        c = self.codes.get(tid, 2)
        return int(c) if isinstance(c, (int, np.integer)) else None

    def set(self, tid, arr):
        if isinstance(arr, np.ndarray):
            u = arr.flat[0]
            if (arr == u).all():
                arr = int(u)
        self.codes[tid] = arr

    def block(self, tid, halo: int, fill: int = 2):
        """Codes of tile `tid` with a halo of `halo` cells from the neighbours (outside the grid: `fill`)."""
        n, g = self.grid.n, self.grid
        m = -(-halo // n)
        out = np.full((n + 2 * halo,) * 3, fill, np.uint8)
        for di in range(-m, m + 1):
            for dj in range(-m, m + 1):
                for dk in range(-m, m + 1):
                    t = (tid[0] + di, tid[1] + dj, tid[2] + dk)
                    if min(t) < 0 or any(t[a] >= g.nt[a] for a in range(3)):
                        continue
                    sl_o, sl_i = [], []
                    for a, dd in enumerate((di, dj, dk)):
                        o0 = halo + dd * n                      # start of neighbour tile in block coordinates
                        lo, hi = max(o0, 0), min(o0 + n, n + 2 * halo)
                        sl_o.append(slice(lo, hi))
                        sl_i.append(slice(lo - o0, hi - o0))
                    c = self.codes.get(t, 2)
                    out[tuple(sl_o)] = c if isinstance(c, (int, np.integer)) else c[tuple(sl_i)]
        return out


_G: dict = {}


def _pmap(fn, items, workers):
    """Map over tile ids with forked workers; fn reads module global _G (inherited at fork). A dead worker raises
    BrokenProcessPool instead of hanging."""
    items = list(items)
    if workers <= 1 or len(items) <= 1:
        return [fn(i) for i in items]
    from concurrent.futures import ProcessPoolExecutor
    with ProcessPoolExecutor(min(workers, len(items)), mp_context=mp.get_context("fork")) as ex:
        return list(ex.map(fn, items, chunksize=1))


# =====================================================================================================================
# stage A: geometry pass
# =====================================================================================================================
def _pass_a(tid):
    dz: Design = _G["design"]
    cfg: ExportConfig = _G["cfg"]
    grid: Grid = _G["grid"]
    n, h, b = grid.n, grid.h, cfg.block
    nb = n // b
    S_thr = _G["delta"] + 2.0 * h
    band_out = dz.d_out_max + S_thr
    lo = grid.tile_lo(tid)
    ax = [lo[k] + (np.arange(nb) + 0.5) * b * h for k in range(3)]
    bc = np.stack(np.meshgrid(*ax, indexing="ij"), -1).reshape(-1, 3)
    phi_c = dz.phi(bc)
    R = 0.5 * b * h * math.sqrt(3.0)
    far_out = phi_c - R > band_out
    blk = np.full(len(bc), 255, np.uint8)
    blk[far_out] = 2
    deep = phi_c + R < -S_thr
    v0 = 0.0
    if deep.any():
        cls = dz.void_class(bc[deep], phi_c[deep], R, h)
        sel = np.flatnonzero(deep)
        blk[sel[cls == 0]] = 0
        blk[sel[cls == 1]] = 1
        v0 += float((cls < 2).sum()) * b ** 3
    ev = blk == 255
    nq = int(ev.sum()) * b ** 3
    full = np.repeat(np.repeat(np.repeat(blk.reshape(nb, nb, nb), b, 0), b, 1), b, 2)
    if not ev.any():
        return tid, (int(blk[0]) if (blk == blk[0]).all() else full), None, v0, 0
    cells = np.flatnonzero((full == 255).ravel())
    X = grid.centres(tid, cells)
    phi, C = dz.phi_closest(X)
    code = full.ravel().copy()
    v0 += float(ramp(-phi, h).sum())
    near = phi <= band_out
    code[cells[~near]] = 2
    ci, phi, C = cells[near], phi[near], C[near]
    if len(ci) == 0:
        return tid, code.reshape(n, n, n), None, v0, nq
    if dz.Vh is None:
        tri = np.full(len(ci), -1, np.int32)
        b1 = b2 = np.zeros(len(ci), np.float32)
    else:
        tri, b1, b2 = dz.locate(C)
    o = phi - dz.d_at(tri, b1, b2)
    psi_in = -phi - dz.a_at(tri, b1, b2)
    code[ci] = np.where(o > 0, 2, np.where(psi_in > 0, 1, 0))
    keep = (np.abs(o) <= S_thr) | (np.abs(psi_in) <= 2.0 * h)
    S = None
    if keep.any():
        S = (ci[keep].astype(np.int32), phi[keep].astype(np.float32), tri[keep], b1[keep], b2[keep])
    return tid, code.reshape(n, n, n), S, v0, nq


# =====================================================================================================================
# stage B: opening and labelling
# =====================================================================================================================
def _open_tile(tid):
    st: TileStore = _G["store"]
    R2 = _G["R2"]
    halo = int(2 * math.ceil(math.sqrt(R2))) + 1
    g = st.grid
    if st.scalar(tid) == 2:
        return tid, None, 0
    # quick skip: tile and neighbours all uniform void
    blk = st.block(tid, halo)
    void = blk == 1
    core = (slice(halo, halo + g.n),) * 3
    if not void[core].any() or void.all():
        return tid, None, 0
    O = morph_open(void, R2)
    removed = void[core] & ~O[core]
    idx = np.flatnonzero(removed.ravel()).astype(np.int32)
    return tid, (idx if len(idx) else None), (int(void[core].sum()))


def _label_tile(args):
    tid, value = args
    st: TileStore = _G["store"]
    g = st.grid
    n = g.n
    s = st.scalar(tid)
    if s is not None and s != value:
        return tid, 0, None, None, None, None
    blk = st.block(tid, 1)
    mask = blk[1:-1, 1:-1, 1:-1] == value
    if not mask.any():
        return tid, 0, None, None, None, None
    lab, m = ndimage.label(mask)
    cnt = np.bincount(lab.ravel(), minlength=m + 1)[1:]
    ext = np.zeros(m + 1, bool)
    if value == 1:                                   # touches the exterior (code 2) through a face
        out = blk == 2
        adj = np.zeros((n, n, n), bool)
        for ax in range(3):
            for sh in (-1, 1):
                sl = [slice(1, -1)] * 3
                sl[ax] = slice(0, n) if sh == -1 else slice(2, n + 2)
                adj |= out[tuple(sl)]
        ext[np.unique(lab[adj & mask])] = True
    faces = (lab[0], lab[-1], lab[:, 0], lab[:, -1], lab[:, :, 0], lab[:, :, -1])
    return tid, m, cnt, ext[1:], faces, None


def _apply_labels(args):
    """Set cells of selected local labels to `newcode`; selection = array of bool per local label."""
    tid, value, newcode, sel = args
    st: TileStore = _G["store"]
    arr = st.get(tid)
    lab, m = ndimage.label(arr == value)
    pick = np.zeros(m + 1, bool)
    pick[1:] = sel
    hit = pick[lab]
    if not hit.any():
        return tid, 0
    arr = arr.copy()
    arr[hit] = newcode
    st.set(tid, arr)       # note: in a forked worker; parent applies from returned indices
    return tid, np.flatnonzero(hit.ravel()).astype(np.int32)


def _label_global(store: TileStore, tids, value: int, workers: int):
    """Global 6-connected components of {code == value}. Returns (comp_of_local dict, ncomp, vol cells, ext flags)."""
    _G["store"] = store
    res = _pmap(_label_tile, [(t, value) for t in tids], workers)
    info = {r[0]: r for r in res}
    off, tot = {}, 0
    for t in tids:
        off[t] = tot
        tot += info[t][1]
    edges = []
    g = store.grid
    for t in tids:
        r = info[t]
        if r[1] == 0:
            continue
        for ax in range(3):
            t2 = list(t)
            t2[ax] += 1
            t2 = tuple(t2)
            if t2 not in info or info[t2][1] == 0:
                continue
            la, lb = r[4][2 * ax + 1], info[t2][4][2 * ax]
            sel = (la > 0) & (lb > 0)
            if sel.any():
                pr = np.unique(np.stack([la[sel] + off[t], lb[sel] + off[t2]], 1), axis=0)
                edges.append(pr - 1)
    if tot == 0:
        return info, off, 0, np.zeros(0, int), np.zeros(0, int), np.zeros(0, bool)
    if edges:
        E = np.concatenate(edges)
        gr = sparse.coo_matrix((np.ones(len(E)), (E[:, 0], E[:, 1])), shape=(tot, tot))
    else:
        gr = sparse.coo_matrix((tot, tot))
    ncomp, comp = connected_components(gr, directed=False)
    cnt = np.zeros(tot)
    ext = np.zeros(tot, bool)
    for t in tids:
        r = info[t]
        if r[1] == 0:
            continue
        cnt[off[t]:off[t] + r[1]] = r[2]
        ext[off[t]:off[t] + r[1]] = r[3]
    vol = np.bincount(comp, weights=cnt, minlength=ncomp)
    cext = np.zeros(ncomp, bool)
    np.logical_or.at(cext, comp, ext)
    return info, off, ncomp, comp, vol, cext


# =====================================================================================================================
# dense ψ of a tile from codes, S, holes
# =====================================================================================================================
def _hole_psi_into(psi_in, grid: Grid, tid, holes, h):
    if holes is None or len(holes["p"]) == 0:
        return
    n = grid.n
    lo = grid.tile_lo(tid)
    hi = lo + n * h
    for p, nv, s0, s1, r in zip(holes["p"], holes["n"], holes["s0"], holes["s1"], holes["r"]):
        ext = s1 + r + h
        if np.any(p + ext < lo) or np.any(p - ext > hi):
            continue
        ij = []
        for a in range(3):
            i0 = max(int(math.floor((p[a] - ext - lo[a]) / h)), 0)
            i1 = min(int(math.ceil((p[a] + ext - lo[a]) / h)) + 1, n)
            ij.append((i0, i1))
        if any(i1 <= i0 for i0, i1 in ij):
            continue
        x = lo[0] + (np.arange(*ij[0]) + 0.5) * h
        y = lo[1] + (np.arange(*ij[1]) + 0.5) * h
        z = lo[2] + (np.arange(*ij[2]) + 0.5) * h
        rx, ry, rz = x[:, None, None] - p[0], y[None, :, None] - p[1], z[None, None, :] - p[2]
        s = rx * nv[0] + ry * nv[1] + rz * nv[2]
        qx, qy, qz = rx - s * nv[0], ry - s * nv[1], rz - s * nv[2]
        rho = np.sqrt(qx * qx + qy * qy + qz * qz)
        ph = np.minimum(np.minimum(r - rho, s - s0), s1 - s).astype(np.float32)
        sl = tuple(slice(i0, i1) for i0, i1 in ij)
        psi_in[sl] = np.maximum(psi_in[sl], ph)


def _psi_parts(tid, delta: float, with_S: bool = False):
    """(o, psi_in[, S]) dense float32 (n³) for tile `tid` at mass shift δ; ±BIG on cells without table entry."""
    dz: Design = _G["design"]
    st: TileStore = _G["store"]
    grid: Grid = _G["grid"]
    h = grid.h
    n = grid.n
    BIG = 3.0 * h
    code = st.get(tid)
    o = np.where(code >= 2, BIG, -BIG).astype(np.float32)
    psi_in = np.where(code == 1, BIG, -BIG).astype(np.float32)
    S = st.S.get(tid)
    if S is not None:
        idx, phi, tri, b1, b2 = S
        phi64 = phi.astype(np.float64)
        cd = code.ravel()[idx]
        oo = phi64 - dz.d_at(tri, b1, b2, delta)
        pin = -phi64 - dz.a_at(tri, b1, b2)
        changed = (pin > 0) & (cd != 1)
        pin = np.where(changed, -0.5 * h, pin)
        dead = cd == 3
        oo = np.where(dead, BIG, oo)
        o.reshape(-1)[idx] = np.clip(oo, -BIG, BIG)
        psi_in.reshape(-1)[idx] = np.clip(pin, -BIG, BIG)
    pi3 = psi_in.reshape(n, n, n)
    _hole_psi_into(pi3, grid, tid, _G.get("holes"), h)
    return o.reshape(n, n, n), pi3


def _psi_tile(tid, delta):
    o, pin = _psi_parts(tid, delta)
    return np.maximum(o, pin)


# =====================================================================================================================
# stage D: per-tile volume constants and S arrays for Brent
# =====================================================================================================================
def _vol_tile(tid):
    st: TileStore = _G["store"]
    grid: Grid = _G["grid"]
    h = grid.h
    n = grid.n
    S = st.S.get(tid)
    has_holes = _G.get("holes") is not None and len(_G["holes"]["p"]) > 0
    c = st.codes[tid]
    if S is None and not has_holes:
        if isinstance(c, (int, np.integer)):
            return tid, (float(n ** 3) if c == 0 else 0.0), None
        return tid, float((c == 0).sum()), None
    o, pin = _psi_parts(tid, 0.0)
    f_in = ramp(-pin, h).astype(np.float32)
    psi = np.maximum(o, pin)
    f = ramp(-psi, h)
    if S is None:
        return tid, float(f.sum()), None
    idx = S[0]
    f_flat = f.reshape(-1).copy()
    fin_S = f_in.reshape(-1)[idx].copy()
    fin_S[st.get(tid).reshape(-1)[idx] == 3] = 0.0           # dead material islands never count
    f_flat[idx] = 0.0
    return tid, float(f_flat.sum()), (S[1], S[2], S[3], S[4], fin_S)


# =====================================================================================================================
# stage E: marching cubes
# =====================================================================================================================
def _psi_block(tid, delta):
    """ψ on the (n+1)³ lattice node block of tile tid (one extra layer from the +1 neighbours); ≈ +3h outside."""
    st: TileStore = _G["store"]
    grid: Grid = _G["grid"]
    n, h = grid.n, grid.h
    BIG = 3.0 * h
    out = np.full((n + 1,) * 3, BIG, np.float32)
    for di in (0, 1):
        for dj in (0, 1):
            for dk in (0, 1):
                t = (tid[0] + di, tid[1] + dj, tid[2] + dk)
                if any(t[a] >= grid.nt[a] for a in range(3)):
                    continue
                sl_o = tuple(slice(0, n) if dd == 0 else slice(n, n + 1) for dd in (di, dj, dk))
                sl_i = tuple(slice(0, n) if dd == 0 else slice(0, 1) for dd in (di, dj, dk))
                c = st.codes.get(t, 2)
                if isinstance(c, (int, np.integer)) and st.S.get(t) is None and not _tile_has_hole(t):
                    out[sl_o] = BIG if c >= 2 or c == 1 else -BIG
                else:
                    out[sl_o] = _psi_tile(t, delta)[sl_i]
    return out


def _tile_has_hole(tid):
    holes = _G.get("holes")
    if holes is None or len(holes["p"]) == 0:
        return False
    grid = _G["grid"]
    lo = grid.tile_lo(tid)
    hi = lo + grid.n * grid.h
    for p, s1, r in zip(holes["p"], holes["s1"], holes["r"]):
        e = s1 + r + grid.h
        if np.all(p + e >= lo) and np.all(p - e <= hi):
            return True
    return False


def _mesh_tile(args):
    from skimage.measure import marching_cubes
    tid, delta, stride, want_keys = args
    grid: Grid = _G["grid"]
    n, h = grid.n, grid.h
    st: TileStore = _G["store"]
    nbr = [(tid[0] + a, tid[1] + b, tid[2] + c) for a in (0, 1) for b in (0, 1) for c in (0, 1)]
    cs = [st.codes.get(t, 2) for t in nbr]
    if all(isinstance(c, (int, np.integer)) for c in cs) and len(set(int(c) for c in cs)) == 1 \
            and all(st.S.get(t) is None and not _tile_has_hole(t) for t in nbr):
        return tid, None
    psi = _psi_block(tid, delta)
    if stride > 1:
        psi = psi[::stride, ::stride, ::stride]
    eps = np.float32(1e-3 * h)                       # keep every crossing strictly inside an edge (robust weld keys)
    psi = np.where(np.abs(psi) < eps, np.where(psi < 0, -eps, eps), psi).astype(np.float32)
    if not (psi.min() < 0 < psi.max()):
        return tid, None
    v, f, _, _ = marching_cubes(psi, 0.0, allow_degenerate=False)
    base = np.asarray(tid) * n
    pos = grid.origin + (base + v.astype(np.float64) * stride + 0.5) * h
    keys = None
    if want_keys:
        fr = np.abs(v - np.rint(v))
        axis = fr.argmax(1)
        lat = np.rint(v).astype(np.int64)
        lat[np.arange(len(v)), axis] = np.floor(v[np.arange(len(v)), axis]).astype(np.int64)
        lat += base
        keys = ((lat[:, 0] * grid.shape[1] + lat[:, 1]) * grid.shape[2] + lat[:, 2]) * 3 + axis
    return tid, (pos.astype(np.float32), f.astype(np.int32), keys)


# =====================================================================================================================
# mesh utilities
# =====================================================================================================================
def mesh_checks(V: np.ndarray, F: np.ndarray) -> dict:
    """Closed, edge-manifold, consistently oriented surface? V (nv,3), F (nf,3) with shared (welded) vertex ids.

    Directed edges must be unique and each must have its reverse. Also: signed volume (float64), Euler number,
    connected component count (bodies: material components and enclosed voids), degenerate faces.
    """
    F = np.asarray(F, np.int64)
    nf, nv = len(F), len(V)
    deg = int(((F[:, 0] == F[:, 1]) | (F[:, 1] == F[:, 2]) | (F[:, 0] == F[:, 2])).sum())
    d = np.concatenate([F[:, 0] * nv + F[:, 1], F[:, 1] * nv + F[:, 2], F[:, 2] * nv + F[:, 0]])
    r = np.concatenate([F[:, 1] * nv + F[:, 0], F[:, 2] * nv + F[:, 1], F[:, 0] * nv + F[:, 2]])
    d.sort()
    dup = int((d[1:] == d[:-1]).sum())
    pos = np.searchsorted(d, r)
    pos[pos >= len(d)] = len(d) - 1
    missing = int((d[pos] != r).sum())
    ne = len(np.unique(np.minimum(d // nv, d % nv) * nv + np.maximum(d // nv, d % nv))) if len(d) < 3e7 else -1
    vol = 0.0
    for s in range(0, nf, 2_000_000):
        t = F[s:s + 2_000_000]
        a, b, c = V[t[:, 0]].astype(np.float64), V[t[:, 1]].astype(np.float64), V[t[:, 2]].astype(np.float64)
        vol += float(np.einsum("ij,ij->i", a, np.cross(b, c)).sum()) / 6.0
    used = np.unique(F)
    g = sparse.coo_matrix((np.ones(3 * nf), (np.concatenate([F[:, 0], F[:, 1], F[:, 2]]),
                                              np.concatenate([F[:, 1], F[:, 2], F[:, 0]]))), shape=(nv, nv))
    nc, lab = connected_components(g, directed=False)
    bodies = len(np.unique(lab[used]))
    return dict(faces=nf, vertices=int(len(used)), degenerate=deg, duplicate_directed_edges=dup,
                missing_reverse_edges=missing, watertight_manifold=bool(dup == 0 and missing == 0 and deg == 0),
                volume=vol, euler=(len(used) - ne + nf if ne >= 0 else None), bodies=bodies)


def drop_material_islands(V: np.ndarray, F: np.ndarray, V0: float, min_vol: float = 0.0):
    """§8.2 step 3 on the final mesh: keep the largest positive-volume body, drop other positive-volume bodies
    (material specks created by the mass shift) and any body smaller than `min_vol` (single-node pockets); enclosed
    voids above `min_vol` have negative volume and stay. Returns (F, info)."""
    nv = len(V)
    g = sparse.coo_matrix((np.ones(3 * len(F)), (np.concatenate([F[:, 0], F[:, 1], F[:, 2]]),
                                                np.concatenate([F[:, 1], F[:, 2], F[:, 0]]))), shape=(nv, nv))
    nc, lab = connected_components(g, directed=False)
    fl = lab[F[:, 0]]
    vol = np.zeros(nc)
    for s in range(0, len(F), 2_000_000):
        t = F[s:s + 2_000_000]
        a, b, c = V[t[:, 0]].astype(np.float64), V[t[:, 1]].astype(np.float64), V[t[:, 2]].astype(np.float64)
        vol += np.bincount(fl[s:s + 2_000_000], weights=np.einsum("ij,ij->i", a, np.cross(b, c)) / 6.0, minlength=nc)
    main = int(np.argmax(vol))
    drop = ((vol > 0) & (np.arange(nc) != main)) | (np.abs(vol) < min_vol)
    drop[main] = False
    info = dict(bodies_before=int(len(np.unique(fl))), islands_dropped=int(drop[np.unique(fl)].sum()),
                islands_volume_mm3=float(vol[drop].sum()))
    return (F[~drop[fl]] if drop.any() else F), info


def write_stl(path: str, V: np.ndarray, F: np.ndarray) -> None:
    """Binary STL, streamed in chunks (float32 vertices, normals from the winding)."""
    dt = np.dtype([("n", "<f4", 3), ("v", "<f4", (3, 3)), ("attr", "<u2")])
    with open(path, "wb") as fh:
        fh.write(b"cshell export".ljust(80, b"\0"))
        fh.write(np.uint32(len(F)).tobytes())
        for s in range(0, len(F), 1_000_000):
            t = F[s:s + 1_000_000]
            rec = np.zeros(len(t), dt)
            tri = V[t].astype(np.float32)
            nrm = np.cross(tri[:, 1] - tri[:, 0], tri[:, 2] - tri[:, 0])
            nrm /= np.maximum(np.linalg.norm(nrm, axis=1, keepdims=True), 1e-30)
            rec["n"], rec["v"] = nrm, tri
            fh.write(rec.tobytes())


def wall_thickness(Vm: np.ndarray, Fm: np.ndarray, n_samples: int, rng, void_side=None) -> dict:
    """Inward ray casting from random surface points (area-weighted): hit distance along −normal.

    `void_side(P, N) -> bool array`: True where the surface faces a void (for the as-built value − b_v)."""
    import igl
    tri = Vm[Fm].astype(np.float64)
    nrm = np.cross(tri[:, 1] - tri[:, 0], tri[:, 2] - tri[:, 0])
    area = 0.5 * np.linalg.norm(nrm, axis=1)
    nrm /= np.maximum(2 * area[:, None], 1e-30)
    pick = rng.choice(len(Fm), size=n_samples, p=area / area.sum())
    u = rng.random((n_samples, 2))
    flip = u.sum(1) > 1
    u[flip] = 1 - u[flip]
    P = tri[pick, 0] + u[:, :1] * (tri[pick, 1] - tri[pick, 0]) + u[:, 1:] * (tri[pick, 2] - tri[pick, 0])
    N = nrm[pick]
    aabb = igl.AABB()
    V64 = np.ascontiguousarray(Vm, np.float64)
    F64 = np.ascontiguousarray(Fm, np.int64)
    aabb.init(V64, F64)
    O = P - 1e-4 * N
    res = aabb.intersect_ray_first(V64, F64, np.ascontiguousarray(O), np.ascontiguousarray(-N))
    hit, t = np.asarray(res[0]), np.asarray(res[1], np.float64)
    ok = (hit >= 0) & np.isfinite(t)
    t = t[ok] + 1e-4
    return dict(t=t, P=P[ok], N=N[ok], n_miss=int((~ok).sum()))


# =====================================================================================================================
# the exporter
# =====================================================================================================================
class Exporter:
    """Tiled export of one `Design` (§8.1-8.5). See the module docstring for the stages."""

    def __init__(self, design: Design, cfg: ExportConfig, normals=None, mount_mesh=None, stress=None, r_ma=None,
                 lo_hi=None):
        self.dz, self.cfg = design, cfg
        self.normals, self.mount_mesh, self.stress, self.r_ma = normals, mount_mesh, stress, r_ma
        self.log = cfg.log
        self.workers = cfg.workers or len(os.sched_getaffinity(0))
        self.timing: dict = {}
        self.summary: dict = {}
        pad = design.d_out_max + cfg.delta_retry + 4 * cfg.h_x + 0.5
        lo = design.V0.min(0) - pad if lo_hi is None else lo_hi[0]
        hi = design.V0.max(0) + pad if lo_hi is None else lo_hi[1]
        self.grid = Grid(lo, hi, cfg.h_x, cfg.tile)
        if cfg.tile % cfg.block:
            raise ValueError("tile must be a multiple of block")

    # -- helpers ---------------------------------------------------------------------------------------------------
    def _t(self, name, t0):
        self.timing[name] = round(time.time() - t0, 2)
        self.log(f"[export] {name}: {self.timing[name]} s")

    def _setG(self):
        _G.update(design=self.dz, cfg=self.cfg, grid=self.grid, store=self.store, delta=self.delta_shell,
                  holes=getattr(self, "holes", None))

    # -- stage A ---------------------------------------------------------------------------------------------------
    def stage_a(self, delta_shell: float):
        t0 = time.time()
        self.delta_shell = delta_shell
        self.store = TileStore(self.grid)
        self._setG()
        res = _pmap(_pass_a, self.grid.tile_ids(), self.workers)
        v0 = 0.0
        nq = 0
        for tid, code, S, v, q in res:
            self.store.set(tid, code)
            if S is not None:
                self.store.S[tid] = S
            v0 += v
            nq += q
        self.V0 = v0 * self.cfg.h_x ** 3
        self.summary.update(V0_mm3=self.V0, n_tiles=len(res), n_cell_queries=nq,
                            n_shell_cells=int(sum(len(s[0]) for s in self.store.S.values())))
        self._t("A_geometry", t0)

    def codes_count(self, value):
        n = self.grid.n
        tot = 0
        for t, c in self.store.codes.items():
            tot += (n ** 3 if c == value else 0) if isinstance(c, (int, np.integer)) else int((c == value).sum())
        return tot

    # -- stage B ---------------------------------------------------------------------------------------------------
    def stage_b(self):
        cfg, h = self.cfg, self.cfg.h_x
        t0 = time.time()
        R2 = ball_r2(0.5 * cfg.v_min, h)
        tids = [t for t in self.grid.tile_ids() if self.store.scalar(t) != 2]
        self._setG()
        _G["R2"] = R2
        rep = {"void_raw_mm3": self.codes_count(1) * h ** 3}
        res = _pmap(_open_tile, tids, min(self.workers, cfg.open_workers))
        n_removed = 0
        for tid, idx, _ in res:
            if idx is None:
                continue
            arr = self.store.get(tid).copy()
            arr.reshape(-1)[idx] = 0
            self.store.set(tid, arr)
            n_removed += len(idx)
        rep["opening_removed_mm3"] = n_removed * h ** 3
        self._t("B1_opening", t0)
        t0 = time.time()
        # fill enclosed voids below void_min_mm3
        self._setG()
        info, off, ncomp, comp, vol, ext = _label_global(self.store, tids, 1, self.workers)
        rep["void_components_after_opening"] = int(ncomp)
        fill = (~ext) & (vol * h ** 3 < cfg.void_min_mm3)
        rep["enclosed_voids_filled"] = int(fill.sum())
        rep["enclosed_filled_mm3"] = float(vol[fill].sum() * h ** 3) if ncomp else 0.0
        jobs = []
        for t in tids:
            r = info[t]
            if r[1] and fill[comp[off[t]:off[t] + r[1]]].any():
                jobs.append((t, 1, 0, fill[comp[off[t]:off[t] + r[1]]]))
        for tid, idx in _pmap(_apply_labels, jobs, self.workers):
            if not isinstance(idx, int):
                arr = self.store.get(tid).copy()
                arr.reshape(-1)[idx] = 0
                self.store.set(tid, arr)
        # largest material component
        self._setG()
        info, off, ncomp, comp, vol, _ = _label_global(self.store, tids, 0, self.workers)
        rep["material_components"] = int(ncomp)
        if ncomp > 1:
            big = int(np.argmax(vol))
            dead = np.ones(ncomp, bool)
            dead[big] = False
            rep["islands_removed"] = int(dead.sum())
            rep["islands_removed_mm3"] = float(vol[dead].sum() * h ** 3)
            jobs = []
            for t in tids:
                r = info[t]
                if r[1] and dead[comp[off[t]:off[t] + r[1]]].any():
                    jobs.append((t, 0, 3, dead[comp[off[t]:off[t] + r[1]]]))
            for tid, idx in _pmap(_apply_labels, jobs, self.workers):
                if not isinstance(idx, int):
                    arr = self.store.get(tid).copy()
                    arr.reshape(-1)[idx] = 3
                    self.store.set(tid, arr)
        else:
            rep["islands_removed"], rep["islands_removed_mm3"] = 0, 0.0
        rep["void_clean_mm3"] = self.codes_count(1) * h ** 3
        self._t("B2_fill_islands", t0)
        if cfg.check_opening:
            t0 = time.time()
            self._setG()
            res = _pmap(_open_tile, tids, min(self.workers, cfg.open_workers))
            rep["reopen_changed_cells"] = int(sum(len(r[1]) for r in res if r[1] is not None))
            rep["opening_rule_ok"] = rep["reopen_changed_cells"] <= 1
            self._t("B3_opening_check", t0)
        self.summary["cleanup"] = rep
        self.log(f"[export] cleanup: {rep}")
        return tids

    # -- stage C ---------------------------------------------------------------------------------------------------
    def stage_c(self, tids):
        cfg, dz, h = self.cfg, self.dz, self.cfg.h_x
        t0 = time.time()
        self.holes = None
        _G["holes"] = None
        self._setG()
        info, off, ncomp, comp, vol, ext = _label_global(self.store, tids, 1, self.workers)
        voids = [dict(id=int(c), volume_mm3=float(vol[c] * h ** 3), enclosed=bool(not ext[c])) for c in range(ncomp)]
        self.summary["voids"] = dict(total=int(ncomp), enclosed=int((~ext).sum()) if ncomp else 0)
        rows = []
        self.holes_info = []
        if dz.Vh is None or ncomp == 0 or not (~ext).any() or not cfg.holes:
            self.summary["voids"]["holes"] = 0
            self.summary["voids"]["without_hole"] = int((~ext).sum()) if ncomp else 0
            self._t("C_holes", t0)
            return
        Vh, Th = dz.Vh, dz.Th
        N = self.normals
        a, d = np.where(dz.closed, np.nan, dz.a), dz.d
        open_v = ~dz.closed
        if self.r_ma is not None:
            open_v &= (a + 3 * h < self.r_ma)
        q = Vh - (a + 3.0 * h)[:, None] * N                    # probe point inside the void behind each vertex
        site = Vh - a[:, None] * N                              # point on the void wall
        vid = np.flatnonzero(open_v)
        cell = self.grid.cell_of(q[vid])
        inside = np.all((cell >= 0) & (cell < self.grid.shape), 1)
        vid, cell = vid[inside], cell[inside]
        n = self.grid.n
        tile = cell // n
        loc = cell - tile * n
        flat = (loc[:, 0] * n + loc[:, 1]) * n + loc[:, 2]
        key = (tile[:, 0] * self.grid.nt[1] + tile[:, 1]) * self.grid.nt[2] + tile[:, 2]
        gcomp = np.full(len(vid), -1, np.int64)
        for k in np.unique(key):
            sel = np.flatnonzero(key == k)
            tid = tuple(int(x) for x in tile[sel[0]])
            arr = self.store.get(tid)
            lab, m = ndimage.label(arr == 1)
            ll = lab.reshape(-1)[flat[sel]]
            ok = ll > 0
            gcomp[sel[ok]] = comp[off[tid] + ll[ok] - 1]
        # candidate filter
        wall = a + d
        okc = (gcomp >= 0) & (wall[vid] >= cfg.t_min + cfg.b_v)
        if self.mount_mesh is not None and len(self.mount_mesh[1]):
            from .surface import MeshQuery
            mq = MeshQuery(*self.mount_mesh)
            okc &= mq.distance(site[vid]) > cfg.mount_clear
        enclosed_ids = [v["id"] for v in voids if v["enclosed"]]
        holes_p, holes_n, holes_s0, holes_s1 = [], [], [], []
        n_without = 0
        for c in enclosed_ids:
            members = np.flatnonzero(gcomp == c)
            cand = members[okc[members]]
            if len(members) < 3 or len(cand) == 0:
                n_without += 1
                voids[c]["holes"] = 0
                continue
            pts = q[vid[members]]
            w, evec = np.linalg.eigh(np.cov(pts.T) if len(pts) > 3 else np.eye(3))
            axis = evec[:, -1]
            if self.stress is None:
                score, kind = a[vid[cand]], "inset_a"
            else:
                score = (self.stress(site[vid[cand]]) if callable(self.stress) else np.asarray(self.stress)[vid[cand]])
                kind = "stress"
            cut = np.quantile(score, cfg.hole_stress_frac)
            low = cand[score <= cut]
            if len(low) < cfg.holes_per_void:
                low = cand[np.argsort(score)[:max(cfg.holes_per_void, len(low))]]
            proj = site[vid[low]] @ axis
            chosen = [int(np.argmin(proj)), int(np.argmax(proj))] if cfg.holes_per_void >= 2 else [int(np.argmin(np.abs(proj - np.median(proj))))]
            while len(chosen) < cfg.holes_per_void and len(chosen) < len(low):
                gap = np.min(np.abs(proj[:, None] - proj[chosen][None, :]), 1)
                gap[chosen] = -1
                chosen.append(int(np.argmax(gap)))
            if len(set(chosen)) < len(chosen):
                chosen = list(dict.fromkeys(chosen))
            sc = dict(zip(cand.tolist(), np.asarray(score).tolist()))
            for hi_, j in enumerate(chosen):
                v = vid[low[j]]
                holes_p.append(site[v]); holes_n.append(N[v])
                holes_s0.append(-0.5); holes_s1.append(wall[v] + 0.5)
                rows.append([c, hi_, *site[v], *N[v], wall[v], sc[int(low[j])], kind, voids[c]["volume_mm3"]])
            voids[c]["holes"] = len(chosen)
        self.holes = dict(p=np.array(holes_p).reshape(-1, 3), n=np.array(holes_n).reshape(-1, 3),
                          s0=np.array(holes_s0), s1=np.array(holes_s1), r=np.full(len(holes_p), 0.5 * cfg.hole_d))
        self.holes_rows = rows
        self.voids = voids
        self.summary["voids"].update(holes=len(holes_p), without_hole=n_without,
                                     per_void=[dict(id=v["id"], volume_mm3=v["volume_mm3"], holes=v.get("holes", 0))
                                               for v in voids if v["enclosed"]])
        self._t("C_holes", t0)

    # -- stage D ---------------------------------------------------------------------------------------------------
    def stage_d(self):
        cfg, dz, h = self.cfg, self.dz, self.cfg.h_x
        t0 = time.time()
        self._setG()
        res = _pmap(_vol_tile, list(self.grid.tile_ids()), self.workers)
        Vc = 0.0
        arrs = []
        for tid, v, A in res:
            Vc += v
            if A is not None:
                arrs.append(A)
        if arrs:
            phi = np.concatenate([a[0] for a in arrs]); tri = np.concatenate([a[1] for a in arrs])
            b1 = np.concatenate([a[2] for a in arrs]); b2 = np.concatenate([a[3] for a in arrs])
            fin = np.concatenate([a[4] for a in arrs])
        else:
            phi = tri = b1 = b2 = fin = np.zeros(0, np.float32)
        del res, arrs
        phi64 = phi.astype(np.float64)

        def V_of(delta):
            f = np.minimum(ramp(-(phi64 - dz.d_at(tri, b1, b2, delta)), h), fin)
            return (Vc + float(f.sum())) * h ** 3

        v00 = V_of(0.0)
        rel0 = v00 / self.V0 - 1.0
        out = dict(V_before_mm3=v00, rel_before=rel0)
        delta = 0.0
        ok = True
        if dz.trivial_solid:
            pass
        elif abs(rel0) > cfg.mass_tol:
            lim = self.delta_shell
            fl, fh = V_of(-lim) / self.V0 - 1.0, V_of(lim) / self.V0 - 1.0
            if fl * fh > 0:
                out["bracket_failed"] = dict(limit=lim, f_lo=fl, f_hi=fh)
                ok = False
            else:
                delta = brentq(lambda x: V_of(x) / self.V0 - 1.0, -lim, lim, xtol=1e-7, rtol=1e-10)
        out.update(delta=delta, V_after_mm3=V_of(delta) if ok else v00, rel_after=(V_of(delta) / self.V0 - 1.0) if ok else rel0,
                   warn=bool(abs(delta) > cfg.delta_warn), fail=bool(abs(delta) > cfg.delta_fail),
                   n_brent_cells=int(len(phi)), ok=ok)
        self.delta = delta
        self.summary["mass_correction"] = out
        self._t("D_mass", t0)
        return ok

    # -- stage E ---------------------------------------------------------------------------------------------------
    def stage_e(self, out_dir, name="final_design"):
        cfg, h = self.cfg, self.cfg.h_x
        t0 = time.time()
        self._setG()
        items = [(t, self.delta, 1, True) for t in self.grid.tile_ids()]
        res = _pmap(_mesh_tile, items, self.workers)
        keys, pos, faces = [], [], []
        base = 0
        for tid, r in res:
            if r is None:
                continue
            p, f, k = r
            pos.append(p); keys.append(k); faces.append(f.astype(np.int64) + base)
            base += len(p)
        del res
        self._t("E1_marching_cubes", t0)
        t0 = time.time()
        if not faces:
            raise RuntimeError("export produced no surface")
        keys = np.concatenate(keys); pos = np.concatenate(pos); F = np.concatenate(faces)
        uk, first, inv = np.unique(keys, return_index=True, return_inverse=True)
        V = pos[first]
        F = inv[F].astype(np.int64)
        del keys, pos, faces
        keep = (F[:, 0] != F[:, 1]) & (F[:, 1] != F[:, 2]) & (F[:, 0] != F[:, 2])
        n_degenerate = int((~keep).sum())          # Lewiner's MC emits a few zero-area triangles (same edge vertex twice)
        if n_degenerate:
            F = F[keep]
        F, isl = drop_material_islands(V, F, self.V0, min_vol=10.0 * h ** 3)
        chk = mesh_checks(V, F)
        chk["degenerate_removed"] = n_degenerate
        chk["islands"] = isl
        chk["volume_rel"] = chk["volume"] / self.V0 - 1.0
        chk["volume_ok"] = bool(abs(chk["volume_rel"]) <= cfg.vol_tol)
        if chk["faces"] <= cfg.trimesh_check_max_faces:
            import trimesh
            tm = trimesh.Trimesh(V.astype(np.float64), F, process=False)
            chk["trimesh"] = dict(watertight=bool(tm.is_watertight), winding_consistent=bool(tm.is_winding_consistent),
                                  volume=float(tm.volume), bodies=int(len(tm.split(only_watertight=False))))
        self.summary["mesh"] = chk
        self._t("E2_weld_checks", t0)
        if cfg.write_stl:
            t0 = time.time()
            path = os.path.join(out_dir, f"{name}.stl")
            write_stl(path, V, F)
            chk["stl"] = path
            chk["stl_bytes"] = os.path.getsize(path)
            self._t("E3_stl", t0)
        self.mesh = (V, F)
        return V, F

    def wall_stats(self, Vm, Fm, out_dir):
        cfg, h = self.cfg, self.cfg.h_x
        t0 = time.time()
        stride = 1
        if len(Fm) > cfg.wall_max_faces:
            stride = int(math.ceil(math.sqrt(len(Fm) / cfg.wall_max_faces)))
            while self.grid.n % stride:
                stride += 1
            self._setG()
            res = _pmap(_mesh_tile, [(t, self.delta, stride, False) for t in self.grid.tile_ids()], self.workers)
            pos, faces, base = [], [], 0
            for tid, r in res:
                if r is None:
                    continue
                pos.append(r[0]); faces.append(r[1].astype(np.int64) + base); base += len(r[0])
            Vm, Fm = np.concatenate(pos), np.concatenate(faces)
        rng = np.random.default_rng(0)
        w = wall_thickness(Vm, Fm, cfg.wall_samples, rng)
        t = w["t"]
        dz = self.dz
        # as-built: − b_v for every side that faces a void
        if dz.has_void:
            Q0 = w["P"] + 3 * h * w["N"]                       # just outside the material at the sample
            Q1 = w["P"] - (t[:, None] + 3 * h) * w["N"]        # just beyond the far wall
            nvoid = np.zeros(len(t))
            for Q in (Q0, Q1):
                phi, C = dz.phi_closest(Q)
                if dz.Vh is None:
                    a = np.full(len(Q), dz.a_const)
                else:
                    tri, b1, b2 = dz.locate(C)
                    a = dz.a_at(tri, b1, b2)
                nvoid += ((-phi - a) > 0) & (phi < 0)
            t_ab = t - cfg.b_v * nvoid
        else:
            t_ab = t.copy()
        edges = np.linspace(0, max(float(t.max()), 1e-3) * 1.0001, 61)
        hist, _ = np.histogram(t, edges)
        hist_ab, _ = np.histogram(t_ab, edges)
        with open(os.path.join(out_dir, "wall_thickness.csv"), "w", newline="") as f:
            wr = csv.writer(f)
            wr.writerow(["bin_lo_mm", "bin_hi_mm", "count_nominal", "count_as_built"])
            for i in range(len(hist)):
                wr.writerow([edges[i], edges[i + 1], int(hist[i]), int(hist_ab[i])])
        self.summary["wall"] = dict(samples=int(len(t)), miss=w["n_miss"], mesh_stride=stride, min_nominal=float(t.min()),
                                    p1_nominal=float(np.quantile(t, 0.01)), median_nominal=float(np.median(t)),
                                    min_as_built=float(t_ab.min()), p1_as_built=float(np.quantile(t_ab, 0.01)),
                                    frac_below_t_min=float((t < cfg.t_min).mean()))
        self._t("E4_wall_thickness", t0)

    # -- driver ----------------------------------------------------------------------------------------------------
    def run(self, out_dir: str, name: str = "final_design") -> dict:
        cfg = self.cfg
        os.makedirs(out_dir, exist_ok=True)
        t_all = time.time()
        for attempt, ds in enumerate((cfg.delta_shell, cfg.delta_retry)):
            self.stage_a(ds)
            tids = self.stage_b() if (cfg.clean and self.dz.has_void) else [t for t in self.grid.tile_ids()]
            if not (cfg.clean and self.dz.has_void):
                self.summary["cleanup"] = {}
            if cfg.holes and cfg.clean and self.dz.has_void:
                self.stage_c(tids)
            ok = self.stage_d()
            if ok or attempt == 1:
                break
            self.log("[export] mass-correction bracket failed with the small shell; retrying with delta_retry")
        if getattr(self, "holes", None) is not None and len(self.holes["p"]):
            with open(os.path.join(out_dir, "escape_holes.csv"), "w", newline="") as f:
                wr = csv.writer(f)
                wr.writerow(["void_id", "hole", "x", "y", "z", "nx", "ny", "nz", "wall_mm", "score", "score_kind",
                             "void_volume_mm3"])
                wr.writerows(self.holes_rows)
        V, F = self.stage_e(out_dir, name)
        if cfg.wall_hist:
            self.wall_stats(V, F, out_dir)
        self.timing["total"] = round(time.time() - t_all, 2)
        ru = resource.getrusage(resource.RUSAGE_SELF)
        rc = resource.getrusage(resource.RUSAGE_CHILDREN)
        self.summary.update(h_x=cfg.h_x, tile=cfg.tile, workers=self.workers, timing_s=self.timing,
                            peak_rss_parent_GB=ru.ru_maxrss / 1e6, peak_rss_child_GB=rc.ru_maxrss / 1e6,
                            grid_shape=[int(x) for x in self.grid.shape])
        ok = (self.summary["mass_correction"]["ok"] and not self.summary["mass_correction"]["fail"]
              and self.summary["mesh"]["watertight_manifold"] and self.summary["mesh"]["volume_ok"])
        self.summary["ok"] = bool(ok)
        with open(os.path.join(out_dir, f"{name}_export.json"), "w") as f:
            json.dump(self.summary, f, indent=1, default=float)
        return self.summary


# =====================================================================================================================
# high-level entry points
# =====================================================================================================================
def vertex_fields(ctx, xi, D_comp: float = 4.0, h_g: float | None = None):
    """Nominal per-vertex (d, a, d_hi, xt_a, passive) of Γ_h for the packed design xi = [ξᵈ, ξᵃ] (THEORY §3.4).

    Mirrors `optimise.PhaseModel` (a_hi = r_MA + γ_c h_g, d_hi from clearance and D_comp, D1 bound for a_lo)."""
    from types import SimpleNamespace
    from . import config as ccfg
    from .design import DesignMap, DesignSpace, passive_thin
    cfg = ctx.cfg
    nv = ctx.surf.n_vertices
    m, dsg = cfg["manufacturing"], cfg["design"]
    if h_g is None:
        h_g = min(p["h_g"] for p in cfg["grid"]["phases"])
    a_hi = ctx.r_ma + dsg["closure_gamma"] * h_g
    d_hi = np.minimum.reduce([ctx.r_out - dsg["gap_mm"], ctx.keepout - dsg["keep_out_clearance_mm"],
                              np.full(nv, D_comp)]).clip(0.0)
    if not dsg["allow_outward"]:
        d_hi[:] = 0.0
    passive = ctx.mount | passive_thin(a_hi, m["t_min_mm"], m["void_min_mm"])
    r_f = ccfg.phase_value(cfg["filter"]["radius_mm"], len(cfg["grid"]["phases"]) - 1)
    ds = DesignSpace(r_ma=ctx.r_ma, a_hi=a_hi, d_hi=d_hi, area=ctx.surf.A, passive=passive, filt=ctx.surf.filter(r_f),
                     t_min=m["t_min_mm"], v_min=m["void_min_mm"], b_v=m["wall_bias_mm"], b_o=m["outer_bias_mm"],
                     soft_w=dsg["inset_soft_mm"])
    xi = np.asarray(xi, np.float64)
    xt_d, xt_a, d, a, _, _ = DesignMap.vertex_fields(SimpleNamespace(xp=np, ds=ds), xi[:nv], xi[nv:])
    return dict(d=d, a=a, d_hi=d_hi, xt_a=xt_a, passive=passive, r_ma=ctx.r_ma)


def export_design(ctx, xi, out_dir: str, D_comp: float = 4.0, stress=None, name: str = "final_design", h_g=None,
                  **cfg_kw) -> dict:
    """Export the nominal design xi of `ctx` to `out_dir` (final_design.stl, escape_holes.csv, wall_thickness.csv,
    <name>_export.json). `stress`: None, (N_v,) array of material stress at the void-wall site below each vertex of
    Γ_h, or a callable points -> stress. Returns the summary dict (also written as JSON)."""
    cfg = ExportConfig.from_run_cfg(ctx.cfg, **cfg_kw)
    f = vertex_fields(ctx, xi, D_comp, h_g)
    s = ctx.surf
    closed = f["xt_a"] >= cfg.closed_thr
    dz = Design(ctx.V, ctx.T, s.V, s.T, f["d"], f["a"], f["d_hi"], active=~f["passive"], closed=closed, sign=cfg.sign)
    N = s.projected_normals() if hasattr(s, "projected_normals") else s.N
    mount_mesh = None
    if len(getattr(ctx, "mount_ids", [])):
        sel = np.isin(ctx.TF, np.asarray(ctx.mount_ids))
        if sel.any():
            mount_mesh = (ctx.V, ctx.T[sel])
    ex = Exporter(dz, cfg, normals=N, mount_mesh=mount_mesh, stress=stress, r_ma=f["r_ma"])
    return ex.run(out_dir, name)


def export_baselines(ctx, out_dir: str, **cfg_kw) -> dict:
    """original.stl (ψ = φ) and uniform_shell.stl (ψ = max(φ − t, −φ), t solved for constant mass, §7.4 S1)."""
    cfg = ExportConfig.from_run_cfg(ctx.cfg, **cfg_kw)
    return baselines(ctx.V, ctx.T, out_dir, cfg)


def solve_shell_thickness(V0, T0, h_c: float = 0.25, t_max: float = 8.0, sign="pseudonormal", workers=None) -> float:
    """t with Vol{0 ≤ φ ≤ t} = Vol{φ ≤ 0} on a coarse grid (φ histogram, linear-ramp volumes)."""
    dz = Design(V0, T0, d_const=0.0, a_const=0.0, sign=sign)
    cfg = ExportConfig(h_x=h_c, tile=64, block=8, workers=workers, sign=sign, log=lambda *_: None)
    grid = Grid(dz.V0.min(0) - t_max - 1, dz.V0.max(0) + t_max + 1, h_c, cfg.tile)
    _G.update(design=dz, grid=grid, cfg=cfg, hist=(t_max + 2 * h_c, 2e-3))
    bw = 2e-3
    edges = np.arange(-2 * h_c, t_max + 2 * h_c + bw, bw)

    def job(tid):
        n, b = grid.n, cfg.block
        nb = n // b
        lo = grid.tile_lo(tid)
        ax = [lo[k] + (np.arange(nb) + 0.5) * b * h_c for k in range(3)]
        bc = np.stack(np.meshgrid(*ax, indexing="ij"), -1).reshape(-1, 3)
        phi_c = dz.phi(bc)
        R = 0.5 * b * h_c * math.sqrt(3)
        far_out = phi_c - R > t_max + 2 * h_c
        deep = phi_c + R < -2 * h_c
        ev = ~(far_out | deep)
        n_deep = int(deep.sum()) * b ** 3
        hist = np.zeros(len(edges) - 1)
        if ev.any():
            blk = np.zeros(nb ** 3, bool); blk[ev] = True
            full = np.repeat(np.repeat(np.repeat(blk.reshape(nb, nb, nb), b, 0), b, 1), b, 2)
            X = grid.centres(tid, np.flatnonzero(full.ravel()))
            phi = dz.phi(X)
            hist = np.histogram(np.clip(phi, edges[0] + 1e-9, edges[-1] - 1e-9), edges)[0].astype(np.float64)
        return tid, n_deep, hist

    global _JOB
    _JOB = job
    res = _pmap(_hist_job, grid.tile_ids(), workers or len(os.sched_getaffinity(0)))
    n_deep = sum(r[1] for r in res)
    hist = sum(r[2] for r in res)
    mid = 0.5 * (edges[1:] + edges[:-1])
    V0 = h_c ** 3 * (n_deep + float((hist * ramp(-mid, h_c)).sum()))

    def V(t):
        return h_c ** 3 * float((hist * np.minimum(ramp(-(mid - t), h_c), ramp(mid, h_c))).sum())
    return brentq(lambda t: V(t) - V0, 1e-3, t_max, xtol=1e-6)


_JOB = None


def _hist_job(tid):
    return _JOB(tid)


def baselines(V0, T0, out_dir: str, cfg: ExportConfig) -> dict:
    os.makedirs(out_dir, exist_ok=True)
    out = {}
    cl = dict(clean=False, holes=False)
    # original
    c = ExportConfig(**{**asdict(cfg), **cl, "log": cfg.log})
    ex = Exporter(Design(V0, T0, d_const=0.0, a_const=A_CLOSED, sign=cfg.sign), c)
    out["original"] = ex.run(out_dir, "original")
    # uniform shell
    t0 = time.time()
    tc = solve_shell_thickness(V0, T0, h_c=max(cfg.h_x, 0.25), sign=cfg.sign, workers=cfg.workers)
    cfg.log(f"[export] uniform shell: coarse t = {tc:.4f} mm ({time.time() - t0:.1f} s)")
    ex = Exporter(Design(V0, T0, d_const=tc, a_const=0.0, sign=cfg.sign), c)
    s = ex.run(out_dir, "uniform_shell")
    s["t_coarse_mm"], s["t_final_mm"] = tc, tc + s["mass_correction"]["delta"]
    out["uniform_shell"] = s
    return out
