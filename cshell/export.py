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

`GpuExporter` / `export_design_gpu` (PLAN_single_run §2-§3, the default export of `cshell.finalise`, h_x = 0.2 mm):
the same stages on one GPU with a few GiB of device and host memory. Warp BVH queries and the optimiser's fused
interpolation (design.cell_psi) per tile, a smooth distance-based opening of the void instead of the lattice opening
(no constant −h/2 cells), Warp marching cubes, δ against the exact solid volume with secant steps on the mesh volume.
`Exporter` (CPU, forked workers) is kept for the legacy `collect` and its tests.
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

MC_HALO = 32                # lattice layers around a tile for the ambiguous-face rule: flips propagate ≤ 1 node per
                            # iteration, so tiles agree on shared nodes while every tile converges in < MC_HALO iterations
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
    # GPU export (GpuExporter, PLAN_single_run §2-§3)
    mc_halo: int = 8                  # lattice layers around a tile for the ambiguous-face rule
    mass_correct: bool = True
    mesh_vol_tol: float = 2.0e-4      # secant steps on δ until |V_mesh/V_target − 1| ≤ this (≤ 2 steps)
    target_volume: float | None = None   # V_target (mm³); None: exact volume of the Γ₀ tessellation
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


def closed_cells(closed_f: np.ndarray, v: np.ndarray, b0, b1, b2) -> np.ndarray:
    """Cells whose closest-triangle vertices with positive (clipped, normalised) weight are all closed: Σ w_k c_k ≥
    1 − 1e-9 (same test as collect's re-scoring kernel, `design.rescore_fwd`)."""
    w0, w1, w2 = np.maximum(b0, 0.0), np.maximum(b1, 0.0), np.maximum(b2, 0.0)
    sw = (w0 + w1) + w2
    cc = ((w0 / sw) * closed_f[v[:, 0]] + (w1 / sw) * closed_f[v[:, 1]]) + (w2 / sw) * closed_f[v[:, 2]]
    return cc >= 1.0 - 1e-9


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
        # closure rule (§8.2, PLAN_collect_fixes P2): a cell is closed (a = A_CLOSED) only when every vertex of its
        # closest Γ_h triangle with positive weight is closed; otherwise a is the linear interpolation of the
        # optimiser's a (closed vertices keep their a ≈ a^hi). Interpolating A_CLOSED itself closed one extra triangle
        # ring along every closed/open boundary (+3.8 % V on the MBB smoke best).
        self.a_lin = np.asarray(a, np.float64).copy()
        self.closed_f = closed.astype(np.float64)
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
        b0 = 1.0 - b1 - b2
        a = b0 * self.a_lin[v[:, 0]] + b1 * self.a_lin[v[:, 1]] + b2 * self.a_lin[v[:, 2]]
        return np.where(closed_cells(self.closed_f, v, b0, b1, b2), A_CLOSED, a)

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
def _psi_block(tid, delta, halo: int = 0):
    """ψ on the lattice nodes of tile tid: the (n+1)³ block (one extra layer from the +1 neighbours) widened by
    `halo` layers on every side (from the ±1 neighbours), i.e. global node indices tid·n − halo … tid·n + n + halo;
    ≈ +3h outside the grid."""
    st: TileStore = _G["store"]
    grid: Grid = _G["grid"]
    n, h = grid.n, grid.h
    BIG = 3.0 * h
    m = n + 1 + 2 * halo
    out = np.full((m,) * 3, BIG, np.float32)
    lo = [tid[a] * n - halo for a in range(3)]                      # global node index of out[0, 0, 0]
    rng_ = (-1, 0, 1) if halo > 0 else (0, 1)
    for di in rng_:
        for dj in rng_:
            for dk in rng_:
                t = (tid[0] + di, tid[1] + dj, tid[2] + dk)
                if any(t[a] < 0 or t[a] >= grid.nt[a] for a in range(3)):
                    continue
                sl_o, sl_i = [], []
                for a in range(3):
                    g0, g1 = max(t[a] * n, lo[a]), min(t[a] * n + n, lo[a] + m)
                    sl_o.append(slice(g0 - lo[a], g1 - lo[a]))
                    sl_i.append(slice(g0 - t[a] * n, g1 - t[a] * n))
                if any(s_.stop <= s_.start for s_ in sl_o):
                    continue
                c = st.codes.get(t, 2)
                if isinstance(c, (int, np.integer)) and st.S.get(t) is None and not _tile_has_hole(t):
                    out[tuple(sl_o)] = BIG if c >= 2 or c == 1 else -BIG
                else:
                    out[tuple(sl_o)] = _psi_tile(t, delta)[tuple(sl_i)]
    return out


def resolve_ambiguous_faces(psi: np.ndarray, value: float, max_iter: int = 64):
    """Remove every ambiguous lattice face (signs + − + − around the face: material {ψ < 0} touching itself only
    across the face diagonal) by moving its material nodes to the void side (ψ ← +value), repeated until none is left.

    Why: marching cubes resolves an ambiguous face by the asymptotic decider; when the decider ties (ac = bd, e.g. the
    ±3h values of code-only cells or clipped ψ) or in skimage's Lewiner case 10 the two cubes sharing the face resolve
    it differently and the mesh gets non-manifold edges (four faces on one edge = duplicated directed edges). Without
    ambiguous faces every cube face has 0 or 2 crossings and the triangulation is consistent. The rule matches the
    clean-up's 6-connectivity (`ndimage.label`): material that touches only diagonally is already two components.
    Removes at most one node per face (≤ h_x of material, sub-cell). In place; returns (nodes flipped, iterations)."""
    xp = np
    if type(psi).__module__.startswith("cupy"):          # same rule on the device (GPU export)
        import cupy as xp
    flipped, it = 0, 0
    for it in range(1, max_iter + 1):
        mat = psi < 0
        flip = xp.zeros_like(mat)
        for ax in range(3):
            a, b = [x for x in range(3) if x != ax]
            q = xp.moveaxis(mat, (a, b), (0, 1))
            amb = (q[:-1, :-1] == q[1:, 1:]) & (q[1:, :-1] == q[:-1, 1:]) & (q[:-1, :-1] != q[1:, :-1])
            if not amb.any():
                continue
            fl = xp.moveaxis(flip, (a, b), (0, 1))                     # view: writes go to `flip`
            fl[:-1, :-1] |= amb & q[:-1, :-1]
            fl[1:, 1:] |= amb & q[1:, 1:]
            fl[1:, :-1] |= amb & q[1:, :-1]
            fl[:-1, 1:] |= amb & q[:-1, 1:]
        k = int(flip.sum())
        if k == 0:
            return flipped, it - 1
        psi[flip] = value
        flipped += k
    return flipped, -it                                                 # not converged (negative iteration count)


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
    H = min(MC_HALO, n)
    psi = _psi_block(tid, delta, H)
    eps = np.float32(1e-3 * h)                       # keep every crossing strictly inside an edge (robust weld keys)
    psi = np.where(np.abs(psi) < eps, np.where(psi < 0, -eps, eps), psi).astype(np.float32)
    n_flip, n_it = resolve_ambiguous_faces(psi, eps)  # global rule, evaluated with a halo: tiles agree on shared nodes
    psi = psi[H:H + n + 1, H:H + n + 1, H:H + n + 1]
    if stride > 1:
        psi = psi[::stride, ::stride, ::stride]
    amb = (n_flip, n_it)
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
        keys[(fr > 1e-6).sum(1) != 1] = -1          # not on a lattice edge (MC33 cube-centre vertex): never welded
    return tid, (pos.astype(np.float32), f.astype(np.int32), keys, amb)


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


def drop_material_islands(V: np.ndarray, F: np.ndarray, V0: float, min_vol: float = 0.0, void_fill: float = 0.0,
                          near_mount=None, mount_tol: float = 0.0):
    """§8.2 steps 2-3 on the final mesh: keep the largest positive-volume body, drop other positive-volume bodies
    (material specks created by the mass shift), any body smaller than `min_vol` (single-node pockets) and enclosed
    voids (negative volume) smaller than `void_fill` (= void_min_mm3: §8.2 fills them; sub-cell slits the grid
    clean-up cannot see). Larger enclosed voids stay. `near_mount(points) -> distance`: dropped material islands with a
    vertex within `mount_tol` of a load/support face are counted (they should not exist). Returns (F, info)."""
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
    present = np.zeros(nc, bool)
    present[np.unique(fl)] = True
    isl = present & (((vol > 0) & (np.arange(nc) != main)) | ((vol >= 0) & (vol < min_vol)))
    vfill = present & (vol < 0) & (-vol < max(void_fill, min_vol))
    isl[main] = vfill[main] = False
    drop = isl | vfill
    info = dict(bodies_before=int(present.sum()), islands_dropped=int(isl.sum()),
                islands_volume_mm3=float(vol[isl].sum()), voids_filled=int(vfill.sum()),
                voids_filled_mm3=float(-vol[vfill].sum()), islands_on_mount=0, islands_on_mount_mm3=0.0)
    if near_mount is not None and isl.any():
        fi = np.flatnonzero(isl[fl])
        vi = np.unique(F[fi])
        close = near_mount(V[vi].astype(np.float64)) <= mount_tol
        if close.any():
            on = np.unique(lab[vi[close]])
            info.update(islands_on_mount=int(len(on)), islands_on_mount_mm3=float(vol[on].sum()))
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
    cos = np.einsum("ij,ij->i", nrm[hit[ok]], N[ok])
    return dict(t=t, P=P[ok], N=N[ok], cos=cos, n_miss=int((~ok).sum()))


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
        self._fill_islands(tids, rep)
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

    def _fill_islands(self, tids, rep: dict) -> None:
        """§8.2 steps 2-3 on the codes: enclosed voids < void_min_mm3 filled (code 1 → 0), every material component
        but the largest marked dead (code 0 → 3)."""
        cfg, h = self.cfg, self.cfg.h_x
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

    def _codes_at(self, X):
        """Cell code (0 material, 1 void, 2 outside, 3 dead) at points X (k, 3); 255 outside the grid."""
        cell = self.grid.cell_of(X)
        n = self.grid.n
        out = np.full(len(X), 255, np.uint8)
        ok = np.all((cell >= 0) & (cell < self.grid.shape), 1)
        tile, loc = cell // n, cell % n
        for i in np.flatnonzero(ok):
            tid = tuple(int(x) for x in tile[i])
            if tid in self.store.codes:
                c = self.store.codes[tid]
                out[i] = c if isinstance(c, (int, np.integer)) else c[tuple(loc[i])]
        return out

    def _hole_extent(self, p, nv, wall, r, s_max_in: float = 3.0):
        """Axial extent (s0, s1) of an escape hole at void-wall point p with outward axis nv such that the inner end
        cap lies wholly in the void and the outer cap wholly outside the part. A flat cap at a fixed s0 = -0.5 mm cut
        the curved void wall at grazing incidence and left films down to 0.0015 mm (B2, research/b2_bodyfit).
        Returns None when no such extent exists within s_max_in (the void is narrower than the hole): reject the site."""
        h = self.cfg.h_x
        e1 = np.cross(nv, [1.0, 0, 0]) if abs(nv[0]) < 0.9 else np.cross(nv, [0, 1.0, 0])
        e1 /= np.linalg.norm(e1)
        e2 = np.cross(nv, e1)
        ang = np.linspace(0, 2 * np.pi, 48, endpoint=False)
        disk = np.concatenate([[[0.0, 0.0]], np.column_stack([np.cos(ang), np.sin(ang)]) * (r + h),
                               np.column_stack([np.cos(ang[::2]), np.sin(ang[::2])]) * 0.5 * r])
        D = disk[:, :1] * e1 + disk[:, 1:] * e2
        s0 = None
        for s in np.arange(-0.5, -s_max_in, -h):
            c = self._codes_at(p + s * nv + D)
            if np.all(c == 1):
                s0 = s - h
                break
            if np.any(c == 2):          # reached the far side: the hole would cut through the opposite wall
                break
        if s0 is None:
            return None
        s1 = None
        for s in np.arange(wall + 0.5, wall + 0.5 + s_max_in, h):
            if np.all(self._codes_at(p + s * nv + D) == 2):
                s1 = s + h
                break
        return None if s1 is None else (float(s0), float(s1))

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
            n_added = 0
            for hi_, j in enumerate(chosen):
                v = vid[low[j]]
                ext_ = self._hole_extent(site[v], N[v], wall[v], 0.5 * cfg.hole_d)
                if ext_ is None:                    # try the next-best candidates of this void (same score order)
                    for j2 in np.argsort(np.abs(proj - proj[j])):
                        if int(j2) in chosen:
                            continue
                        ext_ = self._hole_extent(site[vid[low[j2]]], N[vid[low[j2]]], wall[vid[low[j2]]], 0.5 * cfg.hole_d)
                        if ext_ is not None:
                            j, v = int(j2), vid[low[j2]]
                            break
                if ext_ is None:
                    self.log(f"[export] void {c}: no escape-hole site with a clean extent (hole {hi_} dropped)")
                    continue
                holes_p.append(site[v]); holes_n.append(N[v])
                holes_s0.append(ext_[0]); holes_s1.append(ext_[1])
                rows.append([c, hi_, *site[v], *N[v], wall[v], sc[int(low[j])], kind, voids[c]["volume_mm3"]])
                n_added += 1
            voids[c]["holes"] = n_added
            n_without += int(n_added == 0)
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
        t0 = time.time()
        mc = self._mc_tiles(self.delta)
        self._t("E1_marching_cubes", t0)
        V, F = self._weld_checks(*mc)
        if self.cfg.write_stl:
            self._write(V, F, out_dir, name)
        self.mesh = (V, F)
        return V, F

    def _mc_tiles(self, delta):
        """Marching cubes of every tile at mass shift δ: (pos, keys, faces) lists (faces offset into the concatenated
        positions) and the ambiguous-face statistics (nodes flipped, max iterations, tiles not converged, halo)."""
        self._setG()
        items = [(t, delta, 1, True) for t in self.grid.tile_ids()]
        res = _pmap(_mesh_tile, items, self.workers)
        keys, pos, faces = [], [], []
        base, n_flip, it_max, n_unconv, n_exc = 0, 0, 0, 0, 0
        for tid, r in res:
            if r is None:
                continue
            p, f, k, (nf_, it_) = r
            n_flip += nf_; it_max = max(it_max, abs(it_)); n_unconv += it_ < 0
            n_exc += abs(it_) >= min(MC_HALO, self.grid.n)
            pos.append(p); keys.append(k); faces.append(f.astype(np.int64) + base)
            base += len(p)
        return pos, keys, faces, (n_flip, it_max, n_unconv, n_exc, MC_HALO)

    def _weld_checks(self, pos, keys, faces, amb):
        """Weld the per-tile triangles by lattice-edge keys, drop degenerate faces, material islands and small enclosed
        voids, and run the mesh checks (self.summary["mesh"]). Returns (V, F)."""
        cfg, h = self.cfg, self.cfg.h_x
        n_flip, it_max, n_unconv, n_exc, halo = amb
        t0 = time.time()
        if not faces:
            raise RuntimeError("export produced no surface")
        keys = np.concatenate(keys); pos = np.concatenate(pos); F = np.concatenate(faces)
        free = keys < 0
        keys[free] = -1 - np.arange(int(free.sum()))    # off-lattice-edge vertices: unique keys (not welded)
        uk, first, inv = np.unique(keys, return_index=True, return_inverse=True)
        V = pos[first]
        F = inv[F].astype(np.int64)
        del keys, pos, faces
        keep = (F[:, 0] != F[:, 1]) & (F[:, 1] != F[:, 2]) & (F[:, 0] != F[:, 2])
        n_degenerate = int((~keep).sum())          # should be 0 (skimage drops its own; weld keys are exact)
        if n_degenerate:
            F = F[keep]
        near_mount = None
        if self.mount_mesh is not None and len(self.mount_mesh[1]):
            from .surface import MeshQuery
            near_mount = MeshQuery(*self.mount_mesh).distance
        F, isl = drop_material_islands(V, F, self.V0, min_vol=10.0 * h ** 3, void_fill=cfg.void_min_mm3,
                                       near_mount=near_mount, mount_tol=2.0 * h)
        if isl.get("islands_on_mount", 0):
            self.log(f"[export] WARNING {isl['islands_on_mount']} dropped material island(s) "
                     f"({isl['islands_on_mount_mm3']:.3g} mm³) touch a load/support face")
        chk = mesh_checks(V, F)
        chk["degenerate_removed"] = n_degenerate
        chk["islands"] = isl
        chk["ambiguous_faces"] = dict(nodes_flipped=int(n_flip), flipped_mm3=float(n_flip * h ** 3),
                                      max_iterations=int(it_max), tiles_not_converged=int(n_unconv),
                                      tiles_halo_exceeded=int(n_exc), halo=int(halo))
        chk["volume_rel"] = chk["volume"] / self.V0 - 1.0
        chk["volume_ok"] = bool(abs(chk["volume_rel"]) <= cfg.vol_tol)
        if chk["faces"] <= cfg.trimesh_check_max_faces:
            import trimesh
            tm = trimesh.Trimesh(V.astype(np.float64), F, process=False)
            chk["trimesh"] = dict(watertight=bool(tm.is_watertight), winding_consistent=bool(tm.is_winding_consistent),
                                  volume=float(tm.volume), bodies=int(len(tm.split(only_watertight=False))))
        self.summary["mesh"] = chk
        self._t("E2_weld_checks", t0)
        return V, F

    def _write(self, V, F, out_dir, name):
        t0 = time.time()
        path = os.path.join(out_dir, f"{name}.stl")
        write_stl(path, V, F)
        self.summary["mesh"].update(stl=path, stl_bytes=os.path.getsize(path))
        self._t("E3_stl", t0)

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
        # a ray that leaves through an adjacent face (rims, hole bores, MC creases) is not a wall: keep opposite-wall
        # hits only (cos(n_hit, n) < -0.7). On B2 79 % of the sub-0.5 mm samples were such grazing exits.
        opp = w["cos"] < -0.7
        self.summary["wall_grazing_dropped"] = int((~opp).sum())
        w = {k: (v[opp] if isinstance(v, np.ndarray) and len(v) == len(opp) else v) for k, v in w.items()}
        t = w["t"]
        i_min = int(np.argmin(t))
        self.summary["wall_min_at"] = [float(x) for x in w["P"][i_min]]
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
        # samples on a hole bore / rim (within r + 2 h of a hole axis segment) are flagged and reported separately:
        # the rim of a hole through a thin wall is a legitimate sub-t_min feature (the hole itself), not a film
        rim = np.zeros(len(t), bool)
        H = getattr(self, "holes", None)
        if H is not None and len(H["p"]):
            for p_, n_, s0_, s1_, r_ in zip(H["p"], H["n"], H["s0"], H["s1"], H["r"]):
                d_ = w["P"] - p_
                sax = d_ @ n_
                rad = np.linalg.norm(d_ - sax[:, None] * n_, axis=1)
                rim |= (sax > s0_ - 2 * h) & (sax < s1_ + 2 * h) & (rad < r_ + 2 * h)
        off = ~rim
        self.summary["wall_off_rims"] = dict(
            n_rim=int(rim.sum()), min_as_built=float(t_ab[off].min()) if off.any() else None,
            n_below_t_min_minus_b_v=int((t_ab[off] < cfg.t_min - cfg.b_v).sum()),
            # walls sit at the t_min bound by design: count only those thinner than the export voxel resolution allows
            n_below_t_min_minus_b_v_disc=int((t_ab[off] < cfg.t_min - cfg.b_v - 2.5 * cfg.h_x).sum()))
        self.summary["wall"] = dict(samples=int(len(t)), miss=w["n_miss"], mesh_stride=stride, min_nominal=float(t.min()),
                                    p1_nominal=float(np.quantile(t, 0.01)), p05_nominal=float(np.quantile(t, 0.05)),
                                    median_nominal=float(np.median(t)), max_nominal=float(t.max()),
                                    min_as_built=float(t_ab.min()), p1_as_built=float(np.quantile(t_ab, 0.01)),
                                    frac_below_t_min=float((t < cfg.t_min).mean()))
        self._t("E4_wall_thickness", t0)

    # -- driver ----------------------------------------------------------------------------------------------------
    def run(self, out_dir: str, name: str = "final_design") -> dict:
        try:
            return self._run(out_dir, name)
        finally:
            _G.clear()                  # the tile store etc. would otherwise stay resident (and be forked) after the run

    def _run(self, out_dir: str, name: str) -> dict:
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
# GPU export (PLAN_single_run §2-§3): one GPU, a few GiB of device and host memory
# =====================================================================================================================
_GPU_SRC = r"""
extern "C" {
// Both branches of the nominal ψ per queried cell: o = φ − P d and ψ_in = −φ − P a. The interpolation is
// design.cell_psi's (clipped, normalised barycentrics, float64); ψ_in uses a_closed where every positive-weight vertex
// of the closest Γ_h triangle is closed (closure rule, §8.2). cmode: constant design (d_const, a_const).
// Cells with no surface within the query radius (phi >= far) are outside.
__global__ void exp_psi(long long n, const float* phi, const int* face, const float* bary, const int* Th,
        const double* d, const double* a, const double* cl, double a_closed, int cmode, double dconst, double aconst,
        float far, float* o, float* pin) {
    int v[3]; double w[3]; bool outer;
    for (long long c = blockIdx.x * (long long)blockDim.x + threadIdx.x; c < n;
         c += (long long)gridDim.x * blockDim.x) {
        double ph = (double)phi[c];
        if (ph >= (double)far) { o[c] = 1.0e30f; pin[c] = -1.0e30f; continue; }
        double oo, ii;
        if (cmode) {
            oo = ph - dconst; ii = -ph - aconst;
        } else {
            double psi = cell_psi(c, phi, face, bary, Th, d, a, 0.0, 0, &outer, v, w);
            if (outer) { oo = psi; ii = -ph - (((w[0] * a[v[0]]) + w[1] * a[v[1]]) + w[2] * a[v[2]]); }
            else       { ii = psi; oo = ph - (((w[0] * d[v[0]]) + w[1] * d[v[1]]) + w[2] * d[v[2]]); }
            double cc = ((w[0] * cl[v[0]]) + w[1] * cl[v[1]]) + w[2] * cl[v[2]];
            if (cc >= 1.0 - 1e-9) ii = -ph - a_closed;
        }
        o[c] = (float)oo; pin[c] = (float)ii;
    }
}

// Smooth opening of the void {s > 0} by a ball of radius r (PLAN_single_run §2). s = ψ_in on a region of m0×m1×m2
// cells; for the n0×n1×n2 core cells (offset H ≥ R in every axis):
//   ψ_open(x) = min(s(x), max_c [min(s(c), cap) − |x − c|]),  c ∈ lattice, s(c) ≥ r, |x − c| ≤ R h,
// offsets sorted by |k| (early exit once cap − |k| h ≤ best). With s the depth into the void (1-Lipschitz), the ball
// of radius s(c) ≥ r at c is inscribed, so {ψ_open > 0} is the union of inscribed balls of radius ≥ r: the opening.
// ψ_open is distance-like (slope 1) in removed thin void, so the removed region has no terraces. Cells with s ≥ r
// or s ≤ −2h are unchanged; floor_ (= −2h: continuous with them) where no centre is found.
__global__ void open_void(const float* s, int m1, int m2, int H, int n0, int n1, int n2, const int* off,
        const float* offd, int noff, float r, float cap, float h, float floor_, float* out) {
    long long N = (long long)n0 * n1 * n2;
    for (long long t = blockIdx.x * (long long)blockDim.x + threadIdx.x; t < N;
         t += (long long)gridDim.x * blockDim.x) {
        int x = (int)(t / ((long long)n1 * n2)), y = (int)((t / n2) % n1), z = (int)(t % n2);
        long long c = ((long long)(x + H) * m1 + (y + H)) * m2 + (z + H);
        float sx = s[c];
        if (sx >= r || sx <= -2.0f * h) { out[t] = sx; continue; }
        float best = floor_;
        for (int k = 0; k < noff; k++) {
            float dk = offd[k];
            if (cap - dk <= best) break;
            long long q = c + ((long long)off[3 * k] * m1 + off[3 * k + 1]) * m2 + off[3 * k + 2];
            float sc = s[q];
            if (sc >= r) {
                float val = fminf(sc, cap) - dk;
                if (val > best) best = val;
            }
        }
        out[t] = fminf(sx, best);
    }
}
}
"""
_GPU_MOD = None
_WK: dict = {}
QMAX = 32000                    # int16 quantisation of ψ ∈ [−Q, Q] (step Q/QMAX ≈ 2e-5 mm at h_x = 0.2 mm)


def _gpu_module():
    """design._KERNELS (cell_psi, the optimiser's fused interpolation) + the export kernels, one CuPy module."""
    global _GPU_MOD
    if _GPU_MOD is None:
        import cupy as cp
        from .design import _KERNELS
        _GPU_MOD = cp.RawModule(code=_KERNELS + _GPU_SRC, options=("--fmad=false",))
    return _GPU_MOD


def _warp_kernels():
    """Warp BVH query: signed distance φ to Γ₀ (winding-number sign, as band.py), the closest point on Γ₀ and, when
    Γ_h is given, its closest Γ_h triangle and barycentrics (u, v, 1 − u − v) for (Th[f,0], Th[f,1], Th[f,2])."""
    if _WK:
        return _WK
    import warp as wp
    wp.config.quiet = True
    if "SCRATCH" in os.environ and "WARP_CACHE_PATH" not in os.environ:
        wp.config.kernel_cache_dir = os.path.join(os.environ["SCRATCH"], "warp_cache")
    wp.init()

    @wp.kernel
    def q_cells(mesh0: wp.uint64, meshh: wp.uint64, has_h: int, cells: wp.array(dtype=wp.int32), lo: wp.vec3f,
                h: float, m1: int, m2: int, maxd: float, phi: wp.array(dtype=wp.float32),
                face: wp.array(dtype=wp.int32), bary: wp.array(dtype=wp.vec3f)):
        t = wp.tid()
        c = cells[t]
        i = c // (m1 * m2)
        j = (c // m2) % m1
        k = c % m2
        p = lo + wp.vec3f(float(i), float(j), float(k)) * h
        phi[t] = 1.0e8
        face[t] = 0
        bary[t] = wp.vec3f(1.0, 0.0, 0.0)
        q = wp.mesh_query_point_sign_winding_number(mesh0, p, maxd, 2.0, 0.5)
        if q.result:
            cp0 = wp.mesh_eval_position(mesh0, q.face, q.u, q.v)
            phi[t] = q.sign * wp.length(p - cp0)
            if has_h != 0:
                qh = wp.mesh_query_point_no_sign(meshh, cp0, 1.0e6)
                if qh.result:
                    face[t] = qh.face
                    bary[t] = wp.vec3f(qh.u, qh.v, 1.0 - qh.u - qh.v)

    _WK.update(wp=wp, q_cells=q_cells)
    return _WK


def _launch(n: int) -> tuple:
    import cupy as cp
    sm = cp.cuda.Device().attributes["MultiProcessorCount"]
    return (max(1, min(-(-int(n) // 256), 32 * sm)),), (256,)


def opening_R(r: float, h: float) -> int:
    """Search radius (cells) of the smooth opening: a cell 2h outside the void wall needs a centre with s ∈ [r, r + 2h]
    at distance ≈ r + 2h; 3 cells of lattice slack so that the 2h-thick centre shell is always reached (with
    ceil((r + 2h)/h) alone, cells just outside the wall found no centre and took the floor value)."""
    return int(math.ceil((r + 2.0 * h) / h)) + 3


def tess_volume(V: np.ndarray, T: np.ndarray) -> float:
    """Enclosed volume of a closed, outward-oriented tessellation (float64)."""
    a, b, c = (np.asarray(V, np.float64)[T[:, k]] for k in range(3))
    return float(np.einsum("ij,ij->i", a, np.cross(b, c)).sum()) / 6.0


class GpuExporter(Exporter):
    """Export of one `Design` on one GPU (PLAN_single_run §3); same stages and outputs as `Exporter`.

    A  per tile of n³ cells (+ the opening halo H ≥ R): blocks pruned by φ at their centres, then Warp BVH queries
       (φ, closest Γ_h triangle) and ψ's two branches (`exp_psi`, design.cell_psi); the void branch is opened with
       the smooth distance-based ψ (`open_void`, replaces the lattice opening and its constant −h/2 cells).
       Kept per tile (host): uint8 codes (0 material, 1 void, 2 outside, 3 dead), the outer-shell table
       S = (idx, φ, Γ_h triangle, b1, b2) for |o| ≤ Δ + 2h (re-evaluated for the mass shift δ) and the void-wall table
       P = (idx, ψ_open) for |ψ_open| < 3h. Nothing else: device and host memory stay at a few GiB.
    B  enclosed voids < void_min_mm3 filled, material islands marked dead (`Exporter._fill_islands`, on the codes).
    C  escape holes (`Exporter.stage_c`, unchanged).
    D  δ by Brent on the linear-ramp volume of ψ against V_target (default: the exact volume of the Γ₀ tessellation,
       i.e. the solid's mass), then after meshing ≤ 2 secant steps on the mesh volume (dV/dδ from the ramp volume)
       until |V_mesh/V_target − 1| ≤ mesh_vol_tol.
    E  ψ(δ) of every tile into one int16 lattice on the device; per tile: ambiguous lattice faces removed
       (`resolve_ambiguous_faces`, halo mc_halo), Warp marching cubes, triangles to the host with their lattice-edge
       weld keys; weld, island/void drop, checks and the binary STL as `Exporter`; wall thickness by ray casting."""

    def __init__(self, design: Design, cfg: ExportConfig, normals=None, mount_mesh=None, stress=None, r_ma=None,
                 lo_hi=None):
        import dataclasses as _dc
        cfg = _dc.replace(cfg, wall_max_faces=max(cfg.wall_max_faces, 1 << 40))     # never the CPU re-extraction
        super().__init__(design, cfg, normals, mount_mesh, stress, r_ma, lo_hi)
        if self.grid.n % cfg.block:
            raise ValueError("tile must be a multiple of block")
        self.workers = cfg.workers or 1
        self.Q = 3.0 * cfg.h_x
        self.S_out, self.P_in = {}, {}

    # -- device set-up ---------------------------------------------------------------------------------------------
    def _gpu_setup(self):
        import cupy as cp
        wk = _warp_kernels()
        wp = self._wp = wk["wp"]
        self._q = wk["q_cells"]
        dz = self.dz
        self.c0 = 0.5 * (dz.V0.min(0) + dz.V0.max(0))           # centred float32 frame for the BVH (as band.py)

        def mesh(V, T, wn):
            pts = wp.array(np.ascontiguousarray(V - self.c0, np.float32), dtype=wp.vec3f, device="cuda:0")
            idx = wp.array(np.ascontiguousarray(T, np.int32).ravel(), dtype=wp.int32, device="cuda:0")
            return wp.Mesh(pts, idx, support_winding_number=wn)
        self._m0 = mesh(dz.V0, dz.T0, True)
        self._mh = mesh(dz.Vh, dz.Th, False) if dz.Vh is not None else None
        mod = _gpu_module()
        self._k_psi, self._k_open = mod.get_function("exp_psi"), mod.get_function("open_void")
        if dz.Vh is not None:
            self._Th = cp.asarray(np.ascontiguousarray(dz.Th, np.int32))
            self._d = cp.asarray(np.clip(dz.d, 0.0, np.maximum(dz.d_hi, dz.d)), dtype=cp.float64)   # Design.d_at(δ=0)
            self._a = cp.asarray(dz.a_lin, dtype=cp.float64)
            self._cl = cp.asarray(dz.closed_f, dtype=cp.float64)
        else:
            self._Th = cp.zeros(3, cp.int32)
            self._d = self._a = self._cl = cp.zeros(1, cp.float64)

    def _query(self, cells, lo, h: float, m: int, maxd: float, want_h: bool):
        """φ, closest Γ_h triangle and barycentrics (CuPy) at the centres lo + h (i, j, k) of the region cells `cells`
        (int32 flat indices into an m³ box); lo in mm (global frame)."""
        import cupy as cp
        wp = self._wp
        N = int(cells.size)
        phi = wp.empty(N, dtype=wp.float32, device="cuda:0")
        face = wp.empty(N, dtype=wp.int32, device="cuda:0")
        bary = wp.empty(N, dtype=wp.vec3f, device="cuda:0")
        cells = cp.ascontiguousarray(cells, dtype=cp.int32)
        cp.cuda.Device().synchronize()
        mh = self._mh if (want_h and self._mh is not None) else self._m0
        wp.launch(self._q, dim=N, inputs=[self._m0.id, mh.id, int(want_h and self._mh is not None),
                                          wp.from_dlpack(cells), wp.vec3f(*(np.asarray(lo) - self.c0).tolist()),
                                          float(h), int(m), int(m), float(maxd), phi, face, bary], device="cuda:0")
        wp.synchronize_device("cuda:0")
        out = (cp.asarray(phi).copy(), cp.asarray(face).copy(), cp.asarray(bary).reshape(N, 3).copy())
        del phi, face, bary
        return out

    # -- stage A ---------------------------------------------------------------------------------------------------
    def stage_a(self, delta_shell: float):
        import cupy as cp
        t0 = time.time()
        cfg, dz, g = self.cfg, self.dz, self.grid
        h, n, b, Q = g.h, g.n, cfg.block, self.Q
        self._gpu_setup()
        self.delta_shell = delta_shell
        self.store = TileStore(g)
        self.S_out, self.P_in = {}, {}
        opening = bool(cfg.clean and dz.has_void)
        r = 0.5 * cfg.v_min
        R = opening_R(r, h) if opening else 0
        Hw = -(-R // (b // 2)) * (b // 2)                     # ≥ R, and (n + 2 Hw) a multiple of the block
        m = n + 2 * Hw
        nb = m // b
        rng_ = np.arange(-R, R + 1)
        D = np.stack(np.meshgrid(rng_, rng_, rng_, indexing="ij"), -1).reshape(-1, 3)
        q2 = (D * D).sum(1)
        sel = q2 <= R * R
        order = np.argsort(q2[sel], kind="stable")
        off = cp.asarray(np.ascontiguousarray(D[sel][order], np.int32))
        offd = cp.asarray((np.sqrt(q2[sel][order]) * h).astype(np.float32))
        cap, floor_ = r + 2.0 * h, -2.0 * h   # −2h: 1-Lipschitz next to unchanged cells
        ext = dz.V0.max(0) - dz.V0.min(0)
        band_out = dz.d_out_max + delta_shell + 3.0 * h       # beyond: o > Δ + 2h for every δ in the bracket
        maxd = max(band_out + h, 0.5 * float(ext.min()) + 2.0 * h)   # interior points lie within ½ min extent of Γ₀
        far = np.float32(1.0e7)
        S_thr = delta_shell + 2.0 * h
        Rb = 0.5 * b * h * math.sqrt(3.0)
        blk_cells = cp.arange(nb ** 3, dtype=cp.int32)
        core = (slice(Hw, Hw + n),) * 3
        nq, v0r, void_raw, void_open, n_open_tiles = 0, 0.0, 0, 0, 0
        for tid in g.tile_ids():
            lo_core = g.tile_lo(tid) + 0.5 * h                # centre of core cell (0, 0, 0)
            lo_reg = lo_core - Hw * h
            phb = self._query(blk_cells, lo_reg + 0.5 * (b - 1) * h, b * h, nb, 1.0e6, False)[0]
            need = (phb - Rb <= band_out).reshape(nb, nb, nb)
            if not bool(need.any()):
                self.store.set(tid, 2)
                continue
            mask = cp.repeat(cp.repeat(cp.repeat(need, b, 0), b, 1), b, 2)
            if not bool(mask[core].any()):
                self.store.set(tid, 2)
                continue
            cells = cp.flatnonzero(mask.ravel()).astype(cp.int32)
            del mask
            phi, face, bary = self._query(cells, lo_reg, h, m, maxd, True)
            o_c, i_c = cp.empty(cells.size, cp.float32), cp.empty(cells.size, cp.float32)
            self._k_psi(*_launch(cells.size), (
                np.int64(cells.size), phi, face, bary, self._Th, self._d, self._a, self._cl, np.float64(A_CLOSED),
                np.int32(dz.Vh is None), np.float64(max(dz.d_const, 0.0)), np.float64(dz.a_const), far, o_c, i_c))
            ci, cj, ck = cells // (m * m), (cells // m) % m, cells % m
            inc = (ci >= Hw) & (ci < Hw + n) & (cj >= Hw) & (cj < Hw + n) & (ck >= Hw) & (ck < Hw + n)
            lidx = (((ci - Hw) * n + (cj - Hw)) * n + (ck - Hw)).astype(cp.int32)
            nq += int(cells.size)
            v0r += float(cp.clip(0.5 - phi[inc].astype(cp.float64) / h, 0.0, 1.0).sum())
            ss = inc & (cp.abs(o_c) <= S_thr)                 # outer-shell table
            if bool(ss.any()):
                bb = cp.maximum(bary[ss].astype(cp.float64), 0.0)
                bb /= cp.maximum(bb.sum(1, keepdims=True), 1e-300)
                self.S_out[tid] = (lidx[ss].get(), phi[ss].get(), face[ss].get(),
                                   bb[:, 1].astype(cp.float32).get(), bb[:, 2].astype(cp.float32).get())
            O = cp.full(m ** 3, 1.0e30, cp.float32)
            I = cp.full(m ** 3, -1.0e30, cp.float32)
            O[cells] = o_c
            I[cells] = i_c
            del phi, face, bary, o_c, i_c, ci, cj, ck, inc, lidx, cells
            O, I = O.reshape(m, m, m), I.reshape(m, m, m)
            Ic = I[core]
            void_raw += int((Ic > 0).sum())
            if opening and bool((Ic > -2.0 * h).any()):
                P = cp.empty(n ** 3, cp.float32)
                self._k_open(*_launch(n ** 3), (
                    I, np.int32(m), np.int32(m), np.int32(Hw), np.int32(n), np.int32(n), np.int32(n), off, offd,
                    np.int32(len(offd)), np.float32(r), np.float32(cap), np.float32(h), np.float32(floor_), P))
                P = P.reshape(n, n, n)
                n_open_tiles += 1
            else:
                P = cp.ascontiguousarray(Ic)
            void_open += int((P > 0).sum())
            code = cp.where(O[core] > 0, 2, cp.where(P > 0, 1, 0)).astype(cp.uint8)
            near = cp.flatnonzero((cp.abs(P) < Q).ravel())
            if near.size:
                self.P_in[tid] = (near.astype(cp.int32).get(), P.ravel()[near].get())
            self.store.set(tid, code.get())
            del O, I, P, code, near
        cp.get_default_memory_pool().free_all_blocks()
        self._void_cells = (void_raw, void_open)
        self.V0_ramp = v0r * h ** 3
        self.V0 = float(cfg.target_volume) if cfg.target_volume is not None else tess_volume(dz.V0, dz.T0)
        self.summary.update(V0_mm3=self.V0, V0_ramp_mm3=self.V0_ramp, V0_ramp_rel=self.V0_ramp / self.V0 - 1.0,
                            n_tiles=len(g.tile_ids()), n_cell_queries=nq, opening_halo=Hw, opening_R_cells=R,
                            n_opened_tiles=n_open_tiles,
                            n_shell_cells=int(sum(len(s[0]) for s in self.S_out.values())),
                            n_void_wall_cells=int(sum(len(p[0]) for p in self.P_in.values())))
        self._t("A_geometry_gpu", t0)

    # -- stage B ---------------------------------------------------------------------------------------------------
    def stage_b(self):
        h = self.cfg.h_x
        tids = [t for t in self.grid.tile_ids() if self.store.scalar(t) != 2]
        raw, opened = self._void_cells
        rep = dict(opening="smooth distance-based ψ (GPU)", void_raw_mm3=raw * h ** 3,
                   opening_removed_mm3=(raw - opened) * h ** 3)
        self._fill_islands(tids, rep)
        self.summary["cleanup"] = rep
        self.log(f"[export] cleanup: {rep}")
        return tids

    # -- ψ of a tile from codes and tables ---------------------------------------------------------------------------
    def _parts(self, tid, delta: float):
        """(o, ψ_in) dense float32 (n³, clipped to ±Q) of tile `tid` at mass shift δ, escape holes included."""
        n, Q = self.grid.n, self.Q
        code = self.store.get(tid).reshape(-1)
        o = np.where(code >= 2, Q, -Q).astype(np.float32)
        pin = np.where(code == 1, Q, -Q).astype(np.float32)
        P = self.P_in.get(tid)
        if P is not None:
            idx, val = P
            pin[idx] = np.where((code[idx] == 0) & (val > 0), -Q, np.clip(val, -Q, Q))   # filled voids stay filled
        S = self.S_out.get(tid)
        if S is not None:
            idx, phi, tri, b1, b2 = S
            oo = phi.astype(np.float64) - self.dz.d_at(tri, b1, b2, delta)
            o[idx] = np.clip(np.where(code[idx] == 3, Q, oo), -Q, Q)
        pin = pin.reshape(n, n, n)
        _hole_psi_into(pin, self.grid, tid, getattr(self, "holes", None), self.grid.h)
        return o.reshape(n, n, n), pin

    def _uniform(self, tid):
        """Code of a tile whose ψ is ±Q everywhere (scalar codes, no table entries, no hole), else None."""
        s = self.store.scalar(tid)
        if s is None or tid in self.S_out or tid in self.P_in or _tile_has_hole(tid):
            return None
        return s

    # -- stage D ---------------------------------------------------------------------------------------------------
    def stage_d(self):
        cfg, dz, h, n, Q = self.cfg, self.dz, self.cfg.h_x, self.grid.n, self.Q
        t0 = time.time()
        self._setG()
        Vc = 0.0
        arrs = []
        for tid in self.grid.tile_ids():
            u = self._uniform(tid)
            if u is not None:
                Vc += float(n ** 3) if u == 0 else 0.0
                continue
            o, pin = self._parts(tid, 0.0)
            f = ramp(-np.maximum(o, pin), h).reshape(-1)
            S = self.S_out.get(tid)
            if S is None:
                Vc += float(f.sum(dtype=np.float64))
                continue
            idx = S[0]
            Vc += float(f.sum(dtype=np.float64) - f[idx].sum(dtype=np.float64))
            pS = pin.reshape(-1)[idx].copy()
            pS[self.store.get(tid).reshape(-1)[idx] == 3] = Q       # dead material never counts
            arrs.append((S[1], S[2], S[3], S[4], pS))
        if arrs:
            phi = np.concatenate([a[0] for a in arrs]).astype(np.float64); tri = np.concatenate([a[1] for a in arrs])
            b1 = np.concatenate([a[2] for a in arrs]); b2 = np.concatenate([a[3] for a in arrs])
            pS = np.concatenate([a[4] for a in arrs]).astype(np.float64)
        else:
            phi = pS = np.zeros(0); tri = np.zeros(0, np.int32); b1 = b2 = np.zeros(0, np.float32)
        del arrs

        def V_of(delta):
            oo = phi - dz.d_at(tri, b1, b2, delta)
            return (Vc + float(ramp(-np.maximum(oo, pS), h).sum())) * h ** 3
        self._V_of = V_of
        v00 = V_of(0.0)
        rel0 = v00 / self.V0 - 1.0
        out = dict(V_target_mm3=self.V0, V_before_mm3=v00, rel_before=rel0)
        delta, ok = 0.0, True
        if cfg.mass_correct and not dz.trivial_solid and abs(rel0) > cfg.mass_tol:
            lim = self.delta_shell
            fl, fh = V_of(-lim) / self.V0 - 1.0, V_of(lim) / self.V0 - 1.0
            if fl * fh > 0:
                out["bracket_failed"] = dict(limit=lim, f_lo=fl, f_hi=fh)
                ok = False
            else:
                delta = brentq(lambda x: V_of(x) / self.V0 - 1.0, -lim, lim, xtol=1e-7, rtol=1e-10)
        out.update(delta=delta, V_after_mm3=V_of(delta), rel_after=V_of(delta) / self.V0 - 1.0,
                   warn=bool(abs(delta) > cfg.delta_warn), fail=bool(abs(delta) > cfg.delta_fail),
                   n_brent_cells=int(len(phi)), ok=ok)
        self.delta = delta
        self.summary["mass_correction"] = out
        self._t("D_mass", t0)
        return ok

    # -- stage E ---------------------------------------------------------------------------------------------------
    def _mc_tiles(self, delta):
        import cupy as cp
        from warp.geometry import IsoSurfaceMarchingCubes
        wp = self._wp
        g, cfg = self.grid, self.cfg
        n, h, Q = g.n, g.h, self.Q
        qs = Q / QMAX
        H = int(min(cfg.mc_halo, n))
        self._setG()
        # ψ(δ) on every lattice node, int16 on the device; node i at index i + H (padding: outside)
        L = cp.full(tuple(int(s) + 2 * H + 1 for s in g.shape), QMAX, cp.int16)
        for tid in g.tile_ids():
            sl = tuple(slice(H + tid[a] * n, H + tid[a] * n + n) for a in range(3))
            u = self._uniform(tid)
            if u is not None:
                if u == 0:
                    L[sl] = -QMAX
                continue
            o, pin = self._parts(tid, delta)
            L[sl] = cp.asarray(np.rint(np.clip(np.maximum(o, pin), -Q, Q) / qs).astype(np.int16))
        eps = np.float32(1e-2 * h)                     # crossings strictly inside lattice edges (robust weld keys)
        keys, pos, faces = [], [], []
        base_v, n_flip, it_max, n_unconv, n_exc = 0, 0, 0, 0, 0
        nk = np.asarray(g.shape, np.int64) + 1
        for tid in g.tile_ids():
            i0 = [tid[a] * n for a in range(3)]          # = L index of global node i0 − H
            blk = L[i0[0]:i0[0] + n + 1 + 2 * H, i0[1]:i0[1] + n + 1 + 2 * H, i0[2]:i0[2] + n + 1 + 2 * H]
            core = blk[H:H + n + 1, H:H + n + 1, H:H + n + 1]
            if not (int(core.min()) < 0 < int(core.max())):
                continue
            psi = blk.astype(cp.float32) * np.float32(qs)
            psi = cp.where(cp.abs(psi) < eps, cp.where(psi < 0, -eps, eps), psi).astype(cp.float32)
            nf_, it_ = resolve_ambiguous_faces(psi, eps)
            n_flip += nf_; it_max = max(it_max, abs(it_)); n_unconv += it_ < 0; n_exc += abs(it_) >= H
            psi = cp.ascontiguousarray(psi[H:H + n + 1, H:H + n + 1, H:H + n + 1])
            if not (float(psi.min()) < 0 < float(psi.max())):
                continue
            cp.cuda.Device().synchronize()
            verts, tris = IsoSurfaceMarchingCubes.extract(wp.from_dlpack(psi), 0.0)
            wp.synchronize_device("cuda:0")
            v = verts.numpy().astype(np.float64)
            f = tris.numpy().reshape(-1, 3).astype(np.int64)
            del verts, tris, psi
            if not len(f):
                continue
            fr = np.abs(v - np.rint(v))
            axis = fr.argmax(1)
            lat = np.rint(v).astype(np.int64)
            ar = np.arange(len(v))
            lat[ar, axis] = np.floor(v[ar, axis]).astype(np.int64)
            lat += np.asarray(i0, np.int64)
            k = ((lat[:, 0] * nk[1] + lat[:, 1]) * nk[2] + lat[:, 2]) * 3 + axis
            k[(fr > 1e-5).sum(1) != 1] = -1               # not on a lattice edge: never welded
            pos.append((g.origin + (np.asarray(i0) + v + 0.5) * h).astype(np.float32))
            keys.append(k)
            faces.append(f + base_v)
            base_v += len(v)
        del L
        cp.get_default_memory_pool().free_all_blocks()
        # orientation: material is ψ < 0; Warp's winding is fixed, so one global check suffices
        vol = 0.0
        if faces:
            P_ = np.concatenate(pos).astype(np.float64)
            for f_ in faces:
                a_, b_, c_ = P_[f_[:, 0]], P_[f_[:, 1]], P_[f_[:, 2]]
                vol += float(np.einsum("ij,ij->i", a_, np.cross(b_, c_)).sum()) / 6.0
            del P_
            if vol < 0:
                faces = [f_[:, ::-1].copy() for f_ in faces]
        return pos, keys, faces, (n_flip, it_max, n_unconv, n_exc, H)

    # -- driver ----------------------------------------------------------------------------------------------------
    def _run(self, out_dir: str, name: str) -> dict:
        cfg, dz = self.cfg, self.dz
        os.makedirs(out_dir, exist_ok=True)
        t_all = time.time()
        self.stage_a(cfg.delta_retry)                     # one pass: the shell table covers |δ| ≤ delta_retry
        clean = cfg.clean and dz.has_void
        tids = self.stage_b() if clean else [t for t in self.grid.tile_ids() if self.store.scalar(t) != 2]
        if not clean:
            self.summary["cleanup"] = {}
        if cfg.holes and clean:
            self.stage_c(tids)
        ok = self.stage_d()
        if getattr(self, "holes", None) is not None and len(self.holes["p"]):
            with open(os.path.join(out_dir, "escape_holes.csv"), "w", newline="") as f:
                wr = csv.writer(f)
                wr.writerow(["void_id", "hole", "x", "y", "z", "nx", "ny", "nz", "wall_mm", "score", "score_kind",
                             "void_volume_mm3"])
                wr.writerows(self.holes_rows)
        steps = []
        for it in range(3):
            t0 = time.time()
            mc = self._mc_tiles(self.delta)
            self._t(f"E1_marching_cubes_{it}", t0)
            V, F = self._weld_checks(*mc)
            del mc
            rel = self.summary["mesh"]["volume_rel"]
            steps.append(dict(delta=self.delta, rel_mesh=rel))
            if not (cfg.mass_correct and ok) or dz.trivial_solid or abs(rel) <= cfg.mesh_vol_tol or it == 2:
                break
            e = 0.25 * cfg.h_x
            slope = (self._V_of(self.delta + e) - self._V_of(self.delta - e)) / (2.0 * e)
            if not slope > 0:
                break
            nd = self.delta - (self.summary["mesh"]["volume"] - self.V0) / slope
            if abs(nd) > self.delta_shell:                 # outside the shell table: keep the Brent δ
                break
            self.delta = nd
        mc_ = self.summary["mass_correction"]
        mc_.update(delta=self.delta, mesh_steps=steps, rel_mesh=steps[-1]["rel_mesh"],
                   V_ramp_final_mm3=self._V_of(self.delta),           # ψ volume at the final δ (vs the mesh volume)
                   warn=bool(abs(self.delta) > cfg.delta_warn), fail=bool(abs(self.delta) > cfg.delta_fail))
        self.mesh = (V, F)
        if cfg.write_stl:
            self._write(V, F, out_dir, name)
        if cfg.wall_hist:
            self.wall_stats(V, F, out_dir)
        self.timing["total"] = round(time.time() - t_all, 2)
        ru = resource.getrusage(resource.RUSAGE_SELF)
        self.summary.update(h_x=cfg.h_x, tile=cfg.tile, workers=self.workers, timing_s=self.timing, device="gpu",
                            peak_rss_parent_GB=ru.ru_maxrss / 1e6, grid_shape=[int(x) for x in self.grid.shape])
        ok = (mc_["ok"] and not mc_["fail"] and self.summary["mesh"]["watertight_manifold"]
              and self.summary["mesh"]["volume_ok"])
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


def export_design_gpu(ctx, xi, out_dir: str, D_comp: float = 4.0, stress=None, name: str = "final_design", h_g=None,
                      fields: dict | None = None, closed_thr: float | None = None, **cfg_kw) -> dict:
    """GPU export (`GpuExporter`) of the nominal design xi of `ctx`: <name>.stl, escape_holes.csv,
    wall_thickness.csv, <name>_export.json. cfg_kw: ExportConfig fields (h_x, tile, mc_halo, log, ...); `fields`:
    precomputed `vertex_fields`; closed_thr: ξ̃ᵃ at and above which a vertex is exactly closed. Returns the summary."""
    cfg = ExportConfig.from_run_cfg(ctx.cfg, **cfg_kw)
    f = fields if fields is not None else vertex_fields(ctx, xi, D_comp, h_g)
    s = ctx.surf
    closed = f["xt_a"] >= (cfg.closed_thr if closed_thr is None else closed_thr)
    dz = Design(ctx.V, ctx.T, s.V, s.T, f["d"], f["a"], f["d_hi"], active=~f["passive"], closed=closed, sign=cfg.sign)
    N = s.projected_normals() if hasattr(s, "projected_normals") else s.N
    mount_mesh = None
    if len(getattr(ctx, "mount_ids", [])):
        sel = np.isin(ctx.TF, np.asarray(ctx.mount_ids))
        if sel.any():
            mount_mesh = (ctx.V, ctx.T[sel])
    ex = GpuExporter(dz, cfg, normals=N, mount_mesh=mount_mesh, stress=stress, r_ma=f["r_ma"])
    return ex.run(out_dir, name)


def uniform_shell_fields(ctx, t: float, h_g: float) -> dict:
    """Vertex fields of the uniform-shell baseline on Γ_h (THEORY §7.4 S1, Owen 2026-10-30 D1): d = t, a = 0 on the
    free surface; mount (load/support) vertices kept solid as in the optimiser (d = 0, closed, a = a^hi = r_MA + γ_c h_g).
    `active` vertices carry the mass correction δ."""
    nv = ctx.surf.n_vertices
    mount = np.asarray(getattr(ctx, "mount", np.zeros(nv, bool)), bool)
    a_hi = ctx.r_ma + ctx.cfg["design"]["closure_gamma"] * h_g
    return dict(d=np.where(mount, 0.0, float(t)), a=np.where(mount, a_hi, 0.0), closed=mount.copy(), active=~mount,
                d_hi=np.full(nv, float(t) + 0.5), mount=mount)   # ≥ t + δ_retry


def export_baselines(ctx, out_dir: str, t_coarse: float | None = None, h_g: float | None = None, **cfg_kw) -> dict:
    """original.stl (ψ = φ) and uniform_shell.stl, t for constant mass (§7.4 S1). With mount vertices (ctx.mount) the
    shell keeps them solid (D1): ψ = max(φ − Pd, −φ − Pa) of `uniform_shell_fields`, t = `t_coarse` (collect solves
    it on the optimisation band) refined by the export's mass correction. Without mounts: ψ = max(φ − t, −φ)."""
    cfg = ExportConfig.from_run_cfg(ctx.cfg, **cfg_kw)
    mount = np.asarray(getattr(ctx, "mount", np.zeros(0, bool)), bool)
    if not mount.any():
        return baselines(ctx.V, ctx.T, out_dir, cfg, t_coarse=t_coarse)
    if h_g is None:
        h_g = min(p["h_g"] for p in ctx.cfg["grid"]["phases"])
    if t_coarse is None:
        raise ValueError("export_baselines: a part with mounts needs t_coarse (solved on the band by collect)")
    s = ctx.surf
    f = uniform_shell_fields(ctx, t_coarse, h_g)
    dz = Design(ctx.V, ctx.T, s.V, s.T, f["d"], f["a"], f["d_hi"], active=f["active"], closed=f["closed"],
                sign=cfg.sign)
    return baselines(ctx.V, ctx.T, out_dir, cfg, t_coarse=t_coarse, shell_design=dz)


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
    try:
        res = _pmap(_hist_job, grid.tile_ids(), workers or len(os.sched_getaffinity(0)))
    finally:
        _G.clear()
        _JOB = None
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


def baselines(V0, T0, out_dir: str, cfg: ExportConfig, t_coarse: float | None = None, shell_design=None) -> dict:
    """original and uniform shell. `shell_design`: a Γ_h `Design` at t_coarse (mounts solid) instead of the constant
    shell ψ = max(φ − t, −φ)."""
    os.makedirs(out_dir, exist_ok=True)
    out = {}
    cl = dict(clean=False, holes=False)
    # original
    c = ExportConfig(**{**asdict(cfg), **cl, "log": cfg.log})
    ex = Exporter(Design(V0, T0, d_const=0.0, a_const=A_CLOSED, sign=cfg.sign), c)
    out["original"] = ex.run(out_dir, "original")
    # uniform shell
    t0 = time.time()
    tc = t_coarse
    if tc is None:
        tc = solve_shell_thickness(V0, T0, h_c=max(cfg.h_x, 0.25), sign=cfg.sign, workers=cfg.workers)
        cfg.log(f"[export] uniform shell: coarse t = {tc:.4f} mm ({time.time() - t0:.1f} s)")
    else:
        cfg.log(f"[export] uniform shell: t = {tc:.4f} mm (given){', mounts solid' if shell_design is not None else ''}")
    dz = shell_design if shell_design is not None else Design(V0, T0, d_const=tc, a_const=0.0, sign=cfg.sign)
    ex = Exporter(dz, c)
    s = ex.run(out_dir, "uniform_shell")
    s["t_coarse_mm"], s["t_final_mm"] = tc, tc + s["mass_correction"]["delta"]
    s["mounts_solid"] = shell_design is not None
    out["uniform_shell"] = s
    return out
