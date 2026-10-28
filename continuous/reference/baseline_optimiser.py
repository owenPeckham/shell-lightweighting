"""
The shared code of every optimiser (docs/theory.md §8.1, §12.1). Three classes derive from one neutral base:

    BaseOptimiser                       this file: parameters, run, outputs, diagnostics, buckling bookkeeping
      |-- PCLS_Optimiser                pcls_optimiser.py: the piecewise constant level set method
      `-- BaselineOptimiser             this file: the iteration skeleton of the comparison baselines
            |-- SIMP_Optimiser          simp_optimiser.py
            `-- LevelSet_Optimiser      levelset_optimiser.py

so no optimiser is a special case of another: PCLS "acts identically" to SIMP and the level set wherever the
method does not matter. The file holds, in this order,

  1. the helper blocks the run needs (formerly separate modules): design diagnostics (topology, nucleation,
     thickness; `RunDiagnostics`), the discreteness/crispness metrics, the surface export (alpha wrap, marching
     cubes, STL quality, the background job) and the run provenance (docs/theory.md §8.4);
  2. `BaseOptimiser`;
  3. `BaselineOptimiser`, the loop skeleton the two baselines share.

BaseOptimiser (docs/theory.md §12.1). Everything a run writes -- variables.csv, nucleation_events.csv, history/, the
saved .vtr files, the guaranteed final save and surface export, the exported-design analysis, outcome_stats.csv,
final_fields.npz, run_info.json, threshold_sweep.csv and the postprocess section below -- is produced by BaseOptimiser's methods,
which are written in terms of the design field phi in [V, S] = [1, 2]. A method stores its own design in that form
(eq. 12.5):

    PCLS:       phi itself, updated by the augmented-Lagrangian step
    SIMP:       phi = 1 + x_phys            (x_phys the physical density in [0, 1])
    level set:  phi = 1.5 + clip(psi, -1/2, 1/2)   (psi the signed distance, voxels)

so that rho(phi) = c1 + (1 - c1) x_phys is the design density (SIMP's x, the level set's ersatz volume fraction),
phi >= 1.5 is the exported geometry, and the penalised model rho_m = c1 + (1 - c1) x^p (`_rho_model`, with
current_penalisation_power set to SIMP's p, or 1 for the level set) is exactly the stiffness each method analyses.
A method provides `_setup_phi` (the initial design) and `_outer_loop` (the iteration), plus `_variables_na` for the
variables.csv columns it has no quantity for. The parameters shared by every method (COMMON_PARAMETERS) have one set
of names, meanings, defaults and validation (`initialise_common_parameters`).

BaselineOptimiser. Problem files select the method with optimiser.method ("pcls" when absent) and carry the
method's own parameters in optimiser.simp / optimiser.levelset; optimise.py dispatches (docs/theory.md §8.1).
PCLS-only parameters are not accepted by a baseline (optimise.py drops them with a warning). Only the density
backends ("reg", "cg") are supported: a baseline hands the backend a stiffness density per voxel. Conformal
("conf") meshes a level set of phi and is rejected at start-up.

The linear buckling constraint (docs/theory.md §6, §12.6) is shared by both baselines: BaselineOptimiser owns the
state, the evaluation after each analysis (`_evaluate_buckling`: activation, safety-factor ramp, cadence, the
eigensolve, the constraint value G3, buckling.csv, the three-way status handling), the buckling part of the stopping
test and the augmented-Lagrangian multiplier update (`_update_buckling_multipliers`) for a method that uses one. How
the sensitivity moves the design is the method's: SIMP passes it to MMA as a constraint gradient, the level set adds
it to the front speed.
"""
import os
import inspect
import numpy as np
from typing import Callable
from scipy import ndimage
from interfaces import AnalysisInterface, BucklingAnalysisInterface, BucklingResults
from timeit import default_timer as timer
import json
import traceback
from collections import deque
import io
import sys
import time
import struct
import socket
import hashlib
import platform
import subprocess
from contextlib import redirect_stdout
from datetime import datetime, timezone
from scipy.sparse import coo_matrix
from scipy.sparse.csgraph import connected_components


# ====================================================================================================
# Design diagnostics: topology, nucleation, thickness (formerly diagnostics.py)
# ====================================================================================================
# Geometric and topological diagnostics of a PCLS design (docs/theory.md §8.4.1).
#
# Everything here works on voxel arrays indexed [x, y, z] and needs only NumPy
# and SciPy, so it can run inside the optimisation loop (baseline_optimiser.py calls
# `RunDiagnostics.update` once per iteration) as well as afterwards
# (the postprocess section below). The definitions follow the browser demonstrator's
# (docs/theory.md §11.7):
#
# - The solid is a union of CLOSED unit cubes, so two solid voxels that share a
#   face, an edge or a corner are joined (26-connectivity) and two void voxels
#   are joined only through a shared face (6-connectivity). The space outside
#   the domain counts as void.
# - Euler characteristic of that cubical complex, chi = V - E + F - C, and the
#   Betti numbers b0 (pieces), b2 (enclosed cavities) and b1 = b0 + b2 - chi
#   (tunnels, i.e. handles: a hole through a plate or a closed ring each add 1).
# - A NUCLEATION is a new void cluster (6-connected) none of whose voxels has a
#   26-neighbour that was void in the previous iteration. It is INTERIOR if it
#   also does not touch the skin of the domain, otherwise a SURFACE pit.

S26 = ndimage.generate_binary_structure(3, 3)   # face, edge or corner contact
S6 = ndimage.generate_binary_structure(3, 1)    # face contact only


# ---------------------------------------------------------------------------
# Topology
# ---------------------------------------------------------------------------

def euler_characteristic(solid: np.ndarray) -> int:
    """
    chi = V - E + F - C of the union of closed unit cubes at the solid voxels.

    A vertex (edge, face) of the grid belongs to the complex if any of the 8
    (4, 2) voxels around it is solid. Every count is a vectorised OR of
    shifted views of the zero-padded array, so the cost is a few dozen passes
    over the grid.
    """
    P = np.pad(np.asarray(solid, dtype=bool), 1)
    C = int(np.count_nonzero(P))
    if C == 0:
        return 0
    # Faces normal to each axis: the two voxels on either side.
    F = (np.count_nonzero(P[:-1, 1:-1, 1:-1] | P[1:, 1:-1, 1:-1])
         + np.count_nonzero(P[1:-1, :-1, 1:-1] | P[1:-1, 1:, 1:-1])
         + np.count_nonzero(P[1:-1, 1:-1, :-1] | P[1:-1, 1:-1, 1:]))
    # Edges parallel to each axis: the four voxels around them.
    E = (np.count_nonzero(P[1:-1, :-1, :-1] | P[1:-1, 1:, :-1] | P[1:-1, :-1, 1:] | P[1:-1, 1:, 1:])
         + np.count_nonzero(P[:-1, 1:-1, :-1] | P[1:, 1:-1, :-1] | P[:-1, 1:-1, 1:] | P[1:, 1:-1, 1:])
         + np.count_nonzero(P[:-1, :-1, 1:-1] | P[1:, :-1, 1:-1] | P[:-1, 1:, 1:-1] | P[1:, 1:, 1:-1]))
    # Vertices: the eight voxels around them.
    Vx = (P[:-1, :-1, :-1] | P[1:, :-1, :-1] | P[:-1, 1:, :-1] | P[1:, 1:, :-1]
          | P[:-1, :-1, 1:] | P[1:, :-1, 1:] | P[:-1, 1:, 1:] | P[1:, 1:, 1:])
    V = np.count_nonzero(Vx)
    return int(V - E + F - C)


def count_cavities(solid: np.ndarray) -> int:
    """b2: 6-connected void components that do not reach the outside of the domain."""
    void = ~np.pad(np.asarray(solid, dtype=bool), 1)     # the padding is the outside
    labels, n = ndimage.label(void, structure=S6)
    if n == 0:
        return 0
    outside = labels[0, 0, 0]
    return int(n - (1 if outside > 0 else 0))


def topology(solid: np.ndarray, n_components: int | None = None) -> dict:
    """chi, b0, b1, b2 of the solid (26/6 digital topology, eq. 11.19)."""
    solid = np.asarray(solid, dtype=bool)
    chi = euler_characteristic(solid)
    if n_components is None:
        _, n_components = ndimage.label(solid, structure=S26)
    b0 = int(n_components)
    b2 = count_cavities(solid)
    return dict(euler=chi, b0=b0, b1=int(b0 + b2 - chi), b2=b2)


def interface_faces(solid: np.ndarray) -> int:
    """Voxel faces separating solid from void, including faces on the domain skin."""
    P = np.pad(np.asarray(solid, dtype=bool), 1)
    return int(np.count_nonzero(P[1:] != P[:-1]) + np.count_nonzero(P[:, 1:] != P[:, :-1])
               + np.count_nonzero(P[:, :, 1:] != P[:, :, :-1]))


def skin_mask(shape) -> np.ndarray:
    """The outer layer of voxels of the domain."""
    m = np.zeros(shape, dtype=bool)
    m[0], m[-1] = True, True
    m[:, 0], m[:, -1] = True, True
    m[:, :, 0], m[:, :, -1] = True, True
    return m


def component_sizes(solid: np.ndarray):
    """(labels, sizes) of the 26-connected solid components, sizes in voxels, largest first."""
    labels, n = ndimage.label(np.asarray(solid, dtype=bool), structure=S26)
    if n == 0:
        return labels, np.zeros(0, dtype=np.int64)
    sizes = np.bincount(labels.ravel())[1:]
    return labels, np.sort(sizes)[::-1]


# ---------------------------------------------------------------------------
# Discreteness
# ---------------------------------------------------------------------------

def non_discreteness(rho: np.ndarray, mask: np.ndarray | None = None) -> float:
    """
    Measure of non-discreteness M_nd = 4 mean(rho (1 - rho)) (Sigmund 2007):
    0 for a 0/1 design, 1 for a uniform rho = 0.5. rho is clipped to [0, 1].
    """
    r = np.clip(np.asarray(rho, dtype=np.float64), 0.0, 1.0)
    if mask is not None:
        r = r[np.asarray(mask, dtype=bool)]
    return float(4.0 * np.mean(r * (1.0 - r))) if r.size else 0.0


def band_fraction(phi: np.ndarray, lo: float, hi: float, mask: np.ndarray | None = None) -> float:
    """Fraction of (masked) voxels with lo <= phi < hi."""
    p = np.asarray(phi)
    sel = (p >= lo) & (p < hi)
    if mask is not None:
        m = np.asarray(mask, dtype=bool)
        return float(np.count_nonzero(sel & m)) / max(1, int(np.count_nonzero(m)))
    return float(np.count_nonzero(sel)) / max(1, p.size)


# ---------------------------------------------------------------------------
# Nucleation
# ---------------------------------------------------------------------------

def detect_nucleations(void_prev: np.ndarray, void_now: np.ndarray, design_mask: np.ndarray,
                       skin: np.ndarray | None = None, max_events: int = 1000):
    """
    New void clusters of this iteration that are not the growth of an existing
    void (docs/theory.md eq. 11.17-11.18).

    Returns (n_interior, n_surface, n_voxels, events, labels): the counts over
    ALL nucleated clusters, and for at most `max_events` of them, largest
    first, a dict each (kind, n_voxels, centroid, bounding box, label) with
    `labels` the cluster label array the label refers to (None if no voxel
    became void).
    """
    new = np.asarray(void_now, dtype=bool) & ~np.asarray(void_prev, dtype=bool) & np.asarray(design_mask, dtype=bool)
    if not np.any(new):
        return 0, 0, 0, [], None
    labels, n = ndimage.label(new, structure=S6)
    idx = np.arange(1, n + 1)
    # A cluster is growth if any of its voxels has a 26-neighbour that was void.
    grown_from_prev = ndimage.binary_dilation(np.asarray(void_prev, dtype=bool), structure=S26)
    is_growth = ndimage.maximum(grown_from_prev, labels, idx).astype(bool)
    if skin is None:
        skin = skin_mask(new.shape)
    on_skin = ndimage.maximum(skin, labels, idx).astype(bool)
    sizes = np.bincount(labels.ravel(), minlength=n + 1)[1:].astype(np.int64)
    nucleated = ~is_growth
    n_interior = int(np.count_nonzero(nucleated & ~on_skin))
    n_surface = int(np.count_nonzero(nucleated & on_skin))
    n_voxels = int(sizes[nucleated].sum())
    events = []
    if np.any(nucleated):
        chosen = idx[nucleated]
        order = np.argsort(-sizes[nucleated], kind="stable")[:max_events]
        chosen = chosen[order]
        objs = ndimage.find_objects(labels)
        centroids = ndimage.center_of_mass(new, labels, chosen)
        for lab, com in zip(chosen, centroids):
            sl = objs[lab - 1]
            events.append(dict(kind="surface" if on_skin[lab - 1] else "interior",
                               n_voxels=int(sizes[lab - 1]),
                               centroid=tuple(float(c) for c in com),
                               bbox_min=tuple(int(q.start) for q in sl),
                               bbox_max=tuple(int(q.stop - 1) for q in sl),
                               label=int(lab)))
    return n_interior, n_surface, n_voxels, events, labels


class RunDiagnostics:
    """
    Per-iteration topology and nucleation record of one run.

    `update` is called by every optimiser once per iteration with the new
    field. It returns the per-iteration scalars written to variables.csv and
    appends one row per nucleation event to nucleation_events.csv; every
    `history_every` iterations it also writes a float16 snapshot of phi to
    history/phi_<iteration>.npz (for evolution plots and animations, and to
    re-run any diagnostic afterwards).
    """

    EVENTS_HEADER = ("iteration,outer_iteration,inner_iteration,kind,n_voxels,centroid_x,centroid_y,"
                     "centroid_z,bbox_min_x,bbox_min_y,bbox_min_z,bbox_max_x,bbox_max_y,bbox_max_z,"
                     "mean_phi_before,mean_stiffness_term,mean_volume_term,mean_piecewise_term,"
                     "mean_lambda2\n")

    def __init__(self, run_dir: str, design_mask: np.ndarray, history_every: int = 10,
                 max_events_per_iteration: int = 1000):
        self.run_dir = run_dir
        self.design_mask = np.asarray(design_mask, dtype=bool)
        self.skin = skin_mask(self.design_mask.shape)
        self.history_every = max(0, int(history_every))
        self.max_events = int(max_events_per_iteration)
        self.prev_void = None
        self.prev_phi = None
        self.prev_topology = None
        self.total_interior = 0
        self.total_surface = 0
        self.first_interior_iteration = None
        self.last_topology_change_iteration = None
        self.events_path = os.path.join(run_dir, "nucleation_events.csv")
        with open(self.events_path, "w") as f:
            f.write(self.EVENTS_HEADER)
        self.history_dir = os.path.join(run_dir, "history")
        if self.history_every > 0:
            make_bulk_dir(self.history_dir)

    def snapshot(self, iteration: int, phi: np.ndarray, force: bool = False):
        """Writes history/phi_<iteration>.npz (float16) on the history cadence, or when forced."""
        if self.history_every <= 0:
            return
        if not force and iteration % self.history_every != 0:
            return
        path = os.path.join(self.history_dir, f"phi_{iteration:06d}.npz")
        if os.path.exists(path) and not force:
            return
        np.savez_compressed(path, phi=np.asarray(phi, dtype=np.float16), iteration=int(iteration))

    def update(self, iteration: int, outer_iteration: int, inner_iteration: int,
               phi: np.ndarray, solid: np.ndarray, void_for_nucleation: np.ndarray,
               n_components: int | None = None, fields: dict | None = None) -> dict:
        """
        Args:
            phi: the field at the end of this iteration.
            solid: the exported solid of that field (phi >= level set, plus
                preserve, minus obstacles).
            void_for_nucleation: phi < level set (the void set V_k of eq. 11.17).
            n_components: b0 if already known (the connectivity report counts it).
            fields: optional per-voxel arrays of the update just applied
                ("stiffness_term", "volume_term", "piecewise_term", "lambda2");
                their means over each nucleated cluster are written to
                nucleation_events.csv, i.e. the forces that opened the hole.
        """
        topo = topology(solid, n_components=n_components)
        out = dict(euler_characteristic=topo["euler"], n_tunnels=topo["b1"], n_cavities=topo["b2"],
                   interface_faces=interface_faces(solid),
                   nucleations_interior=0, nucleations_surface=0, nucleated_voxels=0)
        key = (topo["b0"], topo["b1"], topo["b2"])
        if self.prev_topology is not None and key != self.prev_topology:
            self.last_topology_change_iteration = int(iteration)
        self.prev_topology = key

        void_now = np.asarray(void_for_nucleation, dtype=bool)
        if self.prev_void is not None:
            n_int, n_surf, n_vox, events, labels = detect_nucleations(
                self.prev_void, void_now, self.design_mask, skin=self.skin, max_events=self.max_events)
            out.update(nucleations_interior=n_int, nucleations_surface=n_surf, nucleated_voxels=n_vox)
            self.total_interior += n_int
            self.total_surface += n_surf
            if n_int > 0 and self.first_interior_iteration is None:
                self.first_interior_iteration = int(iteration)
            if events:
                self._write_events(iteration, outer_iteration, inner_iteration, events, labels, fields)
        self.prev_void = void_now.copy()
        self.prev_phi = np.asarray(phi, dtype=np.float32).copy()
        self.snapshot(iteration, phi)
        return out

    def update_gpu(self, D, iteration: int, outer_iteration: int, inner_iteration: int,
                   phi: np.ndarray, n_components: int | None = None, fields: dict | None = None) -> dict:
        """
        `update` with the topology and the nucleation counts computed on the device (DiagGPU `D`, whose
        scan of `phi` is current). The integers are exactly the host routines'; a nucleation that has to be reported
        (rare) is handed to the unchanged host `detect_nucleations` on the downloaded masks, so the events written
        are the host's. State (previous topology, void, totals) is kept as in `update`; the previous void lives on
        the device and the previous phi is a reference (it is only needed when an event is written).
        """
        topo = D.topology(phi, n_components)
        out = dict(euler_characteristic=topo["euler"], n_tunnels=topo["b1"], n_cavities=topo["b2"],
                   interface_faces=topo["interface_faces"],
                   nucleations_interior=0, nucleations_surface=0, nucleated_voxels=0)
        key = (topo["b0"], topo["b1"], topo["b2"])
        if self.prev_topology is not None and key != self.prev_topology:
            self.last_topology_change_iteration = int(iteration)
        self.prev_topology = key
        if D.have_prev:
            n_int, n_surf, n_vox = D.nucleation_counts(phi)
            events = []
            if n_int + n_surf > 0:
                if getattr(self, "_prev_phi_ref", None) is not None:
                    self.prev_phi = np.asarray(self._prev_phi_ref, dtype=np.float32)
                n_int, n_surf, n_vox, events, labels = detect_nucleations(
                    D.prev_void_host(), D.void_host(), self.design_mask, skin=self.skin, max_events=self.max_events)
            out.update(nucleations_interior=n_int, nucleations_surface=n_surf, nucleated_voxels=n_vox)
            self.total_interior += n_int
            self.total_surface += n_surf
            if n_int > 0 and self.first_interior_iteration is None:
                self.first_interior_iteration = int(iteration)
            if events:
                self._write_events(iteration, outer_iteration, inner_iteration, events, labels, fields)
        D.advance_void()
        self._prev_phi_ref = phi
        self.prev_phi = None
        self.snapshot(iteration, phi)
        return out

    def _write_events(self, iteration, outer_iteration, inner_iteration, events, labels, fields):
        fields = dict(fields or {})
        if self.prev_phi is not None:
            fields["phi_before"] = self.prev_phi
        index = [ev["label"] for ev in events]
        means = {}
        for name in ("phi_before", "stiffness_term", "volume_term", "piecewise_term", "lambda2"):
            arr = fields.get(name)
            if arr is None:
                means[name] = ["N/A"] * len(index)
                continue
            try:
                vals = ndimage.mean(np.asarray(arr, dtype=np.float64), labels, index)
                means[name] = [f"{float(v):.6g}" for v in np.atleast_1d(vals)]
            except Exception:
                means[name] = ["N/A"] * len(index)
        rows = []
        for i, ev in enumerate(events):
            c, lo, hi = ev["centroid"], ev["bbox_min"], ev["bbox_max"]
            rows.append(f"{iteration},{outer_iteration},{inner_iteration},{ev['kind']},{ev['n_voxels']},"
                        f"{c[0]:.2f},{c[1]:.2f},{c[2]:.2f},{lo[0]},{lo[1]},{lo[2]},{hi[0]},{hi[1]},{hi[2]},"
                        f"{means['phi_before'][i]},{means['stiffness_term'][i]},{means['volume_term'][i]},"
                        f"{means['piecewise_term'][i]},{means['lambda2'][i]}\n")
        with open(self.events_path, "a") as f:
            f.writelines(rows)


# ---------------------------------------------------------------------------
# Geometry of the final structure
# ---------------------------------------------------------------------------

def symmetry_mismatch(solid: np.ndarray) -> dict:
    """
    For each axis, the fraction of the solid whose mirror image in the
    mid-plane normal to that axis is void: 0 for a design symmetric about
    that plane. (Only meaningful if the problem itself is symmetric.)
    """
    s = np.asarray(solid, dtype=bool)
    n = max(1, int(np.count_nonzero(s)))
    return {ax: float(np.count_nonzero(s & ~np.flip(s, axis=i))) / n for i, ax in enumerate("xyz")}


def thickness_samples(solid: np.ndarray):
    """
    Local member thickness in voxels, sampled on the ridge of the distance
    transform (a discrete medial axis).

    d = Euclidean distance from each solid voxel centre to the nearest void
    voxel centre (the outside of the domain counts as void). The ridge is the
    set of solid voxels whose d no 26-neighbour exceeds. Across a member t
    voxels thick, d peaks at (t + 1)/2 on a single middle layer when t is odd
    and at t/2 on two middle layers when t is even, so the thickness at a
    ridge voxel is

        t = 2 d - 1 + [its neighbour across the member is also on the ridge],

    "across" being the axis along which the two neighbours' d sum to the
    least. This is exact for grid-aligned plates and bars of any thickness
    (a two-voxel wall reads 2, not 1) and a voxel-resolution estimate for
    inclined members. Returns (samples, half_width_map), half_width_map =
    d - 0.5 being the distance from each solid voxel centre to the nearest
    solid/void interface (0 outside the solid).
    """
    s = np.pad(np.asarray(solid, dtype=bool), 1)
    if not np.any(s):
        return np.zeros(0), np.zeros(np.asarray(solid).shape, dtype=np.float32)
    d = ndimage.distance_transform_edt(s)
    ridge = s & (d >= ndimage.maximum_filter(d, footprint=S26, mode="constant", cval=0.0))
    # The two neighbours along each axis (the padding is void, d = 0 there).
    sums, twin = [], []
    for a in range(3):
        plus = np.roll(d, -1, axis=a)
        minus = np.roll(d, 1, axis=a)
        sums.append(plus + minus)
        twin.append(np.maximum(plus, minus))
    sums = np.stack(sums)
    twin = np.stack(twin)
    across = np.argmin(sums, axis=0)
    twin_across = np.take_along_axis(twin, across[None], axis=0)[0]
    t = 2.0 * d - 1.0 + (twin_across >= d - 1e-9)
    samples = t[ridge]
    half = (np.where(s, d - 0.5, 0.0))[1:-1, 1:-1, 1:-1].astype(np.float32)
    return samples, half


def lorenz_curve(values: np.ndarray):
    """
    Energy concentration of the solid: with the voxels sorted by decreasing
    value, the cumulative share of the total against the cumulative share of
    the voxels. Returns (volume_share, value_share, gini), gini = 0 for a
    uniform distribution (every voxel equally loaded) and -> 1 when a few
    voxels carry everything.
    """
    v = np.sort(np.clip(np.asarray(values, dtype=np.float64).ravel(), 0.0, None))[::-1]
    if v.size == 0 or v.sum() <= 0:
        return np.array([0.0, 1.0]), np.array([0.0, 1.0]), float("nan")
    cum = np.concatenate([[0.0], np.cumsum(v) / v.sum()])
    x = np.linspace(0.0, 1.0, v.size + 1)
    area = np.trapezoid(cum, x) if hasattr(np, "trapezoid") else np.trapz(cum, x)
    gini = float(2.0 * area - 1.0)
    return x, cum, gini


def share_of_volume_carrying(values: np.ndarray, fraction: float) -> float:
    """Smallest share of the voxels that carries `fraction` of the total (e.g. 0.5, 0.9)."""
    x, cum, _ = lorenz_curve(values)
    k = int(np.searchsorted(cum, fraction))
    return float(x[min(k, x.size - 1)])


# ---------------------------------------------------------------------------
# Images
# ---------------------------------------------------------------------------

def depth_render(solid: np.ndarray, axis: int, from_positive: bool = False,
                 light=(-0.5, 0.6, 1.0)):
    """
    Orthographic depth-shaded view of the solid along one axis: for every ray
    the first solid voxel hit, shaded by a Lambert term from the gradient of
    the depth map and dimmed with depth.

    Returns (img, axes, first, hit): img in [0, 1] (NaN where the ray misses)
    indexed [u, v] over the two remaining axes `axes`; `first` the index
    along `axis` (in the original orientation) of the voxel each ray hits;
    `hit` the rays that hit anything. `first` lets any per-voxel field (a
    component label, a stress) be looked up on the visible surface.
    """
    s = np.asarray(solid, dtype=bool)
    n = s.shape[axis]
    sf = np.flip(s, axis=axis) if from_positive else s
    hit = np.any(sf, axis=axis)
    first_f = np.argmax(sf, axis=axis)
    depth = np.where(hit, first_f.astype(np.float64), np.nan)
    axes = [a for a in range(3) if a != axis]
    filled = np.where(hit, first_f.astype(np.float64), float(n))
    if min(filled.shape) >= 2:
        gu, gv = np.gradient(filled)
    else:
        gu = gv = np.zeros_like(filled)
    gu = np.clip(gu, -3, 3)
    gv = np.clip(gv, -3, 3)
    nrm = np.stack([-gu, -gv, np.ones_like(gu)], axis=-1)
    nrm /= np.linalg.norm(nrm, axis=-1, keepdims=True)
    L = np.asarray(light, dtype=np.float64)
    L = L / np.linalg.norm(L)
    lambert = np.clip(nrm @ L, 0.0, 1.0)
    shade = 0.25 + 0.40 * lambert + 0.35 * (1.0 - depth / max(1, n - 1))
    img = np.where(hit, np.clip(shade, 0.0, 1.0), np.nan)
    first = (n - 1 - first_f) if from_positive else first_f
    return img, axes, first, hit


def lookup_on_surface(field: np.ndarray, axis: int, first: np.ndarray) -> np.ndarray:
    """The value of a per-voxel field at the voxel each ray of depth_render hits."""
    f = np.moveaxis(np.asarray(field), axis, -1)
    return np.take_along_axis(f, first[..., None], axis=-1)[..., 0]


# ====================================================================================================
# Greyness and interface-crispness metrics (formerly crispness.py)
# ====================================================================================================
# Discreteness and interface-crispness metrics of a voxel design, for comparing
# PCLS, SIMP and the classical level set on the same footing (docs/theory.md
# §8.4.7). Method-agnostic: every metric is computed from the design field phi
# that all three optimisers write (phi in [V, S] = [1, 2], eq. 12.5), mapped to
# the design fraction
#
#     x = clip((phi - V) / (S - V), 0, 1)      (SIMP's x_phys, the level set's
#                                               clip(1/2 + psi, 0, 1), PCLS's
#                                               rho without the void floor c1)
#
# on the DESIGN voxels (preserved and obstacle voxels are forced to 1 and 0 and
# would only dilute the numbers). The exported geometry is x >= 1/2.
#
# Metrics (all dimensionless unless stated; 0 is "perfectly crisp" for the first
# five):
#
#   non_discreteness       M_nd = 4 <x (1 - x)> over the design voxels
#                          (Sigmund 2007): 0 for a 0/1 design, 1 for x = 1/2
#                          everywhere.                                 (8.20)
#   grey_fraction          share of design voxels with 0.1 < x < 0.9.  (8.21)
#   interface_width_vox    W = (3/2) sum_d 4 x (1 - x) / A_d [voxels]:
#                          the width of the linear 0 -> 1 density ramp that
#                          would hold the same non-discreteness per unit area
#                          of the exported surface (A_d, the marching-cubes
#                          area of the x = 1/2 isosurface next to design
#                          voxels, in voxel^2). 0 for a binary design, ~0.5-0.9
#                          for an ersatz level set (one partially filled voxel
#                          across the boundary), ~0.8-1 for SIMP with a cone
#                          filter of radius 1.5 (optimiser review, 2026-10).
#                                                                      (8.22)
#   bulk_grey_fraction     share of design voxels that are grey AND more than
#                          2 voxels (centre to centre) from the nearest voxel
#                          of the other phase: grey that no boundary
#                          transition explains.                        (8.23)
#   threshold_volume_sensitivity
#                          S_V = (|x >= 0.3| - |x >= 0.7|) / |x >= 0.5| over
#                          the design voxels: the share of the exported
#                          design volume that depends on the choice of cut in
#                          [0.3, 0.7].                                 (8.24)
#   topology_stable        True when the Betti numbers (pieces, tunnels,
#                          cavities) of the cuts x >= 0.3, 0.5 and 0.7 agree
#                          (the exported topology does not depend on the cut).
#   thin_member_fraction   share of the medial-axis samples of the exported
#                          solid thinner than 2 voxels (diagnostics.
#                          thickness_samples): members a voxel-resolution
#                          surface (alpha wrap, marching cubes) reproduces
#                          poorly.                                     (8.25)
#   interface_area_vox2    A_d [voxel^2].
#   n_tunnels, n_cavities  of the exported solid (26/6 topology); cavities
#                          are enclosed voids, which many processes cannot make.
#
# Usage:
#     from baseline_optimiser import crispness_metrics, design_fraction
#     m = crispness_metrics(design_fraction(phi), design_mask)
#
#     python -c "import sys, baseline_optimiser as b; print(b.format_table({d: b.run_metrics(d) for d in sys.argv[1:]}))" RUN [RUN ...]
#                                                           # the table of finished runs (final_fields.npz)

GREY_LO, GREY_HI = 0.1, 0.9           # the grey band of eq. (8.21), as PCLS's grey test (eq. 5.7)
BULK_DISTANCE = 2.0                   # voxels: grey farther than this from the other phase is "bulk" (8.23)
CUTS = (0.3, 0.5, 0.7)                # cuts of the threshold sensitivity (8.24)
THIN_VOXELS = 2.0                     # members thinner than this count as thin (8.25)

METRIC_KEYS = ("non_discreteness", "grey_fraction", "interface_width_vox", "bulk_grey_fraction",
               "threshold_volume_sensitivity", "topology_stable", "thin_member_fraction",
               "median_member_thickness_vox", "interface_area_vox2", "n_components", "n_tunnels", "n_cavities")


def design_fraction(phi: np.ndarray, solid_value: float = 2.0, void_value: float = 1.0) -> np.ndarray:
    """x = clip((phi - V)/(S - V), 0, 1): the design fraction of every voxel (eq. 8.19)."""
    phi = np.asarray(phi, dtype=np.float64)
    return np.clip((phi - void_value) / (solid_value - void_value), 0.0, 1.0)


def _isosurface_area_near(x: np.ndarray, design: np.ndarray, level: float = 0.5) -> float:
    """
    Area [voxel^2] of the triangles of the marching-cubes isosurface x = level
    (the outside of the domain counts as x = 0, so the surface is closed) that
    lie in a lattice cell with at least one design voxel among its 8 corners.
    """
    try:
        from skimage.measure import marching_cubes
    except Exception:                                  # pragma: no cover
        return float("nan")
    xp = np.pad(np.asarray(x, dtype=np.float64), 1, mode="constant", constant_values=0.0)
    if not (np.any(xp >= level) and np.any(xp < level)):
        return 0.0
    verts, faces, _, _ = marching_cubes(xp, level=level, allow_degenerate=False)
    if len(faces) == 0:
        return 0.0
    tri = verts[faces]                                                  # (F, 3, 3), padded index units
    area = 0.5 * np.linalg.norm(np.cross(tri[:, 1] - tri[:, 0], tri[:, 2] - tri[:, 0]), axis=1)
    cell = np.floor(tri.mean(axis=1)).astype(np.int64)                  # lower corner of the lattice cell
    dp = np.pad(np.asarray(design, dtype=bool), 1, mode="constant", constant_values=False)
    hi = np.array(dp.shape) - 1
    near = np.zeros(len(faces), dtype=bool)
    for o in np.ndindex(2, 2, 2):
        c = np.minimum(cell + np.array(o), hi)
        near |= dp[c[:, 0], c[:, 1], c[:, 2]]
    return float(area[near].sum())


def _distance_to_other_phase(solid: np.ndarray) -> np.ndarray:
    """Centre-to-centre distance [voxels] from each voxel to the nearest voxel of the other phase (outside = void)."""
    s = np.pad(np.asarray(solid, dtype=bool), 1, mode="constant", constant_values=False)
    d_in = ndimage.distance_transform_edt(s)
    d_out = ndimage.distance_transform_edt(~s)
    return np.where(s, d_in, d_out)[1:-1, 1:-1, 1:-1]


def crispness_metrics(x: np.ndarray, design_mask: np.ndarray | None = None, light: bool = False) -> dict:
    """
    The metrics of the module docstring for a design fraction field x (eq.
    8.19) and the design-voxel mask (None = every voxel). Returns a dict of
    plain Python numbers (NaN where undefined, e.g. no interface). light=True
    leaves out the topology of the three cuts and the member thickness (the
    two costly parts), for evolution plots over many snapshots.
    """
    x = np.clip(np.asarray(x, dtype=np.float64), 0.0, 1.0)
    design = np.ones(x.shape, dtype=bool) if design_mask is None else np.asarray(design_mask, dtype=bool)
    nd = int(np.count_nonzero(design))
    out = {k: float("nan") for k in METRIC_KEYS}
    if nd == 0:
        return out
    xd = x[design]
    q = 4.0 * x * (1.0 - x)
    out["non_discreteness"] = float(np.mean(q[design]))
    grey = (x > GREY_LO) & (x < GREY_HI)
    out["grey_fraction"] = float(np.count_nonzero(grey & design)) / nd
    area = _isosurface_area_near(x, design)
    out["interface_area_vox2"] = area
    out["interface_width_vox"] = float(1.5 * np.sum(q[design]) / area) if area > 0 else float("nan")
    solid = x >= 0.5
    dist = _distance_to_other_phase(solid)
    out["bulk_grey_fraction"] = float(np.count_nonzero(grey & design & (dist > BULK_DISTANCE))) / nd
    vols = {c: int(np.count_nonzero(xd >= c)) for c in CUTS}
    out["threshold_volume_sensitivity"] = (float(vols[CUTS[0]] - vols[CUTS[-1]]) / vols[0.5]
                                           if vols[0.5] > 0 else float("nan"))
    if light:
        return out
    betti = {}
    for c in CUTS:
        t = topology(x >= c)
        betti[c] = (int(t["b0"]), int(t["b1"]), int(t["b2"]))
    out["topology_stable"] = bool(len(set(betti.values())) == 1)
    out["topology_by_cut"] = {f"{c:g}": {"n_components": b[0], "n_tunnels": b[1], "n_cavities": b[2]}
                              for c, b in betti.items()}
    out["n_components"], out["n_tunnels"], out["n_cavities"] = betti[0.5]
    samples, _ = thickness_samples(solid)
    if samples.size:
        out["thin_member_fraction"] = float(np.mean(samples < THIN_VOXELS))
        out["median_member_thickness_vox"] = float(np.median(samples))
    return out


def export_fidelity(model_compliance=None, exported_compliance=None,
                    model_load_factor=None, exported_load_factor=None) -> dict:
    """Exported-over-model ratios of compliance and critical load factor (1 = the export is what was optimised)."""
    def ratio(a, b):
        try:
            a, b = float(a), float(b)
        except (TypeError, ValueError):
            return None
        if not (np.isfinite(a) and np.isfinite(b)) or b == 0.0:
            return None
        return a / b
    return {"compliance_exported_over_model": ratio(exported_compliance, model_compliance),
            "load_factor_exported_over_model": ratio(exported_load_factor, model_load_factor)}


def run_metrics(run_dir: str) -> dict:
    """Metrics of a finished run directory from final_fields.npz (+ run_info.json for S, V and the ratios)."""
    info = {}
    p_info = os.path.join(run_dir, "run_info.json")
    if os.path.exists(p_info):
        with open(p_info) as f:
            info = json.load(f)
    params = info.get("optimiser_parameters", {}) or {}
    S = float(params.get("solid_value", 2.0) or 2.0)
    V = float(params.get("void_value", 1.0) or 1.0)
    with np.load(os.path.join(run_dir, "final_fields.npz")) as z:
        phi = z["phi"]
        design = z["design_mask"] if "design_mask" in z.files else None
    m = crispness_metrics(design_fraction(phi, S, V), design)
    final = info.get("final", {}) or {}
    buck = info.get("buckling", {}) or {}
    m.update(export_fidelity(final.get("compliance"), final.get("exported_compliance"),
                             buck.get("last_valid_critical_load_factor"), buck.get("critical_load_factor_exported")))
    return m


def history_metrics(run_dir: str, max_frames: int | None = None, light: bool = True) -> list:
    """
    [(iteration, metrics)] for the snapshots in <run_dir>/history (phi_<iteration>.npz), oldest
    first; at most max_frames of them, evenly spaced (None = all); light metrics by default.
    """
    hdir = os.path.join(run_dir, "history")
    out = []
    if not os.path.isdir(hdir):
        return out
    S, V = 2.0, 1.0
    p_info = os.path.join(run_dir, "run_info.json")
    if os.path.exists(p_info):
        with open(p_info) as f:
            params = (json.load(f).get("optimiser_parameters") or {})
        S = float(params.get("solid_value", 2.0) or 2.0)
        V = float(params.get("void_value", 1.0) or 1.0)
    design = None
    p_ff = os.path.join(run_dir, "final_fields.npz")
    if os.path.exists(p_ff):
        with np.load(p_ff) as z:
            design = z["design_mask"] if "design_mask" in z.files else None
    names = []
    for name in sorted(os.listdir(hdir)):
        if not (name.startswith("phi_") and name.endswith(".npz")):
            continue
        try:
            names.append((int(name[4:-4]), name))
        except ValueError:
            continue
    names.sort()
    if max_frames and len(names) > max_frames:
        keep = np.unique(np.linspace(0, len(names) - 1, max_frames).round().astype(int))
        names = [names[i] for i in keep]
    for it, name in names:
        with np.load(os.path.join(hdir, name)) as z:
            phi = z[z.files[0]] if "phi" not in z.files else z["phi"]
        out.append((it, crispness_metrics(design_fraction(phi, S, V), design, light=light)))
    return out


def format_table(rows: dict) -> str:
    """A plain-text table {label: metrics}."""
    keys = ("non_discreteness", "grey_fraction", "interface_width_vox", "bulk_grey_fraction",
            "threshold_volume_sensitivity", "topology_stable", "thin_member_fraction", "n_tunnels", "n_cavities",
            "compliance_exported_over_model", "load_factor_exported_over_model")
    head = ["run"] + list(keys)
    lines = ["\t".join(head)]
    for label, m in rows.items():
        cells = [label]
        for k in keys:
            v = m.get(k)
            cells.append("-" if v is None else (str(v) if isinstance(v, (bool, int, np.integer)) else f"{v:.4g}"))
        lines.append("\t".join(cells))
    return "\n".join(lines)


# ====================================================================================================
# Closed-surface (STL) export of a final design (formerly surface_export.py)
# ====================================================================================================
# Closed-surface (STL) export of a final design, docs/theory.md §8.4 and §10.6.
#
# The design field is contoured with marching cubes and, optionally, wrapped with
# CGAL's 3D alpha wrap. This module exists because the wrap's resolution used to
# be set as a fraction of the bounding-box diagonal (alpha_proportion = 0.1 of a
# 170 mm diagonal is a 17 mm ball, 17 voxels, far larger than the gaps of a
# design with struts 2-4 voxels thick, so the wrap filled the design in; see
# claude/alpha_wrap_resolution_plan.md). Here the wrap is specified in voxels,
#
#     alpha  = alpha_voxels * h            (h = voxel edge, m)
#     offset = offset_ratio * alpha,
#
# and converted once to the fractions of the diagonal d of the mesh being wrapped
# that the compiled binding takes: alpha_proportion = alpha / d.
#
# Other changes against the old implementation:
#   * small connected components are removed on the voxel grid (milliseconds)
#     before marching cubes, instead of through the binding's cutoff_ratio, whose
#     direction was never verified. The main body is wrapped once. The "raw" file
#     gets the removed specks back as plain marching-cubes shells instead of a
#     second wrap.
#   * method = "marching_cubes" skips CGAL: the marching-cubes mesh of a padded
#     field is already closed and is exact to the field.
#   * the numbers that make a bad setting visible (volume of the surface against
#     the volume of the voxels, faces, wall time, alpha and offset in voxels and
#     mm) are returned and logged.
#
# Only NumPy and SciPy are needed at import; scikit-image, CGAL and trimesh are
# imported inside the functions that use them. run_surface_job() is the entry point of a separate
# interpreter that the optimiser starts, so that a wrap can run in another process without multiprocessing, which would re-execute
# optimise.py on Windows (it has no __main__ guard).


# Defaults (docs/theory.md §10.6). One voxel is d/170 on the 160x40x40 test domain,
# the fine end of the range CGAL's own benchmark covers; offset = alpha/30 is the
# ratio of CGAL's examples (a small fraction of alpha; too small or too large an
# offset raises the output complexity).
DEFAULT_METHOD = "alpha_wrap"
DEFAULT_ALPHA_VOXELS = 1.5
DEFAULT_OFFSET_RATIO = 1.0 / 30.0
DEFAULT_PERIODIC_ALPHA_VOXELS = 1.5
DEFAULT_MIN_COMPONENT_FRACTION = 0.01
VALID_METHODS = ("alpha_wrap", "marching_cubes")

# Warnings raised by check_surface().
VOLUME_RATIO_WARN = 0.10     # |V_surface / V_voxels - 1| above this is reported
ALPHA_VOXELS_WARN = 4.0      # the closing proxy of the plan: features are lost beyond ~6-8 voxels


# --------------------------------------------------------------------------------------
# Geometry helpers
# --------------------------------------------------------------------------------------

def bbox_diagonal(verts: np.ndarray) -> float:
    verts = np.asarray(verts, dtype=float)
    return float(np.linalg.norm(verts.max(axis=0) - verts.min(axis=0)))


def marching_cubes_mesh(field: np.ndarray, level: float, pad_value: float, voxel_size_m: float):
    """
    Marching cubes of `field` at `level` after padding one voxel of `pad_value` on every side, so
    the isosurface closes at the domain boundary. Vertices are returned in metres, with the
    voxel-centre convention of Conformal._surface_mesh (a solid voxel at index i is centred at
    i * h; the padded surface lies half a voxel outside the outermost solid voxels).
    """
    from skimage.measure import marching_cubes
    padded = np.pad(np.asarray(field), pad_width=1, mode="constant", constant_values=pad_value)
    verts, faces, _normals, _ = marching_cubes(volume=padded, level=level, step_size=1,
                                               allow_degenerate=False)
    verts = np.asarray(verts, dtype=np.float64) - 1.0
    verts *= voxel_size_m
    faces = np.asarray(faces, dtype=np.int64)
    # Outward normals (enclosed volume > 0), whichever way skimage winds a field whose solid is
    # the high side; CGAL's wrap is outward-oriented, so the two kinds of surface agree.
    t = verts[faces]
    if np.einsum("ij,ij->i", t[:, 0], np.cross(t[:, 1], t[:, 2])).sum() < 0.0:
        faces = faces[:, ::-1].copy()
    return verts, faces


def split_components(field: np.ndarray, level: float, pad_value: float,
                     min_fraction: float = DEFAULT_MIN_COMPONENT_FRACTION) -> dict:
    """
    Removes small solid components (solid = field >= level, 26-connected like topology()) on
    the voxel grid. A component is removed when it holds fewer than min_fraction of the solid
    voxels; the largest component is always kept. Returns

        filtered_field   field with the removed components set to pad_value
        specks_field     field holding only the removed components (None if nothing removed)
        n_components, n_removed, solid_voxels, kept_voxels, removed_voxels, min_voxels
    """
    field = np.asarray(field)
    solid = field >= level
    labels, n = ndimage.label(solid, structure=S26)
    solid_voxels = int(solid.sum())
    out = {"n_components": int(n), "n_removed": 0, "solid_voxels": solid_voxels,
           "kept_voxels": solid_voxels, "removed_voxels": 0, "min_voxels": 0,
           "filtered_field": field, "specks_field": None}
    if n <= 1 or min_fraction <= 0.0:
        return out
    sizes = np.bincount(labels.ravel())[1:]
    min_voxels = max(1, int(np.ceil(min_fraction * solid_voxels)))
    drop = sizes < min_voxels
    drop[int(np.argmax(sizes))] = False
    if not drop.any():
        out["min_voxels"] = min_voxels
        return out
    removed = np.concatenate([[False], drop])[labels]  # labels 0 (void) -> False
    out.update(n_removed=int(drop.sum()), kept_voxels=int(solid_voxels - sizes[drop].sum()),
               removed_voxels=int(sizes[drop].sum()), min_voxels=min_voxels,
               filtered_field=np.where(removed, pad_value, field).astype(field.dtype, copy=False),
               specks_field=np.where(removed, field, pad_value).astype(field.dtype, copy=False))
    return out


def enclosed_cavities(field: np.ndarray, level: float) -> tuple[np.ndarray, int]:
    """
    Void that does not reach the outside of the domain: (mask, number of cavities). Void is
    6-connected, the counterpart of the 26-connected solid used everywhere else, and the outside is the
    void connected to the padding around the domain, so a void touching a face of the domain is not a
    cavity (the isosurface closes there, like the marching-cubes padding).
    """
    void = np.pad(np.asarray(field) < level, 1, constant_values=True)
    lab, k = ndimage.label(void)
    outside = lab[0, 0, 0]
    cav = (lab != 0) & (lab != outside)
    n = len(np.unique(lab[cav]))
    return cav[1:-1, 1:-1, 1:-1], int(n)


def cavity_surface(field: np.ndarray, level: float, cavity_mask: np.ndarray, voxel_size_m: float):
    """
    Surface of the enclosed cavities, oriented INTO the cavity (out of the material), so that its enclosed
    volume is negative and adds to that of the outer surface correctly. It is contoured from the design
    field itself, reflected about the level so that the cavity is the "solid" side, over the cavity and
    the layer of solid voxels next to it (everything else is set below the level).
    """
    solid_near = ndimage.binary_dilation(cavity_mask, structure=ndimage.generate_binary_structure(3, 1)) \
        & (np.asarray(field) >= level)
    lo = level - 1.0
    g = np.where(cavity_mask | solid_near, 2.0 * level - np.asarray(field), lo)
    verts, faces = marching_cubes_mesh(g, level, lo, voxel_size_m)   # outward of the cavity
    return verts, faces[:, ::-1].copy()                               # into the cavity


def resolve_wrap_parameters(verts: np.ndarray, voxel_size_m: float,
                            alpha_voxels: float = DEFAULT_ALPHA_VOXELS,
                            offset_ratio: float = DEFAULT_OFFSET_RATIO,
                            alpha_proportion: float | None = None,
                            offset_proportion: float | None = None) -> dict:
    """
    Converts the voxel-based wrap settings to the binding's fractions of the diagonal d of `verts`
    (the mesh that will be wrapped). Explicit alpha_proportion / offset_proportion (the old
    parameters) take precedence over the voxel-based ones when given. Returns the fractions and
    the resulting alpha and offset in voxels and in mm.
    """
    d = bbox_diagonal(verts)
    if not d > 0.0:
        raise ValueError("The surface to wrap has a zero-size bounding box.")
    if alpha_proportion is None:
        if not alpha_voxels > 0.0:
            raise ValueError("alpha_voxels must be positive.")
        alpha_proportion = alpha_voxels * voxel_size_m / d
    if offset_proportion is None:
        if not offset_ratio > 0.0:
            raise ValueError("offset_ratio must be positive.")
        offset_proportion = offset_ratio * alpha_proportion
    alpha = alpha_proportion * d
    offset = offset_proportion * d
    return {"alpha_proportion": float(alpha_proportion), "offset_proportion": float(offset_proportion),
            "bbox_diagonal_m": d, "alpha_m": float(alpha), "offset_m": float(offset),
            "alpha_voxels": float(alpha / voxel_size_m), "offset_voxels": float(offset / voxel_size_m),
            "alpha_mm": float(alpha * 1e3), "offset_mm": float(offset * 1e3)}


# --------------------------------------------------------------------------------------
# The wrap itself
# --------------------------------------------------------------------------------------

def cgal_wrap(vertices, faces, alpha_proportion, offset_proportion, printout=False):
    """
    alpha_wrap.wrap(...) with no component cutoff: filtering has already been done on the voxel
    grid, and a call without cutoff_ratio keeps everything (the quiet branch of
    Conformal._surface_mesh relied on exactly that before cutoff_ratio was added). If this build
    of the binding requires cutoff_ratio, a huge one is passed instead (and the caller is told).
    Returns (dict with "vertices" and "faces", log text, used_cutoff_fallback).
    """
    from CGAL import alpha_wrap
    log = io.StringIO()
    fallback = False
    with redirect_stdout(log):
        try:
            wrapped = alpha_wrap.wrap(vertices=vertices, faces=faces,
                                      alpha_proportion=alpha_proportion,
                                      offset_proportion=offset_proportion, printout=printout)
        except TypeError:
            fallback = True
            wrapped = alpha_wrap.wrap(vertices=vertices, faces=faces,
                                      alpha_proportion=alpha_proportion,
                                      offset_proportion=offset_proportion,
                                      cutoff_ratio=1.0e18, printout=printout)
    lines = [ln for ln in log.getvalue().splitlines() if not ln.strip().startswith("Pre-discarding")]
    return wrapped, "\n".join(lines).strip(), fallback


def mesh_stats(verts: np.ndarray, faces: np.ndarray) -> dict:
    """Faces, area, enclosed volume (divergence theorem; > 0 for outward normals), whether every
    edge is shared by exactly two faces, the number of shells and the Euler characteristic."""
    verts = np.asarray(verts, dtype=np.float64)
    faces = np.asarray(faces, dtype=np.int64)
    if len(faces) == 0:
        return {"vertices": int(len(verts)), "faces": 0, "area_m2": 0.0, "volume_m3": 0.0,
                "closed_manifold_edges": False, "shells": 0, "euler": 0}
    t = verts[faces]
    cross = np.cross(t[:, 1] - t[:, 0], t[:, 2] - t[:, 0])
    area = 0.5 * np.linalg.norm(cross, axis=1).sum()
    volume = float(np.einsum("ij,ij->i", t[:, 0], np.cross(t[:, 1], t[:, 2])).sum() / 6.0)
    edges = np.sort(np.concatenate([faces[:, [0, 1]], faces[:, [1, 2]], faces[:, [2, 0]]]), axis=1)
    uniq, counts = np.unique(edges, axis=0, return_counts=True)
    nv = len(verts)
    used = np.unique(faces)
    adj = coo_matrix((np.ones(len(uniq)), (uniq[:, 0], uniq[:, 1])), shape=(nv, nv))
    ncomp, lab = connected_components(adj, directed=False)
    shells = int(len(np.unique(lab[used])))
    return {"vertices": int(len(used)), "faces": int(len(faces)), "area_m2": float(area),
            "volume_m3": volume, "closed_manifold_edges": bool(np.all(counts == 2)),
            "shells": shells, "euler": int(len(used) - len(uniq) + len(faces))}


def voxel_topology(solid: np.ndarray) -> dict | None:
    """
    Topology of the union of closed unit cubes at the solid voxels (topology(): 26-connected
    solid, so voxels that touch only along an edge or at a corner are joined), and the shells and Euler
    characteristic a closed, manifold surface bounding that solid should have: chi_surface = 2 chi_solid
    and one shell per component and per cavity. Marching cubes does NOT reproduce this exactly (it
    separates some edge contacts, which removes tunnels), so the voxel topology, not the marching-cubes
    mesh, is the reference for the wrap. None if the topology cannot be computed.
    """
    try:
        t = topology(np.asarray(solid, dtype=bool))
    except Exception:
        return None
    return {"euler_solid": int(t["euler"]), "b0": int(t["b0"]), "b1": int(t["b1"]), "b2": int(t["b2"]),
            "surface_euler_expected": int(2 * t["euler"]), "shells_expected": int(t["b0"] + t["b2"])}


def _surface_samples(verts, faces):
    t = np.asarray(verts)[np.asarray(faces)]
    w = np.array([[1/3, 1/3, 1/3], [2/3, 1/6, 1/6], [1/6, 2/3, 1/6], [1/6, 1/6, 2/3]])
    pts = np.einsum("kj,fjc->fkc", w, t).reshape(-1, 3)
    return np.vstack([np.asarray(verts), pts])


def surface_distance(verts_a, faces_a, verts_b, faces_b) -> dict:
    """
    Distances between two surfaces, from point samples of both (four per face plus the vertices,
    so the sampling error is a fraction of the triangle size). Returns the mean, 95th percentile and
    maximum of the distance from A to B and from B to A, and the symmetric Hausdorff distance (the
    larger maximum), in the units of the vertices.
    """
    from scipy.spatial import cKDTree
    sa, sb = _surface_samples(verts_a, faces_a), _surface_samples(verts_b, faces_b)
    d_ab = cKDTree(sb).query(sa)[0]
    d_ba = cKDTree(sa).query(sb)[0]
    return {"a_to_b_mean": float(d_ab.mean()), "a_to_b_p95": float(np.percentile(d_ab, 95)),
            "a_to_b_max": float(d_ab.max()),
            "b_to_a_mean": float(d_ba.mean()), "b_to_a_p95": float(np.percentile(d_ba, 95)),
            "b_to_a_max": float(d_ba.max()),
            "hausdorff": float(max(d_ab.max(), d_ba.max())),
            "p95": float(np.percentile(np.concatenate([d_ab, d_ba]), 95)),
            "mean": float(0.5 * (d_ab.mean() + d_ba.mean()))}


def write_stl(path: str, verts: np.ndarray, faces: np.ndarray) -> None:
    """Binary STL (float32), the format trimesh's export wrote before."""
    verts = np.asarray(verts, dtype=np.float64)
    faces = np.asarray(faces, dtype=np.int64)
    t = verts[faces]
    n = np.cross(t[:, 1] - t[:, 0], t[:, 2] - t[:, 0])
    ln = np.linalg.norm(n, axis=1, keepdims=True)
    n = np.divide(n, ln, out=np.zeros_like(n), where=ln > 0)
    rec = np.zeros(len(faces), dtype=np.dtype([("n", "<f4", 3), ("v", "<f4", (3, 3)), ("a", "<u2")]))
    rec["n"] = n
    rec["v"] = t
    with open(path, "wb") as f:
        f.write(b"PCLS surface export".ljust(80, b" "))
        f.write(struct.pack("<I", len(faces)))
        f.write(rec.tobytes())


def _merge(verts_a, faces_a, verts_b, faces_b):
    return (np.vstack([verts_a, verts_b]), np.vstack([faces_a, faces_b + len(verts_a)]))


# --------------------------------------------------------------------------------------
# The pipeline
# --------------------------------------------------------------------------------------

def build_surfaces(field: np.ndarray, level: float, pad_value: float, voxel_size_m: float,
                   method: str = DEFAULT_METHOD,
                   alpha_voxels: float = DEFAULT_ALPHA_VOXELS,
                   offset_ratio: float = DEFAULT_OFFSET_RATIO,
                   alpha_proportion: float | None = None,
                   offset_proportion: float | None = None,
                   min_component_fraction: float = DEFAULT_MIN_COMPONENT_FRACTION,
                   smoothing_iterations: int = 0,
                   printout: bool = False,
                   measure_fidelity: bool = False,
                   measure_distance: bool = True,
                   wrap_fn=None) -> dict:
    """
    Builds the "filtered" and "raw" surfaces of a design field (solid = field >= level).

    filtered: the main body only (components below min_component_fraction of the solid voxels
              removed on the voxel grid), wrapped once (method "alpha_wrap") or marching cubes.
    raw:      the filtered surface plus the removed specks as plain marching-cubes shells (never
              wrapped). Identical to filtered when nothing was removed.

    `wrap_fn(vertices, faces, alpha_proportion, offset_proportion, printout)` replaces the CGAL
    call (tests); like cgal_wrap it returns (dict with "vertices" and "faces", log, fallback).
    Returns {"raw": (verts, faces), "filtered": (verts, faces), "info": {...}}.
    """
    if method not in VALID_METHODS:
        raise ValueError(f"final surface method must be one of {VALID_METHODS}, got {method!r}")
    if method == "alpha_wrap" and wrap_fn is None:
        try:
            from CGAL import alpha_wrap as _aw_probe  # noqa: F401
        except Exception:
            print("CGAL alpha_wrap unavailable on this platform: final surface falls back to method='marching_cubes'")
            method = "marching_cubes"
    t0 = time.perf_counter()
    comp = split_components(field, level, pad_value, min_component_fraction)
    cavity_mask, n_cavities = enclosed_cavities(comp["filtered_field"], level)
    cavity_voxels = int(cavity_mask.sum())
    if n_cavities and method == "alpha_wrap":
        # The wrap carves from the outside and cannot represent a sealed cavity: given one, CGAL fills it
        # or returns an inner shell with the wrong orientation (the enclosed volume then counts twice).
        # So the wrap sees the body with its cavities filled, and the cavities are added back below as
        # plain marching-cubes shells oriented into the void. Those lie inside the material and the
        # wrap outside it, so the two can never intersect.
        fill = max(float(np.max(comp["filtered_field"])), level + 1.0)
        wrap_field = np.where(cavity_mask, fill, comp["filtered_field"]).astype(comp["filtered_field"].dtype)
    else:
        wrap_field = comp["filtered_field"]
    mc_verts, mc_faces = marching_cubes_mesh(wrap_field, level, pad_value, voxel_size_m)
    t_mc = time.perf_counter() - t0

    info = {"method": method, "voxel_size_m": float(voxel_size_m),
            "components": {k: comp[k] for k in ("n_components", "n_removed", "solid_voxels",
                                                "kept_voxels", "removed_voxels", "min_voxels")},
            "min_component_fraction": float(min_component_fraction),
            "cavities": {"count": n_cavities, "voxels": cavity_voxels},
            "marching_cubes_faces": int(len(mc_faces)), "log": ""}
    if method == "alpha_wrap":
        params = resolve_wrap_parameters(mc_verts, voxel_size_m, alpha_voxels, offset_ratio,
                                         alpha_proportion, offset_proportion)
        info.update(params)
        t1 = time.perf_counter()
        wrapped, log, fallback = (wrap_fn or cgal_wrap)(mc_verts, mc_faces,
                                                        params["alpha_proportion"],
                                                        params["offset_proportion"], printout)
        info["wrap_time_s"] = time.perf_counter() - t1
        info["log"] = log
        info["cutoff_ratio_fallback"] = bool(fallback)
        verts = np.asarray(wrapped["vertices"], dtype=np.float64)
        faces = np.asarray(wrapped["faces"], dtype=np.int64)
    else:
        verts, faces = mc_verts, mc_faces
        info["wrap_time_s"] = 0.0
        if smoothing_iterations and smoothing_iterations > 0:
            import trimesh
            from trimesh import smoothing
            m = trimesh.Trimesh(vertices=verts, faces=faces, process=False)
            smoothing.filter_taubin(m, iterations=int(smoothing_iterations))
            verts, faces = np.asarray(m.vertices), np.asarray(m.faces)
            info["smoothing_iterations"] = int(smoothing_iterations)
    info["marching_cubes_time_s"] = t_mc
    if n_cavities and method == "alpha_wrap":
        cv, cf = cavity_surface(comp["filtered_field"], level, cavity_mask, voxel_size_m)
        verts, faces = _merge(verts, faces, cv, cf)

    filtered = (verts, faces)
    if comp["specks_field"] is not None:
        sv, sf = marching_cubes_mesh(comp["specks_field"], level, pad_value, voxel_size_m)
        raw = _merge(verts, faces, sv, sf)
    else:
        raw = filtered

    h3 = voxel_size_m ** 3
    info["filtered"] = mesh_stats(*filtered)
    info["raw"] = info["filtered"] if raw is filtered else mesh_stats(*raw)
    info["voxel_volume_m3_filtered"] = comp["kept_voxels"] * h3
    info["voxel_volume_m3_all"] = comp["solid_voxels"] * h3
    vv = info["voxel_volume_m3_filtered"]
    info["volume_ratio_filtered"] = info["filtered"]["volume_m3"] / vv if vv > 0 else float("nan")
    if measure_fidelity or measure_distance:
        # Reference: the marching-cubes surface of the same filtered field (the voxel surface).
        if wrap_field is not comp["filtered_field"]:
            ref_verts, ref_faces = marching_cubes_mesh(comp["filtered_field"], level, pad_value, voxel_size_m)
        else:
            ref_verts, ref_faces = mc_verts, mc_faces
        if verts is ref_verts and faces is ref_faces:
            dist = {k: 0.0 for k in ("a_to_b_mean", "a_to_b_p95", "a_to_b_max", "b_to_a_mean", "b_to_a_p95",
                                     "b_to_a_max", "hausdorff", "p95", "mean")}
        else:
            dist = surface_distance(verts, faces, ref_verts, ref_faces)
        info["distance_to_voxel_surface_m"] = dist
        if measure_fidelity:
            info["fidelity_to_marching_cubes_m"] = dist
            info["marching_cubes_reference"] = mesh_stats(ref_verts, ref_faces)
    info["voxel_topology"] = voxel_topology(comp["filtered_field"] >= level)
    info["quality"] = surface_quality(info)
    info["total_time_s"] = time.perf_counter() - t0
    return {"raw": raw, "filtered": filtered, "info": info}


def surface_quality(info: dict) -> dict:
    """
    The quality of the exported (filtered) surface from a build_surfaces() info dict, in one place
    (docs/theory.md §8.4.8): volume_ratio (enclosed volume / volume of the solid voxels, 1 = exact),
    closed_manifold (every edge in exactly two faces), shells and euler against the voxels' expected
    values (shells_expected = b0 + b2, euler_expected = 2 chi_solid) with topology_matches, and the
    distance to the voxel surface (the marching-cubes surface of the same field): hausdorff_voxels,
    p95_voxels, mean_voxels (None for a run exported before it was measured).
    """
    f = info.get("filtered") or {}
    vt = info.get("voxel_topology") or {}
    h = info.get("voxel_size_m")
    q = {"volume_ratio": info.get("volume_ratio_filtered"),
         "closed_manifold": f.get("closed_manifold_edges"),
         "shells": f.get("shells"), "shells_expected": vt.get("shells_expected"),
         "euler": f.get("euler"), "euler_expected": vt.get("surface_euler_expected")}
    q["topology_matches"] = (None if not vt or f.get("shells") is None else
                             bool(f["shells"] == vt["shells_expected"] and f["euler"] == vt["surface_euler_expected"]))
    d = info.get("distance_to_voxel_surface_m") or info.get("fidelity_to_marching_cubes_m")
    for key, name in (("hausdorff", "hausdorff_voxels"), ("p95", "p95_voxels"), ("mean", "mean_voxels")):
        if d and h and key in d:
            q[name] = float(d[key]) / float(h)
        elif d and h and key == "hausdorff":
            q[name] = float(max(d["a_to_b_max"], d["b_to_a_max"])) / float(h)
        else:
            q[name] = None
    return q


def check_surface(info: dict) -> list[str]:
    """Plain-language warnings for a build_surfaces() info dict (empty if all is well)."""
    out = []
    r = info.get("volume_ratio_filtered")
    if r is not None and np.isfinite(r) and abs(r - 1.0) > VOLUME_RATIO_WARN:
        out.append(f"the exported surface encloses {r:.2f} times the volume of the solid voxels "
                   f"(expected 1 within {VOLUME_RATIO_WARN:.0%})"
                   + ("; the alpha ball is too large for the gaps in this design" if r > 1 else ""))
    a = info.get("alpha_voxels")
    if info.get("method") == "alpha_wrap" and a is not None and a > ALPHA_VOXELS_WARN:
        out.append(f"alpha is {a:.1f} voxels; features narrower than a few voxels will be filled "
                   f"(recommended <= {ALPHA_VOXELS_WARN:g})")
    vt = info.get("voxel_topology")
    if info.get("method") == "alpha_wrap" and vt is not None:
        f = info["filtered"]
        if f["euler"] != vt["surface_euler_expected"] or f["shells"] != vt["shells_expected"]:
            out.append(f"the wrapped surface has {f['shells']} shell(s) and Euler characteristic {f['euler']}; "
                       f"the solid voxels give {vt['shells_expected']} and {vt['surface_euler_expected']} "
                       f"({vt['b1']} tunnels): the wrap has closed or created tunnels")
    if not info.get("filtered", {}).get("closed_manifold_edges", True):
        out.append("the filtered surface is not closed and edge-manifold")
    if info.get("cutoff_ratio_fallback"):
        out.append("this build of alpha_wrap requires cutoff_ratio; a huge value was passed")
    return out


def make_bulk_dir(path: str) -> None:
    """os.makedirs(path), except that a path under $HOME is created under $PCLS_BULK_ROOT (same path
    relative to $HOME) and symlinked back, keeping bulk outputs (data/, history/) off the 100 GiB $HOME.
    PCLS_BULK_ROOT defaults to /projects/b6bl/owenpeckham.b6bl/home_data; set it empty to disable."""
    root = os.environ.get("PCLS_BULK_ROOT", "/projects/b6bl/owenpeckham.b6bl/home_data")
    path = os.path.abspath(path)
    home = os.path.realpath(os.path.expanduser("~"))
    rel = os.path.relpath(os.path.realpath(os.path.dirname(path)), home)
    if os.path.lexists(path) or not root or not os.path.isdir(root) or rel.startswith(".."):
        os.makedirs(path, exist_ok=True)
        return
    target = os.path.join(root, rel, os.path.basename(path))
    os.makedirs(target, exist_ok=True)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    os.symlink(target, path)


def export_surfaces(field, level, pad_value, voxel_size_m, analysis_dir, iter_name, **kwargs) -> dict:
    """build_surfaces() plus writing {iter_name}_<tag>_raw.stl / _filtered.stl into analysis_dir
    (tag "alphawrap" or "mc"). Returns the info dict with the file names and the warnings."""
    os.makedirs(analysis_dir, exist_ok=True)
    res = build_surfaces(field, level, pad_value, voxel_size_m, **kwargs)
    info = res["info"]
    tag = "alphawrap" if info["method"] == "alpha_wrap" else "mc"
    raw_path = os.path.join(analysis_dir, f"{iter_name}_{tag}_raw.stl")
    filt_path = os.path.join(analysis_dir, f"{iter_name}_{tag}_filtered.stl")
    write_stl(raw_path, *res["raw"])
    write_stl(filt_path, *res["filtered"])
    info["files"] = {"raw": raw_path, "filtered": filt_path}
    info["warnings"] = check_surface(info)
    return info


def summary_line(info: dict) -> str:
    """One line for the console: settings, size, time, volume ratio."""
    f = info["filtered"]
    if info["method"] == "alpha_wrap":
        s = (f"alpha-wrap alpha={info['alpha_voxels']:.2f} voxels ({info['alpha_mm']:.3g} mm), "
             f"offset={info['offset_voxels']:.3f} voxels ({info['offset_mm']:.3g} mm)")
    else:
        s = "marching cubes"
    c = info["components"]
    cav = info.get("cavities", {}).get("count", 0)
    if cav:
        s += f", {cav} sealed cavit{'y' if cav == 1 else 'ies'} kept as marching-cubes shell(s)"
    return (f"{s}; {f['faces']} faces, volume {info['volume_ratio_filtered']:.3f} x the solid voxels, "
            f"{info['total_time_s']:.1f} s (wrap {info['wrap_time_s']:.1f} s); "
            f"{c['n_removed']} of {c['n_components']} component(s) removed from the filtered surface")


def _jsonable(x):
    if isinstance(x, dict):
        return {str(k): _jsonable(v) for k, v in x.items()}
    if isinstance(x, (list, tuple)):
        return [_jsonable(v) for v in x]
    if isinstance(x, np.ndarray):
        return x.tolist()
    if isinstance(x, (np.integer,)):
        return int(x)
    if isinstance(x, (np.floating,)):
        return float(x)
    if isinstance(x, (np.bool_,)):
        return bool(x)
    return x


# --------------------------------------------------------------------------------------
# Command line: one export job in its own process
# --------------------------------------------------------------------------------------

def run_surface_job(job_npz: str) -> int:
    """
    Job file: an .npz with `field` and `config_json`, a JSON object with level, pad_value,
    voxel_size_m, analysis_dir, iter_name, result_path and the keyword arguments of
    build_surfaces (method, alpha_voxels, offset_ratio, ...). The result (the info dict, or the
    error) is written to result_path.
    """
    z = np.load(job_npz, allow_pickle=False)
    cfg = json.loads(str(z["config_json"]))
    result_path = cfg.pop("result_path")
    try:
        info = export_surfaces(z["field"], cfg.pop("level"), cfg.pop("pad_value"),
                               cfg.pop("voxel_size_m"), cfg.pop("analysis_dir"), cfg.pop("iter_name"),
                               **cfg)
        payload = {"ok": True, "info": info}
        code = 0
    except Exception as e:  # reported to the parent, which prints it
        payload = {"ok": False, "error": f"{type(e).__name__}: {e}", "traceback": traceback.format_exc()}
        code = 1
    with open(result_path, "w") as f:
        json.dump(_jsonable(payload), f, indent=2)
    return code


# ====================================================================================================
# Provenance of a run (formerly provenance.py)
# ====================================================================================================
# Provenance of a run (docs/theory.md §8.1): enough to say exactly which code,
# which inputs, which libraries and which machine produced a result, and to
# repeat it.
#
# `collect_provenance()` returns a JSON-serialisable dict; optimise.py prints it into
# problem_description.txt and the optimiser stores it in run_info.json. Every
# item is best effort: a missing git, a package that is not installed or an
# unreadable file is recorded as None rather than raising.

# The modules whose exact contents define a run's behaviour.
# (paths relative to the code directory, which is the repository root: baseline_optimiser.py lives there.)
CODE_FILES = ("optimise.py", "pcls_optimiser.py", "problem_parser.py", "analysis.py", "gpu_analysis.py",
              "interfaces.py", "baseline_optimiser.py", "simp_optimiser.py", "levelset_optimiser.py")

PACKAGES = ("numpy", "scipy", "matplotlib", "scikit-fem", "numba", "pyvista", "trimesh", "tetgen",
            "scikit-image", "pypardiso", "meshio", "Pillow")

SLURM_KEYS = ("SLURM_JOB_ID", "SLURM_ARRAY_JOB_ID", "SLURM_ARRAY_TASK_ID", "SLURM_JOB_NAME",
              "SLURM_JOB_PARTITION", "SLURM_JOB_NODELIST", "SLURM_CPUS_ON_NODE", "SLURM_MEM_PER_NODE",
              "SLURM_JOB_ACCOUNT", "SLURM_SUBMIT_DIR")


def sha256_of(path: str) -> str | None:
    try:
        h = hashlib.sha256()
        with open(path, "rb") as f:
            for chunk in iter(lambda: f.read(1 << 20), b""):
                h.update(chunk)
        return h.hexdigest()
    except Exception:
        return None


def _git(code_dir: str, *args) -> str | None:
    try:
        out = subprocess.run(["git", "-C", code_dir, *args], capture_output=True, text=True, timeout=10)
        return out.stdout.strip() if out.returncode == 0 else None
    except Exception:
        return None


def package_versions() -> dict:
    try:
        from importlib.metadata import version, PackageNotFoundError
    except Exception:  # pragma: no cover
        return {}
    out = {}
    for name in PACKAGES:
        try:
            out[name] = version(name)
        except PackageNotFoundError:
            out[name] = None
        except Exception:
            out[name] = None
    return out


def gpu_name() -> str | None:
    """The CUDA device's name, only if numba.cuda is already loaded (never initialises CUDA itself)."""
    if "numba.cuda" not in sys.modules:
        return None
    try:
        from numba import cuda
        if not cuda.is_available():
            return None
        name = cuda.get_current_device().name
        return name.decode() if isinstance(name, bytes) else str(name)
    except Exception:
        return None


def collect_provenance(code_dir: str | None = None, problem_file: str | None = None, overrides=None,
            random_seed=None) -> dict:
    code_dir = os.path.abspath(code_dir or os.path.dirname(os.path.abspath(__file__)))
    status = _git(code_dir, "status", "--porcelain")
    info = {
        "timestamp_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "host": socket.gethostname(),
        "platform": platform.platform(),
        "python": sys.version.split()[0],
        "command_line": list(sys.argv),
        "working_directory": os.getcwd(),
        "code_directory": code_dir,
        "git": {
            "commit": _git(code_dir, "rev-parse", "HEAD"),
            "branch": _git(code_dir, "rev-parse", "--abbrev-ref", "HEAD"),
            "describe": _git(code_dir, "describe", "--always", "--dirty", "--tags"),
            "dirty": (bool(status) if status is not None else None),
        },
        "code_sha256": {f: sha256_of(os.path.join(code_dir, f)) for f in CODE_FILES
                        if os.path.isfile(os.path.join(code_dir, f))},
        "packages": package_versions(),
        "gpu": gpu_name(),
        "cuda_visible_devices": os.environ.get("CUDA_VISIBLE_DEVICES"),
        "slurm": {k: os.environ[k] for k in SLURM_KEYS if k in os.environ},
    }
    if problem_file:
        info["problem_file"] = {"path": os.path.abspath(problem_file), "sha256": sha256_of(problem_file)}
    if overrides:
        info["overrides"] = list(overrides)
    if random_seed is not None:
        info["random_seed"] = int(random_seed)
    return info


def provenance_as_text(info: dict, indent: str = "    ") -> str:
    """A readable block for problem_description.txt."""
    lines = []
    git = info.get("git") or {}
    lines.append(f"{indent}- Started (UTC): {info.get('timestamp_utc')}")
    lines.append(f"{indent}- Host: {info.get('host')} ({info.get('platform')}), Python {info.get('python')}")
    if info.get("gpu"):
        lines.append(f"{indent}- GPU: {info['gpu']}")
    lines.append(f"{indent}- Code: {info.get('code_directory')}")
    lines.append(f"{indent}- Git: {git.get('describe') or git.get('commit') or 'not a git checkout'}"
                 + (" (UNCOMMITTED CHANGES)" if git.get("dirty") else ""))
    for f, h in (info.get("code_sha256") or {}).items():
        lines.append(f"{indent}    {f}: {h[:16] if h else None}")
    pf = info.get("problem_file")
    if pf:
        lines.append(f"{indent}- Problem file: {pf.get('path')} (sha256 {str(pf.get('sha256'))[:16]})")
    if info.get("overrides"):
        lines.append(f"{indent}- Command-line overrides: {', '.join(info['overrides'])}")
    if info.get("random_seed") is not None:
        lines.append(f"{indent}- Random seed: {info['random_seed']}")
    pk = info.get("packages") or {}
    lines.append(f"{indent}- Packages: " + ", ".join(f"{k} {v}" for k, v in pk.items() if v))
    if info.get("slurm"):
        lines.append(f"{indent}- SLURM: " + ", ".join(f"{k}={v}" for k, v in info["slurm"].items()))
    return "\n".join(lines) + "\n"


class BaseOptimiser:
    """
    The neutral base of every optimiser (docs/theory.md §8.1, §12.1): PCLS_Optimiser (pcls_optimiser.py) and the two
    comparison baselines through BaselineOptimiser (SIMP, simp_optimiser.py; the classical level set,
    levelset_optimiser.py) all derive from it. It owns what does not depend on how a method moves its design:

      * the construction (analysis factory, directories, masks) and the parameters every method shares
        (`initialise_common_parameters`, COMMON_PARAMETERS) with their parameters.csv record;
      * the design-field conventions: phi in [V, S] = [1, 2], the density rho(phi) and the penalised model density;
      * the run (`run_optimisation`) and everything it writes -- variables.csv, nucleation_events.csv, history/,
        outcome_stats.csv, final_fields.npz, run_info.json, threshold_sweep.csv, the surfaces -- in terms of phi;
      * the diagnostics (crispness, connectivity, grey fraction), the stopping tests and the time limit;
      * the buckling bookkeeping that is method-independent.

    A method provides `_setup_phi` (its initial design) and `_outer_loop` (its iteration), and `_variables_na` (the
    variables.csv columns that do not apply to it).
    """

    # Parameters shared by every optimiser (initialise_common_parameters):
    # same names, meanings and defaults (docs/theory.md §12.1, table 10.10).
    #
    # The linear buckling constraint (G3) parameters have PCLS's names and
    # meanings too and are accepted by every method; a method that does not use
    # one lists it in UNUSED_PARAMETERS, which optimise.py drops with the same
    # warning as it gives for PCLS-only keys (docs/theory.md §12.6).
    BUCKLING_PARAMETERS = (
        "buckling_enabled", "buckling_safety_factor", "buckling_safety_factor_initial",
        "buckling_safety_factor_ramp_outer_iters", "buckling_num_modes", "buckling_aggregation_p",
        "buckling_constraint_measure", "buckling_evaluate_every", "buckling_activate_after_outer_iter",
        "buckling_min_energy_fraction", "buckling_mode_continuity_warning_threshold",
        "buckling_min_density_for_geometric_stiffness", "buckling_lambda3_initial",
        "buckling_lambda3_clamp_max", "buckling_mu3", "buckling_mu3_decay_factor", "buckling_mu3_floor",
        "buckling_term_clamp",
        # Since the 2026-10 review: the eigensolver tolerance (was dropped, so a baseline always
        # ran at 1e-12 whatever the file said), the exported design's load factor of an
        # unconstrained run (a twin of a buckling run, docs/theory.md §6.10), and the ramp
        # threshold on the design density (eq. 6.10c).
        "buckling_eigensolver_tol", "buckling_report_exported", "buckling_ramp_on_design_density",
    )
    UNUSED_PARAMETERS = ()

    # A PCLS-only setting that the shared stopping test reads (`_buckling_converged`): PCLS sets it from its
    # parameter buckling_settle_requires_full_ramp; a baseline settles as soon as its own tests pass.
    buckling_settle_requires_full_ramp = False

    COMMON_PARAMETERS = (
        "outer_loop_max_iters", "inner_loop_max_iters", "outer_loop_convergence_threshold",
        "target_volume", "void_density", "solid_value", "void_value", "save_frequency",
        "random_seed", "initial_design", "run_until_converged", "time_limit_hours",
        "convergence_window", "convergence_compliance_tol", "convergence_switch_tol",
        "convergence_require_connected", "convergence_detached_tol", "convergence_exported_volume_tol",
        "volume_multiplier_reset_tol", "history_every", "diagnostics_every", "exported_analysis_every", "postprocess",
        "threshold_sweep_levels", "threshold_sweep_analyses",
        "final_surface_alpha_proportion", "final_surface_offset_proportion",
        "final_surface_num_cells_cutoff", "final_surface_method", "final_surface_alpha_voxels",
        "final_surface_offset_ratio", "final_surface_min_component_fraction",
        "final_surface_smoothing_iterations", "periodic_surface_method",
        "periodic_surface_alpha_voxels", "surface_export_background",
    ) + BUCKLING_PARAMETERS

    def __init__(self, analysis_component_factory: Callable[..., AnalysisInterface],
                 domain_dims: tuple[int, int, int],
                 printouts: str="Normal",
                 optimisation_name: str="PCLS_Optimisation",
                 analysis_dir: str=".", indent="  ",
                 csg_data: dict | None = None):
        """
        Initialise the PCLS Optimiser class.
        
        Args:
        - analysis_component_factory (callable): A factory function to create analysis components.
        - domain_dims (tuple): The dimensions of the optimisation domain.
        - printouts (str): Level of printouts ("Minimal", "Normal", "Verbose").
        - optimisation_name (str): Name of the optimisation run.
        - analysis_dir (str): Directory to run the analysis in.
        - indent (str): Indentation string used for print formatting.
        """
        # Set the printout level
        if printouts not in ["Silent", "Minimal", "Normal", "Verbose"]:
            raise ValueError("printouts must be one of 'Silent', 'Minimal', 'Normal', or 'Verbose'")
        self.printouts = printouts
        
        # Set the analysis class
        if not callable(analysis_component_factory):
            raise ValueError("analysis_component_factory must be a callable that returns an AnalysisInterface instance.")
        test_analysis_component = analysis_component_factory(iter_name="test", printouts="Silent")
        if not isinstance(test_analysis_component, AnalysisInterface):
            raise ValueError("The object returned by analysis_component_factory must implement the AnalysisInterface.")
        self.analysis_component_factory = analysis_component_factory

        # CG warm start (docs/theory.md, "CG warm start"): detect whether this
        # run's backend is a ConjugateGradient constructed with warm_start_cg
        # =True (only that backend ever sets this attribute; Conformal and
        # Regular instances don't have it, so getattr's default keeps them
        # unaffected). When it is, a persistent dict is created here to carry
        # the previous inner iteration's displacement/density across the many
        # ConjugateGradient instances one run creates (a fresh one every
        # inner iteration, per _inner_loop below) -- and is passed in ONLY at
        # the real inner-loop call site, never at a probe or the final/
        # exported analysis, so those always solve cold and never read or
        # write it (they'd otherwise contaminate the state with a
        # displacement field that doesn't belong to the trajectory).
        self._warm_start_cg_enabled = bool(getattr(test_analysis_component, "warm_start_cg", False))
        self._cg_warm_state = {"u_prev": None, "rho_prev": None} if self._warm_start_cg_enabled else None

        # Record whether this analysis backend supports linear buckling, so
        # that enabling the buckling constraint can be rejected up-front in
        # initialise_optimisation_parameters rather than blowing up several
        # hours into a run. Buckling support is optional by design: a
        # backend implementing only AnalysisInterface remains valid.
        self.analysis_supports_buckling = isinstance(test_analysis_component, BucklingAnalysisInterface)
        self._analysis_class_name = type(test_analysis_component).__name__
        self._analysis_field_type = getattr(test_analysis_component, "field_type", None)
        
        # Check the optimisation name is a valid directory name
        if not optimisation_name.isidentifier():
            invalid_chars = list(set(c for c in optimisation_name if not (c.isalnum() or c == '_')))
            error_msg = f"Invalid optimisation_name: '{optimisation_name}'.\n"
            if optimisation_name and optimisation_name[0].isdigit():
                error_msg += "  - Reason: It cannot start with a digit.\n"
            if invalid_chars:
                error_msg += f"  - Reason: It contains invalid characters: {invalid_chars}\n"
            error_msg += "  - Fix: Please use only alphanumeric characters and underscores."
            raise ValueError(error_msg)
            
        self.optimisation_name = optimisation_name
        # Create the optimisation directory
        os.makedirs(self.optimisation_name, exist_ok=True)

        # Check the analysis directory is a valid directory name
        if not analysis_dir.isidentifier():
            invalid_chars = list(set(c for c in analysis_dir if not (c.isalnum() or c == '_')))
            error_msg = f"Invalid analysis_dir: '{analysis_dir}'.\n"
            if analysis_dir and analysis_dir[0].isdigit():
                error_msg += "  - Reason: It cannot start with a digit.\n"
            if invalid_chars:
                error_msg += f"  - Reason: It contains invalid characters: {invalid_chars}\n"
            error_msg += "  - Fix: Please use only alphanumeric characters and underscores."
            raise ValueError(error_msg)
        
        self.analysis_dir = os.path.join(self.optimisation_name, analysis_dir)
        # Create the analysis directory
        make_bulk_dir(self.analysis_dir)

        # Check the validity of the domain dimensions
        if len(domain_dims) != 3 or not all(isinstance(dim, int) and dim > 0 for dim in domain_dims):
            raise ValueError("domain_dims must be a tuple of three positive integers.")
        self.domain_dims = domain_dims

        self.indent = indent

        # (printed here, not beside the detection above: _print_out needs self.indent)
        if self._warm_start_cg_enabled:
            self._print_out("CG warm start ENABLED: the main inner-loop PCG solve will be seeded "
                             "from the previous inner iteration's displacement field when the "
                             "safeguards allow it; probe and final/exported solves always run cold.",
                             "Minimal")

        self.csg_data = csg_data
        if self.csg_data is not None:
            self.preserve_mask = self.csg_data.get("preserve_elements")
            self.obstacle_mask = self.csg_data.get("obstacle_elements")
        else:
            raise Exception("Required domain CSG data is missing.")

        # Voxel masks used by the connectivity diagnostic (_connectivity_report,
        # docs/theory.md §5.6): which voxels carry a load and which carry a
        # Dirichlet constraint. Both are taken straight from the same CSG data
        # the analysis backends smear onto their nodes, so "connected" here
        # means connected between exactly the loads and supports the FE model
        # sees.
        load_vectors = self.csg_data.get("load_vectors")
        fixed_nodes = self.csg_data.get("fixed_nodes")
        self.load_mask = (np.any(np.asarray(load_vectors) != 0.0, axis=-1)
                          if load_vectors is not None else np.zeros(domain_dims, dtype=bool))
        self.support_mask = (np.any(np.asarray(fixed_nodes), axis=-1)
                             if fixed_nodes is not None else np.zeros(domain_dims, dtype=bool))
        self._resolve_passive_overlap()
        self.design_mask = ~(np.asarray(self.preserve_mask, dtype=bool) | np.asarray(self.obstacle_mask, dtype=bool))
        self.n_design_voxels = max(1, int(np.count_nonzero(self.design_mask)))

        self._print_out(f"Initialised {type(self).__name__}:\n"
                          f"  Optimisation Name: {self.optimisation_name}\n"
                          f"  Analysis Directory: {self.analysis_dir}\n"
                          f"  Domain Dimensions: {self.domain_dims}\n"
                          f"  Printout Level: {self.printouts}",
                        "Verbose")

    def _resolve_passive_overlap(self):
        """
        Make preserve and obstacle disjoint, obstacle winning, so that every method
        sees the same passive voxels. A voxel in both (e.g. an obstacle cylinder that
        overlaps a bolted pad's ring after voxelisation) used to be void in PCLS and
        level-set (obstacle applied after preserve) but gave SIMP's MMA a lower bound
        above its upper bound. The counts are kept for run_info.json; a support voxel
        left inside an obstacle is only reported, as its Dirichlet nodes are fixed by
        the backend whatever the density.
        """
        preserve = np.asarray(self.preserve_mask, dtype=bool)
        obstacle = np.asarray(self.obstacle_mask, dtype=bool)
        overlap = preserve & obstacle
        self.n_preserve_obstacle_overlap = int(np.count_nonzero(overlap))
        self.n_support_in_obstacle = int(np.count_nonzero(np.asarray(self.support_mask, dtype=bool) & obstacle))
        self.n_load_in_obstacle = int(np.count_nonzero(np.asarray(self.load_mask, dtype=bool) & obstacle))
        if self.n_preserve_obstacle_overlap:
            self.preserve_mask = preserve & ~obstacle
            self._print_out(f"Warning: {self.n_preserve_obstacle_overlap} voxels are in both a preserve and an "
                            f"obstacle region; they are treated as obstacle (void). Check the problem geometry.",
                            "Minimal")
        if self.n_support_in_obstacle or self.n_load_in_obstacle:
            self._print_out(f"Warning: {self.n_support_in_obstacle} support voxels and {self.n_load_in_obstacle} "
                            f"load voxels lie inside obstacle regions (void), so their constraints/loads act on "
                            f"void material.", "Minimal")

    def _print_out(self, message, message_type):
        """
        Print messages based on the printout level.

        Args:
            message (str): The message to print.
            message_type (str): The type of message ("Minimal", "Normal", "Verbose").
        """
        
        # 1. Define the indent level based *only* on the message_type
        if message_type == "Minimal":
            indent_level = 0
        elif message_type == "Normal":
            indent_level = 1
        elif message_type == "Verbose":
            indent_level = 2
        else:
            # Don't print "Silent" or unknown types
            return

        # 2. Define a hierarchy of levels
        levels = {
            "Minimal": 1,
            "Normal": 2,
            "Verbose": 3
        }

        # Get the numeric level for the current setting and message
        printout_level = levels.get(self.printouts, 0) # 0 for "Silent"
        message_level = levels.get(message_type, 0)

        # 3. Check if the message's level is at or below the global setting
        # (e.g., if setting is "Normal" (2), print "Minimal" (1) and "Normal" (2))
        if message_level > printout_level:
            return
        
        # 4. If we are here, print
        # Build the full indent prefix
        prefix = self.indent * indent_level

        # Apply the indentation to all lines in a multi-line message
        formatted_message = "\n".join(prefix + line for line in str(message).splitlines())

        # Print the result
        print(formatted_message)

    def initialise_common_parameters(self, outer_loop_max_iters: int=25,
                                            inner_loop_max_iters: int=10,
                                            outer_loop_convergence_threshold: float=1e-2,
                                            target_volume: float=0.5,
                                            void_density: float=1e-3,
                                            solid_value: float=2.0,
                                            void_value: float=1.0,
                                            save_frequency: int=10,
                                            buckling_enabled: bool=False,
                                            buckling_safety_factor: float=2.0,
                                            buckling_safety_factor_initial: float|None=None,
                                            buckling_safety_factor_ramp_outer_iters: int=3,
                                            buckling_num_modes: int=6,
                                            buckling_aggregation_p: float=8.0,
                                            buckling_constraint_measure: str="critical",
                                            buckling_evaluate_every: int=1,
                                            buckling_activate_after_outer_iter: int=2,
                                            buckling_min_energy_fraction: float=0.5,
                                            buckling_mode_continuity_warning_threshold: float=0.7,
                                            buckling_min_density_for_geometric_stiffness: float=0.1,
                                            buckling_lambda3_initial: float=0.0,
                                            buckling_lambda3_clamp_max: float=3.0,
                                            buckling_mu3: float=0.1,
                                            buckling_mu3_decay_factor: float=0.9,
                                            buckling_mu3_floor: float=0.01,
                                            buckling_term_clamp: float=10.0,
                                            buckling_eigensolver_tol: float=1e-12,
                                            buckling_report_exported: bool=False,
                                            buckling_ramp_on_design_density: bool=True,
                                            final_surface_alpha_proportion: float|None=None,
                                            final_surface_offset_proportion: float|None=None,
                                            final_surface_num_cells_cutoff: float|None=None,
                                            final_surface_method: str="alpha_wrap",
                                            final_surface_alpha_voxels: float=1.5,
                                            final_surface_offset_ratio: float=1.0/30.0,
                                            final_surface_min_component_fraction: float=0.01,
                                            final_surface_smoothing_iterations: int=0,
                                            periodic_surface_method: str="alpha_wrap",
                                            periodic_surface_alpha_voxels: float=1.5,
                                            surface_export_background: bool=False,
                                            convergence_window: int=10,
                                            convergence_compliance_tol: float=1e-3,
                                            convergence_switch_tol: float=1e-4,
                                            convergence_require_connected: bool=True,
                                            convergence_detached_tol: float=0.005,
                                            convergence_exported_volume_tol: float=0.15,
                                            volume_multiplier_reset_tol: float|None=None,
                                            run_until_converged: bool=False,
                                            time_limit_hours: float|None=None,
                                            random_seed: int|None=42,
                                            initial_design: str|None=None,
                                            history_every: int=10,
                                            diagnostics_every: int=1,
                                            exported_analysis_every: int=10,
                                            postprocess: bool=True,
                                            threshold_sweep_levels: list|None=None,
                                            threshold_sweep_analyses: bool=True):
        """
        Sets, validates and records the parameters every optimiser shares (BaseOptimiser.COMMON_PARAMETERS):
        the target volume, the design-field values, the iteration budget, the stopping and time-limit
        controls, the output cadences, the final/periodic surface export, the initial design and the linear
        buckling constraint. PCLS_Optimiser.initialise_optimisation_parameters calls it with these and sets its
        own; the SIMP and level-set baselines call it with these alone. Names, meanings, defaults and
        validation are the same for every method (docs/theory.md §12.1, table 10.10).

        Args:
        - outer_loop_max_iters (int): Maximum iterations for the outer loop.
        - inner_loop_max_iters (int): Maximum iterations for the inner loop.
        - outer_loop_convergence_threshold (float): tolerance of the settled
          test on the volume error |G1| = |mean(rho) - V0|, a fraction of the
          DOMAIN volume like the grey fraction (0.01 = one percentage point;
          docs/theory.md §5.5 explains why not a fraction of V0) and, with
          buckling on, RELATIVE tolerance on the true critical load factor
          (Lambda_crit >= SF/(1 + tol), docs/theory.md eq. 6.27).
        - target_volume (float): Target volume fraction.
        - void_density (float): Density of void elements.
        - solid_value (float): The value representing the solid phase in the density field.
        - void_value (float): value of the design field in void (default 1.0). Must be 2 - 1: solid_value - void_value = 1 (eq. 2.5);
          SIMP and the level set additionally need solid_value = 2 and void_value = 1 (eq. 12.5).
        - save_frequency (int): every this many iterations the iteration's analysis is saved (.vtr, meshes) and the periodic
          surface exported (default 10, at least 1).
        - buckling_enabled (bool): Turn the buckling constraint on.
        - buckling_safety_factor (float): Target minimum critical load factor.
        - buckling_safety_factor_initial (float|None): Start of the safety
          factor ramp. None means start at the target (no ramp). Ramping is
          strongly recommended: switching the full target on at once fights
          the volume constraint hard and tends to stall the topology.
        - buckling_safety_factor_ramp_outer_iters (int): Number of outer
          iterations over which the safety factor ramps from initial to
          target, counted from activation.
        - buckling_num_modes (int): Modes to retain and aggregate. Must be
          large enough that the retained set does not end part-way through a
          cluster of repeated eigenvalues; the analysis component warns when
          it does. 6-12 is typical in 3D, and symmetric structures need even
          counts because their modes come in degenerate pairs.
        - buckling_aggregation_p (float): p-norm aggregation exponent (8-16).
        - buckling_constraint_measure (str): What the constraint VALUE g3
          measures (docs/theory.md §6.7, eq. 6.23).
          "critical" (default): the true critical factor, g3 = mu_crit -
          1/SF. This is the p-norm aggregate rescaled by mu_crit/mu_tilde at
          every evaluation (adaptive normalisation). The update direction is
          still the smooth aggregate's gradient, whose scale the normalised
          buckling term discards, so the constraint the multiplier drives to
          is exactly the one the settled test checks (eq. 6.27).
          "aggregate": the previous behaviour, g3 = mu_tilde - 1/SF. Because
          mu_tilde >= mu_crit (by up to m^(1/p): 9 % for a degenerate pair at
          p = 8), this demands a hidden margin on top of the safety factor.
          Where that margin does not fit in the volume budget, lambda3 winds
          up to its clamp and the design never settles.
        - buckling_evaluate_every (int): Evaluate buckling every N inner
          iterations, reusing the previous sensitivity in between. The
          eigensolve is by far the most expensive part of an iteration, so
          this is the main cost control.
        - buckling_activate_after_outer_iter (int): Outer iteration from
          which the constraint becomes active. Running buckling on the
          initial near-uniform grey field is both meaningless and expensive.
        - buckling_min_energy_fraction (float): Minimum fraction of a mode's
          modal strain energy carried by elements that contribute geometric
          stiffness -- weighted by the geometric-stiffness ramp r(rho_e) --
          for the mode to be trusted (spurious void-mode filter,
          docs/theory.md §6.5(a)). Default 0.5 (was 0.9 against a hard
          rho >= 0.5 cut, which rejected genuine critical modes: on a test
          column the true critical mode, lambda = 0.50, carried 89 % of its
          energy in such elements and was discarded in favour of a mode at
          lambda = 2.97). If every candidate is rejected the evaluation is
          "unknown" and is never counted as safe.
        - buckling_mode_continuity_warning_threshold (float): Subspace-continuity value
          below which the analysis component notes that mode tracking has
          drifted. This measures each mode's overlap with the SPAN of the
          previous mode set, which (unlike a pairwise MAC) is unaffected by
          the arbitrary rotation of eigenvectors inside a degenerate cluster.
          Logged per evaluation in buckling.csv as min_subspace_continuity.
        - buckling_min_density_for_geometric_stiffness (float): DESIGN density
          theta below which geometric stiffness is ramped out (eq. 6.10: none
          below theta, all above 2 theta).
        - buckling_lambda3_initial (float): Initial Lagrange multiplier.
        - buckling_lambda3_clamp_max (float): Anti-windup ceiling, applied to
          both lambda3 and the modified multiplier lambda3_tilde (default 3.0;
          was 20.0, at which a saturated buckling term outweighed the
          stiffness and volume terms together). The latter
          is what scales the buckling term in the update, so this is the
          practical guard against the buckling constraint overwhelming the
          compliance and volume drivers.
          Scale reference for tuning: 'buckling_term' is normalised to a mean
          absolute value of exactly 1.0 and clamped per-voxel to
          +/-'buckling_term_clamp'. The applied contribution is therefore
          lambda3_tilde on average and up to
          lambda3_tilde * buckling_term_clamp at a localised spike. Compare
          against 'stiffness_term', which is normalised the same way (mean
          1.0) and clamped at 10.0. So this ceiling is roughly 'how many
          times stronger than the compliance driver may the buckling
          constraint push, on average'. Reduce it if the volume fraction
          runs away or phi oscillates once buckling activates.
        - buckling_mu3 (float): Initial penalty parameter (the penalty
          weight is 1/mu3, so smaller mu3 means a harder constraint).
        - buckling_mu3_decay_factor (float): Decay factor for mu3, applied
          after an outer iteration only while G3 is violated and its
          violation shrank by less than 25 % (the adaptive rule mu1 follows).
          Default 0.9 (was 0.75, applied unconditionally).
        - buckling_mu3_floor (float): Lower bound on mu3. Default 0.01 (was
          1e-4, at which lambda3_tilde became bang-bang).
        - buckling_term_clamp (float): Symmetric clamp on the normalised
          buckling update term, mirroring the stiffness term's clamp.
        - buckling_eigensolver_tol (float): ARPACK's relative eigenvalue
          tolerance for the buckling eigensolve (1e-12; 0 = machine precision,
          ARPACK's own default). Each Lanczos iteration is a PARDISO
          back-substitution on the whole stiffness matrix, so a looser value
          saves time, BUT a loose value (1e-6 .. 1e-10) can miss one copy of a
          repeated eigenvalue, which corrupts the sensitivity on a symmetric
          design (a square column): 1e-12 showed no such failure in a stress
          test, 1e-6 failed 3 of 24 and 1e-10 2 of 56. docs/theory.md §6.10.
        - buckling_report_exported (bool): analyse the buckling of the exported
          design at the end of the run even when the constraint is disabled
          (False). Reporting only: one eigensolve after the last iteration,
          written to run_info.json (buckling.critical_load_factor_exported,
          exported_load_factors), the summary and the outcome statistics.
          It does not change the optimisation, so it is how the load factor
          of an unconstrained run is measured for comparison with a
          constrained one. Ignored when the constraint is enabled (the
          exported design is then always analysed).
        - buckling_ramp_on_design_density (bool): True (default since the
          2026-10 review) applies theta to the design density: with the
          penalised model (p > 1) the backend, which only sees the stiffness
          density rho_m = c1 + (1 - c1) x^p, is handed the equivalent model
          threshold c1 + (1 - c1) ((theta - c1)/(1 - c1))^p (docs/theory.md
          §6.5(b), eq. 6.10c), so grey material keeps the geometric stiffness
          its stress gives it. False hands theta to the backend unchanged, the
          earlier behaviour, under which at p = 3 every voxel with design
          density below 0.46 had no geometric stiffness at all and a grey
          column read as up to 100 times stiffer against buckling than it is.
          Identical either way at p = 1 and for binary designs.
        - final_surface_alpha_proportion / final_surface_offset_proportion
          (float|None): The previous way of setting the finished design's
          wrap, as fractions of the bounding-box diagonal of the mesh. When
          given they take precedence over the voxel-based parameters.
        - final_surface_num_cells_cutoff (float|None): DEPRECATED and
          ignored (components are now filtered on the voxel grid by
          final_surface_min_component_fraction); still accepted so that old
          problem files load.
        - final_surface_method (str): "alpha_wrap" (default) or
          "marching_cubes" (no CGAL; the marching-cubes mesh is closed and
          exact to the field).
        - final_surface_alpha_voxels (float): Alpha in voxels, > 0. Default
          1.5; a warning is printed above 4 (features narrower than a few
          voxels are filled).
        - final_surface_offset_ratio (float): offset / alpha, > 0. Default
          1/30, the ratio of CGAL's examples; the offset adds a shell of
          volume ~ offset x area.
        - final_surface_min_component_fraction (float): Components with fewer
          than this fraction of the solid voxels are left out of the
          "filtered" surface and are plain shells in the "raw" one. In
          [0, 1); 0 keeps everything. Default 0.01.
        - final_surface_smoothing_iterations (int): Taubin smoothing
          iterations of the marching-cubes mesh (method "marching_cubes"
          only; needs trimesh). Default 0.
        - periodic_surface_method / periodic_surface_alpha_voxels: the same
          for the surfaces written at every save_frequency iterations
          (default "alpha_wrap" at 1.5 voxels, the same as the finished
          design; a wrap took about 3-5 s on a 160 x 40 x 40 design, so raise
          periodic_surface_alpha_voxels or use surface_export_background if
          saves are frequent).
        - surface_export_background (bool): Run the periodic exports in a
          separate process (python surface_export.py --job ...) so the
          optimisation loop is not blocked. Default False.
        - convergence_window (int): the design must satisfy every
          "settled" test on this many consecutive iterations (10).
        - convergence_compliance_tol (float): maximum relative change of the
          compliance between consecutive iterations (1e-3).
        - convergence_switch_tol (float): maximum fraction of design voxels
          that may change phase (cross the level set) per iteration (1e-4).
        - convergence_require_connected (bool): also require the exported
          geometry to be connected (every loaded voxel joined to a support
          through solid with no void gap -- voxels touching at a face, edge
          or corner count as joined -- and at most convergence_detached_tol
          of the solid volume outside such load paths) (True). With the
          penalised interpolation the design is expected to be connected
          by the time it settles; this is the safety net.
        - convergence_detached_tol (float): see above (0.005).
        - convergence_exported_volume_tol (float): also require the volume of
          the EXPORTED geometry (phi >= level set, plus preserve regions) to
          lie within this fraction of the target volume V0: |V_exported - V0|
          <= tol * V0 (0.15; 0 or less disables the test). The volume error
          |G1| of the settled test is measured on the density field, and a
          uniform void film (docs/theory.md §5.5, "void film") lifts that
          field above the void well without being exported, so a design can
          satisfy |G1| ~ 0 while exporting well under V0. A design that has
          settled and is connected but fails this test is treated like one
          that settled disconnected: lambda2 is reset (sharing the
          reconnection_max_resets budget) and the run continues; once the
          budget is spent the run is accepted as converged, with a
          warning and converged_with_volume_shortfall = true in run_info.json,
          so a run cannot be stranded until the time limit. See
          docs/theory.md §5.5.1.
        - volume_multiplier_reset_tol (float|None): anti-windup of the volume
          multiplier lambda1 of the augmented-Lagrangian methods (PCLS, level
          set). lambda1 integrates G1/mu1 only once per outer iteration, so it
          winds up while the design is far above the target and then keeps
          removing material for many outer iterations after the volume has
          fallen below it (Truss_Cantilever: PCLS starved at a third of its
          budget, the level set removed every design voxel). When the volume
          error has the opposite sign to the stored lambda1 and |G1| exceeds
          this tolerance (in the units of G1, a fraction of the domain), the
          stored lambda1 is dropped, in lambda1_tilde at every iteration and at
          the outer update. None (default) = outer_loop_convergence_threshold,
          the volume tolerance of the settled test, so lambda1 is dropped exactly
          when the volume error is too large for the design to converge; a wider
          band let a column design sit 9 % under its target with lambda1 > 0,
          failing the volume test forever. 0 or less disables it. The dead band
          keeps a converged design, whose G1 only fluctuates about zero, on its
          multiplier. See _volume_multiplier.
        - run_until_converged (bool): ignore the iteration budget
          (outer_loop_max_iters x inner_loop_max_iters) and keep iterating
          until the design converges (False). Combine with time_limit_hours
          on a scheduler.
        - time_limit_hours (float|None): wall-clock budget. When set, the run
          keeps iterating (the iteration budget is ignored) until it
          converges or the budget would be exceeded by the next iteration,
          then saves the final design as usual. Leave headroom below the
          scheduler's own walltime for the final save (None, or 0, = no
          limit). Measured from the start of run_optimisation.
        - random_seed (int|None): seed of the random noise in the initial
          field. The default is 42, so a run is repeatable unless asked
          otherwise; pass None (null in a problem file) for a seed drawn from
          the operating system. The seed used is printed and recorded in
          run_info.json, so any run can be repeated exactly by passing it
          back. The initial noise is
          numpy.random.RandomState(seed).rand(...), the same stream as
          numpy.random.seed(seed) followed by numpy.random.rand(...).
        - initial_design (str|None): start from a previous design instead of
          the default (PCLS: a uniform field plus noise; SIMP: a uniform
          density; level set: a solid block). The path of a .npz holding the
          field `phi` (a run's final_fields.npz or a history/phi_<n>.npz
          snapshot), a .npy of phi, or a run directory (its final_fields.npz);
          relative to the working directory, else to the problem file. phi
          must have the grid of this problem; it can come from any of the three
          methods (it is clipped to [void_value, solid_value]; SIMP uses
          (phi - V)/(S - V) as its design density, the level set the signed
          distance re-computed from phi = level set; docs/theory.md §8.1,
          eq. 8.1a). Preserve and obstacle voxels are re-imposed; PCLS adds no
          noise (initial_noise_multiplier is not used). A history snapshot is
          float16, accurate to ~1e-3. The file's path and SHA-256 are recorded
          in run_info.json. None (default) = the method's own start.
        - history_every (int): write a float16 snapshot of phi to
          history/phi_<iteration>.npz every this many iterations (10), for
          the evolution montage and animation and to re-run any diagnostic
          afterwards. 0 disables the snapshots; the per-iteration topology
          and nucleation record is written regardless.
        - diagnostics_every (int): compute the per-iteration diagnostics (topology,
          nucleation events, discreteness, undecided/void-film fractions,
          connectivity report in variables.csv) every this many iterations
          (1 = every iteration, the default, bit-exact with earlier runs).
          Skipped rows leave those columns 'N/A'. The optimisation path is
          unchanged: the quantities the stopping tests use (grey fraction,
          phase switches, exported volume, the connectivity needed to accept
          convergence) are still computed whenever they are needed, and the
          history snapshots keep their own cadence.
        - exported_analysis_every (int): every this many iterations, also
          analyse the geometry the run would export at that moment (phi >=
          level set made solid, the rest void) and log its compliance beside
          the model's (exported_compliance in variables.csv; density backends
          only -- Conformal already analyses exactly that geometry) (10). The
          ratio shows when the thresholded part stops matching the model --
          a grey link that carries load in the model but is absent from the
          export (docs/theory.md §4.7). One extra linear solve per this many
          iterations; 0 disables it.
        - postprocess (bool): at the end of the run, write final_fields.npz,
          run_info.json, summary.json/.txt, the plots in plots/ and
          report.pdf (the postprocess section below) (True). Any failure there is reported
          and never affects the run's other outputs.
        - threshold_sweep_levels (list|None): cut levels s at which the final
          design is re-thresholded, phi >= V + s (S - V), to test how much
          the result depends on the threshold (threshold_sweep.csv). None
          uses 0.1, 0.2, ..., 0.9; [] disables the sweep. s = 0.5 is the
          exported design.
        - threshold_sweep_analyses (bool): also run a finite element analysis
          of every cut (compliance; with buckling enabled, the critical load
          factor too) (True). Skipped automatically when the run stopped on
          its time limit, or when the analyses would not fit in what is left
          of it.
        """
        # Set target volume
        if not (0.0 < target_volume < 1.0):
            raise ValueError("target_volume must be between 0 and 1.")
        self.target_volume = target_volume

        # Set the field material values
        if solid_value <= void_value:
            raise ValueError("'solid_value' must be greater than 'void_value'.")
        if not (isinstance(solid_value, (int, float)) and isinstance(void_value, (int, float))):
            raise ValueError("solid_value and void_value must be numeric (int or float).")
        # rho(phi) = -c1 (phi - S) + (phi - V) (Wei & Wang 2009; theory eq. 2.5) is c1 at V
        # and 1 at S only when S - V = 1 (at S = 3, V = 1 the void had density
        # 2 c1 and phi = 2 was already solid); the grey band, the void film, the
        # mu1 stability floor and every export at phi = (S + V)/2 assume it too.
        if abs((float(solid_value) - float(void_value)) - 1.0) > 1e-12:
            raise ValueError(f"solid_value - void_value must be 1 (got {solid_value} - {void_value}): the density "
                             f"map rho(phi) of docs/theory.md eq. (2.5) is defined for S - V = 1 (the shipped "
                             f"files use S = 2, V = 1).")
        self.solid_value = solid_value
        self.void_value = void_value

        # Calculate the level set value
        self.level_set_value = (solid_value + void_value) / 2.0

        # Set the remaining parameters. Both loop counts must be at least 1: with
        # inner_loop_max_iters = 0 the inner loop returned at once, before its
        # time check, so a run_until_converged or time-limited run updated the
        # multipliers forever without iterating (optimiser review, 2026-10).
        for _name, _value in (("outer_loop_max_iters", outer_loop_max_iters),
                              ("inner_loop_max_iters", inner_loop_max_iters)):
            if int(_value) != _value or int(_value) < 1:
                raise ValueError(f"{_name} must be an integer >= 1, got {_value}.")
        self.outer_loop_max_iters = int(outer_loop_max_iters)
        self.inner_loop_max_iters = int(inner_loop_max_iters)
        self.outer_loop_convergence_threshold = outer_loop_convergence_threshold
        self.void_density = void_density
        self.save_frequency = max(1, save_frequency)

        # None means "inherit from the analysis component" -- resolved lazily
        # in _export_final_surface_alpha_wrap, since the analysis component
        # (and therefore its own alpha_proportion/offset_proportion/
        # num_cells_cutoff, when the backend has them at all) doesn't exist
        # yet at this point in the run.
        self.final_surface_alpha_proportion = final_surface_alpha_proportion
        self.final_surface_offset_proportion = final_surface_offset_proportion
        self.final_surface_num_cells_cutoff = final_surface_num_cells_cutoff
        if final_surface_num_cells_cutoff is not None:
            print("NOTE: final_surface_num_cells_cutoff is deprecated and ignored; the final surface is "
                  "filtered on the voxel grid by final_surface_min_component_fraction "
                  "(docs/theory.md §8.4).")
        for _name, _method in (("final_surface_method", final_surface_method),
                               ("periodic_surface_method", periodic_surface_method)):
            if _method not in ("alpha_wrap", "marching_cubes"):
                raise ValueError(f"{_name} must be 'alpha_wrap' or 'marching_cubes', got {_method!r}.")
        for _name, _val in (("final_surface_alpha_voxels", final_surface_alpha_voxels),
                            ("final_surface_offset_ratio", final_surface_offset_ratio),
                            ("periodic_surface_alpha_voxels", periodic_surface_alpha_voxels)):
            if not _val > 0.0:
                raise ValueError(f"{_name} must be positive, got {_val}.")
        if not 0.0 <= final_surface_min_component_fraction < 1.0:
            raise ValueError("final_surface_min_component_fraction must be in [0, 1), got "
                             f"{final_surface_min_component_fraction}.")
        self.final_surface_method = final_surface_method
        self.final_surface_alpha_voxels = float(final_surface_alpha_voxels)
        self.final_surface_offset_ratio = float(final_surface_offset_ratio)
        self.final_surface_min_component_fraction = float(final_surface_min_component_fraction)
        self.final_surface_smoothing_iterations = max(0, int(final_surface_smoothing_iterations))
        self.periodic_surface_method = periodic_surface_method
        self.periodic_surface_alpha_voxels = float(periodic_surface_alpha_voxels)
        self.surface_export_background = bool(surface_export_background)
        # --- Loop control and convergence (docs/theory.md §5.4-§5.6) ---
        self.convergence_window = max(1, int(convergence_window))
        self.convergence_compliance_tol = float(convergence_compliance_tol)
        self.convergence_switch_tol = float(convergence_switch_tol)
        self.convergence_require_connected = bool(convergence_require_connected)
        self.convergence_detached_tol = float(convergence_detached_tol)
        self.convergence_exported_volume_tol = float(convergence_exported_volume_tol)
        self.volume_multiplier_reset_tol = (None if volume_multiplier_reset_tol is None
                                            else float(volume_multiplier_reset_tol))
        self.run_until_converged = bool(run_until_converged)
        # None, or any non-positive value (the web app's Batch Edit writes 0
        # for "no limit"), means no time limit.
        if time_limit_hours is None or float(time_limit_hours) <= 0.0:
            self.time_limit_hours = None
        else:
            self.time_limit_hours = float(time_limit_hours)
        if self.time_limit_hours is not None:
            self._print_out(
                f"Time limit: {self.time_limit_hours:.3f} h. The run keeps iterating until it "
                f"converges or the next iteration would exceed the limit; the iteration budget "
                f"({outer_loop_max_iters} x {inner_loop_max_iters}) is ignored.", "Minimal")
        elif self.run_until_converged:
            self._print_out(
                "run_until_converged is set with no time_limit_hours: the iteration budget is "
                "ignored and the run only stops when the design has converged. Set a time limit "
                "(or a scheduler walltime) if a non-converging run must not run indefinitely.",
                "Minimal")

        # --- Diagnostics and post-processing (docs/theory.md §8.4.1-§8.4.3) ---
        if random_seed is not None and not (0 <= int(random_seed) < 2 ** 32):
            raise ValueError("random_seed must be an integer in [0, 2**32).")
        self.random_seed = None if random_seed is None else int(random_seed)
        if initial_design is not None and not (isinstance(initial_design, (str, os.PathLike)) and str(initial_design).strip()):
            raise ValueError("initial_design must be null or the path of a .npz / .npy / run directory.")
        self.initial_design = None if initial_design is None else str(initial_design)
        self.initial_design_info = None
        self.history_every = max(0, int(history_every))
        self.diagnostics_every = max(1, int(diagnostics_every))
        self._diag_fresh = True
        self._conn_fresh = True
        self._diag_pending = None
        self.exported_analysis_every = max(0, int(exported_analysis_every))
        self.postprocess = bool(postprocess)
        if threshold_sweep_levels is None:
            self.threshold_sweep_levels = [round(0.1 * k, 1) for k in range(1, 10)]
        else:
            levels = sorted({float(v) for v in threshold_sweep_levels})
            if any(not (0.0 < v < 1.0) for v in levels):
                raise ValueError("threshold_sweep_levels must lie strictly between 0 and 1.")
            self.threshold_sweep_levels = levels
        self.threshold_sweep_analyses = bool(threshold_sweep_analyses)

        # --- Buckling constraint (G3) configuration ---
        self.buckling_enabled = bool(buckling_enabled)

        if self.buckling_enabled:
            if not self.analysis_supports_buckling:
                raise ValueError(
                    f"buckling_enabled=True, but the configured analysis component "
                    f"('{self._analysis_class_name}') does not implement "
                    f"BucklingAnalysisInterface (it has no 'run_buckling' method). "
                    f"Either choose an analysis type that supports linear buckling, "
                    f"or set buckling_enabled=False in the problem file."
                )
            if buckling_safety_factor <= 0.0:
                raise ValueError("buckling_safety_factor must be positive.")
            if buckling_num_modes < 1:
                raise ValueError("buckling_num_modes must be at least 1.")
            if buckling_aggregation_p < 1.0:
                raise ValueError("buckling_aggregation_p must be >= 1.")
            if buckling_constraint_measure not in ("critical", "aggregate"):
                raise ValueError("buckling_constraint_measure must be 'critical' or 'aggregate', "
                                 f"got {buckling_constraint_measure!r}.")
            if buckling_mu3 <= 0.0:
                raise ValueError("buckling_mu3 must be positive (the penalty weight is 1/mu3).")

        self.buckling_safety_factor = buckling_safety_factor
        # None means "no ramp": start at the target. A ramp start above the
        # target would be a de-facto tightening, so it is clamped down.
        if buckling_safety_factor_initial is None:
            self.buckling_safety_factor_initial = buckling_safety_factor
        else:
            self.buckling_safety_factor_initial = min(buckling_safety_factor_initial,
                                                      buckling_safety_factor)
        self.buckling_safety_factor_ramp_outer_iters = max(0, buckling_safety_factor_ramp_outer_iters)
        self.buckling_num_modes = buckling_num_modes
        self.buckling_aggregation_p = buckling_aggregation_p
        self.buckling_constraint_measure = buckling_constraint_measure
        self.buckling_evaluate_every = max(1, buckling_evaluate_every)
        self.buckling_activate_after_outer_iter = max(0, buckling_activate_after_outer_iter)
        self.buckling_min_energy_fraction = buckling_min_energy_fraction
        self.buckling_mode_continuity_warning_threshold = buckling_mode_continuity_warning_threshold
        self.buckling_min_density_for_geometric_stiffness = buckling_min_density_for_geometric_stiffness
        self.buckling_lambda3_initial = buckling_lambda3_initial
        self.buckling_lambda3_clamp_max = buckling_lambda3_clamp_max
        self.buckling_mu3_initial = buckling_mu3
        self.buckling_mu3_decay_factor = buckling_mu3_decay_factor
        self.buckling_mu3_floor = buckling_mu3_floor
        self.buckling_term_clamp = abs(buckling_term_clamp)
        if buckling_eigensolver_tol < 0.0:
            raise ValueError("buckling_eigensolver_tol must be >= 0 (0 = machine precision).")
        self.buckling_eigensolver_tol = float(buckling_eigensolver_tol)
        self.buckling_report_exported = bool(buckling_report_exported)
        self.buckling_ramp_on_design_density = bool(buckling_ramp_on_design_density)

        if self.buckling_enabled and self.buckling_num_modes % 2 != 0:
            self._print_out(
                f"Note: buckling_num_modes={self.buckling_num_modes} is odd. Buckling "
                f"modes of symmetric structures come in degenerate pairs, and a mode "
                f"set that ends part-way through such a pair has an ill-defined "
                f"sensitivity. The analysis component will warn if this happens; an "
                f"even count is the safer default.", "Minimal"
            )

    def _resolve_initial_design(self) -> str:
        """The file initial_design names: as given, else beside the problem file; a run directory means its final_fields.npz."""
        given = os.path.expanduser(self.initial_design)
        candidates = [given]
        pf = getattr(self, "problem_file", None)
        if pf and not os.path.isabs(given):
            candidates.append(os.path.join(os.path.dirname(os.path.abspath(pf)), given))
        for c in candidates:
            if os.path.isdir(c):
                c = os.path.join(c, "final_fields.npz")
            if os.path.isfile(c):
                return os.path.abspath(c)
        raise FileNotFoundError(f"initial_design {self.initial_design!r} not found (tried: "
                                f"{', '.join(os.path.abspath(c) for c in candidates)}).")

    def _initial_design_phi(self) -> np.ndarray | None:
        """
        The field of the initial_design file as a float64 array clipped to
        [void_value, solid_value], or None when no initial design was asked
        for. Records the file, its hash and the field's range in
        self.initial_design_info (run_info.json provenance).
        """
        if getattr(self, "initial_design", None) is None:
            return None
        path = self._resolve_initial_design()
        if path.lower().endswith(".npy"):
            raw = np.load(path)
        else:
            with np.load(path) as z:
                if "phi" not in z.files:
                    raise ValueError(f"initial_design {path}: no 'phi' array (found {z.files}).")
                raw = z["phi"]
        phi = np.asarray(raw, dtype=np.float64)
        if phi.shape != tuple(self.domain_dims):
            raise ValueError(f"initial_design {path}: phi has shape {phi.shape}, this problem's grid is "
                             f"{tuple(self.domain_dims)}.")
        if not np.all(np.isfinite(phi)):
            raise ValueError(f"initial_design {path}: phi has non-finite values.")
        lo, hi = float(phi.min()), float(phi.max())
        if hi < self.void_value - 1e-6 or lo > self.solid_value + 1e-6:
            raise ValueError(f"initial_design {path}: phi in [{lo:g}, {hi:g}] does not overlap "
                             f"[void_value, solid_value] = [{self.void_value:g}, {self.solid_value:g}].")
        if getattr(self, "_initial_design_from_resume", False):
            # resume_state phi: keep the optimiser's own clamp overshoot (e.g. 0.9 / 2.1 of a truss run)
            phi = np.clip(phi, getattr(self, "phi_clamp_min", self.void_value), getattr(self, "phi_clamp_max", self.solid_value))
        else:
            phi = np.clip(phi, self.void_value, self.solid_value)
        phi[np.asarray(self.preserve_mask, dtype=bool)] = self.solid_value
        phi[np.asarray(self.obstacle_mask, dtype=bool)] = self.void_value
        self.initial_design_info = {"path": path, "sha256": sha256_of(path), "dtype": str(raw.dtype),
                                    "phi_min": lo, "phi_max": hi,
                                    "solid_fraction": float(np.mean(phi >= self.level_set_value))}
        self._print_out(f"Initial design read from {path} (phi in [{lo:.3f}, {hi:.3f}], solid fraction "
                        f"{self.initial_design_info['solid_fraction']:.3f}).", "Minimal")
        return phi

    def _rho(self, phi: np.ndarray | None = None) -> np.ndarray:
        """
        Computes the density field rho based on the phi field.

        This function maps the phi field to a density field 'rho' in the
        range [void_density, 1.0]. The transition is centered at
        self.level_set_value.

        Args:
            phi (np.ndarray | None): The phi field. If None, uses self.phi.

        Returns:
            np.ndarray: The computed density field rho.
        """
        if phi is None:
            phi = self.phi

        # Linear interpolation from the paper: rho(phi) = -c1(phi-2) + c2(phi-1)
        # c1 is void_density, c2 is 1.0 (normalized solid density)
        rho = -self.void_density * (phi - self.solid_value) + 1.0 * (phi - self.void_value)
        
        # Clip to ensure numerical stability keeps it within valid bounds
        rho = np.clip(rho, self.void_density, 1.0)
        return rho

    def _normalised_density(self, rho: np.ndarray) -> np.ndarray:
        """x = (rho - c1)/(1 - c1) in [0, 1]: 0 at void, 1 at solid."""
        c1 = self.void_density
        return np.clip((rho - c1) / (1.0 - c1), 0.0, 1.0)

    def _rho_model(self, phi: np.ndarray | None = None, rho: np.ndarray | None = None) -> np.ndarray:
        """
        The STIFFNESS density handed to a density backend (eq. 4.12):
        c1 + (1 - c1) x^p with x = (rho(phi) - c1)/(1 - c1). At p = 1 this is
        rho itself, and at a binary design (x in {0, 1}) it equals rho for
        every p.
        """
        if rho is None:
            rho = self._rho(phi)
        p = self.current_penalisation_power
        if p <= 1.0 + 1e-12:
            return rho
        c1 = self.void_density
        x = self._normalised_density(rho)
        return (c1 + (1.0 - c1) * x ** p).astype(rho.dtype, copy=False)

    def _geometric_stiffness_threshold(self) -> float:
        """
        The threshold of the geometric-stiffness ramp (eq. 6.10) on the scale of
        the density the backend is given (docs/theory.md §6.5(b), eq. 6.10c).

        buckling_min_density_for_geometric_stiffness, theta, is a DESIGN
        density: below it material counts as void and carries no geometric
        stiffness. A density backend analyses the penalised stiffness density
        rho_m = c1 + (1 - c1) x^p (eq. 4.12) and ramps on that, so the design
        threshold is handed over as its image under the same map,

            theta_m = c1 + (1 - c1) ((theta - c1)/(1 - c1))^p,

        which is theta itself at p = 1 (bit-identical to the earlier code) and,
        at p = 3, theta = 0.1 and c1 = 1e-3, 0.0020 (the ramp then runs from
        x = 0.099 to 0.14): grey material between x = 0.1 and 0.46, which the
        earlier code (theta handed over unchanged) stripped of all geometric
        stiffness while it still carried load, keeps it.
        buckling_ramp_on_design_density = False restores the earlier
        behaviour. Used for the evaluations of the analysed model only: the
        exported design and the threshold sweep are binary, rho in {c1, 1},
        on which theta and theta_m agree.
        """
        theta = float(self.buckling_min_density_for_geometric_stiffness)
        if not getattr(self, "buckling_ramp_on_design_density", True):
            return theta
        if getattr(self, "_analysis_field_type", "density") != "density" or not getattr(self, "_model_penalised", True):
            return theta
        p = float(getattr(self, "current_penalisation_power", 1.0) or 1.0)
        if p <= 1.0 + 1e-12:
            return theta
        c1 = float(self.void_density)
        x = min(1.0, max(0.0, (theta - c1) / (1.0 - c1)))
        return c1 + (1.0 - c1) * x ** p

    def _solid_mask(self, phi: np.ndarray) -> np.ndarray:
        """The geometry as exported: phi >= level set, plus preserve regions, minus obstacles."""
        solid = (phi >= self.level_set_value) | np.asarray(self.preserve_mask, dtype=bool)
        solid &= ~np.asarray(self.obstacle_mask, dtype=bool)
        return solid

    def _grey_fraction(self, phi: np.ndarray) -> float:
        """Fraction of design voxels with phi inside the grey band (V + 0.1(S-V), S - 0.1(S-V))."""
        D = self._gdiag()
        if D is not None:
            return float(D.current(phi)["grey"]) / self.n_design_voxels
        band = 0.1 * (self.solid_value - self.void_value)
        grey = (phi > self.void_value + band) & (phi < self.solid_value - band) & self.design_mask
        return float(np.count_nonzero(grey)) / self.n_design_voxels

    # Full (26-) connectivity: voxels sharing a face, an edge or a corner are
    # joined (docs/theory.md §5.6).
    _CONNECTIVITY_STRUCTURE = ndimage.generate_binary_structure(3, 3)

    def _connectivity_report(self, phi: np.ndarray) -> dict:
        """
        Connectivity of the geometry that will be exported (`_solid_mask`).

        The question is whether a load path has a GAP -- a layer of void at
        least one voxel thick between two pieces, which is what a beam whose
        ends are separated looks like. Voxels that share a face, an edge or a
        corner are therefore counted as joined (26-connectivity). Requiring
        face contact instead flagged every thin diagonal member of a coarse,
        low-volume design (a one-voxel staircase whose steps touch along
        edges) as broken, although the member is continuous and a load it
        carries is transferred -- and would be thickened by the stiffness
        driver if the contact were too weak.

        Returns a dict with
        - n_components: number of solid components;
        - load_path_connected: every loaded voxel lies in a component that
          also contains a supported voxel (True when there are no loads);
        - detached_fraction: fraction of the solid volume lying in
          components that do not join a load to a support, ignoring
          components made only of preserve voxels (which the optimiser
          cannot move) -- these are the floating members and the dangling
          stubs of a broken load path;
        - connected: load_path_connected and detached_fraction <=
          convergence_detached_tol.
        """
        D = self._gdiag()
        if D is not None:
            return D.connectivity(phi, self.convergence_detached_tol, bool(np.any(self.load_mask)))
        return self._connectivity_of_solid(self._solid_mask(phi))

    def _connectivity_of_solid(self, solid: np.ndarray) -> dict:
        """The report of _connectivity_report for a given solid mask (used by the threshold sweep)."""
        n_solid = int(np.count_nonzero(solid))
        if n_solid == 0:
            return dict(n_components=0, load_path_connected=not np.any(self.load_mask),
                        detached_fraction=0.0, connected=not np.any(self.load_mask))
        labels, n = ndimage.label(solid, structure=self._CONNECTIVITY_STRUCTURE)
        idx = np.arange(1, n + 1)
        touches_load = ndimage.maximum(self.load_mask & solid, labels, idx).astype(bool) if n else np.zeros(0, bool)
        touches_support = ndimage.maximum(self.support_mask & solid, labels, idx).astype(bool) if n else np.zeros(0, bool)
        has_design = ndimage.maximum(self.design_mask & solid, labels, idx).astype(bool) if n else np.zeros(0, bool)
        sizes = ndimage.sum(solid, labels, idx) if n else np.zeros(0)

        # A load that lands on no solid voxel at all (e.g. on an obstacle) is
        # not something the optimiser can connect, so it is ignored here.
        loaded_components = touches_load
        load_path_connected = bool(np.all(touches_support[loaded_components])) if np.any(loaded_components) else True
        on_path = touches_load & touches_support
        detached = (~on_path) & has_design
        detached_fraction = float(sizes[detached].sum()) / n_solid if n else 0.0
        connected = load_path_connected and detached_fraction <= self.convergence_detached_tol
        return dict(n_components=int(n), load_path_connected=load_path_connected,
                    detached_fraction=detached_fraction, connected=bool(connected))

    def _current_buckling_safety_factor(self, outer_iter: int) -> float:
        """
        The safety factor actually enforced at this outer iteration.

        Ramping matters more than it looks. A buckling constraint tends to
        push toward thicker, fewer members, which is in direct opposition to
        the volume constraint and to the hole nucleation that the PCLS
        method exists to provide. Demanding the full target the instant the
        constraint activates makes lambda3 spike, swamps the stiffness term
        and can stall the topology before it has formed. Easing the target
        in over a few outer iterations lets the structure reorganise.

        Args:
            outer_iter (int): The current outer loop iteration number.

        Returns:
            float: The safety factor to enforce this iteration.
        """
        start = self.buckling_safety_factor_initial
        target = self.buckling_safety_factor
        ramp_iters = self.buckling_safety_factor_ramp_outer_iters

        if ramp_iters <= 0 or start >= target:
            return target

        iters_since_activation = outer_iter - self.buckling_activate_after_outer_iter
        if iters_since_activation < 0:
            return start

        progress = min(1.0, (iters_since_activation + 1) / float(ramp_iters))
        return start + (target - start) * progress

    def _G3(self, buckling_results: BucklingResults, safety_factor: float) -> float:
        """
        Computes the value of the third constraint function G3. This is the
        linear buckling constraint.

            g3 = c * mu_tilde - 1/safety_factor,   c = mu_crit / mu_tilde,

        i.e. g3 = mu_crit - 1/safety_factor with mu_crit = max_i mu_i = 1 /
        Lambda_crit (buckling_constraint_measure = "critical", the default),
        or c = 1, g3 = mu_tilde - 1/safety_factor ("aggregate", the previous
        behaviour). The factor c is recomputed at every evaluation from that
        evaluation's own modes (adaptive normalisation of the aggregate) and
        treated as a constant in the sensitivity. It changes only the VALUE:
        the update direction is the aggregate's smooth gradient in both cases,
        and _compute_buckling_term divides that gradient by its own mean
        magnitude, so c cancels there. The value is what the multiplier
        update and lambda3_tilde act on, so with "critical" the constraint the
        augmented Lagrangian drives to (mu_crit = 1/SF) is the one the settled
        test checks (docs/theory.md eq. 6.27). With the aggregate it is
        Lambda_crit = SF * mu_tilde/mu_crit, up to m^(1/p) above the target;
        on the test column (a degenerate pair, 9 %) that margin did not fit in
        the volume budget, lambda3 wound up to its clamp and the design never
        settled (docs/theory.md §6.7).

        SIGN CONVENTION: g3 <= 0 means the constraint is SATISFIED (the
        structure is at or above the required buckling safety factor).
        Positive g3 means the structure buckles too early. This matches the
        standard 'g <= 0 is feasible' convention and is what makes the
        non-negative multiplier update lambda3 = max(0, lambda3 + g3/mu3)
        behave correctly: the multiplier only accumulates while the
        constraint is violated.

        Working in mu = 1/lambda rather than lambda directly keeps the
        constraint finite when the structure cannot buckle at all
        (lambda -> inf gives mu -> 0 rather than a divergence).

        Args:
            buckling_results (BucklingResults): Output of run_buckling.
            safety_factor (float): The safety factor to enforce this iteration.

        Returns:
            float: The value of G3.
        """
        mu_target = 1.0 / safety_factor

        status = buckling_results.status
        if status == BucklingResults.STATUS_NO_COMPRESSIVE_MODE:
            # The eigensolver found no positive mu at all: there is no load
            # factor at which this design buckles under this load case (a
            # tension-dominated state). That is a genuine result, so the
            # constraint is reported as comfortably satisfied rather than
            # letting a zero mu_tilde masquerade as a real value.
            self._print_out("Computed G3 (Buckling Constraint): no compressive mode exists, "
                            "constraint satisfied.", "Verbose")
            return -mu_target
        if status in BucklingResults.UNKNOWN_STATUSES:
            # Candidate modes existed but the spurious-mode filter rejected
            # all of them. Nothing is known about the buckling load of this
            # design, and it must NOT be reported as safe: that is how a run
            # used to finish logging lambda_crit = inf for a design that
            # buckled below the applied load (docs/theory.md §6.7). NaN tells
            # the caller to keep the last valid evaluation instead.
            self._print_out("Computed G3 (Buckling Constraint): UNKNOWN (" +
                            ("the eigensolver converged nothing, even after the retry with larger limits)."
                             if status == BucklingResults.STATUS_NOT_CONVERGED else
                             "every candidate mode was rejected by the spurious-mode filter).") , "Verbose")
            return float('nan')

        if self.buckling_constraint_measure == "critical":
            mu_measure = float(np.max(buckling_results.mu_values))
        else:
            mu_measure = float(buckling_results.mu_tilde)
        g3 = mu_measure - mu_target
        self._print_out(f"Computed G3 (Buckling Constraint, {self.buckling_constraint_measure}): {g3:.6f} "
                        f"(required lambda >= {safety_factor:.3f}, "
                        f"aggregated lambda = {buckling_results.aggregated_load_factor:.4f}, "
                        f"true critical lambda = {buckling_results.critical_load_factor:.4f})",
                        "Verbose")
        return g3

    def _log_buckling(self, buckling_results: BucklingResults, safety_factor: float, g3: float):
        """
        Appends a per-evaluation record of the buckling analysis to
        buckling.csv, and prints a summary.

        Kept separate from variables.csv because buckling is evaluated on a
        different cadence (buckling_evaluate_every) and carries per-mode
        data that does not fit the one-row-per-iteration layout.

        `status` distinguishes an evaluation with retained modes ("OK") from
        one with no compressive mode at all ("NO_COMPRESSIVE_MODE", load
        factor inf, genuinely safe) and one whose candidates were all
        rejected by the spurious-mode filter ("ALL_FILTERED", load factor
        and g3 written as nan: unknown, NOT safe). `n_candidates` is the
        number of positive-eigenvalue modes found before filtering.
        """
        file_path = os.path.join(self.optimisation_name, "buckling.csv")
        if not os.path.exists(file_path):
            with open(file_path, 'w') as f:
                f.write("iteration,safety_factor_enforced,status,n_candidates,n_modes_retained,"
                        "critical_load_factor,aggregated_load_factor,mu_tilde,g3,"
                        "lambda3,mu3,min_subspace_continuity,mu_values\n")

        # Subspace continuity, not a pairwise MAC: the latter reads low on any
        # symmetric structure purely because eigenvectors rotate freely inside
        # a degenerate cluster, which is not a tracking failure.
        min_cont = 'N/A'
        if buckling_results.subspace_continuity is not None and buckling_results.subspace_continuity.size:
            min_cont = f"{buckling_results.min_subspace_continuity:.4f}"

        mu_list = ";".join(f"{v:.6e}" for v in buckling_results.mu_values)

        def _fmt(x):
            if np.isnan(x):
                return "nan"
            return "inf" if not np.isfinite(x) else f"{x:.6f}"

        status = buckling_results.status
        try:
            with open(file_path, 'a') as f:
                f.write(f"{self.total_iterations},"
                        f"{safety_factor:.6f},"
                        f"{status},"
                        f"{buckling_results.n_candidates},"
                        f"{buckling_results.mu_values.size},"
                        f"{_fmt(buckling_results.critical_load_factor)},"
                        f"{_fmt(buckling_results.aggregated_load_factor)},"
                        f"{buckling_results.mu_tilde:.6e},"
                        f"{'nan' if np.isnan(g3) else f'{g3:.6e}'},"
                        f"{self.lambda3:.6f},"
                        f"{self.mu3:.6e},"
                        f"{min_cont},"
                        f"{mu_list}\n")
        except Exception as e:
            self._print_out(f"Warning: failed to write buckling.csv. Reason: {e}", "Minimal")

        if status == BucklingResults.STATUS_NO_COMPRESSIVE_MODE:
            self._print_out(
                "  Buckling: no compressive mode exists for this design and load case "
                "(no positive eigenvalue), so it cannot buckle: constraint satisfied.", "Normal")
            return
        if status in BucklingResults.UNKNOWN_STATUSES:
            if status == BucklingResults.STATUS_NOT_CONVERGED:
                self._print_out(
                    "  Buckling: UNKNOWN this evaluation - the eigensolver converged no eigenpair, "
                    "even after a retry with larger ARPACK limits. The last valid evaluation is "
                    "kept for the update, and this one cannot count towards convergence.", "Normal")
                return
            self._print_out(
                f"  Buckling: UNKNOWN this evaluation - all {buckling_results.n_candidates} "
                f"candidate mode(s) were rejected by the spurious-mode filter. The last valid "
                f"evaluation is kept for the update, and this one cannot count towards "
                f"convergence. If this persists, lower buckling_min_energy_fraction.", "Normal")
            return

        verdict = "SATISFIED" if g3 <= 0.0 else "VIOLATED"
        self._print_out(
            f"  Buckling: lambda_crit={_fmt(buckling_results.critical_load_factor)} "
            f"(aggregated {_fmt(buckling_results.aggregated_load_factor)}), "
            f"required >= {safety_factor:.3f} -> {verdict}", "Normal")

    def _compute_compliance(self, F_u: np.ndarray, rho: np.ndarray,
                            analysis_component=None) -> float:
        """
        Returns the structural compliance c = f . u in joules.

        For a linear elastic structure the compliance is exactly twice the
        total stored strain energy, and the inner product of the assembled
        load vector with the displacement solution is the cheapest exact way
        to obtain it. Every analysis backend already holds both vectors at the
        end of its solve, so each one now publishes the product on itself as
        `.compliance`; that exact value is used whenever it is available.

        THE FALLBACK BELOW IS NOT THE COMPLIANCE. sum(SED * rho) was the
        original quantity logged here, and it is retained only so that a
        backend which does not publish an exact value still produces a
        monotonic progress signal. For the ersatz backends (Regular,
        ConjugateGradient) the returned SED is already rho-scaled, so the
        product double-counts the density and carries units of Pa rather than
        J. It must not be compared between backends or against an external
        reference value.

        Args:
            F_u (np.ndarray): The strain energy density in each voxel.
            rho (np.ndarray): The density field.
            analysis_component: The component that produced F_u, if available.

        Returns:
            float: The compliance of the structure (J), or the legacy proxy.
        """
        exact = getattr(analysis_component, "compliance", None)
        if exact is not None and np.isfinite(exact):
            self.compliance_is_exact = True
            self._print_out(f"Computed Compliance (exact, f.u): {float(exact):.6e} J", "Verbose")
            return float(exact)

        self.compliance_is_exact = False
        compliance = float(np.sum(F_u * rho))
        self._print_out(f"Computed Compliance (LEGACY PROXY sum(SED*rho), not in joules): "
                        f"{compliance:.6f}", "Verbose")
        return compliance

    # Columns appended to variables.csv after elapsed_time (docs/theory.md
    # §8.4.1). Each is the value of the attribute of the same name, set in
    # _inner_loop (and for the initial row in _outer_loop); 'N/A' when unset.
    _EXTRA_VARIABLE_COLUMNS = (
        "outer_iteration", "inner_iteration",
        "exported_volume_fraction", "void_film_fraction", "non_discreteness", "undecided_fraction",
        "max_net_force", "mean_abs_dphi", "relative_compliance_change",
        "lambda1_tilde", "stiffness_normalisation_scale",
        "euler_characteristic", "n_tunnels", "n_cavities", "interface_faces",
        "nucleations_interior", "nucleations_surface", "nucleated_voxels",
        "settle_failures",
        "lambda3_tilde", "safety_factor_enforced", "aggregated_load_factor",
        "buckling_time", "solver_iterations", "solver_converged", "solver_final_residual",
        "warm_start_attempted", "warm_start_bailed_to_cold",
        "diffusion_term_min", "diffusion_term_mean", "diffusion_term_max",
        "exported_compliance", "exported_analysis_time",
        "analysis_valid", "floating_volume_fraction", "floating_load_fraction",
        "step_scale", "step_rejected",
        "mesh_elements", "min_tet_quality", "degenerate_tets_removed", "mc_vertices_clamped",
        "compliance_trend", "net_phase_switches",
    )

    def _extra_variable_values(self) -> list:
        """The values of _EXTRA_VARIABLE_COLUMNS for the current iteration, as strings."""
        out = []
        for name in self._EXTRA_VARIABLE_COLUMNS:
            v = getattr(self, name, None)
            if v is None:
                out.append("N/A")
            elif isinstance(v, (bool, np.bool_)):
                out.append(str(int(v)))
            elif isinstance(v, (int, np.integer)):
                out.append(str(int(v)))
            elif isinstance(v, (float, np.floating)):
                v = float(v)
                out.append("nan" if np.isnan(v) else ("inf" if np.isinf(v) and v > 0 else
                           ("-inf" if np.isinf(v) else f"{v:.8g}")))
            else:
                out.append(str(v).replace(",", ";"))
        return out

    def _pcls_cell(self, column: str, value) -> str:
        """
        One of the variables.csv cells that carry the augmented-Lagrangian state (lambda1, the lambda2 field
        min/mean/max, mu1, mu2): the rounded value, 'N/A' for a column the method lists in `_variables_na()` (it has
        no such quantity), and 'nan' when the method lists it as applicable but has not set it.
        """
        if column in self._variables_na():
            return "N/A"
        try:
            v = value()
        except AttributeError:
            return "nan"
        return "nan" if v is None else f"{round(v, 6)}"      # f-string, not str(): a float32 prints in full

    def _save_variables(self, phi: np.ndarray, total_iter_time: float | None, analysis_time: float | None, overwrite=False):
        """
        Saves key optimisation variables to files for post-processing. It saves the data about 'phi_new' as well as the terms used to compute it for each iteration.
        """
        file_path = os.path.join(self.optimisation_name, "variables.csv")
        _conn = bool(getattr(self, "connectivity", None)) and getattr(self, "_conn_fresh", True)
        if not os.path.exists(file_path) or overwrite:
            with open(file_path, 'w') as f:
                # Columns up to analysis_time are unchanged from earlier
                # revisions; the diagnostics of docs/theory.md §5.5-§5.6 are
                # appended after them so older readers keep working.
                f.write("iteration,G1,G2_min,G2_mean,G2_max,phi_min,phi_mean,phi_max,sed_min,sed_mean,sed_max,mesh_volume_fraction,stiffness_term_min,stiffness_term_mean,stiffness_term_max,volume_term_min,volume_term_mean,volume_term_max,piecewise_term_min,piecewise_term_mean,piecewise_term_max,buckling_term_min,buckling_term_mean,buckling_term_max,field_volume_fraction,compliance,lambda1,lambda2_min,lambda2_mean,lambda2_max,mu1,mu2,G3,lambda3,mu3,critical_load_factor,n_buckling_modes,total_iteration_time,analysis_time,"
                        "penalisation_power,grey_fraction,phase_switches,n_components,load_path_connected,detached_fraction,settled_count,elapsed_time,"
                        + ",".join(self._EXTRA_VARIABLE_COLUMNS) + "\n")
        
        try:
            with open(file_path, 'a') as f:
                f.write(f"{self.total_iterations},"
                        f"{round(self.g1, 6) if hasattr(self, 'g1') and self.g1 is not None else 'N/A'},"
                        f"{round(self.g2.min(), 6) if hasattr(self, 'g2') and self.g2 is not None else 'N/A'},"
                        f"{round(self.g2.mean(), 6) if hasattr(self, 'g2') and self.g2 is not None else 'N/A'},"
                        f"{round(self.g2.max(), 6) if hasattr(self, 'g2') and self.g2 is not None else 'N/A'},"
                        f"{round(phi.min(), 6)},"
                        f"{round(phi.mean(), 6)},"
                        f"{round(phi.max(), 6)},"
                        f"{round(self.F_u.min(), 6) if hasattr(self, 'F_u') and self.F_u is not None else 'N/A'},"
                        f"{round(self.F_u.mean(), 6) if hasattr(self, 'F_u') and self.F_u is not None else 'N/A'},"
                        f"{round(self.F_u.max(), 6) if hasattr(self, 'F_u') and self.F_u is not None else 'N/A'},"
                        f"{round(self.mesh_volume_fraction, 6) if hasattr(self, 'mesh_volume_fraction') and self.mesh_volume_fraction is not None else 'N/A'},"
                        f"{round(self.stiffness_term.min(), 6) if hasattr(self, 'stiffness_term') and self.stiffness_term is not None else 'N/A'},"
                        f"{round(self.stiffness_term.mean(), 6) if hasattr(self, 'stiffness_term') and self.stiffness_term is not None else 'N/A'},"
                        f"{round(self.stiffness_term.max(), 6) if hasattr(self, 'stiffness_term') and self.stiffness_term is not None else 'N/A'},"
                        f"{round(self.volume_term.min(), 6) if hasattr(self, 'volume_term') and self.volume_term is not None else 'N/A'},"
                        f"{round(self.volume_term.mean(), 6) if hasattr(self, 'volume_term') and self.volume_term is not None else 'N/A'},"
                        f"{round(self.volume_term.max(), 6) if hasattr(self, 'volume_term') and self.volume_term is not None else 'N/A'},"
                        f"{round(self.piecewise_term.min(), 6) if hasattr(self, 'piecewise_term') and self.piecewise_term is not None else 'N/A'},"
                        f"{round(self.piecewise_term.mean(), 6) if hasattr(self, 'piecewise_term') and self.piecewise_term is not None else 'N/A'},"
                        f"{round(self.piecewise_term.max(), 6) if hasattr(self, 'piecewise_term') and self.piecewise_term is not None else 'N/A'},"
                        f"{round(self.applied_buckling_term.min(), 6) if hasattr(self, 'applied_buckling_term') and self.applied_buckling_term is not None else 'N/A'},"
                        f"{round(self.applied_buckling_term.mean(), 6) if hasattr(self, 'applied_buckling_term') and self.applied_buckling_term is not None else 'N/A'},"
                        f"{round(self.applied_buckling_term.max(), 6) if hasattr(self, 'applied_buckling_term') and self.applied_buckling_term is not None else 'N/A'},"
                        f"{round(self.volume_fraction, 6) if hasattr(self, 'volume_fraction') and self.volume_fraction is not None else 'N/A'},"
                        f"{round(self.compliance, 6) if hasattr(self, 'compliance') and self.compliance is not None else 'N/A'},"
                        f"{self._pcls_cell('lambda1', lambda: self.lambda1)},"
                        f"{self._pcls_cell('lambda2_min', lambda: self.lambda2.min())},"
                        f"{self._pcls_cell('lambda2_mean', lambda: self.lambda2.mean())},"
                        f"{self._pcls_cell('lambda2_max', lambda: self.lambda2.max())},"
                        f"{self._pcls_cell('mu1', lambda: self.mu1)},"
                        f"{self._pcls_cell('mu2', lambda: self.mu2)},"
                        f"{round(self.g3, 8) if hasattr(self, 'g3') and self.g3 is not None else 'N/A'},"
                        f"{round(self.lambda3, 6) if hasattr(self, 'lambda3') else 'N/A'},"
                        f"{round(self.mu3, 8) if hasattr(self, 'mu3') else 'N/A'},"
                        f"{self._fmt_load_factor(self.critical_load_factor) if hasattr(self, 'critical_load_factor') else 'N/A'},"
                        f"{self.n_buckling_modes if hasattr(self, 'n_buckling_modes') else 'N/A'},"
                        f"{round(total_iter_time, 6) if total_iter_time is not None else 'N/A'},"
                        f"{round(analysis_time, 6) if analysis_time is not None else 'N/A'},"
                        f"{round(self.current_penalisation_power, 6)},"
                        f"{round(self.grey_fraction, 6) if getattr(self, 'grey_fraction', None) is not None else 'N/A'},"
                        f"{self.phase_switches if getattr(self, 'phase_switches', None) is not None else 'N/A'},"
                        f"{self.connectivity['n_components'] if _conn else 'N/A'},"
                        f"{int(self.connectivity['load_path_connected']) if _conn else 'N/A'},"
                        f"{round(self.connectivity['detached_fraction'], 6) if _conn else 'N/A'},"
                        f"{getattr(self, 'settled_count', 0)},"
                        f"{round(self._elapsed(), 3)},"
                        + ",".join(self._extra_variable_values()) + "\n"
                )
            f.close()
        except Exception as e:
            raise RuntimeError(f"Failed to save optimisation variables: {e}")
        
        self._print_out(f"Saved variables to csv for iteration {self.total_iterations}", "Verbose")

    @staticmethod
    def _fmt_load_factor(x: float) -> str:
        """inf = no compressive mode (safe); nan = unknown (all modes filtered)."""
        if x is None:
            return 'N/A'
        if np.isnan(x):
            return 'nan'
        return 'inf' if not np.isfinite(x) else str(round(float(x), 6))

    def _elapsed(self) -> float:
        """Seconds since run_optimisation started (0 before it has)."""
        start = getattr(self, "_run_start_time", None)
        return 0.0 if start is None else timer() - start

    def _record_failure(self, reason: str, message: str, trace: str, failed_phi: np.ndarray | None = None):
        """
        Records a failure that ends the run (docs/theory.md §8.4.4): `reason`
        ("ANALYSIS_FAILED" or "ERROR") and `message` go to self.failure (and
        run_info.json), the traceback to error.txt in the run directory, and
        the design the failure happened on to failed_phi.npz, so the failing
        analysis can be reproduced.
        """
        if reason == "ANALYSIS_FAILED":
            self.analysis_failed = True
        self.failure = {"reason": reason, "message": message,
                        "iteration": int(getattr(self, "total_iterations", 0)),
                        "elapsed_s": self._elapsed()}
        self._print_out(f"ERROR ({reason}): {message}", "Minimal")
        try:
            with open(os.path.join(self.optimisation_name, "error.txt"), "a") as f:
                f.write(f"{reason} at iteration {self.failure['iteration']} "
                        f"({self.failure['elapsed_s']:.1f} s into the run)\n{message}\n\n{trace}\n")
        except Exception:
            pass
        if failed_phi is not None:
            try:
                np.savez_compressed(os.path.join(self.optimisation_name, "failed_phi.npz"),
                                    phi=np.asarray(failed_phi, dtype=np.float32),
                                    iteration=self.failure["iteration"])
            except Exception:
                pass

    def _phase_takes_stiffness_density(self, component) -> bool:
        """Whether a phase-field backend's run_analysis accepts stiffness_density (Conformal does)."""
        cached = getattr(self, "_phase_stiffness_density_ok", None)
        if cached is None:
            try:
                cached = "stiffness_density" in inspect.signature(component.run_analysis).parameters
            except (TypeError, ValueError):
                cached = False
            self._phase_stiffness_density_ok = cached
        return cached

    def _record_iteration_diagnostics(self, phi_new: np.ndarray, phi_old: np.ndarray | None,
                                      outer_iter: int, inner_iter: int | None, force: bool = False):
        """
        Sets the per-iteration diagnostics of docs/theory.md §8.4.1 on self
        (they are written to variables.csv by _save_variables) and records
        nucleation events and phi snapshots through self._diagnostics.

        Called once per iteration with the field the iteration produced
        (phi_new) and the one it started from (phi_old), and once for the
        initial field with phi_old None.
        """
        if self._gdiag() is not None:
            return self._record_iteration_diagnostics_gpu(phi_new, phi_old, outer_iter, inner_iter, force)
        self.outer_iteration = int(outer_iter)
        self.inner_iteration = None if inner_iter is None else int(inner_iter)
        solid = self._solid_mask(phi_new)
        n = phi_new.size
        # Needed by the stopping tests (_exported_volume_ok): every iteration.
        self.exported_volume_fraction = float(np.count_nonzero(solid)) / n
        every = getattr(self, "diagnostics_every", 1)
        if not (force or phi_old is None or every <= 1 or self.total_iterations % every == 0):
            # Skipped iteration (diagnostics_every > 1): the rest are blank in variables.csv.
            self._diag_fresh = False
            self._diag_pending = (phi_new, phi_old, outer_iter, inner_iter, self.total_iterations)
            for k in ("void_film_fraction", "non_discreteness", "undecided_fraction", "euler_characteristic",
                      "n_tunnels", "n_cavities", "interface_faces", "nucleations_interior",
                      "nucleations_surface", "nucleated_voxels"):
                setattr(self, k, None)
            if phi_old is not None:
                step = np.abs(phi_new - phi_old)
                step = step[self.design_mask] if np.any(self.design_mask) else step
                dt_step = getattr(self, "_dt_step", None)
                self.max_net_force = float(step.max()) / (self.dt if dt_step is None else dt_step) if step.size else 0.0
                self.mean_abs_dphi = float(step.mean()) if step.size else 0.0
            self._diagnostics.snapshot(self.total_iterations, phi_new)
            return
        self._diag_fresh = True
        self._diag_pending = None
        rho_new = self._rho(phi=phi_new)
        design = self.design_mask
        # The part of the volume constraint that is NOT exported: density
        # above c1 in design voxels below the level set (docs/theory.md §5.5).
        void_design = design & ~solid
        self.void_film_fraction = float(np.sum(rho_new[void_design] - self.void_density)) / n
        self.non_discreteness = non_discreteness(rho_new, design)
        band = 0.25 * (self.solid_value - self.void_value)
        self.undecided_fraction = band_fraction(phi_new, self.void_value + band,
                                                      self.solid_value - band, design)
        if phi_old is not None:
            step = np.abs(phi_new - phi_old)
            step = step[design] if np.any(design) else step
            dt_step = getattr(self, "_dt_step", None)       # PCLS's step; a baseline sets 1 (max|dphi| per iteration)
            self.max_net_force = float(step.max()) / (self.dt if dt_step is None else dt_step) if step.size else 0.0
            self.mean_abs_dphi = float(step.mean()) if step.size else 0.0
        conn = getattr(self, "connectivity", None)
        n_components = conn["n_components"] if (conn is not None and phi_old is not None) else None
        fields = None
        if phi_old is not None:
            fields = {"stiffness_term": getattr(self, "stiffness_term", None),
                      "volume_term": getattr(self, "volume_term", None),
                      "piecewise_term": getattr(self, "piecewise_term", None),
                      "lambda2": getattr(self, "lambda2", None)}
        try:
            d = self._diagnostics.update(
                iteration=self.total_iterations, outer_iteration=outer_iter,
                inner_iteration=-1 if inner_iter is None else inner_iter,
                phi=phi_new, solid=solid, void_for_nucleation=phi_new < self.level_set_value,
                n_components=n_components, fields=fields)
            for k, v in d.items():
                setattr(self, k, v)
        except Exception as e:
            self._print_out(f"Warning: topology/nucleation diagnostics failed this iteration. Error: {e}",
                            "Minimal")

    # ------------------------------------------------------------------
    # GPU update and diagnostics (the former gpu_update.py sections, docs/theory.md §18)
    # ------------------------------------------------------------------

    _gpu_update_kind = None     # "simp" | "pcls" on the classes that have a GPU update; None: host code only

    def _gpu_kind(self):
        """The kind of GPU update this run uses ("simp" | "pcls"), or None for the host code. Decided once per run by
        gpu_analysis.update_engine_choice (FORCE_CLASSIC, optimiser.update_engine, CUDA device, memory model)."""
        sel = getattr(self, "_update_engine_sel", None)
        if sel is None:
            sel = "cpu"
            if self._gpu_update_kind is not None and self._analysis_field_type == "density":
                import gpu_analysis
                sel, _ = gpu_analysis.update_engine_choice(self._gpu_update_kind, int(np.prod(self.domain_dims)),
                                                           unsupported=self._gpu_update_unsupported())
            self._update_engine_sel = sel
        return self._gpu_update_kind if sel == "gpu" else None

    def _gpu_update_unsupported(self):
        """The reason this run's options have no GPU update (the host code then runs, logged and recorded by
        gpu_analysis.update_engine_choice), or None. Overridden by optimisers with such options."""
        return None

    def _gpu_failed(self, what: str, err: Exception):
        """A GPU update component could not be set up: the run continues on the host code (recorded)."""
        import gpu_analysis
        gpu_analysis.record_fallback("update", what, "cpu", err)
        self._update_engine_sel = "cpu"
        self._gdiag_obj = None
        self._print_out(f"WARNING: GPU update engine ({what}) failed: {type(err).__name__}: {err}. "
                        f"Continuing on the host code.", "Minimal")

    def _gdiag(self):
        """The device diagnostics (DiagGPU) or None on the host path."""
        if self._gpu_kind() is None:
            return None
        d = getattr(self, "_gdiag_obj", None)
        if d is None:
            try:
                d = DiagGPU(self.domain_dims, self.design_mask, self.preserve_mask, self.obstacle_mask,
                                       self.load_mask, self.support_mask, self.void_density, self.solid_value,
                                       self.void_value, self.level_set_value)
            except Exception as e:                      # noqa: BLE001  (e.g. device memory): fall back to the host
                self._gpu_failed("diagnostics", e)
                return None
            self._gdiag_obj = d
        return d

    def _gpu_materialise(self):
        """Replaces the lazily downloaded device fields by host arrays (before the end-of-run writers, which test for
        ndarray). A no-op on the host path."""
        for name in ("stiffness_term", "volume_term", "piecewise_term", "applied_buckling_term", "diffusion_term",
                     "net_push", "F0", "dc", "dv", "x", "x_filtered", "_dproj", "buckling_term"):
            v = getattr(self, name, None)
            if v is not None and not isinstance(v, np.ndarray) and hasattr(v, "host"):
                setattr(self, name, np.array(v.host()))

    def _phase_switches(self, phi_new: np.ndarray, phi_old: np.ndarray) -> int:
        """Design voxels whose phase (phi >= level set) differs between phi_old and phi_new."""
        D = self._gdiag()
        if D is not None:
            return int(D.scan(phi_new, phi_old)["switches"])
        ls = self.level_set_value
        return int(np.count_nonzero(((phi_new >= ls) != (phi_old >= ls)) & self.design_mask))

    def _record_iteration_diagnostics_gpu(self, phi_new, phi_old, outer_iter, inner_iter, force=False):
        """_record_iteration_diagnostics on the device: the same quantities and the same diagnostics_every cadence
        (DiagGPU: one scan of the field gives the volume, void film, discreteness, band and step
        statistics; topology and nucleation counts only on the iterations that run the full diagnostics)."""
        self.outer_iteration = int(outer_iter)
        self.inner_iteration = None if inner_iter is None else int(inner_iter)
        D = self._gdiag()
        sc = D.scan(phi_new, phi_old)
        n = phi_new.size
        self.exported_volume_fraction = float(sc["n_solid"]) / n

        def step_stats():
            if phi_old is not None:
                dt_step = getattr(self, "_dt_step", None)
                cnt = sc["step_count"]
                self.max_net_force = sc["step_max"] / (self.dt if dt_step is None else dt_step) if cnt else 0.0
                self.mean_abs_dphi = sc["step_sum"] / cnt if cnt else 0.0

        every = getattr(self, "diagnostics_every", 1)
        if not (force or phi_old is None or every <= 1 or self.total_iterations % every == 0):
            self._diag_fresh = False
            self._diag_pending = (phi_new, phi_old, outer_iter, inner_iter, self.total_iterations)
            for k in ("void_film_fraction", "non_discreteness", "undecided_fraction", "euler_characteristic",
                      "n_tunnels", "n_cavities", "interface_faces", "nucleations_interior",
                      "nucleations_surface", "nucleated_voxels"):
                setattr(self, k, None)
            step_stats()
            self._diagnostics.snapshot(self.total_iterations, phi_new)
            return
        self._diag_fresh = True
        self._diag_pending = None
        n_design = int(np.count_nonzero(self.design_mask))
        self.void_film_fraction = float(sc["void_film"]) / n
        self.non_discreteness = float(4.0 * sc["nd_sum"] / n_design) if n_design else 0.0
        self.undecided_fraction = float(sc["undecided"]) / max(1, n_design)
        step_stats()
        conn = getattr(self, "connectivity", None)
        n_components = conn["n_components"] if (conn is not None and phi_old is not None) else None
        fields = None
        if phi_old is not None:
            fields = {"stiffness_term": getattr(self, "stiffness_term", None),
                      "volume_term": getattr(self, "volume_term", None),
                      "piecewise_term": getattr(self, "piecewise_term", None),
                      "lambda2": getattr(self, "lambda2", None)}
        try:
            d = self._diagnostics.update_gpu(
                D, iteration=self.total_iterations, outer_iteration=outer_iter,
                inner_iteration=-1 if inner_iter is None else inner_iter, phi=phi_new,
                n_components=n_components, fields=fields)
            for k, v in d.items():
                setattr(self, k, v)
        except Exception as e:
            self._print_out(f"Warning: topology/nucleation diagnostics failed this iteration. Error: {e}",
                            "Minimal")

    def _diag_due(self) -> bool:
        """True when this iteration is on the diagnostics_every cadence (iteration 0 always is)."""
        every = getattr(self, "diagnostics_every", 1)
        return every <= 1 or self.total_iterations % every == 0

    def _refresh_connectivity(self, phi_new):
        """Connectivity report of phi_new; run on the cadence, and when a stopping test needs it."""
        self.connectivity = self._connectivity_report(phi_new)
        self._conn_fresh = True

    def _flush_diagnostics(self):
        """Full diagnostics for the last iteration if it was skipped (diagnostics_every > 1), so the
        final-design summary and connectivity warning see the field the run ended on."""
        p = getattr(self, "_diag_pending", None)
        if p is not None:
            if not getattr(self, "_conn_fresh", True):
                self._refresh_connectivity(p[0])
            # total_iterations has moved on since the skipped record: restore the iteration it belonged to
            # (the snapshot file name, nucleation-event iteration and topology-change iteration use it).
            now = self.total_iterations
            self.total_iterations = p[4]
            try:
                self._record_iteration_diagnostics(*p[:4], force=True)
            finally:
                self.total_iterations = now

    def _stop_file_requested(self) -> bool:
        """
        True when a file named STOP exists in the run directory
        (<optimisation_name>/STOP): a graceful stop requested from outside, e.g.
        `touch <run>/STOP` on the cluster. The run then ends as on its time
        limit -- the final design is saved, analysed and post-processed --
        with termination_reason "STOP_FILE" (docs/theory.md §5.7). A scheduler
        cancel (scancel) instead kills the process before any final output.
        """
        try:
            requested = os.path.exists(os.path.join(str(self.optimisation_name), "STOP"))
        except Exception:
            return False
        if requested and not getattr(self, "stop_requested", False):
            self._print_out(f"STOP file found in {self.optimisation_name}/: stopping after "
                            f"{getattr(self, 'total_iterations', 0)} iterations and saving the final design.",
                            "Minimal")
        if requested:
            self.stop_requested = True
        return requested

    def _time_limit_would_be_exceeded(self) -> bool:
        """
        True when another iteration would not finish inside `time_limit_hours`.

        The next iteration is assumed to cost as much as the most expensive of
        the recent ones (tracked in `_iteration_times`), and twice that again
        is reserved for what run_optimisation does after the loops return: the
        guaranteed final save (one more analysis, plus the surface export) and
        the analysis of the exported geometry.
        Stopping on this estimate rather than on the limit itself means a run
        given, say, 23.5 h inside a 24 h scheduler allocation ends cleanly with
        its final design saved instead of being killed mid-iteration.

        With buckling (or buckling_report_exported), twice the longest
        eigensolve so far is reserved as well: the next iteration may be an
        evaluation, and the exported design's load factor is always computed
        at the end. The window of five recent iterations used to miss the
        eigensolve whenever the idle cadence (one evaluation in 10 iterations)
        or report_exported (none in the loop) left it out, and on the
        production column an eigensolve is 20 times an iteration (162 s
        against about 7 s; optimiser review, 2026-10).
        """
        if self.time_limit_hours is None:
            return False
        recent = self._iteration_times[-5:]
        estimate = max(recent) if recent else 0.0
        reserve = 3.0 * estimate
        if getattr(self, "buckling_enabled", False) or getattr(self, "buckling_report_exported", False):
            reserve += 2.0 * float(getattr(self, "_max_buckling_time", 0.0) or 0.0)
        return self._elapsed() + reserve >= 3600.0 * self.time_limit_hours

    def _buckling_converged(self, outer_iter: int) -> bool:
        """
        The buckling part of the settled test (docs/theory.md §5.5): True when
        buckling is off; otherwise the constraint must have been evaluated
        VALIDLY (not "unknown"), against the full, un-ramped safety factor (the
        ramp itself need not have finished unless
        buckling_settle_requires_full_ramp: a design that already meets the
        full target has nothing left to wait for), the
        most recent evaluation must not have been "unknown", and the TRUE
        critical load factor must meet the target to a relative tolerance,
        Lambda_crit >= SF/(1 + tol) (docs/theory.md eq. 6.27). This is the
        quantity the default constraint measure drives to SF, so a design the
        constraint is satisfied on can settle. The p-norm aggregate 1/mu_tilde
        sits below Lambda_crit by up to m^(1/p) and is not used here, whatever
        the measure. (The older test g3 <= tol was also absolute in mu: a 2 %
        margin at SF = 2, 8 % at 8.)
        """
        if not self.buckling_enabled:
            return True
        if not getattr(self, 'buckling_valid_evaluated', False):
            return False
        if not getattr(self, 'last_buckling_evaluation_valid', False):
            return False
        if (self.buckling_settle_requires_full_ramp
                and self._current_buckling_safety_factor(outer_iter) < self.buckling_safety_factor):
            return False
        return self._buckling_target_met()

    def _buckling_target_met(self) -> bool:
        """Lambda_crit of the last valid evaluation >= SF/(1 + tol); inf (no compressive mode) passes."""
        lam = getattr(self, 'last_valid_critical_load_factor', float('nan'))
        if np.isnan(lam):
            return False
        return bool(lam >= self.buckling_safety_factor / (1.0 + self.outer_loop_convergence_threshold))

    def _volume_multiplier(self, lambda1: float, g1: float) -> float:
        """
        The stored volume multiplier lambda1 that an update with volume error g1
        should build on: lambda1 itself, or 0 when it has wound up the wrong way --
        g1 of the opposite sign (the volume already past the target, lambda1 still
        driving it further) by more than volume_multiplier_reset_tol (None:
        outer_loop_convergence_threshold), in G1's units. Used for lambda1_tilde =
        this + g1/mu1 at every iteration and for the outer update lambda1 <- this +
        g1/mu1 (PCLS and level set).
        """
        tol = getattr(self, "volume_multiplier_reset_tol", None)
        if tol is None:
            tol = float(self.outer_loop_convergence_threshold)
        if tol > 0.0 and lambda1 * g1 < 0.0 and abs(g1) > tol:
            return 0.0
        return lambda1

    def _exported_volume_ok(self) -> bool:
        """
        True when the exported geometry's volume fraction is within
        convergence_exported_volume_tol (relative) of the target V0, or the
        test is disabled (tolerance <= 0). Uses self.exported_volume_fraction,
        set by _record_iteration_diagnostics for the field just produced.
        """
        tol = self.convergence_exported_volume_tol
        if tol <= 0.0 or self.target_volume <= 0.0:
            return True
        exported = getattr(self, "exported_volume_fraction", None)
        if exported is None or not np.isfinite(exported):
            return True
        return abs(exported - self.target_volume) <= tol * self.target_volume

    def _reset_settle_windows(self):
        """
        Empty windows of the windowed settle rates (eq. 5.8c). The compliance
        slope needs at least two values, so its window holds max(w, 2) of them:
        with convergence_window = 1 a window of length 1 could never reach the
        two values the slope test waits for, and a phase backend (or the
        level set) could never converge (optimiser review, 2026-10). For w >= 2
        nothing changes.
        """
        self._settle_compliance_window = deque(maxlen=max(self.convergence_window, 2))
        self._settle_solid_window = deque(maxlen=self.convergence_window + 1)

    def _windowed_settle_rates(self, phi_new: np.ndarray):
        """
        The compliance and phase-switch tests of the settled criterion for a
        remeshing (phase) backend (docs/theory.md §5.5, eq. 5.8c).

        TetGen builds a new mesh every iteration, so `Conformal`'s compliance
        changes by a few 1e-4 from one iteration to the next even when the
        design does not move, and a boundary voxel can cross the level set and
        come back. Both tests therefore measure a rate over the last w =
        convergence_window iterations instead of a one-iteration difference:

        - compliance: the least-squares slope of c over the window, relative
          to its mean, must not exceed convergence_compliance_tol per
          iteration (w valid analyses in a row are needed; a rejected or
          failed one empties the window);
        - switches: the design voxels whose phase differs between now and w
          iterations ago -- flips that are undone do not count -- must not
          exceed convergence_switch_tol x n_design x w.

        Sets compliance_trend and net_phase_switches (None until the windows
        are full) and returns (compliance_settled, switches_settled).
        """
        w = self.convergence_window
        if self.analysis_valid and self.compliance is not None and np.isfinite(self.compliance):
            self._settle_compliance_window.append(float(self.compliance))
        else:
            self._settle_compliance_window.clear()
        self._settle_solid_window.append((phi_new >= self.level_set_value) & self.design_mask)
        compliance_settled = switches_settled = False
        self.compliance_trend = self.net_phase_switches = None
        if len(self._settle_compliance_window) >= max(w, 2):
            y = np.asarray(self._settle_compliance_window, dtype=np.float64)
            k = np.arange(y.size) - 0.5 * (y.size - 1)
            slope = float(np.dot(k, y - y.mean()) / np.dot(k, k))
            self.compliance_trend = abs(slope) / max(abs(float(y.mean())), 1e-300)
            compliance_settled = self.compliance_trend <= self.convergence_compliance_tol
        if len(self._settle_solid_window) > w:
            self.net_phase_switches = int(np.count_nonzero(self._settle_solid_window[-1] != self._settle_solid_window[0]))
            switches_settled = self.net_phase_switches <= self.convergence_switch_tol * self.n_design_voxels * w
        return compliance_settled, switches_settled

    def _check_analysis_output(self, analysis_component):
        """
        Fails an analysis whose output would poison the update (docs/theory.md
        §5.8): a strain energy density field with a NaN or infinite value, or
        a published compliance that is not finite. Raises FloatingPointError,
        which the caller handles like any other analysis error: a phase
        backend's step is rejected and retaken smaller, a density backend's
        run ends on the last valid design (ANALYSIS_FAILED).
        """
        F = np.asarray(self.F_u)
        bad = ~np.isfinite(F)
        if np.any(bad):
            raise FloatingPointError(
                f"the analysis returned non-finite strain energy densities in "
                f"{int(np.count_nonzero(bad))} of {F.size} voxels")
        c = getattr(analysis_component, "compliance", None)
        if c is not None and not np.isfinite(c):
            raise FloatingPointError(f"the analysis returned a non-finite compliance ({c})")

    _OUTCOME_STATS_HEADER = ("optimisation_name,total_iterations,volume_error_pct,final_compliance,"
                             "buckling_status,buckling_target_safety_factor,final_critical_load_factor,"
                             "termination_reason,converged,connected,load_path_connected,"
                             "detached_fraction,grey_fraction,final_aggregated_load_factor,"
                             "outer_iterations,reconnection_resets,wall_time_hours,"
                             "final_exported_compliance,final_exported_volume_fraction,"
                             "final_exported_critical_load_factor,"
                             "final_non_discreteness,final_undecided_fraction,final_euler_characteristic,"
                             "final_n_tunnels,final_n_cavities,total_interior_nucleations,"
                             "total_surface_nucleations,first_interior_nucleation_iteration,"
                             "final_void_film_fraction,final_interface_faces")

    def _buckling_outcome_status(self) -> str:
        """
        DISABLED, NEVER_EVALUATED, UNKNOWN (the last evaluation rejected every
        candidate mode, or its eigensolve converged nothing even after the
        retry with larger limits), NO_CRITICAL_MODES (formally feasible, but only because
        no compressive mode exists: reported distinctly so that a tension-only
        result is never mistaken for a design driven to the safety factor),
        SATISFIED or VIOLATED (judged on the TRUE critical load factor, as the
        settled test is, eq. 6.27).
        """
        if not getattr(self, 'buckling_enabled', False):
            return "DISABLED"
        status = getattr(self, 'buckling_status', None)
        if not getattr(self, 'buckling_ever_evaluated', False):
            return "NEVER_EVALUATED"
        if status in BucklingResults.UNKNOWN_STATUSES:
            return "UNKNOWN"
        if status == BucklingResults.STATUS_NO_COMPRESSIVE_MODE:
            return "NO_CRITICAL_MODES"
        return "SATISFIED" if self._buckling_target_met() else "VIOLATED"

    def _save_outcome_stats(self):
        """
        Appends the final outcome statistics of this run to a master CSV file.

        The first seven columns are unchanged from earlier revisions; the
        termination/connectivity columns of docs/theory.md §5.4-§5.6 follow
        them. An existing outcome_stats.csv written with a different header
        is left untouched and the row goes to the first of
        outcome_stats_v2.csv, outcome_stats_v3.csv, ... that is absent or
        has this layout, so that no file ever mixes two layouts.
        """
        header = self._OUTCOME_STATS_HEADER
        # First of outcome_stats.csv, outcome_stats_v2.csv, _v3, ... that is
        # absent or already has this column layout.
        file_path = "outcome_stats.csv"
        for k in range(1, 100):
            file_path = "outcome_stats.csv" if k == 1 else f"outcome_stats_v{k}.csv"
            if not os.path.isfile(file_path):
                break
            try:
                with open(file_path, "r") as f:
                    existing_header = f.readline().strip()
            except Exception:
                existing_header = header
            if existing_header == header:
                break
        if file_path != "outcome_stats.csv":
            self._print_out(f"Note: outcome_stats.csv has a different column layout; writing this run's "
                            f"outcome to {file_path} instead.", "Minimal")
        file_exists = os.path.isfile(file_path)

        # Get metrics, defaulting to 'N/A' if the run failed prematurely
        iters = getattr(self, 'total_iterations', 'N/A')
        comp = getattr(self, 'compliance', 'N/A')

        # Calculate volume fraction error relative to the target volume (%)
        if hasattr(self, 'g1') and self.target_volume > 0:
            vol_error_pct = (self.g1 / self.target_volume) * 100.0
            vol_error_str = f"{vol_error_pct:.4f}"
        else:
            vol_error_str = 'N/A'

        comp_str = f"{comp:.6f}" if isinstance(comp, float) else str(comp)

        # Buckling outcome. Reported as the TRUE critical load factor (the
        # physically meaningful number, and what the default constraint
        # measure acts on) and the aggregated one (the smooth surrogate whose
        # gradient drives the update, conservative by up to m^(1/p)), with a
        # status that never reports an unknown result as safe.
        agg_str = "N/A"
        buckling_status = self._buckling_outcome_status()
        if getattr(self, 'buckling_enabled', False):
            lam_crit = getattr(self, 'critical_load_factor', float('inf'))
            lam_crit_str = self._fmt_load_factor(lam_crit)
            agg_str = self._fmt_load_factor(getattr(self, 'aggregated_load_factor', float('inf')))
            target_str = f"{self.buckling_safety_factor:.4f}"
        else:
            lam_crit_str = "N/A"
            target_str = "N/A"

        conn = getattr(self, 'connectivity', None) or {}
        def _b(v):
            return "N/A" if v is None else str(int(bool(v)))
        grey = getattr(self, 'grey_fraction', None)
        row = [
            self.optimisation_name, str(iters), vol_error_str, comp_str,
            buckling_status, target_str, lam_crit_str,
            str(getattr(self, 'termination_reason', 'N/A')),
            _b(getattr(self, 'converged', None)),
            _b(conn.get('connected')),
            _b(conn.get('load_path_connected')),
            f"{conn['detached_fraction']:.6f}" if 'detached_fraction' in conn else "N/A",
            f"{grey:.6f}" if grey is not None else "N/A",
            agg_str,
            str(getattr(self, 'outer_iterations_run', 'N/A')),
            str(getattr(self, 'reconnection_resets_done', 'N/A')),
            f"{self._elapsed() / 3600.0:.4f}",
            (f"{self.final_exported_compliance:.6f}"
             if getattr(self, 'final_exported_compliance', None) is not None else "N/A"),
            (f"{self.final_exported_volume_fraction:.6f}"
             if getattr(self, 'final_exported_volume_fraction', None) is not None else "N/A"),
            self._fmt_load_factor(getattr(self, 'final_exported_critical_load_factor', None)),
        ]
        # Final-design diagnostics (docs/theory.md §8.4.1).
        diag = getattr(self, '_diagnostics', None)

        def _v(x, fmt="{:.6f}"):
            if x is None:
                return "N/A"
            if isinstance(x, (int, np.integer)):
                return str(int(x))
            return fmt.format(float(x))
        row += [
            _v(getattr(self, 'non_discreteness', None)),
            _v(getattr(self, 'undecided_fraction', None)),
            _v(getattr(self, 'euler_characteristic', None)),
            _v(getattr(self, 'n_tunnels', None)),
            _v(getattr(self, 'n_cavities', None)),
            _v(diag.total_interior if diag is not None else None),
            _v(diag.total_surface if diag is not None else None),
            _v(diag.first_interior_iteration if diag is not None else None),
            _v(getattr(self, 'void_film_fraction', None)),
            _v(getattr(self, 'interface_faces', None)),
        ]

        try:
            # Append mode ('a') creates the file if it doesn't exist
            with open(file_path, "a") as f:
                if not file_exists:
                    f.write(header + "\n")
                f.write(",".join(row) + "\n")
            self._print_out(f"Appended final stats to {file_path}", "Normal")
        except Exception as e:
            self._print_out(f"Warning: Failed to write to {file_path}. Error: {e}", "Minimal")

    # Surface export (docs/theory.md §8.4, §10.6). The wrap is specified in voxels and the
    # work is done by surface_export.py, which is also what alpha_wrap_sweep.py and
    # reexport_surfaces.py use.
    _SURFACE_JOB_TIMEOUT_S = 3600.0

    def _surface_export_settings(self, tier: str) -> dict:
        """The keyword arguments of surface_export.build_surfaces for the "final" or "periodic" tier."""
        if tier == "periodic":
            return dict(method=self.periodic_surface_method,
                        alpha_voxels=self.periodic_surface_alpha_voxels,
                        offset_ratio=self.final_surface_offset_ratio,
                        min_component_fraction=self.final_surface_min_component_fraction,
                        smoothing_iterations=self.final_surface_smoothing_iterations)
        return dict(method=self.final_surface_method,
                    alpha_voxels=self.final_surface_alpha_voxels,
                    offset_ratio=self.final_surface_offset_ratio,
                    alpha_proportion=self.final_surface_alpha_proportion,
                    offset_proportion=self.final_surface_offset_proportion,
                    min_component_fraction=self.final_surface_min_component_fraction,
                    smoothing_iterations=self.final_surface_smoothing_iterations)

    def _report_surface_export(self, info: dict, label: str) -> None:
        """Console report of one finished export (settings, size, time, volume check, warnings)."""
        self._print_out(f"Saved {label} surface meshes: {summary_line(info)} -> {info['files']['filtered']}",
                        "Normal")
        if info.get("log"):
            self._print_out(f"[{label} surface] " + info["log"], "Verbose")
        for w in info.get("warnings", []):
            self._print_out(f"Warning ({label} surface): {w}.", "Minimal")
        if label == "final":
            self.final_surface_info = {k: v for k, v in info.items() if k != "log"}
        else:
            self.periodic_surface_infos = getattr(self, "periodic_surface_infos", []) + [
                {k: v for k, v in info.items() if k not in ("log", "files")}]

    def _finish_surface_jobs(self) -> None:
        """Waits for the background surface exports started by periodic saves and reports them."""
        jobs = getattr(self, "_surface_jobs", [])
        self._surface_jobs = []
        for job in jobs:
            try:
                job["proc"].wait(timeout=self._SURFACE_JOB_TIMEOUT_S)
            except Exception as e:
                job["proc"].kill()
                self._print_out(f"Warning: the background surface export of {job['iter_name']} did not "
                                f"finish ({e}); it was stopped.", "Minimal")
                continue
            try:
                with open(job["result_path"]) as f:
                    res = json.load(f)
            except Exception as e:
                self._print_out(f"Warning: the background surface export of {job['iter_name']} left no "
                                f"result. Error: {e}", "Minimal")
                continue
            finally:
                for p in (job["job_path"], job["result_path"]):
                    try:
                        os.remove(p)
                    except OSError:
                        pass
            if res.get("ok"):
                self._report_surface_export(res["info"], "periodic")
            else:
                self._print_out(f"Warning: the background surface export of {job['iter_name']} failed: "
                                f"{res.get('error')}", "Minimal")

    def _export_final_surface_alpha_wrap(self, final_analysis_component, field: np.ndarray,
                                          level_set_value: float, pad_value: float,
                                          tier: str = "final") -> None:
        """
        Marching-cubes (+ CGAL alpha-wrap) surface export of a design, run regardless of which
        analysis backend (Conformal/Regular/ConjugateGradient) produced 'field'. The work is done
        by surface_export.export_surfaces (docs/theory.md §8.4):

        - The wrap is specified in VOXELS (final_surface_alpha_voxels, offset =
          final_surface_offset_ratio x alpha), not as a fraction of the bounding-box diagonal
          (the old default of 0.1 was a 17-voxel ball on a 160x40x40 domain and filled the design
          in). final_surface_alpha_proportion / final_surface_offset_proportion still override
          this for the final tier when given.
        - Components smaller than final_surface_min_component_fraction of the solid voxels are
          removed on the voxel grid; the main body is wrapped ONCE. The "raw" file adds the
          removed specks back as plain marching-cubes shells.
        - final_surface_method = "marching_cubes" skips CGAL.
        - Two tiers: "final" (the finished design; fine alpha) and "periodic" (the saves every
          save_frequency iterations; periodic_surface_alpha_voxels, which may be set
          coarser than the final one to save time, or marching cubes). With surface_export_background the periodic tier runs in a separate process so
          the optimisation loop is not blocked (finished by _finish_surface_jobs).

        Files, next to the design's other outputs (final_analysis_component.analysis_dir):
        '{iter_name}_alphawrap_raw.stl' / '_filtered.stl' (method "alpha_wrap") or
        '{iter_name}_mc_raw.stl' / '_filtered.stl' (method "marching_cubes").

        Args:
            final_analysis_component: supplies voxel_size_m, analysis_dir and iter_name.
            field: the design field to contour, NOT padded: final_phi for a "phase" backend, the
                design density for a "density" backend.
            level_set_value: isosurface threshold, in the units of 'field' (solid = field >= level).
            pad_value: the field's own void value, used to pad one voxel on every side so the
                isosurface closes at the domain boundary (void_value for a phase field, 0.0 for a
                density: padding a density with 1.0 would treat the boundary as solid).
            tier: "final" or "periodic".
        """
        analysis_dir = final_analysis_component.analysis_dir
        iter_name = final_analysis_component.iter_name
        voxel_size_m = final_analysis_component.voxel_size_m
        kwargs = self._surface_export_settings(tier)
        printout = (self.printouts == "Verbose")

        if tier == "periodic" and self.surface_export_background:
            try:
                import subprocess
                import sys
                os.makedirs(analysis_dir, exist_ok=True)
                job_path = os.path.join(analysis_dir, f"{iter_name}_surface_job.npz")
                result_path = os.path.join(analysis_dir, f"{iter_name}_surface_job.json")
                cfg = dict(kwargs, level=float(level_set_value), pad_value=float(pad_value),
                           voxel_size_m=float(voxel_size_m), analysis_dir=analysis_dir,
                           iter_name=iter_name, result_path=result_path, printout=printout)
                np.savez(job_path, field=np.ascontiguousarray(field),
                         config_json=np.array(json.dumps(cfg)))
                code_dir = os.path.dirname(os.path.abspath(__file__))
                launcher = ("import sys; sys.path.insert(0, sys.argv[1]); "
                            "import baseline_optimiser; sys.exit(baseline_optimiser.run_surface_job(sys.argv[2]))")
                proc = subprocess.Popen([sys.executable, "-c", launcher, code_dir, job_path],
                                        cwd=code_dir, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                self._surface_jobs = getattr(self, "_surface_jobs", []) + [
                    {"proc": proc, "job_path": job_path, "result_path": result_path,
                     "iter_name": iter_name}]
                self._print_out(f"Started the periodic surface export of {iter_name} in the background.",
                                "Verbose")
                return
            except Exception as e:
                self._print_out(f"Warning: could not start the background surface export ({e}); "
                                f"running it in this process.", "Minimal")

        # A finished design is exported only after any background exports have finished, so
        # that two processes never write the same directory at once.
        if tier == "final":
            self._finish_surface_jobs()
        info = export_surfaces(field, level_set_value, pad_value, voxel_size_m, analysis_dir,
                                  iter_name, printout=printout, **kwargs)
        self._report_surface_export(info, tier)

    @staticmethod
    def _capture_pending_fields(component) -> dict | None:
        """
        Copies of a component's grid-native output arrays, taken before
        finalize_iteration_output writes and clears them: {"cell": {name:
        array on the voxel grid}, "point": {name: array on the grid points},
        "grid_shape", "spacing_m"}. None if the component saved nothing.
        """
        if component is None or getattr(component, "_grid_shape", None) is None:
            return None
        return {"cell": {k: np.array(v, copy=True) for k, v in component._pending_cell_data.items()},
                "point": {k: np.array(v, copy=True) for k, v in component._pending_point_data.items()},
                "grid_shape": tuple(int(v) for v in component._grid_shape),
                "spacing_m": tuple(float(v) for v in (component._grid_hx, component._grid_hy,
                                                      component._grid_hz))}

    @staticmethod
    def _jsonable(x):
        """Converts numpy scalars/arrays to JSON types; +inf -> "inf", -inf -> "-inf", nan -> None."""
        if isinstance(x, dict):
            return {str(k): BaseOptimiser._jsonable(v) for k, v in x.items()}
        if isinstance(x, (list, tuple)):
            return [BaseOptimiser._jsonable(v) for v in x]
        if isinstance(x, np.ndarray):
            return [BaseOptimiser._jsonable(v) for v in x.tolist()]
        if isinstance(x, (bool, np.bool_)):
            return bool(x)
        if isinstance(x, (int, np.integer)):
            return int(x)
        if isinstance(x, (float, np.floating)):
            x = float(x)
            if np.isnan(x):
                return None
            if np.isinf(x):
                return "inf" if x > 0 else "-inf"
            return x
        if x is None or isinstance(x, str):
            return x
        return str(x)

    @staticmethod
    def _preconditioner_used_record():
        """{"preconditioner_used": {"requested->used": n_solves}} of the GPU PCG solves of this process; {} when none ran
        (CPU stand-in or other backends), so those run_info.json files carry no new key."""
        try:
            from gpu_analysis import PRECONDITIONER_USED
        except Exception:
            return {}
        used = {k: v for k, v in PRECONDITIONER_USED.items() if not k.startswith("_")}
        return {"preconditioner_used": used} if used else {}

    @staticmethod
    def _engines_record():
        """{"engines": gpu_analysis.RECORD} when an engine other than "classic" was selected in this process (capability,
        selection and reason, fallbacks, buckling timing); {} otherwise, so classic run_info.json files are unchanged."""
        try:
            import gpu_analysis
        except Exception:
            return {}
        return {"engines": json.loads(json.dumps(gpu_analysis.RECORD, default=str))} if gpu_analysis.RECORD else {}

    def _analysis_keyword(self, name, default=None):
        """A keyword the analysis factory was built with (functools.partial), else default."""
        return getattr(self.analysis_component_factory, "keywords", {}).get(name, default)

    def _write_final_fields(self, final_phi: np.ndarray):
        """final_fields.npz: every per-voxel field of the final design and its analyses."""
        self._gpu_materialise()
        rho = self._rho(phi=final_phi)
        solid = self._solid_mask(final_phi)
        labels, _ = component_sizes(solid)
        arrays = {
            "phi": np.asarray(final_phi, dtype=np.float32),
            "rho": rho.astype(np.float32),
            "solid": solid,
            "component_labels": labels.astype(np.int32),
            "design_mask": self.design_mask,
            "preserve_mask": np.asarray(self.preserve_mask, dtype=bool),
            "obstacle_mask": np.asarray(self.obstacle_mask, dtype=bool),
            "load_mask": np.asarray(self.load_mask, dtype=bool),
            "support_mask": np.asarray(self.support_mask, dtype=bool),
        }
        load_vectors = self.csg_data.get("load_vectors") if self.csg_data is not None else None
        if load_vectors is not None:
            arrays["load_magnitude"] = np.linalg.norm(np.asarray(load_vectors, dtype=np.float64),
                                                      axis=-1).astype(np.float32)
        if self._analysis_field_type == "density":
            rho_m = self._rho_model(rho=rho)
            arrays["rho_model"] = rho_m.astype(np.float32)
        # State of the last optimisation iteration: its analysis and the
        # terms of the update it applied (eq. 5.1).
        for name, attr in (("F", "F_u"), ("F0", "F0"), ("stiffness_term", "stiffness_term"),
                           ("volume_term", "volume_term"), ("piecewise_term", "piecewise_term"),
                           ("buckling_term", "applied_buckling_term"), ("diffusion_term", "diffusion_term"),
                           ("net_push", "net_push"), ("lambda2", "lambda2"), ("G2", "g2")):
            v = getattr(self, attr, None)
            if isinstance(v, np.ndarray) and v.shape == final_phi.shape:
                arrays[name] = v.astype(np.float32)
        for prefix, captured in (("model", self._final_model_fields), ("exported", self._exported_fields)):
            if not captured:
                continue
            for k, v in captured.get("cell", {}).items():
                arrays[f"{prefix}_{k}"] = np.asarray(v)
            for k, v in captured.get("point", {}).items():
                arrays[f"{prefix}_point_{k}"] = np.asarray(v)
            for k, v in captured.get("modes", {}).items():
                arrays[f"{prefix}_point_{k}"] = np.asarray(v)
        path = os.path.join(self.optimisation_name, "final_fields.npz")
        np.savez_compressed(path, **arrays)
        self._print_out(f"Saved the final fields to {path} ({len(arrays)} arrays).", "Normal")

    def _write_run_info(self, final_phi: np.ndarray):
        """run_info.json: the run's configuration and final scalars, for the postprocess section."""
        conn = getattr(self, "connectivity", None) or {}
        diag = getattr(self, "_diagnostics", None)
        load_vectors = self.csg_data.get("load_vectors") if self.csg_data is not None else None
        exported_buckling = getattr(self, "_exported_buckling", None)
        params = {k: v for k, v in getattr(self, "_optimiser_params_dict", {}).items()
                  if isinstance(v, (int, float, str, bool, list, tuple, type(None)))}
        # Topology of the exported design (the solid of final_fields.npz): components b0, tunnels b1 and
        # internal voids b2 (void that does not reach the outside), docs/theory.md eq. (11.19).
        try:
            _t = topology(self._solid_mask(final_phi))
            final_topology = {"euler_characteristic": int(_t["euler"]), "n_components": int(_t["b0"]),
                              "n_tunnels": int(_t["b1"]), "n_cavities": int(_t["b2"])}
        except Exception:
            final_topology = None
        info = {
            "optimisation_name": self.optimisation_name,
            "analysis_class": self._analysis_class_name,
            "field_type": self._analysis_field_type,
            "domain_dims": list(self.domain_dims),
            "n_voxels": int(final_phi.size),
            "n_design_voxels": int(self.n_design_voxels),
            "n_preserve_voxels": int(np.count_nonzero(self.preserve_mask)),
            "n_obstacle_voxels": int(np.count_nonzero(self.obstacle_mask)),
            "n_preserve_obstacle_overlap_voxels": int(getattr(self, "n_preserve_obstacle_overlap", 0)),
            "n_support_in_obstacle_voxels": int(getattr(self, "n_support_in_obstacle", 0)),
            "n_load_in_obstacle_voxels": int(getattr(self, "n_load_in_obstacle", 0)),
            "n_load_voxels": int(np.count_nonzero(self.load_mask)),
            "n_support_voxels": int(np.count_nonzero(self.support_mask)),
            "total_load_N": (np.asarray(load_vectors, dtype=np.float64).reshape(-1, 3).sum(axis=0)
                             if load_vectors is not None else None),
            "voxel_size_mm": float(self._analysis_keyword("voxel_size_mm", 1.0)),
            "youngs_modulus": self._analysis_keyword("youngs_modulus"),
            "poisson_ratio": self._analysis_keyword("poisson_ratio"),
            **self._preconditioner_used_record(),
            **self._engines_record(),
            "optimiser_parameters": params,
            "level_set_value": self.level_set_value,
            "termination_reason": getattr(self, "termination_reason", None),
            "converged": bool(getattr(self, "converged", False)),
            "total_iterations": int(getattr(self, "total_iterations", 0)),
            "outer_iterations": getattr(self, "outer_iterations_run", None),
            "reconnection_resets": getattr(self, "reconnection_resets_done", None),
            "reconnection_reset_iterations": list(getattr(self, "reconnection_reset_iterations", [])),
            "settled_disconnected_events": getattr(self, "settled_disconnected_events", None),
            "settled_volume_shortfall_events": getattr(self, "settled_volume_shortfall_events", None),
            "freeze_volume_shortfall_events": getattr(self, "freeze_volume_shortfall_events", None),
            "buckling_stall_exit_events": getattr(self, "buckling_stall_exit_events", None),
            "converged_with_volume_shortfall": bool(getattr(self, "converged_with_volume_shortfall", False)),
            "nonphysical_analyses": getattr(self, "nonphysical_analyses", None),
            "invalid_analyses": getattr(self, "invalid_analyses", None),
            "rejected_steps": getattr(self, "rejected_steps", None),
            "retried_analysis_errors": getattr(self, "retried_analysis_errors", None),
            "min_step_scale": getattr(self, "min_step_scale", None),
            "failure": getattr(self, "failure", None),
            "random_seed": getattr(self, "random_seed", None),
            "initial_design": getattr(self, "initial_design_info", None),
            "overrides": getattr(self, "overrides", None),
            "provenance": getattr(self, "provenance", None),
            "wall_time_s": self._elapsed(),
            "final": {
                "compliance": getattr(self, "compliance", None),
                "compliance_is_exact": getattr(self, "compliance_is_exact", None),
                "G1": getattr(self, "g1", None),
                "mean_density": float(np.mean(self._rho(phi=final_phi))),
                "exported_volume_fraction": getattr(self, "final_exported_volume_fraction", None),
                "exported_over_target_volume": (
                    getattr(self, "final_exported_volume_fraction", None) / self.target_volume
                    if getattr(self, "final_exported_volume_fraction", None) is not None and self.target_volume > 0
                    else None),
                "void_film_fraction": getattr(self, "void_film_fraction", None),
                "exported_compliance": getattr(self, "final_exported_compliance", None),
                "grey_fraction": getattr(self, "grey_fraction", None),
                "settled_count": getattr(self, "settled_count", None),
                "settle_failures": getattr(self, "settle_failures", None),
                "penalisation_power": getattr(self, "current_penalisation_power", None),
                "lambda1": getattr(self, "lambda1", None),
                "mu1": getattr(self, "mu1", None),
                "mu2": getattr(self, "mu2", None),
                "lambda2_min": float(np.min(self.lambda2)) if hasattr(self, "lambda2") else None,
                "lambda2_mean": float(np.mean(self.lambda2)) if hasattr(self, "lambda2") else None,
                "lambda2_max": float(np.max(self.lambda2)) if hasattr(self, "lambda2") else None,
                "connectivity": conn,
                "topology": final_topology,
            },
            "buckling": {
                "enabled": bool(self.buckling_enabled),
                "status": self._buckling_outcome_status(),
                "safety_factor": self.buckling_safety_factor if self.buckling_enabled else None,
                "constraint_measure": self.buckling_constraint_measure if self.buckling_enabled else None,
                "critical_load_factor_model": getattr(self, "critical_load_factor", None) if self.buckling_enabled else None,
                "last_valid_critical_load_factor": getattr(self, "last_valid_critical_load_factor", None) if self.buckling_enabled else None,
                "aggregated_load_factor_model": getattr(self, "aggregated_load_factor", None) if self.buckling_enabled else None,
                "critical_load_factor_exported": getattr(self, "final_exported_critical_load_factor", None),
                "exported_load_factors": ([1.0 / m for m in exported_buckling.mu_values if m > 0]
                                          if exported_buckling is not None else None),
                "exported_status": exported_buckling.status if exported_buckling is not None else None,
                "unknown_evaluations": getattr(self, "buckling_unknown_evaluations", None),
                "evaluations": getattr(self, "buckling_evaluations_done", None) if self.buckling_enabled else None,
                "idle_skips": getattr(self, "buckling_idle_skips", None) if self.buckling_enabled else None,
                "lambda3": getattr(self, "lambda3", None) if self.buckling_enabled else None,
                "mu3": getattr(self, "mu3", None) if self.buckling_enabled else None,
            },
            "diagnostics": {
                "total_interior_nucleations": diag.total_interior if diag is not None else None,
                "total_surface_nucleations": diag.total_surface if diag is not None else None,
                "first_interior_nucleation_iteration": diag.first_interior_iteration if diag is not None else None,
                "last_topology_change_iteration": diag.last_topology_change_iteration if diag is not None else None,
            },
            "threshold_sweep_levels": list(self.threshold_sweep_levels),
            # Settings, size, time and volume check of the surface exports (docs/theory.md §8.4);
            # final_surface.volume_ratio_filtered is the volume of the STL over that of the solid voxels.
            "final_surface": getattr(self, "final_surface_info", None),
            "periodic_surfaces": getattr(self, "periodic_surface_infos", None),
        }
        # Discreteness and interface crispness of the final design, and how faithfully the
        # exported geometry reproduces the analysed model (docs/theory.md §8.4.7); the same
        # definitions for every method, so runs of PCLS, SIMP and the level set compare directly.
        try:
            crisp = crispness_metrics(
                design_fraction(final_phi, self.solid_value, self.void_value), self.design_mask)
            crisp.update(export_fidelity(
                getattr(self, "compliance", None), getattr(self, "final_exported_compliance", None),
                getattr(self, "last_valid_critical_load_factor", None) if self.buckling_enabled else None,
                getattr(self, "final_exported_critical_load_factor", None)))
        except Exception as e:
            crisp = {"error": f"{type(e).__name__}: {e}"}
        info["crispness"] = crisp
        if getattr(self, "restart_source", None):   # resume_state (PCLS) only; absent otherwise
            info["restart_source"] = self.restart_source
        path = os.path.join(self.optimisation_name, "run_info.json")
        with open(path, "w") as f:
            json.dump(self._jsonable(info), f, indent=2)
        self._print_out(f"Saved the run record to {path}.", "Normal")
        if "error" not in crisp:
            self._print_out(
                f"Crispness of the final design (docs/theory.md §8.4.7): M_nd {crisp['non_discreteness']:.4f}, "
                f"grey {crisp['grey_fraction']:.4f}, interface width {crisp['interface_width_vox']:.2f} voxel, "
                f"bulk grey {crisp['bulk_grey_fraction']:.4f}, threshold sensitivity "
                f"{crisp['threshold_volume_sensitivity']:.4f}, topology stable over cuts 0.3-0.7: "
                f"{crisp['topology_stable']}.", "Minimal")
        if final_topology is not None:
            self._print_out(
                f"Final design topology: {final_topology['n_components']} component(s), "
                f"{final_topology['n_tunnels']} tunnel(s), {final_topology['n_cavities']} internal void(s) "
                f"(Euler characteristic {final_topology['euler_characteristic']}).", "Minimal")

    def _threshold_sweep(self, final_phi: np.ndarray):
        """
        threshold_sweep.csv: the final design cut at phi >= V + s (S - V) for
        each s in threshold_sweep_levels (s = 0.5 is the exported design), with
        its volume fraction, topology, connectivity and -- if
        threshold_sweep_analyses -- compliance and, with buckling enabled,
        critical load factor (docs/theory.md §8.4.2). A converged binary design
        gives the same answer at every cut; a design that still holds grey
        material does not, and the spread measures how much the exported
        result depends on the choice of threshold.
        """
        levels = list(self.threshold_sweep_levels)
        if not levels:
            return
        run_fe = self.threshold_sweep_analyses
        if run_fe and getattr(self, "termination_reason", None) == "TIME_LIMIT":
            run_fe = False
            self._print_out("Threshold sweep: skipping the analyses because the run stopped on its "
                            "time limit (geometry and topology only).", "Minimal")
        if run_fe and self.time_limit_hours is not None and self._iteration_times:
            per = float(np.median(self._iteration_times))
            need = per * len(levels) * (2.0 if self.buckling_enabled else 1.0)
            left = self.time_limit_hours * 3600.0 - self._elapsed()
            if need > left:
                run_fe = False
                self._print_out(f"Threshold sweep: skipping the analyses, which would take about "
                                f"{need / 60.0:.1f} min with {max(0.0, left) / 60.0:.1f} min of the time "
                                f"limit left (geometry and topology only).", "Minimal")
        S, V = self.solid_value, self.void_value
        preserve = np.asarray(self.preserve_mask, dtype=bool)
        obstacle = np.asarray(self.obstacle_mask, dtype=bool)
        phi = np.asarray(final_phi, dtype=np.float32).copy()
        phi[preserve] = S
        phi[obstacle] = V
        header = ["level", "phi_threshold", "volume_fraction", "euler_characteristic", "n_components",
                  "n_tunnels", "n_cavities", "interface_faces", "load_path_connected",
                  "detached_fraction", "connected", "compliance", "critical_load_factor",
                  "aggregated_load_factor", "buckling_status", "floating_load_fraction", "analysis_time", "note"]
        rows = []
        for level in levels:
            threshold = V + level * (S - V)
            solid = ((phi >= threshold) | preserve) & ~obstacle
            conn = self._connectivity_of_solid(solid)
            topo = topology(solid, n_components=conn["n_components"])
            row = {"level": level, "phi_threshold": threshold,
                   "volume_fraction": float(np.count_nonzero(solid)) / solid.size,
                   "euler_characteristic": topo["euler"], "n_components": topo["b0"],
                   "n_tunnels": topo["b1"], "n_cavities": topo["b2"],
                   "interface_faces": interface_faces(solid),
                   "load_path_connected": int(conn["load_path_connected"]),
                   "detached_fraction": conn["detached_fraction"], "connected": int(conn["connected"]),
                   "compliance": None, "critical_load_factor": None, "aggregated_load_factor": None,
                   "buckling_status": None, "floating_load_fraction": None, "analysis_time": None, "note": ""}
            is_exported = abs(level - 0.5) < 1e-9 and self._analysis_field_type == "density"
            if run_fe and is_exported and getattr(self, "final_exported_compliance", None) is not None:
                # Already analysed: the exported design.
                row["compliance"] = self.final_exported_compliance
                row["critical_load_factor"] = getattr(self, "final_exported_critical_load_factor", None)
                eb = getattr(self, "_exported_buckling", None)
                if eb is not None:
                    row["aggregated_load_factor"] = eb.aggregated_load_factor
                    row["buckling_status"] = eb.status
                row["note"] = "exported design"
            elif run_fe:
                t0 = timer()
                try:
                    comp = self.analysis_component_factory(iter_name="test", printouts="Silent",
                                                           data_dir=self.analysis_dir, save_meshes=False)
                    if comp.field_type == "density":
                        comp.run_analysis(field=np.where(solid, 1.0, self.void_density).astype(np.float32),
                                          level_set_value=None)
                    else:
                        comp.run_analysis(field=phi, level_set_value=threshold)
                    row["compliance"] = getattr(comp, "compliance", None)
                    floating = getattr(comp, "floating_load_fraction", None)
                    row["floating_load_fraction"] = floating
                    if floating is not None and floating > 0.0:
                        # The cut design's load path is broken: no compliance.
                        row["compliance"] = float("nan")
                        row["note"] = f"load path cut ({100 * floating:.1f}% of the load on unsupported pieces)" 
                    if (self.buckling_enabled and getattr(self, "buckling_ever_evaluated", False)
                            and hasattr(comp, "run_buckling")):
                        br = comp.run_buckling(
                            num_modes=self.buckling_num_modes, aggregation_p=self.buckling_aggregation_p,
                            prev_modes=None, min_energy_fraction=self.buckling_min_energy_fraction,
                            mode_continuity_warning_threshold=self.buckling_mode_continuity_warning_threshold,
                            min_density_for_geometric_stiffness=self.buckling_min_density_for_geometric_stiffness)
                        row["critical_load_factor"] = br.critical_load_factor
                        row["aggregated_load_factor"] = br.aggregated_load_factor
                        row["buckling_status"] = br.status
                except Exception as e:
                    row["note"] = f"analysis failed: {str(e)[:120]}"
                row["analysis_time"] = timer() - t0
            rows.append(row)

        def _cell(v):
            if v is None:
                return "N/A"
            if isinstance(v, (float, np.floating)):
                v = float(v)
                return "nan" if np.isnan(v) else ("inf" if np.isinf(v) else f"{v:.8g}")
            return str(v).replace(",", ";")
        path = os.path.join(self.optimisation_name, "threshold_sweep.csv")
        with open(path, "w") as f:
            f.write(",".join(header) + "\n")
            for row in rows:
                f.write(",".join(_cell(row[h]) for h in header) + "\n")
        self._print_out(f"Saved the threshold sweep ({len(rows)} cuts) to {path}.", "Normal")

    def _write_pvd(self):
        """data/iterations.pvd: the saved .vtr files as one ParaView time series (timestep = iteration)."""
        files = sorted(getattr(self, "_saved_iteration_files", []), key=lambda t: t[0])
        if not files:
            return
        path = os.path.join(self.analysis_dir, "iterations.pvd")
        with open(path, "w") as f:
            f.write('<?xml version="1.0"?>\n'
                    '<VTKFile type="Collection" version="0.1" byte_order="LittleEndian">\n'
                    '  <Collection>\n')
            for it, file_path in files:
                rel = os.path.relpath(file_path, self.analysis_dir).replace(os.sep, "/")
                f.write(f'    <DataSet timestep="{it}" group="" part="0" file="{rel}"/>\n')
            f.write('  </Collection>\n</VTKFile>\n')
        self._print_out(f"Saved the ParaView time series of {len(files)} saved iterations to {path}.", "Normal")

    def _write_research_outputs(self, final_phi: np.ndarray):
        """Runs each research-output step in its own try/except (docs/theory.md §8.4.1-§8.4.3)."""
        steps = [("final fields", self._write_final_fields),
                 ("run record", self._write_run_info),
                 ("threshold sweep", self._threshold_sweep)]
        for label, fn in steps:
            try:
                fn(final_phi)
            except Exception as e:
                self._print_out(f"Warning: failed to write the {label}. Error: {e}", "Minimal")
        try:
            self._write_pvd()
        except Exception as e:
            self._print_out(f"Warning: failed to write the ParaView time series. Error: {e}", "Minimal")
        if self.postprocess:
            try:
                t0 = timer()
                postprocess_run(self.optimisation_name, verbose=self.printouts != "Silent")
                self._print_out(f"Post-processing (plots, summary, report) finished in {timer() - t0:.1f} s: "
                                f"see {os.path.join(self.optimisation_name, 'plots')} and "
                                f"{os.path.join(self.optimisation_name, 'report.pdf')}.", "Minimal")
            except Exception as e:
                self._print_out(f"Warning: post-processing failed; the run's data files are complete and "
                                f"'python optimise.py --postprocess-only {self.optimisation_name}' can be re-run. Error: {e}",
                                "Minimal")

    def run_optimisation(self) -> np.ndarray:
        """
        Runs the PCLS optimisation process by executing the nested outer and inner loops.

        Returns:
            np.ndarray: The final phi field after optimisation.
        """
        self._run_start_time = timer()
        self._setup_phi()
        # Provenance (docs/theory.md §8.1): optimise.py collects it first (for
        # problem_description.txt) and hands it over; otherwise collect it here.
        try:
            if getattr(self, "provenance", None) is None:
                self.provenance = collect_provenance(problem_file=getattr(self, "problem_file", None),
                                                overrides=getattr(self, "overrides", None))
            self.provenance["random_seed"] = self.random_seed
            if getattr(self, "initial_design_info", None):
                self.provenance["initial_design"] = self.initial_design_info
        except Exception as e:
            self.provenance = {"error": str(e)}

        try:
            final_phi = self._outer_loop()
        except Exception as e:
            # Anything else that goes wrong inside the loops (docs/theory.md
            # §8.4.4): record it and still write the final records for the
            # last design the run adopted. optimise.py exits non-zero.
            self._record_failure("ERROR", f"{type(e).__name__}: {e}", traceback.format_exc(),
                                 failed_phi=getattr(self, "phi", None))
            self.termination_reason = "ERROR"
            # The last design whose analysis was VALID: after a rejected
            # (cut) Conformal step _last_analysed_phi is the design the run
            # threw away (optimiser review, 2026-10).
            final_phi = getattr(self, "_last_valid_phi", None)
            if final_phi is None:
                final_phi = getattr(self, "_last_analysed_phi", None)
            if final_phi is None:
                final_phi = getattr(self, "phi", None)
            if final_phi is None:
                raise
            final_phi = np.asarray(final_phi, dtype=np.float32).copy()
            for name, default in (("outer_iterations_run", None), ("converged", False)):
                if not hasattr(self, name):
                    setattr(self, name, default)

        # Unconditionally save the final structure, independent of
        # `save_frequency` and of how _outer_loop/_inner_loop returned.
        #
        # Every mesh saved during the loops above is of the design ENTERING
        # an iteration: `_inner_loop` builds and runs the analysis component
        # on `phi_old` (the field the iteration starts with, before that
        # iteration's own update is applied), decided by `save_this_iter`
        # BEFORE the update happens. Both `_inner_loop` and `_outer_loop` can
        # also return early through their own convergence checks, with no
        # further iteration left to catch and save the update that was just
        # computed. Put together, the literal array returned by this method
        # -- the actual final result of the optimisation -- was frequently
        # never analysed with `save_meshes=True` at all: a converged run
        # exits via the early-return branch far more often than it exhausts
        # `inner_loop_max_iters`, and any `save_frequency > 1` (let alone
        # `save_frequency` values close to or above `inner_loop_max_iters`)
        # makes it correspondingly less likely that whichever inner_iter the
        # loop happened to stop on also satisfied the modulo test. This is
        # fixed here by always running one extra, dedicated analysis pass on
        # `final_phi` with `save_meshes=True`, decoupled entirely from
        # `save_frequency`. It does not touch `self.F_u`/`self.g1`/
        # `self.compliance` etc. (it never assigns into them), so it cannot
        # perturb the metrics `_save_outcome_stats` below reports for the
        # run's actual last optimisation step.
        final_analysis_component = None
        export_field = None
        export_level_set_value = None
        export_pad_value = None
        self._final_model_fields = None
        self._exported_fields = None
        self._exported_buckling = None
        try:
            self._diagnostics.snapshot(self.total_iterations, final_phi, force=True)
        except Exception as e:
            self._print_out(f"Warning: failed to save the final phi snapshot. Error: {e}", "Minimal")
        try:
            final_analysis_component = self.analysis_component_factory(
                iter_name=os.path.join("outer_final", "inner_final"),
                printouts=self.printouts,
                data_dir=self.analysis_dir,
                save_meshes=True,
            )
            final_field_data = {"termination_reason": str(getattr(self, "termination_reason", "")),
                                "converged": int(bool(getattr(self, "converged", False))),
                                "penalisation_power": self.current_penalisation_power}
            if final_analysis_component.field_type == "phase":
                # The model the run finished with (the penalised material
                # interpolation, docs/theory.md §4.8); the exported design is
                # the same mesh with the full modulus, analysed just below.
                if self._phase_takes_stiffness_density(final_analysis_component):
                    final_analysis_component.run_analysis(
                        field=final_phi, level_set_value=self.level_set_value,
                        stiffness_density=self._rho_model(rho=self._rho(phi=final_phi)))
                else:
                    final_analysis_component.run_analysis(field=final_phi, level_set_value=self.level_set_value)
                self._final_phase_model_compliance = getattr(final_analysis_component, "compliance", None)
                try:
                    if getattr(final_analysis_component, "material_interpolation_applied", False):
                        self._final_phase_exported_compliance = float(
                            final_analysis_component.uniform_material_compliance())
                    else:
                        self._final_phase_exported_compliance = self._final_phase_model_compliance
                except Exception as e:
                    self._final_phase_exported_compliance = None
                    self._print_out(f"Warning: could not analyse the exported design. Error: {e}", "Minimal")
                # Buckling of the same exported geometry (full modulus in every tet;
                # docs/theory.md §6.10), for the load factor of the finished design.
                self._final_phase_exported_buckling = None
                if (((self.buckling_enabled and getattr(self, "buckling_ever_evaluated", False))
                     or self.buckling_report_exported)
                        and hasattr(final_analysis_component, "uniform_material_buckling")):
                    try:
                        self._final_phase_exported_buckling = final_analysis_component.uniform_material_buckling(
                            num_modes=self.buckling_num_modes,
                            aggregation_p=self.buckling_aggregation_p,
                            prev_modes=None,
                            min_energy_fraction=self.buckling_min_energy_fraction,
                            mode_continuity_warning_threshold=self.buckling_mode_continuity_warning_threshold,
                            min_density_for_geometric_stiffness=self.buckling_min_density_for_geometric_stiffness,
                        )
                    except Exception as e:
                        self._print_out(f"Warning: could not run the buckling analysis of the exported "
                                        f"design. Error: {e}", "Minimal")
                # The standardised .vtr is only written by
                # finalize_iteration_output (docs/theory.md §7.6); the final
                # pass previously never called it, so the run's final design
                # had no iteration-data file at all.
                self._final_model_fields = self._capture_pending_fields(final_analysis_component)
                try:
                    written = final_analysis_component.finalize_iteration_output(field_data=final_field_data)
                    if written:
                        self._saved_iteration_files.append((self.total_iterations, written))
                except Exception as e:
                    self._print_out(f"Warning: failed to write the final iteration data file. Error: {e}", "Minimal")
                # Reuse for the backend-agnostic alpha-wrap export below: the
                # phase field's own isosurface threshold and void padding.
                export_field = final_phi
                export_level_set_value = self.level_set_value
                export_pad_value = self.void_value
            elif final_analysis_component.field_type == "density":
                final_rho = self._rho(phi=final_phi)
                # Analysed with the same (penalised) stiffness interpolation
                # the run finished with, so the saved fields match the last
                # logged iteration; for a converged, binary design this is
                # identical to final_rho. The EXPORTED geometry below is the
                # design density final_rho >= 0.5, i.e. phi >= level set.
                final_analysis_component.run_analysis(field=self._rho_model(rho=final_rho), level_set_value=None)
                try:
                    if getattr(final_analysis_component, "_grid_shape", None) is not None:
                        final_analysis_component._pending_cell_data["Design_Density"] = final_rho.astype(np.float32)
                        final_analysis_component._pending_cell_data["Phi"] = np.asarray(final_phi, dtype=np.float32)
                    self._final_model_fields = self._capture_pending_fields(final_analysis_component)
                    written = final_analysis_component.finalize_iteration_output(field_data=final_field_data)
                    if written:
                        self._saved_iteration_files.append((self.total_iterations, written))
                except Exception as e:
                    self._print_out(f"Warning: failed to write the final iteration data file. Error: {e}", "Minimal")
                # A density field has no self.level_set_value/void_value of
                # its own (those are phase-field concepts). The design density
                # rho(phi) is linear in phi, so rho = 0.5 is (to within
                # void_density) the level set phi = (S + V)/2 itself, and 0.0
                # is the density field's own void value. See docs/theory.md
                # §8.4.
                export_field = final_rho
                export_level_set_value = 0.5
                export_pad_value = 0.0
            else:
                raise ValueError("Analysis component is requesting an invalid (or nonexistent) field type.")
            self._print_out("Saved final structure to outer_final/inner_final.", "Normal")
        except Exception as e:
            # The optimisation itself already completed successfully; failing
            # to write this extra, purely-for-output mesh should not prevent
            # the outcome stats below from being recorded.
            self._print_out(f"Warning: failed to save the final structure. Error: {e}", "Minimal")

        # Backend-agnostic closed-surface export of the final structure via
        # marching cubes + CGAL alpha-wrap, regardless of which analysis
        # backend the run actually used (only Conformal builds a surface mesh
        # of its own, during the run). Kept in its own try/except: this is
        # purely an extra output artifact, so a failure here (e.g. CGAL not
        # installed, or a degenerate/empty isosurface) must not affect the
        # native final-structure save above or the outcome stats below.
        if (
            final_analysis_component is not None
            and export_field is not None
            and export_level_set_value is not None
            and export_pad_value is not None
        ):
            try:
                self._export_final_surface_alpha_wrap(
                    final_analysis_component=final_analysis_component,
                    field=export_field,
                    level_set_value=export_level_set_value,
                    pad_value=export_pad_value,
                )
            except Exception as e:
                self._print_out(f"Warning: failed to export final alpha-wrap surface meshes. Error: {e}", "Minimal")
        # Background periodic exports are always collected, also when no final export ran.
        try:
            self._finish_surface_jobs()
        except Exception as e:
            self._print_out(f"Warning: could not collect the background surface exports. Error: {e}", "Minimal")

        # Compliance of the geometry that is actually EXPORTED (phi >= level
        # set, i.e. design density >= 0.5, made fully solid; everything else
        # void). For a converged, binary design this equals the last logged
        # compliance. For a run that stopped early it can be very different:
        # a grey link below the level set carries load in the analysis model
        # but is absent from the exported geometry, and on the house
        # cantilever the exported design of the previous revision was 10-20
        # times less stiff than the logged compliance for most of the run
        # (docs/theory.md §4.7). Density backends only -- Conformal already
        # analyses exactly the exported geometry.
        self.final_exported_compliance = None
        if self._analysis_field_type == "phase":
            # Conformal: the final mesh with the full modulus in every tet
            # (the model itself when there is no material interpolation).
            c_exp = getattr(self, "_final_phase_exported_compliance", None)
            if c_exp is not None and np.isfinite(c_exp):
                self.final_exported_compliance = float(c_exp)
                c_mod = getattr(self, "_final_phase_model_compliance", None)
                self._print_out(f"Compliance of the exported design (the final mesh, full modulus): "
                                f"{self.final_exported_compliance:.6e} J"
                                + (f" ({self.final_exported_compliance / c_mod:.3f} x the final model's)"
                                   if c_mod else "") + ".", "Minimal")
        # Buckling load factor of the same exported geometry (buckling runs on
        # a density backend only). The logged critical_load_factor is that of
        # the analysed model rho_m, which a design that is not yet binary
        # overstates: on the test column with the linear interpolation (p = 1)
        # the exported geometry buckled at 1.77 against 2.00 logged
        # (docs/theory.md §6.7). None = not evaluated ("N/A" in the stats).
        self.final_exported_critical_load_factor = None
        _pb = getattr(self, "_final_phase_exported_buckling", None) if self._analysis_field_type == "phase" else None
        if _pb is not None:
            # Conformal: the final mesh at full modulus (analysed with the final model above).
            self.final_exported_critical_load_factor = float(_pb.critical_load_factor)
            self._exported_buckling = _pb
            _lam = self.final_exported_critical_load_factor
            self._print_out(f"Critical buckling load factor of the exported design (the final mesh, full modulus"
                            f"{'' if self.buckling_enabled else '; constraint disabled, reported only'}): "
                            f"{self._fmt_load_factor(_lam)}"
                            + (f" (analysed model: "
                               f"{self._fmt_load_factor(getattr(self, 'last_valid_critical_load_factor', float('nan')))}, "
                               f"target {self.buckling_safety_factor})" if self.buckling_enabled else "") + ".",
                            "Minimal")
        # Volume fraction of the exported geometry. It can sit below the
        # volume fraction the constraint sees, mean(rho): void voxels that the
        # piecewise multiplier has not yet pulled all the way back to V carry
        # a small density that counts towards G1 but is not exported
        # (docs/theory.md §5.5).
        self.final_exported_volume_fraction = float(np.count_nonzero(self._solid_mask(final_phi))) / final_phi.size
        self._print_out(f"Volume fraction of the exported design: {self.final_exported_volume_fraction:.4f} "
                        f"(target {self.target_volume}, mean design density {float(np.mean(self._rho(final_phi))):.4f}; "
                        f"exported = {100 * self.final_exported_volume_fraction / max(self.target_volume, 1e-300):.1f}% of the target, "
                        f"void film carries {getattr(self, 'void_film_fraction', float('nan')):.4f} of the domain volume).",
                        "Minimal")
        try:
            # Density backends analyse the exported geometry on a component of
            # its own, saved as data/outer_final/exported/..._iteration_data.vtr
            # (stress, displacement and buckling modes of the part as exported).
            is_density = (self._analysis_field_type == "density")
            probe = self.analysis_component_factory(
                iter_name=os.path.join("outer_final", "exported") if is_density else "test",
                printouts="Silent", data_dir=self.analysis_dir, save_meshes=is_density)
            if probe.field_type == "density":
                binary = np.where(self._solid_mask(final_phi), 1.0, self.void_density).astype(np.float32)
                exported_sed, _ = probe.run_analysis(field=binary, level_set_value=None)
                if getattr(probe, "_grid_shape", None) is not None:
                    probe._pending_cell_data["Design_Density"] = np.where(
                        self._solid_mask(final_phi), 1.0, self.void_density).astype(np.float32)
                exported_mode_shapes = None
                exported = getattr(probe, "compliance", None)
                if exported is not None and np.isfinite(exported):
                    self.final_exported_compliance = float(exported)
                    ratio = self.final_exported_compliance / self.compliance if self.compliance else float('nan')
                    self._print_out(f"Compliance of the exported (thresholded) design: "
                                    f"{self.final_exported_compliance:.6e} J "
                                    f"({ratio:.3f} x the last optimisation compliance).", "Minimal")
                if (((self.buckling_enabled and getattr(self, 'buckling_ever_evaluated', False))
                     or self.buckling_report_exported)
                        and hasattr(probe, "run_buckling")):
                    # Same probe instance: run_buckling reuses the primal
                    # solution and the stiffness matrix of run_analysis above.
                    try:
                        exported_buckling = probe.run_buckling(
                            num_modes=self.buckling_num_modes,
                            aggregation_p=self.buckling_aggregation_p,
                            prev_modes=None,
                            min_energy_fraction=self.buckling_min_energy_fraction,
                            mode_continuity_warning_threshold=self.buckling_mode_continuity_warning_threshold,
                            min_density_for_geometric_stiffness=self.buckling_min_density_for_geometric_stiffness,
                        )
                        self.final_exported_critical_load_factor = float(exported_buckling.critical_load_factor)
                        self._exported_buckling = exported_buckling
                        if exported_buckling.has_modes:
                            exported_mode_shapes = exported_buckling.mode_shapes
                        lam_exp = self.final_exported_critical_load_factor
                        if self.buckling_enabled:
                            self._print_out(f"Critical buckling load factor of the exported design: "
                                            f"{self._fmt_load_factor(lam_exp)} (analysed model: "
                                            f"{self._fmt_load_factor(getattr(self, 'last_valid_critical_load_factor', float('nan')))}, "
                                            f"target {self.buckling_safety_factor}).", "Minimal")
                        else:
                            self._print_out(f"Critical buckling load factor of the exported design "
                                            f"(constraint disabled, reported only): "
                                            f"{self._fmt_load_factor(lam_exp)}.", "Minimal")
                        if (self._buckling_target_met() and np.isfinite(lam_exp)
                                and lam_exp < self.buckling_safety_factor / (1.0 + self.outer_loop_convergence_threshold)):
                            self._print_out(
                                "WARNING: the analysed model meets the buckling target but the exported "
                                "geometry does not. Its grey material carries compressive stiffness in the "
                                "model that thresholding removes: check the grey fraction and the "
                                "penalisation power (docs/theory.md §4.7, §6.7).", "Minimal")
                    except Exception as e:
                        self._print_out(f"Warning: could not run the buckling analysis of the exported "
                                        f"design. Error: {e}", "Minimal")
                # Keep the exported design's fields for final_fields.npz, then
                # write its .vtr (with up to four buckling mode magnitudes).
                self._exported_fields = self._capture_pending_fields(probe)
                if self._exported_fields is not None:
                    self._exported_fields["cell"]["Exported_SED"] = np.asarray(exported_sed, dtype=np.float32)
                    if exported_mode_shapes is not None:
                        try:
                            n_modes = min(4, exported_mode_shapes.shape[1])
                            self._exported_fields["modes"] = {
                                k: np.asarray(v, dtype=np.float32) for k, v in
                                probe._map_mode_shapes_to_grid(exported_mode_shapes[:, :n_modes]).items()}
                        except Exception as e:
                            self._print_out(f"Warning: could not map the exported design's buckling modes "
                                            f"onto the grid. Error: {e}", "Minimal")
                try:
                    written = probe.finalize_iteration_output(
                        field_data={"exported": 1,
                                    "compliance": float(self.final_exported_compliance or float('nan')),
                                    "critical_load_factor": float(self.final_exported_critical_load_factor
                                                                  if self.final_exported_critical_load_factor
                                                                  is not None else float('nan'))},
                        mode_shapes=exported_mode_shapes, max_modes_to_save=4)
                    if written:
                        self._print_out(f"Saved the exported design's analysis to {written}.", "Normal")
                except Exception as e:
                    self._print_out(f"Warning: failed to write the exported design's data file. Error: {e}",
                                    "Minimal")
        except Exception as e:
            self._print_out(f"Warning: could not analyse the exported design. Error: {e}", "Minimal")

        # Export the final statistics
        self._save_outcome_stats()

        # Research outputs (docs/theory.md §8.4.1-§8.4.3): the final fields,
        # a JSON record of the run, the threshold sweep, a ParaView time
        # series, and the plots, summary and report of the postprocess section. Each
        # step is independent; a failure is reported and never affects the
        # outputs written above.
        self._write_research_outputs(final_phi)

        return final_phi

    # ------------------------------------------------------------------
    # Parameters
    # ------------------------------------------------------------------

    @classmethod
    def common_parameter_defaults(cls) -> dict:
        """{name: default} of COMMON_PARAMETERS, read from `initialise_common_parameters`."""
        sig = inspect.signature(BaseOptimiser.initialise_common_parameters)
        return {n: sig.parameters[n].default for n in cls.COMMON_PARAMETERS}

    @classmethod
    def parameter_names(cls) -> set:
        """Every keyword the class's parameter initialiser accepts (used by optimise.py's name check)."""
        return set(cls.COMMON_PARAMETERS)

    def _write_parameters_csv(self):
        """parameters.csv in one format for every method: one header row and one value row."""
        all_params = {"optimiser_optimisation_name": self.optimisation_name,
                      "optimiser_domain_dims": self.domain_dims}
        for key, value in self._optimiser_params_dict.items():
            all_params[f"optimiser_{key}"] = value
        if hasattr(self.analysis_component_factory, "keywords"):
            func = getattr(self.analysis_component_factory, "func", None)
            if func is not None and hasattr(func, "__name__"):
                all_params["analysis_class"] = func.__name__
            for key, value in self.analysis_component_factory.keywords.items():
                all_params[f"analysis_{key}"] = value
        try:
            with open(os.path.join(self.optimisation_name, "parameters.csv"), "w") as f:
                f.write(",".join(all_params.keys()) + "\n")
                f.write(",".join(str(v).replace(",", "_").replace(" ", "") for v in all_params.values()) + "\n")
        except Exception as e:
            self._print_out(f"Error: Failed to save parameters.csv. Reason: {e}", "Minimal")

    # ------------------------------------------------------------------
    # What a method provides
    # ------------------------------------------------------------------

    def _setup_phi(self):  # pragma: no cover - overridden
        raise NotImplementedError("a method builds its initial design in _setup_phi")

    def _outer_loop(self):  # pragma: no cover - overridden
        raise NotImplementedError("a method runs its iteration in _outer_loop")

    def _variables_na(self) -> tuple:
        """variables.csv columns that do not apply to the method (written as N/A); none for PCLS."""
        return ()


class BaselineOptimiser(BaseOptimiser):
    """
    Base class of the comparison baselines (docs/theory.md §12.1). Not used on
    its own: SIMP_Optimiser and LevelSet_Optimiser implement the method.

    A subclass provides
      METHOD, METHOD_LABEL      the problem-file name ("simp") and a label;
      _set_method_parameters()  its own keyword parameters, all prefixed with
                                METHOD + "_" (the parser flattens the
                                optimiser.<method> block onto these names);
      _initialise_design()      sets self.phi (and its own state);
      _initialise_loop_state()  multipliers etc. before the first iteration;
      _iterate(outer, inner)    one analysis and one design update; returns
                                False if the analysis failed;
      _settle_tests()           the method's own stopping tests;
    and optionally _end_of_outer_iteration(), _variables_na(),
    _method_record(), _method_final_fields(), _field_data_extras(),
    _extra_cell_data().
    """

    METHOD = None
    METHOD_LABEL = "baseline"

    # PCLS-only state that the inherited variables.csv writer reads without a
    # guard; written as N/A in a baseline's rows (see _save_variables).
    _VARIABLES_NA = ("lambda2_min", "lambda2_mean", "lambda2_max", "mu2")

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        if self._analysis_field_type != "density":
            raise ValueError(
                f"The {self.METHOD_LABEL} optimiser needs a density backend ('reg' or 'cg'): it hands the "
                f"analysis a stiffness density per voxel. The configured backend "
                f"'{self._analysis_class_name}' (field_type '{self._analysis_field_type}') meshes a level "
                f"set instead. Use analysis.type 'reg' or 'cg' (docs/theory.md §12.1).")

    # ------------------------------------------------------------------
    # Parameters
    # ------------------------------------------------------------------

    @classmethod
    def method_parameter_defaults(cls) -> dict:
        """{name: default} of the method's own (prefixed) parameters."""
        sig = inspect.signature(cls._set_method_parameters)
        return {n: p.default for n, p in sig.parameters.items() if n != "self"}

    @classmethod
    def parameter_names(cls) -> set:
        """Every keyword initialise_optimisation_parameters accepts (used by optimise.py's name check)."""
        return set(cls.COMMON_PARAMETERS) | set(cls.method_parameter_defaults())

    def _set_method_parameters(self, **kwargs):  # pragma: no cover - overridden
        raise NotImplementedError

    def _effective_method_parameters(self) -> dict:
        """Parameters the method set itself (overriding the given or default ones), recorded in parameters.csv."""
        return {}

    def initialise_optimisation_parameters(self, **kwargs):
        """
        Sets the common parameters (COMMON_PARAMETERS: the same names,
        meanings, defaults and validation for every optimiser, by
        `initialise_common_parameters`) and the method's
        own, prefixed parameters (`_set_method_parameters`), then records the
        parameters actually used in parameters.csv and run_info.json.

        Raises ValueError for a name the method does not accept, and for
        solid_value / void_value other than 2 / 1: the encodings of eq. (12.5)
        rely on S - V = 1 with rho(V) = c1 (docs/theory.md §12.1).
        """
        valid = self.parameter_names()
        unknown = sorted(set(kwargs) - valid)
        if unknown:
            raise ValueError(f"Unknown {self.METHOD_LABEL} optimiser parameter(s) {unknown}. "
                             f"Valid names: {', '.join(sorted(valid))}.")
        common = {k: kwargs[k] for k in self.COMMON_PARAMETERS if k in kwargs}
        method = {k: v for k, v in kwargs.items() if k not in self.COMMON_PARAMETERS}
        S = common.get("solid_value", 2.0)
        V = common.get("void_value", 1.0)
        if float(S) != 2.0 or float(V) != 1.0:
            raise ValueError(f"The {self.METHOD_LABEL} optimiser stores its design as phi in [1, 2] "
                             f"(docs/theory.md eq. 12.5) and needs solid_value = 2 and void_value = 1 "
                             f"(got {S} and {V}).")

        # The common part: the same names, defaults and validation as every other optimiser.
        self.initialise_common_parameters(**common)
        self._set_method_parameters(**method)

        defaults = self.common_parameter_defaults()
        defaults.update(self.method_parameter_defaults())
        params = {"method": self.METHOD}
        for name in list(self.COMMON_PARAMETERS) + list(self.method_parameter_defaults()):
            params[name] = kwargs.get(name, defaults[name])
        params.update(self._effective_method_parameters())
        self._optimiser_params_dict = params
        self._write_parameters_csv()

        self._print_out(f"{self.METHOD_LABEL} optimiser parameters:\n" +
                        "\n".join(f"  {k}: {v}" for k, v in params.items()), "Verbose")

    # ------------------------------------------------------------------
    # Initial design
    # ------------------------------------------------------------------

    def _setup_phi(self):
        """
        Resolves the random seed (recorded in run_info.json like PCLS's, although
        neither baseline draws random numbers: both start from a deterministic
        design) and builds the method's initial design (`_initialise_design`).
        """
        if getattr(self, "random_seed", None) is None:
            self.random_seed = int(np.random.SeedSequence().entropy % (2 ** 32))
        self._initial_phi = self._initial_design_phi()      # None unless initial_design was given
        self._initialise_design()
        self._print_out(f"Initial {self.METHOD_LABEL} design: phi min={self.phi.min():.6f}, "
                        f"mean={self.phi.mean():.6f}, max={self.phi.max():.6f}", "Verbose")

    # ------------------------------------------------------------------
    # One analysis (shared by both methods)
    # ------------------------------------------------------------------

    def _analyse(self, model_rho: np.ndarray, current_rho: np.ndarray, outer_iter: int, inner_iter: int):
        """
        Analyses the current design (self.phi, design density `current_rho`)
        with the stiffness density `model_rho`, the part of PCLS's _inner_loop
        that runs and checks one density-backend analysis (docs/theory.md
        §5.4, §8.4.1): mesh saving on the save_frequency cadence, CG warm start,
        the output check, the exact compliance, the exported-design compliance
        every exported_analysis_every iterations, the standardised .vtr and the
        periodic surface export.

        Sets self.F_u, self.compliance, self.g1, self.volume_fraction and the
        per-iteration analysis columns. Returns the analysis component, or
        None after a failure (recorded as ANALYSIS_FAILED; self.phi is set back
        to the last design whose analysis was valid, on which the run ends).
        """
        phi = self.phi
        iter_dir = os.path.join(f"outer_{outer_iter:03d}", f"inner_{inner_iter:03d}")
        save_this_iter = (self.total_iterations % self.save_frequency == 0)
        iter_name = iter_dir if save_this_iter else "test"
        factory_kwargs = {"iter_name": iter_name, "printouts": self.printouts,
                          "data_dir": self.analysis_dir, "save_meshes": save_this_iter}
        if self._warm_start_cg_enabled:
            factory_kwargs["warm_start_state"] = self._cg_warm_state

        self.g1 = float(np.mean(current_rho)) - self.target_volume
        self.volume_fraction = float(np.mean(current_rho))
        self._print_out(f"Iteration {self.total_iterations} (outer {outer_iter}, inner {inner_iter}): "
                        f"volume fraction {self.volume_fraction:.6f}", "Normal")
        try:
            component = self.analysis_component_factory(**factory_kwargs)
            self.F_u, _ = component.run_analysis(field=np.asarray(model_rho), level_set_value=None)
            self._check_analysis_output(component)
        except Exception as e:
            message = (f"the analysis failed at iteration {self.total_iterations} (outer {outer_iter}, "
                       f"inner {inner_iter}): {type(e).__name__}: {e}")
            self._save_variables(phi, None, None)
            self._record_failure("ANALYSIS_FAILED", message, traceback.format_exc(), failed_phi=phi)
            last = getattr(self, "_last_valid_phi", None)
            if last is not None:
                self.phi = last.copy()
                self._restore_last_valid_method_state()
            return None
        self._last_analysed_phi = phi.copy()

        # mesh_volume_fraction: the exported geometry's voxel count relative to
        # the first iteration's, counted from phi as for PCLS (docs/theory.md §8.4).
        mesh_volume = float(np.count_nonzero(self._solid_mask(phi)))
        if self.total_iterations == 0:
            self.initial_mesh_volume = mesh_volume
        self.mesh_volume_fraction = (mesh_volume / self.initial_mesh_volume
                                     if getattr(self, "initial_mesh_volume", 0) else float("nan"))
        self.solver_iterations = getattr(component, "solver_iterations", None)
        self.solver_converged = getattr(component, "solver_converged", None)
        self.solver_final_residual = getattr(component, "solver_final_residual", None)
        self.warm_start_attempted = getattr(component, "warm_start_attempted", None)
        self.warm_start_bailed_to_cold = getattr(component, "warm_start_bailed_to_cold", None)
        self.compliance = self._compute_compliance(self.F_u, current_rho, analysis_component=component)
        self.analysis_valid = True
        self._last_valid_phi = phi.copy()
        self._last_valid_F_u = self.F_u
        self._last_valid_method_state = self._method_state()
        if not np.isfinite(self.compliance) or self.compliance <= 0.0:
            self.nonphysical_analyses += 1
            if self.nonphysical_analyses == 1 or self.nonphysical_analyses % 25 == 0:
                self._print_out(f"WARNING: the analysis returned a non-physical compliance "
                                f"({self.compliance:.4e}) at iteration {self.total_iterations}.", "Minimal")

        # Compliance of the geometry this design would export (eq. 8.4i).
        self.exported_compliance = None
        self.exported_analysis_time = None
        if self.exported_analysis_every > 0 and self.total_iterations % self.exported_analysis_every == 0:
            t_exp = timer()
            try:
                probe = self.analysis_component_factory(iter_name="test", printouts="Silent",
                                                        data_dir=self.analysis_dir, save_meshes=False)
                probe.run_analysis(field=np.where(self._solid_mask(phi), 1.0, self.void_density).astype(np.float32),
                                   level_set_value=None)
                c_exp = getattr(probe, "compliance", None)
                if c_exp is not None and np.isfinite(c_exp):
                    self.exported_compliance = float(c_exp)
            except Exception as e:
                self._print_out(f"Warning: the analysis of the exported design failed this iteration. "
                                f"Error: {e}", "Minimal")
            self.exported_analysis_time = timer() - t_exp

        # Linear buckling (G3), on the same component as the analysis above.
        mode_shapes = None
        self.buckling_results = None
        if self.buckling_enabled:
            ok, mode_shapes = self._evaluate_buckling(component, outer_iter, inner_iter)
            if not ok:
                return None

        # Standardised per-iteration output (docs/theory.md §7.6).
        field_data = {"outer_iter": outer_iter, "inner_iter": inner_iter,
                      "total_iterations": self.total_iterations, "compliance": self.compliance,
                      "volume_fraction": self.volume_fraction, "mesh_volume_fraction": self.mesh_volume_fraction,
                      "penalisation_power": self.current_penalisation_power, "method": self.METHOD}
        field_data.update(self._field_data_extras())
        if self.buckling_enabled:
            field_data["lambda3"] = getattr(self, "lambda3", 0.0)
            field_data["mu3"] = getattr(self, "mu3", 0.0)
            if self.safety_factor_enforced is not None:
                field_data["buckling_safety_factor"] = self.safety_factor_enforced
            if self.buckling_ever_evaluated:
                field_data["critical_load_factor"] = self.critical_load_factor
                field_data["aggregated_load_factor"] = self.aggregated_load_factor
                field_data["n_buckling_modes"] = self.n_buckling_modes
        if save_this_iter:
            try:
                if getattr(component, "_grid_shape", None) is not None:
                    component._pending_cell_data["Design_Density"] = np.asarray(current_rho, dtype=np.float32)
                    component._pending_cell_data["Phi"] = np.asarray(phi, dtype=np.float32)
                    for k, v in self._extra_cell_data().items():
                        component._pending_cell_data[k] = np.asarray(v, dtype=np.float32)
            except Exception:
                pass
        try:
            written = component.finalize_iteration_output(field_data=field_data, mode_shapes=mode_shapes)
            if written:
                self._saved_iteration_files.append((self.total_iterations, written))
        except Exception as e:
            self._print_out(f"Warning: failed to finalize this iteration's standardised output. Error: {e}",
                            "Minimal")
        if save_this_iter:
            try:
                self._export_final_surface_alpha_wrap(final_analysis_component=component,
                                                      field=np.asarray(current_rho), level_set_value=0.5,
                                                      pad_value=0.0, tier="periodic")
            except Exception as e:
                self._print_out(f"Warning: failed to export this iteration's surface mesh. Error: {e}", "Minimal")
        return component

    # ------------------------------------------------------------------
    # Linear buckling (G3): shared by both methods (docs/theory.md §6, §12.6)
    # ------------------------------------------------------------------

    def _initialise_buckling_state(self):
        """The buckling state (what PCLS_Optimiser._outer_loop sets up too), for a run with buckling enabled."""
        self.lambda3 = self.buckling_lambda3_initial
        self.mu3 = self.buckling_mu3_initial
        self.lambda3_tilde = 0.0
        self.g3 = 0.0
        self.last_valid_g3 = 0.0
        self.last_valid_critical_load_factor = float("nan")
        self.last_buckling_evaluation_valid = False
        self.buckling_term = np.zeros(self.domain_dims, dtype=np.float32)
        self.applied_buckling_term = np.zeros(self.domain_dims, dtype=np.float32)
        self.prev_buckling_modes = None
        self.buckling_evaluated_this_outer = False
        self.buckling_valid_evaluated = False
        self.buckling_unknown_evaluations = 0
        self.buckling_status = None
        self.critical_load_factor = float("inf")
        self.aggregated_load_factor = float("inf")
        self.n_buckling_modes = 0
        self.buckling_ever_evaluated = False
        self.buckling_results = None
        self.buckling_time = 0.0
        self.safety_factor_enforced = None
        self.previous_g3_violation = float("inf")
        # run_info.json's buckling.evaluations (PCLS counts its own; a baseline's was always null).
        self.buckling_evaluations_done = 0

    def _buckling_active(self, outer_iter: int) -> bool:
        """True from outer iteration buckling_activate_after_outer_iter (eq. 6.24)."""
        return self.buckling_enabled and outer_iter >= self.buckling_activate_after_outer_iter

    def _evaluate_buckling(self, component, outer_iter: int, inner_iter: int):
        """
        One buckling evaluation on the analysis `component` that has just run
        (docs/theory.md §6.5-6.7, §12.6): the logic of PCLS_Optimiser's
        inline block built from its inherited pieces -- activation,
        `_current_buckling_safety_factor`, the evaluate_every cadence and the
        analysis_valid guard, `run_buckling`, `_G3`, `_log_buckling`, and the
        three-way status handling that never reads an UNKNOWN evaluation
        (every candidate mode rejected by the spurious-mode filter) as safe.

        Sets the buckling attributes of PCLS (g3, last_valid_g3,
        last_valid_critical_load_factor, critical_load_factor,
        aggregated_load_factor, n_buckling_modes, buckling_status,
        buckling_ever_evaluated, buckling_valid_evaluated, lambda3_tilde,
        buckling_time, safety_factor_enforced) and:

          self.buckling_results   the BucklingResults of THIS iteration's
                                  evaluation when it was valid (status OK or
                                  NO_COMPRESSIVE_MODE), else None -- the
                                  method builds its driver from it;
          self.buckling_active_now True when the constraint is active this
                                  iteration, whether or not it was evaluated.

        Returns (ok, mode_shapes): ok is False after a failed eigensolve
        (recorded as ANALYSIS_FAILED, self.phi set back to the last valid
        design, as _analyse does for a failed analysis); mode_shapes are the
        retained modes for the .vtr, or None.
        """
        self.buckling_results = None
        self.buckling_time = 0.0
        self.safety_factor_enforced = None
        self.buckling_active_now = self._buckling_active(outer_iter)
        mode_shapes = None
        if not self.buckling_active_now:
            # Inactive: contribute nothing.
            self.g3 = 0.0
            self.lambda3_tilde = 0.0
            self.buckling_term = np.zeros_like(self.phi, dtype=np.float32)
            self._print_out(f"  Buckling: inactive until outer iteration "
                            f"{self.buckling_activate_after_outer_iter}.", "Verbose")
            return True, None

        safety_factor = self._current_buckling_safety_factor(outer_iter)
        self.safety_factor_enforced = float(safety_factor)
        evaluate_now = (inner_iter % self.buckling_evaluate_every == 0) and self.analysis_valid
        if not evaluate_now:
            self._print_out(f"  Buckling: reusing the last evaluation (evaluate_every="
                            f"{self.buckling_evaluate_every}).", "Verbose")
            return True, None

        if not hasattr(component, "run_buckling"):
            raise ValueError(f"buckling_enabled=True, but the analysis component "
                             f"'{type(component).__name__}' has no run_buckling.")
        t_buckling = timer()
        try:
            results = component.run_buckling(
                num_modes=self.buckling_num_modes,
                aggregation_p=self.buckling_aggregation_p,
                prev_modes=self.prev_buckling_modes,
                min_energy_fraction=self.buckling_min_energy_fraction,
                mode_continuity_warning_threshold=self.buckling_mode_continuity_warning_threshold,
                min_density_for_geometric_stiffness=self._geometric_stiffness_threshold(),
                eigensolver_tol=self.buckling_eigensolver_tol,
            )
        except Exception as e:
            phi = self.phi
            self._save_variables(phi, None, None)
            self._record_failure(
                "ANALYSIS_FAILED",
                f"the buckling analysis failed at iteration {self.total_iterations} "
                f"(outer {outer_iter}, inner {inner_iter}): {type(e).__name__}: {e}",
                traceback.format_exc(), failed_phi=phi)
            last = getattr(self, "_last_valid_phi", None)
            if last is not None:
                self.phi = last.copy()
                self._restore_last_valid_method_state()
            return False, None
        self.buckling_time = timer() - t_buckling
        self._max_buckling_time = max(getattr(self, "_max_buckling_time", 0.0), self.buckling_time)
        self.buckling_evaluations_done = getattr(self, "buckling_evaluations_done", 0) + 1

        g3 = self._G3(results, safety_factor)
        self.critical_load_factor = results.critical_load_factor
        self.aggregated_load_factor = results.aggregated_load_factor
        self.n_buckling_modes = results.mu_values.size
        self.buckling_status = results.status
        self.buckling_ever_evaluated = True
        self._log_buckling(results, safety_factor, g3)
        if np.isnan(g3):
            # UNKNOWN: keep the last valid evaluation's state; this one neither
            # updates a multiplier nor counts towards convergence.
            self.buckling_unknown_evaluations += 1
            self.last_buckling_evaluation_valid = False
        else:
            self.g3 = g3
            self.last_valid_g3 = g3
            self.last_valid_critical_load_factor = results.critical_load_factor
            self.last_buckling_evaluation_valid = True
            if np.isfinite(self.mu3):
                self.lambda3_tilde = float(np.clip(self.lambda3 + g3 / self.mu3,
                                                   0.0, self.buckling_lambda3_clamp_max))
            else:
                # A method whose multiplier is not an augmented-Lagrangian one (SIMP's MMA
                # dual, mu3 = nan) sets lambda3_tilde itself after its update.
                self.lambda3_tilde = float(self.lambda3)
            self.buckling_evaluated_this_outer = True
            self.buckling_valid_evaluated = True
            self.buckling_results = results
        if results.has_modes:
            self.prev_buckling_modes = results.mode_shapes
            mode_shapes = results.mode_shapes
        return True, mode_shapes

    def _buckling_density_sensitivity(self, results: BucklingResults) -> np.ndarray:
        """
        s_e = d(mu_tilde)/d(rho_m,e) per voxel from the backend (dimensionless;
        the derivative of the aggregate with respect to the stiffness density
        it was given), checked against the design grid. Zeros for a result
        without modes.
        """
        s = np.asarray(results.sensitivity)
        if s.shape != tuple(self.domain_dims):
            raise ValueError(f"Buckling sensitivity shape {s.shape} does not match the design grid "
                             f"{tuple(self.domain_dims)}.")
        return s

    def _buckling_scale_factor(self, results: BucklingResults) -> float:
        """c = mu_max / mu_tilde of eq. (6.23): the constraint measure's value over the aggregate's, held fixed in the gradient."""
        if not results.has_modes or results.mu_tilde <= 0.0:
            return 1.0
        if self.buckling_constraint_measure == "critical":
            return float(np.max(results.mu_values) / results.mu_tilde)
        return 1.0

    def _update_buckling_multipliers(self):
        """
        Once per outer iteration, for a method that has an augmented-Lagrangian
        multiplier for G3 (the level set; SIMP's MMA has its own dual):
        PCLS's update (docs/theory.md eq. 4.8 and 6.26) on the last VALID
        evaluation, only after an outer iteration that had one.

            lambda3 <- clip(lambda3 + g3 / mu3, 0, lambda3_max)
            mu3     <- max(eta3 mu3, mu3_floor)  while violated and not shrinking by 25 %.
        """
        if not self.buckling_enabled:
            return
        if getattr(self, "buckling_evaluated_this_outer", False):
            g3 = self.last_valid_g3
            self.lambda3 = min(max(0.0, self.lambda3 + g3 / self.mu3), self.buckling_lambda3_clamp_max)
            violation = max(g3, 0.0)
            previous = getattr(self, "previous_g3_violation", float("inf"))
            if violation > 0.0 and violation > 0.75 * previous:
                self.mu3 = max(self.mu3 * self.buckling_mu3_decay_factor, self.buckling_mu3_floor)
            self.previous_g3_violation = violation
            self._print_out(f"  Lambda3: {self.lambda3:.6f}, Mu3: {self.mu3:.6e}", "Verbose")
        self.buckling_evaluated_this_outer = False

    # ------------------------------------------------------------------
    # The loop
    # ------------------------------------------------------------------

    def _outer_loop(self) -> np.ndarray:
        """
        The baseline driver (docs/theory.md §12.1, eq. 12.9). Iteration n is
        outer iteration n // inner_loop_max_iters, inner iteration n % N_in;
        the budget, the time limit and run_until_converged work exactly as for
        PCLS (§5.4, §5.7). Each iteration analyses the current design and
        updates it (`_iterate`); then the shared diagnostics are taken and the
        row of variables.csv written. The run converges once the method's own
        stopping tests (`_settle_tests`) have held for `_settle_window`
        consecutive iterations, the exported design is connected (if
        convergence_require_connected) and its volume is within
        convergence_exported_volume_tol of V0 -- the PCLS rules with no
        reconnection resets, since neither baseline has a lambda2 to release
        (§5.5.1, §5.6): a disconnected design keeps iterating, one that fails
        only the volume check is accepted with converged_with_volume_shortfall.
        """
        self._print_out(f"Starting {self.METHOD_LABEL} optimiser...", "Minimal")
        t0 = timer()
        self.total_iterations = 0
        self.converged = False
        self.time_limit_reached = False
        self.stop_requested = False
        self.analysis_failed = False
        self.failure = None
        self.termination_reason = None
        self.settled_count = 0
        self.settled_disconnected_events = 0
        self.settled_volume_shortfall_events = 0
        self.converged_with_volume_shortfall = False
        self.reconnection_resets_done = 0
        self.reconnection_reset_iterations = []
        self.nonphysical_analyses = 0
        self.invalid_analyses = 0
        self.rejected_steps = 0
        self.retried_analysis_errors = 0
        self.min_step_scale = 1.0
        self.analysis_valid = True
        self._last_valid_phi = None
        self._last_valid_F_u = None
        self._last_valid_method_state = None
        self._step_scale = 1.0
        # max_net_force of _record_iteration_diagnostics is max|dphi| / _dt_step:
        # for a baseline the largest change of phi in one iteration.
        self._dt_step = 1.0
        self._reset_settle_windows()
        self.compliance_trend = None
        self.net_phase_switches = None
        self.grey_fraction = None
        self.phase_switches = None
        self.connectivity = None
        self._iteration_times = []
        unlimited = self.run_until_converged or self.time_limit_hours is not None
        self.iteration_budget = None if unlimited else self.outer_loop_max_iters * self.inner_loop_max_iters
        self.budget_exhausted = False
        if self.buckling_enabled:
            self._initialise_buckling_state()
        for name in self._EXTRA_VARIABLE_COLUMNS:
            # lambda3_tilde and aggregated_load_factor are buckling state set by
            # _initialise_buckling_state (as in PCLS_Optimiser); the rest start unset.
            if self.buckling_enabled and name in ("lambda3_tilde", "aggregated_load_factor"):
                continue
            setattr(self, name, None)
        self._diagnostics = RunDiagnostics(self.optimisation_name, self.design_mask,
                                                 history_every=self.history_every)
        self._saved_iteration_files = []
        self.final_surface_info = None
        self.periodic_surface_infos = []
        self._surface_jobs = []
        self._initialise_loop_state()
        self._record_iteration_diagnostics(self.phi, None, outer_iter=0, inner_iter=None)
        self._save_variables(self.phi, None, None, overwrite=True)

        n_in = max(1, int(self.inner_loop_max_iters))
        outer_iter = 0
        while True:
            outer_iter, inner_iter = divmod(self.total_iterations, n_in)
            if self.iteration_budget is not None and self.total_iterations >= self.iteration_budget:
                self.budget_exhausted = True
                self.termination_reason = "MAX_ITERATIONS"
                break
            if inner_iter == 0 and self.total_iterations > 0:
                self._end_of_outer_iteration(outer_iter - 1)
            if self._stop_file_requested():
                self.time_limit_reached = True
                self.termination_reason = "STOP_FILE"
                break
            if self._time_limit_would_be_exceeded():
                self.time_limit_reached = True
                self.termination_reason = "TIME_LIMIT"
                self._print_out(f"Time limit of {self.time_limit_hours:.3f} h reached "
                                f"({self._elapsed() / 3600.0:.3f} h elapsed).", "Minimal")
                break
            phi_old = self.phi.copy()
            t1 = timer()
            ok = self._iterate(outer_iter, inner_iter)
            if not ok:
                self.termination_reason = "ANALYSIS_FAILED"
                break
            t_analysis = getattr(self, "_analysis_time", None)
            phi_new = self.phi
            # PCLS's net push (phi^{n+1} - phi^n) / dt with dt = 1 (eq. 12.8):
            # written to final_fields.npz and shown by the postprocess section.
            self.net_push = (phi_new - phi_old).astype(np.float32)
            self.phase_switches = self._phase_switches(phi_new, phi_old)
            self.grey_fraction = self._grey_fraction(phi_new)
            if self._diag_due():
                self._refresh_connectivity(phi_new)
            else:
                self._conn_fresh = False
            settle_tests = list(self._settle_tests(phi_new, phi_old)) + [("analysis", self.analysis_valid)]
            if self.buckling_enabled:
                settle_tests.append(("buckling", self._buckling_converged(outer_iter)))
            settled_now = all(passed for _, passed in settle_tests)
            self.settle_failures = ";".join(name for name, passed in settle_tests if not passed) or "none"
            self.settled_count = self.settled_count + 1 if settled_now else 0
            self._iteration_times.append(timer() - t1)
            self._record_iteration_diagnostics(phi_new, phi_old, outer_iter=outer_iter, inner_iter=inner_iter)
            self._save_variables(phi_new, timer() - t1, t_analysis)
            self._print_out(
                f"  compliance {self.compliance:.6e} J, G1 {self.g1:+.5f}, grey {self.grey_fraction:.4f}, "
                f"switches {self.phase_switches}, connected {self.connectivity['connected']}, "
                f"settled {self.settled_count}/{self._settle_window} ({self.settle_failures})", "Normal")

            if self.settled_count >= self._settle_window:
                if self.convergence_require_connected and not self._conn_fresh:
                    self._refresh_connectivity(phi_new)      # the cadence skipped it; the acceptance needs it
                if self.convergence_require_connected and not self.connectivity["connected"]:
                    self.settled_disconnected_events += 1
                    self.settled_count = 0
                    if self.settled_disconnected_events == 1 or self.settled_disconnected_events % 25 == 0:
                        self._print_out(
                            f"The design meets the {self.METHOD_LABEL} stopping tests but is not connected "
                            f"(load path connected: {self.connectivity['load_path_connected']}, detached "
                            f"{self.connectivity['detached_fraction']:.4f}); continuing to iterate "
                            f"({self.settled_disconnected_events} time(s) so far).", "Minimal")
                elif not self._exported_volume_ok():
                    self.settled_volume_shortfall_events += 1
                    self.converged = True
                    self.converged_with_volume_shortfall = True
                    self.termination_reason = "CONVERGED"
                    self.total_iterations += 1
                    self._print_out(
                        f"WARNING: accepted as converged although the exported volume is "
                        f"{100 * self.exported_volume_fraction / self.target_volume:.1f}% of the target "
                        f"(tolerance +-{100 * self.convergence_exported_volume_tol:.0f}%).", "Minimal")
                    break
                else:
                    self.converged = True
                    self.termination_reason = "CONVERGED"
                    self.total_iterations += 1
                    self._print_out(f"Converged after {self.total_iterations} iterations.", "Minimal")
                    break
            self.total_iterations += 1

        if self.termination_reason in ("ANALYSIS_FAILED", "TIME_LIMIT", "STOP_FILE"):
            self.outer_iterations_run = outer_iter + 1       # the outer iteration it stopped in
        else:
            self.outer_iterations_run = (self.total_iterations - 1) // n_in + 1 if self.total_iterations else 0
        if self.termination_reason == "ANALYSIS_FAILED":
            self.analysis_failed = True
            self._print_out(f"Stopped after {self.total_iterations} iterations because "
                            f"{(self.failure or {}).get('message', 'the analysis failed')}.", "Minimal")
        elif self.termination_reason == "MAX_ITERATIONS":
            self._print_out(f"Iteration budget ({self.iteration_budget} iterations) used up without "
                            f"convergence, in {timer() - t0:.2f} seconds.", "Minimal")
        elif self.termination_reason == "TIME_LIMIT":
            self._print_out(f"Stopped by the time limit after {self.total_iterations} iterations.", "Minimal")
        elif self.termination_reason == "STOP_FILE":
            self._print_out(f"Stopped by the STOP file after {self.total_iterations} iterations.", "Minimal")
        self._flush_diagnostics()
        if self.connectivity is not None and not self.connectivity["connected"]:
            self._print_out(f"WARNING: the final design is NOT connected (load path connected: "
                            f"{self.connectivity['load_path_connected']}, detached solid fraction: "
                            f"{self.connectivity['detached_fraction']:.4f}).", "Minimal")
        return self.phi

    # Hooks with defaults.
    _settle_window = 1

    def _method_state(self) -> dict:
        """
        Copies of the method's own design state that go with self.phi (SIMP's x
        and its filtered/projected fields, the level set's psi). Snapshotted at
        every valid analysis and restored with phi after a failed one, so that
        final_fields.npz and run_info.json describe one design (before the
        2026-10 review phi went back to the last valid design and x / psi stayed
        at the failed one).
        """
        return {}

    def _restore_last_valid_method_state(self):
        state = getattr(self, "_last_valid_method_state", None)
        for k, v in (state or {}).items():
            setattr(self, k, v.copy() if isinstance(v, np.ndarray) else v)

    def _initialise_loop_state(self):
        pass

    def _end_of_outer_iteration(self, outer_iter: int):
        pass

    def _field_data_extras(self) -> dict:
        return {}

    def _extra_cell_data(self) -> dict:
        return {}

    def _method_record(self) -> dict:
        return {}

    def _method_final_fields(self) -> dict:
        return {}

    # ------------------------------------------------------------------
    # Thin wrappers around writers that read PCLS-only state
    # ------------------------------------------------------------------

    def _variables_na(self) -> tuple:
        """variables.csv columns that do not apply to this method (written as N/A)."""
        return self._VARIABLES_NA

    def _write_run_info(self, final_phi):
        """BaseOptimiser._write_run_info, plus `method`, the method's own record, and N/A for PCLS-only state."""
        super()._write_run_info(final_phi)
        path = os.path.join(self.optimisation_name, "run_info.json")
        with open(path) as f:
            info = json.load(f)
        info["method"] = self.METHOD
        info["settle_window"] = int(self._settle_window)
        final = info.get("final", {})
        for key in ("lambda2_min", "lambda2_mean", "lambda2_max", "mu2"):
            final[key] = None
        if "mu1" in self._variables_na():
            final["mu1"] = None
        info["final"] = final
        info["method_record"] = self._method_record()
        with open(path, "w") as f:
            json.dump(self._jsonable(info), f, indent=2)

    def _write_final_fields(self, final_phi):
        """BaseOptimiser._write_final_fields, plus the method's own arrays (`_method_final_fields`)."""
        super()._write_final_fields(final_phi)
        extra = {k: v for k, v in self._method_final_fields().items() if v is not None}
        if not extra:
            return
        path = os.path.join(self.optimisation_name, "final_fields.npz")
        with np.load(path) as z:
            arrays = {k: z[k] for k in z.files}
        for k, v in extra.items():
            arrays[k] = np.asarray(v)
        np.savez_compressed(path, **arrays)
        self._print_out(f"Added the {self.METHOD_LABEL} fields {sorted(extra)} to {path}.", "Normal")


# ======================================================================================================================
# === former gpu_update.py: device helpers, _Reducer, DeviceField, _Base, DiagGPU, update_memory_bytes ===
# ======================================================================================================================
"""
GPU (numba-CUDA) ports of the per-iteration host work of the SIMP and PCLS optimisers
(docs/PLAN_phase1.md step 3B; formulas and the memory model in docs/theory.md §18).

What is ported, and where it plugs in (the CPU code stays as the bit-exact fallback; every GPU entry point is
guarded by `gpu_analysis.update_engine_choice`):

  SimpGPU    SIMP: density filter and its transpose (cone-kernel correlation), Heaviside projection, the
             compliance/volume sensitivities, the method of moving asymptotes (class MMAGPU: the p/q/alpha/beta
             set-up, the dual evaluations and the primal update are device kernels, only the m-dimensional dual
             maximisation, scipy L-BFGS-B on 2 variables, stays on the host), the update terms and the change
             statistics of the stopping rule.
  PclsGPU    PCLS: density, driver, winsorised normalisation (bisection on the device), stiffness / volume /
             piecewise terms, the explicit step, the AOS/TV smoothing (np.gradient diffusivity and three batched
             Thomas solves, one thread per line), clamping, net push and the phase-switch / grey counts.
  DiagGPU    the per-iteration diagnostics of baseline_optimiser.py: exported solid, volume, void film,
             non-discreteness, band and step statistics, Euler characteristic, interface faces, 26- and
             6-connected component labelling (union-find with pointer jumping) for the connectivity report,
             the cavity count and the nucleation counts. A nucleation event that must be reported (rare) is
             written by the unchanged host routine on the downloaded masks.

Fields stay on the device between the analysis result (F_u, uploaded once: the analysis hands it over as a host
array) and the update; the arrays the host still reads are downloaded once (x_phys / phi) or lazily (DeviceField:
a host-array look-alike that downloads on first use and answers min / max / mean from the device).

numba-CUDA only (no CuPy): imports and runs on Windows and on small cards. Reductions are two-stage with a fixed
block assignment and a fixed-order tree, so a run is repeatable.
"""

import math

import numpy as np
from numba import cuda, float32, float64, int32, int64  # noqa: F401

BS = 256                 # threads per block
MAXB = 2048              # blocks of a reduction
NSLOT = 32               # scalars one reduction returns
_BIG = 1.0e300


def available():
    """True when numba finds a usable CUDA device."""
    try:
        return bool(cuda.is_available())
    except Exception:
        return False


def _nb(n):
    return int(min(MAXB, max(1, (int(n) + BS - 1) // BS)))


# ============================================================================ device helpers
@cuda.jit(device=True)
def _mx(a, b):
    return a if a > b else b


@cuda.jit(device=True)
def _mn(a, b):
    return a if a < b else b


@cuda.jit(device=True)
def _bsum(sh, v):
    t = cuda.threadIdx.x
    sh[t] = v
    cuda.syncthreads()
    s = BS // 2
    while s > 0:
        if t < s:
            sh[t] += sh[t + s]
        cuda.syncthreads()
        s //= 2
    r = sh[0]
    cuda.syncthreads()
    return r


@cuda.jit(device=True)
def _bmax(sh, v):
    t = cuda.threadIdx.x
    sh[t] = v
    cuda.syncthreads()
    s = BS // 2
    while s > 0:
        if t < s:
            if sh[t + s] > sh[t]:
                sh[t] = sh[t + s]
        cuda.syncthreads()
        s //= 2
    r = sh[0]
    cuda.syncthreads()
    return r


@cuda.jit(device=True)
def _bmin(sh, v):
    t = cuda.threadIdx.x
    sh[t] = v
    cuda.syncthreads()
    s = BS // 2
    while s > 0:
        if t < s:
            if sh[t + s] < sh[t]:
                sh[t] = sh[t + s]
        cuda.syncthreads()
        s //= 2
    r = sh[0]
    cuda.syncthreads()
    return r


@cuda.jit
def _k_combine(part, nblk, ops, out):
    """out[s] = reduce(part[s, :nblk]) with ops[s] 0 = sum, 1 = max, 2 = min; one block per slot."""
    sh = cuda.shared.array(BS, float64)
    s = cuda.blockIdx.x
    t = cuda.threadIdx.x
    op = ops[s]
    if op == 0:
        acc = 0.0
    elif op == 1:
        acc = -_BIG
    else:
        acc = _BIG
    i = t
    while i < nblk:
        v = part[s, i]
        if op == 0:
            acc += v
        elif op == 1:
            if v > acc:
                acc = v
        else:
            if v < acc:
                acc = v
        i += BS
    if op == 0:
        r = _bsum(sh, acc)
    elif op == 1:
        r = _bmax(sh, acc)
    else:
        r = _bmin(sh, acc)
    if t == 0:
        out[s] = r


class _Reducer:
    """Partial-sum buffer and the second stage of every reduction (one device-to-host copy of NSLOT doubles)."""

    def __init__(self):
        self.part = cuda.device_array((NSLOT, MAXB), np.float64)
        self.out = cuda.device_array(NSLOT, np.float64)
        self.host = cuda.pinned_array(NSLOT, np.float64)
        self._ops = {}

    def ops_dev(self, ops):
        key = tuple(ops)
        d = self._ops.get(key)
        if d is None:
            d = cuda.to_device(np.array(key, dtype=np.int32))
            self._ops[key] = d
        return d

    def finish(self, ops, nb):
        k = len(ops)
        _k_combine[k, BS](self.part, nb, self.ops_dev(ops), self.out)
        self.out.copy_to_host(self.host)
        return self.host[:k].copy()


def _sync_free_zero(arr):
    """Zero a device array (kernel; numba has no memset)."""
    n = arr.size
    _k_zero[_nb(n), BS](arr.reshape(-1) if arr.ndim != 1 else arr, n)


@cuda.jit
def _k_zero(a, n):
    for i in range(cuda.grid(1), n, cuda.gridsize(1)):
        a[i] = 0


@cuda.jit
def _k_copy(src, dst, n):
    for i in range(cuda.grid(1), n, cuda.gridsize(1)):
        dst[i] = src[i]


def _dev_copy(src, dst):
    n = src.size
    _k_copy[_nb(n), BS](src, dst, n)


# ============================================================================ lazily downloaded fields
class DeviceField:
    """
    A device-resident per-voxel array that behaves like the host ndarray the optimiser's host code expects for the
    handful of operations it performs on it (np.asarray / astype / indexing / min / max / mean / copy). The download
    happens on first use. `version` ties the field to the owner's buffer generation: a field read after its buffer
    was reused raises instead of returning another iteration's data.
    """
    __array_priority__ = 1000

    def __init__(self, owner, key, dev, shape, dtype, stats=None):
        self._owner, self._key = owner, key
        self._dev, self.shape, self.dtype = dev, tuple(shape), np.dtype(dtype)
        self._gen = owner.generation(key)
        self._host = None
        self._stats = stats                  # (min, max, sum) preloaded by the owner, or None

    # --- ndarray look-alike
    @property
    def size(self):
        return int(np.prod(self.shape))

    @property
    def ndim(self):
        return len(self.shape)

    def _check(self):
        if self._owner.generation(self._key) != self._gen:
            raise RuntimeError(f"DeviceField {self._key!r} was read after its device buffer was reused")

    def host(self):
        if self._host is None:
            self._check()
            self._host = self._dev.copy_to_host().reshape(self.shape)
        return self._host

    def __array__(self, dtype=None, copy=None):
        h = self.host()
        return h if dtype is None else h.astype(dtype)

    def astype(self, dtype, *a, **k):
        return self.host().astype(dtype, *a, **k)

    def copy(self):
        return self.host().copy()

    def __getitem__(self, item):
        return self.host()[item]

    def __len__(self):
        return self.shape[0]

    def _st(self):
        if self._stats is None:
            self._check()
            self._stats = self._owner.field_stats(self._dev, self.dtype)
        return self._stats

    def min(self):
        return self.dtype.type(self._st()[0])

    def max(self):
        return self.dtype.type(self._st()[1])

    def mean(self):
        return self.dtype.type(self._st()[2] / self.size)


@cuda.jit
def _k_stats(a, n, part):
    sh = cuda.shared.array(BS, float64)
    lo = _BIG
    hi = -_BIG
    s = 0.0
    for i in range(cuda.grid(1), n, cuda.gridsize(1)):
        v = float64(a[i])
        s += v
        if v < lo:
            lo = v
        if v > hi:
            hi = v
    r0 = _bmin(sh, lo)
    r1 = _bmax(sh, hi)
    r2 = _bsum(sh, s)
    if cuda.threadIdx.x == 0:
        part[0, cuda.blockIdx.x] = r0
        part[1, cuda.blockIdx.x] = r1
        part[2, cuda.blockIdx.x] = r2


class _Base:
    """Shared plumbing: the reducer, field generations, statistics of a device array."""

    def __init__(self, dims):
        self.dims = tuple(int(d) for d in dims)
        self.nx, self.ny, self.nz = self.dims
        self.n = self.nx * self.ny * self.nz
        self.nb = _nb(self.n)
        self.red = _Reducer()
        self._gen = {}

    def generation(self, key):
        return self._gen.get(key, 0)

    def bump(self, key):
        self._gen[key] = self._gen.get(key, 0) + 1

    def field_stats(self, dev, dtype):
        _k_stats[self.nb, BS](dev, self.n, self.red.part)
        r = self.red.finish((2, 1, 0), self.nb)         # slots 0,1,2 combined as min, max, sum
        return float(r[0]), float(r[1]), float(r[2])

    def alloc(self, dtype):
        return cuda.device_array(self.n, dtype)

    def up(self, host, dtype=None, out=None):
        h = np.ascontiguousarray(host if dtype is None else np.asarray(host, dtype=dtype)).reshape(-1)
        if out is None:
            return cuda.to_device(h)
        cuda.to_device(h, to=out)
        return out

    def mask_up(self, m):
        return cuda.to_device(np.ascontiguousarray(np.asarray(m, dtype=np.uint8)).reshape(-1))


# ============================================================================ diagnostics
@cuda.jit
def _k_scan(phi, phi_old, has_old, design, preserve, obstacle, ls, c1, S, V, one, band_lo, band_hi, grey_lo, grey_hi,
            step_all, solid, voidnow, n, part):
    """Per-voxel scan of the new field; slots: 0 solid count, 1 void film sum, 2 sum rho(1-rho) over the design,
    3 undecided count, 4 step max, 5 step sum, 6 grey count, 7 phase switches, 8 step count."""
    sh = cuda.shared.array(BS, float64)
    ns = 0.0
    vf = 0.0
    nd = 0.0
    und = 0.0
    smax = 0.0
    ssum = 0.0
    gr = 0.0
    sw = 0.0
    sc = 0.0
    for i in range(cuda.grid(1), n, cuda.gridsize(1)):
        ph = phi[i]
        sol = (ph >= ls) or (preserve[i] != 0)
        if obstacle[i] != 0:
            sol = False
        solid[i] = 1 if sol else 0
        voidnow[i] = 1 if (ph < ls) else 0
        if sol:
            ns += 1.0
        d = design[i] != 0
        if d:
            r = -c1 * (ph - S) + (ph - V)
            r = _mn(_mx(r, c1), one)
            if not sol:
                vf += float64(r - c1)
            rr = float64(r)
            nd += rr * (1.0 - rr)
            if ph >= band_lo and ph < band_hi:
                und += 1.0
            if ph > grey_lo and ph < grey_hi:
                gr += 1.0
            if has_old != 0:
                if (ph >= ls) != (phi_old[i] >= ls):
                    sw += 1.0
        if has_old != 0 and (d or step_all != 0):
            dd = float64(abs(ph - phi_old[i]))
            ssum += dd
            sc += 1.0
            if dd > smax:
                smax = dd
    r0 = _bsum(sh, ns)
    r1 = _bsum(sh, vf)
    r2 = _bsum(sh, nd)
    r3 = _bsum(sh, und)
    r4 = _bmax(sh, smax)
    r5 = _bsum(sh, ssum)
    r6 = _bsum(sh, gr)
    r7 = _bsum(sh, sw)
    r8 = _bsum(sh, sc)
    if cuda.threadIdx.x == 0:
        b = cuda.blockIdx.x
        part[0, b] = r0
        part[1, b] = r1
        part[2, b] = r2
        part[3, b] = r3
        part[4, b] = r4
        part[5, b] = r5
        part[6, b] = r6
        part[7, b] = r7
        part[8, b] = r8


@cuda.jit
def _k_euler(solid, nx, ny, nz, part, off):
    """Euler characteristic V - E + F - C of the union of unit cubes at the solid voxels (the host
    euler_characteristic), slot 0; slot 1: interface faces (solid/void changes along each axis, domain skin
    included)."""
    sh = cuda.shared.array(BS, float64)
    chi = 0.0
    faces = 0.0
    sx = nx + 1
    sy = ny + 1
    sz = nz + 1
    total = sx * sy * sz
    for idx in range(cuda.grid(1), total, cuda.gridsize(1)):
        k = idx % sz
        t = idx // sz
        j = t % sy
        i = t // sy
        # solid flags of the 8 voxels around the lattice site (i, j, k): voxel (i-a, j-b, k-c)
        s000 = False
        s100 = False
        s010 = False
        s110 = False
        s001 = False
        s101 = False
        s011 = False
        s111 = False
        # naming: sABC = voxel (i-1+A, j-1+B, k-1+C)
        if i >= 1 and j >= 1 and k >= 1:
            s000 = solid[((i - 1) * ny + (j - 1)) * nz + (k - 1)] != 0
        if i < nx and j >= 1 and k >= 1:
            s100 = solid[(i * ny + (j - 1)) * nz + (k - 1)] != 0
        if i >= 1 and j < ny and k >= 1:
            s010 = solid[((i - 1) * ny + j) * nz + (k - 1)] != 0
        if i < nx and j < ny and k >= 1:
            s110 = solid[(i * ny + j) * nz + (k - 1)] != 0
        if i >= 1 and j >= 1 and k < nz:
            s001 = solid[((i - 1) * ny + (j - 1)) * nz + k] != 0
        if i < nx and j >= 1 and k < nz:
            s101 = solid[(i * ny + (j - 1)) * nz + k] != 0
        if i >= 1 and j < ny and k < nz:
            s011 = solid[((i - 1) * ny + j) * nz + k] != 0
        if i < nx and j < ny and k < nz:
            s111 = solid[(i * ny + j) * nz + k] != 0
        # vertex
        if s000 or s100 or s010 or s110 or s001 or s101 or s011 or s111:
            chi += 1.0
        # voxel (i, j, k) itself
        if i < nx and j < ny and k < nz and s111:
            chi -= 1.0
        # faces normal to x at (i; j, k): voxels (i-1, j, k) and (i, j, k)
        if j < ny and k < nz:
            if s011 or s111:
                chi += 1.0
            if s011 != s111:
                faces += 1.0
        # faces normal to y at (j; i, k): voxels (i, j-1, k), (i, j, k)
        if i < nx and k < nz:
            if s101 or s111:
                chi += 1.0
            if s101 != s111:
                faces += 1.0
        # faces normal to z at (k; i, j): voxels (i, j, k-1), (i, j, k)
        if i < nx and j < ny:
            if s110 or s111:
                chi += 1.0
            if s110 != s111:
                faces += 1.0
        # edges parallel to x at (i; j, k): voxels (i, j-1..j, k-1..k)
        if i < nx:
            if s100 or s110 or s101 or s111:
                chi -= 1.0
        # edges parallel to y at (j; i, k): voxels (i-1..i, j, k-1..k)
        if j < ny:
            if s010 or s110 or s011 or s111:
                chi -= 1.0
        # edges parallel to z at (k; i, j): voxels (i-1..i, j-1..j, k)
        if k < nz:
            if s001 or s101 or s011 or s111:
                chi -= 1.0
    r0 = _bsum(sh, chi)
    r1 = _bsum(sh, faces)
    if cuda.threadIdx.x == 0:
        part[off, cuda.blockIdx.x] = r0
        part[off + 1, cuda.blockIdx.x] = r1


@cuda.jit
def _k_cc_init(mask, f, n):
    for i in range(cuda.grid(1), n, cuda.gridsize(1)):
        f[i] = i if mask[i] != 0 else -1


@cuda.jit
def _k_cc_hook(mask, f, nx, ny, nz, conn26, changed, slot):
    n = nx * ny * nz
    for i in range(cuda.grid(1), n, cuda.gridsize(1)):
        if mask[i] == 0:
            continue
        fi = f[i]
        z = i % nz
        t = i // nz
        y = t % ny
        x = t // ny
        for dx in range(-1, 2):
            xx = x + dx
            if xx < 0 or xx >= nx:
                continue
            for dy in range(-1, 2):
                yy = y + dy
                if yy < 0 or yy >= ny:
                    continue
                for dz in range(-1, 2):
                    if dx == 0 and dy == 0 and dz == 0:
                        continue
                    if conn26 == 0 and (abs(dx) + abs(dy) + abs(dz)) != 1:
                        continue
                    zz = z + dz
                    if zz < 0 or zz >= nz:
                        continue
                    j = (xx * ny + yy) * nz + zz
                    if mask[j] != 0:
                        fj = f[j]
                        if fj < fi:
                            old = cuda.atomic.min(f, fi, fj)
                            if fj < old:
                                changed[slot] = 1


@cuda.jit
def _k_cc_compress(mask, f, n, changed, slot):
    for i in range(cuda.grid(1), n, cuda.gridsize(1)):
        if mask[i] != 0:
            p = f[i]
            pp = f[p]
            if pp != p:
                f[i] = pp
                changed[slot] = 1


@cuda.jit
def _k_cc_flags_conn(mask, f, load, support, design, flags, size, n):
    for i in range(cuda.grid(1), n, cuda.gridsize(1)):
        if mask[i] != 0:
            r = f[i]
            bits = 0
            if load[i] != 0:
                bits |= 1
            if support[i] != 0:
                bits |= 2
            if design[i] != 0:
                bits |= 4
            if bits != 0:
                cuda.atomic.or_(flags, r, bits)
            cuda.atomic.add(size, r, 1)


@cuda.jit
def _k_cc_final_conn(mask, f, flags, size, n, part):
    """slots: 0 components, 1 components touching a load, 2 loaded components without a support,
    3 size of the detached components (no load-support path, holding design voxels)."""
    sh = cuda.shared.array(BS, float64)
    nc = 0.0
    nl = 0.0
    nbad = 0.0
    det = 0.0
    for i in range(cuda.grid(1), n, cuda.gridsize(1)):
        if mask[i] != 0 and f[i] == i:
            fl = flags[i]
            nc += 1.0
            if (fl & 1) != 0:
                nl += 1.0
                if (fl & 2) == 0:
                    nbad += 1.0
            if not ((fl & 1) != 0 and (fl & 2) != 0) and (fl & 4) != 0:
                det += float64(size[i])
    r0 = _bsum(sh, nc)
    r1 = _bsum(sh, nl)
    r2 = _bsum(sh, nbad)
    r3 = _bsum(sh, det)
    if cuda.threadIdx.x == 0:
        b = cuda.blockIdx.x
        part[0, b] = r0
        part[1, b] = r1
        part[2, b] = r2
        part[3, b] = r3


@cuda.jit
def _k_cc_flags_skin(mask, f, nx, ny, nz, prev_void, use_prev, flags, size, n):
    """Cluster flags for the cavity count (bit 8 = touches the domain skin) and for nucleation (bit 8 = skin,
    bit 16 = growth: a voxel of the cluster has a 26-neighbour (or itself) that was void)."""
    for i in range(cuda.grid(1), n, cuda.gridsize(1)):
        if mask[i] == 0:
            continue
        r = f[i]
        z = i % nz
        t = i // nz
        y = t % ny
        x = t // ny
        bits = 0
        if x == 0 or x == nx - 1 or y == 0 or y == ny - 1 or z == 0 or z == nz - 1:
            bits |= 8
        if use_prev != 0:
            grown = False
            for dx in range(-1, 2):
                xx = x + dx
                if xx < 0 or xx >= nx:
                    continue
                for dy in range(-1, 2):
                    yy = y + dy
                    if yy < 0 or yy >= ny:
                        continue
                    for dz in range(-1, 2):
                        zz = z + dz
                        if zz < 0 or zz >= nz:
                            continue
                        if prev_void[(xx * ny + yy) * nz + zz] != 0:
                            grown = True
            if grown:
                bits |= 16
        if bits != 0:
            cuda.atomic.or_(flags, r, bits)
        cuda.atomic.add(size, r, 1)


@cuda.jit
def _k_cc_final_cav(mask, f, flags, n, part):
    """slot 0: components with no skin voxel (cavities)."""
    sh = cuda.shared.array(BS, float64)
    nc = 0.0
    for i in range(cuda.grid(1), n, cuda.gridsize(1)):
        if mask[i] != 0 and f[i] == i and (flags[i] & 8) == 0:
            nc += 1.0
    r0 = _bsum(sh, nc)
    if cuda.threadIdx.x == 0:
        part[0, cuda.blockIdx.x] = r0


@cuda.jit
def _k_cc_final_nuc(mask, f, flags, size, n, part):
    """slots: 0 nucleated interior clusters, 1 nucleated surface clusters, 2 voxels of the nucleated clusters."""
    sh = cuda.shared.array(BS, float64)
    ni = 0.0
    ns = 0.0
    nv = 0.0
    for i in range(cuda.grid(1), n, cuda.gridsize(1)):
        if mask[i] != 0 and f[i] == i:
            fl = flags[i]
            if (fl & 16) == 0:
                if (fl & 8) == 0:
                    ni += 1.0
                else:
                    ns += 1.0
                nv += float64(size[i])
    r0 = _bsum(sh, ni)
    r1 = _bsum(sh, ns)
    r2 = _bsum(sh, nv)
    if cuda.threadIdx.x == 0:
        part[0, cuda.blockIdx.x] = r0
        part[1, cuda.blockIdx.x] = r1
        part[2, cuda.blockIdx.x] = r2


@cuda.jit
def _k_new_mask(voidnow, prev_void, design, new, n, part):
    sh = cuda.shared.array(BS, float64)
    c = 0.0
    for i in range(cuda.grid(1), n, cuda.gridsize(1)):
        v = 1 if (voidnow[i] != 0 and prev_void[i] == 0 and design[i] != 0) else 0
        new[i] = v
        c += float64(v)
    r = _bsum(sh, c)
    if cuda.threadIdx.x == 0:
        part[0, cuda.blockIdx.x] = r


@cuda.jit
def _k_not(src, dst, n):
    for i in range(cuda.grid(1), n, cuda.gridsize(1)):
        dst[i] = 0 if src[i] != 0 else 1


class DiagGPU(_Base):
    """The diagnostics of docs/theory.md §8.4.1 on the device (module docstring)."""

    def __init__(self, dims, design_mask, preserve, obstacle, load, support, void_density, solid_value, void_value,
                 level_set_value):
        super().__init__(dims)
        self.c1, self.S, self.V, self.ls = float(void_density), float(solid_value), float(void_value), \
            float(level_set_value)
        design = np.asarray(design_mask, dtype=bool)
        self.any_design = bool(np.any(design))
        self.n_design = int(np.count_nonzero(design))
        self.d_design = self.mask_up(design)
        self.d_pres, self.d_obs = self.mask_up(preserve), self.mask_up(obstacle)
        self.d_load, self.d_sup = self.mask_up(load), self.mask_up(support)
        a = self.alloc
        u8 = np.uint8
        self.d_solid, self.d_void, self.d_prev_void = a(u8), a(u8), a(u8)
        self.have_prev = False
        self.d_new, self.d_inv = a(u8), a(u8)
        self.d_f = a(np.int32)
        self.d_flags, self.d_size = a(np.int32), a(np.int32)
        self.d_changed = cuda.device_array(8, np.int32)
        self.d_phi_new = {}
        self.d_prev_phi = None
        self._host_new = None
        self._scan = None
        self._cc = None
        self._topo = None
        self.dtype = None

    # --- device mirrors of host fields
    def _dev_phi(self, phi, which):
        """Device copy of the host field `phi` (identity-cached for the last field registered)."""
        slot = self.d_phi_new.get(which)
        if slot is not None and slot[0] is phi:
            return slot[1]
        dt = np.float32 if phi.dtype == np.float32 else np.float64
        if slot is not None and slot[2] and slot[1].dtype == dt:
            dev = slot[1]                                   # our own buffer: reuse
        else:
            dev = cuda.device_array(self.n, dt)
        self.up(phi, dt, out=dev)
        self.d_phi_new[which] = (phi, dev, True)
        return dev

    def register(self, which, phi_host, dev):
        """Tell the cache that device array `dev` (not ours) holds the host array `phi_host`."""
        self.d_phi_new[which] = (phi_host, dev, False)

    def _cast(self, dev, v):
        return dev.dtype.type(v)

    # --- the per-iteration scan
    def scan(self, phi_new, phi_old, dt_for_step=None):
        """Scalars of one iteration; cached per phi_new object (grey fraction, phase switches, diagnostics and the
        connectivity report all read it)."""
        if self._scan is not None and self._scan[0] is phi_new and self._scan[1] is phi_old:
            return self._scan[2]
        d_new = self._dev_phi(phi_new, "new")
        has_old = phi_old is not None
        d_old = self._dev_phi(phi_old, "old") if has_old else d_new
        band = 0.25 * (self.S - self.V)
        gband = 0.1 * (self.S - self.V)
        ct = lambda v: d_new.dtype.type(v)       # noqa: E731  compare in the field's own precision
        _k_scan[self.nb, BS](d_new, d_old, 1 if has_old else 0, self.d_design, self.d_pres, self.d_obs,
                             ct(self.ls), ct(self.c1), ct(self.S), ct(self.V), ct(1.0), ct(self.V + band), ct(self.S - band),
                             ct(self.V + gband), ct(self.S - gband), 0 if self.any_design else 1, self.d_solid,
                             self.d_void, self.n, self.red.part)
        _k_euler[self.nb, BS](self.d_solid, self.nx, self.ny, self.nz, self.red.part, 9)
        r = self.red.finish((0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0), self.nb)
        out = dict(n_solid=int(round(r[0])), void_film=float(r[1]), nd_sum=float(r[2]), undecided=int(round(r[3])),
                   step_max=float(r[4]), step_sum=float(r[5]), grey=int(round(r[6])), switches=int(round(r[7])),
                   step_count=float(r[8]), euler=int(round(r[9])), faces=int(round(r[10])))
        self._scan = (phi_new, phi_old, out)
        self._cc = None
        self._topo = None
        return out

    # --- connected components
    def _label(self, mask, conn26):
        """Connected components of `mask` into self.d_f (label = the smallest voxel index of the component).
        Union-find by hooking roots to smaller neighbouring roots with atomicMin, then pointer jumping; four rounds are
        queued per host read (one device-to-host copy per block, not per kernel). A round that hooks nothing means
        every neighbouring pair carries one label: the labelling is final."""
        c = self
        n = self.n
        _k_cc_init[c.nb, BS](mask, self.d_f, n)
        flags = self.d_changed
        for _ in range(2500):
            _k_zero[1, 32](flags, 8)
            for r in range(4):
                _k_cc_hook[c.nb, BS](mask, self.d_f, self.nx, self.ny, self.nz, 1 if conn26 else 0, flags, r)
                for _k in range(3):
                    _k_cc_compress[c.nb, BS](mask, self.d_f, n, flags, 4 + r)
            if int(flags.copy_to_host()[3]) == 0:
                break

    def current(self, phi):
        """The scan of `phi` if it is the field of the last scan, else a scan without a previous field."""
        if self._scan is not None and self._scan[0] is phi:
            return self._scan[2]
        return self.scan(phi, None)

    def connectivity(self, phi, tol, any_load):
        """The connectivity report of the exported solid of phi (baseline_optimiser._connectivity_of_solid)."""
        sc = self.current(phi)
        if self._cc is not None:
            return self._cc
        n_solid = sc["n_solid"]
        if n_solid == 0:
            self._cc = dict(n_components=0, load_path_connected=not any_load, detached_fraction=0.0,
                            connected=not any_load)
            return self._cc
        self._label(self.d_solid, True)
        _sync_free_zero(self.d_flags)
        _sync_free_zero(self.d_size)
        _k_cc_flags_conn[self.nb, BS](self.d_solid, self.d_f, self.d_load, self.d_sup, self.d_design, self.d_flags,
                                      self.d_size, self.n)
        _k_cc_final_conn[self.nb, BS](self.d_solid, self.d_f, self.d_flags, self.d_size, self.n, self.red.part)
        r = self.red.finish((0, 0, 0, 0), self.nb)
        ncomp, nl, nbad, det = int(round(r[0])), int(round(r[1])), int(round(r[2])), float(r[3])
        lpc = (nbad == 0) if nl > 0 else True
        frac = det / n_solid
        self._cc = dict(n_components=ncomp, load_path_connected=bool(lpc), detached_fraction=float(frac),
                        connected=bool(lpc and frac <= tol))
        return self._cc

    def topology(self, phi, n_components):
        """euler, b0, b1, b2 and interface_faces of the exported solid of phi."""
        sc = self.current(phi)
        if self._topo is not None and self._topo[0] == n_components:
            return self._topo[1]
        chi = sc["euler"] if sc["n_solid"] > 0 else 0
        faces = sc["faces"]
        if n_components is None:
            if sc["n_solid"] == 0:
                n_components = 0
            else:
                n_components = self.connectivity(phi, 0.0, False)["n_components"]
        # cavities: 6-connected void components that never touch the domain skin
        _k_not[self.nb, BS](self.d_solid, self.d_inv, self.n)
        self._label(self.d_inv, False)
        _sync_free_zero(self.d_flags)
        _sync_free_zero(self.d_size)
        _k_cc_flags_skin[self.nb, BS](self.d_inv, self.d_f, self.nx, self.ny, self.nz, self.d_inv, 0, self.d_flags,
                                      self.d_size, self.n)
        _k_cc_final_cav[self.nb, BS](self.d_inv, self.d_f, self.d_flags, self.n, self.red.part)
        b2 = int(round(self.red.finish((0,), self.nb)[0]))
        b0 = int(n_components)
        topo = dict(euler=chi, b0=b0, b1=int(b0 + b2 - chi), b2=b2, interface_faces=faces)
        self._topo = (n_components, topo)
        return topo

    def nucleation_counts(self, phi):
        """(n_interior, n_surface, n_voxels) of the new void clusters of this iteration that are not the growth of an
        existing void (host detect_nucleations), or None when there is no previous void yet."""
        self.current(phi)
        if not self.have_prev:
            return None
        _k_new_mask[self.nb, BS](self.d_void, self.d_prev_void, self.d_design, self.d_new, self.n, self.red.part)
        cnt = int(round(self.red.finish((0,), self.nb)[0]))
        if cnt == 0:
            return 0, 0, 0
        self._label(self.d_new, False)
        _sync_free_zero(self.d_flags)
        _sync_free_zero(self.d_size)
        _k_cc_flags_skin[self.nb, BS](self.d_new, self.d_f, self.nx, self.ny, self.nz, self.d_prev_void, 1,
                                      self.d_flags, self.d_size, self.n)
        _k_cc_final_nuc[self.nb, BS](self.d_new, self.d_f, self.d_flags, self.d_size, self.n, self.red.part)
        r = self.red.finish((0, 0, 0), self.nb)
        return int(round(r[0])), int(round(r[1])), int(round(r[2]))

    def prev_void_host(self):
        return self.d_prev_void.copy_to_host().reshape(self.dims).astype(bool)

    def void_host(self):
        return self.d_void.copy_to_host().reshape(self.dims).astype(bool)

    def solid_host(self):
        return self.d_solid.copy_to_host().reshape(self.dims).astype(bool)

    def advance_void(self):
        """The current void becomes the previous one (the end of RunDiagnostics.update)."""
        _dev_copy(self.d_void, self.d_prev_void)
        self.have_prev = True


# ============================================================================ memory model
def update_memory_bytes(engine, n, kind="simp"):
    """Device memory of the update engine for n voxels (arrays of docs/theory.md §18 §4); the diagnostics
    add 11 bytes per voxel (masks, labels, flags)."""
    diag = 4 * n + 4 * n + 4 * n + 12 * n          # int32 label / flags / size + uint8 masks
    if kind == "simp":
        return int(8 * n * 28 + 4 * n * 3 + diag)
    return int(4 * n * 20 + 8 * n * 3 + diag)


# ======================================================================================================================
# === former postprocess.py: statistics, plots, report.pdf, evolution.gif (former utils/postprocess.py) ===
# ======================================================================================================================
"""
postprocess -- statistics, plots and a PDF report for a PCLS optimisation run
(docs/theory.md §8.4.3).

Runs automatically at the end of PCLS_Optimiser.run_optimisation(), and on its
own to regenerate everything from a run directory's files (for instance after
copying a run back from the cluster, or for a run made before these outputs
existed -- whatever its files allow is produced):

    python optimise.py --postprocess-only <run_dir> [<run_dir> ...] [--dpi 150] [--no-render] [--no-gif]

Reads, when present: variables.csv, parameters.csv, buckling.csv,
nucleation_events.csv, threshold_sweep.csv, run_info.json, final_fields.npz and
history/phi_*.npz. Writes into the run directory:

    summary.json, summary.txt    every statistic below, machine- and human-readable
    plots/*.png                  one file per figure
    report.pdf                   the summary page followed by every figure
    evolution.gif                the design's evolution (needs history/)

Only NumPy, SciPy and Matplotlib are required (Matplotlib is driven through
its object-oriented API with the Agg canvas, so no display is needed and no
global backend is changed). scikit-image, if installed, is used for the 3D
surface render; Pillow (a Matplotlib dependency) writes the GIF. Every figure
is produced in its own try/except: one that fails is listed in
plots/postprocess_log.txt and the rest are still written.
"""
import os
import sys
import csv
import glob
import json
import argparse
import traceback
from timeit import default_timer as timer

import numpy as np
from scipy import ndimage


# Colours (colour-blind safe: no red/green pairs carry meaning on their own).
C_MODEL = "#1f77b4"
C_EXPORT = "#ff7f0e"
C_TARGET = "#555555"
C_ALT = "#9467bd"
C_ALT2 = "#8c564b"
C_BAD = "#d62728"
C_OK = "#2c7fb8"
SETTLE_TESTS = ("p_rising", "volume", "multiplier", "grey", "compliance", "switches", "buckling", "analysis")


# ---------------------------------------------------------------------------
# Reading the run's files
# ---------------------------------------------------------------------------

def load_table(path: str) -> dict:
    """A CSV as {column: array}; numeric where every entry parses ('N/A' -> nan), else object."""
    if not os.path.isfile(path):
        return {}
    with open(path, newline="") as f:
        rows = list(csv.reader(f))
    if len(rows) < 1:
        return {}
    header, data = rows[0], rows[1:]
    cols = {}
    for j, name in enumerate(header):
        vals = [r[j].strip() if j < len(r) else "" for r in data]
        num = np.empty(len(vals), dtype=np.float64)
        ok = True
        for i, v in enumerate(vals):
            if v in ("", "N/A", "None"):
                num[i] = np.nan
                continue
            try:
                num[i] = float(v)
            except ValueError:
                ok = False
                break
        cols[name] = num if ok else np.array(vals, dtype=object)
    return cols


def _first_row(table: dict, key: str, default=None):
    v = table.get(key)
    if v is None or len(v) == 0:
        return default
    return v[0]


class RunContext:
    """Everything postprocessing knows about one run."""

    def __init__(self, run_dir: str, dpi: int = 150, render3d: bool = True, gif: bool = True,
                 verbose: bool = True, create_plot_dir: bool = True):
        self.run_dir = os.path.abspath(run_dir)
        self.name = os.path.basename(os.path.normpath(run_dir))
        self.plot_dir = os.path.join(self.run_dir, "plots")
        if create_plot_dir:
            os.makedirs(self.plot_dir, exist_ok=True)
        self.dpi = dpi
        self.render3d = render3d
        self.gif = gif
        self.verbose = verbose
        self.log = []
        self.written = []
        p = lambda f: os.path.join(self.run_dir, f)  # noqa: E731
        self.var = load_table(p("variables.csv"))
        self.params = load_table(p("parameters.csv"))
        self.buck = load_table(p("buckling.csv"))
        self.events = load_table(p("nucleation_events.csv"))
        self.sweep = load_table(p("threshold_sweep.csv"))
        self.info = {}
        if os.path.isfile(p("run_info.json")):
            try:
                with open(p("run_info.json")) as f:
                    self.info = json.load(f)
            except Exception as e:
                self.note(f"run_info.json unreadable: {e}")
        self.fields = {}
        if os.path.isfile(p("final_fields.npz")):
            try:
                with np.load(p("final_fields.npz")) as z:
                    self.fields = {k: z[k] for k in z.files}
            except Exception as e:
                self.note(f"final_fields.npz unreadable: {e}")
        self.history = sorted(glob.glob(os.path.join(self.run_dir, "history", "phi_*.npz")))
        self._setup()

    def note(self, msg: str):
        self.log.append(msg)
        if self.verbose:
            print(f"  [postprocess] {msg}")

    # -- parameters -------------------------------------------------------
    def param(self, name: str, default=None):
        """An optimiser parameter: run_info.json first, then parameters.csv."""
        op = self.info.get("optimiser_parameters", {})
        if name in op and op[name] is not None:
            return op[name]
        v = _first_row(self.params, f"optimiser_{name}")
        if v is None or (isinstance(v, float) and np.isnan(v)):
            return default
        return v

    def _setup(self):
        self.target = float(self.param("target_volume", 0.5))
        self.S = float(self.param("solid_value", 2.0))
        self.V = float(self.param("void_value", 1.0))
        self.c1 = float(self.param("void_density", 1e-3))
        self.ls = 0.5 * (self.S + self.V)
        self.voxel_mm = float(self.info.get("voxel_size_mm") or _first_row(self.params, "analysis_voxel_size_mm", 1.0) or 1.0)
        self.tol_outer = float(self.param("outer_loop_convergence_threshold", 0.01))
        self.tol_inner = float(self.param("inner_loop_convergence_threshold", 0.01))
        self.tol_grey = float(self.param("convergence_grey_tol", 0.01))
        self.tol_comp = float(self.param("convergence_compliance_tol", 1e-3))
        self.tol_switch = float(self.param("convergence_switch_tol", 1e-4))
        # The method of the run (run_info.json; absent = PCLS) and the number of consecutive settled
        # iterations it requires: the baselines write it (SIMP's is 1), older files fall back on the method.
        self.method = str(self.info.get("method") or "pcls").lower()
        self.method_label = {"pcls": "PCLS", "simp": "SIMP", "levelset": "Level-set"}.get(
            self.method, self.method.capitalize())
        win = self.info.get("settle_window")
        if win is None:
            win = 1 if self.method == "simp" else self.param("convergence_window", 10)
        self.window = int(win)
        self.buckling = bool(self.info.get("buckling", {}).get("enabled", False)) or (
            len(self.buck.get("iteration", [])) > 0)
        dims = self.info.get("domain_dims")
        if dims is None and "phi" in self.fields:
            dims = list(self.fields["phi"].shape)
        if dims is None and self.history:
            # A run that stopped before its final records (e.g. a crash): the snapshots still tell.
            try:
                with np.load(self.history[0]) as z:
                    dims = list(z["phi"].shape)
            except Exception:
                dims = None
        self.dims = tuple(int(d) for d in dims) if dims else None
        nd = self.info.get("n_design_voxels")
        if nd is None and "design_mask" in self.fields:
            nd = int(np.count_nonzero(self.fields["design_mask"]))
        self.n_design = nd
        # Per-iteration rows: the first row of variables.csv is the initial
        # state, written before any iteration ran.
        it = self.var.get("iteration")
        t = self.var.get("total_iteration_time")
        if it is not None and t is not None:
            self.rows = np.where(~np.isnan(t))[0]
        elif it is not None:
            self.rows = np.arange(len(it))
        else:
            self.rows = np.zeros(0, dtype=int)
        self.x = self.var["iteration"][self.rows] if it is not None else np.zeros(0)
        # Outer-iteration boundaries: where outer_iteration changes (or, for
        # older files, where lambda1 was updated).
        self.outer_starts = []
        oi = self.col("outer_iteration")
        if oi is not None and np.any(~np.isnan(oi)):
            ch = np.where(np.diff(oi) != 0)[0] + 1
            self.outer_starts = list(self.x[ch])
        else:
            l1 = self.col("lambda1")
            if l1 is not None and len(l1) > 1:
                ch = np.where(np.abs(np.diff(l1)) > 0)[0] + 1
                self.outer_starts = list(self.x[ch])
        self.resets = list(self.info.get("reconnection_reset_iterations") or [])

    def col(self, name: str):
        """A variables.csv column over the per-iteration rows (None if absent)."""
        v = self.var.get(name)
        if v is None:
            return None
        return v[self.rows]

    def has(self, *names) -> bool:
        for n in names:
            v = self.col(n)
            if v is None:
                return False
            if v.dtype != object and np.all(np.isnan(v)):
                return False
        return True


# ---------------------------------------------------------------------------
# Figure helpers
# ---------------------------------------------------------------------------

def new_figure(width=11.0, height=6.0):
    from matplotlib.figure import Figure
    from matplotlib.backends.backend_agg import FigureCanvasAgg
    fig = Figure(figsize=(width, height))
    FigureCanvasAgg(fig)
    return fig


def decorate(ax, ctx: RunContext, xlabel=True):
    """Outer-iteration boundaries (light), reconnection resets (dashed) and a grid."""
    for x0 in ctx.outer_starts:
        ax.axvline(x0, color="0.85", lw=0.6, zorder=0)
    for x0 in ctx.resets:
        ax.axvline(x0, color=C_BAD, lw=0.9, ls="--", zorder=1)
    ax.grid(True, alpha=0.25)
    if xlabel:
        ax.set_xlabel("iteration")


def logy_if_positive(ax, *arrays):
    """Log y-axis when every value is positive and they span more than a factor of 30."""
    parts = []
    for a in arrays:
        if a is None:
            continue
        a = np.asarray(a)
        if a.dtype == object:
            continue
        a = a.astype(float)
        parts.append(a[np.isfinite(a)])
    vals = np.concatenate(parts) if parts else np.zeros(0)
    if vals.size and np.all(vals > 0) and vals.max() / max(vals.min(), 1e-300) > 30:
        ax.set_yscale("log")


def _legend(ax, **kw):
    """A legend only if something on the axes is labelled."""
    handles, labels = ax.get_legend_handles_labels()
    if handles:
        ax.legend(**kw)


def finite(a):
    a = np.asarray(a, dtype=float)
    return a[np.isfinite(a)]


class Plotter:
    """Collects figures into plots/*.png and report.pdf."""

    def __init__(self, ctx: RunContext, pdf: "PdfPages | None"):
        self.ctx = ctx
        self.pdf = pdf
        self.n = 0

    def run(self, name: str, fn):
        t0 = timer()
        try:
            figs = fn(self.ctx)
        except Exception as e:
            self.ctx.note(f"figure '{name}' failed: {type(e).__name__}: {e}")
            with open(os.path.join(self.ctx.plot_dir, "postprocess_log.txt"), "a") as f:
                f.write(f"--- {name} ---\n{traceback.format_exc()}\n")
            return
        if figs is None:
            return
        if not isinstance(figs, (list, tuple)):
            figs = [(name, figs)]
        for fname, fig in figs:
            path = os.path.join(self.ctx.plot_dir, f"{fname}.png")
            try:
                fig.savefig(path, dpi=self.ctx.dpi, bbox_inches="tight")
                self.ctx.written.append(os.path.relpath(path, self.ctx.run_dir))
                if self.pdf is not None:
                    self.pdf.savefig(fig, bbox_inches="tight")
                self.n += 1
            except Exception as e:
                self.ctx.note(f"saving figure '{fname}' failed: {e}")
        if self.ctx.verbose and timer() - t0 > 20:
            print(f"  [postprocess] {name}: {timer() - t0:.0f} s")


# ---------------------------------------------------------------------------
# Summary statistics
# ---------------------------------------------------------------------------

def _num(v):
    if v is None:
        return None
    if isinstance(v, str):
        if v == "inf":
            return float("inf")
        try:
            return float(v)
        except ValueError:
            return None
    try:
        return float(v)
    except Exception:
        return None


def surface_summary(info: dict) -> dict:
    """Quality metrics of the run's final STL (baseline_optimiser.surface_quality) from run_info.json; {} if none was exported."""
    fs = (info or {}).get("final_surface")
    if not fs:
        return {}
    try:
        q = fs.get("quality") or surface_quality(fs)
    except Exception:
        return {}
    out = {k: v for k, v in q.items() if v is not None}
    if fs.get("method"):
        out["method"] = fs["method"]
    return out


def compute_summary(ctx: RunContext) -> dict:
    """Every statistic of the run and of its final design (docs/theory.md §8.4.3)."""
    info, f = ctx.info, ctx.fields
    fin = info.get("final", {})
    S = {"run": {}, "volume": {}, "discreteness": {}, "crispness": {}, "objective": {}, "topology": {},
         "geometry": {}, "surface": {}, "mechanics": {}, "buckling": {}, "multipliers": {},
         "convergence": {}, "timing": {}, "threshold_sweep": {}}
    r = S["run"]
    if ctx.method != "pcls":
        r["method"] = ctx.method
    r.update(name=ctx.name, analysis_class=info.get("analysis_class"), field_type=info.get("field_type"),
             domain_dims=ctx.dims, voxel_size_mm=ctx.voxel_mm, n_voxels=info.get("n_voxels"),
             n_design_voxels=ctx.n_design, termination_reason=info.get("termination_reason"),
             converged=info.get("converged"), total_iterations=info.get("total_iterations"),
             outer_iterations=info.get("outer_iterations"), reconnection_resets=info.get("reconnection_resets"),
             reconnection_reset_iterations=ctx.resets,
             wall_time_h=(info.get("wall_time_s") or 0) / 3600.0 if info.get("wall_time_s") is not None else None,
             total_load_N=info.get("total_load_N"))
    if r["termination_reason"] is None and len(ctx.x):
        r["total_iterations"] = int(np.nanmax(ctx.x)) + 1
    fail = info.get("failure")
    if fail:
        r.update(failure_reason=fail.get("reason"), failure_message=fail.get("message"),
                 failure_iteration=fail.get("iteration"))
    elif os.path.isfile(os.path.join(ctx.run_dir, "error.txt")):
        r["failure_message"] = "see error.txt"
    r.update(random_seed=info.get("random_seed"), overrides=info.get("overrides"))

    # --- timing ---
    t = ctx.col("total_iteration_time")
    if t is not None and len(finite(t)):
        a = ctx.col("analysis_time")
        bt = ctx.col("buckling_time")
        S["timing"].update(mean_iteration_s=float(np.nanmean(t)), median_iteration_s=float(np.nanmedian(t)),
                           max_iteration_s=float(np.nanmax(t)), total_iteration_h=float(np.nansum(t)) / 3600.0)
        if a is not None and len(finite(a)):
            S["timing"]["analysis_fraction"] = float(np.nansum(a) / max(np.nansum(t), 1e-300))
        if bt is not None and len(finite(bt)):
            S["timing"]["buckling_fraction"] = float(np.nansum(bt) / max(np.nansum(t), 1e-300))
        si = ctx.col("solver_iterations")
        if si is not None and len(finite(si)):
            S["timing"].update(mean_solver_iterations=float(np.nanmean(si)), max_solver_iterations=float(np.nanmax(si)))
        sc = ctx.col("solver_converged")
        if sc is not None and len(sc):
            if sc.dtype == object:
                sc = np.array([{"true": 1.0, "1": 1.0, "false": 0.0, "0": 0.0}.get(str(v).strip().lower(), np.nan)
                               for v in sc])
            if len(finite(sc)):
                # A solve stopped by pcg_max_iters is used as it is: its share says how
                # much of the run's sensitivity came from unconverged displacements.
                S["timing"]["unconverged_solve_share"] = float(np.mean(finite(sc) == 0))
    # Conformal's mesh (docs/theory.md §7.1): size, quality and conditioning over the run.
    me = ctx.col("mesh_elements")
    if me is not None and me.dtype != object and len(finite(me)):
        S["timing"].update(median_mesh_elements=float(np.nanmedian(me)), max_mesh_elements=float(np.nanmax(me)))
        q = ctx.col("min_tet_quality")
        if q is not None and q.dtype != object and len(finite(q)):
            S["timing"]["min_tet_quality"] = float(np.nanmin(q))
        dr = ctx.col("degenerate_tets_removed")
        if dr is not None and dr.dtype != object and len(finite(dr)):
            S["timing"]["degenerate_tets_removed_total"] = int(np.nansum(dr))
        mc = ctx.col("mc_vertices_clamped")
        if mc is not None and mc.dtype != object and len(finite(mc)):
            S["timing"]["median_mc_vertices_clamped"] = float(np.nanmedian(mc))

    # --- objective ---
    c = ctx.col("compliance")
    o = S["objective"]
    o["compliance_model_J"] = _num(fin.get("compliance")) if fin.get("compliance") is not None else (
        float(c[-1]) if c is not None and len(c) else None)
    o["compliance_exported_J"] = _num(fin.get("exported_compliance"))
    if o["compliance_exported_J"] is None and info.get("field_type") == "phase" and o["compliance_model_J"] is not None:
        # A Conformal run without its material interpolation (or from before it) meshes
        # the exported geometry at full modulus: its model compliance is the exported one.
        o["compliance_exported_J"] = o["compliance_model_J"]
        o["compliance_exported_source"] = "model (Conformal without material interpolation analyses the exported geometry)"
    if o["compliance_model_J"] and o["compliance_exported_J"]:
        o["exported_over_model"] = o["compliance_exported_J"] / o["compliance_model_J"]
    if c is not None and len(c):
        bad = ~np.isfinite(c) | (c <= 0)
        # A rejected step's compliance is NaN by design (§5.8): not a non-physical analysis.
        av0 = ctx.col("analysis_valid")
        rejected = (av0 == 0) if (av0 is not None and av0.dtype != object) else np.zeros(len(c), bool)
        o["nonphysical_compliance_iterations"] = int(np.count_nonzero(bad & ~rejected))
        if np.any(bad & ~rejected):
            o["first_nonphysical_compliance_iteration"] = int(ctx.x[np.argmax(bad & ~rejected)])
        c = np.where(bad, np.nan, c)
    if c is not None and len(finite(c)):
        k = int(np.nanargmin(c))
        o["compliance_min_over_run_J"] = float(c[k])
        o["iteration_of_min_compliance"] = int(ctx.x[k])
    av = ctx.col("analysis_valid")
    if av is not None and av.dtype != object and len(av):
        invalid = av == 0
        o["invalid_analysis_iterations"] = int(np.count_nonzero(invalid))
        if np.any(invalid):
            o["first_invalid_analysis_iteration"] = int(ctx.x[np.argmax(invalid)])
            o["last_invalid_analysis_iteration"] = int(ctx.x[len(invalid) - 1 - np.argmax(invalid[::-1])])
    # Step rejection on a cut load path (docs/theory.md §5.8).
    sr = ctx.col("step_rejected")
    if sr is not None and sr.dtype != object and np.any(np.isfinite(sr)):
        o["rejected_steps"] = int(np.nansum(sr))
    elif info.get("rejected_steps") is not None:
        o["rejected_steps"] = info.get("rejected_steps")
    ss = ctx.col("step_scale")
    if ss is not None and ss.dtype != object and np.any(np.isfinite(ss)):
        o["min_step_scale"] = float(np.nanmin(ss))
        o["iterations_with_reduced_step"] = int(np.count_nonzero(ss < 1.0))
    ec = ctx.col("exported_compliance")
    if c is not None and ec is not None and ec.dtype != object:
        ok = np.isfinite(ec) & np.isfinite(c) & (c > 0)
        if np.any(ok):
            ratio = ec[ok] / c[ok]
            xs = ctx.x[ok]
            o["exported_over_model_max"] = float(np.max(ratio))
            o["iteration_of_max_exported_over_model"] = int(xs[int(np.argmax(ratio))])
            above = xs[ratio > 1.1]
            o["last_iteration_exported_over_model_above_1p1"] = int(above[-1]) if above.size else None
    ev = _num(fin.get("exported_volume_fraction"))
    if ev is None and "solid" in f:
        ev = float(np.mean(f["solid"]))
    if o.get("compliance_exported_J") and ev:
        o["compliance_x_volume_exported"] = o["compliance_exported_J"] * ev
    if o.get("compliance_model_J") and fin.get("mean_density"):
        o["compliance_x_volume_model"] = o["compliance_model_J"] * float(fin["mean_density"])

    # --- volume ---
    v = S["volume"]
    v.update(target=ctx.target, mean_density=_num(fin.get("mean_density")), exported_volume_fraction=ev,
             G1=_num(fin.get("G1")))
    if ev is not None:
        v["exported_error_pct_of_target"] = 100.0 * (ev - ctx.target) / ctx.target
    vf = ctx.col("void_film_fraction")
    if vf is not None and len(finite(vf)):
        v["void_film_fraction"] = float(finite(vf)[-1])
        v["void_film_pct_of_target"] = 100.0 * v["void_film_fraction"] / ctx.target

    # --- discreteness ---
    d = S["discreteness"]
    if "phi" in f:
        phi = f["phi"]
        dm = f.get("design_mask", np.ones(phi.shape, bool))
        rho = f.get("rho")
        pd_ = phi[dm]
        band = 0.1 * (ctx.S - ctx.V)
        d["grey_fraction"] = band_fraction(phi, ctx.V + band + 1e-12, ctx.S - band, dm)
        d["undecided_fraction"] = band_fraction(phi, ctx.V + 0.25 * (ctx.S - ctx.V), ctx.S - 0.25 * (ctx.S - ctx.V), dm)
        if rho is not None:
            d["non_discreteness"] = non_discreteness(rho, dm)
        d["near_wells_fraction"] = float(np.mean((np.abs(pd_ - ctx.V) < 0.05) | (np.abs(pd_ - ctx.S) < 0.05))) if pd_.size else None
        d["phi_quantiles"] = {str(q): float(np.quantile(pd_, q)) for q in (0.0, 0.01, 0.05, 0.25, 0.5, 0.75, 0.95, 0.99, 1.0)} if pd_.size else None
        d["phi_below_void_fraction"] = float(np.mean(pd_ < ctx.V - 0.05)) if pd_.size else None
        d["phi_above_solid_fraction"] = float(np.mean(pd_ > ctx.S + 0.05)) if pd_.size else None
    else:
        for key, name in (("grey_fraction", "grey_fraction"), ("non_discreteness", "non_discreteness"),
                          ("undecided_fraction", "undecided_fraction")):
            col = ctx.col(name)
            if col is not None and len(finite(col)):
                d[key] = float(finite(col)[-1])

    # --- crispness and CAD readiness (docs/theory.md §8.4.7): the same definitions for every method ---
    cr = S["crispness"]
    try:
        if "phi" in f:
            m = crispness_metrics(design_fraction(f["phi"], ctx.S, ctx.V), f.get("design_mask"))
        else:
            m = dict(info.get("crispness") or {})
        for k in ("non_discreteness", "grey_fraction", "interface_width_vox", "bulk_grey_fraction",
                  "threshold_volume_sensitivity", "topology_stable", "thin_member_fraction",
                  "median_member_thickness_vox", "interface_area_vox2", "n_tunnels", "n_cavities"):
            if m.get(k) is not None:
                cr[k] = m[k]
        cm, ce = _num(fin.get("compliance")), _num(fin.get("exported_compliance"))
        if cm and ce:
            cr["compliance_exported_over_model"] = ce / cm
        _b = info.get("buckling", {}) or {}
        lm, le = _num(_b.get("last_valid_critical_load_factor")), _num(_b.get("critical_load_factor_exported"))
        if _b.get("enabled") and lm and le and np.isfinite(lm) and np.isfinite(le):
            cr["load_factor_exported_over_model"] = le / lm
    except Exception as e:
        ctx.note(f"crispness metrics failed: {e}")

    # --- the exported STL's quality (docs/theory.md §8.4.8) ---
    S["surface"].update(surface_summary(info))

    # --- topology & connectivity ---
    tp = S["topology"]
    conn = fin.get("connectivity") or {}
    if "solid" in f:
        solid = f["solid"]
        labels = f.get("component_labels")
        if labels is None:
            labels, _ = component_sizes(solid)
        sizes = np.bincount(labels.ravel())[1:] if labels.max() > 0 else np.zeros(0, int)
        topo = topology(solid, n_components=int(np.count_nonzero(sizes)))
        tp.update(euler_characteristic=topo["euler"], n_components=topo["b0"], n_tunnels=topo["b1"],
                  n_cavities=topo["b2"])
        order = np.argsort(sizes)[::-1]
        tp["component_sizes_voxels"] = [int(sizes[i]) for i in order[:20]]
        tp["largest_component_share"] = float(sizes.max() / max(1, sizes.sum())) if sizes.size else None
    tp.update(load_path_connected=conn.get("load_path_connected"), connected=conn.get("connected"),
              detached_fraction=conn.get("detached_fraction"))
    dg = info.get("diagnostics", {})
    tp.update(total_interior_nucleations=dg.get("total_interior_nucleations"),
              total_surface_nucleations=dg.get("total_surface_nucleations"),
              first_interior_nucleation_iteration=dg.get("first_interior_nucleation_iteration"),
              last_topology_change_iteration=dg.get("last_topology_change_iteration"))
    if tp.get("total_interior_nucleations") is None:
        ni = ctx.col("nucleations_interior")
        if ni is not None and len(finite(ni)):
            tp["total_interior_nucleations"] = int(np.nansum(ni))
            tp["total_surface_nucleations"] = int(np.nansum(ctx.col("nucleations_surface")))
    if len(ctx.events.get("iteration", [])):
        kinds = ctx.events["kind"]
        its = ctx.events["iteration"]
        inter = its[kinds == "interior"]
        # Nucleations after the design has formed are the interesting ones:
        # report those after the first 10 % of the run separately.
        if len(ctx.x):
            late = float(np.nanmax(ctx.x)) * 0.1
            tp["interior_nucleations_after_first_10pct"] = int(np.count_nonzero(inter > late))

    # --- geometry ---
    g = S["geometry"]
    if "solid" in f:
        solid = f["solid"]
        n_solid = int(np.count_nonzero(solid))
        faces = interface_faces(solid)
        h = ctx.voxel_mm
        g.update(solid_voxels=n_solid, interface_faces=faces, surface_area_mm2=faces * h * h,
                 volume_mm3=n_solid * h ** 3,
                 surface_to_volume_per_mm=(faces * h * h) / max(1e-300, n_solid * h ** 3) if n_solid else None)
        if n_solid:
            idx = np.argwhere(solid)
            g["bbox_min_voxel"] = idx.min(axis=0).tolist()
            g["bbox_max_voxel"] = idx.max(axis=0).tolist()
            g["centroid_voxel"] = idx.mean(axis=0).tolist()
            samples, _half = thickness_samples(solid)
            if samples.size:
                g["member_thickness_voxels"] = {str(q): float(np.quantile(samples, q)) for q in (0.0, 0.05, 0.25, 0.5, 0.75, 0.95, 1.0)}
                g["member_thickness_mm_median"] = float(np.median(samples)) * h
                g["thin_member_share"] = float(np.mean(samples <= 1.0))
                g["thickness_note"] = ("2d - 1 on the ridge of the distance transform: exact for an odd "
                                       "number of voxels, one voxel low for an even number")
            g["symmetry_mismatch"] = symmetry_mismatch(solid)

    # --- mechanics ---
    m = S["mechanics"]
    if "solid" in f:
        solid = f["solid"]
        for prefix, label in (("exported", "exported"), ("model", "model")):
            vm = f.get(f"{prefix}_Von_Mises_Stress")
            if vm is not None and vm.shape == solid.shape and np.any(solid):
                vs = vm[solid]
                m[f"von_mises_{label}_Pa"] = {"max": float(vs.max()), "p99": float(np.quantile(vs, 0.99)),
                                              "p95": float(np.quantile(vs, 0.95)), "median": float(np.median(vs)),
                                              "mean": float(vs.mean())}
            disp = f.get(f"{prefix}_point_Displacement_Magnitude")
            if disp is not None:
                m[f"max_displacement_{label}_m"] = float(np.max(disp))
        sed = f.get("exported_Exported_SED")
        src = "exported"
        if sed is None:
            sed = f.get("F0")
            src = "model F0"
        if sed is not None and sed.shape == solid.shape and np.any(solid):
            vals = sed[solid]
            _, _, gini = lorenz_curve(vals)
            m["energy_source"] = src
            m["energy_gini"] = gini
            m["solid_share_carrying_50pct_energy"] = share_of_volume_carrying(vals, 0.5)
            m["solid_share_carrying_90pct_energy"] = share_of_volume_carrying(vals, 0.9)
            m["sed_solid_J_per_m3"] = {"max": float(vals.max()), "median": float(np.median(vals)),
                                       "mean": float(vals.mean())}

    # --- optimality check (eq. 8.4m) ---
    st = f.get("stiffness_term")
    l1t = ctx.col("lambda1_tilde")
    if st is not None and "solid" in f and l1t is not None and np.any(np.isfinite(l1t)):
        dm = f.get("design_mask", np.ones(st.shape, bool))
        pressure = float(finite(l1t)[-1]) * (1.0 - ctx.c1)
        solid_d = f["solid"] & dm
        void_d = (~f["solid"]) & dm
        m["volume_pressure_final"] = pressure
        # The buckling constraint is part of the balance: the update applies
        # S - lambda1~ rho' - lambda3~ b, and in a critically compressed member
        # lambda3~ b < 0 ADDS material. Leaving it out flagged 89 % of the
        # Column_Buckling solid as on the wrong side, against 10 % with it.
        drive = st
        bt, l3t = f.get("buckling_term"), ctx.col("lambda3_tilde")
        if (bt is not None and bt.shape == st.shape and l3t is not None and l3t.dtype != object
                and finite(l3t).size):
            applied = float(finite(l3t)[-1]) * bt
            drive = st - applied
            m["buckling_term_in_balance"] = True
        if np.any(solid_d):
            m["wrong_side_share_solid"] = float(np.mean(drive[solid_d] < pressure))
        if np.any(void_d):
            m["wrong_side_share_void"] = float(np.mean(drive[void_d] > pressure))

    # --- buckling ---
    b = S["buckling"]
    bi = info.get("buckling", {})
    if bi.get("enabled") or ctx.buckling or bi.get("critical_load_factor_exported") is not None:
        b.update(status=bi.get("status"), safety_factor=bi.get("safety_factor"),
                 constraint_measure=bi.get("constraint_measure"),
                 critical_load_factor_model=_num(bi.get("critical_load_factor_model")),
                 aggregated_load_factor_model=_num(bi.get("aggregated_load_factor_model")),
                 critical_load_factor_exported=_num(bi.get("critical_load_factor_exported")),
                 exported_load_factors=bi.get("exported_load_factors"),
                 unknown_evaluations=bi.get("unknown_evaluations"))
        if len(ctx.buck.get("iteration", [])):
            b["evaluations"] = int(len(ctx.buck["iteration"]))
            st = ctx.buck.get("status")
            if st is not None and st.dtype == object:
                b["evaluations_by_status"] = {k: int(np.count_nonzero(st == k)) for k in np.unique(st)}

    # --- multipliers ---
    mu = S["multipliers"]
    for k in ("lambda1", "mu1", "mu2", "lambda2_min", "lambda2_mean", "lambda2_max", "penalisation_power"):
        if fin.get(k) is not None:
            mu[k] = fin.get(k)
    if bi.get("lambda3") is not None:
        mu["lambda3"] = bi.get("lambda3")
        mu["mu3"] = bi.get("mu3")

    # --- convergence at the end ---
    cv = S["convergence"]
    cv.update(settled_count=fin.get("settled_count"), window=ctx.window, settle_failures=fin.get("settle_failures"),
              grey_fraction=fin.get("grey_fraction"))
    for name in ("max_net_force", "relative_compliance_change", "phase_switches"):
        col = ctx.col(name)
        if col is not None and len(finite(col)):
            cv[f"final_{name}"] = float(finite(col)[-1])
    sf = ctx.col("settle_failures")
    if sf is not None and sf.dtype == object and len(sf):
        tail = sf[-min(len(sf), 200):]
        cv["failing_share_last_200"] = {t: float(np.mean([t in str(x).split(";") for x in tail])) for t in SETTLE_TESTS}

    # --- threshold sweep ---
    if len(ctx.sweep.get("level", [])):
        sw = ctx.sweep
        lv = sw["level"]
        k5 = int(np.argmin(np.abs(lv - 0.5)))
        comp = sw.get("compliance")
        S["threshold_sweep"]["levels"] = lv.tolist()
        if comp is not None and np.any(np.isfinite(comp)) and np.isfinite(comp[k5]):
            rel = comp / comp[k5]
            S["threshold_sweep"]["compliance_relative_to_exported"] = [None if not np.isfinite(x) else float(x) for x in rel]
            inner = (lv >= 0.3 - 1e-9) & (lv <= 0.7 + 1e-9) & np.isfinite(rel)
            if np.any(inner):
                S["threshold_sweep"]["max_compliance_change_0p3_to_0p7"] = float(np.max(np.abs(rel[inner] - 1.0)))
        vfr = sw.get("volume_fraction")
        if vfr is not None:
            S["threshold_sweep"]["volume_fraction"] = vfr.tolist()
            inner = (lv >= 0.3 - 1e-9) & (lv <= 0.7 + 1e-9)
            if np.any(inner):
                S["threshold_sweep"]["volume_spread_0p3_to_0p7"] = float(np.ptp(vfr[inner]))
        for key in ("n_components", "n_tunnels", "n_cavities", "connected", "critical_load_factor"):
            if key in sw:
                S["threshold_sweep"][key] = [None if (isinstance(x, float) and not np.isfinite(x)) else x
                                             for x in np.asarray(sw[key]).tolist()]
    return S


def _fmt(v, nd=4):
    if v is None:
        return "-"
    if isinstance(v, bool):
        return "yes" if v else "no"
    if isinstance(v, (int, np.integer)):
        return str(int(v))
    if isinstance(v, (float, np.floating)):
        v = float(v)
        if not np.isfinite(v):
            return "inf" if v > 0 else ("nan" if np.isnan(v) else "-inf")
        if v != 0 and (abs(v) < 1e-3 or abs(v) >= 1e5):
            return f"{v:.{nd}e}"
        return f"{v:.{nd}g}"
    if isinstance(v, dict):
        return ", ".join(f"{k}: {_fmt(x, nd)}" for k, x in v.items())
    if isinstance(v, (list, tuple)):
        return "[" + ", ".join(_fmt(x, nd) for x in v[:12]) + (", ..." if len(v) > 12 else "") + "]"
    return str(v)


SUMMARY_LABELS = {
    "run": "Run", "timing": "Timing", "objective": "Objective (compliance)", "volume": "Volume",
    "discreteness": "Discreteness", "crispness": "Crispness and CAD readiness (theory.md §8.4.7)",
    "topology": "Topology and connectivity", "geometry": "Geometry",
    "surface": "Exported STL quality (theory.md §8.4.8)",
    "mechanics": "Mechanics of the final design", "buckling": "Buckling", "multipliers": "Multipliers",
    "convergence": "Convergence state at the end", "threshold_sweep": "Threshold sweep",
}


def summary_text(S: dict) -> str:
    label = {"simp": "SIMP", "levelset": "Level-set"}.get(str(S["run"].get("method")), "PCLS")
    lines = [f"{label} run summary: {S['run'].get('name')}", "=" * 72]
    for sec, label in SUMMARY_LABELS.items():
        d = S.get(sec) or {}
        if not d:
            continue
        lines.append("")
        lines.append(label)
        lines.append("-" * len(label))
        for k, v in d.items():
            lines.append(f"  {k:42s} {_fmt(v)}")
    return "\n".join(lines) + "\n"


def summary_figures(ctx: RunContext, S: dict):
    """The summary as one or more text pages for the report."""
    text = summary_text(S).splitlines()
    per_page = 78
    figs = []
    for i in range(0, len(text), per_page):
        fig = new_figure(8.27, 11.69)
        fig.text(0.04, 0.98, "\n".join(text[i:i + per_page]), family="monospace", fontsize=7.2,
                 va="top", ha="left")
        figs.append((f"00_summary_p{i // per_page + 1}", fig))
    return figs


# ---------------------------------------------------------------------------
# Convergence history figures
# ---------------------------------------------------------------------------

def _plot_line(ax, ctx, name, label=None, color=None, **kw):
    y = ctx.col(name)
    if y is None or y.dtype == object or not np.any(np.isfinite(y)):
        return None
    return ax.plot(ctx.x, y, label=label or name, color=color, lw=kw.pop("lw", 1.1), **kw)


def _volume_panel(ax, ctx: RunContext, field_label, export_label, film=True, band=False, log=True):
    """
    The constrained (field) and exported volume, the void film and the target, with the
    final exported error annotated, on a LOG axis (or linear with log=False). On a
    linear axis running from the initial volume (~1) down to a target of a few percent,
    an export at half the target is a hair's breadth from the target line
    (Truss_Cantilever_19453760: 0.0128 against 0.025 looked converged).
    """
    fv = _plot_line(ax, ctx, "field_volume_fraction", field_label, C_MODEL)
    ev = _plot_line(ax, ctx, "exported_volume_fraction", export_label, C_EXPORT)
    film_y = ctx.col("void_film_fraction") if film else None
    if film_y is not None and film_y.dtype != object and np.any(np.isfinite(film_y) & (film_y > 0)):
        ax.plot(ctx.x, np.where(film_y > 0, film_y, np.nan), color=C_ALT, lw=1.0, ls=":",
                label="void film (counted by $G_1$, not exported)")
    ax.axhline(ctx.target, color=C_TARGET, ls="--", lw=0.9, label=f"target $V_0$ = {ctx.target:g}")
    if band:
        ax.axhspan(max(ctx.target - ctx.tol_outer, 1e-12), ctx.target + ctx.tol_outer, color="0.92", zorder=0)
    vals = [finite(ctx.col(c)) for c in ("field_volume_fraction", "exported_volume_fraction") if ctx.has(c)]
    vals = np.concatenate(vals) if vals else np.zeros(0)
    if log and vals.size and np.all(vals > 0):
        ax.set_yscale("log")
    exp_col = ctx.col("exported_volume_fraction")
    if ev is not None and ctx.target > 0:
        last = finite(exp_col)
        if last.size:
            err = 100.0 * (last[-1] - ctx.target) / ctx.target
            bad = abs(err) > 100.0 * float(ctx.param("convergence_exported_volume_tol", 0.15) or 0.15)
            ax.annotate(f"exported {last[-1]:.4g} ({err:+.1f}% of $V_0$)", xy=(ctx.x[-1], last[-1]),
                        xytext=(-6, -14 if (log and last[-1] < ctx.target) else 8), textcoords="offset points",
                        ha="right", fontsize=8, color=C_BAD if bad else C_EXPORT,
                        fontweight="bold" if bad else "normal")


def fig_dashboard(ctx: RunContext):
    if not len(ctx.x):
        return None
    fig = new_figure(15, 13)
    axes = fig.subplots(4, 3).ravel()
    fig.suptitle(f"{ctx.name}: convergence overview", fontsize=13)
    # 1 compliance
    ax = axes[0]
    cc = ctx.col("compliance")
    if cc is not None:
        cc = np.where(np.isfinite(cc) & (cc > 0), cc, np.nan)
        ax.plot(ctx.x, cc, color=C_MODEL, lw=1.1, label="model")
    ec = ctx.col("exported_compliance")
    if ec is not None and ec.dtype != object and np.any(np.isfinite(ec)):
        ax.plot(ctx.x, ec, ".", color=C_EXPORT, ms=3, label="exported")
    ce = _num(ctx.info.get("final", {}).get("exported_compliance"))
    if ce:
        ax.plot([ctx.x[-1]], [ce], "*", color=C_EXPORT, ms=11, label="exported (final)")
    logy_if_positive(ax, cc, ec)
    ax.set_title("Compliance $c = f^\\top u$ [J]")
    _legend(ax, fontsize=7)
    # 2 volume
    ax = axes[1]
    _volume_panel(ax, ctx, r"$\overline{\rho}$ (constraint)", "exported")
    ax.set_title("Volume fraction (log)")
    _legend(ax, fontsize=7)
    # 3 G1
    ax = axes[2]
    _plot_line(ax, ctx, "G1", "$G_1$", C_MODEL)
    ax.axhspan(-ctx.tol_outer, ctx.tol_outer, color="0.9", zorder=0, label="settled tolerance")
    ax.axhline(0, color="k", lw=0.5)
    ax.set_title(r"Volume error $G_1 = \overline{\rho} - V_0$")
    _legend(ax, fontsize=7)
    # 4 discreteness
    ax = axes[3]
    _plot_line(ax, ctx, "grey_fraction", r"grey $\gamma$ (5.7)", C_MODEL)
    _plot_line(ax, ctx, "undecided_fraction", "undecided (1.25-1.75)", C_ALT)
    _plot_line(ax, ctx, "non_discreteness", r"$M_{nd}$", C_ALT2)
    ax.axhline(ctx.tol_grey, color=C_TARGET, ls="--", lw=0.9, label="grey tolerance")
    logy_if_positive(ax, ctx.col("grey_fraction"))
    ax.set_title("Discreteness (share of design voxels)")
    _legend(ax, fontsize=7)
    # 5 p and settled count
    ax = axes[4]
    _plot_line(ax, ctx, "penalisation_power", "p", C_MODEL)
    ax.set_ylabel("p")
    ax2 = ax.twinx()
    sc = ctx.col("settled_count")
    if sc is not None:
        ax2.fill_between(ctx.x, 0, sc, color=C_EXPORT, alpha=0.3, step="post", label="settled count")
        ax2.axhline(ctx.window, color=C_EXPORT, ls=":", lw=0.9)
        ax2.set_ylabel("settled count", color=C_EXPORT)
    ax.set_title("Penalisation power and settled count")
    # 6 topology
    ax = axes[5]
    _plot_line(ax, ctx, "n_components", "$b_0$ pieces", C_MODEL, drawstyle="steps-post")
    _plot_line(ax, ctx, "n_tunnels", "$b_1$ tunnels", C_EXPORT, drawstyle="steps-post")
    _plot_line(ax, ctx, "n_cavities", "$b_2$ cavities", C_ALT, drawstyle="steps-post")
    logy_if_positive(ax, ctx.col("n_components"), ctx.col("n_tunnels"))
    ax.set_title("Topology of the exported solid")
    _legend(ax, fontsize=7)
    # 7 connectivity
    ax = axes[6]
    df = ctx.col("detached_fraction")
    lp = ctx.col("load_path_connected")
    if df is not None:
        ax.plot(ctx.x, 100 * df, color=C_MODEL, lw=1.1, label="detached solid [%]")
    if lp is not None:
        broken = np.where(lp == 0, 1.0, 0.0)
        ax.fill_between(ctx.x, 0, 1, where=broken > 0, transform=ax.get_xaxis_transform(), color=C_BAD,
                        alpha=0.12, step="post", label="load path broken")
    av = ctx.col("analysis_valid")
    if av is not None and av.dtype != object and np.any(av == 0):
        ax.fill_between(ctx.x, 0, 1, where=av == 0, transform=ax.get_xaxis_transform(), color=C_ALT,
                        alpha=0.25, step="post", label="analysed geometry cut (step rejected)")
    ax.set_title("Connectivity of the exported solid")
    _legend(ax, fontsize=7)
    # 8 lambda1
    ax = axes[7]
    _plot_line(ax, ctx, "lambda1", r"$\lambda_1$", C_MODEL)
    _plot_line(ax, ctx, "lambda1_tilde", r"$\tilde\lambda_1$", C_EXPORT, alpha=0.7)
    ax.set_title("Volume multiplier")
    _legend(ax, fontsize=7)
    # 9 penalties
    ax = axes[8]
    _plot_line(ax, ctx, "mu1", r"$\mu_1$", C_MODEL)
    _plot_line(ax, ctx, "mu2", r"$\mu_2$", C_EXPORT)
    ax.set_yscale("log")
    ax.set_title("Penalty parameters")
    _legend(ax, fontsize=7)
    # 10 lambda2
    ax = axes[9]
    lo, me, hi = ctx.col("lambda2_min"), ctx.col("lambda2_mean"), ctx.col("lambda2_max")
    if me is not None:
        ax.fill_between(ctx.x, lo, hi, color=C_MODEL, alpha=0.2, label="min-max")
        ax.plot(ctx.x, me, color=C_MODEL, lw=1.1, label="mean")
    ax.set_title(r"Piecewise multiplier field $\lambda_2$")
    _legend(ax, fontsize=7)
    # 11 stationarity
    ax = axes[10]
    _plot_line(ax, ctx, "max_net_force", r"max $|\Delta\phi|/\Delta t$", C_MODEL)
    ax.axhline(ctx.tol_inner, color=C_TARGET, ls="--", lw=0.9, label="inner threshold")
    if ctx.has("max_net_force"):
        ax.set_yscale("log")
    ps = ctx.col("phase_switches")
    if ps is not None:
        ax2 = ax.twinx()
        ax2.plot(ctx.x, np.maximum(ps, 0.5), color=C_EXPORT, lw=0.8, alpha=0.7)
        ax2.set_yscale("log")
        ax2.set_ylabel("phase switches", color=C_EXPORT)
    ax.set_title("Stationarity")
    _legend(ax, fontsize=7, loc="upper right")
    # 12 time
    ax = axes[11]
    _plot_line(ax, ctx, "total_iteration_time", "iteration", C_MODEL)
    _plot_line(ax, ctx, "analysis_time", "analysis", C_EXPORT)
    _plot_line(ax, ctx, "buckling_time", "buckling", C_ALT)
    logy_if_positive(ax, ctx.col("total_iteration_time"))
    ax.set_title("Time per iteration [s]")
    _legend(ax, fontsize=7)
    for a in axes:
        decorate(a, ctx)
    fig.tight_layout(rect=(0, 0, 1, 0.97))
    return fig


def fig_objective(ctx: RunContext):
    if not ctx.has("compliance"):
        return None
    ec = ctx.col("exported_compliance")
    has_ec = ec is not None and ec.dtype != object and np.any(np.isfinite(ec))
    fig = new_figure(11, 10.5 if has_ec else 7.5)
    if has_ec:
        ax1, ax3, ax2 = fig.subplots(3, 1, sharex=True)
    else:
        ax1, ax2 = fig.subplots(2, 1, sharex=True)
    c = ctx.col("compliance")
    bad = ~np.isfinite(c) | (c <= 0)
    if np.any(bad):
        # A singular FE model (a piece no support reaches) gives f.u <= 0.
        c = np.where(bad, np.nan, c)
        for x0 in ctx.x[bad]:
            ax1.axvline(x0, color=C_BAD, lw=0.5, alpha=0.5, zorder=0)
        ax1.plot([], [], color=C_BAD, lw=0.8, label=f"non-physical compliance ({int(bad.sum())} iterations)")
    ax1.plot(ctx.x, c, color=C_MODEL, lw=1.2, label="model compliance (logged)")
    if has_ec:
        ax1.plot(ctx.x, ec, "o", color=C_EXPORT, ms=3, label="exported geometry (every few iterations)")
        ok = np.isfinite(ec) & np.isfinite(c) & (c > 0)
        ratio = np.where(ok, ec / np.where(ok, c, 1.0), np.nan)
        ax3.plot(ctx.x[ok], ratio[ok], "o-", color=C_EXPORT, ms=3, lw=0.9)
        ax3.axhline(1.0, color="k", lw=0.6)
        if np.nanmax(ratio) / max(np.nanmin(ratio[ok]) if np.any(ok) else 1.0, 1e-12) > 3:
            ax3.set_yscale("log")
        ax3.set_ylabel("exported / model")
        ax3.set_title("Exported part / model: > 1 while grey material carries load that thresholding removes")
        decorate(ax3, ctx, xlabel=False)
    ce = _num(ctx.info.get("final", {}).get("exported_compliance"))
    if ce:
        ax1.plot([ctx.x[-1]], [ce], "*", color=C_EXPORT, ms=12, label=f"exported design: {ce:.4g} J")
    if np.any(np.isfinite(c)):
        k = int(np.nanargmin(c))
        ax1.plot([ctx.x[k]], [c[k]], "v", color=C_ALT, ms=8, label=f"minimum {c[k]:.4g} J (it {int(ctx.x[k])})")
    logy_if_positive(ax1, c, ec if has_ec else None)
    ax1.set_ylabel("compliance [J]")
    if ctx.has("penalisation_power"):
        a = ax1.twinx()
        a.plot(ctx.x, ctx.col("penalisation_power"), color="0.5", lw=0.8, ls=":")
        a.set_ylabel("p (dotted)")
    _legend(ax1, fontsize=8)
    ax1.set_title("Objective: compliance of the analysed model; star = the exported (thresholded) design")
    decorate(ax1, ctx, xlabel=False)
    rc = ctx.col("relative_compliance_change")
    if rc is None and len(c) > 1:
        rc = np.r_[np.nan, np.abs(np.diff(c)) / np.maximum(np.abs(c[1:]), 1e-300)]
    if rc is not None:
        ax2.plot(ctx.x, np.where(np.isfinite(rc) & (rc > 0), rc, np.nan), color=C_MODEL, lw=0.9)
        ax2.axhline(ctx.tol_comp, color=C_TARGET, ls="--", lw=0.9, label="settled tolerance")
        ax2.set_yscale("log")
        ax2.set_ylabel("|relative change| per iteration")
        _legend(ax2, fontsize=8)
    decorate(ax2, ctx)
    fig.tight_layout()
    return fig


def fig_volume(ctx: RunContext):
    if not ctx.has("field_volume_fraction"):
        return None
    fig = new_figure(11, 11)
    ax0, ax1, ax2 = fig.subplots(3, 1, sharex=True)
    for ax, log in ((ax0, False), (ax1, True)):
        _volume_panel(ax, ctx, r"$\overline{\rho}$ (what $G_1$ constrains)",
                      r"exported ($\phi \geq \phi_{LS}$ + preserve)", band=True, log=log)
        scale = "log" if log else "linear"
        ax.set_ylabel(f"fraction of the domain ({scale})")
        _legend(ax, fontsize=8)
        ax.set_title(f"Volume: constrained vs exported, {scale} axis "
                     f"(grey band = settled tolerance on $G_1$)")
        decorate(ax, ctx, xlabel=False)
    if ctx.has("void_film_fraction"):
        ax2.plot(ctx.x, 100 * ctx.col("void_film_fraction") / ctx.target, color=C_ALT, lw=1.1,
                 label="void film (density above $c_1$ below the level set)")
    ev, fv = ctx.col("exported_volume_fraction"), ctx.col("field_volume_fraction")
    if ev is not None and fv is not None:
        ax2.plot(ctx.x, 100 * (ev - ctx.target) / ctx.target, color=C_EXPORT, lw=1.1, label="exported - target")
        ax2.plot(ctx.x, 100 * (fv - ctx.target) / ctx.target, color=C_MODEL, lw=1.1, label=r"$\overline{\rho}$ - target")
    ax2.axhline(0, color="k", lw=0.5)
    tol_pct = 100.0 * float(ctx.param("convergence_exported_volume_tol", 0.15) or 0.15)
    ax2.axhspan(-tol_pct, tol_pct, color="0.92", zorder=0, label=f"exported-volume tolerance $\\pm${tol_pct:g}%")
    # Symmetric log: the start (+3900 % for a 2.5 % target) no longer flattens the end of the run.
    ax2.set_yscale("symlog", linthresh=10.0)
    ax2.set_ylabel("% of $V_0$ (symlog, linear within $\\pm$10%)")
    _legend(ax2, fontsize=8)
    ax2.set_title("Where the volume is (docs/theory.md §5.5)")
    decorate(ax2, ctx)
    fig.tight_layout()
    return fig


def _history_frames(ctx: RunContext, max_frames=None):
    files = ctx.history
    if max_frames and len(files) > max_frames:
        idx = np.unique(np.linspace(0, len(files) - 1, max_frames).round().astype(int))
        files = [files[i] for i in idx]
    for path in files:
        try:
            with np.load(path) as z:
                yield int(z["iteration"]), z["phi"].astype(np.float32)
        except Exception:
            continue


def fig_discreteness(ctx: RunContext):
    from matplotlib import colors as mcolors
    have_cols = ctx.has("grey_fraction")
    if not have_cols and not ctx.history:
        return None
    fig = new_figure(11, 10.5)
    ax1, ax2, ax3 = fig.subplots(3, 1)
    if have_cols:
        _plot_line(ax1, ctx, "grey_fraction", r"grey $\gamma$: $\phi \in (1.1, 1.9)$", C_MODEL)
        _plot_line(ax1, ctx, "undecided_fraction", r"undecided: $\phi \in [1.25, 1.75)$", C_ALT)
        _plot_line(ax1, ctx, "non_discreteness", r"$M_{nd} = 4\,\overline{\rho(1-\rho)}$", C_ALT2)
        ax1.axhline(ctx.tol_grey, color=C_TARGET, ls="--", lw=0.9, label="grey tolerance")
        logy_if_positive(ax1, ctx.col("grey_fraction"), ctx.col("non_discreteness"))
        _legend(ax1, fontsize=8)
        ax1.set_title("Discreteness of the design voxels")
        decorate(ax1, ctx)
    lo, me, hi = ctx.col("phi_min"), ctx.col("phi_mean"), ctx.col("phi_max")
    if me is not None:
        ax2.fill_between(ctx.x, lo, hi, color=C_MODEL, alpha=0.15, label=r"$\phi$ min-max")
        ax2.plot(ctx.x, me, color=C_MODEL, lw=1.1, label=r"$\phi$ mean")
        for yv in (ctx.V, ctx.ls, ctx.S):
            ax2.axhline(yv, color="0.5", lw=0.6, ls=":")
        _legend(ax2, fontsize=8)
        ax2.set_title(r"Range of $\phi$ (wells at V and S, level set dotted between)")
        decorate(ax2, ctx)
    # Distribution of phi over time from the history snapshots.
    if ctx.history and ctx.fields.get("design_mask") is not None:
        dm = ctx.fields["design_mask"]
        bins = np.linspace(0.0, max(3.0, ctx.S + 1.0), 121)
        its, hists = [], []
        for it, phi in _history_frames(ctx, max_frames=400):
            h, _ = np.histogram(phi[dm], bins=bins)
            its.append(it)
            hists.append(h)
        if hists:
            H = np.array(hists, dtype=float).T
            H = H / max(1, dm.sum())
            its = np.array(its)
            edges_x = np.r_[its, its[-1] + max(1, (its[-1] - its[0]) / max(1, len(its) - 1))]
            im = ax3.pcolormesh(edges_x, bins, np.where(H > 0, H, np.nan), shading="flat",
                                norm=mcolors.LogNorm(vmin=max(1e-5, np.nanmin(H[H > 0])), vmax=1.0),
                                cmap="viridis")
            fig.colorbar(im, ax=ax3, label="share of design voxels")
            ax3.axhline(ctx.ls, color="w", lw=0.6, ls=":")
            ax3.set_ylim(bins[0], bins[-1])
            ax3.set_ylabel(r"$\phi$")
            ax3.set_title(r"Distribution of $\phi$ over the run (history snapshots)")
            decorate(ax3, ctx)
    else:
        ax3.set_visible(False)
    fig.tight_layout()
    return fig


def fig_crispness(ctx: RunContext):
    """
    Discreteness and interface crispness over the run (docs/theory.md §8.4.7), from the history
    snapshots: the interface width W (eq. 8.22), M_nd and the grey and bulk-grey shares (8.20, 8.21,
    8.23), the threshold sensitivity S_V (8.24); and the distribution of the design fraction x at the end.
    """
    if not ctx.history and "phi" not in ctx.fields:
        return None
    rows = []
    try:
        rows = history_metrics(ctx.run_dir, max_frames=60, light=True)
    except Exception as e:
        ctx.note(f"crispness history failed: {e}")
    fig = new_figure(11, 8.5)
    (a1, a2), (a3, a4) = fig.subplots(2, 2)
    if rows:
        its = np.array([r[0] for r in rows], dtype=float)
        get = lambda k: np.array([np.nan if r[1].get(k) is None else r[1][k] for r in rows], dtype=float)  # noqa: E731
        a1.plot(its, get("interface_width_vox"), color=C_MODEL, lw=1.2, marker=".", ms=3)
        a1.axhline(1.0, color=C_TARGET, ls=":", lw=0.8, label="one voxel")
        a1.set_title(r"Interface width $W$ [voxels] (0 = sharp; eq. 8.22)")
        _legend(a1, fontsize=8)
        decorate(a1, ctx)
        for k, lab, col in (("non_discreteness", r"$M_{nd}$", C_MODEL), ("grey_fraction", "grey (0.1 < x < 0.9)", C_ALT),
                            ("bulk_grey_fraction", "bulk grey (> 2 voxels from the interface)", C_ALT2)):
            a2.plot(its, get(k), color=col, lw=1.1, label=lab)
        logy_if_positive(a2, get("non_discreteness"), get("grey_fraction"))
        a2.set_title("Discreteness of the design voxels")
        _legend(a2, fontsize=8)
        decorate(a2, ctx)
        a3.plot(its, get("threshold_volume_sensitivity"), color=C_EXPORT, lw=1.2)
        a3.set_title(r"Threshold sensitivity $S_V = (V_{0.3} - V_{0.7}) / V_{0.5}$ (eq. 8.24)")
        decorate(a3, ctx)
    else:
        for a in (a1, a2, a3):
            a.set_visible(False)
    if "phi" in ctx.fields:
        dm = ctx.fields.get("design_mask", np.ones(ctx.fields["phi"].shape, bool))
        x = design_fraction(ctx.fields["phi"], ctx.S, ctx.V)[dm]
        a4.hist(x, bins=np.linspace(0, 1, 41), color=C_MODEL, log=True)
        a4.axvline(0.5, color=C_TARGET, ls="--", lw=0.8)
        a4.set_xlabel("design fraction x of the design voxels (final design)")
        a4.set_ylabel("voxels")
        a4.set_title("Distribution of x at the end (0/1 = crisp)")
    else:
        a4.set_visible(False)
    fig.tight_layout()
    return fig


def fig_topology(ctx: RunContext):
    if not ctx.has("n_components"):
        return None
    fig = new_figure(11, 9)
    ax1, ax2, ax3 = fig.subplots(3, 1, sharex=True)
    _plot_line(ax1, ctx, "n_components", "$b_0$ pieces", C_MODEL, drawstyle="steps-post")
    _plot_line(ax1, ctx, "n_tunnels", "$b_1$ tunnels (handles)", C_EXPORT, drawstyle="steps-post")
    _plot_line(ax1, ctx, "n_cavities", "$b_2$ enclosed cavities", C_ALT, drawstyle="steps-post")
    logy_if_positive(ax1, ctx.col("n_components"), ctx.col("n_tunnels"))
    _legend(ax1, fontsize=8)
    ax1.set_title("Betti numbers of the exported solid (26/6 connectivity, eq. 11.19)")
    decorate(ax1, ctx, xlabel=False)
    _plot_line(ax2, ctx, "euler_characteristic", r"$\chi = V - E + F - C = b_0 - b_1 + b_2$", C_MODEL, drawstyle="steps-post")
    _legend(ax2, fontsize=8)
    ax2.set_title("Euler characteristic")
    decorate(ax2, ctx, xlabel=False)
    ni, ns = ctx.col("nucleations_interior"), ctx.col("nucleations_surface")
    if ni is not None and ns is not None:
        ax3.bar(ctx.x, np.nan_to_num(ni), width=1.0, color=C_MODEL, label="interior nucleations")
        ax3.bar(ctx.x, np.nan_to_num(ns), width=1.0, bottom=np.nan_to_num(ni), color=C_EXPORT,
                label="surface pits")
        if np.nanmax(np.nan_to_num(ni) + np.nan_to_num(ns)) > 20:
            ax3.set_yscale("symlog", linthresh=1)
        _legend(ax3, fontsize=8)
        ax3.set_title("New holes per iteration (eq. 11.17-11.18)")
    decorate(ax3, ctx)
    fig.tight_layout()
    return fig


def fig_multipliers(ctx: RunContext):
    if not ctx.has("lambda1"):
        return None
    rows = 4 if ctx.buckling else 3
    fig = new_figure(11, 2.8 * rows)
    axes = fig.subplots(rows, 1, sharex=True)
    ax = axes[0]
    _plot_line(ax, ctx, "lambda1", r"$\lambda_1$", C_MODEL)
    _plot_line(ax, ctx, "lambda1_tilde", r"$\tilde\lambda_1 = \lambda_1 + G_1/\mu_1$", C_EXPORT, alpha=0.8)
    ax.axhline(0, color="k", lw=0.5)
    _legend(ax, fontsize=8)
    ax.set_title("Volume multiplier (units of the normalised stiffness driver)")
    ax = axes[1]
    _plot_line(ax, ctx, "mu1", r"$\mu_1$", C_MODEL)
    _plot_line(ax, ctx, "mu2", r"$\mu_2$", C_EXPORT)
    # PCLS and the level set normalise by a mean strain energy density [J/m^3]; SIMP's scale is
    # Lambda * mean(dV/dx), an energy [J] (docs/theory.md §12.2).
    scale_unit = "J" if ctx.method == "simp" else "J/m$^3$"
    _plot_line(ax, ctx, "stiffness_normalisation_scale", f"stiffness scale $m$ [{scale_unit}]", C_ALT)
    ax.set_yscale("log")
    _legend(ax, fontsize=8)
    ax.set_title("Penalties and the stiffness normalisation scale")
    ax = axes[2]
    lo, me, hi = ctx.col("lambda2_min"), ctx.col("lambda2_mean"), ctx.col("lambda2_max")
    if me is not None:
        ax.fill_between(ctx.x, lo, hi, color=C_MODEL, alpha=0.2, label="min-max")
        ax.plot(ctx.x, me, color=C_MODEL, lw=1.1, label="mean")
    _legend(ax, fontsize=8)
    ax.set_title(r"Piecewise multiplier field $\lambda_2$ (dashed red: reconnection resets)")
    if ctx.buckling:
        ax = axes[3]
        _plot_line(ax, ctx, "lambda3", r"$\lambda_3$", C_MODEL)
        _plot_line(ax, ctx, "lambda3_tilde", r"$\tilde\lambda_3$ (applied)", C_EXPORT)
        ax2 = ax.twinx()
        m3 = ctx.col("mu3")
        if m3 is not None and np.any(m3 > 0):
            ax2.plot(ctx.x, m3, color=C_ALT, lw=0.9, ls=":")
            ax2.set_yscale("log")
            ax2.set_ylabel(r"$\mu_3$ (dotted)", color=C_ALT)
        _legend(ax, fontsize=8)
        ax.set_title("Buckling multiplier")
    for i, a in enumerate(axes):
        decorate(a, ctx, xlabel=(i == rows - 1))
    fig.tight_layout()
    return fig


def fig_terms(ctx: RunContext):
    terms = [("stiffness_term", "Stiffness term (adds material)"), ("volume_term", r"Volume term $\tilde\lambda_1\rho'$"),
             ("piecewise_term", r"Piecewise term $\tilde\lambda_2 K'$"), ("buckling_term", r"Applied buckling term $\tilde\lambda_3\,\partial\tilde\mu$"),
             ("diffusion_term", r"Diffusion (AOS curvature) term $(\phi^{n+1} - \phi^*)/\Delta t$")]
    terms = [t for t in terms if ctx.has(f"{t[0]}_mean")]
    if not terms:
        return None
    fig = new_figure(11, 2.6 * len(terms))
    axes = np.atleast_1d(fig.subplots(len(terms), 1, sharex=True))
    for ax, (name, title) in zip(axes, terms):
        lo, me, hi = ctx.col(f"{name}_min"), ctx.col(f"{name}_mean"), ctx.col(f"{name}_max")
        ax.fill_between(ctx.x, lo, hi, color=C_MODEL, alpha=0.2, label="min-max")
        ax.plot(ctx.x, me, color=C_MODEL, lw=1.1, label="mean")
        ax.axhline(0, color="k", lw=0.5)
        ax.set_title(title)
        _legend(ax, fontsize=7)
        decorate(ax, ctx, xlabel=(ax is axes[-1]))
    fig.suptitle("Driving terms of the update $\\phi^* = \\phi + \\Delta t\\,(\\ldots)$ (eq. 5.1)")
    fig.tight_layout(rect=(0, 0, 1, 0.97))
    return fig


def fig_stationarity(ctx: RunContext):
    if not ctx.has("phase_switches"):
        return None
    ss = ctx.col("step_scale")
    # The step controller of docs/theory.md §5.8, shown only when it acted.
    show_step = ss is not None and ss.dtype != object and bool(np.any(np.isfinite(ss) & (ss < 1.0)))
    n_panels = 5 if show_step else 4
    fig = new_figure(11, 10 + (2.2 if show_step else 0.0))
    axes = fig.subplots(n_panels, 1, sharex=True)
    ax = axes[0]
    if _plot_line(ax, ctx, "max_net_force", r"max $|\Delta\phi|/\Delta t$ (design voxels)", C_MODEL):
        ax.axhline(ctx.tol_inner, color=C_TARGET, ls="--", lw=0.9, label="inner-loop threshold (5.4)")
        ax.set_yscale("log")
        _legend(ax, fontsize=8)
    ax.set_title("Largest net driving force")
    ax = axes[1]
    if _plot_line(ax, ctx, "mean_abs_dphi", r"mean $|\Delta\phi|$", C_MODEL):
        ax.set_yscale("log")
    ax.set_title("Mean change of $\\phi$ per iteration")
    ax = axes[2]
    ps = ctx.col("phase_switches")
    ax.plot(ctx.x, np.where(ps > 0, ps, np.nan), color=C_MODEL, lw=1.0, label="voxels crossing the level set")
    if ctx.n_design:
        ax.axhline(ctx.tol_switch * ctx.n_design, color=C_TARGET, ls="--", lw=0.9, label="settled tolerance")
    ax.set_yscale("log")
    _legend(ax, fontsize=8)
    ax.set_title("Phase switches per iteration")
    ax = axes[3]
    rc = ctx.col("relative_compliance_change")
    if rc is not None:
        ax.plot(ctx.x, np.where(np.isfinite(rc) & (rc > 0), rc, np.nan), color=C_MODEL, lw=0.9)
        ax.axhline(ctx.tol_comp, color=C_TARGET, ls="--", lw=0.9, label="settled tolerance")
        ax.set_yscale("log")
        _legend(ax, fontsize=8)
    ax.set_title("Relative compliance change per iteration")
    if show_step:
        ax = axes[4]
        ax.plot(ctx.x, ss, color=C_MODEL, lw=1.0, drawstyle="steps-post", label=r"$\Delta t_k / \Delta t$")
        rej = ctx.col("step_rejected")
        if rej is not None and rej.dtype != object and np.any(rej == 1):
            ax.plot(ctx.x[rej == 1], ss[rej == 1], "v", color=C_BAD, ms=5, label="rejected step (cut or failed analysis)")
        ax.set_yscale("log", base=2)
        ax.set_ylim(float(np.nanmin(ss)) / 1.6, 1.6)
        _legend(ax, fontsize=8)
        ax.set_title("Step size: rejected steps are retaken smaller (§5.8)")
    for i, a in enumerate(axes):
        decorate(a, ctx, xlabel=(i == n_panels - 1))
    fig.tight_layout()
    return fig


def fig_settled_tests(ctx: RunContext):
    from matplotlib import colors as mcolors
    sf = ctx.col("settle_failures")
    if sf is None or sf.dtype != object:
        return None
    tests = list(SETTLE_TESTS) if ctx.buckling else [t for t in SETTLE_TESTS if t != "buckling"]
    if not any("analysis" in str(v).split(";") for v in sf):
        tests = [t for t in tests if t != "analysis"]
    M = np.array([[t in str(v).split(";") for v in sf] for t in tests], dtype=float)
    fig = new_figure(11, 5.5)
    ax1, ax2 = fig.subplots(2, 1, sharex=True, gridspec_kw={"height_ratios": [2.2, 1]})
    cmap = mcolors.ListedColormap([C_OK, "#fdae61"])
    x = ctx.x
    edges = np.r_[x, x[-1] + 1] if len(x) else x
    ax1.pcolormesh(edges, np.arange(len(tests) + 1), M, cmap=cmap, vmin=0, vmax=1, shading="flat")
    ax1.set_yticks(np.arange(len(tests)) + 0.5)
    ax1.set_yticklabels(tests)
    ax1.invert_yaxis()
    ax1.set_title("Settled tests per iteration (blue = passes, orange = fails; docs/theory.md §5.5)")
    sc = ctx.col("settled_count")
    if sc is not None:
        ax2.fill_between(x, 0, sc, color=C_MODEL, alpha=0.4, step="post")
        ax2.axhline(ctx.window, color=C_TARGET, ls="--", lw=0.9, label=f"window = {ctx.window}")
        ax2.set_ylabel("settled count")
        _legend(ax2, fontsize=8)
    decorate(ax2, ctx)
    fig.tight_layout()
    return fig


def fig_timing(ctx: RunContext):
    if not ctx.has("total_iteration_time"):
        return None
    fig = new_figure(11, 7.5)
    ax1, ax2 = fig.subplots(2, 1, sharex=True)
    _plot_line(ax1, ctx, "total_iteration_time", "whole iteration", C_MODEL)
    _plot_line(ax1, ctx, "analysis_time", "analysis (+ buckling + output)", C_EXPORT)
    _plot_line(ax1, ctx, "buckling_time", "buckling eigensolve", C_ALT)
    ax1.set_ylabel("seconds")
    logy_if_positive(ax1, ctx.col("total_iteration_time"))
    _legend(ax1, fontsize=8)
    ax1.set_title("Time per iteration")
    decorate(ax1, ctx, xlabel=False)
    et = ctx.col("elapsed_time")
    if et is not None:
        ax2.plot(ctx.x, et / 3600.0, color=C_MODEL, lw=1.2, label="elapsed [h]")
        ax2.set_ylabel("hours")
    si = ctx.col("solver_iterations")
    if si is not None and np.any(np.isfinite(si)):
        a = ax2.twinx()
        a.plot(ctx.x, si, color=C_EXPORT, lw=0.8, alpha=0.8)
        a.set_ylabel("PCG iterations per solve", color=C_EXPORT)
    ax2.set_title("Elapsed time (and linear-solver iterations, if iterative)")
    decorate(ax2, ctx)
    fig.tight_layout()
    return fig


def fig_buckling(ctx: RunContext):
    B = ctx.buck
    if not len(B.get("iteration", [])):
        return None
    it = B["iteration"]
    figs = []
    fig = new_figure(11, 10)
    axes = fig.subplots(4, 1, sharex=True)
    ax = axes[0]
    ax.plot(it, B.get("critical_load_factor"), ".-", color=C_MODEL, lw=0.9, ms=3, label=r"$\Lambda_{crit}$ (model)")
    ax.plot(it, B.get("aggregated_load_factor"), ".-", color=C_ALT, lw=0.7, ms=2, alpha=0.8, label=r"aggregated $1/\tilde\mu$")
    ax.plot(it, B.get("safety_factor_enforced"), color=C_TARGET, ls="--", lw=0.9, label="safety factor enforced")
    le = _num(ctx.info.get("buckling", {}).get("critical_load_factor_exported"))
    if le is not None and np.isfinite(le):
        ax.plot([it[-1]], [le], "*", color=C_EXPORT, ms=12, label=f"exported design: {le:.4g}")
    _legend(ax, fontsize=8)
    ax.set_title("Critical buckling load factor")
    ax = axes[1]
    ax.plot(it, B.get("g3"), ".-", color=C_MODEL, lw=0.9, ms=3)
    ax.axhline(0, color="k", lw=0.5)
    ax.set_title(r"$G_3$ ($\leq 0$: satisfied)")
    ax = axes[2]
    ax.plot(it, B.get("n_modes_retained"), ".-", color=C_MODEL, lw=0.9, ms=3, label="modes retained")
    if "n_candidates" in B:
        ax.plot(it, B["n_candidates"], ".-", color=C_EXPORT, lw=0.7, ms=2, label="candidates before filtering")
    _legend(ax, fontsize=8)
    ax.set_title("Mode counts per evaluation")
    ax = axes[3]
    msc = B.get("min_subspace_continuity")
    if msc is not None and msc.dtype != object:
        ax.plot(it, msc, ".-", color=C_MODEL, lw=0.9, ms=3)
        ax.set_ylim(-0.02, 1.02)
    ax.set_title("Minimum subspace continuity of the retained modes")
    for i, a in enumerate(axes):
        decorate(a, ctx, xlabel=(i == 3))
    fig.tight_layout()
    figs.append(("11_buckling", fig))
    # Spectrum: every retained load factor of every evaluation.
    mv = B.get("mu_values")
    if mv is not None:
        fig = new_figure(11, 5.5)
        ax = fig.subplots()
        xs, ys, ks = [], [], []
        for x0, cell in zip(it, mv):
            if not isinstance(cell, str) or not cell:
                continue
            vals = [float(v) for v in cell.split(";") if v]
            for k, m in enumerate(vals):
                if m > 0:
                    xs.append(x0)
                    ys.append(1.0 / m)
                    ks.append(k)
        if xs:
            sc = ax.scatter(xs, ys, c=ks, s=6, cmap="viridis")
            fig.colorbar(sc, ax=ax, label="mode rank (0 = critical)")
            ax.plot(it, B.get("safety_factor_enforced"), color=C_TARGET, ls="--", lw=0.9, label="safety factor enforced")
            ax.set_yscale("log")
            _legend(ax, fontsize=8)
        ax.set_ylabel("load factor $1/\\mu_i$")
        ax.set_title("Retained buckling spectrum per evaluation")
        decorate(ax, ctx)
        fig.tight_layout()
        figs.append(("12_buckling_spectrum", fig))
    return figs


# ---------------------------------------------------------------------------
# Final-design figures
# ---------------------------------------------------------------------------

def _views_layout(dims, width=11.0):
    """
    Engineering-drawing layout: top view (x-y) above the front view (x-z),
    side view (y-z) to the right of the front view, the top-right cell free
    for a colour bar. Returns (fig, ax_top, ax_front, ax_side, ax_free).
    """
    nx, ny, nz = dims
    wr = [nx, max(ny, 1)]
    hr = [ny, nz]
    total_w = nx + ny
    total_h = ny + nz
    height = min(15.0, max(4.0, width * total_h / total_w * 1.05))
    fig = new_figure(width, height)
    gs = fig.add_gridspec(2, 2, width_ratios=wr, height_ratios=hr, wspace=0.10, hspace=0.28)
    ax_top = fig.add_subplot(gs[0, 0])
    ax_front = fig.add_subplot(gs[1, 0])
    ax_side = fig.add_subplot(gs[1, 1])
    ax_free = fig.add_subplot(gs[0, 1])
    ax_free.axis("off")
    ax_top.sharex(ax_front)
    ax_top.tick_params(labelbottom=False)
    return fig, ax_top, ax_front, ax_side, ax_free


def _free_colorbar(fig, a_free, im, label=""):
    """A horizontal colour bar in the free top-right cell of _views_layout."""
    cax = a_free.inset_axes([0.08, 0.45, 0.84, 0.14])
    cb = fig.colorbar(im, cax=cax, orientation="horizontal")
    cax.tick_params(labelsize=7)
    if label:
        cb.set_label(label, fontsize=7.5)
    return cb


def _slices(arr, idx=None):
    """Mid-plane slices (z = zc: [x, y]; y = yc: [x, z]; x = xc: [y, z])."""
    nx, ny, nz = arr.shape[:3]
    i, j, k = idx if idx is not None else (nx // 2, ny // 2, nz // 2)
    return arr[:, :, k], arr[:, j, :], arr[i, :, :], (i, j, k)


def _show(ax, img2d, xlabel, ylabel, **kw):
    """imshow of an [u, v] (or [u, v, rgb]) array with u horizontal and v vertical."""
    a = np.asarray(img2d)
    a = a.transpose(1, 0, 2) if a.ndim == 3 else a.T
    im = ax.imshow(a, origin="lower", interpolation="nearest", **kw)
    ax.set_xlabel(xlabel, fontsize=8)
    ax.set_ylabel(ylabel, fontsize=8)
    ax.tick_params(labelsize=7)
    return im


def _cell_centred(point_field):
    """Nodal (nx+1, ny+1, nz+1) field averaged to the voxels."""
    p = np.asarray(point_field, dtype=np.float32)
    return 0.125 * (p[:-1, :-1, :-1] + p[1:, :-1, :-1] + p[:-1, 1:, :-1] + p[:-1, :-1, 1:]
                    + p[1:, 1:, :-1] + p[1:, :-1, 1:] + p[:-1, 1:, 1:] + p[1:, 1:, 1:])


def _field_specs(ctx: RunContext):
    """(key, array, title, norm/cmap kwargs) for every field worth slicing."""
    from matplotlib import colors as mcolors
    f = ctx.fields
    specs = []
    if "phi" not in f:
        return specs
    solid = f.get("solid")

    def sym(a, pct=98):
        v = np.nanpercentile(np.abs(a[f.get("design_mask", np.ones(a.shape, bool))]), pct) if a.size else 1.0
        v = float(v) if np.isfinite(v) and v > 0 else 1.0
        return dict(cmap="PuOr_r", norm=mcolors.TwoSlopeNorm(vcenter=0.0, vmin=-v, vmax=v))

    def logn(a, mask=None, cmap="viridis"):
        vals = a[mask] if mask is not None and np.any(mask) else a
        vals = vals[np.isfinite(vals) & (vals > 0)]
        if vals.size == 0:
            return dict(cmap=cmap)
        vmax = float(np.nanmax(vals))
        vmin = max(vmax * 1e-4, float(np.nanmin(vals)))
        return dict(cmap=cmap, norm=mcolors.LogNorm(vmin=vmin, vmax=vmax))

    specs.append(("phi", f["phi"], r"design field $\phi$ (level set at the white contour)",
                  dict(cmap="RdBu_r", norm=mcolors.TwoSlopeNorm(vcenter=ctx.ls, vmin=min(ctx.V - 0.05, float(np.min(f["phi"]))),
                                                              vmax=max(ctx.S + 0.05, float(np.max(f["phi"])))))))
    if "rho" in f:
        specs.append(("rho", f["rho"], r"design density $\rho(\phi)$", dict(cmap="Greys", vmin=0, vmax=1)))
    if "rho_model" in f:
        specs.append(("rho_model", f["rho_model"], r"model stiffness density $\rho_m = c_1 + (1-c_1)x^p$",
                      dict(cmap="Greys", vmin=0, vmax=1)))
    for key, title in (("exported_Von_Mises_Stress", "von Mises stress of the exported design [Pa] (void blank)"),
                       ("model_Von_Mises_Stress", "von Mises stress of the analysed model [Pa] (void blank)")):
        if key in f and f[key].shape == f["phi"].shape:
            a = np.where(solid, f[key], np.nan) if solid is not None else f[key]
            specs.append((key, a, title, logn(f[key], solid, "magma")))
            break
    if "F0" in f:
        specs.append(("F0", f["F0"], r"base-material strain energy density $F_0$ [J/m$^3$]", logn(f["F0"], solid, "viridis")))
    for key, title in (("exported_point_Displacement_Magnitude", "displacement magnitude, exported design [m]"),
                       ("model_point_Displacement_Magnitude", "displacement magnitude, model [m]")):
        if key in f:
            a = f[key]
            if a.shape != f["phi"].shape:
                a = _cell_centred(a)
            specs.append((key, a, title, dict(cmap="cividis")))
            break
    for key, title in (("stiffness_term", "stiffness term (last iteration)"),
                       ("volume_term", "volume term (last iteration)"),
                       ("piecewise_term", "piecewise term (last iteration)"),
                       ("buckling_term", "applied buckling term (last iteration)"),
                       ("diffusion_term", "diffusion (AOS curvature) term (last iteration)"),
                       ("net_push", r"net push $(\phi^{n+1} - \phi^n)/\Delta t$ (last iteration)"),
                       ("lambda2", r"piecewise multiplier field $\lambda_2$"),
                       ("G2", r"piecewise constraint $G_2 = K(\phi) = (\phi - V)(\phi - S)$")):
        if key in f and f[key].shape == f["phi"].shape:
            a = f[key]
            if key == "buckling_term" and not np.any(a):
                continue
            if key == "stiffness_term":
                specs.append((key, a, title, dict(cmap="viridis", vmin=0, vmax=max(1e-9, float(np.nanpercentile(a, 99.5))))))
            else:
                specs.append((key, a, title, sym(a)))
    if "component_labels" in f and f["component_labels"].max() > 0:
        specs.append(("component_labels", f["component_labels"], "solid components (0 = void)", "labels"))
    return specs


def _label_cmap(labels):
    from matplotlib import colors as mcolors
    import matplotlib
    n = int(labels.max())
    base = matplotlib.colormaps["tab20"]
    cols = [(1, 1, 1, 1)] + [base(i % 20) for i in range(n)]
    return mcolors.ListedColormap(cols), mcolors.BoundaryNorm(np.arange(-0.5, n + 1.5), n + 1)


def fig_slices(ctx: RunContext):
    specs = _field_specs(ctx)
    if not specs:
        return None
    dims = ctx.fields["phi"].shape
    phi = ctx.fields["phi"]
    figs = []
    for n, (key, arr, title, kw) in enumerate(specs):
        fig, a_top, a_front, a_side, a_free = _views_layout(dims)
        s_top, s_front, s_side, (i, j, k) = _slices(arr)
        p_top, p_front, p_side, _ = _slices(phi)
        if kw == "labels":
            cmap, norm = _label_cmap(arr)
            kw = dict(cmap=cmap, norm=norm)
        im = _show(a_top, s_top, "", "y", **kw)
        _show(a_front, s_front, "x", "z", **kw)
        _show(a_side, s_side, "y", "z", **kw)
        a_top.tick_params(labelbottom=False)
        for ax, p in ((a_top, p_top), (a_front, p_front), (a_side, p_side)):
            if np.nanmin(p) < ctx.ls < np.nanmax(p):
                ax.contour(np.asarray(p).T, levels=[ctx.ls], colors="w" if key != "rho" else "r", linewidths=0.6)
        a_top.set_title(f"z = {k}", fontsize=8)
        a_front.set_title(f"y = {j}", fontsize=8)
        a_side.set_title(f"x = {i}", fontsize=8)
        _free_colorbar(fig, a_free, im)
        fig.suptitle(f"{ctx.name}: {title}, mid-plane slices", fontsize=10)
        figs.append((f"20_slice_{n:02d}_{key}", fig))
    return figs


def fig_views(ctx: RunContext):
    from matplotlib import colors as mcolors
    f = ctx.fields
    if "solid" not in f:
        return None
    solid = f["solid"]
    dims = solid.shape
    labels = f.get("component_labels")
    sizes = np.bincount(labels.ravel()) if labels is not None else None
    main = int(np.argmax(sizes[1:]) + 1) if sizes is not None and sizes.size > 1 else 1
    figs = []
    # Depth-shaded views, detached components in orange.
    fig, a_top, a_front, a_side, a_free = _views_layout(dims)
    for ax, axis, fp, (xl, yl), title in ((a_top, 2, True, ("x", "y"), "top (from +z)"),
                                          (a_front, 1, False, ("x", "z"), "front (from -y)"),
                                          (a_side, 0, True, ("y", "z"), "side (from +x)")):
        img, _, first, hit = depth_render(solid, axis, from_positive=fp)
        rgb = np.ones(img.shape + (3,))
        base = np.array(mcolors.to_rgb("#4a78a8"))
        other = np.array(mcolors.to_rgb(C_EXPORT))
        if labels is not None:
            lab = lookup_on_surface(labels, axis, first)
            is_main = (lab == main)
        else:
            is_main = np.ones(img.shape, bool)
        col = np.where(is_main[..., None], base, other)
        shade = np.nan_to_num(img, nan=1.0)[..., None]
        rgb = np.where(hit[..., None], col * (0.35 + 0.9 * shade), 1.0)
        _show(ax, np.clip(rgb, 0, 1), xl, yl)
        ax.set_title(title, fontsize=8)
    fig.suptitle(f"{ctx.name}: depth-shaded views of the exported solid (blue: largest piece; orange: any other piece)",
                 fontsize=10)
    figs.append(("21_views_depth", fig))
    # X-ray projections of the design density.
    rho = f.get("rho", solid.astype(float))
    fig, a_top, a_front, a_side, a_free = _views_layout(dims)
    for ax, axis, (xl, yl), title in ((a_top, 2, ("x", "y"), "mean over z"), (a_front, 1, ("x", "z"), "mean over y"),
                                      (a_side, 0, ("y", "z"), "mean over x")):
        im = _show(ax, rho.mean(axis=axis), xl, yl, cmap="Greys", vmin=0, vmax=1)
        ax.set_title(title, fontsize=8)
    _free_colorbar(fig, a_free, im, "mean density along the view")
    fig.suptitle(f"{ctx.name}: X-ray projections of the design density", fontsize=10)
    figs.append(("22_views_xray", fig))
    # Max-intensity projections of stress and energy inside the solid: load paths.
    for key, title, cmap in (("exported_Von_Mises_Stress", "max von Mises stress in the solid along each axis [Pa]", "magma"),
                             ("model_Von_Mises_Stress", "max von Mises stress in the solid along each axis [Pa]", "magma"),
                             ("F0", r"max base-material SED $F_0$ in the solid along each axis [J/m$^3$]", "viridis")):
        if key not in f or f[key].shape != solid.shape:
            continue
        if key == "model_Von_Mises_Stress" and "exported_Von_Mises_Stress" in f:
            continue
        a = np.where(solid, f[key], 0.0)
        pos = a[a > 0]
        if pos.size == 0:
            continue
        norm = mcolors.LogNorm(vmin=max(pos.max() * 1e-3, pos.min()), vmax=pos.max())
        fig, a_top, a_front, a_side, a_free = _views_layout(dims)
        for ax, axis, (xl, yl) in ((a_top, 2, ("x", "y")), (a_front, 1, ("x", "z")), (a_side, 0, ("y", "z"))):
            proj = a.max(axis=axis)
            im = _show(ax, np.where(proj > 0, proj, np.nan), xl, yl, cmap=cmap, norm=norm)
        _free_colorbar(fig, a_free, im)
        fig.suptitle(f"{ctx.name}: load paths, {title}", fontsize=10)
        figs.append((f"23_loadpaths_{'stress' if 'Mises' in key else 'sed'}", fig))
    return figs


def isosurface_triangles(fields: dict, S: float, V: float, ls: float, max_faces: float = 120000.0):
    """
    The exported design's isosurface phi = level set as triangles, for a 3D
    render: (tri (M, 3, 3), colours (M, 3), unit normals (M, 3), marching-cubes
    step). The largest piece is blue, every other piece orange. None when
    there is no phi field or scikit-image is missing.
    """
    from matplotlib import colors as mcolors
    if "phi" not in fields:
        return None
    try:
        from skimage.measure import marching_cubes
    except Exception:
        return None
    solid = fields.get("solid")
    phi = np.asarray(fields["phi"], dtype=np.float32).copy()
    if "preserve_mask" in fields:
        phi[fields["preserve_mask"]] = S
    if "obstacle_mask" in fields:
        phi[fields["obstacle_mask"]] = V
    if not np.any(phi >= ls):
        return None
    padded = np.pad(phi, 1, constant_values=V)
    faces_est = interface_faces(solid) if solid is not None else 50000
    step = int(max(1, np.ceil(np.sqrt(2.0 * faces_est / max_faces))))
    verts, faces, _normals, _ = marching_cubes(padded, level=ls, step_size=step, allow_degenerate=False)
    verts -= 1.0
    tri = verts[faces]
    # Colour: largest piece blue, other pieces orange, by the voxel under each triangle.
    labels = fields.get("component_labels")
    base = np.array(mcolors.to_rgb("#4a78a8"))
    other = np.array(mcolors.to_rgb(C_EXPORT))
    cols = np.tile(base, (len(faces), 1))
    if labels is not None and labels.max() > 1:
        sizes = np.bincount(labels.ravel())
        main = int(np.argmax(sizes[1:]) + 1)
        cen = tri.mean(axis=1)
        idx = np.clip(np.round(cen).astype(int), 0, np.array(labels.shape) - 1)
        lab = labels[idx[:, 0], idx[:, 1], idx[:, 2]]
        if np.any(lab == 0):
            # Triangle centres just outside the solid: nearest solid voxel's label.
            _, inds = ndimage.distance_transform_edt(labels == 0, return_indices=True)
            li = labels[inds[0], inds[1], inds[2]]
            lab = li[idx[:, 0], idx[:, 1], idx[:, 2]]
        cols[lab != main] = other
    nrm = np.cross(tri[:, 1] - tri[:, 0], tri[:, 2] - tri[:, 0])
    nrm /= np.maximum(np.linalg.norm(nrm, axis=1, keepdims=True), 1e-12)
    return tri, cols, nrm, step


def draw_isosurface(ax, tri, cols, nrm, dims, elev=25.0, azim=-60.0):
    """Draws isosurface_triangles on a 3D axes, lit from near the viewer."""
    from mpl_toolkits.mplot3d.art3d import Poly3DCollection
    dims = np.asarray(dims, dtype=float)
    view = np.array([np.cos(np.radians(elev)) * np.cos(np.radians(azim)),
                     np.cos(np.radians(elev)) * np.sin(np.radians(azim)), np.sin(np.radians(elev))])
    light = view + np.array([0.3, -0.3, 0.6])
    light /= np.linalg.norm(light)
    lam = np.abs(nrm @ light)
    fc = np.clip(cols * (0.3 + 0.75 * lam[:, None]), 0, 1)
    coll = Poly3DCollection(tri, facecolors=fc, edgecolors="none", linewidths=0)
    ax.add_collection3d(coll)
    ax.set_xlim(0, dims[0])
    ax.set_ylim(0, dims[1])
    ax.set_zlim(0, dims[2])
    ax.set_box_aspect(tuple(dims))
    ax.view_init(elev=elev, azim=azim)


def fig_render3d(ctx: RunContext):
    if not ctx.render3d or "phi" not in ctx.fields:
        return None
    import importlib.util
    if importlib.util.find_spec("skimage") is None:
        ctx.note("scikit-image not available: 3D render skipped")
        return None
    iso = isosurface_triangles(ctx.fields, ctx.S, ctx.V, ctx.ls)
    if iso is None:
        return None
    tri, cols, nrm, step = iso
    fig = new_figure(13, 6.5)
    dims = np.array(ctx.fields["phi"].shape, dtype=float)
    for n, (elev, azim) in enumerate(((25, -60), (25, 120))):
        ax = fig.add_subplot(1, 2, n + 1, projection="3d")
        draw_isosurface(ax, tri, cols, nrm, dims, elev, azim)
        ax.set_xlabel("x")
        ax.set_ylabel("y")
        ax.set_zlabel("z")
        ax.tick_params(labelsize=6)
    fig.suptitle(f"{ctx.name}: exported design (isosurface $\\phi = \\phi_{{LS}}$; orange = pieces other than the largest)"
                 + (f", marching-cubes step {step}" if step > 1 else ""), fontsize=10)
    fig.tight_layout()
    return fig


def fig_histograms(ctx: RunContext):
    f = ctx.fields
    if "phi" not in f:
        return None
    dm = f.get("design_mask", np.ones(f["phi"].shape, bool))
    fig = new_figure(11, 7.5)
    ax1, ax2 = fig.subplots(2, 1)
    p = f["phi"][dm]
    band = 0.1 * (ctx.S - ctx.V)
    ax1.axvspan(ctx.V + band, ctx.S - band, color="0.92", label="grey band (5.7)")
    ax1.axvspan(ctx.V + 0.25 * (ctx.S - ctx.V), ctx.S - 0.25 * (ctx.S - ctx.V), color="0.84", label="undecided band")
    ax1.hist(p, bins=240, color=C_MODEL, log=True)
    ax1.axvline(ctx.ls, color="k", lw=0.8, ls=":", label="level set")
    ax1.set_xlabel(r"$\phi$")
    ax1.set_ylabel("design voxels")
    _legend(ax1, fontsize=8)
    ax1.set_title(r"Final $\phi$ over the design voxels")
    if "rho" in f:
        ax2.hist(f["rho"][dm], bins=200, color=C_ALT, log=True, label=r"design density $\rho$")
        if "rho_model" in f:
            ax2.hist(f["rho_model"][dm], bins=200, histtype="step", color=C_EXPORT, log=True,
                     label=r"model stiffness density $\rho_m$")
        ax2.axvline(0.5, color="k", lw=0.8, ls=":")
        ax2.set_xlabel("density")
        _legend(ax2, fontsize=8)
        ax2.set_title("Final densities over the design voxels")
    fig.tight_layout()
    return fig


def fig_mechanics(ctx: RunContext):
    f = ctx.fields
    solid = f.get("solid")
    if solid is None or not np.any(solid):
        return None
    fig = new_figure(12, 9)
    axes = fig.subplots(2, 2).ravel()
    # (a) von Mises distribution in the solid
    ax = axes[0]
    for key, label, colr in (("exported_Von_Mises_Stress", "exported design", C_EXPORT),
                             ("model_Von_Mises_Stress", "analysed model", C_MODEL)):
        if key in f and f[key].shape == solid.shape:
            vals = f[key][solid]
            vals = vals[vals > 0]
            if vals.size:
                bins = np.logspace(np.log10(max(vals.min(), vals.max() * 1e-6)), np.log10(vals.max()), 80)
                ax.hist(vals, bins=bins, histtype="step", color=colr, lw=1.3, label=label)
    ax.set_xscale("log")
    ax.set_xlabel("von Mises stress [Pa]")
    ax.set_ylabel("solid voxels")
    ax.set_title("Stress distribution in the solid")
    if ax.get_legend_handles_labels()[0]:
        _legend(ax, fontsize=8)
    # (b) SED distribution
    ax = axes[1]
    sed = f.get("exported_Exported_SED", f.get("F0"))
    if sed is not None and sed.shape == solid.shape:
        vals = sed[solid]
        vals = vals[vals > 0]
        if vals.size:
            bins = np.logspace(np.log10(max(vals.min(), vals.max() * 1e-8)), np.log10(vals.max()), 80)
            ax.hist(vals, bins=bins, color=C_ALT, alpha=0.8)
            ax.set_xscale("log")
    ax.set_xlabel("strain energy density [J/m$^3$]")
    ax.set_title("Strain energy density in the solid")
    # (c) Lorenz curve
    ax = axes[2]
    if sed is not None and sed.shape == solid.shape:
        x, cum, gini = lorenz_curve(sed[solid])
        ax.plot(x, cum, color=C_MODEL, lw=1.5, label=f"Gini = {gini:.3f}")
        ax.plot([0, 1], [0, 1], color="0.6", ls="--", lw=0.8, label="uniformly loaded")
        for q in (0.5, 0.9):
            sq = share_of_volume_carrying(sed[solid], q)
            ax.plot([sq], [q], "o", color=C_EXPORT)
            ax.annotate(f"{100 * sq:.1f}% of the solid\ncarries {100 * q:.0f}% of the energy", (sq, q),
                        textcoords="offset points", xytext=(8, -18), fontsize=7)
        ax.set_xlabel("share of the solid (most loaded first)")
        ax.set_ylabel("share of the strain energy")
        _legend(ax, fontsize=8, loc="lower right")
    ax.set_title("Energy concentration (Lorenz curve)")
    # (d) optimality: stiffness term vs volume pressure
    ax = axes[3]
    st = f.get("stiffness_term")
    if st is not None and st.shape == solid.shape:
        dm = f.get("design_mask", np.ones(solid.shape, bool))
        er = ndimage.binary_erosion(solid, structure=S6, border_value=0)
        boundary_solid = solid & ~er & dm
        interior = er & dm
        void = (~solid) & dm
        bins = np.linspace(0, max(1e-6, float(np.nanpercentile(st[dm], 99.9))), 80)
        for m_, lab_, c_ in ((interior, "solid interior", C_MODEL), (boundary_solid, "solid boundary", C_EXPORT),
                             (void, "void", "0.5")):
            if np.any(m_):
                ax.hist(st[m_], bins=bins, histtype="step", lw=1.2, color=c_, label=lab_, log=True)
        l1t = ctx.col("lambda1_tilde")
        if l1t is not None and np.any(np.isfinite(l1t)):
            ax.axvline(float(finite(l1t)[-1]) * (1 - ctx.c1), color="k", ls="--", lw=0.9,
                       label=r"volume pressure $\tilde\lambda_1\rho'$")
        ax.set_xlabel("normalised stiffness term (last iteration)")
        _legend(ax, fontsize=8)
    ax.set_title("Optimality: stiffness driver against the volume pressure")
    fig.tight_layout()
    return fig


def fig_thickness(ctx: RunContext):
    f = ctx.fields
    solid = f.get("solid")
    if solid is None or not np.any(solid):
        return None
    samples, half = thickness_samples(solid)
    figs = []
    fig = new_figure(11, 4.8)
    ax = fig.subplots()
    if samples.size:
        mx = int(np.ceil(samples.max()))
        ax.hist(samples, bins=np.arange(0.5, mx + 1.5, 1.0), color=C_MODEL, rwidth=0.85)
        med = float(np.median(samples))
        ax.axvline(med, color=C_EXPORT, lw=1.2, label=f"median {med:.1f} voxels = {med * ctx.voxel_mm:.2f} mm")
        ax.set_xlabel("local member thickness [voxels], sampled on the medial ridge of the distance transform")
        ax.set_ylabel("ridge voxels")
        _legend(ax, fontsize=8)
    ax.set_title(f"Member thickness distribution ({100 * float(np.mean(samples <= 1.0)) if samples.size else 0:.1f}% "
                 f"of the ridge is one voxel thick)")
    fig.tight_layout()
    figs.append(("27_thickness_hist", fig))
    dims = solid.shape
    fig, a_top, a_front, a_side, a_free = _views_layout(dims)
    s_top, s_front, s_side, (i, j, k) = _slices(np.where(solid, half, np.nan))
    vmax = max(1.0, float(np.nanmax(half)))
    for ax, sl, (xl, yl), t in ((a_top, s_top, ("x", "y"), f"z = {k}"), (a_front, s_front, ("x", "z"), f"y = {j}"),
                                (a_side, s_side, ("y", "z"), f"x = {i}")):
        im = _show(ax, sl, xl, yl, cmap="plasma", vmin=0, vmax=vmax)
        ax.set_title(t, fontsize=8)
    _free_colorbar(fig, a_free, im, "voxels")
    fig.suptitle(f"{ctx.name}: distance to the nearest surface (half-thickness), mid-plane slices", fontsize=10)
    figs.append(("27_thickness_map", fig))
    return figs


def fig_threshold_sweep(ctx: RunContext):
    sw = ctx.sweep
    if not len(sw.get("level", [])):
        return None
    lv = sw["level"]
    has_b = "critical_load_factor" in sw and np.any(np.isfinite(sw["critical_load_factor"]))
    fig = new_figure(12, 9 if not has_b else 12)
    rows = 3 if has_b else 2
    axes = fig.subplots(rows, 2).ravel()
    k5 = int(np.argmin(np.abs(lv - 0.5)))
    ax = axes[0]
    comp = sw.get("compliance")
    if comp is not None and np.any(np.isfinite(comp)):
        ax.plot(lv, comp, "o-", color=C_MODEL)
        if np.isfinite(comp[k5]):
            ax.plot([lv[k5]], [comp[k5]], "*", color=C_EXPORT, ms=12, label="exported (s = 0.5)")
        logy_if_positive(ax, comp)
        _legend(ax, fontsize=8)
    else:
        ax.text(0.5, 0.5, "no analyses in the sweep", ha="center", transform=ax.transAxes)
    ax.set_title("Compliance of the cut design [J]")
    ax = axes[1]
    ax.plot(lv, sw["volume_fraction"], "o-", color=C_MODEL, label="cut design")
    ax.axhline(ctx.target, color=C_TARGET, ls="--", lw=0.9, label="target")
    _legend(ax, fontsize=8)
    ax.set_title("Volume fraction of the cut design")
    ax = axes[2]
    for key, lab, c_ in (("n_components", "$b_0$", C_MODEL), ("n_tunnels", "$b_1$", C_EXPORT), ("n_cavities", "$b_2$", C_ALT)):
        if key in sw:
            ax.plot(lv, sw[key], "o-", color=c_, label=lab)
    _legend(ax, fontsize=8)
    ax.set_title("Topology of the cut design")
    ax = axes[3]
    ax.plot(lv, 100 * sw["detached_fraction"], "o-", color=C_MODEL, label="detached solid [%]")
    broken = sw["load_path_connected"] == 0
    if np.any(broken):
        ax.plot(lv[broken], 100 * sw["detached_fraction"][broken], "x", color=C_BAD, ms=9, label="load path broken")
    _legend(ax, fontsize=8)
    ax.set_title("Connectivity of the cut design")
    if has_b:
        ax = axes[4]
        ax.plot(lv, sw["critical_load_factor"], "o-", color=C_MODEL, label=r"$\Lambda_{crit}$")
        sf = ctx.info.get("buckling", {}).get("safety_factor")
        if sf:
            ax.axhline(float(sf), color=C_TARGET, ls="--", lw=0.9, label="target")
        _legend(ax, fontsize=8)
        ax.set_title("Critical load factor of the cut design")
        axes[5].axis("off")
    for a in axes[: 2 * rows]:
        if a.axison:
            a.set_xlabel(r"cut level $s$: solid where $\phi \geq V + s(S - V)$")
            a.axvline(0.5, color="0.8", lw=0.8, zorder=0)
            a.grid(True, alpha=0.25)
    fig.suptitle("Threshold sweep: how much the exported result depends on the threshold", fontsize=11)
    fig.tight_layout(rect=(0, 0, 1, 0.96))
    return fig


def fig_nucleation(ctx: RunContext):
    from matplotlib import cm as mcm
    from matplotlib import colors as mcolors
    E = ctx.events
    if not len(E.get("iteration", [])):
        return None
    it = E["iteration"]
    kind = E["kind"]
    figs = []
    fig = new_figure(11, 6.5)
    ax1, ax2 = fig.subplots(2, 1, sharex=True)
    ni, ns = ctx.col("nucleations_interior"), ctx.col("nucleations_surface")
    if ni is not None:
        ax1.plot(ctx.x, np.cumsum(np.nan_to_num(ni)), color=C_MODEL, label="interior nucleations")
        ax1.plot(ctx.x, np.cumsum(np.nan_to_num(ns)), color=C_EXPORT, label="surface pits")
        ax1.set_yscale("symlog", linthresh=1)
        _legend(ax1, fontsize=8)
        ax1.set_title("Cumulative number of new holes")
        decorate(ax1, ctx, xlabel=False)
    sz = E["n_voxels"]
    for k_, c_, m_ in (("interior", C_MODEL, "o"), ("surface", C_EXPORT, "^")):
        sel = kind == k_
        if np.any(sel):
            ax2.scatter(it[sel], sz[sel], s=8, color=c_, marker=m_, alpha=0.6, label=k_)
    ax2.set_yscale("log")
    ax2.set_ylabel("voxels in the new hole")
    _legend(ax2, fontsize=8)
    ax2.set_title("Size of each nucleated hole")
    decorate(ax2, ctx)
    fig.tight_layout()
    figs.append(("29_nucleation_history", fig))
    # Where: centroids over the final design's silhouette.
    if ctx.dims is not None:
        dims = ctx.dims
        fig, a_top, a_front, a_side, a_free = _views_layout(dims)
        solid = ctx.fields.get("solid")
        cx, cy, cz = E["centroid_x"], E["centroid_y"], E["centroid_z"]
        norm = mcolors.Normalize(vmin=float(np.nanmin(it)), vmax=float(np.nanmax(it)))
        for ax, (u, v), (xl, yl), axis in ((a_top, (cx, cy), ("x", "y"), 2), (a_front, (cx, cz), ("x", "z"), 1),
                                           (a_side, (cy, cz), ("y", "z"), 0)):
            if solid is not None:
                _show(ax, solid.any(axis=axis).astype(float), xl, yl, cmap="Greys", vmin=0, vmax=4, alpha=0.9)
            for k_, m_ in (("interior", "o"), ("surface", "^")):
                sel = kind == k_
                if np.any(sel):
                    ax.scatter(u[sel], v[sel], c=it[sel], cmap="viridis", norm=norm, s=14 + 4 * np.sqrt(sz[sel]),
                                    marker=m_, edgecolors="none", alpha=0.8)
            ax.set_xlim(-0.5, dims["xyz".index(xl)] - 0.5)
            ax.set_ylim(-0.5, dims["xyz".index(yl)] - 0.5)
        _free_colorbar(fig, a_free, mcm.ScalarMappable(norm=norm, cmap="viridis"), "iteration")
        fig.suptitle(f"{ctx.name}: where holes nucleated (o interior, ^ surface pit; grey: final design)", fontsize=10)
        figs.append(("29_nucleation_map", fig))
    return figs


def _evolution_axis(dims):
    """View along the axis whose perpendicular plane has the largest area."""
    areas = [dims[1] * dims[2], dims[0] * dims[2], dims[0] * dims[1]]
    return int(np.argmax(areas))


def fig_evolution(ctx: RunContext):
    if not ctx.history or ctx.dims is None:
        return None
    frames = list(_history_frames(ctx, max_frames=12))
    if not frames:
        return None
    dims = ctx.dims
    axis = _evolution_axis(dims)
    other = [a for a in range(3) if a != axis]
    pres = ctx.fields.get("preserve_mask")
    obst = ctx.fields.get("obstacle_mask")
    ncol = 3 if len(frames) > 4 else len(frames)
    nrow = int(np.ceil(len(frames) / ncol))
    w = 13.0
    cell_w = w / ncol
    cell_h = cell_w * dims[other[1]] / dims[other[0]] + 0.35
    fig = new_figure(w, min(22.0, max(3.0, cell_h * nrow)))
    axes = np.atleast_1d(fig.subplots(nrow, ncol)).ravel()
    for n, (ax, (it, phi)) in enumerate(zip(axes, frames)):
        solid = phi >= ctx.ls
        if pres is not None:
            solid |= pres
        if obst is not None:
            solid &= ~obst
        img, _, _, hit = depth_render(solid, axis, from_positive=False)
        _show(ax, np.where(hit, img, np.nan), "", "", cmap="bone", vmin=0, vmax=1.1)
        ax.set_title(f"iteration {it}", fontsize=8)
        ax.set_xticks([])
        ax.set_yticks([])
    for ax in axes[len(frames):]:
        ax.axis("off")
    fig.suptitle(f"{ctx.name}: evolution of the exported solid (depth-shaded view along {'xyz'[axis]}, "
                 f"horizontal {'xyz'[other[0]]}, vertical {'xyz'[other[1]]})", fontsize=10)
    fig.tight_layout(rect=(0, 0, 1, 0.96))
    return fig


def fig_buckling_modes(ctx: RunContext):
    f = ctx.fields
    keys = sorted(k for k in f if k.startswith("exported_point_Buckling_Mode_") and k.endswith("_Magnitude"))
    if not keys:
        return None
    lf = ctx.info.get("buckling", {}).get("exported_load_factors") or []
    figs = []
    for n, key in enumerate(keys[:4]):
        m = _cell_centred(f[key]) if f[key].shape != f["phi"].shape else f[key]
        solid = f.get("solid")
        a = np.where(solid, m, 0.0) if solid is not None else m
        fig, a_top, a_front, a_side, a_free = _views_layout(a.shape)
        vmax = float(a.max()) if a.max() > 0 else 1.0
        for ax, axis, (xl, yl) in ((a_top, 2, ("x", "y")), (a_front, 1, ("x", "z")), (a_side, 0, ("y", "z"))):
            proj = a.max(axis=axis)
            im = _show(ax, np.where(proj > 0, proj, np.nan), xl, yl, cmap="inferno", vmin=0, vmax=vmax)
        _free_colorbar(fig, a_free, im, "|mode| (normalised)")
        lam = f"{lf[n]:.4g}" if n < len(lf) else "?"
        fig.suptitle(f"{ctx.name}: buckling mode {n + 1} of the exported design, load factor {lam} "
                     f"(max of |mode| in the solid along each view)", fontsize=10)
        figs.append((f"31_buckling_mode_{n + 1}", fig))
    return figs


def fig_components(ctx: RunContext):
    f = ctx.fields
    labels = f.get("component_labels")
    if labels is None or labels.max() == 0:
        return None
    sizes = np.bincount(labels.ravel())[1:]
    if sizes.size <= 1:
        return None
    fig = new_figure(11, 4.5)
    ax = fig.subplots()
    order = np.argsort(sizes)[::-1]
    ax.bar(np.arange(len(order)), sizes[order], color=[C_MODEL] + [C_EXPORT] * (len(order) - 1))
    ax.set_yscale("log")
    ax.set_xlabel("component (largest first)")
    ax.set_ylabel("voxels")
    ax.set_title(f"{len(sizes)} solid components (26-connected); the largest holds "
                 f"{100 * sizes.max() / sizes.sum():.2f}% of the solid")
    fig.tight_layout()
    return fig


# ---------------------------------------------------------------------------
# Animation
# ---------------------------------------------------------------------------

def write_gif(ctx: RunContext, max_frames: int = 150) -> str | None:
    if not ctx.gif or not ctx.history or ctx.dims is None:
        return None
    try:
        from PIL import Image, ImageDraw
    except Exception:
        ctx.note("Pillow not available: evolution.gif skipped")
        return None
    dims = ctx.dims
    axis = _evolution_axis(dims)
    pres = ctx.fields.get("preserve_mask")
    obst = ctx.fields.get("obstacle_mask")
    frames = []
    other = [a for a in range(3) if a != axis]
    scale = max(1, int(np.ceil(360 / max(dims[other[0]], dims[other[1]]))))
    for it, phi in _history_frames(ctx, max_frames=max_frames):
        solid = phi >= ctx.ls
        if pres is not None:
            solid |= pres
        if obst is not None:
            solid &= ~obst
        img, _, _, hit = depth_render(solid, axis, from_positive=False)
        g = np.where(hit, 40 + 200 * np.nan_to_num(img), 255).astype(np.uint8).T[::-1]
        g = np.kron(g, np.ones((scale, scale), dtype=np.uint8))
        im = Image.fromarray(g).convert("P")
        ImageDraw.Draw(im).text((4, 2), f"it {it}", fill=0)
        frames.append(im)
    if not frames:
        return None
    path = os.path.join(ctx.run_dir, "evolution.gif")
    frames[0].save(path, save_all=True, append_images=frames[1:], duration=120, loop=0)
    ctx.written.append(os.path.relpath(path, ctx.run_dir))
    return path


# ---------------------------------------------------------------------------
# Driver
# ---------------------------------------------------------------------------

FIGURES = [
    ("01_convergence_dashboard", fig_dashboard),
    ("02_objective", fig_objective),
    ("03_volume", fig_volume),
    ("04_discreteness", fig_discreteness),
    ("04b_crispness", fig_crispness),
    ("05_topology", fig_topology),
    ("06_multipliers", fig_multipliers),
    ("07_driving_terms", fig_terms),
    ("08_stationarity", fig_stationarity),
    ("09_settled_tests", fig_settled_tests),
    ("10_timing", fig_timing),
    ("11_buckling", fig_buckling),
    ("20_slices", fig_slices),
    ("21_views", fig_views),
    ("24_render3d", fig_render3d),
    ("25_histograms", fig_histograms),
    ("26_mechanics", fig_mechanics),
    ("27_thickness", fig_thickness),
    ("28_threshold_sweep", fig_threshold_sweep),
    ("29_nucleation", fig_nucleation),
    ("30_evolution", fig_evolution),
    ("31_buckling_modes", fig_buckling_modes),
    ("32_components", fig_components),
]


def postprocess_run(run_dir: str, dpi: int = 150, render3d: bool = True, gif: bool = True,
                    verbose: bool = True) -> dict:
    """Writes summary.json/.txt, plots/*.png, report.pdf and evolution.gif for one run directory."""
    from matplotlib.backends.backend_pdf import PdfPages
    t0 = timer()
    ctx = RunContext(run_dir, dpi=dpi, render3d=render3d, gif=gif, verbose=verbose)
    log_path = os.path.join(ctx.plot_dir, "postprocess_log.txt")
    if os.path.exists(log_path):
        os.remove(log_path)
    # Figures of an earlier pass (possibly under other names) would be stale.
    for old_png in glob.glob(os.path.join(ctx.plot_dir, "[0-9][0-9]_*.png")):
        try:
            os.remove(old_png)
        except OSError:
            pass
    if not ctx.var:
        ctx.note("no variables.csv: convergence plots skipped")
    try:
        S = compute_summary(ctx)
    except Exception as e:
        ctx.note(f"summary failed: {e}")
        with open(log_path, "a") as f:
            f.write(f"--- summary ---\n{traceback.format_exc()}\n")
        S = {"run": {"name": ctx.name}}
    with open(os.path.join(ctx.run_dir, "summary.json"), "w") as f:
        json.dump(_clean_json(S), f, indent=2)
    with open(os.path.join(ctx.run_dir, "summary.txt"), "w") as f:
        f.write(summary_text(S))
    pdf_path = os.path.join(ctx.run_dir, "report.pdf")
    with PdfPages(pdf_path) as pdf:
        plotter = Plotter(ctx, pdf)
        plotter.run("00_summary", lambda c: summary_figures(c, S))
        for name, fn in FIGURES:
            plotter.run(name, fn)
        info = pdf.infodict()
        info["Title"] = f"{ctx.method_label} run report: {ctx.name}"
    try:
        write_gif(ctx)
    except Exception as e:
        ctx.note(f"evolution.gif failed: {e}")
    if ctx.log:
        with open(log_path, "a") as f:
            f.write("\n".join(ctx.log) + "\n")
    if verbose:
        print(f"  [postprocess] {ctx.name}: {plotter.n} figures, summary.json/.txt, report.pdf"
              f"{', evolution.gif' if os.path.exists(os.path.join(ctx.run_dir, 'evolution.gif')) else ''}"
              f" in {timer() - t0:.1f} s")
    return S


def _clean_json(x):
    if isinstance(x, dict):
        return {str(k): _clean_json(v) for k, v in x.items()}
    if isinstance(x, (list, tuple)):
        return [_clean_json(v) for v in x]
    if isinstance(x, np.ndarray):
        return _clean_json(x.tolist())
    if isinstance(x, (np.integer,)):
        return int(x)
    if isinstance(x, (float, np.floating)):
        x = float(x)
        if np.isnan(x):
            return None
        if np.isinf(x):
            return "inf" if x > 0 else "-inf"
        return x
    if isinstance(x, np.bool_):
        return bool(x)
    return x


def postprocess_main(argv=None):
    ap = argparse.ArgumentParser(prog="optimise.py --postprocess-only",
                                 description="Plots, statistics and a PDF report for PCLS run directories.")
    ap.add_argument("run_dirs", nargs="+", help="run directories (the optimisation_name folders)")
    ap.add_argument("--dpi", type=int, default=150)
    ap.add_argument("--no-render", action="store_true", help="skip the 3D surface render")
    ap.add_argument("--no-gif", action="store_true", help="skip evolution.gif")
    args = ap.parse_args(argv)
    for d in args.run_dirs:
        if not os.path.isdir(d):
            print(f"Not a directory: {d}")
            continue
        print(f"Post-processing {d} ...")
        postprocess_run(d, dpi=args.dpi, render3d=not args.no_render, gif=not args.no_gif)
