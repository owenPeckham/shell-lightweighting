"""Adapter to Owen's reference FE stack (continuous/reference/analysis.py, gpu_analysis.py).

THEORY §4.1-4.3 and Appendix A. The reference files are imported, never copied or edited.

Internal units are mm, N, MPa; the reference stack works in SI (m, Pa, N, J). Conversions live here only:
compliance J -> N·mm (x 1e3), strain-energy density J/m^3 -> MPa (x 1e-6).

Grid conventions (Appendix A): node counts (nx, ny, nz) = element counts + 1, C order, node id
i·(ny·nz) + j·nz + k, DOFs interleaved [ux, uy, uz] per node. Node (i, j, k) sits at origin + h·(i, j, k).
"""
from __future__ import annotations

import os
import sys
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np

REFERENCE_DIR = Path(__file__).resolve().parents[1] / "continuous" / "reference"
if str(REFERENCE_DIR) not in sys.path:
    sys.path.insert(0, str(REFERENCE_DIR))

import analysis  # noqa: E402  (reference stack)
import gpu_analysis  # noqa: E402

MPa_PER_Pa = 1e-6
Nmm_PER_J = 1e3


@dataclass
class SolveResult:
    """One elastic solve, in mm/N/MPa. THEORY §4.4(a), §6.2."""
    compliance: float                 # C = f·u (N·mm)
    U_e: np.ndarray                   # element strain energy ½uᵀK_e u (N·mm), element-grid shape
    u: np.ndarray                     # nodal displacement (mm), (nx, ny, nz, 3)
    iterations: int
    converged: bool
    final_residual: float
    engine: str
    seconds: float
    warm: bool = False
    info: dict = field(default_factory=dict)

    def dC_drho_hat(self, rho_hat: np.ndarray) -> np.ndarray:
        """∂C/∂ρ̂_e = −2 U_e / ρ̂_e (THEORY §6.2), N·mm per unit ρ̂."""
        return -2.0 * self.U_e / rho_hat


class ShellCG(analysis.ConjugateGradient):
    """`analysis.ConjugateGradient` with node-level boundary conditions (THEORY §4.3).

    The reference class smears element-wise CSG masks onto the 8 element corners, clamping a one-element slab.
    Here `fixed_dof` (bool, (nx, ny, nz, 3) on NODES) and `nodal_load` (N, same shape) are given directly, as
    vshell does. An element-shaped zero `csg_domain_data` is still passed because `__init__` requires one and
    `_apply_elastic_engine` sizes the grid from it.

    Only q1_im with analysis_refinement = 1 is supported for production (THEORY §4.1).
    """

    def __init__(self, fixed_dof: np.ndarray, nodal_load: np.ndarray, h_mm: float, E_MPa: float, nu: float, *,
                 rel_tol: float = 1e-6, max_iters: int = 10_000, engine: str = "auto", pcg_engine: str = "auto",
                 element_type: str = "q1_im", warm_start: bool = True, iter_name: str = "cshell",
                 data_dir: str | None = None, printouts: str = "Silent", **kw):
        fixed_dof = np.asarray(fixed_dof, bool)
        nodal_load = np.asarray(nodal_load, np.float64)
        if fixed_dof.ndim != 4 or fixed_dof.shape[3] != 3 or nodal_load.shape != fixed_dof.shape:
            raise ValueError("fixed_dof and nodal_load must both have shape (nx, ny, nz, 3) on nodes")
        if not fixed_dof.any():
            raise ValueError("no fixed DOF: the problem is singular")
        self.node_shape = tuple(int(n) for n in fixed_dof.shape[:3])
        self.elem_shape = tuple(n - 1 for n in self.node_shape)
        self._fixed_dof = fixed_dof
        self._nodal_load = nodal_load
        self.h_mm = float(h_mm)
        csg = {"fixed_nodes": np.zeros(self.elem_shape + (3,), bool),
               "load_vectors": np.zeros(self.elem_shape + (3,), np.float32)}
        if data_dir is None:
            data_dir = os.environ.get("CSHELL_FE_SCRATCH") or tempfile.gettempdir()
        super().__init__(iter_name, printouts=printouts, data_dir=data_dir, youngs_modulus=E_MPa / MPa_PER_Pa,
                         poisson_ratio=nu, voxel_size_mm=h_mm, pcg_max_iters=max_iters, pcg_tol=0.0,
                         pcg_rel_tol=rel_tol, csg_domain_data=csg, warm_start_cg=warm_start,
                         warm_start_state={} if warm_start else None, element_type=element_type,
                         preconditioner="gmg", analysis_refinement=1, engine=engine, pcg_engine=pcg_engine, **kw)

    # -- reference-stack hook ------------------------------------------------------------------------------------
    def _volume_mesh(self, rho: np.ndarray):
        """Same 11-tuple and dtypes as the parent (THEORY §4.3): fixed_mask int32 per DOF (3·N_nodes), b float32
        per DOF, E_elem/nu_elem float32 per element. Loads on clamped DOF are dropped with a warning."""
        if rho.ndim != 3 or tuple(rho.shape) != self.elem_shape:
            raise ValueError(f"rho must have the element shape {self.elem_shape}, got {rho.shape}")
        if self.analysis_refinement != 1:
            raise RuntimeError(f"ShellCG needs analysis_refinement = 1 (engine set it to {self.analysis_refinement})")
        self._fine_elem_shape = tuple(rho.shape)
        nx, ny, nz = self.node_shape
        h = np.float32(self.voxel_size_m)
        rho_flat = np.ascontiguousarray(rho, dtype=np.float32).ravel(order="C")
        lam_s, mu_s = analysis.lame_parameters(self.youngs_modulus, self.poisson_ratio)
        lam, mu = rho_flat * lam_s, rho_flat * mu_s                   # parent's formula: both Lamé scaled by ρ̂
        den = np.where(np.abs(lam + mu) < 1e-12, 1e-12, lam + mu)
        E_elem = (mu * (3.0 * lam + 2.0 * mu) / den).astype(np.float32)
        nu_elem = (lam / (2.0 * den)).astype(np.float32)
        fixed_mask = self._fixed_dof.reshape(-1).astype(np.int32)
        b = self._nodal_load.reshape(-1).astype(np.float32)
        clash = int(np.count_nonzero((fixed_mask == 1) & (b != 0.0)))
        if clash:
            self._print_out(f"Warning: {clash} loaded DOF are clamped; their load is removed.", "Minimal")
            b[fixed_mask == 1] = 0.0
        return nx, ny, nz, h, h, h, E_elem, nu_elem, b, fixed_mask, rho_flat

    # -- cshell interface ----------------------------------------------------------------------------------------
    def solve(self, rho_hat: np.ndarray) -> SolveResult:
        """Elastic solve for stiffness density ρ̂ (must already contain ρ_min; THEORY §4.1)."""
        t0 = time.time()
        rho_hat = np.asarray(rho_hat, np.float32)
        sed, _ = self.run_analysis(rho_hat)
        eng = getattr(self, "pcg_engine_used", None) or ("cpu_direct" if not hasattr(self, "gmg_info") else "gpu")
        h3 = self.h_mm ** 3
        U = sed.astype(np.float64) * MPa_PER_Pa * h3                 # J/m³ -> MPa = N·mm/mm³, times element volume
        u = (np.asarray(self._last_u_h, np.float64) * 1e3).reshape(self.node_shape + (3,))
        return SolveResult(compliance=self.compliance * Nmm_PER_J, U_e=U, u=u,
                           iterations=int(getattr(self, "solver_iterations", 0)),
                           converged=bool(getattr(self, "solver_converged", True)),
                           final_residual=float(getattr(self, "solver_final_residual", 0.0)),
                           engine=str(eng), seconds=time.time() - t0,
                           warm=bool(getattr(self, "warm_start_attempted", False)))


# ---------------------------------------------------------------------------------------------------------------
# Boundary conditions from B-rep faces (THEORY §4.3)
# ---------------------------------------------------------------------------------------------------------------
def nodes_near_faces(origin, h: float, node_shape, V: np.ndarray, T: np.ndarray, TF: np.ndarray, face_ids,
                     tol_factor: float = 0.75) -> np.ndarray:
    """Flat node ids within tol_factor·h of the triangles of `face_ids` (vshell's rule)."""
    import igl
    T_sel = np.ascontiguousarray(T[np.isin(TF, list(face_ids))], np.int64)
    if T_sel.size == 0:
        return np.zeros(0, np.int64)
    tol = tol_factor * h
    pts = V[np.unique(T_sel)]
    lo = np.floor((pts.min(0) - tol - origin) / h).astype(int).clip(0)
    hi = np.minimum(np.ceil((pts.max(0) + tol - origin) / h).astype(int), np.asarray(node_shape) - 1)
    ii, jj, kk = (np.arange(lo[d], hi[d] + 1) for d in range(3))
    I, J, K = np.meshgrid(ii, jj, kk, indexing="ij")
    ijk = np.stack([I.ravel(), J.ravel(), K.ravel()], 1)
    X = np.ascontiguousarray(origin + h * ijk, np.float64)
    d2 = igl.point_mesh_squared_distance(X, np.ascontiguousarray(V, np.float64), T_sel)[0]
    sel = ijk[d2 <= tol * tol]
    ny, nz = node_shape[1], node_shape[2]
    return sel[:, 0] * ny * nz + sel[:, 1] * nz + sel[:, 2]


def build_bcs(origin, h: float, node_shape, V, T, TF, support_faces, loads):
    """Node-level fixed-DOF mask and nodal load (THEORY §4.3, R3). `loads` is a list of dicts with `faces` and
    `force` (N); each force is shared equally among its load-face nodes. Returns (fixed_dof, nodal_load, info)."""
    n = int(np.prod(node_shape))
    fixed = nodes_near_faces(origin, h, node_shape, V, T, TF, support_faces)
    if fixed.size == 0:
        raise RuntimeError("no FE nodes found on the support faces")
    fixed_dof = np.zeros((n, 3), bool)
    fixed_dof[fixed] = True
    load = np.zeros((n, 3))
    n_load = []
    for L in loads:
        ln = nodes_near_faces(origin, h, node_shape, V, T, TF, L["faces"])
        if ln.size == 0:
            raise RuntimeError(f"no FE nodes found on load faces {L['faces']}")
        load[ln] += np.asarray(L["force"], float) / ln.size
        n_load.append(int(ln.size))
    clash = int(np.count_nonzero(fixed_dof & (load != 0)))
    info = dict(n_fixed_nodes=int(fixed.size), n_load_nodes=n_load, loaded_clamped_dof=clash)
    return fixed_dof.reshape(tuple(node_shape) + (3,)), load.reshape(tuple(node_shape) + (3,)), info


# ---------------------------------------------------------------------------------------------------------------
# Reference-code fix: q1_im buckling sensitivity (THEORY §6.4, OPEN_QUESTIONS B1)
# ---------------------------------------------------------------------------------------------------------------
class _ElementDataConsistent(analysis.ElementData):
    """`analysis.ElementData` with the solid stiffness K0 = E_s·Khat (E_s = 2μ_s(1 + ν)) for q1 / q1_im, i.e. the
    ∂K/∂ρ̂ of the matrix actually assembled. The reference uses the Gauss K0 = λ_s Klam + μ_s Kmu in Term A and in
    the explicit adjoint term, which biases ∂μ̃/∂ρ̂ by ≈ 5.8 % for q1_im."""

    def __init__(self, xp, grid, E_elem, nu_elem, u, rho_flat, threshold, lam_s, mu_s, element_type, nu_run, Khat):
        super().__init__(xp, grid, E_elem, nu_elem, u, rho_flat, threshold, lam_s, mu_s, element_type, nu_run, Khat)
        if element_type in ("q1", "q1_im"):
            self.K0 = analysis._to(xp, 2.0 * float(mu_s) * (1.0 + float(nu_run)) * np.asarray(Khat, np.float64))


def _run_buckling_fixed(self, *args, fix_q1im_sensitivity: bool = True, **kw):
    """`ConjugateGradient.run_buckling` with the q1_im sensitivity fix applied for the duration of the call."""
    if not fix_q1im_sensitivity:
        return analysis.ConjugateGradient.run_buckling(self, *args, **kw)
    orig = analysis.ElementData
    analysis.ElementData = _ElementDataConsistent
    try:
        return analysis.ConjugateGradient.run_buckling(self, *args, **kw)
    finally:
        analysis.ElementData = orig


ShellCG.run_buckling = _run_buckling_fixed
