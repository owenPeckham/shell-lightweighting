"""Active-set voxel FE with a sparse direct (Cholesky) solver: cshell's production FE engine.

Decision of 28 October 2026 (continuous/STATUS.md §2.2, 0.3; research in continuous/research/fe_engine/REPORT.md):
the reference stack solves the dense bounding box, of which the material fills 1.7-4 %. Here only the elements of an
**active set** S (every element with ρ_e > 0, usually dilated by a margin so the sparsity pattern survives several
iterations) are assembled. Elements outside S carry ρ̂ = ρ_min in the dense model, so dropping them changes K by
O(ρ_min) and the gradient not at all: ∂χ/∂ψ = 0 away from the surface, so every element with a non-zero design
derivative has ρ_e > 0 and lies in S (THEORY §3.7, §6.1).

Element: the same q1_im (Wilson-Taylor, condensed) unit-E matrix as the reference, `gpu_analysis.element_khat`,
so K_e = E ρ̂_e K̂ exactly as `ConjugateGradient._cg_assemble_stiffness` (THEORY §4.1). One factorisation per design
serves the primal, the stress adjoint and the buckling eigen-solve (THEORY §6.3-6.4).

Backends: "gpu" = CuPy CSR + NVIDIA cuDSS (nvmath-python DirectSolver, SPD); "cpu" = SciPy CSR + the reference
host Cholesky (`analysis.HostCholesky`) for the CPU test suite (R8).

Grid conventions are those of the reference (fea.py docstring): dense element box (nx_e, ny_e, nz_e), node box one
larger, C order, the reference's local corner order, DOFs interleaved per node. Units mm, N, MPa.
"""
from __future__ import annotations

import time
from dataclasses import dataclass

import numpy as np

from . import fea  # noqa: F401  (puts continuous/reference on sys.path)
import gpu_analysis

# local corner order of the reference kernels (gpu_analysis.xi_arr/eta_arr/zeta_arr, hex_dof_map n0..n7)
CORNERS = np.array([[0, 0, 0], [1, 0, 0], [1, 1, 0], [0, 1, 0],
                    [0, 0, 1], [1, 0, 1], [1, 1, 1], [0, 1, 1]], np.int64)
import os as _os
_EIG_VERBOSE = bool(_os.environ.get("CSHELL_EIG_VERBOSE"))
_SELF = 13                                    # stencil slot of offset (0, 0, 0) in 13 + 9 dx + 3 dy + dz


def _stencil_slot(d):
    return 13 + 9 * d[..., 0] + 3 * d[..., 1] + d[..., 2]


def centre_B(h: float) -> np.ndarray:
    """6 x 24 compatible strain matrix at the element centre, Voigt (xx, yy, zz, yz, xz, xy), engineering shear.
    For q1_im the incompatible-mode gradients vanish at the centre, so B(0)u_e is the element's exact centre strain
    (THEORY §4.4(b))."""
    s = 2.0 * CORNERS - 1.0
    dN = 0.125 * s * (2.0 / h)                # (8, 3): dN_a/dx_i at ξ = η = ζ = 0
    B = np.zeros((6, 24))
    for a in range(8):
        dx, dy, dz = dN[a]
        c = 3 * a
        B[0, c] = dx
        B[1, c + 1] = dy
        B[2, c + 2] = dz
        B[3, c + 1], B[3, c + 2] = dz, dy
        B[4, c], B[4, c + 2] = dz, dx
        B[5, c], B[5, c + 1] = dy, dx
    return B


def unit_D(nu: float) -> np.ndarray:
    """Isotropic constitutive matrix for E = 1, Voigt (xx, yy, zz, yz, xz, xy), engineering shear."""
    lam = nu / ((1 + nu) * (1 - 2 * nu))
    mu = 1.0 / (2 * (1 + nu))
    D = np.zeros((6, 6))
    D[:3, :3] = lam
    D[np.arange(3), np.arange(3)] += 2 * mu
    D[np.arange(3, 6), np.arange(3, 6)] = mu
    return D


class ActiveGrid:
    """Active elements of a dense element box, their nodes and a fixed CSR sparsity pattern.

    The CSR row of DOF 3n+p lists, for each neighbour node of n (in stencil order), its three DOFs. Block (n, k)
    entry (p, q) therefore sits at 9·bptr[n] + 3p·nnb[n] + 3k + q, with bptr the block-row pointer and nnb[n] the
    neighbour count. Assembly scatters 3×3 element blocks into block slots and permutes once into CSR order.
    """

    def __init__(self, elem_mask: np.ndarray):
        elem_mask = np.asarray(elem_mask, bool)
        self.elem_shape = tuple(int(n) for n in elem_mask.shape)
        self.node_shape = tuple(n + 1 for n in self.elem_shape)
        nxn, nyn, nzn = self.node_shape
        self.elem_dense = np.flatnonzero(elem_mask.ravel())              # C-order flat ids in the element box
        if self.elem_dense.size == 0:
            raise ValueError("empty active set")
        eijk = np.stack(np.unravel_index(self.elem_dense, self.elem_shape), 1)
        nijk = eijk[:, None, :] + CORNERS[None]                            # (Ne, 8, 3)
        ndense = (nijk[..., 0] * nyn + nijk[..., 1]) * nzn + nijk[..., 2]
        self.node_dense, conn = np.unique(ndense.ravel(), return_inverse=True)
        self.conn = conn.reshape(-1, 8).astype(np.int64)
        self.Ne, self.Nn = self.conn.shape[0], self.node_dense.size
        self.ndof = 3 * self.Nn
        # neighbour presence per node (27 stencil slots)
        slot_ab = _stencil_slot(CORNERS[None, :, :] - CORNERS[:, None, :])   # (8, 8): slot of node b seen from a
        present = np.zeros((self.Nn, 27), bool)
        for a in range(8):
            for b in range(8):
                present[self.conn[:, a], slot_ab[a, b]] = True
        self.nnb = present.sum(1).astype(np.int64)
        self.bptr = np.concatenate([[0], np.cumsum(self.nnb)])
        rank = np.cumsum(present, 1) - 1                                     # rank of each present slot
        self.nblk = int(self.bptr[-1])
        self.nnz = 9 * self.nblk
        # block slot of every element node pair
        self.pos_blk = np.empty((self.Ne, 8, 8), np.int64)
        for a in range(8):
            na = self.conn[:, a]
            for b in range(8):
                self.pos_blk[:, a, b] = self.bptr[na] + rank[na, slot_ab[a, b]]
        # block row and neighbour node of each block
        blk_row = np.repeat(np.arange(self.Nn), self.nnb)
        nidx = np.stack(np.unravel_index(self.node_dense, self.node_shape), 1)
        offs = np.stack(np.unravel_index(np.arange(27), (3, 3, 3)), 1) - 1     # slot -> (dx, dy, dz)
        node_of_dense = -np.ones(int(np.prod(self.node_shape)), np.int64)
        node_of_dense[self.node_dense] = np.arange(self.Nn)
        rr, ss = np.nonzero(present)                                          # row-major: blocks in bptr order
        nb_ijk = nidx[rr] + offs[ss]
        nb_dense = (nb_ijk[:, 0] * nyn + nb_ijk[:, 1]) * nzn + nb_ijk[:, 2]
        blk_col = node_of_dense[nb_dense]
        assert np.all(blk_col >= 0) and np.array_equal(rr, blk_row)
        # CSR arrays
        k = np.arange(self.nblk) - self.bptr[blk_row]                       # rank of the block within its row
        p = np.arange(3)
        perm = (9 * self.bptr[blk_row][:, None, None] + 3 * p[None, :, None] * self.nnb[blk_row][:, None, None]
                + 3 * k[:, None, None] + p[None, None, :])                  # (nblk, 3, 3) -> CSR data index
        self.perm = perm.ravel()
        self.indices = np.empty(self.nnz, np.int64)
        self.indices[self.perm] = (3 * blk_col[:, None, None] + p[None, None, :] + 0 * p[None, :, None]).ravel()
        self.indptr = np.zeros(self.ndof + 1, np.int64)
        self.indptr[1:] = np.cumsum(np.repeat(3 * self.nnb, 3))
        assert self.indptr[-1] == self.nnz
        self.row_of = np.repeat(np.arange(self.ndof), np.diff(self.indptr))
        self.dofs = (3 * self.conn[:, :, None] + p[None, None, :]).reshape(self.Ne, 24)

    def gather_nodes(self, node_field: np.ndarray) -> np.ndarray:
        """(nx, ny, nz, ...) dense node array -> (Nn, ...) on active nodes."""
        a = np.asarray(node_field)
        return a.reshape((-1,) + a.shape[3:])[self.node_dense]

    def scatter_nodes(self, vals: np.ndarray, fill=0.0) -> np.ndarray:
        """(Nn, ...) -> dense (nx, ny, nz, ...) node array."""
        out = np.full((int(np.prod(self.node_shape)),) + vals.shape[1:], fill, dtype=vals.dtype)
        out[self.node_dense] = vals
        return out.reshape(self.node_shape + vals.shape[1:])

    def gather_elems(self, elem_field: np.ndarray) -> np.ndarray:
        return np.asarray(elem_field).reshape(-1)[self.elem_dense]

    def scatter_elems(self, vals, fill=0.0) -> np.ndarray:
        out = np.full(int(np.prod(self.elem_shape)), fill, dtype=np.asarray(vals).dtype)
        out[self.elem_dense] = vals
        return out.reshape(self.elem_shape)


@dataclass
class Primal:
    compliance: float          # f·u (N·mm)
    u: object                  # (ndof,) on active DOFs, mm (xp array)
    U_e: np.ndarray            # ½ u_eᵀ K_e u_e per active element (N·mm), host float64
    seconds: dict


class SparseFE:
    """K(ρ̂) on an ActiveGrid, factorised once per design; solves for any number of right-hand sides."""

    def __init__(self, grid: ActiveGrid, fixed_dof: np.ndarray, nodal_load: np.ndarray, h_mm: float, E_MPa: float,
                 nu: float, backend: str = "auto", element_type: str = "q1_im", hybrid: bool = False):
        self.g, self.h, self.E, self.nu = grid, float(h_mm), float(E_MPa), float(nu)
        self.element_type = element_type
        if backend == "auto":
            backend = "gpu" if _cuda_ok() else "cpu"
        self.backend = backend
        if backend == "gpu":
            import cupy as cp
            import cupyx.scipy.sparse as cps
            self.xp, self._sp = cp, cps
        else:
            import scipy.sparse as sps
            self.xp, self._sp = np, sps
        xp = self.xp
        self.Khat = gpu_analysis.element_khat(element_type, (self.h,) * 3, self.nu)          # unit E, N/mm per MPa
        fx = grid.gather_nodes(np.asarray(fixed_dof, bool)).reshape(-1)
        f = grid.gather_nodes(np.asarray(nodal_load, float)).reshape(-1)
        tot_dense = np.asarray(nodal_load, float).reshape(-1, 3).sum(0)
        if not np.allclose(f.reshape(-1, 3).sum(0), tot_dense):
            raise ValueError("some loaded nodes are outside the active set")
        if not fx.any():
            raise ValueError("no fixed DOF in the active set")
        clash = fx & (f != 0)
        if clash.any():
            print(f"fe_sparse: warning: {int(clash.sum())} loaded DOF are clamped; load removed", flush=True)
            f[clash] = 0.0
        self.fixed = fx
        self.f = xp.asarray(f)
        # entries kept (free-free) and diagonal entries of fixed DOF (set to 1)
        col = grid.indices
        keep = ~(fx[grid.row_of] | fx[col])
        diag_fixed = np.flatnonzero((grid.row_of == col) & fx[grid.row_of])
        self._keep = xp.asarray(keep.astype(np.float64))
        self._diag_fixed = xp.asarray(diag_fixed)
        self._pos_blk = xp.asarray(grid.pos_blk.reshape(grid.Ne, 64))
        self._perm = xp.asarray(grid.perm)
        self._dofs = xp.asarray(grid.dofs)
        self._Khat = xp.asarray(self.Khat)
        B_all, Bg_all, detJ = gpu_analysis.hex_reference_data(self.h, self.h, self.h)   # 2x2x2 Gauss, weights 1
        self._Bg_np, self._detJ = Bg_all, float(detJ)
        self._DB = xp.asarray(np.einsum("ij,gjk->gik", unit_D(self.nu), B_all))        # (8, 6, 24) unit-E D B_g
        self._Bg = xp.asarray(Bg_all)                                                    # (8, 3, 8) dN_a/dx_i
        self._K = None
        self._solver = None
        self.hybrid = bool(hybrid)
        self.stats = dict(Ne=grid.Ne, ndof=grid.ndof, nnz=grid.nnz, backend=backend)

    # -- assembly and factorisation --------------------------------------------------------------------------------
    def _values(self, rho_hat):
        """CSR data of K = Σ_e E ρ̂_e K̂ with fixed DOF replaced by identity rows/columns."""
        xp, g = self.xp, self.g
        Ee = self.E * xp.asarray(rho_hat, dtype=xp.float64)
        blocks = xp.zeros((g.nblk, 3, 3))
        Kb = self._Khat.reshape(8, 3, 8, 3).transpose(0, 2, 1, 3).reshape(64, 3, 3)
        for ab in range(64):
            contrib = Ee[:, None, None] * Kb[ab][None]
            if xp is np:
                np.add.at(blocks, self._pos_blk[:, ab], contrib)
            else:
                import cupyx
                cupyx.scatter_add(blocks, self._pos_blk[:, ab], contrib)
        data = xp.empty(g.nnz)
        data[self._perm] = blocks.reshape(-1)
        data *= self._keep
        data[self._diag_fixed] = 1.0
        return data

    def factor(self, rho_hat_active):
        """Assemble and factorise K(ρ̂). ρ̂ is given per active element (length Ne) and must contain ρ_min."""
        t0 = time.time()
        rho = np.asarray(rho_hat_active) if self.xp is np else rho_hat_active
        if rho.shape[0] != self.g.Ne:
            raise ValueError("rho_hat must have one value per active element")
        self.rho_hat = self.xp.asarray(rho, dtype=self.xp.float64)
        data = self._values(self.rho_hat)
        t1 = time.time()
        g = self.g
        if self.backend == "gpu":
            from nvmath.sparse.advanced import DirectSolver, DirectSolverMatrixType, DirectSolverOptions
            cp = self.xp
            if self._K is None:
                self._K = self._sp.csr_matrix((data, cp.asarray(g.indices.astype(np.int32)),
                                               cp.asarray(g.indptr.astype(np.int32))), shape=(g.ndof, g.ndof))
                self._b = cp.asfortranarray(cp.zeros((g.ndof, 1)))
                opts = DirectSolverOptions(sparse_system_type=DirectSolverMatrixType.SPD,
                                           multithreading_lib=_cudss_mtlayer())          # threaded host planning
                self._solver = DirectSolver(self._K, self._b, options=opts, execution=self._execution())
                self._solver.plan()
                self.stats["plan_s"] = time.time() - t1
            else:
                self._K.data[...] = data                     # same buffers -> numeric refactorisation only
            try:
                self._solver.factorize()
            except Exception as e:                       # no silent fallback: log, switch to hybrid memory, re-plan
                if "ALLOC_FAILED" not in repr(e) or self.hybrid:
                    raise
                print(f"fe_sparse: WARNING: cuDSS factor does not fit in GPU memory ({g.ndof} DOF); retrying in "
                      f"hybrid memory mode (factor partly in host memory)", flush=True)
                self.free()
                self._K = None
                self.hybrid = True
                self.stats["hybrid_fallback"] = True
                return self.factor(rho_hat_active)
            cp.cuda.Device().synchronize()
        else:
            import analysis
            self._K = self._sp.csr_matrix((data, g.indices, g.indptr), shape=(g.ndof, g.ndof))
            if self._solver is not None:
                self._solver.free()
            self._solver = analysis.HostCholesky().factor(self._K)
        self.stats.update(assemble_s=t1 - t0, factor_s=time.time() - t1)
        return self

    def _execution(self):
        """cuDSS execution space: device memory only, or hybrid (factor spills to host memory) when it does not fit."""
        from nvmath.sparse.advanced import ExecutionCUDA, HybridMemoryModeOptions
        if not self.hybrid:
            return ExecutionCUDA()
        return ExecutionCUDA(hybrid_memory_mode_options=HybridMemoryModeOptions(hybrid_memory_mode=True))

    def solve(self, rhs):
        """K x = rhs for (ndof,) or (ndof, k) right-hand sides; fixed DOF of rhs are zeroed (x = 0 there)."""
        xp = self.xp
        b = xp.array(rhs, dtype=xp.float64)
        one = b.ndim == 1
        if one:
            b = b[:, None]
        b[xp.asarray(self.fixed)] = 0.0
        if self.backend == "gpu":
            # nvmath operands must keep the planned buffer and strides: copy each column into it in place
            x = xp.empty_like(b)
            for j in range(b.shape[1]):
                self._b[:, 0] = b[:, j]
                x[:, j] = xp.asarray(self._solver.solve()).reshape(-1)
        else:
            x = np.column_stack([self._solver.solve(b[:, j]) for j in range(b.shape[1])])
        return x[:, 0] if one else x

    def matvec(self, x):
        return self._K @ x

    # -- responses -------------------------------------------------------------------------------------------------
    def primal(self) -> Primal:
        t0 = time.time()
        u = self.solve(self.f)
        C = float(self.f @ u)
        U = self.element_bilinear(u, u, solid=False) * 0.5
        return Primal(compliance=C, u=u, U_e=U, seconds=dict(solve=time.time() - t0, **{
            k: self.stats[k] for k in ("assemble_s", "factor_s") if k in self.stats}))

    def element_bilinear(self, a, b, solid: bool = True) -> np.ndarray:
        """a_eᵀ K_e b_e per active element (host float64). solid=True uses K⁰_e = E K̂ (the ∂K/∂ρ̂_e of THEORY
        §6.3); solid=False uses K_e = E ρ̂_e K̂."""
        xp = self.xp
        ae, be = a[self._dofs], b[self._dofs]
        val = self.E * xp.einsum("ei,ij,ej->e", ae, self._Khat, be)
        if not solid:
            val = val * self.rho_hat
        return _host(val)

    def centre_stress(self, u, solid: bool = True):
        """(Ne, 6) centre stress, Voigt (xx, yy, zz, yz, xz, xy); material stress (D⁰) if solid else ρ̂-scaled."""
        xp = self.xp
        DB = xp.asarray(self.E * unit_D(self.nu) @ centre_B(self.h))
        s = u[self._dofs] @ DB.T
        if not solid:
            s = s * self.rho_hat[:, None]
        return s

    def dC_drho_hat(self, p: Primal) -> np.ndarray:
        """∂C/∂ρ̂_e = −u_eᵀ K⁰_e u_e = −2 U_e/ρ̂_e (THEORY §6.2)."""
        return -2.0 * p.U_e / _host(self.rho_hat)

    def free(self):
        if self._solver is not None and hasattr(self._solver, "free"):
            self._solver.free()
        self._solver = None


def _host(a) -> np.ndarray:
    try:
        import cupy as cp
        if isinstance(a, cp.ndarray):
            return cp.asnumpy(a).astype(np.float64)
    except ImportError:
        pass
    return np.asarray(a, np.float64)


def _cuda_ok() -> bool:
    try:
        import cupy as cp
        return cp.cuda.runtime.getDeviceCount() > 0
    except Exception:
        return False


# ---------------------------------------------------------------------------------------------------------------------
# Linear buckling on the same factorisation (THEORY §4.4(c), §6.4)
# ---------------------------------------------------------------------------------------------------------------------
@dataclass
class BucklingResult:
    mu: np.ndarray             # retained inverse load factors μ_j = 1/λ_j, most critical first
    mu_tilde: float            # (Σ μ_j^p)^{1/p} over positive retained μ_j; 0 if none
    modes: object              # (ndof, m) K-normalised (xp array)
    dmu_tilde: np.ndarray      # ∂μ̃/∂ρ̂_e per active element (host)
    energy_fraction: np.ndarray
    n_candidates: int
    converged: bool
    iterations: int
    seconds: float


def _ramp(rho, t):
    """Smoothstep ramp of the geometric stiffness, 0 below t, 1 above 2t (reference analysis.py, ElementData)."""
    x = ((rho - t) / t).clip(0.0, 1.0)
    return x * x * (3.0 - 2.0 * x), 6.0 * x * (1.0 - x) / t


def buckling(fe: SparseFE, primal: Primal, num_modes: int = 6, p: float = 8.0, min_density: float = 0.1,
             min_energy_fraction: float = 0.5, tol: float = 1e-6, maxiter: int = 100, X0=None,
             extra: int = 4, seed: int = 0) -> BucklingResult:
    """Inverse load factors of (−K_σ) φ = μ K φ and the sensitivity of μ̃ (THEORY §4.4(c), §6.4).

    Same discrete model as the reference (`analysis.ElementData`): K = Σ E ρ̂ K̂ (q1_im); stress σ_g = ρ̂ E D B_g u at
    the 8 Gauss points from the compatible B, ramped by the smoothstep r(ρ̂) at `min_density`; K_σ,e = Σ_g detJ
    ∇Nᵀ S_g ∇N ⊗ I₃. With φᵀKφ = 1:

        ∂μ/∂ρ̂_e = −μ φ_eᵀK⁰_eφ_e − Σ_g detJ (r + ρ̂ r′) E (D B_g u_e)·Γ_g(φ) + η_eᵀ K⁰_e u_e,   K η = c(φ),

    with Γ_g(φ) the Voigt geometric strain and c = Σ_e Σ_g detJ r ρ̂ E B_gᵀ D Γ_g(φ). K⁰_e = E K̂ is the element's
    own (q1_im) solid stiffness, so the bias of the reference sensitivity (THEORY §6.4, D6) does not arise.
    Eigen-solve: restarted block Krylov on K⁻¹G with K-orthonormal Rayleigh-Ritz (`_block_krylov_largest`).
    """
    xp, g = fe.xp, fe.g
    t0 = time.time()
    rho = fe.rho_hat
    r, rp = _ramp(rho, min_density)
    u_e = primal.u[fe._dofs]                                                    # (Ne, 24)
    sig = fe.E * xp.einsum("gij,ej->egi", fe._DB, u_e) * rho[:, None, None]    # actual stress (Ne, 8, 6)
    sig_r = sig * r[:, None, None]
    S = _voigt_to_tensor(xp, sig_r)                                             # (Ne, 8, 3, 3)
    H = fe._detJ * xp.einsum("gia,egij,gjb->eab", fe._Bg, S, fe._Bg)           # (Ne, 8, 8)
    blocks = xp.zeros((g.nblk,))
    for ab in range(64):
        a, b = divmod(ab, 8)
        _scatter_add(xp, blocks, fe._pos_blk[:, ab], H[:, a, b])
    eye = xp.eye(3).reshape(-1)
    data = xp.empty(g.nnz)
    data[fe._perm] = (blocks[:, None] * eye[None, :]).reshape(-1)
    data *= fe._keep
    Gm = fe._sp.csr_matrix((-data, fe._K.indices, fe._K.indptr), shape=fe._K.shape)    # G = −K_σ
    # LOBPCG, largest μ, B = K, preconditioner K⁻¹
    m = num_modes + extra
    rng = np.random.default_rng(seed)
    X = xp.asarray(rng.standard_normal((g.ndof, m))) if X0 is None else xp.asarray(X0)
    if X.shape[1] < m:
        X = xp.concatenate([X, xp.asarray(rng.standard_normal((g.ndof, m - X.shape[1])))], 1)
    X[xp.asarray(fe.fixed)] = 0.0
    vals, vecs, res, n_it = _block_krylov_largest(fe, Gm, X, num_modes, tol, maxiter)
    converged = bool(np.all(res[:num_modes] <= tol))
    hist = range(n_it)
    # K-normalise
    nrm = xp.sqrt(xp.einsum("ij,ij->j", vecs, fe._K @ vecs))
    vecs = vecs / nrm[None, :]
    n_candidates = int(np.count_nonzero(vals > 0))
    # spurious-mode filter: share of modal strain energy in elements with full geometric stiffness
    frac = np.zeros(vals.size)
    for j in range(vals.size):
        ej = fe.element_bilinear(vecs[:, j], vecs[:, j], solid=False)
        frac[j] = float((_host(r) * ej).sum() / max(ej.sum(), 1e-300))
    keep = (vals > 0) & (frac >= min_energy_fraction)
    idx = np.flatnonzero(keep)
    n_ret = min(num_modes, idx.size)
    while n_ret < idx.size and abs(vals[idx[n_ret]] - vals[idx[n_ret - 1]]) <= 1e-3 * abs(vals[idx[n_ret - 1]]):
        n_ret += 1                     # never cut a degenerate cluster: the p-norm gradient is basis-dependent otherwise
    idx = idx[:n_ret]
    mu = vals[idx]
    if mu.size == 0:
        return BucklingResult(mu=mu, mu_tilde=0.0, modes=vecs[:, :0], dmu_tilde=np.zeros(g.Ne), energy_fraction=frac,
                              n_candidates=n_candidates, converged=converged, iterations=len(hist),
                              seconds=time.time() - t0)
    mu_max = mu.max()
    Ssum = float(np.sum((mu / mu_max) ** p))
    mu_tilde = mu_max * Ssum ** (1.0 / p)
    w = (mu / mu_tilde) ** (p - 1.0)                                            # ∂μ̃/∂μ_j
    # sensitivities
    dmu = xp.zeros(g.Ne)
    cvec = xp.zeros(g.ndof)
    drs = (r + rho * rp) * fe.E
    for k, j in enumerate(idx):
        phi = vecs[:, j]
        phi_e = phi[fe._dofs].reshape(-1, 8, 3)
        grad = xp.einsum("gia,eak->egki", fe._Bg, phi_e)                       # ∂φ_k/∂x_i at Gauss points
        Gam = _gamma_voigt(xp, grad)                                             # (Ne, 8, 6)
        sig_unit = xp.einsum("gij,ej->egi", fe._DB, u_e)                         # D B u (unit E)
        term_a = -mu[k] * fe.E * xp.einsum("ei,ij,ej->e", phi[fe._dofs], fe._Khat, phi[fe._dofs])
        term_b = -fe._detJ * drs * xp.einsum("egi,egi->e", sig_unit, Gam)
        dmu += w[k] * (term_a + term_b)
        ce = fe._detJ * fe.E * (r * rho)[:, None] * xp.einsum("gij,egi->ej", fe._DB, Gam)   # B_gᵀ D Γ_g, (Ne, 24)
        _scatter_add(xp, cvec, fe._dofs.reshape(-1), (w[k] * ce).reshape(-1))
    eta = fe.solve(cvec)
    dmu += fe.E * xp.einsum("ei,ij,ej->e", eta[fe._dofs], fe._Khat, u_e)
    return BucklingResult(mu=mu, mu_tilde=float(mu_tilde), modes=vecs[:, xp.asarray(idx)], dmu_tilde=_host(dmu),
                          energy_fraction=frac, n_candidates=n_candidates, converged=converged,
                          iterations=len(hist), seconds=time.time() - t0)


def _scatter_add(xp, out, idx, vals):
    if xp is np:
        np.add.at(out, idx, vals)
    else:
        import cupyx
        cupyx.scatter_add(out, idx, vals)


def _voigt_to_tensor(xp, s):
    """(..., 6) Voigt (xx, yy, zz, yz, xz, xy) stress -> (..., 3, 3)."""
    T = xp.empty(s.shape[:-1] + (3, 3))
    T[..., 0, 0], T[..., 1, 1], T[..., 2, 2] = s[..., 0], s[..., 1], s[..., 2]
    T[..., 1, 2] = T[..., 2, 1] = s[..., 3]
    T[..., 0, 2] = T[..., 2, 0] = s[..., 4]
    T[..., 0, 1] = T[..., 1, 0] = s[..., 5]
    return T


def _gamma_voigt(xp, grad):
    """Γ_ij = Σ_k ∂_iφ_k ∂_jφ_k as Voigt (Γxx, Γyy, Γzz, 2Γyz, 2Γxz, 2Γxy); grad[..., k, i] = ∂φ_k/∂x_i."""
    G = xp.einsum("...ki,...kj->...ij", grad, grad)
    return xp.stack([G[..., 0, 0], G[..., 1, 1], G[..., 2, 2], 2 * G[..., 1, 2], 2 * G[..., 0, 2], 2 * G[..., 0, 1]], -1)


def _k_orthonormalise(xp, Q, KQ, drop=1e-10):
    """Q ← Q W with (QW)ᵀK(QW) = I, dropping directions whose K-norm² is below drop·max (rank-revealing)."""
    M = Q.T @ KQ
    M = 0.5 * (M + M.T)
    w, U = xp.linalg.eigh(M)
    keep = w > drop * w.max()
    W = U[:, keep] / xp.sqrt(w[keep])[None, :]
    return Q @ W, KQ @ W


def _block_krylov_largest(fe: SparseFE, G, X, nev: int, tol: float, maxiter: int, steps: int = 3,
                          n_neg: int = 4):
    """Largest eigenpairs of G φ = μ K φ (K SPD, G symmetric indefinite).

    Each outer iteration builds the block Krylov space [X, AX, …, A^s X] of A = K⁻¹G (exact factor), K-orthonormalises
    it and does Rayleigh-Ritz on G. The restart block keeps the m most positive Ritz vectors **and** the m most
    negative ones: the negative end of K⁻¹G (tensile or surface-element modes) can dominate |μ|, and keeping it in
    the basis deflates it instead of letting it absorb the power steps. No shift is used (a shift by |μ_min| compresses
    the wanted end towards the rest of the spectrum). Convergence: dual-norm residual ‖Gφ − μKφ‖_{K⁻¹}/|μ| ≤ tol for
    the nev wanted pairs (eigenvalue error ≈ tol²)."""
    xp = fe.xp
    m = X.shape[1]
    res = np.full(nev, np.inf)
    it = 0
    Xn = None
    for it in range(1, maxiter + 1):
        X0 = X if Xn is None else xp.concatenate([X, Xn], 1)
        blocks, V = [X0], X0
        for _ in range(steps):
            V = fe.solve(G @ V)
            V = V / xp.linalg.norm(V, axis=0)[None, :]
            blocks.append(V)
        Q = xp.concatenate(blocks, 1)
        KQ = fe._K @ Q
        Q, KQ = _k_orthonormalise(xp, Q, KQ)
        Q, KQ = _k_orthonormalise(xp, Q, KQ)               # twice is enough (Cholesky-QR style)
        GQ = G @ Q
        Hs = Q.T @ GQ
        theta, Y = xp.linalg.eigh(0.5 * (Hs + Hs.T))      # ascending
        Y_pos, th = Y[:, ::-1][:, :m], theta[::-1][:m]
        Xn = Q @ Y[:, :n_neg]                               # most negative Ritz vectors, kept for deflation
        X = Q @ Y_pos
        R = (GQ @ Y_pos[:, :nev]) - (KQ @ Y_pos[:, :nev]) * th[None, :nev]
        del Q, KQ, GQ, blocks, V
        if xp is not np:
            xp.get_default_memory_pool().free_all_blocks()
        # dual-norm residual ‖r‖_{K⁻¹}/|μ| (φ K-normalised): bounds the eigenvalue error by ≈ res² (THEORY §4.4(c))
        res = _host(xp.sqrt(xp.abs(xp.einsum("ij,ij->j", R, fe.solve(R)))) / (xp.abs(th[:nev]) + 1e-300))
        if _EIG_VERBOSE:
            print(f"  eig it {it}: mu {np.round(_host(th[:nev]), 6).tolist()} res {np.max(res[:nev]):.2e} "
                  f"min-theta {float(theta[0]):.3e}", flush=True)
        if np.all(res[:nev] <= tol):
            break
    return _host(th), X, res, it


def _cudss_mtlayer() -> str | None:
    """Path of cuDSS's OpenMP threading layer (speeds up plan() on the 72-core Grace host), if installed."""
    import glob
    import os
    import sys
    for d in sys.path:
        hits = glob.glob(os.path.join(d, "nvidia", "cu12", "lib", "libcudss_mtlayer_gomp.so*"))
        if hits:
            return sorted(hits)[0]
    return None
