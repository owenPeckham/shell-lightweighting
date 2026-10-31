"""Two-grid PCG engine for the optimiser: matrix-free fine level, Galerkin coarse level, no fine factor.

Plan `continuous/plans/PLAN_twogrid.md` (THEORY §4.1-4.2); prototype and measurements in
`continuous/research/gpu_memory/` (twogrid.py, REPORT.md option 4). Same discrete model as `SparseFE`: q1_im element,
K = Σ_e E ρ̂_e K̂ on an active element set, clamped DOF as identity rows/columns. Same public interface, so
`objective.evaluate`, `buckling()` and `optimise.PhaseModel` use it unchanged.

Fine level (h_e): K u is applied element by element on the GPU (one fused CUDA kernel: gather u_e, K̂ u_e scaled by
E ρ̂_e, scatter-add) for k right-hand sides at once; no assembled fine K and no fine factor. The Jacobi diagonal
is computed element-wise.

Determinism (`deterministic` = "off" (default: atomic kernels, fastest) | "kernels" | "full"; config
`fe.twogrid.deterministic`, env CSHELL_TWOGRID_DET when the argument is None). In "kernels"/"full" every accumulation has a
fixed order (full: bitwise repeatable FE evaluations on the same GPU, but not end to end: design.py still uses atomics). Elements are split into 8 colours by the parity of their (i, j, k): two elements of one colour never
share a node, so within a colour each node is updated by at most one element with a plain (non-atomic) add, and the
colours are applied one kernel launch after another. The K and G products, the Jacobi diagonal, the coarse Galerkin
assembly (coarse elements coloured the same way) and the buckling adjoint load use this. "off" keeps the float64 atomicAdd
kernels (run-to-run noise ~1e-16 per product, amplified by the optimiser's MMA). cuDSS (coarse
factor and solve) is itself not reproducible run to run; "full" switches on its DETERMINISTIC_MODE, which supports one
right-hand side per solve only (the coarse solve then loops over the PCG block's columns); "kernels" keeps the
colour-ordered kernels with the default (faster, non-reproducible) cuDSS.

Coarse level (2 h_e): trilinear prolongation P (node-level, ⊗ I₃) from the coarse nodes of every coarse element that
contains an active fine element. A coarse DOF that interpolates to any clamped fine DOF is fixed (identity row), so
the free coarse columns never touch a clamped fine row and the Galerkin operator K_c = Pᵀ K P is assembled
element-wise without any fine matrix: a fine element is child c of one coarse element, so K_c = Σ_e E ρ̂_e P_cᵀ K̂ P_c
with 8 precomputed 24×24 matrices. K_c is factorised by cuDSS (fp32, fp64 if fp32 breaks down), one factor per design.

Preconditioner: symmetric two-grid V(1,1) cycle, Chebyshev(Jacobi) smoothing of degree `degree` on
[1/r, 1]·1.1 λ_max(D⁻¹K) (λ_max by Lanczos, r = `smoothing_range`), exact coarse solve. PCG in fp64 with the Polak-Ribière β (robust to the
fp32 coarse factor's rounding), one column per right-hand side, run as a block. Non-convergence in `maxiter`
iterations raises; a column whose residual has stagnated (reduction < 2× over 20 iterations, rounding level of an
ill-conditioned K) is accepted with a logged warning if its relative residual is ≤ `stall_accept` (1e-7), and counted
in `stats`/`pcg_summary()`.

Memory (device): _dofs (Ne×24 int32), ρ̂, diag, f, the coarse factor and the PCG/Krylov work blocks; no fine matrix.
Host: only the active element/node ids of the grid (interface of `ActiveGrid`), built on the GPU (`LeanGrid`).
Units mm, N, MPa.
"""
from __future__ import annotations

import os
import time

import numpy as np

from .fe_sparse import (CORNERS, Primal, SparseFE, _cuda_ok, _cudss_mtlayer, _host, _release_device_memory,
                        _scatter_add, _stencil_slot)
from . import fea  # noqa: F401  (puts continuous/reference on sys.path)
import gpu_analysis

_CHUNK = 1 << 16                      # elements per batched host-side (NumPy) or set-up chunk

# ---------------------------------------------------------------------------------------------------------------------
# CUDA kernels: y += Σ_e s_e K̂ u_e (q1_im, constant K̂) and y += Σ_e sgn H_e ⊗ I₃ v_e (node-level element matrices)
# ---------------------------------------------------------------------------------------------------------------------
_KERNELS = r"""
extern "C" __global__ void k_apply(const double* __restrict__ X, double* __restrict__ Y,
                                   const int* __restrict__ dofs, const double* __restrict__ s,
                                   const unsigned char* __restrict__ fixed, const double* __restrict__ Khat,
                                   const int* __restrict__ elist, const long long ne, const int k, const int atom)
{
    __shared__ double Ks[576];
    for (int i = threadIdx.x; i < 576; i += blockDim.x) Ks[i] = Khat[i];
    __syncthreads();
    long long t = (long long)blockIdx.x * blockDim.x + threadIdx.x;
    if (t >= ne * k) return;
    long long e = t / k;
    int j = (int)(t - e * k);
    if (!atom) e = elist[e];
    int d[24];
    double u[24];
    #pragma unroll
    for (int a = 0; a < 24; a++) {
        d[a] = dofs[e * 24 + a];
        u[a] = fixed[d[a]] ? 0.0 : X[(long long)d[a] * k + j];
    }
    const double se = s[e];
    #pragma unroll 4
    for (int a = 0; a < 24; a++) {
        if (fixed[d[a]]) continue;
        double acc = 0.0;
        #pragma unroll
        for (int b = 0; b < 24; b++) acc += Ks[a * 24 + b] * u[b];
        if (atom) atomicAdd(&Y[(long long)d[a] * k + j], se * acc);
        else Y[(long long)d[a] * k + j] += se * acc;
    }
}

extern "C" __global__ void g_apply(const double* __restrict__ X, double* __restrict__ Y,
                                   const int* __restrict__ dofs, const double* __restrict__ H,
                                   const unsigned char* __restrict__ fixed, const double sgn,
                                   const int* __restrict__ elist, const long long ne, const int k, const int atom)
{
    long long t = (long long)blockIdx.x * blockDim.x + threadIdx.x;
    if (t >= ne * k) return;
    long long e = t / k;
    int j = (int)(t - e * k);
    if (!atom) e = elist[e];
    double h[64];
    #pragma unroll
    for (int i = 0; i < 64; i++) h[i] = H[e * 64 + i];
    for (int c = 0; c < 3; c++) {
        int d[8];
        double v[8];
        #pragma unroll
        for (int b = 0; b < 8; b++) {
            d[b] = dofs[e * 24 + 3 * b + c];
            v[b] = fixed[d[b]] ? 0.0 : X[(long long)d[b] * k + j];
        }
        #pragma unroll
        for (int a = 0; a < 8; a++) {
            if (fixed[d[a]]) continue;
            double acc = 0.0;
            #pragma unroll
            for (int b = 0; b < 8; b++) acc += h[a * 8 + b] * v[b];
            if (atom) atomicAdd(&Y[(long long)d[a] * k + j], sgn * acc);
            else Y[(long long)d[a] * k + j] += sgn * acc;
        }
    }
}
"""
_MOD = None


def _kernels():
    global _MOD
    if _MOD is None:
        import cupy as cp
        _MOD = cp.RawModule(code=_KERNELS)
    return _MOD.get_function("k_apply"), _MOD.get_function("g_apply")


_EW = {}


def _ew(name):
    """Fused in-place block updates (no n×k temporaries): axpy y += a x, xpay p = z + b p, cheb d = a d + c D⁻¹ r;
    a, b, c broadcast (scalars or (1, k) per-column factors, D⁻¹ as (n, 1))."""
    if not _EW:
        import cupy as cp
        _EW["axpy"] = cp.ElementwiseKernel("T x, T a", "T y", "y += a * x", "tg_axpy")
        _EW["xpay"] = cp.ElementwiseKernel("T z, T b", "T p", "p = z + b * p", "tg_xpay")
        _EW["cheb"] = cp.ElementwiseKernel("T r, T dinv, T a, T c", "T d", "d = a * d + c * dinv * r", "tg_cheb")
        # in-place indexed adds out[idx] += v; callers guarantee distinct idx within one launch (one colour)
        _EW["idxadd"] = cp.ElementwiseKernel("T v, S idx", "raw T out", "out[idx] += v", "tg_idxadd")
        _EW["idxadd9"] = cp.ElementwiseKernel("raw T v, S idx", "raw T out",
                                              "for (int q = 0; q < 9; q++) out[idx * 9 + q] += v[i * 9 + q];",
                                              "tg_idxadd9")
        _EW["dot"] = cp.ReductionKernel("T x, T y", "T z", "x * y", "a + b", "z = a", "0", "tg_coldot")
    return _EW[name]


def _coldot(xp, A, B):
    """Column-wise dot products Σ_i A_ij B_ij without an n×k temporary."""
    if xp is np:
        return np.einsum("ij,ij->j", A, B)
    return _ew("dot")(A, B, axis=0)


def _axpy(xp, Y, X, a):
    if xp is np:
        Y += a * X
    else:
        _ew("axpy")(X, a, Y)


def _xpay(xp, P, Z, b):
    if xp is np:
        P *= b
        P += Z
    else:
        _ew("xpay")(Z, b, P)


def _idx_add(xp, out, idx, vals):
    """out[idx] += vals (1-D out, 1-D idx, vals same length) for DISTINCT idx: plain in-place add, no atomics."""
    if xp is np:
        out[idx] += vals
    else:
        _ew("idxadd")(vals, idx, out)


def _blk_add(xp, out, pos, vals):
    """out[pos] += vals for out (nblk, 3, 3), pos (m,) DISTINCT, vals (m, 3, 3)."""
    if xp is np:
        out[pos] += vals
    else:
        _ew("idxadd9")(xp.ascontiguousarray(vals).reshape(-1), pos, out)


def _colour_lists(xp, ei, ej, ek):
    """Element ids (ascending, int32) of each of the 8 parity colours of the element indices (i, j, k)."""
    col = 4 * (ei % 2) + 2 * (ej % 2) + ek % 2
    return [l for l in (xp.flatnonzero(col == c).astype(xp.int32) for c in range(8)) if l.size]


def _bincount_add(Y, idx, vals):
    """Y[idx] += vals with repeated indices (NumPy; np.add.at is ~50x slower)."""
    for j in range(Y.shape[1]):
        Y[:, j] += np.bincount(idx, weights=vals[:, j], minlength=Y.shape[0])


def _xp_default():
    if _cuda_ok():
        import cupy as cp
        return cp
    return np


def _sparse(xp):
    if xp is np:
        import scipy.sparse as sps
        return sps
    import cupyx.scipy.sparse as cps
    return cps


# ---------------------------------------------------------------------------------------------------------------------
# Grids
# ---------------------------------------------------------------------------------------------------------------------
class LeanGrid:
    """Active elements and nodes of a dense element box without the CSR pattern of `ActiveGrid` (same numbering:
    sorted dense ids, the reference's corner order). Connectivity is built on the GPU when one is present; the host
    keeps only the element and node ids that gather/scatter need."""

    def __init__(self, elem_mask: np.ndarray, xp=None):
        xp = _xp_default() if xp is None else xp
        elem_mask = np.asarray(elem_mask, bool)
        self.elem_shape = tuple(int(n) for n in elem_mask.shape)
        self.node_shape = tuple(n + 1 for n in self.elem_shape)
        self.elem_dense = np.flatnonzero(elem_mask.ravel())
        if self.elem_dense.size == 0:
            raise ValueError("empty active set")
        node_dense, conn = _connectivity(xp, xp.asarray(self.elem_dense), self.elem_shape)
        self.node_dense = _host_int(node_dense)
        self.conn = conn                                               # (Ne, 8) int32 on xp; consumed by TwoGridFE
        self.Ne, self.Nn = int(conn.shape[0]), int(node_dense.size)
        self.ndof = 3 * self.Nn
        self.nnz = None

    # same gather/scatter semantics as ActiveGrid
    def gather_nodes(self, node_field):
        a = np.asarray(node_field)
        return a.reshape((-1,) + a.shape[3:])[self.node_dense]

    def scatter_nodes(self, vals, fill=0.0):
        out = np.full((int(np.prod(self.node_shape)),) + vals.shape[1:], fill, dtype=vals.dtype)
        out[self.node_dense] = vals
        return out.reshape(self.node_shape + vals.shape[1:])

    def gather_elems(self, elem_field):
        return np.asarray(elem_field).reshape(-1)[self.elem_dense]

    def scatter_elems(self, vals, fill=0.0):
        out = np.full(int(np.prod(self.elem_shape)), fill, dtype=np.asarray(vals).dtype)
        out[self.elem_dense] = vals
        return out.reshape(self.elem_shape)


def _host_int(a) -> np.ndarray:
    return a.get() if hasattr(a, "get") else np.asarray(a)


def _unravel(xp, flat, shape):
    ny, nz = shape[1], shape[2]
    return flat // (ny * nz), (flat // nz) % ny, flat % nz


def _connectivity(xp, elem_dense, elem_shape):
    """Sorted active node ids (dense node box) and the (Ne, 8) int32 connectivity in the reference corner order."""
    nyn, nzn = elem_shape[1] + 1, elem_shape[2] + 1
    ei, ej, ek = _unravel(xp, elem_dense.astype(xp.int64), elem_shape)
    C = xp.asarray(CORNERS)
    nd = ((ei[:, None] + C[:, 0]) * nyn + ej[:, None] + C[:, 1]) * nzn + ek[:, None] + C[:, 2]
    del ei, ej, ek
    node_dense, conn = xp.unique(nd.ravel(), return_inverse=True)
    return node_dense, conn.reshape(-1, 8).astype(xp.int32)


class _BlockPattern:
    """CSR pattern of a 3-DOF-per-node voxel operator, built with `xp` (the GPU port of `ActiveGrid`'s pattern, used
    for the coarse level). Block (n, k) entry (p, q) sits at 9·bptr[n] + 3p·nnb[n] + 3k + q."""

    def __init__(self, xp, elem_dense, elem_shape):
        self.elem_shape = tuple(elem_shape)
        self.node_shape = tuple(n + 1 for n in elem_shape)
        self.node_dense, conn = _connectivity(xp, elem_dense, elem_shape)
        conn = conn.astype(xp.int64)
        self.colours = _colour_lists(xp, *_unravel(xp, elem_dense.astype(xp.int64), elem_shape))
        self.Ne, self.Nn = int(conn.shape[0]), int(self.node_dense.size)
        self.ndof = 3 * self.Nn
        slot_ab = _stencil_slot(CORNERS[None, :, :] - CORNERS[:, None, :])
        present = xp.zeros((self.Nn, 27), bool)
        for a in range(8):
            for b in range(8):
                present[conn[:, a], int(slot_ab[a, b])] = True
        nnb = present.sum(1).astype(xp.int64)
        bptr = xp.concatenate([xp.zeros(1, xp.int64), xp.cumsum(nnb)])
        rank = xp.cumsum(present, 1).astype(xp.int64) - 1
        self.nblk = int(bptr[-1])
        self.nnz = 9 * self.nblk
        it = np.int32 if self.nnz < 2 ** 31 else np.int64
        pos = xp.empty((self.Ne, 64), it)
        for a in range(8):
            na = conn[:, a]
            for b in range(8):
                pos[:, 8 * a + b] = bptr[na] + rank[na, int(slot_ab[a, b])]
        self.pos_blk = pos
        rr, ss = xp.nonzero(present)
        del present
        ni, nj, nk = _unravel(xp, self.node_dense[rr], self.node_shape)
        offs = xp.asarray(np.stack(np.unravel_index(np.arange(27), (3, 3, 3)), 1) - 1)
        nb_dense = ((ni + offs[ss, 0]) * self.node_shape[1] + nj + offs[ss, 1]) * self.node_shape[2] + nk + offs[ss, 2]
        blk_col = xp.searchsorted(self.node_dense, nb_dense)
        del ni, nj, nk, nb_dense, ss
        kk = xp.arange(self.nblk, dtype=xp.int64) - bptr[rr]
        p = xp.arange(3, dtype=xp.int64)
        perm = (9 * bptr[rr][:, None, None] + 3 * p[None, :, None] * nnb[rr][:, None, None]
                + 3 * kk[:, None, None] + p[None, None, :])
        self.perm = perm.ravel().astype(it)
        ind = xp.empty(self.nnz, np.int32)
        ind[self.perm] = (3 * blk_col[:, None, None] + p[None, None, :] + 0 * p[None, :, None]).ravel().astype(np.int32)
        self.indices = ind
        del perm, kk, blk_col, rr
        self.indptr = xp.zeros(self.ndof + 1, it)
        self.indptr[1:] = xp.cumsum(xp.repeat(3 * nnb, 3))
        k13 = rank[:, 13]
        self.diag_pos = (9 * bptr[:-1, None] + 3 * p[None, :] * nnb[:, None] + 3 * k13[:, None] + p[None, :]).ravel()


# ---------------------------------------------------------------------------------------------------------------------
# Coarse direct solver
# ---------------------------------------------------------------------------------------------------------------------
class _CoarseSolver:
    """K_c factorised by cuDSS (SPD, fp32 or fp64, `nrhs` columns per call) or SciPy SuperLU on the CPU."""

    def __init__(self, xp, pat: _BlockPattern, precision: str = "auto", nrhs: int = 14, deterministic: bool = True):
        if precision not in ("auto", "fp32", "fp64"):
            raise ValueError("coarse precision must be 'auto', 'fp32' or 'fp64'")
        self.xp, self.pat, self.nrhs = xp, pat, max(int(nrhs), 1)
        self.deterministic = bool(deterministic)
        self.auto = precision == "auto"
        self.precision = "fp32" if precision == "auto" else precision
        self.dt = np.float32 if self.precision == "fp32" else np.float64
        self._K = self._b = self._solver = None
        self.stats = dict(coarse_precision=self.precision)

    def factor(self, data):
        """Factorise K_c. "auto" precision: fp32, checked by one probe solve; a non-finite result (the fp32 Cholesky
        broke down: K_c inherits the ρ_min contrast of the active-set margin) switches to fp64 for this engine, with a
        logged warning."""
        self._factor(data)
        if self.xp is np or self.precision == "fp64":
            return
        xp = self.xp
        probe = self.solve(xp.asarray(np.random.default_rng(5).standard_normal((self.pat.ndof, 1))))
        if bool(xp.isfinite(probe).all()):
            return
        if not self.auto:
            raise FloatingPointError("two-grid: fp32 coarse factor broke down (non-finite solve); use coarse "
                                     "precision 'auto' or 'fp64'")
        print(f"fe_twogrid: WARNING: fp32 coarse factor broke down ({self.pat.ndof} coarse DOF); refactorising in "
              f"fp64", flush=True)
        self.free()
        self.precision, self.dt = "fp64", np.float64
        self.stats.update(coarse_precision="fp64", coarse_fp32_breakdown=True)
        self._factor(data)

    def _factor(self, data):
        xp, pat = self.xp, self.pat
        t0 = time.time()
        if xp is np:
            import scipy.sparse as sps
            import scipy.sparse.linalg as spla
            K = sps.csr_matrix((data, pat.indices, pat.indptr), shape=(pat.ndof, pat.ndof))
            self._solver = spla.splu(K.tocsc())
        elif self._solver is None:
            from nvmath.sparse.advanced import DirectSolver, DirectSolverMatrixType, DirectSolverOptions, ExecutionCUDA
            sp = _sparse(xp)
            self._K = sp.csr_matrix((data.astype(self.dt), pat.indices, pat.indptr.astype(np.int32)),
                                    shape=(pat.ndof, pat.ndof))
            self._b = xp.asfortranarray(xp.zeros((pat.ndof, self.nrhs), dtype=self.dt))
            opts = DirectSolverOptions(sparse_system_type=DirectSolverMatrixType.SPD,
                                       multithreading_lib=_cudss_mtlayer())
            self._solver = DirectSolver(self._K, self._b, options=opts, execution=ExecutionCUDA())
            if self.deterministic:      # cuDSS run-to-run reproducibility (not exposed by nvmath's option classes)
                from nvmath.bindings import cudss
                one = np.array(1, dtype=np.int32)
                cudss.config_set(self._solver.config_ptr, cudss.ConfigParam.DETERMINISTIC_MODE, one.ctypes.data,
                                 one.dtype.itemsize)
            self._solver.plan()
            est = self._solver.plan_info.memory_estimates
            self.stats["coarse_factor_estimate_GiB"] = float(est["peak_device_memory"]) / 2 ** 30
            self.stats["coarse_plan_s"] = time.time() - t0
            self._solver.factorize()
        else:
            self._K.data[...] = data.astype(self.dt)
            self._solver.factorize()
        if xp is not np:
            xp.cuda.Device().synchronize()
        self.stats["coarse_factor_s"] = time.time() - t0

    def solve(self, R):
        """(nc, k) fp64 → (nc, k) fp64."""
        xp = self.xp
        if xp is np:
            return self._solver.solve(np.asarray(R, np.float64))
        X = xp.empty(R.shape, xp.float64)
        for j0 in range(0, R.shape[1], self.nrhs):
            j1 = min(j0 + self.nrhs, R.shape[1])
            k = j1 - j0
            if k < self.nrhs:
                self._b[:, k:] = 0.0
            self._b[:, :k] = R[:, j0:j1]
            X[:, j0:j1] = xp.asarray(self._solver.solve())[:, :k]
        return X

    def free(self):
        if self._solver is not None and hasattr(self._solver, "free"):
            self._solver.free()
        self._solver = self._K = self._b = None
        if self.xp is not np:
            _release_device_memory(self.xp)


def child_P() -> np.ndarray:
    """P_c (8, 24, 24): fine DOF (a, k) of child c (fine offset c + C[a] ∈ {0,1,2}³ inside its coarse element) from
    coarse DOF (A, k) with trilinear weight N_A(ξ), ξ = (c + C[a])/2."""
    out = np.zeros((8, 24, 24))
    for ci, c in enumerate(CORNERS):
        for a in range(8):
            xi = (c + CORNERS[a]) / 2.0
            for A in range(8):
                w = np.prod(np.where(CORNERS[A] == 1, xi, 1 - xi))
                for k in range(3):
                    out[ci, 3 * a + k, 3 * A + k] = w
    return out


# ---------------------------------------------------------------------------------------------------------------------
# Operators
# ---------------------------------------------------------------------------------------------------------------------
class _MatFreeK:
    """K = Σ_e E ρ̂_e K̂ (identity on clamped DOF), applied element by element; `K @ x` for (n,) or (n, k)."""

    def __init__(self, fe: "TwoGridFE"):
        self.fe = fe
        self.shape = (fe.g.ndof, fe.g.ndof)

    def __matmul__(self, X):
        return self.fe._apply_K(X)

    def diagonal(self):
        return self.fe._diag


class _ElemGeomOp:
    """G = −K_σ = −Σ_e H_e ⊗ I₃ (rows/columns of clamped DOF zero), applied element by element."""

    def __init__(self, fe: "TwoGridFE", H):
        self.fe, self.H = fe, fe.xp.ascontiguousarray(H.reshape(-1, 64))

    def __matmul__(self, V):
        fe, xp = self.fe, self.fe.xp
        one = V.ndim == 1
        X = xp.ascontiguousarray(V.reshape(-1, 1) if one else V, dtype=xp.float64)
        k = X.shape[1]
        Y = xp.zeros_like(X)
        if xp is np:
            X0 = X.copy()
            X0[fe.fixed] = 0.0
            Xn = X0.reshape(-1, 3 * k)
            nodes = fe._dofs[:, 0::3] // 3
            Yn = Y.reshape(-1, 3 * k)
            for s in range(0, fe.g.Ne, _CHUNK):
                n_ = nodes[s:s + _CHUNK]
                ye = -np.matmul(self.H[s:s + _CHUNK].reshape(-1, 8, 8), Xn[n_])     # batched GEMM, not einsum
                _bincount_add(Yn, n_.ravel(), ye.reshape(-1, 3 * k))
            Y[fe.fixed] = 0.0
        else:
            _, g_apply = _kernels()
            for el, ne, atom in fe._launches():
                n = ne * k
                g_apply(((n + 127) // 128,), (128,), (X, Y, fe._dofs, self.H, fe._fixed_u8, np.float64(-1.0),
                                                      el, np.int64(ne), np.int32(k), np.int32(atom)))
        return Y[:, 0] if one else Y


# ---------------------------------------------------------------------------------------------------------------------
# Engine
# ---------------------------------------------------------------------------------------------------------------------
class TwoGridFE(SparseFE):
    """K(ρ̂) on an active grid (`LeanGrid` or `ActiveGrid`), solved by two-grid PCG; interface of `SparseFE`.

    factor(ρ̂) assembles the coarse operator and factorises it, computes the Jacobi diagonal and λ_max(D⁻¹K);
    solve(rhs) runs PCG to `tol` (relative residual per column); solve_inner(rhs) to `krylov_tol` (used by the
    buckling block-Krylov expansion, whose Rayleigh-Ritz step and residual test keep the eigenpairs exact).
    primal() warm-starts from the previous design's displacement."""

    engine = "twogrid"

    def __init__(self, grid, fixed_dof: np.ndarray, nodal_load: np.ndarray, h_mm: float, E_MPa: float, nu: float,
                 backend: str = "auto", element_type: str = "q1_im", coarse_precision: str = "auto",
                 degree: int = 5, tol: float = 1e-10, krylov_tol: float = 1e-2, maxiter: int = 100,
                 nrhs: int = 7, warm_start: bool = True, lanczos_steps: int = 20, smoothing_range: float = 30.0,
                 deterministic=None, stall_accept: float = 1e-7, stall_window: int = 20):
        t0 = time.time()
        self.g, self.h, self.E, self.nu = grid, float(h_mm), float(E_MPa), float(nu)
        self.element_type = element_type
        if backend == "auto":
            backend = "gpu" if _cuda_ok() else "cpu"
        self.backend = backend
        if backend == "gpu":
            import cupy as cp
            self.xp = cp
        else:
            self.xp = np
        xp = self.xp
        self._sp = _sparse(xp)
        self.degree, self.tol, self.krylov_tol, self.maxiter = int(degree), float(tol), float(krylov_tol), int(maxiter)
        self.nrhs, self.warm_start, self.lanczos_steps = max(int(nrhs), 1), bool(warm_start), int(lanczos_steps)
        self.smoothing_range = float(smoothing_range)
        mode = os.environ.get("CSHELL_TWOGRID_DET", "off") if deterministic is None else deterministic
        mode = {True: "full", False: "off"}.get(mode, mode)
        if mode not in ("full", "kernels", "off"):
            raise ValueError("deterministic must be 'full', 'kernels' or 'off' (or True/False)")
        self.det_mode = mode
        self.deterministic = mode != "off"                                  # colour-ordered kernels and assembly
        self.cudss_deterministic = mode == "full"                           # + cuDSS DETERMINISTIC_MODE (nrhs = 1)
        self.stall_accept, self.stall_window = float(stall_accept), int(stall_window)
        self.precision = "pcg-fp64"
        self.hybrid = False
        self.Khat = gpu_analysis.element_khat(element_type, (self.h,) * 3, self.nu)
        fx = grid.gather_nodes(np.asarray(fixed_dof, bool)).reshape(-1)
        f = grid.gather_nodes(np.asarray(nodal_load, float)).reshape(-1)
        tot_dense = np.asarray(nodal_load, float).reshape(-1, 3).sum(0)
        if not np.allclose(f.reshape(-1, 3).sum(0), tot_dense):
            raise ValueError("some loaded nodes are outside the active set")
        if not fx.any():
            raise ValueError("no fixed DOF in the active set")
        clash = fx & (f != 0)
        if clash.any():
            print(f"fe_twogrid: warning: {int(clash.sum())} loaded DOF are clamped; load removed", flush=True)
            f[clash] = 0.0
        self.fixed = xp.asarray(fx)                                         # device mask (interface: indexable)
        self._fixed_x = self.fixed
        self._fixed_u8 = self.fixed.astype(xp.uint8)
        self.f = xp.asarray(f)
        del fx, f
        conn = xp.asarray(_host_int(grid.conn) if xp is np else grid.conn).astype(xp.int32)   # grid may be on the GPU
        p3 = xp.arange(3, dtype=xp.int32)
        self._dofs = (3 * conn[:, :, None] + p3[None, None, :]).reshape(grid.Ne, 24)
        if isinstance(grid, LeanGrid):
            grid.conn = None                                                  # the engine owns the connectivity now
        self._Khat = xp.asarray(self.Khat)
        B_all, Bg_all, detJ = gpu_analysis.hex_reference_data(self.h, self.h, self.h)
        self._detJ = float(detJ)
        from .fe_sparse import unit_D
        self._DB = xp.asarray(np.einsum("ij,gjk->gik", unit_D(self.nu), B_all))
        self._Bg = xp.asarray(Bg_all)
        self._K = _MatFreeK(self)
        self._u_prev = None
        self.rho_hat = None
        self._setup_coarse(conn)
        del conn
        self.stats = dict(engine="twogrid", Ne=grid.Ne, ndof=grid.ndof, backend=backend, coarse_ndof=self._pat.ndof,
                          coarse_nnz=self._pat.nnz, coarse_fixed_dof=self._nb, smoothing_range=self.smoothing_range,
                          deterministic=self.det_mode, maxiter=self.maxiter, degree=self.degree, tol=self.tol,
                          krylov_tol=self.krylov_tol, setup_s=time.time() - t0)
        self._co = _CoarseSolver(xp, self._pat, coarse_precision, 1 if self.cudss_deterministic else self.nrhs,
                                 self.cudss_deterministic)      # cuDSS deterministic mode supports one right-hand side
        self.stats["coarse_precision"] = self._co.precision
        self._reset_counters()

    # -- set-up (design independent) ---------------------------------------------------------------------------
    def _setup_coarse(self, conn):
        """Coarse elements, child index of each fine element, coarse pattern, prolongation and coarse clamps."""
        xp, g = self.xp, self.g
        es = g.elem_shape
        esc = tuple((n + 1) // 2 for n in es)
        ed = xp.asarray(g.elem_dense, dtype=xp.int64)
        i, j, k = _unravel(xp, ed, es)
        del ed
        pdense = ((i // 2) * esc[1] + j // 2) * esc[2] + k // 2
        lut = np.zeros(8, np.int64)
        for c, (x, y, z) in enumerate(CORNERS):
            lut[4 * x + 2 * y + z] = c
        child = xp.asarray(lut)[4 * (i % 2) + 2 * (j % 2) + k % 2].astype(xp.int8)
        del i, j, k
        cdense, parent = xp.unique(pdense, return_inverse=True)
        del pdense
        self._parent, self._child = parent.astype(xp.int32), child
        # fine-element colours = child index (the parity of (i, j, k)); a colour has no two elements sharing a node
        self._colours = [l for l in (xp.flatnonzero(child == c).astype(xp.int32) for c in range(8)) if l.size]
        self._null = xp.zeros(1, xp.int32)                                      # "no element list" kernel argument
        self._pat = _BlockPattern(xp, cdense, esc)
        del cdense
        pat = self._pat
        # prolongation, node level: fine node at ijk interpolates from coarse nodes ijk//2 (+1 on odd axes)
        nd = xp.asarray(g.node_dense, dtype=xp.int64)
        fi = xp.stack(_unravel(xp, nd, g.node_shape), 1)
        del nd
        odd = fi % 2
        nnc = pat.node_shape
        rows, cols, vals = [], [], []
        for corner in CORNERS:
            cr = xp.asarray(corner)
            take = xp.flatnonzero(xp.all((cr == 0)[None, :] | (odd == 1), 1))
            cn = fi[take] // 2 + cr[None, :] * odd[take]
            w = xp.prod(xp.where(odd[take] == 1, 0.5, 1.0), 1)
            cdn = (cn[:, 0] * nnc[1] + cn[:, 1]) * nnc[2] + cn[:, 2]
            cd = xp.searchsorted(pat.node_dense, cdn)
            if not bool(xp.all(pat.node_dense[xp.minimum(cd, pat.Nn - 1)] == cdn)):
                raise RuntimeError("two-grid: a fine node interpolates from an inactive coarse node")
            rows.append(take); cols.append(cd); vals.append(w)
            del cn, cdn
        del fi, odd
        sp = self._sp
        Pn = sp.coo_matrix((xp.concatenate(vals), (xp.concatenate(rows).astype(xp.int32),
                                                    xp.concatenate(cols).astype(xp.int32))),
                           shape=(g.Nn, pat.Nn)).tocsr()
        del rows, cols, vals
        self._Pn, self._PnT = Pn, Pn.T.tocsr()
        # Coarse clamps: a coarse DOF is fixed if any fine DOF it interpolates to (same component) is clamped. The
        # remaining columns of P then never touch a clamped fine row, so P_m = P on them and K_c = Pᵀ K P is assembled
        # from the unmasked child matrices; P keeps full column rank on any subset of columns, so K_c is SPD. (Fixing
        # only coarse DOF with an entirely clamped support, the plan's first rule, made masked columns linearly
        # dependent near the MBB's cylindrical clamps: singular K_c, NaN in fp32 and fp64.)
        clamped = self.fixed.astype(xp.float64).reshape(-1, 3)
        self._fixed_c = ((self._PnT @ clamped) > 0).reshape(-1)
        del clamped
        self._diag_fixed_c = pat.diag_pos[self._fixed_c]
        rows = xp.searchsorted(pat.indptr[1:], xp.arange(pat.nnz, dtype=pat.indptr.dtype), side="right")
        self._drop_c = xp.flatnonzero(self._fixed_c[rows] | self._fixed_c[pat.indices]).astype(pat.perm.dtype)
        del rows
        Pc = child_P()                                                      # child matrices P_cᵀ K̂ P_c (unit E)
        self._Mc = xp.asarray(np.einsum("cia,ij,cjb->cab", Pc, self.Khat, Pc).reshape(8, 576))
        self._nb = int(self._fixed_c.sum())

    # -- factorisation ---------------------------------------------------------------------------------------------
    def factor(self, rho_hat_active):
        """Coarse Galerkin operator + factor, Jacobi diagonal and the Chebyshev interval for ρ̂ (length Ne)."""
        xp, g = self.xp, self.g
        t0 = time.time()
        rho = xp.asarray(rho_hat_active, dtype=xp.float64)
        if rho.shape[0] != g.Ne:
            raise ValueError("rho_hat must have one value per active element")
        self.rho_hat = rho
        self._s = self.E * rho                                              # element scale E ρ̂_e
        if xp is np:                    # CPU (tests, small problems): assembled K, a NumPy element loop is ~100× slower
            d = self._dofs.astype(np.int64)
            Ke = self._s[:, None, None] * self.Khat[None]
            self._Kcpu = self._sp.csr_matrix((Ke.ravel(), (np.repeat(d, 24, 1).ravel(), np.tile(d, (1, 24)).ravel())),
                                             shape=(g.ndof, g.ndof))
            del d, Ke
        # Jacobi diagonal
        dK = xp.asarray(np.diag(self.Khat).copy())
        diag = xp.zeros(g.ndof)
        if self.deterministic:
            for lst in self._colours:
                for s in range(0, lst.size, 1 << 19):
                    sel = lst[s:s + (1 << 19)]
                    _idx_add(xp, diag, self._dofs[sel].reshape(-1), (self._s[sel, None] * dK[None, :]).reshape(-1))
        else:
            for s in range(0, g.Ne, 1 << 20):
                _scatter_add(xp, diag, self._dofs[s:s + (1 << 20)].reshape(-1),
                             (self._s[s:s + (1 << 20), None] * dK[None, :]).reshape(-1))
        diag[self.fixed] = 1.0
        self._diag = diag
        self._dinv = 1.0 / diag
        t1 = time.time()
        data = self._coarse_values()
        self._co.factor(data)
        del data
        t2 = time.time()
        lam = self._lanczos_lmax()
        self._lmax = 1.1 * lam
        self._lmin = self._lmax / self.smoothing_range
        if xp is not np:
            xp.cuda.Device().synchronize()
        self._reset_counters()
        self.stats.update(assemble_s=t1 - t0, factor_s=time.time() - t0, coarse_s=t2 - t1, lmax_DinvK=lam,
                          lanczos_s=time.time() - t2, **self._co.stats)
        return self

    def _coarse_values(self):
        """CSR data of K_c = Σ_e E ρ̂_e P_cᵀ K̂ P_c, rows/columns of fixed coarse DOF replaced by identity."""
        xp, pat = self.xp, self._pat
        rho_c = xp.zeros((pat.Ne, 8))
        rho_c[self._parent, self._child.astype(xp.int64)] = self._s
        blocks = xp.zeros((pat.nblk, 3, 3))
        if self.deterministic:      # colour by colour: within one colour every (element, a, b) hits a distinct block
            chunks = [(lst[s:s + _CHUNK]) for lst in pat.colours for s in range(0, lst.size, _CHUNK)]
        else:
            chunks = [slice(s, s + _CHUNK) for s in range(0, pat.Ne, _CHUNK)]
        for sel in chunks:
            Ke = rho_c[sel] @ self._Mc                                      # (n, 576) coarse-element matrices
            Kb = Ke.reshape(-1, 8, 3, 8, 3).transpose(0, 1, 3, 2, 4).reshape(-1, 3, 3)
            if self.deterministic:
                _blk_add(xp, blocks, pat.pos_blk[sel].reshape(-1), Kb)
            else:
                _scatter_add(xp, blocks, pat.pos_blk[sel].reshape(-1), Kb)
            del Ke, Kb
        del rho_c
        data = xp.empty(pat.nnz)
        data[pat.perm] = blocks.reshape(-1)
        del blocks
        data[self._drop_c] = 0.0
        data[self._diag_fixed_c] = 1.0
        return data

    def _lanczos_lmax(self) -> float:
        """Largest eigenvalue of D⁻¹K (Lanczos on D^{-1/2} K D^{-1/2}, fixed seed); the Ritz value is a lower bound,
        so the Chebyshev interval uses 1.1 times it (PETSc's default safety)."""
        xp = self.xp
        dh = xp.sqrt(self._dinv)
        rng = np.random.default_rng(11)
        v = xp.asarray(rng.standard_normal(self.g.ndof))
        v /= xp.linalg.norm(v)
        v_old = xp.zeros_like(v)
        al, be = [], []
        beta = 0.0
        for _ in range(self.lanczos_steps):
            w = dh * self._apply_K(dh * v) - beta * v_old
            a = float(w @ v)
            w -= a * v
            beta = float(xp.linalg.norm(w))
            al.append(a)
            if beta < 1e-14:
                break
            be.append(beta)
            v_old, v = v, w / beta
        T = np.diag(al) + np.diag(be[:len(al) - 1], 1) + np.diag(be[:len(al) - 1], -1)
        return float(np.linalg.eigvalsh(T).max())

    # -- operators -------------------------------------------------------------------------------------------------
    def _apply_K(self, V):
        xp = self.xp
        one = V.ndim == 1
        X = xp.ascontiguousarray(V.reshape(-1, 1) if one else V, dtype=xp.float64)
        k = X.shape[1]
        Y = xp.zeros_like(X)
        if xp is np:
            X0 = X.copy()
            X0[self.fixed] = 0.0
            Y = np.asarray(self._Kcpu @ X0)
        else:
            k_apply, _ = _kernels()
            for el, ne, atom in self._launches():
                n = ne * k
                k_apply(((n + 127) // 128,), (128,), (X, Y, self._dofs, self._s, self._fixed_u8, self._Khat,
                                                      el, np.int64(ne), np.int32(k), np.int32(atom)))
        Y[self.fixed] = X[self.fixed]
        return Y[:, 0] if one else Y

    def _launches(self):
        """(element list, count, atomic flag) per kernel launch: the 8 colours in fixed order (deterministic) or one
        launch over all elements with atomics."""
        if self.deterministic:
            return [(l, int(l.size), 0) for l in self._colours]
        return [(self._null, self.g.Ne, 1)]

    def _scatter_dofs(self, out, vals):
        """out (ndof,) += Σ_e (per-element DOF values vals (Ne, 24)) scattered to `_dofs`, in a fixed order."""
        xp = self.xp
        if not self.deterministic:
            return SparseFE._scatter_dofs(self, out, vals)
        for lst in self._colours:
            for s in range(0, lst.size, 1 << 19):
                sel = lst[s:s + (1 << 19)]
                _idx_add(xp, out, self._dofs[sel].reshape(-1), vals[sel].reshape(-1))
        return out

    def matvec(self, x):
        return self._apply_K(x)

    def _geometric_op(self, H):
        return _ElemGeomOp(self, H)

    def _cheb(self, B, overwrite: bool = False):
        """Chebyshev(Jacobi) approximation of K⁻¹B, zero initial guess, `degree` terms (degree − 1 products). Work
        blocks: R, D, X and one K·D; with overwrite=True, B itself serves as R."""
        xp = self.xp
        lmax, lmin = self._lmax, self._lmin
        th, de = 0.5 * (lmax + lmin), 0.5 * (lmax - lmin)
        sig = th / de
        ro = 1.0 / sig
        Dinv = self._dinv[:, None]
        D = B * (Dinv / th)
        X = D.copy()
        R = B if overwrite else B.copy()
        for _ in range(self.degree - 1):
            KD = self._apply_K(D)
            xp.subtract(R, KD, out=R)
            del KD
            rn = 1.0 / (2.0 * sig - ro)
            if xp is np:
                D *= rn * ro
                D += (2.0 * rn / de) * (Dinv * R)
            else:
                _ew("cheb")(R, Dinv, rn * ro, 2.0 * rn / de, D)
            X += D
            ro = rn
        return X

    def _coarse_correction(self, R, overwrite: bool = False):
        """P_m K_c⁻¹ P_mᵀ R (node-level P ⊗ I₃, clamped fine DOF masked)."""
        xp = self.xp
        k = R.shape[1]
        Rm = R if overwrite else R.copy()
        Rm[self.fixed] = 0.0
        rc = (self._PnT @ Rm.reshape(-1, 3 * k)).reshape(-1, k)
        del Rm
        rc[self._fixed_c] = 0.0
        xc = self._co.solve(rc)
        del rc
        xc[self._fixed_c] = 0.0
        Y = xp.ascontiguousarray((self._Pn @ xc.reshape(-1, 3 * k)).reshape(-1, k))
        Y[self.fixed] = 0.0
        return Y

    def _residual(self, B, X):
        """B − K X in one new block."""
        T = self._apply_K(X)
        self.xp.subtract(B, T, out=T)
        return T

    def precondition(self, B):
        """Symmetric two-grid V(1,1) cycle on (n, k) residuals: x = S b, x += C(b − Kx), x += S(b − Kx)."""
        X = self._cheb(B)
        Y = self._coarse_correction(self._residual(B, X), overwrite=True)
        X += Y
        del Y
        Y = self._cheb(self._residual(B, X), overwrite=True)
        X += Y
        del Y
        X[self.fixed] = B[self.fixed]
        self._vcycles += B.shape[1]
        return X

    # -- solves ----------------------------------------------------------------------------------------------------
    def _reset_counters(self):
        self._vcycles = 0
        self._t_pcg = 0.0
        self.stats_solves = []                     # (columns, iterations, tol) per solve since the last factor
        self.stall_events = []                     # stagnation acceptances since the last factor (see _pcg)
        if hasattr(self, "stats"):
            self.stats.update(pcg_stall_accepted=0, pcg_stall_max_rel=0.0)

    def _pcg(self, B, X0=None, tol: float = 1e-10):
        """Column-wise PCG on (n, k) blocks (converged columns frozen); Polak-Ribière β. Raises when a column does not
        reach ‖r‖ ≤ tol ‖b‖ within maxiter iterations, or on an indefinite step. A column whose residual stagnates
        (reduction < 2× over `stall_window` iterations: rounding level of the operator) is accepted with a warning if
        its relative residual is ≤ `stall_accept`; it is then counted in `stall_events`."""
        xp = self.xp
        t_start = time.time()
        k = B.shape[1]
        nb = xp.linalg.norm(B, axis=0)
        if not bool(xp.isfinite(nb).all()):
            raise FloatingPointError("two-grid PCG: non-finite right-hand side")
        X = xp.zeros_like(B) if X0 is None else xp.array(X0, dtype=xp.float64).reshape(B.shape)
        R = B - self._apply_K(X) if X0 is not None else B.copy()
        tgt = tol * nb
        rn = xp.linalg.norm(R, axis=0)
        active = (rn > tgt) & (nb > 0)
        if X0 is not None:
            X[:, ~(nb > 0)] = 0.0
        its = 0
        if bool(active.any()):
            Z = self.precondition(R)
            P = Z.copy()
            rz = _coldot(xp, R, Z)
            hist = [rn]                                           # residual norms of the last stall_window + 1 iterations
            for its in range(1, self.maxiter + 1):
                AP = self._apply_K(P)
                pap = _coldot(xp, P, AP)
                if bool(xp.any(active & ((pap <= 0) | (rz <= 0)))):
                    raise RuntimeError("two-grid PCG: non-positive curvature (preconditioner not SPD?)")
                alpha = xp.where(active, rz / xp.where(pap > 0, pap, 1.0), 0.0)[None, :]
                _axpy(xp, X, P, alpha)
                _axpy(xp, R, AP, -alpha)
                del AP
                rn = xp.linalg.norm(R, axis=0)
                if not bool(xp.isfinite(rn).all()):
                    raise FloatingPointError(f"two-grid PCG: non-finite residual at iteration {its}")
                active &= rn > tgt
                hist.append(rn)
                if len(hist) > self.stall_window + 1:
                    hist.pop(0)
                    stalled = active & (rn > 0.5 * hist[0]) & (rn <= self.stall_accept * nb)
                    if bool(stalled.any()):
                        self._accept_stalled(stalled, rn, nb, its, tol)
                        active &= ~stalled
                if not bool(active.any()):
                    break
                rzo = _coldot(xp, R, Z)
                del Z
                Z = self.precondition(R)
                rz_new = _coldot(xp, R, Z)
                beta = xp.where(active, (rz_new - rzo) / xp.where(rz > 0, rz, 1.0), 0.0)
                _xpay(xp, P, Z, beta[None, :])
                rz = rz_new
            else:
                rel = float(xp.max(rn / xp.where(nb > 0, nb, 1.0)))
                raise RuntimeError(f"two-grid PCG did not converge in {self.maxiter} iterations "
                                   f"(rel. residual {rel:.2e}, tol {tol:.0e}, {k} columns)")
        self.stats_solves.append((k, its, tol))
        self._t_pcg += time.time() - t_start
        return X

    def _accept_stalled(self, stalled, rn, nb, its, tol):
        rel = float(self.xp.max(self.xp.where(stalled, rn / self.xp.where(nb > 0, nb, 1.0), 0.0)))
        n = int(stalled.sum())
        self.stall_events.append(dict(columns=n, iteration=its, rel_residual=rel, tol=tol))
        self.stats.update(pcg_stall_accepted=len(self.stall_events),
                          pcg_stall_max_rel=max(e["rel_residual"] for e in self.stall_events))
        print(f"fe_twogrid: WARNING: PCG stagnated at iteration {its} ({n} column(s)): accepting relative residual "
              f"{rel:.2e} (target {tol:.0e}, accept limit {self.stall_accept:.0e})", flush=True)

    def solve(self, rhs, x0=None, tol: float | None = None):
        """K x = rhs for (ndof,) or (ndof, k); clamped DOF of rhs are zeroed (x = 0 there). Columns are solved in
        blocks of `nrhs`."""
        xp = self.xp
        b = xp.array(rhs, dtype=xp.float64)
        one = b.ndim == 1
        if one:
            b = b[:, None]
        b[self.fixed] = 0.0
        t = self.tol if tol is None else float(tol)
        if x0 is not None:
            x0 = xp.asarray(x0, dtype=xp.float64).reshape(b.shape)
        X = xp.empty_like(b)
        for j0 in range(0, b.shape[1], self.nrhs):
            j1 = min(j0 + self.nrhs, b.shape[1])
            X[:, j0:j1] = self._pcg(xp.ascontiguousarray(b[:, j0:j1]),
                                    None if x0 is None else xp.ascontiguousarray(x0[:, j0:j1]), t)
        return X[:, 0] if one else X

    def solve_inner(self, rhs, eig_res: float = np.inf):
        """Krylov-expansion and residual-norm solves of the buckling eigen-solver, to min(krylov_tol, 0.1·res) (not
        below `tol`), res the eigen-solver's current dual-norm residual: a fixed inner tolerance would cap the
        attainable eigen-residual near it (inexact subspace iteration), a looser one early costs nothing."""
        return self.solve(rhs, tol=max(self.tol, min(self.krylov_tol, 0.1 * eig_res)))

    def primal(self) -> Primal:
        t0 = time.time()
        x0 = self._u_prev if (self.warm_start and self._u_prev is not None) else None
        u = self.solve(self.f, x0=x0)
        self._u_prev = u.copy() if self.warm_start else None
        C = float(self.f @ u)
        U = self.element_bilinear(u, u, solid=False) * 0.5
        return Primal(compliance=C, u=u, U_e=U, seconds=dict(solve=time.time() - t0, **{
            k: self.stats[k] for k in ("assemble_s", "factor_s") if k in self.stats}))

    def pcg_summary(self) -> dict:
        """Iterations per solve since the last factor (columns, iterations, tol)."""
        s = self.stats_solves
        full = [i for _, i, t in s if t <= self.tol * 1.0001]
        inner = [i for _, i, t in s if t > self.tol * 1.0001]
        return dict(n_solves=len(s), its_exact=full, its_krylov_mean=float(np.mean(inner)) if inner else 0.0,
                    its_krylov_max=int(max(inner)) if inner else 0, vcycles=int(self._vcycles),
                    pcg_s=float(self._t_pcg), stalled=len(self.stall_events),
                    stall_max_rel=max([e["rel_residual"] for e in self.stall_events], default=0.0))

    def free(self):
        """Release the coarse factor and the device arrays of this engine."""
        if getattr(self, "_co", None) is not None:
            self._co.free()
        self._co = None
        self._u_prev = None
        if self.xp is not np:
            for a in ("_Pn", "_PnT", "_pat", "_diag", "_dinv", "_s", "_dofs", "_parent", "_child", "_drop_c",
                      "_fixed_c", "_diag_fixed_c"):
                setattr(self, a, None)
            _release_device_memory(self.xp)
