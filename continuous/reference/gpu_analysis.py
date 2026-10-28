"""
gpu_analysis.py -- the GPU / engine layer of the elastic and buckling analysis (lowest module of the toolchain).

Sections (each a former module, kept as one headed block; `import gpu_analysis` imports neither jax nor initialises CUDA):
  former element_matrices.py  element stiffness tables (q1_gauss / q1 / q1_im)
  former gpu_CG.py            numba-CUDA matrix-free PCG, GMG preconditioner, fast engine
  former gpu_jax.py           JAX/XLA GMG-PCG engine (JAX imported lazily inside its functions)
  former engines.py           capability probe and engine selection; process-wide engine settings and records
"""
from __future__ import annotations


# ======================================================================================================================
# === former element_matrices.py ===
# ======================================================================================================================
"""
Host-side (numpy, float64) unit-E element stiffness matrices for the axis-aligned
box hex of the CG backend, shared by the GPU kernels (gpu_CG.py), the CPU mirror
(utils/cpu_backend.py) and the buckling assembly (analysis.py).

Element types (analysis.parameters.element_type):

    "q1_gauss"  trilinear hex, 2x2x2 Gauss, B recomputed per element in the kernels
                (the original path; never touches this module).
    "q1"        the same element as a precomputed 24x24 matrix  K = E * khat_q1(h, nu).
    "q1_im"     Wilson-Taylor incompatible-modes hex (QM6), 9 internal modes statically
                condensed:  Khat = K_uu - K_ua K_aa^-1 K_au   (docs/theory.md §13).

Conventions (identical to gpu_CG.py): local node order (xi, eta, zeta) corners
(-,-,-) (+,-,-) (+,+,-) (-,+,-) (-,-,+) (+,-,+) (+,+,+) (-,+,+); DOF order
[ux0, uy0, uz0, ux1, ...]; Voigt order (xx, yy, zz, yz, xz, xy) with ENGINEERING
shear; C = lam*[1 1 1 0 0 0]-block + mu*diag(2,2,2,1,1,1).

K_e = E_e * Khat(nu): the elements of one run share (hx, hy, hz) and nu; only E varies.
"""
import numpy as np

ELEMENT_TYPES = ("q1_gauss", "q1", "q1_im")

_XI = np.array([-1.0, 1.0, 1.0, -1.0, -1.0, 1.0, 1.0, -1.0])
_ETA = np.array([-1.0, -1.0, 1.0, 1.0, -1.0, -1.0, 1.0, 1.0])
_ZETA = np.array([-1.0, -1.0, -1.0, -1.0, 1.0, 1.0, 1.0, 1.0])
_G = 1.0 / np.sqrt(3.0)


def _gauss_points():
    return [(a, b, c) for a in (-_G, _G) for b in (-_G, _G) for c in (-_G, _G)]


def unit_C(nu):
    """6x6 Voigt elasticity matrix for E = 1."""
    lam = nu / ((1.0 + nu) * (1.0 - 2.0 * nu))
    mu = 1.0 / (2.0 * (1.0 + nu))
    C = np.zeros((6, 6))
    C[:3, :3] = lam
    C[0, 0] = C[1, 1] = C[2, 2] = lam + 2.0 * mu
    C[3, 3] = C[4, 4] = C[5, 5] = mu
    return C


def _B_compatible(h, xi, eta, zeta):
    """6x24 Voigt strain-displacement matrix of the trilinear hex at (xi, eta, zeta)."""
    hx, hy, hz = h
    dxi = 0.125 * _XI * (1 + _ETA * eta) * (1 + _ZETA * zeta) * (2.0 / hx)
    deta = 0.125 * _ETA * (1 + _XI * xi) * (1 + _ZETA * zeta) * (2.0 / hy)
    dzeta = 0.125 * _ZETA * (1 + _XI * xi) * (1 + _ETA * eta) * (2.0 / hz)
    B = np.zeros((6, 24))
    for a in range(8):
        c = 3 * a
        B[0, c] = dxi[a]; B[1, c + 1] = deta[a]; B[2, c + 2] = dzeta[a]
        B[3, c + 1] = dzeta[a]; B[3, c + 2] = deta[a]
        B[4, c] = dzeta[a]; B[4, c + 2] = dxi[a]
        B[5, c] = deta[a]; B[5, c + 1] = dxi[a]
    return B


def _B_incompatible(h, xi, eta, zeta):
    """
    6x9 Voigt strain matrix of the incompatible modes u_j = sum_k (1 - xi_k^2) alpha_kj
    (column 3*k + j).  d u_j / d x_i = delta_ik * (-4 xi_i / h_i) alpha_ij, so the mode (k, j)
    contributes G_k = -4 xi_k / h_k to d_k u_j.

    Taylor's correction (subtract the element mean of B_alpha, equivalently evaluate the
    mode derivatives with the centre Jacobian) is the identity here: the Jacobian of a box is
    constant and the mean of xi_k over the element is zero, so it changes nothing; the
    condition it enforces (int B_alpha dV = 0, i.e. the patch test) holds exactly.
    """
    hx, hy, hz = h
    G = np.array([-4.0 * xi / hx, -4.0 * eta / hy, -4.0 * zeta / hz])
    Bi = np.zeros((6, 9))
    for k in range(3):
        for j in range(3):
            col = 3 * k + j
            if k == j:
                Bi[j, col] = G[k]                       # normal strain eps_kk
            else:
                # engineering shear row for the pair (k, j): yz=3, xz=4, xy=5
                row = {(1, 2): 3, (2, 1): 3, (0, 2): 4, (2, 0): 4, (0, 1): 5, (1, 0): 5}[(k, j)]
                Bi[row, col] = G[k]
    return Bi


def khat_q1(h, nu):
    """Unit-E 24x24 stiffness of the trilinear box hex, 2x2x2 Gauss (float64)."""
    C = unit_C(nu)
    detJ = h[0] * h[1] * h[2] / 8.0
    K = np.zeros((24, 24))
    for g in _gauss_points():
        B = _B_compatible(h, *g)
        K += B.T @ C @ B * detJ
    return K


def qm6_blocks(h, nu):
    """(K_uu, K_ua, K_aa) of the QM6 element, unit E, 2x2x2 Gauss."""
    C = unit_C(nu)
    detJ = h[0] * h[1] * h[2] / 8.0
    Kuu = np.zeros((24, 24)); Kua = np.zeros((24, 9)); Kaa = np.zeros((9, 9))
    for g in _gauss_points():
        B = _B_compatible(h, *g)
        Bi = _B_incompatible(h, *g)
        Kuu += B.T @ C @ B * detJ
        Kua += B.T @ C @ Bi * detJ
        Kaa += Bi.T @ C @ Bi * detJ
    return Kuu, Kua, Kaa


def khat_qm6(h, nu):
    """Unit-E 24x24 condensed QM6 stiffness: K_uu - K_ua K_aa^-1 K_au (float64)."""
    Kuu, Kua, Kaa = qm6_blocks(h, nu)
    K = Kuu - Kua @ np.linalg.solve(Kaa, Kua.T)
    return 0.5 * (K + K.T)


def element_khat(element_type, h, nu):
    """Unit-E 24x24 float64 matrix for element_type in {'q1', 'q1_im'}."""
    h = (float(h[0]), float(h[1]), float(h[2]))
    if element_type == "q1":
        return khat_q1(h, float(nu))
    if element_type == "q1_im":
        return khat_qm6(h, float(nu))
    raise ValueError(f"element_khat: element_type must be 'q1' or 'q1_im', got {element_type!r}")


def resolve_nu(nu_elem, E_elem, nu=None, tol=1e-4):
    """
    The single Poisson ratio of a run. If `nu` is None it is taken from the elements with
    E > 0 (the density-scaled lam/mu pair keeps nu constant; void elements with E = 0 carry
    a meaningless nu). Raises if the stiff elements disagree (the matrix kernels hold one K).
    """
    nu_elem = np.asarray(nu_elem, dtype=np.float64)
    E_elem = np.asarray(E_elem)
    live = nu_elem[E_elem > 0.0]
    if nu is None:
        if live.size == 0:
            raise ValueError("resolve_nu: no element with E > 0 and no nu given")
        nu = float(np.median(live))
    if live.size and np.max(np.abs(live - nu)) > tol:
        raise ValueError(f"element_type != 'q1_gauss' needs a single Poisson ratio: "
                         f"elements span nu = {live.min():.6f}..{live.max():.6f}, expected {nu:.6f}")
    return float(nu)


# ======================================================================================================================
# === former gpu_CG.py ===
# ======================================================================================================================
import warnings
import numpy as np
from math import sqrt
from numba import cuda, float32, float64, int32
from numba.core.errors import NumbaPerformanceWarning

warnings.filterwarnings("ignore", category=NumbaPerformanceWarning)


THREADS_PER_BLOCK = 256

ELEMS_PER_BLOCK_I = 4
ELEMS_PER_BLOCK_J = 4
ELEMS_PER_BLOCK_K = 4

_B_ROWS = 6
_B_COLS = 24
_PE_SIZE = 24
_B_SIZE = _B_ROWS * _B_COLS

_PE_SIZE2 = _PE_SIZE * _PE_SIZE
_6PE_SIZE = _PE_SIZE * 6

# local arrays for reference hex corners (row-major)
xi_arr = np.array([-1.0,  1.0,  1.0, -1.0, -1.0,  1.0,  1.0, -1.0], dtype=np.float32)
eta_arr = np.array([-1.0, -1.0,  1.0,  1.0, -1.0, -1.0,  1.0,  1.0], dtype=np.float32)
zeta_arr = np.array([-1.0, -1.0, -1.0, -1.0,  1.0,  1.0,  1.0,  1.0], dtype=np.float32)

@cuda.jit(device=True)
def _compute_C_voigt(E, nu, C_out):
    denom = (1.0 + nu) * (1.0 - 2.0 * nu)
    if denom == 0.0:
        denom = 1e-12
    lam = E * nu / denom
    mu = E / (2.0 * (1.0 + nu))

    # zero
    for r in range(6):
        for c in range(6):
            C_out[r*6 + c] = 0.0

    # normal terms
    C_out[0*6 + 0] = lam + 2.0*mu
    C_out[1*6 + 1] = lam + 2.0*mu
    C_out[2*6 + 2] = lam + 2.0*mu

    # off-diagonal normals
    C_out[0*6 + 1] = lam; C_out[0*6 + 2] = lam
    C_out[1*6 + 0] = lam; C_out[1*6 + 2] = lam
    C_out[2*6 + 0] = lam; C_out[2*6 + 1] = lam

    # shear terms
    C_out[3*6 + 3] = mu
    C_out[4*6 + 4] = mu
    C_out[5*6 + 5] = mu

@cuda.jit(device=True)
def _compute_B_at_gauss(hx, hy, hz, gp_xi, gp_eta, gp_zeta, B_out):
    invJ_x = 2.0 / hx
    invJ_y = 2.0 / hy
    invJ_z = 2.0 / hz

    # zero
    for r in range(_B_ROWS):
        for c in range(_B_COLS):
            B_out[r * _B_COLS + c] = 0.0

    for a in range(8):
        xi_i = xi_arr[a]; eta_i = eta_arr[a]; zeta_i = zeta_arr[a]

        dN_dxi = 0.125 * xi_i * (1.0 + eta_i * gp_eta) * (1.0 + zeta_i * gp_zeta)
        dN_deta = 0.125 * eta_i * (1.0 + xi_i * gp_xi) * (1.0 + zeta_i * gp_zeta)
        dN_dzeta = 0.125 * zeta_i * (1.0 + xi_i * gp_xi) * (1.0 + eta_i * gp_eta)

        dN_dx = dN_dxi * invJ_x
        dN_dy = dN_deta * invJ_y
        dN_dz = dN_dzeta * invJ_z

        col_base = 3 * a
        B_out[0 * _B_COLS + (col_base + 0)] = dN_dx
        B_out[1 * _B_COLS + (col_base + 1)] = dN_dy
        B_out[2 * _B_COLS + (col_base + 2)] = dN_dz

        B_out[3 * _B_COLS + (col_base + 1)] = dN_dz
        B_out[3 * _B_COLS + (col_base + 2)] = dN_dy

        B_out[4 * _B_COLS + (col_base + 0)] = dN_dz
        B_out[4 * _B_COLS + (col_base + 2)] = dN_dx

        B_out[5 * _B_COLS + (col_base + 0)] = dN_dy
        B_out[5 * _B_COLS + (col_base + 1)] = dN_dx

@cuda.jit(fastmath=True)
def _element_action_kernel(nx, ny, nz, hx, hy, hz, E_elem, nu_elem, p_global, Ap_global, elems_per_block_i, elems_per_block_j, elems_per_block_k):
    """
    Element-wise elasticity action with per-block accumulation to reduce atomics.
    Each thread still computes its element's f_e but accumulation into Ap_global is done
    by first collecting per-block unique nodes & contributions into shared memory and then 
    performing atomic adds once per unique node per block.
    """
    # thread / block ids
    tid = cuda.threadIdx.x
    bidx = cuda.blockIdx.x
    bdim = cuda.blockDim.x

    nx_elems = nx - 1
    ny_elems = ny - 1
    nz_elems = nz - 1
    Ne = nx_elems * ny_elems * nz_elems
    # total elements partitioning in linear blocks: compute how many blocks we launched for grid
    # here each block handles contiguous chunk of elements sized about bdim threads,
    # and the shared memory per block will gather unique nodes from those elements.
    e_global = bidx * bdim + tid
    if e_global >= Ne:
        return

    # compute element indices (ei,ej,ek)
    ek = e_global % nz_elems
    ej = (e_global // nz_elems) % ny_elems
    ei = e_global // (ny_elems * nz_elems)

    # node ids for the element
    n0 = (ei    ) * (ny * nz) + (ej    ) * nz + (ek    )
    n1 = (ei + 1) * (ny * nz) + (ej    ) * nz + (ek    )
    n2 = (ei + 1) * (ny * nz) + (ej + 1) * nz + (ek    )
    n3 = (ei    ) * (ny * nz) + (ej + 1) * nz + (ek    )
    n4 = (ei    ) * (ny * nz) + (ej    ) * nz + (ek + 1)
    n5 = (ei + 1) * (ny * nz) + (ej    ) * nz + (ek + 1)
    n6 = (ei + 1) * (ny * nz) + (ej + 1) * nz + (ek + 1)
    n7 = (ei    ) * (ny * nz) + (ej + 1) * nz + (ek + 1)

    # gather p_e (24)
    p_e = cuda.local.array(_PE_SIZE, float32)
    nodes_local = cuda.local.array(8, int32)
    nodes_local[0]=n0; nodes_local[1]=n1; nodes_local[2]=n2; nodes_local[3]=n3
    nodes_local[4]=n4; nodes_local[5]=n5; nodes_local[6]=n6; nodes_local[7]=n7

    for a in range(8):
        nid = nodes_local[a]
        base = 3*nid
        p_e[3*a + 0] = p_global[base + 0]
        p_e[3*a + 1] = p_global[base + 1]
        p_e[3*a + 2] = p_global[base + 2]

    # compute f_e
    f_e = cuda.local.array(_PE_SIZE, float32)
    for i in range(_PE_SIZE):
        f_e[i] = 0.0

    Cloc = cuda.local.array(36, float32)
    _compute_C_voigt(E_elem[e_global], nu_elem[e_global], Cloc)

    g = 0.5773502691896257
    gauss_vals0 = (-g, g)
    for gi in range(2):
        gp_xi = gauss_vals0[gi]
        for gj in range(2):
            gp_eta = gauss_vals0[gj]
            for gk in range(2):
                gp_zeta = gauss_vals0[gk]
                B = cuda.local.array(_B_SIZE, float32)
                _compute_B_at_gauss(hx, hy, hz, gp_xi, gp_eta, gp_zeta, B)
                # eps = B*p_e
                eps0 = 0.0; eps1 = 0.0; eps2 = 0.0; eps3 = 0.0; eps4 = 0.0; eps5 = 0.0
                # unroll for speed (6 rows)
                for jj in range(_B_COLS):
                    v = p_e[jj]
                    eps0 += B[0*_B_COLS + jj] * v
                    eps1 += B[1*_B_COLS + jj] * v
                    eps2 += B[2*_B_COLS + jj] * v
                    eps3 += B[3*_B_COLS + jj] * v
                    eps4 += B[4*_B_COLS + jj] * v
                    eps5 += B[5*_B_COLS + jj] * v

                sigma0 = Cloc[0*6 + 0]*eps0 + Cloc[0*6 + 1]*eps1 + Cloc[0*6 + 2]*eps2 + Cloc[0*6 + 3]*eps3 + Cloc[0*6 + 4]*eps4 + Cloc[0*6 + 5]*eps5
                sigma1 = Cloc[1*6 + 0]*eps0 + Cloc[1*6 + 1]*eps1 + Cloc[1*6 + 2]*eps2 + Cloc[1*6 + 3]*eps3 + Cloc[1*6 + 4]*eps4 + Cloc[1*6 + 5]*eps5
                sigma2 = Cloc[2*6 + 0]*eps0 + Cloc[2*6 + 1]*eps1 + Cloc[2*6 + 2]*eps2 + Cloc[2*6 + 3]*eps3 + Cloc[2*6 + 4]*eps4 + Cloc[2*6 + 5]*eps5
                sigma3 = Cloc[3*6 + 0]*eps0 + Cloc[3*6 + 1]*eps1 + Cloc[3*6 + 2]*eps2 + Cloc[3*6 + 3]*eps3 + Cloc[3*6 + 4]*eps4 + Cloc[3*6 + 5]*eps5
                sigma4 = Cloc[4*6 + 0]*eps0 + Cloc[4*6 + 1]*eps1 + Cloc[4*6 + 2]*eps2 + Cloc[4*6 + 3]*eps3 + Cloc[4*6 + 4]*eps4 + Cloc[4*6 + 5]*eps5
                sigma5 = Cloc[5*6 + 0]*eps0 + Cloc[5*6 + 1]*eps1 + Cloc[5*6 + 2]*eps2 + Cloc[5*6 + 3]*eps3 + Cloc[5*6 + 4]*eps4 + Cloc[5*6 + 5]*eps5

                # Bt_sigma = B^T * sigma
                detJ = (hx * hy * hz) * 0.125
                for col in range(_B_COLS):
                    val = B[0*_B_COLS + col] * sigma0 + B[1*_B_COLS + col] * sigma1 + B[2*_B_COLS + col] * sigma2 \
                          + B[3*_B_COLS + col] * sigma3 + B[4*_B_COLS + col] * sigma4 + B[5*_B_COLS + col] * sigma5
                    f_e[col] += val * detJ

    # --------------------------
    # Per-block accumulation:
    # We'll have shared arrays: shared_node_ids[], shared_vals[].
    # Each thread will insert its 8 node ids into shared_node_ids and its contributions into shared_vals.
    # We'll do a simple parallel insertion: each thread tries to find node id in shared_node_ids; if not present, it appends.
    # After all threads inserted, we'll perform reduction over contributions per shared node and then atomic-add once to global Ap.
    # --------------------------
    # Shared memory declarations (dynamic)
    smem_node_ids = cuda.shared.array(shape=0, dtype=int32)  # dynamic: we'll index via pointer arithmetic; not directly used
    # Numba doesn't support dynamic shared arrays via python variable; instead we'll emulate by allocating fixed-size shared arrays
    # sized by maximum possible unique nodes per block (compute at runtime on host and set kernel launch accordingly).
    # For simplicity inside this kernel we assume the host has launched with enough static shared memory sized to MAX_UNIQUE.

    # Because Numba limits dynamic shared creation via python-level, a simpler approach is:
    # perform atomic adds but group elements per block into contiguous indices minimizing collisions.
    # Given complexity and portability, fallback to atomic adds here for correctness, but still keep optimized internal multiplies.
    # (This keeps the rest of the solver fully GPU-resident and reduces host traffic.)
    # So final step: atomic add per node (one atomic per DOF per element) -- still correct.
    for a in range(8):
        nid = nodes_local[a]
        base = 3*nid
        cuda.atomic.add(Ap_global, base + 0, f_e[3*a + 0])
        cuda.atomic.add(Ap_global, base + 1, f_e[3*a + 1])
        cuda.atomic.add(Ap_global, base + 2, f_e[3*a + 2])

@cuda.jit(fastmath=True)
def _element_energy_kernel(nx, ny, nz, hx, hy, hz, E_elem, nu_elem, u_global, energy_out):
    """
    Computes total Strain Energy (Joules) per element.
    To get Density (J/m^3), divide the result by element volume (hx*hy*hz).
    """
    tid = cuda.threadIdx.x
    bidx = cuda.blockIdx.x
    bdim = cuda.blockDim.x

    nx_elems = nx - 1
    ny_elems = ny - 1
    nz_elems = nz - 1
    Ne = nx_elems * ny_elems * nz_elems

    e_global = bidx * bdim + tid
    if e_global >= Ne:
        return

    # compute element indices
    ek = e_global % nz_elems
    ej = (e_global // nz_elems) % ny_elems
    ei = e_global // (ny_elems * nz_elems)

    # node ids
    n0 = (ei    ) * (ny * nz) + (ej    ) * nz + (ek    )
    n1 = (ei + 1) * (ny * nz) + (ej    ) * nz + (ek    )
    n2 = (ei + 1) * (ny * nz) + (ej + 1) * nz + (ek    )
    n3 = (ei    ) * (ny * nz) + (ej + 1) * nz + (ek    )
    n4 = (ei    ) * (ny * nz) + (ej    ) * nz + (ek + 1)
    n5 = (ei + 1) * (ny * nz) + (ej    ) * nz + (ek + 1)
    n6 = (ei + 1) * (ny * nz) + (ej + 1) * nz + (ek + 1)
    n7 = (ei    ) * (ny * nz) + (ej + 1) * nz + (ek + 1)

    # gather u_element (24)
    u_e = cuda.local.array(_PE_SIZE, float32)
    nodes_local = cuda.local.array(8, int32)
    nodes_local[0]=n0; nodes_local[1]=n1; nodes_local[2]=n2; nodes_local[3]=n3
    nodes_local[4]=n4; nodes_local[5]=n5; nodes_local[6]=n6; nodes_local[7]=n7

    for a in range(8):
        nid = nodes_local[a]
        base = 3*nid
        u_e[3*a + 0] = u_global[base + 0]
        u_e[3*a + 1] = u_global[base + 1]
        u_e[3*a + 2] = u_global[base + 2]

    # Material matrix C
    Cloc = cuda.local.array(36, float32)
    _compute_C_voigt(E_elem[e_global], nu_elem[e_global], Cloc)

    total_energy = 0.0
    
    # Gauss Integration
    g = 0.5773502691896257
    gauss_vals0 = (-g, g)
    
    # Precompute determinant (constant for rectangular hex)
    detJ = (hx * hy * hz) * 0.125

    for gi in range(2):
        gp_xi = gauss_vals0[gi]
        for gj in range(2):
            gp_eta = gauss_vals0[gj]
            for gk in range(2):
                gp_zeta = gauss_vals0[gk]
                
                # Compute B matrix
                B = cuda.local.array(_B_SIZE, float32)
                _compute_B_at_gauss(hx, hy, hz, gp_xi, gp_eta, gp_zeta, B)

                # Compute Strain: eps = B * u_e
                eps0 = 0.0; eps1 = 0.0; eps2 = 0.0; eps3 = 0.0; eps4 = 0.0; eps5 = 0.0
                for jj in range(_B_COLS):
                    v = u_e[jj]
                    eps0 += B[0*_B_COLS + jj] * v
                    eps1 += B[1*_B_COLS + jj] * v
                    eps2 += B[2*_B_COLS + jj] * v
                    eps3 += B[3*_B_COLS + jj] * v
                    eps4 += B[4*_B_COLS + jj] * v
                    eps5 += B[5*_B_COLS + jj] * v

                # Compute Stress: sigma = C * eps
                sigma0 = Cloc[0*6 + 0]*eps0 + Cloc[0*6 + 1]*eps1 + Cloc[0*6 + 2]*eps2 + Cloc[0*6 + 3]*eps3 + Cloc[0*6 + 4]*eps4 + Cloc[0*6 + 5]*eps5
                sigma1 = Cloc[1*6 + 0]*eps0 + Cloc[1*6 + 1]*eps1 + Cloc[1*6 + 2]*eps2 + Cloc[1*6 + 3]*eps3 + Cloc[1*6 + 4]*eps4 + Cloc[1*6 + 5]*eps5
                sigma2 = Cloc[2*6 + 0]*eps0 + Cloc[2*6 + 1]*eps1 + Cloc[2*6 + 2]*eps2 + Cloc[2*6 + 3]*eps3 + Cloc[2*6 + 4]*eps4 + Cloc[2*6 + 5]*eps5
                sigma3 = Cloc[3*6 + 0]*eps0 + Cloc[3*6 + 1]*eps1 + Cloc[3*6 + 2]*eps2 + Cloc[3*6 + 3]*eps3 + Cloc[3*6 + 4]*eps4 + Cloc[3*6 + 5]*eps5
                sigma4 = Cloc[4*6 + 0]*eps0 + Cloc[4*6 + 1]*eps1 + Cloc[4*6 + 2]*eps2 + Cloc[4*6 + 3]*eps3 + Cloc[4*6 + 4]*eps4 + Cloc[4*6 + 5]*eps5
                sigma5 = Cloc[5*6 + 0]*eps0 + Cloc[5*6 + 1]*eps1 + Cloc[5*6 + 2]*eps2 + Cloc[5*6 + 3]*eps3 + Cloc[5*6 + 4]*eps4 + Cloc[5*6 + 5]*eps5

                # Energy Density w = 0.5 * (sigma dot eps)
                w = 0.5 * (sigma0*eps0 + sigma1*eps1 + sigma2*eps2 + sigma3*eps3 + sigma4*eps4 + sigma5*eps5)

                # Accumulate integral
                total_energy += w * detJ

    energy_out[e_global] = total_energy

# =========================================================================
# --- PRECOMPUTED-MATRIX ELEMENT KERNELS (element_type "q1" and "q1_im") ---
#
# K_e = E_e * Khat, with Khat the unit-E 24x24 element matrix built on the host
# (element_matrices.py, float64 -> float32). The elements share (hx, hy, hz) and
# nu; only E varies. Node/DOF order and the atomic accumulation are those of the
# Gauss-point kernels above, which remain the default ("q1_gauss") path.
# Khat is staged into shared memory once per block (all threads of a warp read
# the same address, so the reads broadcast). Threads past Ne must not return
# before the barrier, hence the `valid` flag instead of an early return.
# =========================================================================
@cuda.jit(fastmath=True)
def _element_action_kernel_mat(nx, ny, nz, E_elem, Khat, p_global, Ap_global):
    sK = cuda.shared.array(_PE_SIZE2, float32)
    tid = cuda.threadIdx.x
    for i in range(tid, _PE_SIZE2, cuda.blockDim.x):
        sK[i] = Khat[i]
    cuda.syncthreads()

    ny_elems = ny - 1
    nz_elems = nz - 1
    Ne = (nx - 1) * ny_elems * nz_elems
    e_global = cuda.blockIdx.x * cuda.blockDim.x + tid
    if e_global >= Ne:
        return

    ek = e_global % nz_elems
    ej = (e_global // nz_elems) % ny_elems
    ei = e_global // (ny_elems * nz_elems)

    nodes_local = cuda.local.array(8, int32)
    nodes_local[0] = (ei    ) * (ny * nz) + (ej    ) * nz + (ek    )
    nodes_local[1] = (ei + 1) * (ny * nz) + (ej    ) * nz + (ek    )
    nodes_local[2] = (ei + 1) * (ny * nz) + (ej + 1) * nz + (ek    )
    nodes_local[3] = (ei    ) * (ny * nz) + (ej + 1) * nz + (ek    )
    nodes_local[4] = (ei    ) * (ny * nz) + (ej    ) * nz + (ek + 1)
    nodes_local[5] = (ei + 1) * (ny * nz) + (ej    ) * nz + (ek + 1)
    nodes_local[6] = (ei + 1) * (ny * nz) + (ej + 1) * nz + (ek + 1)
    nodes_local[7] = (ei    ) * (ny * nz) + (ej + 1) * nz + (ek + 1)

    p_e = cuda.local.array(_PE_SIZE, float32)
    for a in range(8):
        base = 3 * nodes_local[a]
        p_e[3*a + 0] = p_global[base + 0]
        p_e[3*a + 1] = p_global[base + 1]
        p_e[3*a + 2] = p_global[base + 2]

    E = E_elem[e_global]
    for a in range(8):
        base = 3 * nodes_local[a]
        for r in range(3):
            row = (3 * a + r) * _PE_SIZE
            s = 0.0
            for j in range(_PE_SIZE):
                s += sK[row + j] * p_e[j]
            cuda.atomic.add(Ap_global, base + r, E * s)


@cuda.jit(fastmath=True)
def _element_energy_kernel_mat(nx, ny, nz, E_elem, Khat, u_global, energy_out):
    """0.5 * E_e * u_e^T Khat u_e per element (Joules): consistent with the action kernel."""
    sK = cuda.shared.array(_PE_SIZE2, float32)
    tid = cuda.threadIdx.x
    for i in range(tid, _PE_SIZE2, cuda.blockDim.x):
        sK[i] = Khat[i]
    cuda.syncthreads()

    ny_elems = ny - 1
    nz_elems = nz - 1
    Ne = (nx - 1) * ny_elems * nz_elems
    e_global = cuda.blockIdx.x * cuda.blockDim.x + tid
    if e_global >= Ne:
        return

    ek = e_global % nz_elems
    ej = (e_global // nz_elems) % ny_elems
    ei = e_global // (ny_elems * nz_elems)

    nodes_local = cuda.local.array(8, int32)
    nodes_local[0] = (ei    ) * (ny * nz) + (ej    ) * nz + (ek    )
    nodes_local[1] = (ei + 1) * (ny * nz) + (ej    ) * nz + (ek    )
    nodes_local[2] = (ei + 1) * (ny * nz) + (ej + 1) * nz + (ek    )
    nodes_local[3] = (ei    ) * (ny * nz) + (ej + 1) * nz + (ek    )
    nodes_local[4] = (ei    ) * (ny * nz) + (ej    ) * nz + (ek + 1)
    nodes_local[5] = (ei + 1) * (ny * nz) + (ej    ) * nz + (ek + 1)
    nodes_local[6] = (ei + 1) * (ny * nz) + (ej + 1) * nz + (ek + 1)
    nodes_local[7] = (ei    ) * (ny * nz) + (ej + 1) * nz + (ek + 1)

    u_e = cuda.local.array(_PE_SIZE, float32)
    for a in range(8):
        base = 3 * nodes_local[a]
        u_e[3*a + 0] = u_global[base + 0]
        u_e[3*a + 1] = u_global[base + 1]
        u_e[3*a + 2] = u_global[base + 2]

    quad = 0.0
    for i in range(_PE_SIZE):
        row = i * _PE_SIZE
        s = 0.0
        for j in range(_PE_SIZE):
            s += sK[row + j] * u_e[j]
        quad += u_e[i] * s
    energy_out[e_global] = 0.5 * E_elem[e_global] * quad


@cuda.jit(fastmath=True)
def _element_diag_kernel_mat(nx, ny, nz, E_elem, Khat, diag_blocks):
    """Per-node 3x3 diagonal blocks of K for the block-Jacobi preconditioner from E_e * Khat."""
    e = cuda.grid(1)
    ny_elems = ny - 1
    nz_elems = nz - 1
    Ne = (nx - 1) * ny_elems * nz_elems
    if e >= Ne:
        return

    ek = e % nz_elems
    ej = (e // nz_elems) % ny_elems
    ei = e // (ny_elems * nz_elems)

    nodes = cuda.local.array(8, int32)
    nodes[0] = (ei    ) * (ny * nz) + (ej    ) * nz + (ek    )
    nodes[1] = (ei + 1) * (ny * nz) + (ej    ) * nz + (ek    )
    nodes[2] = (ei + 1) * (ny * nz) + (ej + 1) * nz + (ek    )
    nodes[3] = (ei    ) * (ny * nz) + (ej + 1) * nz + (ek    )
    nodes[4] = (ei    ) * (ny * nz) + (ej    ) * nz + (ek + 1)
    nodes[5] = (ei + 1) * (ny * nz) + (ej    ) * nz + (ek + 1)
    nodes[6] = (ei + 1) * (ny * nz) + (ej + 1) * nz + (ek + 1)
    nodes[7] = (ei    ) * (ny * nz) + (ej + 1) * nz + (ek + 1)

    E = E_elem[e]
    for a in range(8):
        base = nodes[a] * 9
        for rr in range(3):
            for cc in range(3):
                cuda.atomic.add(diag_blocks, base + rr*3 + cc,
                                E * Khat[(3*a + rr) * _PE_SIZE + (3*a + cc)])

@cuda.jit(fastmath=True)
def _vec_dot_block_kernel(a, b, block_sums):
    """
    Each block computes partial dot product of vectors a and b and writes sum to block_sums[blockIdx.x].
    """
    sdata = cuda.shared.array(THREADS_PER_BLOCK, float32)
    tid = cuda.threadIdx.x
    i = cuda.blockIdx.x * cuda.blockDim.x + tid
    n = a.size

    tmp = 0.0
    if i < n:
        tmp = a[i] * b[i]
    sdata[tid] = tmp
    cuda.syncthreads()

    # reduce in shared mem
    stride = cuda.blockDim.x // 2
    while stride > 0:
        if tid < stride and (tid + stride) < cuda.blockDim.x:
            sdata[tid] += sdata[tid + stride]
        cuda.syncthreads()
        stride //= 2

    if tid == 0:
        block_sums[cuda.blockIdx.x] = sdata[0]

@cuda.jit(fastmath=True)
def _vec_dot_final_kernel(block_sums, out_scalar):
    """
    Reduce block_sums to one scalar and store at out_scalar[0].
    """
    tid = cuda.threadIdx.x
    n = block_sums.size
    sdata = cuda.shared.array(THREADS_PER_BLOCK, float32)

    idx = tid
    s = 0.0
    while idx < n:
        s += block_sums[idx]
        idx += cuda.blockDim.x
    sdata[tid] = s
    cuda.syncthreads()

    stride = cuda.blockDim.x // 2
    while stride > 0:
        if tid < stride:
            sdata[tid] += sdata[tid + stride]
        cuda.syncthreads()
        stride //= 2

    if tid == 0:
        out_scalar[0] = sdata[0]

@cuda.jit(fastmath=True)
def _vec_axpby_kernel(alpha, x, y):  # x = x + alpha * y
    i = cuda.blockIdx.x * cuda.blockDim.x + cuda.threadIdx.x
    if i >= x.size:
        return
    x[i] = x[i] + alpha * y[i]

@cuda.jit(fastmath=True)
def _vec_axpy_kernel(y, alpha, x):  # y = y - alpha * x  (we'll use with alpha negative or adjust)
    i = cuda.blockIdx.x * cuda.blockDim.x + cuda.threadIdx.x
    if i >= x.size:
        return
    y[i] = y[i] - alpha * x[i]

@cuda.jit(fastmath=True)
def _vec_copy_kernel(dest, src):
    i = cuda.blockIdx.x * cuda.blockDim.x + cuda.threadIdx.x
    if i >= dest.size:
        return
    dest[i] = src[i]

@cuda.jit(fastmath=True)
def _vec_scaled_add_kernel(dest, a, src, b, src2):
    # dest = a*src + b*src2
    i = cuda.blockIdx.x * cuda.blockDim.x + cuda.threadIdx.x
    if i >= dest.size:
        return
    dest[i] = a * src[i] + b * src2[i]

@cuda.jit(fastmath=True)
def _element_diag_kernel(nx, ny, nz, hx, hy, hz, E_elem, nu_elem, diag_blocks):
    # Similar to previous, but only atomic-add diagonal 3x3 per node
    e = cuda.grid(1)
    nx_elems = nx - 1; ny_elems = ny - 1; nz_elems = nz - 1
    Ne = nx_elems * ny_elems * nz_elems
    if e >= Ne:
        return

    ek = e % nz_elems
    ej = (e // nz_elems) % ny_elems
    ei = e // (ny_elems * nz_elems)

    n0 = (ei    ) * (ny * nz) + (ej    ) * nz + (ek    )
    n1 = (ei + 1) * (ny * nz) + (ej    ) * nz + (ek    )
    n2 = (ei + 1) * (ny * nz) + (ej + 1) * nz + (ek    )
    n3 = (ei    ) * (ny * nz) + (ej + 1) * nz + (ek    )
    n4 = (ei    ) * (ny * nz) + (ej    ) * nz + (ek + 1)
    n5 = (ei + 1) * (ny * nz) + (ej    ) * nz + (ek + 1)
    n6 = (ei + 1) * (ny * nz) + (ej + 1) * nz + (ek + 1)
    n7 = (ei    ) * (ny * nz) + (ej + 1) * nz + (ek + 1)

    nodes = cuda.local.array(8, int32)
    nodes[0]=n0; nodes[1]=n1; nodes[2]=n2; nodes[3]=n3
    nodes[4]=n4; nodes[5]=n5; nodes[6]=n6; nodes[7]=n7

    Cloc = cuda.local.array(36, float32)
    _compute_C_voigt(E_elem[e], nu_elem[e], Cloc)

    # compute Ke diag parts only via B^T C B -> extract diag blocks
    Ke_local = cuda.local.array((_PE_SIZE2,), float32)
    for i in range(_PE_SIZE * _PE_SIZE):
        Ke_local[i] = 0.0

    g = 0.5773502691896257
    for gi in range(2):
        gp_xi = -g if gi==0 else g
        for gj in range(2):
            gp_eta = -g if gj==0 else g
            for gk in range(2):
                gp_zeta = -g if gk==0 else gk and g or -g
                # compute B
                B = cuda.local.array((_B_SIZE,), float32)
                _compute_B_at_gauss(hx, hy, hz, gp_xi, gp_eta, gp_zeta, B)
                # BtC = B^T * C
                BtC = cuda.local.array((_6PE_SIZE,), float32)
                for row in range(_B_COLS):
                    for col in range(6):
                        s = 0.0
                        for ii in range(6):
                            s += B[ii * _B_COLS + row] * Cloc[ii * 6 + col]
                        BtC[row * 6 + col] = s
                # Ke += BtC * B * detJ
                detJ = (hx * hy * hz) * 0.125
                for r in range(_B_COLS):
                    for c in range(_B_COLS):
                        s = 0.0
                        for k2 in range(6):
                            s += BtC[r * 6 + k2] * B[k2 * _B_COLS + c]
                        Ke_local[r * _B_COLS + c] += s * detJ
    # add diagonal blocks
    for a in range(8):
        nid = nodes[a]
        base = nid * 9
        for rr in range(3):
            for cc in range(3):
                val = Ke_local[(3*a + rr) * _B_COLS + (3*a + cc)]
                cuda.atomic.add(diag_blocks, base + rr*3 + cc, val)

@cuda.jit(fastmath=True)
def _zero_array_kernel(arr):
    """
    Sets all elements of a 1D device array to 0.0
    """
    i = cuda.grid(1)
    if i < arr.size:
        arr[i] = 0.0

@cuda.jit(fastmath=True)
def _invert_diag_blocks_kernel(Nnodes, diag_blocks, Minv_blocks, fixed_mask):
    nid = cuda.grid(1)
    if nid >= Nnodes:
        return
        
    # Uncompressed mask is length Ndofs
    fix0 = fixed_mask[nid*3 + 0]
    fix1 = fixed_mask[nid*3 + 1]
    fix2 = fixed_mask[nid*3 + 2]
    
    if fix0 == 1 and fix1 == 1 and fix2 == 1:
        for i in range(9):
            Minv_blocks[nid * 9 + i] = 0.0
        return

    a11 = diag_blocks[nid*9 + 0]; a12 = diag_blocks[nid*9 + 1]; a13 = diag_blocks[nid*9 + 2]
    a21 = diag_blocks[nid*9 + 3]; a22 = diag_blocks[nid*9 + 4]; a23 = diag_blocks[nid*9 + 5]
    a31 = diag_blocks[nid*9 + 6]; a32 = diag_blocks[nid*9 + 7]; a33 = diag_blocks[nid*9 + 8]

    # Zero out the off-diagonals and set diagonal to 1.0 for partially fixed DOFs
    if fix0 == 1:
        a11 = 1.0; a12 = 0.0; a13 = 0.0
        a21 = 0.0; a31 = 0.0
    if fix1 == 1:
        a22 = 1.0; a12 = 0.0; a23 = 0.0
        a21 = 0.0; a32 = 0.0
    if fix2 == 1:
        a33 = 1.0; a13 = 0.0; a23 = 0.0
        a31 = 0.0; a32 = 0.0

    det = (a11 * (a22 * a33 - a23 * a32)
           - a12 * (a21 * a33 - a23 * a31)
           + a13 * (a21 * a32 - a22 * a31))
    if abs(det) < 1e-12:
        det = 1e-12
    invdet = 1.0 / det

    b11 = (a22 * a33 - a23 * a32) * invdet
    b12 = (a13 * a32 - a12 * a33) * invdet
    b13 = (a12 * a23 - a13 * a22) * invdet

    b21 = (a23 * a31 - a21 * a33) * invdet
    b22 = (a11 * a33 - a13 * a31) * invdet
    b23 = (a13 * a21 - a11 * a23) * invdet

    b31 = (a21 * a32 - a22 * a31) * invdet
    b32 = (a12 * a31 - a11 * a32) * invdet
    b33 = (a11 * a22 - a12 * a21) * invdet

    Minv_blocks[nid*9 + 0] = b11; Minv_blocks[nid*9 + 1] = b12; Minv_blocks[nid*9 + 2] = b13
    Minv_blocks[nid*9 + 3] = b21; Minv_blocks[nid*9 + 4] = b22; Minv_blocks[nid*9 + 5] = b23
    Minv_blocks[nid*9 + 6] = b31; Minv_blocks[nid*9 + 7] = b32; Minv_blocks[nid*9 + 8] = b33

@cuda.jit(fastmath=True)
def _apply_precond_kernel_per_dof(Ndofs, Minv_blocks, fixed_mask, r_global, z_global):
    i = cuda.grid(1)
    if i >= Ndofs:
        return

    nid = i // 3
    dof = i % 3

    # Check the specific DOF
    if fixed_mask[i] == 1:
        z_global[i] = 0.0
        return

    ri = r_global[i]

    mbase = nid * 9 + dof * 3
    m00 = Minv_blocks[mbase + 0]
    m01 = Minv_blocks[mbase + 1]
    m02 = Minv_blocks[mbase + 2]

    r0 = r_global[nid*3 + 0]
    r1 = r_global[nid*3 + 1]
    r2 = r_global[nid*3 + 2]

    z_global[i] = m00 * r0 + m01 * r1 + m02 * r2

@cuda.jit(fastmath=True)
def _zero_fixed_ap_kernel(Ndofs, fixed_mask, Ap):
    i = cuda.blockIdx.x * cuda.blockDim.x + cuda.threadIdx.x
    if i >= Ndofs:
        return
    if fixed_mask[i] == 1:
        Ap[i] = 0.0

# ---------------------------------------------------------------------------
# GEOMETRIC MULTIGRID PRECONDITIONER (analysis.parameters.preconditioner = "gmg")
# docs/theory.md §14.  Everything below is only used when requested; the
# default block-Jacobi path never touches it.
#
#   * 2x element coarsening per axis (odd counts: the last coarse element is
#     half void-padded, i.e. n_c = ceil(n_f / 2); fine elements outside the
#     domain are skipped = E 0).
#   * trilinear prolongation P (scalar, applied to each displacement component),
#     fine Dirichlet DOFs masked:  P_eff = F P G,  restriction R = P_eff^T.
#   * Galerkin coarse operators  A_c = P_eff^T A P_eff  assembled ONCE per solve as
#     24x24 coarse-ELEMENT matrices (a coarse element contains whole fine
#     elements, so the Galerkin product stays element-wise: the same
#     matrix-free action/diagonal kernels serve every coarse level).
#   * Chebyshev smoother on D^-1 A (D = 3x3 nodal blocks), lambda_max by power
#     iteration, same polynomial pre and post => symmetric V-cycle (SPD M).
#   * coarsest level: dense inverse (host float64, Jacobi-scaled) applied as one
#     matrix-vector product.
# ---------------------------------------------------------------------------

PRECONDITIONERS = ("block_jacobi", "gmg")
# Process-wide count of solves per "requested->used" preconditioner (written to run_info.json as preconditioner_used);
# "_announced" marks that the one-line too-small-grid fallback message was printed.
PRECONDITIONER_USED = {}

_HX_DI = np.array([0, 1, 1, 0, 0, 1, 1, 0], dtype=np.int32)   # hex corner offsets (same order as n0..n7)
_HX_DJ = np.array([0, 0, 1, 1, 0, 0, 1, 1], dtype=np.int32)
_HX_DK = np.array([0, 0, 0, 0, 1, 1, 1, 1], dtype=np.int32)

_GMG_POWER_ITERS = 12        # power iterations for lambda_max(D^-1 A) per level
_GMG_CHEB_HI = 1.1           # upper Chebyshev bound = this x estimated lambda_max
_GMG_CHEB_RATIO = 6.0        # lower bound = upper / this (smoother damps the upper 5/6 of the spectrum; two-grid factor 0.25 vs 0.31 at 4)


@cuda.jit(fastmath=True)
def _gmg_galerkin_kernel(nxf, nyf, nzf, nxc, nyc, nzc, E_f, Kel_f, Khat, use_kel, fixed_f, Kc_out):
    """
    One thread per COARSE element: Kc = sum over its (up to 8) fine children of
    W^T F K_child F W  with the trilinear weights W (8 child nodes -> 8 coarse nodes) and F the
    fine free-DOF flags. Children outside the fine domain (odd counts) are skipped.
    Fine child matrix: E_f[e] * Khat (use_kel == 0) or the per-element Kel_f[(i*24+j)*Nef + e].
    Output layout Kc_out[(i*24+j)*Nec + ce] (entry-major: coalesced across threads).
    """
    ce = cuda.grid(1)
    nxec = nxc - 1; nyec = nyc - 1; nzec = nzc - 1
    Nec = nxec * nyec * nzec
    if ce >= Nec:
        return
    ck = ce % nzec
    cj = (ce // nzec) % nyec
    ci = ce // (nyec * nzec)
    nxef = nxf - 1; nyef = nyf - 1; nzef = nzf - 1
    Nef = nxef * nyef * nzef

    Kc = cuda.local.array(576, float32)
    T = cuda.local.array(576, float32)
    W = cuda.local.array(64, float32)
    fr = cuda.local.array(24, float32)
    for i in range(576):
        Kc[i] = 0.0

    for m in range(2):
        for n in range(2):
            for l in range(2):
                fi = 2 * ci + m
                fj = 2 * cj + n
                fk = 2 * ck + l
                if fi >= nxef or fj >= nyef or fk >= nzef:
                    continue
                fe = (fi * nyef + fj) * nzef + fk
                for a in range(8):
                    node = (fi + _HX_DI[a]) * (nyf * nzf) + (fj + _HX_DJ[a]) * nzf + (fk + _HX_DK[a])
                    for r in range(3):
                        if fixed_f[3 * node + r] == 1:
                            fr[3 * a + r] = 0.0
                        else:
                            fr[3 * a + r] = 1.0
                    for b in range(8):
                        wx = 1.0 - 0.5 * abs(m + _HX_DI[a] - 2 * _HX_DI[b])
                        wy = 1.0 - 0.5 * abs(n + _HX_DJ[a] - 2 * _HX_DJ[b])
                        wz = 1.0 - 0.5 * abs(l + _HX_DK[a] - 2 * _HX_DK[b])
                        if wx < 0.0:
                            wx = 0.0
                        if wy < 0.0:
                            wy = 0.0
                        if wz < 0.0:
                            wz = 0.0
                        W[a * 8 + b] = wx * wy * wz
                E = E_f[fe]
                # T[i, 3b+s] = sum_beta W[beta, b] fr[3 beta + s] K[i, 3 beta + s]
                for i in range(24):
                    for b in range(8):
                        for s in range(3):
                            acc = 0.0
                            for beta in range(8):
                                j = 3 * beta + s
                                if use_kel == 1:
                                    kv = Kel_f[(i * 24 + j) * Nef + fe]
                                else:
                                    kv = E * Khat[i * 24 + j]
                                acc += W[beta * 8 + b] * fr[j] * kv
                            T[i * 24 + 3 * b + s] = acc
                # Kc[3a+r, :] += sum_alpha W[alpha, a] fr[3 alpha + r] T[3 alpha + r, :]
                for a in range(8):
                    for r in range(3):
                        for c in range(24):
                            acc = 0.0
                            for alpha in range(8):
                                acc += W[alpha * 8 + a] * fr[3 * alpha + r] * T[(3 * alpha + r) * 24 + c]
                            Kc[(3 * a + r) * 24 + c] += acc
    for i in range(576):
        Kc_out[i * Nec + ce] = Kc[i]


@cuda.jit(fastmath=True)
def _gmg_action_kernel(nx, ny, nz, Kel, p_global, Ap_global):
    """Ap += K p for an operator stored as per-element 24x24 matrices Kel[(i*24+j)*Ne + e]."""
    e = cuda.grid(1)
    ny_e = ny - 1
    nz_e = nz - 1
    Ne = (nx - 1) * ny_e * nz_e
    if e >= Ne:
        return
    ek = e % nz_e
    ej = (e // nz_e) % ny_e
    ei = e // (ny_e * nz_e)
    nodes = cuda.local.array(8, int32)
    p_e = cuda.local.array(24, float32)
    for a in range(8):
        nodes[a] = (ei + _HX_DI[a]) * (ny * nz) + (ej + _HX_DJ[a]) * nz + (ek + _HX_DK[a])
        for r in range(3):
            p_e[3 * a + r] = p_global[3 * nodes[a] + r]
    for i in range(24):
        s = 0.0
        for j in range(24):
            s += Kel[(i * 24 + j) * Ne + e] * p_e[j]
        cuda.atomic.add(Ap_global, 3 * nodes[i // 3] + (i % 3), s)


@cuda.jit(fastmath=True)
def _gmg_diag_kernel(nx, ny, nz, Kel, diag_blocks):
    """Nodal 3x3 diagonal blocks of an element-matrix operator (atomics into diag_blocks[9 * node])."""
    e = cuda.grid(1)
    ny_e = ny - 1
    nz_e = nz - 1
    Ne = (nx - 1) * ny_e * nz_e
    if e >= Ne:
        return
    ek = e % nz_e
    ej = (e // nz_e) % ny_e
    ei = e // (ny_e * nz_e)
    for a in range(8):
        node = (ei + _HX_DI[a]) * (ny * nz) + (ej + _HX_DJ[a]) * nz + (ek + _HX_DK[a])
        for rr in range(3):
            for cc in range(3):
                cuda.atomic.add(diag_blocks, node * 9 + rr * 3 + cc,
                                Kel[((3 * a + rr) * 24 + 3 * a + cc) * Ne + e])


@cuda.jit(fastmath=True)
def _gmg_coarse_mask_kernel(nxf, nyf, nzf, nxc, nyc, nzc, fixed_f, fixed_c):
    """
    Coarse Dirichlet mask, one thread per coarse node: the coarse DOF is fixed if the coincident fine
    node's DOF is fixed, or if no free fine DOF lies in its trilinear support (e.g. a padded coarse node
    next to a fixed fine boundary), which would make the Galerkin column vanish.
    """
    c = cuda.grid(1)
    if c >= nxc * nyc * nzc:
        return
    K = c % nzc
    J = (c // nzc) % nyc
    I = c // (nyc * nzc)
    for d in range(3):
        fixed = 0
        if 2 * I < nxf and 2 * J < nyf and 2 * K < nzf:
            if fixed_f[3 * ((2 * I) * (nyf * nzf) + (2 * J) * nzf + (2 * K)) + d] == 1:
                fixed = 1
        if fixed == 0:
            wfree = 0.0
            for a in range(-1, 2):
                ii = 2 * I + a
                if ii < 0 or ii >= nxf:
                    continue
                wx = 1.0 - 0.5 * abs(a)
                for b in range(-1, 2):
                    jj = 2 * J + b
                    if jj < 0 or jj >= nyf:
                        continue
                    wy = 1.0 - 0.5 * abs(b)
                    for g in range(-1, 2):
                        kk = 2 * K + g
                        if kk < 0 or kk >= nzf:
                            continue
                        wz = 1.0 - 0.5 * abs(g)
                        if fixed_f[3 * (ii * (nyf * nzf) + jj * nzf + kk) + d] == 0:
                            wfree += wx * wy * wz
            if wfree <= 0.0:
                fixed = 1
        fixed_c[3 * c + d] = fixed


@cuda.jit(fastmath=True)
def _gmg_restrict_kernel(nxf, nyf, nzf, nxc, nyc, nzc, rf, fixed_c, bc):
    """bc = P^T rf (trilinear weights; gather, no atomics); zero at coarse fixed DOFs. rf must be 0 at fixed fine DOFs."""
    c = cuda.grid(1)
    if c >= nxc * nyc * nzc:
        return
    K = c % nzc
    J = (c // nzc) % nyc
    I = c // (nyc * nzc)
    s0 = 0.0
    s1 = 0.0
    s2 = 0.0
    for a in range(-1, 2):
        ii = 2 * I + a
        if ii < 0 or ii >= nxf:
            continue
        wx = 1.0 - 0.5 * abs(a)
        for b in range(-1, 2):
            jj = 2 * J + b
            if jj < 0 or jj >= nyf:
                continue
            wy = wx * (1.0 - 0.5 * abs(b))
            for g in range(-1, 2):
                kk = 2 * K + g
                if kk < 0 or kk >= nzf:
                    continue
                w = wy * (1.0 - 0.5 * abs(g))
                base = 3 * (ii * (nyf * nzf) + jj * nzf + kk)
                s0 += w * rf[base]
                s1 += w * rf[base + 1]
                s2 += w * rf[base + 2]
    if fixed_c[3 * c] == 1:
        s0 = 0.0
    if fixed_c[3 * c + 1] == 1:
        s1 = 0.0
    if fixed_c[3 * c + 2] == 1:
        s2 = 0.0
    bc[3 * c] = s0
    bc[3 * c + 1] = s1
    bc[3 * c + 2] = s2


@cuda.jit(fastmath=True)
def _gmg_prolong_add_kernel(nxf, nyf, nzf, nxc, nyc, nzc, xc, fixed_f, xf):
    """xf += P xc (trilinear; gather); not added at fixed fine DOFs."""
    n = cuda.grid(1)
    if n >= nxf * nyf * nzf:
        return
    k = n % nzf
    j = (n // nzf) % nyf
    i = n // (nyf * nzf)
    s0 = 0.0
    s1 = 0.0
    s2 = 0.0
    for a in range(2):
        if i % 2 == 0:
            if a == 1:
                continue
            ci = i // 2
            wx = 1.0
        else:
            ci = (i - 1) // 2 + a
            wx = 0.5
        for b in range(2):
            if j % 2 == 0:
                if b == 1:
                    continue
                cj = j // 2
                wy = wx
            else:
                cj = (j - 1) // 2 + b
                wy = wx * 0.5
            for g in range(2):
                if k % 2 == 0:
                    if g == 1:
                        continue
                    ck = k // 2
                    w = wy
                else:
                    ck = (k - 1) // 2 + g
                    w = wy * 0.5
                base = 3 * (ci * (nyc * nzc) + cj * nzc + ck)
                s0 += w * xc[base]
                s1 += w * xc[base + 1]
                s2 += w * xc[base + 2]
    if fixed_f[3 * n] == 0:
        xf[3 * n] += s0
    if fixed_f[3 * n + 1] == 0:
        xf[3 * n + 1] += s1
    if fixed_f[3 * n + 2] == 0:
        xf[3 * n + 2] += s2


@cuda.jit(fastmath=True)
def _gmg_cheb_kernel(Ndofs, Minv_blocks, fixed_mask, r, d, x, c1, c2, mode):
    """
    Chebyshev update with the block-Jacobi scaling: t = D^-1 r; d = (mode&1 ? 0 : c1 d) + c2 t;
    x = (mode&2 ? d : x + d). Zero at fixed DOFs.
    """
    i = cuda.grid(1)
    if i >= Ndofs:
        return
    if fixed_mask[i] == 1:
        d[i] = 0.0
        if (mode & 2) != 0:
            x[i] = 0.0
        return
    nid = i // 3
    mb = nid * 9 + (i % 3) * 3
    t = Minv_blocks[mb] * r[3 * nid] + Minv_blocks[mb + 1] * r[3 * nid + 1] + Minv_blocks[mb + 2] * r[3 * nid + 2]
    if (mode & 1) != 0:
        dn = c2 * t
    else:
        dn = c1 * d[i] + c2 * t
    d[i] = dn
    if (mode & 2) != 0:
        x[i] = dn
    else:
        x[i] += dn


@cuda.jit(fastmath=True)
def _gmg_resid_kernel(Ndofs, fixed_mask, b, Ad, r):
    """r = b - Ad, zero at fixed DOFs (the unmasked action output is only consumed here)."""
    i = cuda.grid(1)
    if i >= Ndofs:
        return
    if fixed_mask[i] == 1:
        r[i] = 0.0
    else:
        r[i] = b[i] - Ad[i]


@cuda.jit(fastmath=True)
def _gmg_seed_kernel(fixed_mask, v):
    """Deterministic pseudo-random start vector in (-0.5, 0.5), zero at fixed DOFs."""
    i = cuda.grid(1)
    if i >= v.size:
        return
    h = (i * 2654435761 + 12345) & 0xFFFFFFFF
    h = h ^ (h >> 13)
    h = (h * 1274126177) & 0xFFFFFFFF
    h = h ^ (h >> 16)
    if fixed_mask[i] == 1:
        v[i] = 0.0
    else:
        v[i] = h / 4294967296.0 - 0.5


@cuda.jit(fastmath=True)
def _gmg_dense_apply_kernel(N, Minv, r, z):
    """z = Minv r for a dense symmetric N x N matrix (column access: coalesced)."""
    i = cuda.grid(1)
    if i >= N:
        return
    s = 0.0
    for j in range(N):
        s += Minv[j * N + i] * r[j]
    z[i] = s


_GMG_CACHE = {}


def gmg_cache_clear():
    """Drop the stored multigrid hierarchy of `gmg_reuse_max_rel_change` (frees its device memory)."""
    _GMG_CACHE.clear()


def gmg_level_dims(nx, ny, nz, levels=None, coarse_max_nodes=400):
    """
    Node-count triples [(nx, ny, nz), ...] of the multigrid hierarchy, finest first. Each level has
    ceil(n_elems / 2) elements per axis (odd counts are padded by a void half-element). `levels` (int >= 2,
    total number of levels incl. the finest) fixes the depth; None coarsens until the node count is
    <= coarse_max_nodes (or every axis has a single element).
    """
    dims = [(int(nx), int(ny), int(nz))]
    while True:
        cx, cy, cz = dims[-1]
        if levels is not None:
            if len(dims) >= levels:
                break
        elif cx * cy * cz <= coarse_max_nodes:
            break
        if max(cx, cy, cz) <= 2:       # one element per axis: cannot coarsen further
            break
        dims.append(tuple(((n - 1) + 1) // 2 + 1 for n in (cx, cy, cz)))
    if len(dims) < 2:
        raise ValueError("gmg: the grid is too small to coarsen; use preconditioner='block_jacobi'")
    return dims


class _GMGLevel:
    pass


class GMGPreconditioner:
    """
    Symmetric geometric-multigrid V-cycle, z = M^-1 r, for the operator of `matrix_free_pcg`
    (Dirichlet DOFs = zero rows/columns). Level 0 uses the caller's matrix-free action
    `apply_K0(d_in, d_out)` (d_out += K d_in; zeroing and fixed-DOF masking are done here) and
    block-Jacobi inverse `d_Minv0`; its Galerkin coarsening needs the element data (`d_E`, `d_Khat`).
    Reusable for any structured-hex fine grid (e.g. an analysis refined 2x) by passing its own
    dims, mask, action and element data.
    """

    def __init__(self, nx, ny, nz, d_fixed0, apply_K0, d_Minv0, d_E, d_Khat,
                 levels=None, smoother_steps=2, coarse_max_nodes=400, threads=THREADS_PER_BLOCK):
        if int(smoother_steps) < 1:
            raise ValueError("gmg_smoother_steps must be >= 1")
        self.threads = threads
        self.deg = int(smoother_steps)
        dims = gmg_level_dims(nx, ny, nz, levels, coarse_max_nodes)
        self.levels = []
        for l, (a, b, c) in enumerate(dims):
            lv = _GMGLevel()
            lv.nx, lv.ny, lv.nz = a, b, c
            lv.Nn = a * b * c
            lv.Nd = 3 * lv.Nn
            lv.Ne = (a - 1) * (b - 1) * (c - 1)
            lv.blk_d = (lv.Nd + threads - 1) // threads
            lv.blk_n = (lv.Nn + threads - 1) // threads
            lv.blk_e = (lv.Ne + threads - 1) // threads
            lv.r = cuda.device_array(lv.Nd, dtype=np.float32)
            lv.d = cuda.device_array(lv.Nd, dtype=np.float32)
            lv.Ad = cuda.device_array(lv.Nd, dtype=np.float32)
            if l > 0:
                lv.x = cuda.device_array(lv.Nd, dtype=np.float32)
                lv.b = cuda.device_array(lv.Nd, dtype=np.float32)
            else:
                lv.x = None
                lv.b = None
            self.levels.append(lv)
        L0 = self.levels[0]
        L0.fixed = d_fixed0
        L0.Minv = d_Minv0
        L0.Kel = None
        self._apply_K0 = apply_K0
        # coarse levels: mask, Galerkin element matrices, nodal block-Jacobi inverse
        for l in range(1, len(self.levels)):
            f, c = self.levels[l - 1], self.levels[l]
            c.fixed = cuda.device_array(c.Nd, dtype=np.int32)
            _gmg_coarse_mask_kernel[c.blk_n, threads](f.nx, f.ny, f.nz, c.nx, c.ny, c.nz, f.fixed, c.fixed)
            c.Kel = cuda.device_array(576 * c.Ne, dtype=np.float32)
            if l == 1:
                _gmg_galerkin_kernel[c.blk_e, threads](f.nx, f.ny, f.nz, c.nx, c.ny, c.nz,
                                                       d_E, d_E, d_Khat, 0, f.fixed, c.Kel)
            else:
                _gmg_galerkin_kernel[c.blk_e, threads](f.nx, f.ny, f.nz, c.nx, c.ny, c.nz,
                                                       d_E, f.Kel, d_Khat, 1, f.fixed, c.Kel)
            diag = cuda.device_array(c.Nn * 9, dtype=np.float32)
            _zero_array_kernel[(diag.size + threads - 1) // threads, threads](diag)
            _gmg_diag_kernel[c.blk_e, threads](c.nx, c.ny, c.nz, c.Kel, diag)
            c.Minv = cuda.device_array(c.Nn * 9, dtype=np.float32)
            _invert_diag_blocks_kernel[c.blk_n, threads](c.Nn, diag, c.Minv, c.fixed)
        self._build_coarse_solver()
        self._scratch = {}
        for l in range(len(self.levels) - 1):
            self._estimate_spectrum(self.levels[l], l)

    # -- operator on level l: out = A_l in  (fixed rows zeroed) --------------------------------
    def _apply_A(self, l, d_in, d_out, mask=True):
        """d_out = A_l d_in. mask=False leaves the fixed rows unzeroed: only for callers that mask themselves."""
        lv = self.levels[l]
        th = self.threads
        _zero_array_kernel[lv.blk_d, th](d_out)
        if l == 0:
            self._apply_K0(d_in, d_out)
        else:
            _gmg_action_kernel[lv.blk_e, th](lv.nx, lv.ny, lv.nz, lv.Kel, d_in, d_out)
        if mask:
            _zero_fixed_ap_kernel[lv.blk_d, th](lv.Nd, lv.fixed, d_out)

    def _dot(self, a, b):
        n = a.size
        nb = (n + self.threads - 1) // self.threads
        bs = self._scratch.get(nb)
        if bs is None:
            bs = self._scratch[nb] = (cuda.device_array(nb, dtype=np.float32), cuda.device_array(1, dtype=np.float32))
        _vec_dot_block_kernel[nb, self.threads](a, b, bs[0])
        _vec_dot_final_kernel[(1,), self.threads](bs[0], bs[1])
        return float(bs[1].copy_to_host()[0])

    def _estimate_spectrum(self, lv, l):
        """Power iteration on D^-1 A (similar to an SPD matrix): lambda_max estimate and Chebyshev interval."""
        th = self.threads
        v, w = lv.r, lv.d
        _gmg_seed_kernel[lv.blk_d, th](lv.fixed, v)
        lam = 1.0
        for _ in range(_GMG_POWER_ITERS):
            self._apply_A(l, v, lv.Ad)
            _apply_precond_kernel_per_dof[lv.blk_d, th](lv.Nd, lv.Minv, lv.fixed, lv.Ad, w)
            nw = sqrt(max(self._dot(w, w), 1e-60))
            nv = sqrt(max(self._dot(v, v), 1e-60))
            lam = nw / nv
            _vec_scaled_add_kernel[lv.blk_d, th](v, np.float32(1.0 / nw), w, np.float32(0.0), w)
        lv.lam_max = lam
        hi = _GMG_CHEB_HI * lam
        lo = hi / _GMG_CHEB_RATIO
        lv.theta = 0.5 * (hi + lo)
        lv.delta = 0.5 * (hi - lo)

    def _build_coarse_solver(self):
        """Dense Jacobi-scaled inverse of the coarsest Galerkin operator (host float64) -> device float32."""
        lv = self.levels[-1]
        if len(self.levels) == 1:
            raise ValueError("gmg: need at least two levels")
        N = lv.Nd
        Kel = lv.Kel.copy_to_host().reshape(576, lv.Ne).T.reshape(lv.Ne, 24, 24).astype(np.float64)
        dm = hex_dof_map(lv.nx, lv.ny, lv.nz)
        A = np.zeros((N, N))
        np.add.at(A, (dm[:, :, None], dm[:, None, :]), Kel)
        fixed = lv.fixed.copy_to_host() == 1
        free = np.where(~fixed)[0]
        Aff = 0.5 * (A[np.ix_(free, free)] + A[np.ix_(free, free)].T)
        dg = np.diag(Aff).copy()
        dg[dg <= 0.0] = 1.0
        s = 1.0 / np.sqrt(dg)
        S = Aff * s[:, None] * s[None, :]
        try:
            Lc = np.linalg.cholesky(S)
            Li = np.linalg.solve(Lc, np.eye(free.size))
            Sinv = Li.T @ Li
        except np.linalg.LinAlgError:
            w_, V_ = np.linalg.eigh(S)
            keep = w_ > 1e-12 * w_.max()
            Sinv = (V_[:, keep] / w_[keep]) @ V_[:, keep].T
        Minv = np.zeros((N, N))
        Minv[np.ix_(free, free)] = Sinv * s[:, None] * s[None, :]
        lv.dense = cuda.to_device(np.ascontiguousarray(Minv, dtype=np.float32).ravel())
        lv.blk_dense = (N + self.threads - 1) // self.threads

    # -- smoother ------------------------------------------------------------------------------
    def _cheb(self, l, zero_init, want_r):
        """x <- x + p(D^-1 A) D^-1 (b - A x), p = degree-`deg` Chebyshev polynomial; leaves r = b - A x if want_r."""
        lv = self.levels[l]
        th = self.threads
        deg = self.deg
        theta, delta = lv.theta, lv.delta
        sigma = theta / delta
        rho = 1.0 / sigma
        if zero_init:
            rsrc = lv.b
        else:
            self._apply_A(l, lv.x, lv.Ad, mask=False)
            _gmg_resid_kernel[lv.blk_d, th](lv.Nd, lv.fixed, lv.b, lv.Ad, lv.r)
            rsrc = lv.r
        _gmg_cheb_kernel[lv.blk_d, th](lv.Nd, lv.Minv, lv.fixed, rsrc, lv.d, lv.x,
                                       np.float32(0.0), np.float32(1.0 / theta), 3 if zero_init else 1)
        for _ in range(1, deg):
            self._apply_A(l, lv.d, lv.Ad, mask=False)
            _gmg_resid_kernel[lv.blk_d, th](lv.Nd, lv.fixed, rsrc, lv.Ad, lv.r)
            rsrc = lv.r
            rho_new = 1.0 / (2.0 * sigma - rho)
            _gmg_cheb_kernel[lv.blk_d, th](lv.Nd, lv.Minv, lv.fixed, lv.r, lv.d, lv.x,
                                           np.float32(rho_new * rho), np.float32(2.0 * rho_new / delta), 0)
            rho = rho_new
        if want_r:
            self._apply_A(l, lv.d, lv.Ad, mask=False)
            _gmg_resid_kernel[lv.blk_d, th](lv.Nd, lv.fixed, rsrc, lv.Ad, lv.r)

    # -- V-cycle -----------------------------------------------------------------------------------
    def apply(self, d_r, d_z):
        """d_z = M^-1 d_r (one V-cycle from a zero initial guess)."""
        th = self.threads
        Ls = self.levels
        nl = len(Ls)
        Ls[0].b = d_r
        Ls[0].x = d_z
        for l in range(nl - 1):
            self._cheb(l, True, True)
            f, c = Ls[l], Ls[l + 1]
            _gmg_restrict_kernel[c.blk_n, th](f.nx, f.ny, f.nz, c.nx, c.ny, c.nz, f.r, c.fixed, c.b)
        cl = Ls[-1]
        _gmg_dense_apply_kernel[cl.blk_dense, th](cl.Nd, cl.dense, cl.b, cl.x)
        for l in range(nl - 2, -1, -1):
            f, c = Ls[l], Ls[l + 1]
            _gmg_prolong_add_kernel[f.blk_n, th](f.nx, f.ny, f.nz, c.nx, c.ny, c.nz, c.x, f.fixed, f.x)
            self._cheb(l, False, False)
        Ls[0].b = None
        Ls[0].x = None

    def rebind(self, d_fixed0, apply_K0, d_Minv0):
        """Re-attach the (stale-coarse-level) hierarchy to a new solve's level-0 mask, action and block-Jacobi inverse."""
        L0 = self.levels[0]
        L0.fixed = d_fixed0
        L0.Minv = d_Minv0
        self._apply_K0 = apply_K0

    def info(self):
        nbytes = 0
        for l, lv in enumerate(self.levels):
            for nm in ("r", "d", "Ad", "x", "b", "Kel", "dense") + (("Minv", "fixed") if l > 0 else ()):
                a = getattr(lv, nm, None)
                if a is not None and hasattr(a, "size"):
                    nbytes += int(a.size) * int(a.dtype.itemsize)
        return dict(levels=[(lv.nx - 1, lv.ny - 1, lv.nz - 1) for lv in self.levels],
                    lam_max=[getattr(lv, "lam_max", None) for lv in self.levels],
                    smoother_steps=self.deg, coarsest_solve="dense_inverse", extra_bytes=nbytes)


def matrix_free_pcg(nx, ny, nz, hx, hy, hz, E_elem_h, nu_elem_h, b_h, fixed_mask_h,
                    max_iters=200, tol=1e-6, rel_tol=0.0, verbose=True,
                    u0=None, warm_start_probe_iters=8, warm_start_min_reduction=0.5,
                    element_type="q1_gauss", nu=None,
                    preconditioner="block_jacobi", gmg_levels=None, gmg_smoother_steps=2,
                    gmg_coarse_max_nodes=400, gmg_reuse_max_rel_change=0.0,
                    pcg_engine="classic", pcg_check_every=None):
    """
    Matrix-free block-Jacobi (or geometric-multigrid) preconditioned conjugate gradient solve of K u = b
    on a regular trilinear-hex grid, entirely on the GPU.

    Convergence test:  ||r||_2 <= max(tol, rel_tol * ||b||_2)

    `tol` is ABSOLUTE (newtons) and `rel_tol` RELATIVE. rel_tol defaults to 0.0
    so that existing callers keep their exact previous behaviour; pass a small
    value (1e-6 is the ConjugateGradient backend's default) to make the solve
    accuracy independent of how the loads happen to be scaled.

    NOTE ON THE RESIDUAL. Dirichlet conditions are imposed by zeroing Ap at
    fixed DOFs, so a residual component seeded at a fixed DOF by a nonzero b
    never decays and would put an unreachable floor under ||r||. Any load on a
    constrained DOF is therefore stripped from b below -- it does no work in any
    case, being carried by the reaction -- which makes ||r|| a true measure of
    the free-DOF error. `ConjugateGradient._volume_mesh` already does this and
    warns the user; the guard is repeated here so that any other caller is safe.

    WARM START (docs/theory.md, "CG warm start"). Pass `u0` (an (Ndofs,) host
    array, typically the previous inner iteration's converged displacement) to
    seed the solve from it instead of zero. This never changes what the solve
    converges TO -- CG with a different starting point still solves the same
    K u = b -- only how fast it gets there, so it is safe to use whenever the
    stiffness matrix and load are close to whatever u0 solved. When they are
    not close (the void/solid stiffness contrast has grown a lot since u0 was
    captured, typically late in a run as the penalisation power ramps up),
    warm-starting can converge SLOWER than from zero, sometimes by a large
    margin.

    This function does not try to predict that in advance; instead it bounds
    the downside directly. After `warm_start_probe_iters` iterations of a
    warm-started solve, the residual is checked against its own starting
    value ||r0|| = ||b - A u0||. If it has not fallen by at least
    `warm_start_min_reduction` (a fraction) by then, the attempt is judged not
    to be paying off and the solve is reset in place -- u, r, p rebuilt exactly
    as the u0=None path builds them -- and continues as an ordinary cold solve
    for the remaining budget. Worst case, a warm start that never helps costs
    `warm_start_probe_iters` extra iterations relative to a pure cold solve;
    it can never run the full max_iters budget twice. The probe/reduction
    values are tuned on a downscaled CPU benchmark (see the published "Warm-
    start CG benchmark" report) and should be treated as a starting point,
    not a theoretical result, until reproduced at production grid sizes.

    ELEMENT TYPE. `element_type` selects the element: "q1_gauss" (default; the
    Gauss-point kernels, unchanged), "q1" (same element, precomputed 24x24
    matrix) or "q1_im" (Wilson-Taylor incompatible-modes hex, condensed 24x24
    matrix); see element_matrices.py and docs/theory.md §13. The matrix
    types use one Poisson ratio `nu` (taken from the stiff elements of
    `nu_elem_h` if None) and apply to the action, the block-Jacobi diagonal and
    -- via compute_element_strain_energy -- the strain energy.

    PRECONDITIONER. "block_jacobi" (default; 3x3 nodal blocks, bit-identical to
    before) or "gmg": a symmetric geometric-multigrid V-cycle (Galerkin coarse
    operators, Chebyshev smoother of degree `gmg_smoother_steps` pre and post,
    dense coarsest solve; see GMGPreconditioner and docs/theory.md §14).
    `gmg_levels` (int >= 2, total levels) fixes the depth, None coarsens until
    <= `gmg_coarse_max_nodes` nodes. Works with every element_type and the warm
    start; the hierarchy is built once per call.

    SETUP REUSE. With gmg_reuse_max_rel_change = c > 0 the hierarchy (Galerkin coarse
    operators, Chebyshev bounds, dense coarsest inverse) of the previous gmg solve in this
    process is reused when the grid, constraint mask, element/nu/h and level settings are
    unchanged and max|E - E_prev| / max(E_prev) <= c; only the level-0 operator, its
    block-Jacobi inverse and mask are rebound to the new E. A stale M is still symmetric
    positive definite (CG stays exact; only the iteration count may rise). c = 0 (default)
    rebuilds every time and keeps nothing between calls. info["gmg"] reports setup_s,
    reused, rel_change and cg_loop_s (the CG loop wall time).

    A "gmg" request on a grid too small to coarsen (fewer than 2 levels) falls back to
    "block_jacobi" automatically (one printed line); info["preconditioner_used"] says which ran.

    PCG ENGINE. "classic" (default; this function, bit-identical to before) or "fast" (docs/theory.md §16): the
    same GMG-preconditioned PCG with gather kernels (no atomics, bitwise reproducible), fused smoother kernels, device-side
    scalars and one host read per batch of `pcg_check_every` iterations (None: adaptive), batched Galerkin setup. Needs
    preconditioner="gmg", element_type "q1"/"q1_im" and a coarsenable grid; otherwise this call runs the classic engine and
    info["pcg_engine_used"] = "classic" says so (info["pcg_engine_note"] why). The fast engine stops at the same exact first
    iteration with ||r|| < stop (float round-off aside, iteration counts agree within +-2 of classic).

    Returns (u_h, history, info) where info = {"warm_started": bool,
    "bailed_to_cold": bool, "bailout_iter": int | None, "preconditioner_used": str}; with "gmg" it also has
    info["gmg"] = {levels, lam_max, setup_s, extra_bytes, ...}.
    """
    if pcg_engine != "classic":
        if pcg_engine not in PCG_ENGINES:
            raise ValueError(f"pcg_engine must be one of {PCG_ENGINES}, got {pcg_engine!r}")
        _jax_note = None
        if pcg_engine in ("jax", "auto"):
            _res, _jax_note = _try_jax(pcg_engine, nx, ny, nz, hx, hy, hz, E_elem_h, nu_elem_h, b_h, fixed_mask_h, max_iters, tol,
                                       rel_tol, verbose, u0, warm_start_probe_iters, warm_start_min_reduction, element_type, nu,
                                       preconditioner, gmg_levels, gmg_smoother_steps, gmg_coarse_max_nodes)
            if _res is not None:
                return _res
            pcg_engine = "fast"            # jax declined or failed: the numba fast engine (then classic), see _try_jax
        _ok, _why = _fast_engine_ok(nx, ny, nz, element_type, preconditioner, gmg_levels, gmg_coarse_max_nodes)
        if _ok:
            if _jax_note:
                _res = _pcg_fast(nx, ny, nz, hx, hy, hz, E_elem_h, nu_elem_h, b_h, fixed_mask_h, max_iters, tol, rel_tol, verbose, u0,
                                 warm_start_probe_iters, warm_start_min_reduction, element_type, nu, gmg_levels, gmg_smoother_steps,
                                 gmg_coarse_max_nodes, gmg_reuse_max_rel_change, pcg_check_every)
                _res[2]["pcg_engine_note"] = _jax_note
                return _res
            return _pcg_fast(nx, ny, nz, hx, hy, hz, E_elem_h, nu_elem_h, b_h, fixed_mask_h, max_iters, tol, rel_tol, verbose, u0,
                             warm_start_probe_iters, warm_start_min_reduction, element_type, nu, gmg_levels, gmg_smoother_steps,
                             gmg_coarse_max_nodes, gmg_reuse_max_rel_change, pcg_check_every)
        _res = matrix_free_pcg(nx, ny, nz, hx, hy, hz, E_elem_h, nu_elem_h, b_h, fixed_mask_h, max_iters=max_iters, tol=tol,
                               rel_tol=rel_tol, verbose=verbose, u0=u0, warm_start_probe_iters=warm_start_probe_iters,
                               warm_start_min_reduction=warm_start_min_reduction, element_type=element_type, nu=nu,
                               preconditioner=preconditioner, gmg_levels=gmg_levels, gmg_smoother_steps=gmg_smoother_steps,
                               gmg_coarse_max_nodes=gmg_coarse_max_nodes, gmg_reuse_max_rel_change=gmg_reuse_max_rel_change)
        _res[2]["pcg_engine_used"] = "classic"
        _res[2]["pcg_engine_note"] = _why
        return _res
    if preconditioner not in PRECONDITIONERS:
        raise ValueError(f"preconditioner must be one of {PRECONDITIONERS}, got {preconditioner!r}")
    preconditioner_requested = preconditioner
    if preconditioner == "gmg":
        try:
            gmg_level_dims(nx, ny, nz, gmg_levels, gmg_coarse_max_nodes)
        except ValueError:
            # a grid that cannot be coarsened (fewer than 2 levels): block-Jacobi, announced once per process
            preconditioner = "block_jacobi"
            if not PRECONDITIONER_USED.get("_announced"):
                print(f"gmg: grid {nx - 1}x{ny - 1}x{nz - 1} elements is too small to coarsen; "
                      f"using preconditioner='block_jacobi' instead", flush=True)
                PRECONDITIONER_USED["_announced"] = True
    _k = preconditioner_requested + "->" + preconditioner
    PRECONDITIONER_USED[_k] = PRECONDITIONER_USED.get(_k, 0) + 1
    fixed_mask_h = np.asarray(fixed_mask_h).astype(np.int32)
    b_h =np.asarray(b_h, dtype=np.float32).copy()
    n_load_on_fixed = int(np.count_nonzero((fixed_mask_h == 1) & (b_h != 0.0)))
    if n_load_on_fixed:
        b_h[fixed_mask_h == 1] = 0.0
        if verbose:
            print(f"Note: removed load from {n_load_on_fixed} constrained DOF(s) "
                  f"(no work is done there; it would otherwise floor the residual).")

    stop_tol = max(float(tol), float(rel_tol) * float(np.linalg.norm(b_h)))
    if verbose:
        print(f"PCG stopping tolerance: {stop_tol:.6e} "
              f"(absolute {tol:.3e}, relative {rel_tol:.3e} x ||b||)")

    Nnodes = nx * ny * nz
    Ndofs = 3 * Nnodes
    nx_elems = nx - 1; ny_elems = ny - 1; nz_elems = nz - 1
    Ne = nx_elems * ny_elems * nz_elems

    # thread / block sizing (explicit, single place)
    threads = THREADS_PER_BLOCK
    blocks_elems = (Ne + threads - 1) // threads    # for element kernels
    blocks_nodes = (Nnodes + threads - 1) // threads # for node kernels (diag/invert/zero_fixed)
    blocks_dofs  = (Ndofs + threads - 1) // threads  # for vector kernels (dot, axpy, copy, per-DOF precond)

    # Transfer data to device
    d_E = cuda.to_device(E_elem_h.astype(np.float32))
    d_nu = cuda.to_device(nu_elem_h.astype(np.float32))
    d_b = cuda.to_device(b_h)

    use_mat = element_type != "q1_gauss"
    if use_mat:
        d_Khat = _khat_to_device(element_type, hx, hy, hz, E_elem_h, nu_elem_h, nu)

    def _apply_K(d_in, d_out):
        # d_out += K d_in (caller zeroes d_out first)
        if use_mat:
            _element_action_kernel_mat[blocks_elems, threads](nx, ny, nz, d_E, d_Khat, d_in, d_out)
        else:
            _element_action_kernel[blocks_elems, threads](nx, ny, nz, hx, hy, hz, d_E, d_nu, d_in, d_out, ELEMS_PER_BLOCK_I, ELEMS_PER_BLOCK_J, ELEMS_PER_BLOCK_K)

    # allocate gpu vectors
    d_u = cuda.device_array(Ndofs, dtype=np.float32)
    d_r = cuda.device_array(Ndofs, dtype=np.float32)
    d_p = cuda.device_array(Ndofs, dtype=np.float32)
    d_z = cuda.device_array(Ndofs, dtype=np.float32)
    d_Ap = cuda.device_array(Ndofs, dtype=np.float32)

    # Dirichlet mask moved up from its original spot (just below, with the
    # rest of the preconditioner arrays) because the warm-start r0 = b - A u0
    # matvec needs it to enforce the boundary condition on Ap0.
    d_fixed = cuda.to_device(fixed_mask_h)

    warm_started = u0 is not None
    bailed_to_cold = False
    bailout_iter = None
    r0_norm_warm = None

    if warm_started:
        # Seed u from u0 rather than zero. u0 must be exactly zero at fixed
        # (Dirichlet) DOFs -- it should already be, coming from a converged
        # solve on the same fixed-DOF pattern, but this is cheap insurance
        # against a stale or mismatched caller.
        u0_h = np.asarray(u0, dtype=np.float32).copy()
        if u0_h.shape[0] != Ndofs:
            raise ValueError(f"u0 has {u0_h.shape[0]} DOFs, expected {Ndofs}.")
        u0_h[fixed_mask_h == 1] = 0.0
        d_u = cuda.to_device(u0_h)

        # r0 = b - A u0 (one extra matvec; the u0=None path below gets this
        # for free because A*0 = 0, so r0 = b exactly).
        d_Ap0 = cuda.device_array(Ndofs, dtype=np.float32)
        _zero_array_kernel[blocks_dofs, threads](d_Ap0)
        _apply_K(d_u, d_Ap0)
        _zero_fixed_ap_kernel[blocks_dofs, threads](Ndofs, d_fixed, d_Ap0)
        d_r = cuda.to_device(b_h)
        _vec_axpy_kernel[blocks_dofs, threads](d_r, np.float32(1.0), d_Ap0)  # r = r - 1*Ap0 = b - A u0

        block_sums0 = cuda.device_array(blocks_dofs, dtype=np.float32)
        out_scalar0 = cuda.device_array(1, dtype=np.float32)
        _vec_dot_block_kernel[blocks_dofs, threads](d_r, d_r, block_sums0)
        _vec_dot_final_kernel[(1,), threads](block_sums0, out_scalar0)
        r0_norm_warm = sqrt(float(out_scalar0.copy_to_host()[0]))
        if verbose:
            print(f"Warm start: ||r0|| = {r0_norm_warm:.6e} "
                  f"(will bail to cold at iter {warm_start_probe_iters} if it hasn't "
                  f"fallen {warm_start_min_reduction:.0%} by then)")
    else:
        # Zero out d_u using the custom kernel
        _zero_array_kernel[blocks_dofs, threads](d_u)

        # set r = b (copy host b_h to device d_r)
        # (use to_device on host array to avoid unnecessary roundtrip)
        d_r = cuda.to_device(b_h)

    # preconditioner arrays
    d_diag = cuda.device_array(Nnodes * 9, dtype=np.float32)
    d_Minv = cuda.device_array(Nnodes * 9, dtype=np.float32)

    # Zero out d_diag using the custom kernel (size Nnodes*9)
    blocks_diag = (d_diag.size + threads - 1) // threads
    _zero_array_kernel[blocks_diag, threads](d_diag)

    # assemble diagonal blocks
    if use_mat:
        _element_diag_kernel_mat[blocks_elems, threads](nx, ny, nz, d_E, d_Khat, d_diag)
    else:
        _element_diag_kernel[blocks_elems, threads](nx, ny, nz, hx, hy, hz, d_E, d_nu, d_diag)

    # invert diag -> Minv
    _invert_diag_blocks_kernel[blocks_nodes, threads](Nnodes, d_diag, d_Minv, d_fixed)

    # copy Minv to host for quick diag check
    Minv_h = d_Minv.copy_to_host().reshape(Nnodes, 9)

    gmg = None
    gmg_setup_s = None
    gmg_reused = False
    gmg_rel_change = None
    if preconditioner == "gmg":
        from time import perf_counter
        cuda.synchronize()
        _t0 = perf_counter()

        def _apply_K_fine(d_in, d_out):
            _apply_K(d_in, d_out)
        _E32 = np.asarray(E_elem_h, dtype=np.float32)
        _key = (nx, ny, nz, element_type, gmg_levels, int(gmg_smoother_steps), int(gmg_coarse_max_nodes),
                None if nu is None else float(nu), float(hx), float(hy), float(hz))
        _ent = _GMG_CACHE.get("entry") if gmg_reuse_max_rel_change > 0.0 else None
        if _ent is not None and _ent["key"] == _key and np.array_equal(_ent["fixed"], fixed_mask_h):
            _rel = float(np.max(np.abs(_E32 - _ent["E"])) / max(float(np.max(_ent["E"])), 1e-30))
            if _rel <= gmg_reuse_max_rel_change:
                gmg = _ent["gmg"]
                gmg.rebind(d_fixed, _apply_K_fine, d_Minv)
                gmg_reused = True
                gmg_rel_change = _rel
        if gmg is None and use_mat:
            d_Khat_g = d_Khat
        elif gmg is None:
            # Galerkin coarsening uses the same trilinear element as the Gauss-point kernels, as a 24x24 matrix
            live = nu_elem_h[E_elem_h > 0]
            nu_g = float(nu) if nu is not None else (float(np.median(live)) if live.size else 0.3)
            d_Khat_g = _khat_to_device("q1", hx, hy, hz, E_elem_h, nu_elem_h, nu_g)
        if gmg is None:
            gmg = GMGPreconditioner(nx, ny, nz, d_fixed, _apply_K_fine, d_Minv, d_E, d_Khat_g,
                                    levels=gmg_levels, smoother_steps=gmg_smoother_steps,
                                    coarse_max_nodes=gmg_coarse_max_nodes, threads=threads)
            if gmg_reuse_max_rel_change > 0.0:
                _GMG_CACHE["entry"] = dict(key=_key, fixed=np.array(fixed_mask_h, copy=True), E=_E32.copy(), gmg=gmg)
            else:
                _GMG_CACHE.clear()
        cuda.synchronize()
        gmg_setup_s = perf_counter() - _t0

    def _apply_M(d_rr, d_zz):
        # z = M^{-1} r
        if gmg is None:
            _zero_array_kernel[blocks_dofs, threads](d_zz)
            _apply_precond_kernel_per_dof[blocks_dofs, threads](Ndofs, d_Minv, d_fixed, d_rr, d_zz)
        else:
            gmg.apply(d_rr, d_zz)

    # apply preconditioner z = M^{-1} r
    _apply_M(d_r, d_z)

    # p = z
    _vec_copy_kernel[blocks_dofs, threads](d_p, d_z)

    # prepare reduction scratch arrays (size must match blocks_dofs)
    block_sums = cuda.device_array(blocks_dofs, dtype=np.float32)
    block_sums2 = cuda.device_array((blocks_dofs + THREADS_PER_BLOCK - 1)//THREADS_PER_BLOCK, dtype=np.float32)
    out_scalar = cuda.device_array(1, dtype=np.float32)

    # compute initial r^T z (use blocks_dofs)
    _vec_dot_block_kernel[blocks_dofs, threads](d_r, d_z, block_sums)
    _vec_dot_final_kernel[(1,), threads](block_sums, out_scalar)
    rtz_old = float(out_scalar.copy_to_host()[0])

    if verbose:
        print("Initial r^T z:", rtz_old)

    history = []
    _t_loop0 = None
    if gmg is not None:
        cuda.synchronize()
        _t_loop0 = perf_counter()
    for it in range(max_iters):
        # compute Ap = K * p
        _zero_array_kernel[blocks_dofs, threads](d_Ap)
        _apply_K(d_p, d_Ap)

        # enforce Dirichlet: zero Ap at fixed DOFs using Ndofs scaling
        _zero_fixed_ap_kernel[blocks_dofs, threads](Ndofs, d_fixed, d_Ap)

        # p^T Ap
        _vec_dot_block_kernel[blocks_dofs, threads](d_p, d_Ap, block_sums)
        _vec_dot_final_kernel[(1,), threads](block_sums, out_scalar)
        pAp = float(out_scalar.copy_to_host()[0])
        if abs(pAp) < 1e-18:
            print("pAp ~ 0, break")
            break
        
        alpha = np.float32(rtz_old / pAp)

        # u = u + alpha * p
        _vec_axpby_kernel[blocks_dofs, threads](alpha, d_u, d_p)

        # r = r - alpha * Ap   (use axpy kernel with alpha)
        _vec_axpy_kernel[blocks_dofs, threads](d_r, alpha, d_Ap)

        # compute norm for stopping (r^T r)
        _vec_dot_block_kernel[blocks_dofs, threads](d_r, d_r, block_sums)
        _vec_dot_final_kernel[(1,), threads](block_sums, out_scalar)
        rnorm = sqrt(float(out_scalar.copy_to_host()[0]))
        history.append(rnorm)
        if verbose:
            print(f"Iter {it:4d} rnorm {rnorm:.6e}")

        if rnorm < stop_tol:
            if verbose:
                print("Converged")
            break

        # Warm-start bailout check (see the WARM START note in the docstring).
        # Only ever considered once, at the probe checkpoint, and only when
        # the solve hasn't already converged (the check above runs first).
        if warm_started and not bailed_to_cold and (it + 1) == warm_start_probe_iters:
            # max(..., tiny) guards the degenerate r0_norm_warm ~ 0 case (u0
            # already essentially the solution); it would otherwise already
            # have broken out via the convergence check above.
            required = (1.0 - warm_start_min_reduction) * max(r0_norm_warm, 1e-30)
            if rnorm > required:
                bailed_to_cold = True
                bailout_iter = it
                if verbose:
                    print(f"Warm start not converging fast enough (||r||={rnorm:.6e} > "
                          f"{required:.6e} required by iter {it}); restarting cold with "
                          f"the remaining {max_iters - it - 1} iterations.")
                # Reset to exactly the u0=None initial state (u=0, r=b,
                # z=M^-1 r, p=z, rtz_old=r^T z) and let the loop continue from
                # the next iteration as an ordinary cold solve. This is the
                # same handful of kernel calls as the initial setup above.
                _zero_array_kernel[blocks_dofs, threads](d_u)
                d_r = cuda.to_device(b_h)
                _apply_M(d_r, d_z)
                _vec_copy_kernel[blocks_dofs, threads](d_p, d_z)
                _vec_dot_block_kernel[blocks_dofs, threads](d_r, d_z, block_sums)
                _vec_dot_final_kernel[(1,), threads](block_sums, out_scalar)
                rtz_old = float(out_scalar.copy_to_host()[0])
                continue

        # z = M^{-1} r
        _apply_M(d_r, d_z)

        # compute r^T z (new)
        _vec_dot_block_kernel[blocks_dofs, threads](d_r, d_z, block_sums)
        _vec_dot_final_kernel[(1,), threads](block_sums, out_scalar)
        rtz_new = float(out_scalar.copy_to_host()[0])

        beta = np.float32(rtz_new / rtz_old)

        # p = z + beta * p   (dest = a*src + b*src2)
        _vec_scaled_add_kernel[blocks_dofs, threads](d_p, np.float32(1.0), d_z, beta, d_p)

        rtz_old = rtz_new

    u_h = d_u.copy_to_host()

    info = {"warm_started": warm_started, "bailed_to_cold": bailed_to_cold, "bailout_iter": bailout_iter,
            "preconditioner_used": preconditioner}
    if gmg is not None:
        info["gmg"] = dict(gmg.info(), setup_s=gmg_setup_s, reused=gmg_reused, rel_change=gmg_rel_change,
                           cg_loop_s=perf_counter() - _t_loop0)
    return u_h, history, info

# =========================================================================
# FAST PCG ENGINE  (pcg_engine="fast"; analysis.parameters.pcg_engine; docs/theory.md §16)
#
# Same mathematics as matrix_free_pcg(preconditioner="gmg") with a matrix element type (q1, q1_im), reorganised for
# speed and determinism:
#   * every operator application is a node-centred GATHER (one thread per node sums its <= 8 incident elements):
#     no atomics, results independent of thread scheduling (bitwise run-to-run reproducible);
#   * the Chebyshev smoother is fused: operator + residual + Jacobi scaling + recurrence + solution update in one
#     kernel per smoothing step (classic: 5 kernels per step), the restriction is fused with the next level's
#     smoother initialisation, the last post-smoothing kernel of the fine level also emits the partial sums of r.z;
#   * PCG scalars (alpha, beta, r.z, ||r||^2, pAp), the stopping flag and the residual history live on the device;
#     the host enqueues a batch of iterations and reads the flag once per batch. After the flag is set every remaining
#     enqueued iteration is an exact no-op (alpha = beta = 0), so the returned iterate is the one of the exact first
#     iteration with ||r|| < stop, and the stored residual history is re-checked on the host;
#   * the Galerkin coarse operators of level 1 are a batched product sum_c E_c (W_c^T Khat W_c) with 8 precomputed
#     24x24 matrices wherever no fixed DOF touches the coarse element; elsewhere (and on levels >= 2) the per-element
#     W^T F K F W product of the classic kernel;
#   * the Chebyshev spectrum estimate runs on the device (no host reads inside the power iteration);
#   * dot-product partial sums are float64 in a fixed-order tree; vectors and operators are float32 as in classic.
# No cupy, no Linux-only calls; extra device memory = classic's (the ping-pong buffer replaces the `Ad` buffer).
# =========================================================================
PCG_ENGINES = ("classic", "fast", "jax", "auto")      # "jax": gpu_jax.py (JAX/XLA); "auto": engines.pcg_choice picks jax or fast
_FB = 128                                                   # block size of the fast-path kernels
_GF_RB = 256                                                 # threads of the single-block scalar reductions
# accumulator type of the node-gather sums. float32: a float64 accumulator (tested on truss, flexure and column final and the void-contrast unit problems) changes
# neither accuracy nor iteration counts measurably, but is 1/64 rate on consumer GPUs.
_GF_ACC = float32
_GF_Q = np.array([0, 4, 3, 7, 1, 5, 2, 6], dtype=np.int32)   # local node index of a node at element-corner offset (a,b,c), key a*4+b*2+c


@cuda.jit(device=True)
def _gf_gather(nx, ny, nz, n, E, Kel, sK, use_kel, v):
    """(A v) at node n (3 components, unmasked): sum over the incident elements of E_e Khat or Kel[e] applied to v."""
    nyz = ny * nz
    k = n % nz
    j = (n // nz) % ny
    i = n // nyz
    nex = nx - 1
    ney = ny - 1
    nez = nz - 1
    Ne = nex * ney * nez
    a0 = _GF_ACC(0.0)
    a1 = _GF_ACC(0.0)
    a2 = _GF_ACC(0.0)
    for a in range(2):
        ei = i - a
        if ei < 0 or ei >= nex:
            continue
        for b in range(2):
            ej = j - b
            if ej < 0 or ej >= ney:
                continue
            for c in range(2):
                ek = k - c
                if ek < 0 or ek >= nez:
                    continue
                e = (ei * ney + ej) * nez + ek
                q = _GF_Q[a * 4 + b * 2 + c]
                n0 = ei * nyz + ej * nz + ek
                s0 = _GF_ACC(0.0)
                s1 = _GF_ACC(0.0)
                s2 = _GF_ACC(0.0)
                if use_kel == 0:
                    Ee = E[e]
                    for qq in range(8):
                        nb = 3 * (n0 + _HX_DI[qq] * nyz + _HX_DJ[qq] * nz + _HX_DK[qq])
                        p0 = v[nb]
                        p1 = v[nb + 1]
                        p2 = v[nb + 2]
                        cc = 72 * q + 3 * qq
                        s0 += sK[cc] * p0 + sK[cc + 1] * p1 + sK[cc + 2] * p2
                        s1 += sK[cc + 24] * p0 + sK[cc + 25] * p1 + sK[cc + 26] * p2
                        s2 += sK[cc + 48] * p0 + sK[cc + 49] * p1 + sK[cc + 50] * p2
                    a0 += Ee * s0
                    a1 += Ee * s1
                    a2 += Ee * s2
                else:
                    for qq in range(8):
                        nb = 3 * (n0 + _HX_DI[qq] * nyz + _HX_DJ[qq] * nz + _HX_DK[qq])
                        p0 = v[nb]
                        p1 = v[nb + 1]
                        p2 = v[nb + 2]
                        cc = (72 * q + 3 * qq) * Ne + e
                        s0 += Kel[cc] * p0 + Kel[cc + Ne] * p1 + Kel[cc + 2 * Ne] * p2
                        cc += 24 * Ne
                        s1 += Kel[cc] * p0 + Kel[cc + Ne] * p1 + Kel[cc + 2 * Ne] * p2
                        cc += 24 * Ne
                        s2 += Kel[cc] * p0 + Kel[cc + Ne] * p1 + Kel[cc + 2 * Ne] * p2
                    a0 += s0
                    a1 += s1
                    a2 += s2
    return a0, a1, a2


@cuda.jit(device=True)
def _gf_block_reduce(sh, tid, loc, part):
    """Fixed-order tree reduction of one float64 value per thread; thread 0 stores the block sum in part[blockIdx.x]."""
    sh[tid] = loc
    cuda.syncthreads()
    s = cuda.blockDim.x // 2
    while s > 0:
        if tid < s:
            sh[tid] += sh[tid + s]
        cuda.syncthreads()
        s //= 2
    if tid == 0:
        part[cuda.blockIdx.x] = sh[0]


@cuda.jit(fastmath=True)
def _gf_step_kernel(nx, ny, nz, E, Kel, Khat, use_kel, din, fixed, rsrc, Minv, dout, rout, x, dotv, part, c1, c2, mode):
    """
    Fused Chebyshev step at one node (3 DOFs):  r = rsrc - A din (0 at fixed DOFs, written to rout; rout may alias rsrc);
    mode bit 0: d = c1 din + c2 Minv r written to dout; bit 1: x += d; bit 2: also x += din (deferred first step of the
    post-smoother); bit 3: block partial sums of (updated x) . dotv -> part (float64).
    """
    sK = cuda.shared.array(576, float32)
    sh = cuda.shared.array(_FB, float64)
    tid = cuda.threadIdx.x
    if use_kel == 0:
        for q in range(tid, 576, cuda.blockDim.x):
            sK[q] = Khat[q]
        cuda.syncthreads()
    n = cuda.grid(1)
    loc = float64(0.0)
    if n < nx * ny * nz:
        a0, a1, a2 = _gf_gather(nx, ny, nz, n, E, Kel, sK, use_kel, din)
        n3 = 3 * n
        r0 = rsrc[n3] - a0
        r1 = rsrc[n3 + 1] - a1
        r2 = rsrc[n3 + 2] - a2
        f0 = fixed[n3]
        f1 = fixed[n3 + 1]
        f2 = fixed[n3 + 2]
        if f0 == 1:
            r0 = float32(0.0)
        if f1 == 1:
            r1 = float32(0.0)
        if f2 == 1:
            r2 = float32(0.0)
        rout[n3] = r0
        rout[n3 + 1] = r1
        rout[n3 + 2] = r2
        if (mode & 1) != 0:
            m9 = 9 * n
            t0 = Minv[m9] * r0 + Minv[m9 + 1] * r1 + Minv[m9 + 2] * r2
            t1 = Minv[m9 + 3] * r0 + Minv[m9 + 4] * r1 + Minv[m9 + 5] * r2
            t2 = Minv[m9 + 6] * r0 + Minv[m9 + 7] * r1 + Minv[m9 + 8] * r2
            g0 = din[n3]
            g1 = din[n3 + 1]
            g2 = din[n3 + 2]
            d0 = c1 * g0 + c2 * t0
            d1 = c1 * g1 + c2 * t1
            d2 = c1 * g2 + c2 * t2
            if f0 == 1:
                d0 = float32(0.0)
            if f1 == 1:
                d1 = float32(0.0)
            if f2 == 1:
                d2 = float32(0.0)
            dout[n3] = d0
            dout[n3 + 1] = d1
            dout[n3 + 2] = d2
            if (mode & 2) != 0:
                u0 = x[n3] + d0
                u1 = x[n3 + 1] + d1
                u2 = x[n3 + 2] + d2
                if (mode & 4) != 0:
                    u0 += g0
                    u1 += g1
                    u2 += g2
                x[n3] = u0
                x[n3 + 1] = u1
                x[n3 + 2] = u2
                if (mode & 8) != 0:
                    loc = float64(u0) * float64(dotv[n3]) + float64(u1) * float64(dotv[n3 + 1]) + float64(u2) * float64(dotv[n3 + 2])
    if (mode & 8) != 0:
        _gf_block_reduce(sh, tid, loc, part)


@cuda.jit(fastmath=True)
def _gf_apply_dot_kernel(nx, ny, nz, E, Kel, Khat, use_kel, p, fixed, Ap, part):
    """Ap = A p (0 at fixed DOFs) and the block partial sums of p . Ap (float64)."""
    sK = cuda.shared.array(576, float32)
    sh = cuda.shared.array(_FB, float64)
    tid = cuda.threadIdx.x
    if use_kel == 0:
        for q in range(tid, 576, cuda.blockDim.x):
            sK[q] = Khat[q]
        cuda.syncthreads()
    n = cuda.grid(1)
    loc = float64(0.0)
    if n < nx * ny * nz:
        a0, a1, a2 = _gf_gather(nx, ny, nz, n, E, Kel, sK, use_kel, p)
        n3 = 3 * n
        if fixed[n3] == 1:
            a0 = float32(0.0)
        if fixed[n3 + 1] == 1:
            a1 = float32(0.0)
        if fixed[n3 + 2] == 1:
            a2 = float32(0.0)
        Ap[n3] = a0
        Ap[n3 + 1] = a1
        Ap[n3 + 2] = a2
        loc = float64(p[n3]) * float64(a0) + float64(p[n3 + 1]) * float64(a1) + float64(p[n3 + 2]) * float64(a2)
    _gf_block_reduce(sh, tid, loc, part)


@cuda.jit(fastmath=True)
def _gf_update_kernel(Nn, x, r, p, Ap, sc, part):
    """x += alpha p; r -= alpha Ap (3 DOFs per thread); block partial sums of r.r (float64). alpha = sc[2] (device scalar)."""
    sh = cuda.shared.array(_FB, float64)
    tid = cuda.threadIdx.x
    n = cuda.grid(1)
    al = float32(sc[2])
    loc = float64(0.0)
    if n < Nn:
        n3 = 3 * n
        r0 = r[n3] - al * Ap[n3]
        r1 = r[n3 + 1] - al * Ap[n3 + 1]
        r2 = r[n3 + 2] - al * Ap[n3 + 2]
        x[n3] += al * p[n3]
        x[n3 + 1] += al * p[n3 + 1]
        x[n3 + 2] += al * p[n3 + 2]
        r[n3] = r0
        r[n3 + 1] = r1
        r[n3 + 2] = r2
        loc = float64(r0) * float64(r0) + float64(r1) * float64(r1) + float64(r2) * float64(r2)
    _gf_block_reduce(sh, tid, loc, part)


@cuda.jit(fastmath=True)
def _gf_p_kernel(Nn, z, p, sc):
    """p = z + beta p (3 DOFs per thread), beta = sc[3] (device scalar)."""
    n = cuda.grid(1)
    if n < Nn:
        be = float32(sc[3])
        n3 = 3 * n
        p[n3] = z[n3] + be * p[n3]
        p[n3 + 1] = z[n3 + 1] + be * p[n3 + 1]
        p[n3 + 2] = z[n3 + 2] + be * p[n3 + 2]


@cuda.jit(fastmath=True)
def _gf_dot_kernel(Nn, a, b, part):
    """Block partial sums of a.b (float64), 3 DOFs per thread."""
    sh = cuda.shared.array(_FB, float64)
    tid = cuda.threadIdx.x
    n = cuda.grid(1)
    loc = float64(0.0)
    if n < Nn:
        n3 = 3 * n
        loc = float64(a[n3]) * float64(b[n3]) + float64(a[n3 + 1]) * float64(b[n3 + 1]) + float64(a[n3 + 2]) * float64(b[n3 + 2])
    _gf_block_reduce(sh, tid, loc, part)


@cuda.jit
def _gf_scalar_kernel(part, nparts, sc, hist, which):
    """
    One block: s = sum(part[:nparts]) (fixed order, float64), then the scalar recurrence of PCG on the device.
    sc: 0 rz, 1 pAp, 2 alpha, 3 beta, 4 r2, 5 iteration count, 6 flag (0 run, 1 converged/max_iters, 2 breakdown), 7 stop^2,
    8 max_iters, 9 rz_new.   which: 0 alpha, 1 ||r||^2 + stopping test + history, 2 beta, 3 plain rz.
    """
    sh = cuda.shared.array(_GF_RB, float64)
    t = cuda.threadIdx.x
    acc = float64(0.0)
    i = t
    while i < nparts:
        acc += part[i]
        i += _GF_RB
    sh[t] = acc
    cuda.syncthreads()
    s = _GF_RB // 2
    while s > 0:
        if t < s:
            sh[t] += sh[t + s]
        cuda.syncthreads()
        s //= 2
    if t == 0:
        v = sh[0]
        if which == 0:
            if sc[6] != 0.0:
                sc[2] = 0.0
            else:
                sc[1] = v
                if abs(v) < 1e-18:
                    sc[2] = 0.0
                    sc[6] = 2.0
                else:
                    sc[2] = sc[0] / v
        elif which == 1:
            if sc[6] == 0.0:
                sc[4] = v
                c = int(sc[5])
                if c < hist.size:
                    hist[c] = v
                sc[5] = c + 1.0
                if v < sc[7] or c + 1 >= sc[8]:
                    sc[6] = 1.0
        elif which == 2:
            if sc[6] != 0.0:
                sc[3] = 0.0
            else:
                sc[9] = v
                if sc[0] != 0.0:
                    sc[3] = v / sc[0]
                else:
                    sc[3] = 0.0
                sc[0] = v
        else:
            sc[0] = v


@cuda.jit(device=True)
def _gf_inv3(a11, a12, a13, a21, a22, a23, a31, a32, a33, fix0, fix1, fix2, Minv, base):
    """Inverse of a nodal 3x3 block with the classic Dirichlet convention (_invert_diag_blocks_kernel); stored row-major."""
    if fix0 == 1:
        a11 = float32(1.0)
        a12 = float32(0.0)
        a13 = float32(0.0)
        a21 = float32(0.0)
        a31 = float32(0.0)
    if fix1 == 1:
        a22 = float32(1.0)
        a12 = float32(0.0)
        a23 = float32(0.0)
        a21 = float32(0.0)
        a32 = float32(0.0)
    if fix2 == 1:
        a33 = float32(1.0)
        a13 = float32(0.0)
        a23 = float32(0.0)
        a31 = float32(0.0)
        a32 = float32(0.0)
    det = (a11 * (a22 * a33 - a23 * a32) - a12 * (a21 * a33 - a23 * a31) + a13 * (a21 * a32 - a22 * a31))
    if abs(det) < 1e-12:
        det = float32(1e-12)
    invdet = float32(1.0) / det
    Minv[base + 0] = (a22 * a33 - a23 * a32) * invdet
    Minv[base + 1] = (a13 * a32 - a12 * a33) * invdet
    Minv[base + 2] = (a12 * a23 - a13 * a22) * invdet
    Minv[base + 3] = (a23 * a31 - a21 * a33) * invdet
    Minv[base + 4] = (a11 * a33 - a13 * a31) * invdet
    Minv[base + 5] = (a13 * a21 - a11 * a23) * invdet
    Minv[base + 6] = (a21 * a32 - a22 * a31) * invdet
    Minv[base + 7] = (a12 * a31 - a11 * a32) * invdet
    Minv[base + 8] = (a11 * a22 - a12 * a21) * invdet


@cuda.jit(fastmath=True)
def _gf_diag_inv_kernel(nx, ny, nz, E, Kel, Khat, use_kel, fixed, Minv):
    """Nodal 3x3 diagonal blocks of the operator by gather (deterministic), inverted in place of the classic
    atomics + _invert_diag_blocks_kernel."""
    n = cuda.grid(1)
    if n >= nx * ny * nz:
        return
    nyz = ny * nz
    k = n % nz
    j = (n // nz) % ny
    i = n // nyz
    nex = nx - 1
    ney = ny - 1
    nez = nz - 1
    Ne = nex * ney * nez
    m = cuda.local.array(9, float32)
    for q in range(9):
        m[q] = float32(0.0)
    for a in range(2):
        ei = i - a
        if ei < 0 or ei >= nex:
            continue
        for b in range(2):
            ej = j - b
            if ej < 0 or ej >= ney:
                continue
            for c in range(2):
                ek = k - c
                if ek < 0 or ek >= nez:
                    continue
                e = (ei * ney + ej) * nez + ek
                q = _GF_Q[a * 4 + b * 2 + c]
                for s in range(3):
                    for t in range(3):
                        if use_kel == 0:
                            m[3 * s + t] += E[e] * Khat[(3 * q + s) * 24 + 3 * q + t]
                        else:
                            m[3 * s + t] += Kel[((3 * q + s) * 24 + 3 * q + t) * Ne + e]
    n3 = 3 * n
    f0 = fixed[n3]
    f1 = fixed[n3 + 1]
    f2 = fixed[n3 + 2]
    if f0 == 1 and f1 == 1 and f2 == 1:
        for q in range(9):
            Minv[9 * n + q] = float32(0.0)
        return
    _gf_inv3(m[0], m[1], m[2], m[3], m[4], m[5], m[6], m[7], m[8], f0, f1, f2, Minv, 9 * n)


@cuda.jit(device=True)
def _gf_galerkin_general(ce, ci, cj, ck, nxf, nyf, nzf, Nec, E_f, Kel_f, Khat, use_kel, fixed_f, Kc_out, Kc, T, W, fr):
    """Classic per-coarse-element Galerkin product sum_children W^T F K F W (any fixed pattern, any child matrices)."""
    nxef = nxf - 1
    nyef = nyf - 1
    nzef = nzf - 1
    Nef = nxef * nyef * nzef
    for i in range(576):
        Kc[i] = 0.0
    for m in range(2):
        for n in range(2):
            for l in range(2):
                fi = 2 * ci + m
                fj = 2 * cj + n
                fk = 2 * ck + l
                if fi >= nxef or fj >= nyef or fk >= nzef:
                    continue
                fe = (fi * nyef + fj) * nzef + fk
                for a in range(8):
                    node = (fi + _HX_DI[a]) * (nyf * nzf) + (fj + _HX_DJ[a]) * nzf + (fk + _HX_DK[a])
                    for r in range(3):
                        if fixed_f[3 * node + r] == 1:
                            fr[3 * a + r] = 0.0
                        else:
                            fr[3 * a + r] = 1.0
                    for b in range(8):
                        wx = 1.0 - 0.5 * abs(m + _HX_DI[a] - 2 * _HX_DI[b])
                        wy = 1.0 - 0.5 * abs(n + _HX_DJ[a] - 2 * _HX_DJ[b])
                        wz = 1.0 - 0.5 * abs(l + _HX_DK[a] - 2 * _HX_DK[b])
                        if wx < 0.0:
                            wx = 0.0
                        if wy < 0.0:
                            wy = 0.0
                        if wz < 0.0:
                            wz = 0.0
                        W[a * 8 + b] = wx * wy * wz
                E = E_f[fe]
                for i in range(24):
                    for b in range(8):
                        for s in range(3):
                            acc = 0.0
                            for beta in range(8):
                                j = 3 * beta + s
                                if use_kel == 1:
                                    kv = Kel_f[(i * 24 + j) * Nef + fe]
                                else:
                                    kv = E * Khat[i * 24 + j]
                                acc += W[beta * 8 + b] * fr[j] * kv
                            T[i * 24 + 3 * b + s] = acc
                for a in range(8):
                    for r in range(3):
                        for c in range(24):
                            acc = 0.0
                            for alpha in range(8):
                                acc += W[alpha * 8 + a] * fr[3 * alpha + r] * T[(3 * alpha + r) * 24 + c]
                            Kc[(3 * a + r) * 24 + c] += acc
    for i in range(576):
        Kc_out[i * Nec + ce] = Kc[i]


@cuda.jit(fastmath=True)
def _gf_galerkin_kernel(nxf, nyf, nzf, nxc, nyc, nzc, E_f, Kel_f, Khat, use_kel, fixed_f, G, Kc_out):
    """
    One thread per coarse element. Level 1 (use_kel == 0) and no fixed fine DOF in the 3x3x3 fine-node footprint:
    Kc = sum_c E_c G_c with the 8 precomputed matrices G_c = W_c^T Khat W_c (batched Galerkin product, G: 8x576 floats).
    Otherwise the classic per-element product (boundary elements next to constraints, and every level >= 2).
    Output layout Kc_out[(i*24+j)*Nec + ce].
    """
    ce = cuda.grid(1)
    nxec = nxc - 1
    nyec = nyc - 1
    nzec = nzc - 1
    Nec = nxec * nyec * nzec
    if ce >= Nec:
        return
    ck = ce % nzec
    cj = (ce // nzec) % nyec
    ci = ce // (nyec * nzec)
    clean = 0
    if use_kel == 0:
        clean = 1
        for a in range(3):
            ii = 2 * ci + a
            if ii >= nxf:
                continue
            for b in range(3):
                jj = 2 * cj + b
                if jj >= nyf:
                    continue
                for c in range(3):
                    kk = 2 * ck + c
                    if kk >= nzf:
                        continue
                    nb = 3 * (ii * (nyf * nzf) + jj * nzf + kk)
                    if fixed_f[nb] == 1 or fixed_f[nb + 1] == 1 or fixed_f[nb + 2] == 1:
                        clean = 0
    if clean == 1:
        nxef = nxf - 1
        nyef = nyf - 1
        nzef = nzf - 1
        e0 = float32(0.0)
        e1 = float32(0.0)
        e2 = float32(0.0)
        e3 = float32(0.0)
        e4 = float32(0.0)
        e5 = float32(0.0)
        e6 = float32(0.0)
        e7 = float32(0.0)
        for m in range(2):
            for n in range(2):
                for l in range(2):
                    fi = 2 * ci + m
                    fj = 2 * cj + n
                    fk = 2 * ck + l
                    if fi < nxef and fj < nyef and fk < nzef:
                        ev = E_f[(fi * nyef + fj) * nzef + fk]
                        cidx = m * 4 + n * 2 + l
                        if cidx == 0:
                            e0 = ev
                        elif cidx == 1:
                            e1 = ev
                        elif cidx == 2:
                            e2 = ev
                        elif cidx == 3:
                            e3 = ev
                        elif cidx == 4:
                            e4 = ev
                        elif cidx == 5:
                            e5 = ev
                        elif cidx == 6:
                            e6 = ev
                        else:
                            e7 = ev
        for ent in range(576):
            s = (e0 * G[ent] + e1 * G[576 + ent] + e2 * G[1152 + ent] + e3 * G[1728 + ent]
                 + e4 * G[2304 + ent] + e5 * G[2880 + ent] + e6 * G[3456 + ent] + e7 * G[4032 + ent])
            Kc_out[ent * Nec + ce] = s
    else:
        Kc = cuda.local.array(576, float32)
        T = cuda.local.array(576, float32)
        W = cuda.local.array(64, float32)
        fr = cuda.local.array(24, float32)
        _gf_galerkin_general(ce, ci, cj, ck, nxf, nyf, nzf, Nec, E_f, Kel_f, Khat, use_kel, fixed_f, Kc_out, Kc, T, W, fr)


@cuda.jit(fastmath=True)
def _gf_restrict_init_kernel(nxf, nyf, nzf, nxc, nyc, nzc, rf, fixed_c, Minv_c, bc, d, x, c2):
    """bc = P^T rf (gather, zero at coarse fixed DOFs) fused with the smoother initialisation d = x = c2 Minv_c bc."""
    c = cuda.grid(1)
    if c >= nxc * nyc * nzc:
        return
    K = c % nzc
    J = (c // nzc) % nyc
    I = c // (nyc * nzc)
    s0 = float32(0.0)
    s1 = float32(0.0)
    s2 = float32(0.0)
    for a in range(-1, 2):
        ii = 2 * I + a
        if ii < 0 or ii >= nxf:
            continue
        wx = 1.0 - 0.5 * abs(a)
        for b in range(-1, 2):
            jj = 2 * J + b
            if jj < 0 or jj >= nyf:
                continue
            wy = wx * (1.0 - 0.5 * abs(b))
            for g in range(-1, 2):
                kk = 2 * K + g
                if kk < 0 or kk >= nzf:
                    continue
                w = float32(wy * (1.0 - 0.5 * abs(g)))
                base = 3 * (ii * (nyf * nzf) + jj * nzf + kk)
                s0 += w * rf[base]
                s1 += w * rf[base + 1]
                s2 += w * rf[base + 2]
    c3 = 3 * c
    f0 = fixed_c[c3]
    f1 = fixed_c[c3 + 1]
    f2 = fixed_c[c3 + 2]
    if f0 == 1:
        s0 = float32(0.0)
    if f1 == 1:
        s1 = float32(0.0)
    if f2 == 1:
        s2 = float32(0.0)
    bc[c3] = s0
    bc[c3 + 1] = s1
    bc[c3 + 2] = s2
    m9 = 9 * c
    t0 = c2 * (Minv_c[m9] * s0 + Minv_c[m9 + 1] * s1 + Minv_c[m9 + 2] * s2)
    t1 = c2 * (Minv_c[m9 + 3] * s0 + Minv_c[m9 + 4] * s1 + Minv_c[m9 + 5] * s2)
    t2 = c2 * (Minv_c[m9 + 6] * s0 + Minv_c[m9 + 7] * s1 + Minv_c[m9 + 8] * s2)
    if f0 == 1:
        t0 = float32(0.0)
    if f1 == 1:
        t1 = float32(0.0)
    if f2 == 1:
        t2 = float32(0.0)
    d[c3] = t0
    d[c3 + 1] = t1
    d[c3 + 2] = t2
    x[c3] = t0
    x[c3 + 1] = t1
    x[c3 + 2] = t2


@cuda.jit(fastmath=True)
def _gf_power_kernel(nx, ny, nz, E, Kel, Khat, use_kel, v, w, sc, slot, fixed, Minv, pw, pv):
    """w = s Minv A v (0 at fixed DOFs), s = sc[slot]; block partial sums of w.w and (s v).(s v) (float64)."""
    sK = cuda.shared.array(576, float32)
    sh = cuda.shared.array(_FB, float64)
    tid = cuda.threadIdx.x
    if use_kel == 0:
        for q in range(tid, 576, cuda.blockDim.x):
            sK[q] = Khat[q]
        cuda.syncthreads()
    n = cuda.grid(1)
    lw = float64(0.0)
    lv = float64(0.0)
    if n < nx * ny * nz:
        s = float32(sc[slot])
        a0, a1, a2 = _gf_gather(nx, ny, nz, n, E, Kel, sK, use_kel, v)
        n3 = 3 * n
        if fixed[n3] == 1:
            a0 = float32(0.0)
        if fixed[n3 + 1] == 1:
            a1 = float32(0.0)
        if fixed[n3 + 2] == 1:
            a2 = float32(0.0)
        m9 = 9 * n
        t0 = s * (Minv[m9] * a0 + Minv[m9 + 1] * a1 + Minv[m9 + 2] * a2)
        t1 = s * (Minv[m9 + 3] * a0 + Minv[m9 + 4] * a1 + Minv[m9 + 5] * a2)
        t2 = s * (Minv[m9 + 6] * a0 + Minv[m9 + 7] * a1 + Minv[m9 + 8] * a2)
        if fixed[n3] == 1:
            t0 = float32(0.0)
        if fixed[n3 + 1] == 1:
            t1 = float32(0.0)
        if fixed[n3 + 2] == 1:
            t2 = float32(0.0)
        w[n3] = t0
        w[n3 + 1] = t1
        w[n3 + 2] = t2
        lw = float64(t0) * float64(t0) + float64(t1) * float64(t1) + float64(t2) * float64(t2)
        v0 = s * v[n3]
        v1 = s * v[n3 + 1]
        v2 = s * v[n3 + 2]
        lv = float64(v0) * float64(v0) + float64(v1) * float64(v1) + float64(v2) * float64(v2)
    _gf_block_reduce(sh, tid, lw, pw)
    cuda.syncthreads()
    _gf_block_reduce(sh, tid, lv, pv)


@cuda.jit
def _gf_power_reduce_kernel(pw, pv, nparts, sc, lam, slot):
    """lam[slot] = sqrt(sum pw / sum pv); sc[slot] = 1/sqrt(sum pw) (normalisation of the next power iteration)."""
    sh = cuda.shared.array((2, _GF_RB), float64)
    t = cuda.threadIdx.x
    a = float64(0.0)
    b = float64(0.0)
    i = t
    while i < nparts:
        a += pw[i]
        b += pv[i]
        i += _GF_RB
    sh[0, t] = a
    sh[1, t] = b
    cuda.syncthreads()
    s = _GF_RB // 2
    while s > 0:
        if t < s:
            sh[0, t] += sh[0, t + s]
            sh[1, t] += sh[1, t + s]
        cuda.syncthreads()
        s //= 2
    if t == 0:
        ww = sh[0, 0]
        if ww < 1e-60:
            ww = 1e-60
        vv = sh[1, 0]
        if vv < 1e-60:
            vv = 1e-60
        lam[slot] = (ww / vv) ** 0.5
        sc[slot] = 1.0 / ww ** 0.5


_GF_G_CACHE = {}


def _gf_child_galerkin_mats(Khat_flat):
    """The 8 matrices G_c = W_c^T Khat W_c (W_c: the 24x24 trilinear child-to-coarse transfer of child c = m*4+n*2+l), float32, flat 8*576."""
    key = Khat_flat.tobytes()
    G_ = _GF_G_CACHE.get(key)
    if G_ is None:
        Kh = np.asarray(Khat_flat, dtype=np.float64).reshape(24, 24)
        out = np.zeros((8, 24, 24))
        for m in range(2):
            for n in range(2):
                for l in range(2):
                    Wm = np.zeros((24, 24))
                    for a in range(8):
                        for b in range(8):
                            wx = max(0.0, 1.0 - 0.5 * abs(m + _HX_DI[a] - 2 * _HX_DI[b]))
                            wy = max(0.0, 1.0 - 0.5 * abs(n + _HX_DJ[a] - 2 * _HX_DJ[b]))
                            wz = max(0.0, 1.0 - 0.5 * abs(l + _HX_DK[a] - 2 * _HX_DK[b]))
                            for s in range(3):
                                Wm[3 * a + s, 3 * b + s] = wx * wy * wz
                    out[m * 4 + n * 2 + l] = Wm.T @ Kh @ Wm
        G_ = np.ascontiguousarray(out.astype(np.float32)).ravel()
        if len(_GF_G_CACHE) > 8:
            _GF_G_CACHE.clear()
        _GF_G_CACHE[key] = G_
    return G_


class GMGFastPreconditioner(GMGPreconditioner):
    """
    Same V-cycle as GMGPreconditioner (same hierarchy, Galerkin operators, Chebyshev interval from 12 power iterations, dense
    coarsest inverse) with fused gather kernels (see the block comment above). Level 0 is the E-weighted element matrix `d_Khat`
    (matrix element types only). Not a drop-in for `apply_K0`-style callers: it owns its level-0 operator.
    """

    def __init__(self, nx, ny, nz, d_fixed0, d_E, d_Khat, levels=None, smoother_steps=2, coarse_max_nodes=400, threads=_FB, profile=False):
        if int(smoother_steps) < 1:
            raise ValueError("gmg_smoother_steps must be >= 1")
        self.threads = threads
        self.deg = int(smoother_steps)
        self.prof = {} if profile else None     # stage timings (synchronising) when profile=True
        self._tp = None
        self._tick("start")
        self.d_Khat = d_Khat
        self.d_E = d_E
        self._dummy64 = cuda.device_array(1, dtype=np.float64)
        dims = gmg_level_dims(nx, ny, nz, levels, coarse_max_nodes)
        self.levels = []
        th = threads
        for l, (a, b, c) in enumerate(dims):
            lv = _GMGLevel()
            lv.nx, lv.ny, lv.nz = a, b, c
            lv.Nn = a * b * c
            lv.Nd = 3 * lv.Nn
            lv.Ne = (a - 1) * (b - 1) * (c - 1)
            lv.blk_d = (lv.Nd + th - 1) // th
            lv.blk_n = (lv.Nn + th - 1) // th
            lv.blk_e = (lv.Ne + th - 1) // th
            lv.r = cuda.device_array(lv.Nd, dtype=np.float32)
            lv.d = cuda.device_array(lv.Nd, dtype=np.float32)
            lv.d2 = cuda.device_array(lv.Nd, dtype=np.float32)
            if l > 0:
                lv.x = cuda.device_array(lv.Nd, dtype=np.float32)
                lv.b = cuda.device_array(lv.Nd, dtype=np.float32)
            else:
                lv.x = None
                lv.b = None
            lv.Kel = None
            self.levels.append(lv)
        Ls = self.levels
        L0 = Ls[0]
        L0.fixed = d_fixed0
        L0.Minv = cuda.device_array(L0.Nn * 9, dtype=np.float32)
        self._level0_minv()
        self._tick("alloc+minv0")
        G_dev = cuda.to_device(_gf_child_galerkin_mats(np.asarray(d_Khat.copy_to_host(), dtype=np.float32)))
        for l in range(1, len(Ls)):
            f, c = Ls[l - 1], Ls[l]
            c.fixed = cuda.device_array(c.Nd, dtype=np.int32)
            _gmg_coarse_mask_kernel[c.blk_n, th](f.nx, f.ny, f.nz, c.nx, c.ny, c.nz, f.fixed, c.fixed)
            c.Kel = cuda.device_array(576 * c.Ne, dtype=np.float32)
            if l == 1:
                _gf_galerkin_kernel[c.blk_e, th](f.nx, f.ny, f.nz, c.nx, c.ny, c.nz, d_E, d_E, d_Khat, 0, f.fixed, G_dev, c.Kel)
            else:
                _gf_galerkin_kernel[c.blk_e, th](f.nx, f.ny, f.nz, c.nx, c.ny, c.nz, d_E, f.Kel, d_Khat, 1, f.fixed, G_dev, c.Kel)
            c.Minv = cuda.device_array(c.Nn * 9, dtype=np.float32)
            _gf_diag_inv_kernel[c.blk_n, th](c.nx, c.ny, c.nz, d_E, c.Kel, d_Khat, 1, c.fixed, c.Minv)
            self._tick(f"galerkin_L{l}")
        # coarsest operator to the host (waits for the kernels above), then the spectrum estimates are enqueued
        # asynchronously and overlap the host-side dense inverse
        lc = Ls[-1]
        Kel_h = lc.Kel.copy_to_host().reshape(576, lc.Ne).T.reshape(lc.Ne, 24, 24).astype(np.float64)
        fixed_h = lc.fixed.copy_to_host()
        self._tick("coarsest_to_host")
        nl = len(Ls)
        self._sc = cuda.to_device(np.ones(nl, dtype=np.float64))
        self._lam = cuda.device_array(nl, dtype=np.float64)
        for l in range(nl - 1):
            self._power_iterations(l)
        self._tick("power_enqueued")
        dense_h = _dense_coarse_inverse(Kel_h, fixed_h, lc)
        self._tick("dense_inverse_host")
        lc.dense = cuda.to_device(dense_h)
        lc.blk_dense = (lc.Nd + th - 1) // th
        lam = self._lam.copy_to_host()
        self._set_chebyshev(lam)
        self._tick("power_done")

    def _tick(self, name):
        if self.prof is not None:
            from time import perf_counter
            cuda.synchronize()
            t = perf_counter()
            if self._tp is not None:
                self.prof[name] = t - self._tp
            self._tp = t

    def _level0_minv(self):
        L0 = self.levels[0]
        _gf_diag_inv_kernel[L0.blk_n, self.threads](L0.nx, L0.ny, L0.nz, self.d_E, self.d_E, self.d_Khat, 0, L0.fixed, L0.Minv)

    def _op_args(self, l):
        """(E, Kel, use_kel) for gather kernels of level l (arrays are dummies where unused; all float32 1-D)."""
        lv = self.levels[l]
        if l == 0:
            return self.d_E, self.d_E, 0
        return self.d_E, lv.Kel, 1

    def _power_iterations(self, l):
        lv = self.levels[l]
        th = self.threads
        E, Kel, uk = self._op_args(l)
        pw = cuda.device_array(lv.blk_n, dtype=np.float64)
        pv = cuda.device_array(lv.blk_n, dtype=np.float64)
        v, w = lv.r, lv.d
        _gmg_seed_kernel[lv.blk_d, th](lv.fixed, v)
        for _ in range(_GMG_POWER_ITERS):
            _gf_power_kernel[lv.blk_n, th](lv.nx, lv.ny, lv.nz, E, Kel, self.d_Khat, uk, v, w, self._sc, l, lv.fixed, lv.Minv, pw, pv)
            _gf_power_reduce_kernel[(1,), _GF_RB](pw, pv, lv.blk_n, self._sc, self._lam, l)
            v, w = w, v

    def _set_chebyshev(self, lam):
        """Per-level Chebyshev interval and the (c1, c2) coefficients of every smoothing step (same recurrence as the classic)."""
        for l, lv in enumerate(self.levels[:-1]):
            lv.lam_max = float(lam[l])
            hi = _GMG_CHEB_HI * lv.lam_max
            lo = hi / _GMG_CHEB_RATIO
            lv.theta = 0.5 * (hi + lo)
            lv.delta = 0.5 * (hi - lo)
            sigma = lv.theta / lv.delta
            rho = 1.0 / sigma
            co = [(np.float32(0.0), np.float32(1.0 / lv.theta))]
            for _ in range(1, self.deg):
                rho_new = 1.0 / (2.0 * sigma - rho)
                co.append((np.float32(rho_new * rho), np.float32(2.0 * rho_new / lv.delta)))
                rho = rho_new
            lv.coef = co

    # -- smoothing -----------------------------------------------------------------------------------
    def _step(self, l, din, rsrc, dout, x, c1, c2, mode, dotv=None, part=None):
        lv = self.levels[l]
        E, Kel, uk = self._op_args(l)
        _gf_step_kernel[lv.blk_n, self.threads](lv.nx, lv.ny, lv.nz, E, Kel, self.d_Khat, uk, din, lv.fixed, rsrc, lv.Minv,
                                                dout, lv.r, x, lv.r if dotv is None else dotv, self._dummy64 if part is None else part,
                                                c1, c2, mode)

    def _pre(self, l):
        """Pre-smoothing from the zero guess; the init (d = x = c2_0 Minv b) was done by the caller. Leaves lv.r = b - A x."""
        lv = self.levels[l]
        din, dout = lv.d, lv.d2
        rsrc = lv.b
        for k in range(1, self.deg):
            c1, c2 = lv.coef[k]
            self._step(l, din, rsrc, dout, lv.x, c1, c2, 3)
            rsrc = lv.r
            din, dout = dout, din
        self._step(l, din, rsrc, din, lv.x, np.float32(0.0), np.float32(0.0), 0)

    def _post(self, l, dotv=None, part=None):
        """Post-smoothing of lv.x (nonzero guess); at level 0 with dotv/part also emits the partial sums of x_final . dotv."""
        lv = self.levels[l]
        th = self.threads
        c1, c2 = lv.coef[0]
        din, dout = lv.d, lv.d2
        self._step(l, lv.x, lv.b, din, lv.x, c1, c2, 1)
        fused = False
        for k in range(1, self.deg):
            c1, c2 = lv.coef[k]
            last = (k == self.deg - 1) and part is not None
            mode = 3 | (4 if k == 1 else 0) | (8 if last else 0)
            self._step(l, din, lv.r, dout, lv.x, c1, c2, mode, dotv if last else None, part if last else None)
            fused = fused or last
            din, dout = dout, din
        if self.deg == 1:
            _vec_axpby_kernel[lv.blk_d, th](np.float32(1.0), lv.x, din)
        return fused

    def apply(self, d_r, d_z, part=None):
        """
        d_z = M^-1 d_r. With `part` (float64 array, one slot per level-0 block) the block partial sums of d_r . d_z are written to it
        (fused into the last smoothing kernel when deg >= 2, by a separate kernel otherwise).
        """
        th = self.threads
        Ls = self.levels
        nl = len(Ls)
        L0 = Ls[0]
        L0.b = d_r
        L0.x = d_z
        _gmg_cheb_kernel[L0.blk_d, th](L0.Nd, L0.Minv, L0.fixed, d_r, L0.d, d_z, np.float32(0.0), L0.coef[0][1], 3)
        for l in range(nl - 1):
            f, c = Ls[l], Ls[l + 1]
            self._pre(l)
            if l + 1 < nl - 1:
                _gf_restrict_init_kernel[c.blk_n, th](f.nx, f.ny, f.nz, c.nx, c.ny, c.nz, f.r, c.fixed, c.Minv, c.b, c.d, c.x, c.coef[0][1])
            else:
                _gmg_restrict_kernel[c.blk_n, th](f.nx, f.ny, f.nz, c.nx, c.ny, c.nz, f.r, c.fixed, c.b)
        cl = Ls[-1]
        _gmg_dense_apply_kernel[cl.blk_dense, th](cl.Nd, cl.dense, cl.b, cl.x)
        for l in range(nl - 2, -1, -1):
            f, c = Ls[l], Ls[l + 1]
            _gmg_prolong_add_kernel[f.blk_n, th](f.nx, f.ny, f.nz, c.nx, c.ny, c.nz, c.x, f.fixed, f.x)
            if l == 0 and part is not None:
                fused = self._post(0, d_r, part)
                if not fused:
                    _gf_dot_kernel[L0.blk_n, th](L0.Nn, d_r, d_z, part)
            else:
                self._post(l)
        L0.b = None
        L0.x = None

    def rebind(self, d_fixed0, d_E):
        """Re-attach the (stale-coarse-level) hierarchy to a new solve's E (mask unchanged): level-0 operator and block-Jacobi inverse."""
        L0 = self.levels[0]
        L0.fixed = d_fixed0
        self.d_E = d_E
        self._level0_minv()

    def info(self):
        nbytes = 0
        for l, lv in enumerate(self.levels):
            for nm in ("r", "d", "d2", "x", "b", "Kel", "dense") + (("Minv", "fixed") if l > 0 else ()):
                a = getattr(lv, nm, None)
                if a is not None and hasattr(a, "size"):
                    nbytes += int(a.size) * int(a.dtype.itemsize)
        return dict(levels=[(lv.nx - 1, lv.ny - 1, lv.nz - 1) for lv in self.levels],
                    lam_max=[getattr(lv, "lam_max", None) for lv in self.levels],
                    smoother_steps=self.deg, coarsest_solve="dense_inverse", extra_bytes=nbytes, engine="fast")


_DENSE_PROF = None     # set to a dict to collect stage timings of _dense_coarse_inverse (profiling only)


def _dense_coarse_inverse(Kel_h, fixed_h, lv):
    """Dense Jacobi-scaled inverse of the coarsest Galerkin operator (host float64) -> float32 flat; same definition as
    GMGPreconditioner._build_coarse_solver, evaluated faster: the element matrices are summed straight into the dense matrix of the FREE
    DOFs by one sparse COO assembly (no N x N intermediate, no np.add.at), the SPD inverse comes from one Cholesky + LAPACK potri
    (about 1 N^3 flops) instead of Cholesky + a general triangular solve with the identity + a product (about 3 N^3), and the scaling
    is applied in place."""
    from time import perf_counter as _pc
    _t = [_pc()]

    def _mark(nm):
        if _DENSE_PROF is not None:
            _DENSE_PROF[nm] = _pc() - _t[0]
            _t[0] = _pc()
    N = lv.Nd
    dm = hex_dof_map(lv.nx, lv.ny, lv.nz)
    fixed = fixed_h == 1
    free = np.where(~fixed)[0]
    nf = free.size
    try:
        import scipy.sparse as _sp
        import scipy.linalg as _sl
    except Exception:       # numpy-only fallback (slower)
        _sp = _sl = None
    fmap = np.full(N, -1, dtype=np.int64)
    fmap[free] = np.arange(nf)
    if _sp is not None:
        rr = np.broadcast_to(fmap[dm][:, :, None], Kel_h.shape).ravel()
        cc = np.broadcast_to(fmap[dm][:, None, :], Kel_h.shape).ravel()
        keep = (rr >= 0) & (cc >= 0)
        Aff = _sp.coo_matrix((Kel_h.ravel()[keep], (rr[keep], cc[keep])), shape=(nf, nf)).toarray()
    else:
        A = np.zeros((N, N))
        np.add.at(A, (dm[:, :, None], dm[:, None, :]), Kel_h)
        Aff = A[np.ix_(free, free)]
    _mark('assemble')
    Aff += Aff.T
    Aff *= 0.5
    dg = np.diag(Aff).copy()
    dg[dg <= 0.0] = 1.0
    sc = 1.0 / np.sqrt(dg)
    Aff *= sc[:, None]
    Aff *= sc[None, :]
    S = Aff
    _mark('scale')
    Sinv = None
    try:
        if _sl is not None:
            c, _low = _sl.cho_factor(S, lower=True, check_finite=False)
            Si, info = _sl.lapack.dpotri(c, lower=1, overwrite_c=True)
            if info == 0:
                Sinv = np.tril(Si)
                Sinv += np.tril(Si, -1).T
        else:
            Lc = np.linalg.cholesky(S)
            Li = np.linalg.solve(Lc, np.eye(nf))
            Sinv = Li.T @ Li
    except (np.linalg.LinAlgError, ValueError):
        Sinv = None
    _mark('chol+potri')
    if Sinv is None:
        pass
        w_, V_ = np.linalg.eigh(S)
        keep_ = w_ > 1e-12 * w_.max()
        Sinv = (V_[:, keep_] / w_[keep_]) @ V_[:, keep_].T
    Sinv *= sc[:, None]
    Sinv *= sc[None, :]
    Minv = np.zeros((N, N), dtype=np.float32)
    Minv[np.ix_(free, free)] = Sinv
    out = Minv.ravel()
    _mark('scatter+cast')
    return out


def _fast_engine_ok(nx, ny, nz, element_type, preconditioner, gmg_levels, gmg_coarse_max_nodes):
    """(True, "") when the fast engine can run this solve, else (False, reason): it needs preconditioner='gmg', a matrix element type and a coarsenable grid."""
    if preconditioner != "gmg":
        return False, f"pcg_engine='fast' needs preconditioner='gmg' (got {preconditioner!r})"
    if element_type == "q1_gauss":
        return False, "pcg_engine='fast' needs a matrix element type ('q1' or 'q1_im'), not 'q1_gauss'"
    try:
        gmg_level_dims(nx, ny, nz, gmg_levels, gmg_coarse_max_nodes)
    except ValueError:
        return False, f"grid {nx - 1}x{ny - 1}x{nz - 1} elements is too small to coarsen"
    return True, ""


def _pcg_fast(nx, ny, nz, hx, hy, hz, E_elem_h, nu_elem_h, b_h, fixed_mask_h, max_iters, tol, rel_tol, verbose, u0,
              warm_start_probe_iters, warm_start_min_reduction, element_type, nu, gmg_levels, gmg_smoother_steps,
              gmg_coarse_max_nodes, gmg_reuse_max_rel_change, check_every):
    """PCG with the fast engine (see the block comment at the top of this section); arguments as matrix_free_pcg."""
    from time import perf_counter
    _k = "gmg->gmg"
    PRECONDITIONER_USED[_k] = PRECONDITIONER_USED.get(_k, 0) + 1
    fixed_mask_h = np.asarray(fixed_mask_h).astype(np.int32)
    b_h = np.asarray(b_h, dtype=np.float32).copy()
    n_load_on_fixed = int(np.count_nonzero((fixed_mask_h == 1) & (b_h != 0.0)))
    if n_load_on_fixed:
        b_h[fixed_mask_h == 1] = 0.0
        if verbose:
            print(f"Note: removed load from {n_load_on_fixed} constrained DOF(s) "
                  f"(no work is done there; it would otherwise floor the residual).")
    stop_tol = max(float(tol), float(rel_tol) * float(np.linalg.norm(b_h)))
    if verbose:
        print(f"PCG (fast engine) stopping tolerance: {stop_tol:.6e} (absolute {tol:.3e}, relative {rel_tol:.3e} x ||b||)")
    Nn = nx * ny * nz
    Nd = 3 * Nn
    th = _FB
    nb0 = (Nn + th - 1) // th
    nbd = (Nd + th - 1) // th
    max_iters = int(max_iters)

    d_E = cuda.to_device(np.asarray(E_elem_h, dtype=np.float32))
    d_fixed = cuda.to_device(fixed_mask_h)
    d_Khat = _khat_to_device(element_type, hx, hy, hz, E_elem_h, nu_elem_h, nu)

    cuda.synchronize()
    _t0 = perf_counter()
    gmg = None
    reused = False
    rel_change = None
    _E32 = np.asarray(E_elem_h, dtype=np.float32)
    _key = ("fast", nx, ny, nz, element_type, gmg_levels, int(gmg_smoother_steps), int(gmg_coarse_max_nodes),
            None if nu is None else float(nu), float(hx), float(hy), float(hz))
    _ent = _GMG_CACHE.get("entry") if gmg_reuse_max_rel_change > 0.0 else None
    if _ent is not None and _ent["key"] == _key and np.array_equal(_ent["fixed"], fixed_mask_h):
        rel_change = float(np.max(np.abs(_E32 - _ent["E"])) / max(float(np.max(_ent["E"])), 1e-30))
        if rel_change <= gmg_reuse_max_rel_change:
            gmg = _ent["gmg"]
            gmg.rebind(d_fixed, d_E)
            reused = True
    if gmg is None:
        gmg = GMGFastPreconditioner(nx, ny, nz, d_fixed, d_E, d_Khat, levels=gmg_levels, smoother_steps=gmg_smoother_steps,
                                    coarse_max_nodes=gmg_coarse_max_nodes, threads=th)
        if gmg_reuse_max_rel_change > 0.0:
            _GMG_CACHE["entry"] = dict(key=_key, fixed=np.array(fixed_mask_h, copy=True), E=_E32.copy(), gmg=gmg)
        else:
            _GMG_CACHE.clear()
    cuda.synchronize()
    setup_s = perf_counter() - _t0

    d_x = cuda.device_array(Nd, dtype=np.float32)
    d_r = cuda.device_array(Nd, dtype=np.float32)
    d_p = cuda.device_array(Nd, dtype=np.float32)
    d_z = cuda.device_array(Nd, dtype=np.float32)
    d_Ap = cuda.device_array(Nd, dtype=np.float32)
    d_b = cuda.to_device(b_h)
    part = cuda.device_array(nb0, dtype=np.float64)
    hist = cuda.device_array(max(max_iters, 1), dtype=np.float64)
    sc = cuda.to_device(np.array([0, 0, 0, 0, 0, 0, 0, stop_tol * stop_tol, float(max_iters), 0], dtype=np.float64))
    L0 = gmg.levels[0]
    warm_started = u0 is not None
    bailed_to_cold = False
    bailout_iter = None
    r0_norm_warm = None

    def _cold_start():
        # u = 0, r = b, z = M^-1 r, p = z, rz = r.z
        _zero_array_kernel[nbd, th](d_x)
        _vec_copy_kernel[nbd, th](d_r, d_b)
        gmg.apply(d_r, d_z, part)
        _gf_scalar_kernel[(1,), _GF_RB](part, nb0, sc, hist, 3)
        _vec_copy_kernel[nbd, th](d_p, d_z)

    cuda.synchronize()
    _t_loop0 = perf_counter()
    if warm_started:
        u0_h = np.asarray(u0, dtype=np.float32).copy()
        if u0_h.shape[0] != Nd:
            raise ValueError(f"u0 has {u0_h.shape[0]} DOFs, expected {Nd}.")
        u0_h[fixed_mask_h == 1] = 0.0
        d_x.copy_to_device(u0_h)
        # r = b - A u0
        gmg._step(0, d_x, d_b, d_x, d_x, np.float32(0.0), np.float32(0.0), 0)       # writes r into level-0 buffer lv.r
        _vec_copy_kernel[nbd, th](d_r, L0.r)
        _gf_dot_kernel[nb0, th](Nn, d_r, d_r, part)
        _gf_scalar_kernel[(1,), _GF_RB](part, nb0, sc, hist, 3)
        r0_norm_warm = sqrt(max(float(sc.copy_to_host()[0]), 0.0))
        if verbose:
            print(f"Warm start: ||r0|| = {r0_norm_warm:.6e}")
        gmg.apply(d_r, d_z, part)
        _gf_scalar_kernel[(1,), _GF_RB](part, nb0, sc, hist, 3)
        _vec_copy_kernel[nbd, th](d_p, d_z)
    else:
        _cold_start()

    def _iteration():
        _gf_apply_dot_kernel[nb0, th](nx, ny, nz, d_E, d_E, d_Khat, 0, d_p, d_fixed, d_Ap, part)
        _gf_scalar_kernel[(1,), _GF_RB](part, nb0, sc, hist, 0)
        _gf_update_kernel[nb0, th](Nn, d_x, d_r, d_p, d_Ap, sc, part)
        _gf_scalar_kernel[(1,), _GF_RB](part, nb0, sc, hist, 1)
        gmg.apply(d_r, d_z, part)
        _gf_scalar_kernel[(1,), _GF_RB](part, nb0, sc, hist, 2)
        _gf_p_kernel[nb0, th](Nn, d_z, d_p, sc)

    n_reads = 0
    n_enq = 0
    batch = int(check_every) if check_every else (int(warm_start_probe_iters) if warm_started else 4)
    batch = max(1, batch)
    probe_pending = warm_started
    auto = not check_every
    count = 0
    while True:
        if probe_pending:
            batch = max(1, min(batch, int(warm_start_probe_iters) - count))
        for _ in range(batch):
            _iteration()
        n_enq += batch
        s = sc.copy_to_host()
        n_reads += 1
        count = int(s[5])
        flag = s[6]
        if probe_pending and count == int(warm_start_probe_iters):
            probe_pending = False
            if flag == 0:
                rn = sqrt(float(hist[count - 1:count].copy_to_host()[0]))
                required = (1.0 - warm_start_min_reduction) * max(r0_norm_warm, 1e-30)
                if rn > required:
                    bailed_to_cold = True
                    bailout_iter = count - 1
                    if verbose:
                        print(f"Warm start not converging fast enough (||r||={rn:.6e} > {required:.6e} required by iter "
                              f"{count - 1}); restarting cold.")
                    _cold_start()
                    continue
        if flag != 0:
            break
        if auto:
            batch = _next_batch(hist, count, float(s[7]))
    cuda.synchronize()
    cg_loop_s = perf_counter() - _t_loop0
    u_h = d_x.copy_to_host()
    s = sc.copy_to_host()
    count = int(s[5])
    h2 = hist[:count].copy_to_host() if count > 0 else np.zeros(0)
    history = [float(sqrt(v)) for v in h2]
    # exactness re-check of the stored residual history: no entry before the last is below the stop, the last is (unless max_iters / breakdown)
    stop2 = float(s[7])
    exact = True
    if count > 0 and s[6] == 1.0 and count < max_iters:
        exact = bool(h2[-1] < stop2 and (count == 1 or not np.any(h2[:-1] < stop2)))
    info = {"warm_started": warm_started, "bailed_to_cold": bailed_to_cold, "bailout_iter": bailout_iter,
            "preconditioner_used": "gmg", "pcg_engine_used": "fast"}
    info["gmg"] = dict(gmg.info(), setup_s=setup_s, reused=reused, rel_change=rel_change, cg_loop_s=cg_loop_s)
    info["fast"] = dict(host_reads=n_reads, enqueued_iters=n_enq, wasted_iters=max(n_enq - count, 0), stop_exact=exact,
                        breakdown=bool(s[6] == 2.0))
    return u_h, history, info


def _try_jax(requested, nx, ny, nz, hx, hy, hz, E_elem_h, nu_elem_h, b_h, fixed_mask_h, max_iters, tol, rel_tol, verbose, u0,
             warm_start_probe_iters, warm_start_min_reduction, element_type, nu, preconditioner, gmg_levels, gmg_smoother_steps,
             gmg_coarse_max_nodes):
    """pcg_engine "jax" / "auto" (docs/theory.md §17): ((u_h, history, info), "") from gpu_jax.py, or (None, reason) when the
    solve should run the fast engine instead (engines.pcg_choice declined: no JAX / no CUDA JAX / memory / not faster; or the JAX
    solve raised, e.g. out of memory, which is recorded in engines.RECORD["fallbacks"] and disables jax for that grid)."""
    eng, variant, form, why = pcg_choice(requested, nx, ny, nz, hx, hy, hz, element_type, preconditioner, gmg_levels,
                                                 gmg_coarse_max_nodes)
    if eng != "jax":
        return None, why
    try:
        _res = pcg(nx, ny, nz, hx, hy, hz, E_elem_h, nu_elem_h, b_h, fixed_mask_h, max_iters=max_iters, tol=tol, rel_tol=rel_tol,
                           verbose=verbose, u0=u0, warm_start_probe_iters=warm_start_probe_iters,
                           warm_start_min_reduction=warm_start_min_reduction, element_type=element_type, nu=nu, gmg_levels=gmg_levels,
                           gmg_smoother_steps=gmg_smoother_steps, gmg_coarse_max_nodes=gmg_coarse_max_nodes, variant=variant, form=form)
        record_use("pcg", "jax")
        return _res, ""
    except Exception as e:  # noqa: BLE001
        FAILED.add(("pcg_jax", nx, ny, nz))
        record_fallback("pcg", "jax", "fast", f"{type(e).__name__}: {e}")
        print(f"engines: pcg jax failed on grid {nx - 1}x{ny - 1}x{nz - 1} ({type(e).__name__}: {str(e)[:200]}); using the fast engine.", flush=True)
        return None, f"jax failed: {type(e).__name__}"


def _next_batch(hist, count, stop2):
    """Iterations to enqueue before the next host read: predicted from the geometric decay of the stored residual history (a bit short of it, so that
    iterations enqueued after convergence stay rare); at least 1, at most 16."""
    if count < 2:
        return 2
    m = min(count - 1, 3)
    h2 = hist[count - 1 - m:count].copy_to_host()
    if not (h2[0] > 0.0 and h2[-1] > 0.0):
        return 2
    rate = (h2[-1] / h2[0]) ** (1.0 / m)          # per-iteration decay of ||r||^2
    if not (rate < 1.0):
        return 4
    remaining = np.log(stop2 / h2[-1]) / np.log(rate)
    return int(max(1, min(16, np.floor(0.8 * remaining))))


def _khat_to_device(element_type, hx, hy, hz, E_elem_h, nu_elem_h, nu=None):
    """Unit-E 24x24 element matrix (float64 on the host -> float32 flat device array)."""
    if element_type not in ELEMENT_TYPES:
        raise ValueError(f"element_type must be one of {ELEMENT_TYPES}, got {element_type!r}")
    nu0 = resolve_nu(nu_elem_h, E_elem_h, nu)
    Kh = element_khat(element_type, (float(hx), float(hy), float(hz)), nu0)
    return cuda.to_device(np.ascontiguousarray(Kh, dtype=np.float32).ravel())


def compute_element_strain_energy(nx, ny, nz, hx, hy, hz, E_elem_h, nu_elem_h, u_h,
                                  element_type="q1_gauss", nu=None):
    nx_elems = nx - 1; ny_elems = ny - 1; nz_elems = nz - 1
    Ne = nx_elems * ny_elems * nz_elems
    
    # 1. Setup device arrays
    d_u = cuda.to_device(u_h.astype(np.float32))
    d_E = cuda.to_device(E_elem_h.astype(np.float32))
    d_nu = cuda.to_device(nu_elem_h.astype(np.float32))
    d_energy = cuda.device_array(Ne, dtype=np.float32)

    # 2. Configure launch
    threads = 128
    blocks = (Ne + threads - 1) // threads

    # 3. Launch
    if element_type == "q1_gauss":
        _element_energy_kernel[blocks, threads](nx, ny, nz, hx, hy, hz, d_E, d_nu, d_u, d_energy)
    else:
        d_Khat = _khat_to_device(element_type, hx, hy, hz, E_elem_h, nu_elem_h, nu)
        _element_energy_kernel_mat[blocks, threads](nx, ny, nz, d_E, d_Khat, d_u, d_energy)
    
    # 4. Return results
    return d_energy.copy_to_host()


# =========================================================================
# --- HOST-SIDE REFERENCE-ELEMENT DATA (shared with analysis.py) ---
#
# analysis.py's linear buckling analysis needs the same trilinear hex
# shape-function data that the CUDA kernels above use, but on the host and
# in float64 (an eigensolver in single precision is not viable). Rather
# than let analysis.py keep its own transcription of the formulas -- which
# would silently drift the moment anyone edited `_compute_B_at_gauss` or
# `_compute_C_voigt` -- the canonical host-side definitions live HERE,
# next to the kernels they must agree with, and analysis.py imports them.
#
# If you change the element formulation in the kernels above, change it
# here too, then run `verify_host_kernel_consistency()` (bottom of this
# file) on a machine with a GPU to confirm the two still agree.
# =========================================================================

# Voigt ordering used by _compute_C_voigt / _compute_B_at_gauss above:
#   (xx, yy, zz, yz, xz, xy)
# with ENGINEERING shear strains (rows 3-5 of B carry both cross terms,
# so B @ u gives gamma_yz = 2*eps_yz, and C[3,3] = mu maps that to
# sigma_yz correctly).
VOIGT_C_LAMBDA_BASIS = np.zeros((6, 6), dtype=np.float64)
VOIGT_C_LAMBDA_BASIS[:3, :3] = 1.0
VOIGT_C_MU_BASIS = np.diag([2.0, 2.0, 2.0, 1.0, 1.0, 1.0]).astype(np.float64)

# 2-point Gauss rule, matching the kernels' `g = 0.5773502691896257`.
GAUSS_POINT_1D = 1.0 / sqrt(3.0)


def lame_from_E_nu(E, nu):
    """
    Host-side (vectorised, float64) equivalent of the device function
    `_compute_C_voigt`'s Lame conversion. Kept bit-for-bit consistent with
    it, including the degenerate-denominator guard.
    """
    E = np.asarray(E, dtype=np.float64)
    nu = np.asarray(nu, dtype=np.float64)
    denom = (1.0 + nu) * (1.0 - 2.0 * nu)
    denom = np.where(np.abs(denom) < 1e-12, 1e-12, denom)
    lam = E * nu / denom
    mu = E / (2.0 * (1.0 + nu))
    return lam, mu


def hex_reference_data(hx, hy, hz):
    """
    Host-side (float64) shape-function data for the regular, axis-aligned
    trilinear hex element used by the matrix-free kernels above.

    Mirrors `_compute_B_at_gauss` exactly: same corner ordering (the
    module-level `xi_arr`/`eta_arr`/`zeta_arr` are reused directly rather
    than re-typed, so they cannot drift), same shape functions, and the
    same 2x2x2 Gauss rule traversed in the same (xi, eta, zeta) nesting
    order used by `_element_energy_kernel`.

    Args:
    - hx, hy, hz (float): element edge lengths in metres.

    Returns:
    - B_all (8, 6, 24): Voigt strain-displacement matrix per Gauss point.
    - Bg_all (8, 3, 8): raw shape-function gradients dN_a/dx_i per Gauss
      point. Needed for the geometric stiffness, which couples spatial
      gradients directly rather than through the Voigt strain.
    - detJ (float): Jacobian determinant, constant for a uniform hex.
    """
    inv_j = np.array([2.0 / hx, 2.0 / hy, 2.0 / hz], dtype=np.float64)
    gauss_vals = (-GAUSS_POINT_1D, GAUSS_POINT_1D)

    # Reuse the kernels' own corner arrays (promoted to float64) so the
    # host and device definitions cannot diverge.
    xi_c = xi_arr.astype(np.float64)
    eta_c = eta_arr.astype(np.float64)
    zeta_c = zeta_arr.astype(np.float64)

    B_all = np.zeros((8, _B_ROWS, _B_COLS), dtype=np.float64)
    Bg_all = np.zeros((8, 3, 8), dtype=np.float64)

    gp = 0
    for gp_xi in gauss_vals:
        for gp_eta in gauss_vals:
            for gp_zeta in gauss_vals:
                dN_dxi = 0.125 * xi_c * (1.0 + eta_c * gp_eta) * (1.0 + zeta_c * gp_zeta)
                dN_deta = 0.125 * eta_c * (1.0 + xi_c * gp_xi) * (1.0 + zeta_c * gp_zeta)
                dN_dzeta = 0.125 * zeta_c * (1.0 + xi_c * gp_xi) * (1.0 + eta_c * gp_eta)

                dN_dx = dN_dxi * inv_j[0]
                dN_dy = dN_deta * inv_j[1]
                dN_dz = dN_dzeta * inv_j[2]

                Bg_all[gp, 0, :] = dN_dx
                Bg_all[gp, 1, :] = dN_dy
                Bg_all[gp, 2, :] = dN_dz

                for a in range(8):
                    col = 3 * a
                    B_all[gp, 0, col + 0] = dN_dx[a]
                    B_all[gp, 1, col + 1] = dN_dy[a]
                    B_all[gp, 2, col + 2] = dN_dz[a]
                    B_all[gp, 3, col + 1] = dN_dz[a]
                    B_all[gp, 3, col + 2] = dN_dy[a]
                    B_all[gp, 4, col + 0] = dN_dz[a]
                    B_all[gp, 4, col + 2] = dN_dx[a]
                    B_all[gp, 5, col + 0] = dN_dy[a]
                    B_all[gp, 5, col + 1] = dN_dx[a]

                gp += 1

    detJ = (hx * hy * hz) * 0.125
    return B_all, Bg_all, float(detJ)


def hex_dof_map(nx, ny, nz):
    """
    Vectorised element -> global DOF map for the regular grid, matching
    the node-numbering convention hard-coded in the kernels above:
    C-order numbering with z varying fastest, nx/ny/nz are NODE counts,
    and the 24 local DOFs are ordered [ux0, uy0, uz0, ux1, uy1, uz1, ...].

    This is the single definition of that convention for host-side code;
    it is derived from the same n0..n7 expressions used in
    `_element_action_kernel` and `_element_energy_kernel`.

    Returns:
    - dof_map (Ne, 24) int64 array of global DOF indices, with elements in
      the same C-order (z fastest) that `E_elem`/`nu_elem` use.
    """
    nx_e, ny_e, nz_e = nx - 1, ny - 1, nz - 1
    ei, ej, ek = np.meshgrid(np.arange(nx_e), np.arange(ny_e), np.arange(nz_e),
                             indexing='ij')
    ei = ei.ravel(); ej = ej.ravel(); ek = ek.ravel()

    n0 = ei * (ny * nz) + ej * nz + ek
    n1 = (ei + 1) * (ny * nz) + ej * nz + ek
    n2 = (ei + 1) * (ny * nz) + (ej + 1) * nz + ek
    n3 = ei * (ny * nz) + (ej + 1) * nz + ek
    n4 = ei * (ny * nz) + ej * nz + (ek + 1)
    n5 = (ei + 1) * (ny * nz) + ej * nz + (ek + 1)
    n6 = (ei + 1) * (ny * nz) + (ej + 1) * nz + (ek + 1)
    n7 = ei * (ny * nz) + (ej + 1) * nz + (ek + 1)

    nodes = np.stack([n0, n1, n2, n3, n4, n5, n6, n7], axis=1)
    dof_map = (3 * nodes[:, :, None] + np.arange(3)[None, None, :])
    return dof_map.reshape(nodes.shape[0], 24).astype(np.int64)


@cuda.jit(fastmath=True)
def _element_stress_kernel(nx, ny, nz, hx, hy, hz, E_elem, nu_elem, u_global, stress_out):
    """
    Computes the element-averaged Cauchy stress in Voigt form
    (xx, yy, zz, yz, xz, xy) for every element, writing 6 floats per
    element into `stress_out` (shape (Ne, 6)).

    Structurally identical to `_element_energy_kernel` -- same node
    gathering, same 2x2x2 Gauss loop, same `_compute_B_at_gauss` and
    `_compute_C_voigt` calls -- but it keeps the stress that the energy
    kernel computes internally and then discards, instead of contracting
    it against the strain. Averaged over the 8 Gauss points, which for a
    trilinear hex under a linear stress field is the element centroid
    value.

    This exists because the linear buckling analysis in analysis.py needs
    the pre-buckling stress state to build the geometric stiffness matrix
    K_sigma; see section 2.1 of the buckling implementation plan.
    """
    tid = cuda.threadIdx.x
    bidx = cuda.blockIdx.x
    bdim = cuda.blockDim.x

    nx_elems = nx - 1
    ny_elems = ny - 1
    nz_elems = nz - 1
    Ne = nx_elems * ny_elems * nz_elems

    e_global = bidx * bdim + tid
    if e_global >= Ne:
        return

    ek = e_global % nz_elems
    ej = (e_global // nz_elems) % ny_elems
    ei = e_global // (ny_elems * nz_elems)

    n0 = (ei    ) * (ny * nz) + (ej    ) * nz + (ek    )
    n1 = (ei + 1) * (ny * nz) + (ej    ) * nz + (ek    )
    n2 = (ei + 1) * (ny * nz) + (ej + 1) * nz + (ek    )
    n3 = (ei    ) * (ny * nz) + (ej + 1) * nz + (ek    )
    n4 = (ei    ) * (ny * nz) + (ej    ) * nz + (ek + 1)
    n5 = (ei + 1) * (ny * nz) + (ej    ) * nz + (ek + 1)
    n6 = (ei + 1) * (ny * nz) + (ej + 1) * nz + (ek + 1)
    n7 = (ei    ) * (ny * nz) + (ej + 1) * nz + (ek + 1)

    u_e = cuda.local.array(_PE_SIZE, float32)
    nodes_local = cuda.local.array(8, int32)
    nodes_local[0]=n0; nodes_local[1]=n1; nodes_local[2]=n2; nodes_local[3]=n3
    nodes_local[4]=n4; nodes_local[5]=n5; nodes_local[6]=n6; nodes_local[7]=n7

    for a in range(8):
        nid = nodes_local[a]
        base = 3*nid
        u_e[3*a + 0] = u_global[base + 0]
        u_e[3*a + 1] = u_global[base + 1]
        u_e[3*a + 2] = u_global[base + 2]

    Cloc = cuda.local.array(36, float32)
    _compute_C_voigt(E_elem[e_global], nu_elem[e_global], Cloc)

    acc0 = 0.0; acc1 = 0.0; acc2 = 0.0; acc3 = 0.0; acc4 = 0.0; acc5 = 0.0

    g = 0.5773502691896257
    gauss_vals0 = (-g, g)

    for gi in range(2):
        gp_xi = gauss_vals0[gi]
        for gj in range(2):
            gp_eta = gauss_vals0[gj]
            for gk in range(2):
                gp_zeta = gauss_vals0[gk]

                B = cuda.local.array(_B_SIZE, float32)
                _compute_B_at_gauss(hx, hy, hz, gp_xi, gp_eta, gp_zeta, B)

                eps0 = 0.0; eps1 = 0.0; eps2 = 0.0; eps3 = 0.0; eps4 = 0.0; eps5 = 0.0
                for jj in range(_B_COLS):
                    v = u_e[jj]
                    eps0 += B[0*_B_COLS + jj] * v
                    eps1 += B[1*_B_COLS + jj] * v
                    eps2 += B[2*_B_COLS + jj] * v
                    eps3 += B[3*_B_COLS + jj] * v
                    eps4 += B[4*_B_COLS + jj] * v
                    eps5 += B[5*_B_COLS + jj] * v

                acc0 += Cloc[0*6+0]*eps0 + Cloc[0*6+1]*eps1 + Cloc[0*6+2]*eps2 + Cloc[0*6+3]*eps3 + Cloc[0*6+4]*eps4 + Cloc[0*6+5]*eps5
                acc1 += Cloc[1*6+0]*eps0 + Cloc[1*6+1]*eps1 + Cloc[1*6+2]*eps2 + Cloc[1*6+3]*eps3 + Cloc[1*6+4]*eps4 + Cloc[1*6+5]*eps5
                acc2 += Cloc[2*6+0]*eps0 + Cloc[2*6+1]*eps1 + Cloc[2*6+2]*eps2 + Cloc[2*6+3]*eps3 + Cloc[2*6+4]*eps4 + Cloc[2*6+5]*eps5
                acc3 += Cloc[3*6+0]*eps0 + Cloc[3*6+1]*eps1 + Cloc[3*6+2]*eps2 + Cloc[3*6+3]*eps3 + Cloc[3*6+4]*eps4 + Cloc[3*6+5]*eps5
                acc4 += Cloc[4*6+0]*eps0 + Cloc[4*6+1]*eps1 + Cloc[4*6+2]*eps2 + Cloc[4*6+3]*eps3 + Cloc[4*6+4]*eps4 + Cloc[4*6+5]*eps5
                acc5 += Cloc[5*6+0]*eps0 + Cloc[5*6+1]*eps1 + Cloc[5*6+2]*eps2 + Cloc[5*6+3]*eps3 + Cloc[5*6+4]*eps4 + Cloc[5*6+5]*eps5

    inv8 = 0.125
    stress_out[e_global, 0] = acc0 * inv8
    stress_out[e_global, 1] = acc1 * inv8
    stress_out[e_global, 2] = acc2 * inv8
    stress_out[e_global, 3] = acc3 * inv8
    stress_out[e_global, 4] = acc4 * inv8
    stress_out[e_global, 5] = acc5 * inv8


def compute_element_stress(nx, ny, nz, hx, hy, hz, E_elem_h, nu_elem_h, u_h):
    """
    Host driver for `_element_stress_kernel`. Mirrors
    `compute_element_strain_energy`'s calling convention exactly.

    Returns:
    - (Ne, 6) float32 array of element-averaged Voigt stresses
      (xx, yy, zz, yz, xz, xy), in the same C-order (z fastest) element
      ordering as `E_elem_h` / `nu_elem_h`.
    """
    nx_elems = nx - 1; ny_elems = ny - 1; nz_elems = nz - 1
    Ne = nx_elems * ny_elems * nz_elems

    d_u = cuda.to_device(u_h.astype(np.float32))
    d_E = cuda.to_device(E_elem_h.astype(np.float32))
    d_nu = cuda.to_device(nu_elem_h.astype(np.float32))
    d_stress = cuda.device_array((Ne, 6), dtype=np.float32)

    threads = 128
    blocks = (Ne + threads - 1) // threads

    _element_stress_kernel[blocks, threads](nx, ny, nz, hx, hy, hz, d_E, d_nu, d_u, d_stress)

    return d_stress.copy_to_host()


def verify_host_kernel_consistency(nx_e=4, ny_e=3, nz_e=5, voxel_size_m=1e-3,
                                    E=210.0e9, nu=0.3, tol=2e-4, verbose=True):
    """
    Self-test that the host-side reference-element data above agrees with
    what the CUDA kernels actually compute. REQUIRES A GPU.

    This is the guard against the one real failure mode of splitting the
    element formulation across two implementations: someone edits
    `_compute_B_at_gauss`, `_compute_C_voigt` or the n0..n7 node
    expressions in a kernel, and `hex_reference_data` / `hex_dof_map`
    silently keep the old formulation. Nothing would raise; the buckling
    eigenvalues would just quietly become wrong.

    Three things are checked against a random displacement field on a
    fully-solid block:

      1. `hex_dof_map` + `hex_reference_data` reproduce the kernel's
         per-element strain energy (`compute_element_strain_energy`).
      2. The same reproduce the kernel's element stress
         (`compute_element_stress`).
      3. A host-assembled sparse K reproduces the kernel's matrix-free
         action K@u (`_element_action_kernel` via one PCG matvec).

    Tolerance defaults to 2e-4 relative, which is loose because the
    kernels run in float32 while the host path runs in float64; the two
    should not agree more closely than single precision allows.

    Returns True if every check passes, and raises AssertionError with a
    diagnostic if any fails.
    """
    from scipy.sparse import coo_matrix

    rng = np.random.default_rng(0)
    nx, ny, nz = nx_e + 1, ny_e + 1, nz_e + 1
    Ne = nx_e * ny_e * nz_e
    Nnodes = nx * ny * nz
    Ndofs = 3 * Nnodes
    hx = hy = hz = np.float32(voxel_size_m)

    rho = rng.uniform(0.2, 1.0, size=Ne).astype(np.float32)
    lam_s = E * nu / ((1.0 + nu) * (1.0 - 2.0 * nu))
    mu_s = E / (2.0 * (1.0 + nu))
    lam_map, mu_map = rho * lam_s, rho * mu_s
    den = np.where(np.abs(lam_map + mu_map) < 1e-12, 1e-12, lam_map + mu_map)
    E_elem = (mu_map * (3.0 * lam_map + 2.0 * mu_map) / den).astype(np.float32)
    nu_elem = (lam_map / (2.0 * den)).astype(np.float32)

    # Small displacements keep float32 round-off from dominating.
    u_h = (rng.standard_normal(Ndofs) * 1e-9).astype(np.float32)

    dof_map = hex_dof_map(nx, ny, nz)
    B_all, Bg_all, detJ = hex_reference_data(float(hx), float(hy), float(hz))
    lam_e, mu_e = lame_from_E_nu(E_elem, nu_elem)

    u64 = u_h.astype(np.float64)
    u_e = u64[dof_map]                                     # (Ne, 24)
    strain = np.einsum('gvc,ec->egv', B_all, u_e)          # (Ne, 8, 6)
    s_lam = np.einsum('vw,egw->egv', VOIGT_C_LAMBDA_BASIS, strain)
    s_mu = np.einsum('vw,egw->egv', VOIGT_C_MU_BASIS, strain)
    stress = lam_e[:, None, None] * s_lam + mu_e[:, None, None] * s_mu

    results = {}

    # --- Check 1: strain energy ---
    host_energy = 0.5 * detJ * np.sum(np.einsum('egv,egv->eg', stress, strain), axis=1)
    gpu_energy = compute_element_strain_energy(nx, ny, nz, hx, hy, hz,
                                                E_elem, nu_elem, u_h).astype(np.float64)
    scale = max(np.abs(gpu_energy).max(), 1e-30)
    results['strain_energy'] = np.abs(host_energy - gpu_energy).max() / scale

    # --- Check 2: element stress ---
    host_stress = np.mean(stress, axis=1)
    gpu_stress = compute_element_stress(nx, ny, nz, hx, hy, hz,
                                         E_elem, nu_elem, u_h).astype(np.float64)
    scale = max(np.abs(gpu_stress).max(), 1e-30)
    results['element_stress'] = np.abs(host_stress - gpu_stress).max() / scale

    # --- Check 3: matrix-free K action vs host-assembled K ---
    Ke_lam = np.einsum('gvc,vw,gwd->cd', B_all, VOIGT_C_LAMBDA_BASIS, B_all) * detJ
    Ke_mu = np.einsum('gvc,vw,gwd->cd', B_all, VOIGT_C_MU_BASIS, B_all) * detJ
    Ke_all = lam_e[:, None, None] * Ke_lam[None] + mu_e[:, None, None] * Ke_mu[None]
    rows = np.repeat(dof_map[:, :, None], 24, axis=2)
    cols = np.repeat(dof_map[:, None, :], 24, axis=1)
    K_host = coo_matrix((Ke_all.ravel(), (rows.ravel(), cols.ravel())),
                        shape=(Ndofs, Ndofs)).tocsr()
    host_Ku = K_host @ u64

    d_u = cuda.to_device(u_h)
    d_E = cuda.to_device(E_elem)
    d_nu = cuda.to_device(nu_elem)
    d_Ku = cuda.device_array(Ndofs, dtype=np.float32)
    threads = THREADS_PER_BLOCK
    blocks_dofs = (Ndofs + threads - 1) // threads
    nblk_i = (nx_e + ELEMS_PER_BLOCK_I - 1) // ELEMS_PER_BLOCK_I
    nblk_j = (ny_e + ELEMS_PER_BLOCK_J - 1) // ELEMS_PER_BLOCK_J
    nblk_k = (nz_e + ELEMS_PER_BLOCK_K - 1) // ELEMS_PER_BLOCK_K
    blocks_elems = nblk_i * nblk_j * nblk_k
    _zero_array_kernel[blocks_dofs, threads](d_Ku)
    _element_action_kernel[blocks_elems, threads](
        nx, ny, nz, hx, hy, hz, d_E, d_nu, d_u, d_Ku,
        ELEMS_PER_BLOCK_I, ELEMS_PER_BLOCK_J, ELEMS_PER_BLOCK_K)
    gpu_Ku = d_Ku.copy_to_host().astype(np.float64)
    scale = max(np.abs(gpu_Ku).max(), 1e-30)
    results['stiffness_action'] = np.abs(host_Ku - gpu_Ku).max() / scale

    if verbose:
        print("Host <-> kernel consistency check "
              f"({nx_e}x{ny_e}x{nz_e} elements, tol={tol:.1e}):")
        for name, err in results.items():
            print(f"  {name:20s} max relative error {err:.3e}  "
                  f"{'PASS' if err < tol else 'FAIL'}")

    failures = {k: v for k, v in results.items() if not (v < tol)}
    if failures:
        raise AssertionError(
            "gpu_CG.py host-side reference data disagrees with the CUDA "
            f"kernels: {failures}. The host formulation in "
            "hex_reference_data/hex_dof_map/lame_from_E_nu has drifted from "
            "_compute_B_at_gauss/_compute_C_voigt/the kernel node numbering. "
            "analysis.py's buckling results are NOT trustworthy until this "
            "is reconciled."
        )
    return True


# ======================================================================================================================
# === former gpu_jax.py ===
# ======================================================================================================================
"""
gpu_jax.py -- JAX/XLA geometric-multigrid PCG for the elastic analysis (pcg_engine="jax"; docs/theory.md §17).

The same algorithm as the numba engines of gpu_CG.py (node-centred 2:1 coarsening with void-padded odd counts, Galerkin
coarse operators from coarse ELEMENT matrices, block-Jacobi-scaled Chebyshev smoother, dense coarsest inverse, hash-seeded
power iteration for lambda_max), written as array operations so that the whole solve -- hierarchy setup (one jit) and the PCG
loop with the V-cycle (a second jit, `lax.while_loop`) -- runs on the device with no host round trip until the end.
Port of bench/topopt_methods/jaxpcg.py + jaxgmg.py (benchmarked there), generalised to the production element matrices
(q1 and q1_im: the 24x24 Khat of element_matrices.py, permuted to the corner order of this module).

Everything JAX is imported lazily: `import gpu_jax` never imports jax, so the code runs where JAX is absent (Windows).
jax_enable_x64 is switched on at the first use (process-wide; nothing else in this code base uses JAX).
XLA_PYTHON_CLIENT_PREALLOCATE defaults to false and XLA_PYTHON_CLIENT_ALLOCATOR to "platform" (buffers are allocated and freed exactly, so
the memory is returned to numba-CUDA between solves; the BFC pool kept GiB-granular regions and 1.4x the memory at the same speed). Compiled executables are cached on
disk (PCLS_JAX_CACHE_DIR, default ~/.cache/pcls_jax; set it empty to disable): the compile of a grid takes ~7 s, a cache hit ~0.5 s.

Precision variants (`variant`):
  "fp64"   all float64 (reference quality, 2x the bytes of fp32);
  "fp32"   all float32 with precision=HIGHEST (the default matmul precision of float32 on H100/Ampere+ is TF32: NEVER use it,
           the recursive residual converges to a displacement with an 84 % compliance error);
  "mixed"  outer PCG recursion (x, r, p, the level-0 operator in the CG step) in float64, the V-cycle (setup + application)
           in float32: float64 residual accuracy at nearly float32 cost.
Level-0 operator form (`form`): "slice" (gather by 8 shifted slices + batched matmul, never builds the 27x3x3 stencil of the
finest level) or "stencil" (assembled 27-point stencil; more memory, fewer flops). A third form, the E-weighted stencil evaluated on the fly
as 64 unrolled 3x3 block products in one fused loop (no (Ne,24) arrays), was 1.4x slower than "slice" in fp32 (matmul runs on tensor cores)
and is not kept (docs/theory.md §17).

Public API: available(), ok(...), pcg(...) (same arguments/returns as gpu_CG.matrix_free_pcg), memory_estimate(...),
to_numba_device(x) (zero-copy hand-off of a JAX device array to numba-CUDA), CACHE.
"""
import itertools
import os
import time

import numpy as np

VARIANTS = ("fp64", "fp32", "mixed")
FORMS = ("slice", "stencil")
DEFAULT_VARIANT = "mixed"
DEFAULT_FORM = "slice"

CORNERS = [(a, b, c) for a in (0, 1) for b in (0, 1) for c in (0, 1)]
OFFS = [(a, b, c) for a in (-1, 0, 1) for b in (-1, 0, 1) for c in (-1, 0, 1)]
OFF_INDEX = {o: i for i, o in enumerate(OFFS)}
HI_FACTOR, CHEB_RATIO, POWER_ITERS = 1.1, 6.0, 12      # as gpu_CG.GMGPreconditioner (1.1 lambda_max, interval [hi/6, hi])

CACHE = {}           # compiled solvers by key (shapes are fixed over an optimisation run: compiled once per grid/variant)
LAST = {}            # timing of the most recent call (profilers read it)
PHASE_HOOK = [None]  # called (no args) when the PCG loop starts, i.e. after the hierarchy setup (tools/profile_run.py)
_J = {}


# ============================================================================ lazy JAX
def _jax():
    """(jax, jnp), imported on first use; raises ImportError when JAX is absent."""
    if "jax" not in _J:
        os.environ.setdefault("XLA_PYTHON_CLIENT_PREALLOCATE", "false")
        os.environ.setdefault("XLA_PYTHON_CLIENT_ALLOCATOR", "platform")     # exact alloc/free: no GiB-granular pool kept next to numba; same speed (measured)
        import jax
        import jax.numpy as jnp
        jax.config.update("jax_enable_x64", True)
        cdir = os.environ.get("PCLS_JAX_CACHE_DIR", os.path.join(os.environ.get("XDG_CACHE_HOME") or os.path.expanduser("~/.cache"), "pcls_jax"))
        if cdir:                                  # persistent compilation cache: the ~7 s compile of a grid is paid once per machine, not per process
            try:
                os.makedirs(cdir, exist_ok=True)
                jax.config.update("jax_compilation_cache_dir", cdir)
            except Exception:  # noqa: BLE001
                pass
        _J.update(jax=jax, jnp=jnp)
    return _J["jax"], _J["jnp"]


def available(require_gpu=True):
    """(True, "") when JAX imports and (if require_gpu) has a CUDA device; else (False, reason). Initialises XLA."""
    if "avail" not in _J:
        try:
            jax, _ = _jax()
            devs = jax.devices()
            plat = devs[0].platform if devs else "none"
            _J["avail"] = (True, plat, jax.__version__)
        except Exception as e:  # noqa: BLE001
            _J["avail"] = (False, f"{type(e).__name__}: {str(e)[:150]}", None)
    ok, plat, ver = _J["avail"]
    if not ok:
        return False, f"JAX not importable ({plat})"
    if require_gpu and plat not in ("gpu", "cuda"):
        return False, f"JAX {ver} has no CUDA device (platform {plat})"
    return True, ""


# ============================================================================ element matrix in corner order
def ke_corner_order(element_type, h, nu):
    """24x24 unit-E element matrix (h included) with local node k = 4a + 2b + c <-> corner (a,b,c) (a slowest), DOF 3k+i.
    The production Khat (element_matrices.py) orders its nodes (-,-,-) (+,-,-) (+,+,-) (-,+,-) (-,-,+) (+,-,+) (+,+,+) (-,+,+)."""
    kh = element_khat(element_type, h, nu)
    perm = np.zeros(24, dtype=int)
    for a in range(8):
        k = 4 * int((_XI[a] + 1) / 2) + 2 * int((_ETA[a] + 1) / 2) + int((_ZETA[a] + 1) / 2)
        perm[3 * a:3 * a + 3] = 3 * k + np.arange(3)
    out = np.zeros((24, 24))
    out[np.ix_(perm, perm)] = kh
    return 0.5 * (out + out.T)


def _child_prolongators():
    """Pc[m][n][l]: (24,24) child(m,n,l)-node -> coarse-corner trilinear weights, dof-expanded (identity on components)."""
    out = np.zeros((2, 2, 2, 24, 24))
    for m, n, l in itertools.product(range(2), repeat=3):
        for a, (di, dj, dk) in enumerate(CORNERS):
            for b, (ei, ej, ek) in enumerate(CORNERS):
                w = (max(0.0, 1 - 0.5 * abs(m + di - 2 * ei)) * max(0.0, 1 - 0.5 * abs(n + dj - 2 * ej))
                     * max(0.0, 1 - 0.5 * abs(l + dk - 2 * ek)))
                for s in range(3):
                    out[m, n, l, 3 * a + s, 3 * b + s] = w
    return out


def level_dims(dims, levels=None, coarse_max_nodes=400):
    """Element-count triples of the hierarchy, finest first (mirrors gpu_CG.gmg_level_dims, which works on node counts)."""
    out = [tuple(int(d) for d in dims)]
    while True:
        nodes = [n + 1 for n in out[-1]]
        if levels is not None:
            if len(out) >= levels:
                break
        elif np.prod(nodes) <= coarse_max_nodes:
            break
        if max(nodes) <= 2:
            break
        out.append(tuple((n + 1) // 2 for n in out[-1]))
    if len(out) < 2:
        raise ValueError("grid too small to coarsen")
    return out


# ============================================================================ operators (built per dtype)
def _make_ops(dims, dtype, ke, form):
    """setup(E, free) -> params and op(params, v) -> K v (identity on fixed DOFs); fields are (nx+1,ny+1,nz+1,3) / (nx,ny,nz)."""
    jax, jnp = _jax()
    HIGH = jax.lax.Precision.HIGHEST
    kej = jnp.asarray(ke, dtype)
    kd = jnp.asarray(np.stack([ke[3 * k:3 * k + 3, 3 * k:3 * k + 3] for k in range(8)]), dtype)
    nx, ny, nz = dims
    nn = (nx + 1, ny + 1, nz + 1)

    def gather(u):
        return jnp.concatenate([u[a:a + nx, b:b + ny, c:c + nz] for a, b, c in CORNERS], -1)

    def scatter(fe):
        return sum(jnp.pad(fe[..., 3 * k:3 * k + 3], ((a, 1 - a), (b, 1 - b), (c, 1 - c), (0, 0)))
                   for k, (a, b, c) in enumerate(CORNERS))

    def block_diag(E):
        out = 0.0
        for k, (a, b, c) in enumerate(CORNERS):
            out = out + jnp.pad(E[..., None, None] * kd[k], ((a, 1 - a), (b, 1 - b), (c, 1 - c), (0, 0), (0, 0)))
        return out

    def stencil_build(E, free):
        Ep = jnp.pad(E, ((1, 1), (1, 1), (1, 1)))
        acc = [0.0] * 27
        for k, ck in enumerate(CORNERS):
            Es = Ep[1 - ck[0]:1 - ck[0] + nn[0], 1 - ck[1]:1 - ck[1] + nn[1], 1 - ck[2]:1 - ck[2] + nn[2]]
            for kp, ckp in enumerate(CORNERS):
                d = (ckp[0] - ck[0], ckp[1] - ck[1], ckp[2] - ck[2])
                acc[OFF_INDEX[d]] = acc[OFF_INDEX[d]] + Es[..., None, None] * jnp.asarray(ke[3 * k:3 * k + 3, 3 * kp:3 * kp + 3], dtype)
        A = jnp.stack(acc)
        fr = free.astype(dtype)
        frp = jnp.pad(fr, ((1, 1), (1, 1), (1, 1), (0, 0)))
        rows = []
        for i, (da, db, dc) in enumerate(OFFS):
            fcol = frp[1 + da:1 + da + nn[0], 1 + db:1 + db + nn[1], 1 + dc:1 + dc + nn[2]]
            Ai = A[i] * (fr[..., :, None] * fcol[..., None, :])
            if (da, db, dc) == (0, 0, 0):
                Ai = Ai + (1 - fr)[..., :, None] * jnp.eye(3, dtype=dtype)
            rows.append(Ai)
        return jnp.stack(rows)

    def setup(E, free):
        E = E.astype(dtype)
        p = {"E": E, "free": free}
        if form == "stencil":
            p["A"] = stencil_build(E, free)
        return p

    def op(p, v):
        if form == "slice":
            return jnp.where(p["free"], scatter(jnp.matmul(gather(v), kej, precision=HIGH) * p["E"][..., None]), v)
        vp = jnp.pad(v, ((1, 1), (1, 1), (1, 1), (0, 0)))
        out = 0.0
        for i, (da, db, dc) in enumerate(OFFS):
            vs = vp[1 + da:1 + da + nn[0], 1 + db:1 + db + nn[1], 1 + dc:1 + dc + nn[2]]
            out = out + jnp.einsum("xyzij,xyzj->xyzi", p["A"][i], vs, precision=HIGH)
        return out

    def level0_block_diag(p):
        """masked 3x3 nodal blocks of the finest level (what the smoother inverts)."""
        fr = p["free"].astype(dtype)
        bd = block_diag(p["E"])
        return bd * fr[..., :, None] * fr[..., None, :] + (1 - fr)[..., :, None] * jnp.eye(3, dtype=dtype)

    return setup, op, level0_block_diag


# ============================================================================ the V-cycle
def _make_gmg(dims, dtype, ke, op0, level0_block_diag, deg, levels, coarse_max_nodes):
    """(setup2(p) -> hierarchy dict, pre(p, g, r) -> M^-1 r) with the finest operator op0(p, v) of `dtype`."""
    jax, jnp = _jax()
    HIGH = jax.lax.Precision.HIGHEST
    edims = level_dims(dims, levels, coarse_max_nodes)
    L = len(edims)
    ndims = [tuple(n + 1 for n in e) for e in edims]
    ke0 = jnp.asarray(ke, dtype)
    Pc = jnp.asarray(_child_prolongators(), dtype)

    def prolong_axis(c, axis, nfn):
        c = jnp.moveaxis(c, axis, 0)
        ne, no = (nfn + 1) // 2, nfn // 2
        ev = c[:ne]
        od = 0.5 * (c[:no] + c[1:no + 1])
        if ne > no:
            od = jnp.concatenate([od, jnp.zeros_like(od[:1])], 0)
        out = jnp.stack([ev, od], 1).reshape((2 * ne,) + c.shape[1:])[:nfn]
        return jnp.moveaxis(out, 0, axis)

    def restrict_axis(f, axis, ncn):
        f = jnp.moveaxis(f, axis, 0)
        nfn = f.shape[0]
        ne, no = (nfn + 1) // 2, nfn // 2
        out = jnp.zeros((ncn,) + f.shape[1:], f.dtype)
        out = out.at[:ne].add(f[0::2])
        out = out.at[:no].add(0.5 * f[1::2])
        out = out.at[1:no + 1].add(0.5 * f[1::2])
        return jnp.moveaxis(out, 0, axis)

    def prolong(c, fine_nodes):
        for ax in range(3):
            c = prolong_axis(c, ax, fine_nodes[ax])
        return c

    def restrict(f, coarse_nodes):
        for ax in range(3):
            f = restrict_axis(f, ax, coarse_nodes[ax])
        return f

    def stencil_apply(A, v, nn):
        vp = jnp.pad(v, ((1, 1), (1, 1), (1, 1), (0, 0)))
        out = 0.0
        for i, (da, db, dc) in enumerate(OFFS):
            vs = vp[1 + da:1 + da + nn[0], 1 + db:1 + db + nn[1], 1 + dc:1 + dc + nn[2]]
            out = out + jnp.einsum("xyzij,xyzj->xyzi", A[i], vs, precision=HIGH)
        return out

    def stencil_from_kel(Kel, free):
        nx_, ny_, nz_ = Kel.shape[:3]
        nn = (nx_ + 1, ny_ + 1, nz_ + 1)
        Kp = jnp.pad(Kel, ((1, 1), (1, 1), (1, 1), (0, 0), (0, 0)))
        acc = [0.0] * 27
        for k, ck in enumerate(CORNERS):
            Ks = Kp[1 - ck[0]:1 - ck[0] + nn[0], 1 - ck[1]:1 - ck[1] + nn[1], 1 - ck[2]:1 - ck[2] + nn[2]]
            for kp, ckp in enumerate(CORNERS):
                d = (ckp[0] - ck[0], ckp[1] - ck[1], ckp[2] - ck[2])
                i = OFF_INDEX[d]
                acc[i] = acc[i] + Ks[..., 3 * k:3 * k + 3, 3 * kp:3 * kp + 3]
        A = jnp.stack(acc)
        fr = free.astype(dtype)
        frp = jnp.pad(fr, ((1, 1), (1, 1), (1, 1), (0, 0)))
        rows = []
        for i, (da, db, dc) in enumerate(OFFS):
            fcol = frp[1 + da:1 + da + nn[0], 1 + db:1 + db + nn[1], 1 + dc:1 + dc + nn[2]]
            Ai = A[i] * (fr[..., :, None] * fcol[..., None, :])
            if (da, db, dc) == (0, 0, 0):
                Ai = Ai + (1 - fr)[..., :, None] * jnp.eye(3, dtype=dtype)
            rows.append(Ai)
        return jnp.stack(rows)

    def coarse_free(free_f, nc_nodes):
        """coarse DOF fixed if the coincident fine DOF is fixed or no free fine DOF lies in its trilinear support."""
        co = free_f[::2, ::2, ::2]
        pad = [(0, nc_nodes[i] - co.shape[i]) for i in range(3)] + [(0, 0)]
        co = jnp.pad(co, pad, constant_values=True)
        wfree = restrict(free_f.astype(dtype), nc_nodes)
        return co & (wfree > 0)

    def galerkin(src, is_E, free_f, ce):
        """coarse element matrices (ce..., 24, 24) from the fine element data `src` (E field, or the Kel array)."""
        pe = [2 * c for c in ce]
        pad3 = [(0, pe[i] - src.shape[i]) for i in range(3)]
        srcp = jnp.pad(src, pad3 + ([(0, 0), (0, 0)] if not is_E else []))
        frp = jnp.pad(free_f.astype(dtype), [(0, pe[i] + 2 - free_f.shape[i]) for i in range(3)] + [(0, 0)])
        Kc = 0.0
        for m, n, l in itertools.product(range(2), repeat=3):
            fr24 = jnp.concatenate(
                [frp[m + di:m + di + pe[0]:2, n + dj:n + dj + pe[1]:2, l + dk:l + dk + pe[2]:2] for di, dj, dk in CORNERS], -1)
            Pm = fr24[..., :, None] * Pc[m, n, l]
            if is_E:
                Ec = srcp[m::2, n::2, l::2]
                T = jnp.einsum("...ia,ij->...aj", Pm, ke0, precision=HIGH)
                Kt = jnp.einsum("...aj,...jb->...ab", T, Pm, precision=HIGH) * Ec[..., None, None]
            else:
                Kch = srcp[m::2, n::2, l::2]
                T = jnp.einsum("...ia,...ij->...aj", Pm, Kch, precision=HIGH)
                Kt = jnp.einsum("...aj,...jb->...ab", T, Pm, precision=HIGH)
            Kc = Kc + Kt
        return Kc

    def block_inv(bd):
        return jnp.linalg.inv(bd)

    def make_A(p, g, l):
        if l == 0:
            return lambda v: op0(p, v)
        return lambda v: stencil_apply(g["A"][l], v, ndims[l])

    def setup2(p):
        E, free0 = p["E"], p["free"]
        g = {"free": [free0], "A": [None], "Minv": [block_inv(level0_block_diag(p))]}
        Kprev, is_E = E, True
        for l in range(1, L):
            free_c = coarse_free(g["free"][-1], ndims[l])
            Kc = galerkin(Kprev, is_E, g["free"][-1], edims[l])
            A = stencil_from_kel(Kc, free_c)
            g["free"].append(free_c)
            g["A"].append(A)
            g["Minv"].append(jnp.linalg.inv(A[OFF_INDEX[(0, 0, 0)]]))
            Kprev, is_E = Kc, False
        g["theta"], g["delta"] = [], []
        for l in range(L - 1):
            A = make_A(p, g, l)
            fr = g["free"][l].astype(dtype)
            hh = jnp.arange(fr.size, dtype=jnp.uint32) * jnp.uint32(2654435761) + jnp.uint32(12345)   # hash seed of gpu_CG
            hh = hh ^ (hh >> 13)
            hh = hh * jnp.uint32(1274126177)
            hh = hh ^ (hh >> 16)
            v = (hh.astype(dtype) / 4294967296.0 - 0.5).reshape(fr.shape) * fr

            def it(_, v, A=A, Mi=g["Minv"][l], fr=fr):
                w = jnp.einsum("xyzij,xyzj->xyzi", Mi, A(v), precision=HIGH) * fr
                return w / jnp.linalg.norm(w)

            v = jax.lax.fori_loop(0, POWER_ITERS - 1, it, v / jnp.linalg.norm(v))
            w = jnp.einsum("xyzij,xyzj->xyzi", g["Minv"][l], A(v), precision=HIGH) * fr
            lam = jnp.linalg.norm(w) / jnp.linalg.norm(v)
            hi = HI_FACTOR * lam
            lo = hi / CHEB_RATIO
            g["theta"].append(0.5 * (hi + lo))
            g["delta"].append(0.5 * (hi - lo))
        Ac = make_A(p, g, L - 1)
        shape = (*ndims[-1], 3)
        N = int(np.prod(shape))
        basis = jnp.eye(N, dtype=dtype).reshape((N,) + shape)
        Amat = jax.vmap(Ac)(basis).reshape(N, N)
        Amat = 0.5 * (Amat + Amat.T)
        g["dense"] = jnp.linalg.inv(Amat.astype(jnp.float64)).astype(dtype)
        return g

    def pre(p, g, r):
        As = [make_A(p, g, l) for l in range(L)]
        Dinv = lambda l, x: jnp.einsum("xyzij,xyzj->xyzi", g["Minv"][l], x, precision=HIGH)
        fr = [f.astype(dtype) for f in g["free"]]

        def cheb(l, b, x):
            th, de = g["theta"][l], g["delta"][l]
            sigma = th / de
            res = b if x is None else (b - As[l](x)) * fr[l]
            d = Dinv(l, res) / th
            x = d if x is None else x + d
            rho = 1.0 / sigma
            for _ in range(1, deg):
                res = (res - As[l](d)) * fr[l]
                rho1 = 1.0 / (2.0 * sigma - rho)
                d = rho1 * rho * d + (2.0 * rho1 / de) * Dinv(l, res)
                x = x + d
                rho = rho1
            return x

        xs, bs = [None] * L, [None] * L
        bs[0] = r
        for l in range(L - 1):
            xs[l] = cheb(l, bs[l], None)
            res = (bs[l] - As[l](xs[l])) * fr[l]
            bs[l + 1] = restrict(res, ndims[l + 1]) * fr[l + 1]
        shape = (*ndims[-1], 3)
        xs[L - 1] = (g["dense"] @ bs[L - 1].reshape(-1)).reshape(shape)
        for l in range(L - 2, -1, -1):
            xs[l] = xs[l] + prolong(xs[l + 1], ndims[l]) * fr[l]
            xs[l] = cheb(l, bs[l], xs[l])
        return xs[0]

    return setup2, pre, edims


# ============================================================================ the solver (setup jit + solve jit)
class JaxSolver:
    """Compiled hierarchy setup and PCG loop for one grid / element matrix / variant (AOT compiled on construction)."""

    def __init__(self, dims, ke, variant, form, deg, levels, coarse_max_nodes, maxit, warm, probe=8, min_red=0.5):
        jax, jnp = _jax()
        self.dims, self.variant, self.form, self.maxit, self.warm = tuple(dims), variant, form, int(maxit), warm
        dt = jnp.float32 if variant == "fp32" else jnp.float64
        dti = jnp.float64 if variant == "fp64" else jnp.float32
        self.dt, self.dti = dt, dti
        self.np_dt = np.float32 if variant == "fp32" else np.float64
        nx, ny, nz = dims
        nshape = (nx + 1, ny + 1, nz + 1, 3)
        so, op, bd_o = _make_ops(dims, dt, ke, form)
        if dti == dt:
            si, opi, bd_i = so, op, bd_o
        else:
            si, opi, bd_i = _make_ops(dims, dti, ke, form)
        setup2, pre_i, edims = _make_gmg(dims, dti, ke, opi, bd_i, deg, levels, coarse_max_nodes)
        self.edims = edims
        mixed = dti != dt
        self.parts = {}

        def setup(E, free):
            po = so(E, free)
            pi = si(E, free) if mixed else po
            g = setup2(pi)
            return {"o": po, "i": pi, "g": g}

        def pre(p, r):
            z = pre_i(p["i"], p["g"], r.astype(dti) if mixed else r)
            return z.astype(dt) if mixed else z

        maxit_ = self.maxit
        self.parts.update(op=op, opi=opi, pre=pre, setup=setup)       # closures for tools/jax_profile.py

        def solve(p, b, u0, stop2):
            free = p["o"]["free"]
            b = jnp.where(free, b, 0).astype(dt)
            stop2 = stop2.astype(dt)
            if warm:
                u = jnp.where(free, u0, 0).astype(dt)
                r = b - op(p["o"], u)
            else:
                u, r = jnp.zeros_like(b), b
            rn2_0 = jnp.vdot(r, r)
            z = pre(p, r)
            rz0 = jnp.vdot(r, z)
            hist = jnp.zeros(maxit_, dt)

            def cold(_):
                zc = pre(p, b)
                return (jnp.zeros_like(b), b, zc, zc, jnp.vdot(b, zc), jnp.vdot(b, b))

            def body(s):
                k, u, r, z, d, rz, rn2, ok, hist, bail = s
                ap = op(p["o"], d)
                dap = jnp.vdot(d, ap)
                alpha = rz / dap
                u, r = u + alpha * d, r - alpha * ap
                rn2n = jnp.vdot(r, r)
                zn = pre(p, r)
                rzn = jnp.vdot(r, zn)
                dn = zn + (rzn / rz) * d
                hist = hist.at[k].set(jnp.sqrt(rn2n))
                ok = (dap > 0) & jnp.isfinite(rn2n) & jnp.isfinite(rzn) & (rz > 0)
                if warm:
                    required2 = ((1.0 - min_red) ** 2) * rn2_0
                    do_bail = (k + 1 == probe) & (rn2n > required2) & (rn2n > stop2) & ok
                    u, r, zn, dn, rzn, rn2n = jax.lax.cond(
                        do_bail, cold, lambda _: (u, r, zn, dn, rzn, rn2n), None)
                    bail = bail | do_bail
                return k + 1, u, r, zn, dn, rzn, rn2n, ok, hist, bail

            def cond(s):
                return (s[0] < maxit_) & (s[6] > stop2) & s[7]

            init = (jnp.int32(0), u, r, z, z, rz0, rn2_0, rz0 > 0, hist, jnp.bool_(False))
            k, u, r, _, _, _, rn2, ok, hist, bail = jax.lax.while_loop(cond, body, init)
            return u.astype(jnp.float32), k, jnp.sqrt(rn2), hist, ok, bail

        t0 = time.perf_counter()
        sd = jax.ShapeDtypeStruct
        self.setup_c = jax.jit(setup).lower(sd(dims, jnp.float64), sd(nshape, jnp.bool_)).compile()
        t1 = time.perf_counter()
        p_abs = jax.eval_shape(setup, sd(dims, jnp.float64), sd(nshape, jnp.bool_))
        self.solve_c = jax.jit(solve).lower(p_abs, sd(nshape, dt), sd(nshape, dt), sd((), jnp.float64)).compile()
        t2 = time.perf_counter()
        self.compile_s = dict(setup=t1 - t0, solve=t2 - t1, total=t2 - t0)
        try:
            ms, mv = self.setup_c.memory_analysis(), self.solve_c.memory_analysis()
            tot = lambda m: m.temp_size_in_bytes + m.argument_size_in_bytes + m.output_size_in_bytes
            self.mem_analysis = dict(setup_bytes=tot(ms), solve_bytes=tot(mv), setup_temp=ms.temp_size_in_bytes,
                                     solve_temp=mv.temp_size_in_bytes)
        except Exception:  # noqa: BLE001
            self.mem_analysis = {}


def _get_solver(dims, ke_key, ke, variant, form, deg, levels, coarse_max_nodes, maxit, warm, probe, min_red):
    key = (tuple(dims), ke_key, variant, form, int(deg), levels, int(coarse_max_nodes), int(maxit), bool(warm),
           int(probe) if warm else 0, float(min_red) if warm else 0.0)
    s = CACHE.get(key)
    new = s is None
    if new:
        s = CACHE[key] = JaxSolver(dims, ke, variant, form, deg, levels, coarse_max_nodes, maxit, warm, probe, min_red)
    return s, new


# ============================================================================ entry points
def ok(nx, ny, nz, hx, hy, hz, element_type, preconditioner, gmg_levels, gmg_coarse_max_nodes, require_gpu=True):
    """(True, "") when this engine can run the solve, else (False, reason)."""
    if preconditioner != "gmg":
        return False, f"pcg_engine='jax' needs preconditioner='gmg' (got {preconditioner!r})"
    if element_type not in ("q1", "q1_im"):
        return False, f"pcg_engine='jax' needs a matrix element type ('q1' or 'q1_im'), not {element_type!r}"
    if not (abs(float(hx) - float(hy)) <= 1e-6 * abs(float(hx)) and abs(float(hx) - float(hz)) <= 1e-6 * abs(float(hx))):
        return False, "pcg_engine='jax' needs cubic voxels (hx == hy == hz)"
    try:
        level_dims((nx - 1, ny - 1, nz - 1), gmg_levels, gmg_coarse_max_nodes)
    except ValueError:
        return False, f"grid {nx - 1}x{ny - 1}x{nz - 1} elements is too small to coarsen"
    return available(require_gpu)


def memory_estimate(nx, ny, nz, variant=DEFAULT_VARIANT, form=DEFAULT_FORM):
    """Rough peak device bytes of one solve (engines.JAX_BYTES_PER_ELEMENT is the selection model; calibrated on GH200)."""
    return pcg_jax_memory(nx, ny, nz, variant, form) / SAFETY


def to_numba_device(x):
    """Zero-copy view of a (single-device, CUDA) JAX array as a numba-CUDA device array (via __cuda_array_interface__ when
    JAX provides it, else via the raw buffer pointer). The view keeps `x` alive; the caller must not use it after x is freed."""
    from numba import cuda
    if hasattr(x, "__cuda_array_interface__"):
        return cuda.as_cuda_array(x)
    import ctypes
    ptr = x.addressable_data(0).unsafe_buffer_pointer()
    dt = np.dtype(x.dtype)
    n = int(np.prod(x.shape)) * dt.itemsize
    mem = cuda.cudadrv.driver.MemoryPointer(cuda.current_context(), ctypes.c_void_p(ptr), n)
    return cuda.cudadrv.devicearray.DeviceNDArray(x.shape, tuple(int(s) for s in np.empty(x.shape, dt).strides), dt, gpu_data=mem)


def pcg(nx, ny, nz, hx, hy, hz, E_elem_h, nu_elem_h, b_h, fixed_mask_h, max_iters=200, tol=1e-6, rel_tol=0.0, verbose=True,
        u0=None, warm_start_probe_iters=8, warm_start_min_reduction=0.5, element_type="q1", nu=None,
        gmg_levels=None, gmg_smoother_steps=2, gmg_coarse_max_nodes=400, variant=None, form=None):
    """JAX GMG-PCG solve of K u = b; arguments as gpu_CG.matrix_free_pcg (host arrays in). Returns (u_h float32 host,
    history list of ||r|| per iteration, info). info["gmg"] has setup_s / cg_loop_s as the numba engines; info["jax"] the JAX
    detail (variant, form, compile_s when this call compiled, h2d_s, d2h_s, device array u_device)."""
    jax, jnp = _jax()
    variant = variant or os.environ.get("PCLS_JAX_VARIANT", DEFAULT_VARIANT)
    form = form or os.environ.get("PCLS_JAX_FORM", DEFAULT_FORM)
    if variant not in VARIANTS or form not in FORMS:
        raise ValueError(f"jax variant must be in {VARIANTS} and form in {FORMS}, got {variant!r}, {form!r}")
    pc = time.perf_counter
    t0 = pc()
    dims = (nx - 1, ny - 1, nz - 1)
    Nd = 3 * nx * ny * nz
    nshape = (nx, ny, nz, 3)
    fixed = np.asarray(fixed_mask_h).reshape(nshape)
    b = np.asarray(b_h, dtype=np.float32).reshape(nshape).copy()
    n_load_on_fixed = int(np.count_nonzero((fixed == 1) & (b != 0.0)))
    if n_load_on_fixed:
        b[fixed == 1] = 0.0
        if verbose:
            print(f"Note: removed load from {n_load_on_fixed} constrained DOF(s) "
                  f"(no work is done there; it would otherwise floor the residual).")
    stop_tol = max(float(tol), float(rel_tol) * float(np.linalg.norm(b)))
    if verbose:
        print(f"PCG (jax engine, {variant}/{form}) stopping tolerance: {stop_tol:.6e} "
              f"(absolute {tol:.3e}, relative {rel_tol:.3e} x ||b||)")
    nu0 = resolve_nu(nu_elem_h, E_elem_h, nu)
    h = float(hx)
    ke = ke_corner_order(element_type, (h, h, h), nu0)
    warm = u0 is not None
    solver, compiled_now = _get_solver(dims, (element_type, h, nu0), ke, variant, form, gmg_smoother_steps, gmg_levels,
                                       gmg_coarse_max_nodes, max_iters, warm, warm_start_probe_iters,
                                       warm_start_min_reduction)
    npdt = solver.np_dt
    t_h2d0 = pc()
    E_d = jax.device_put(np.asarray(E_elem_h, dtype=np.float64).reshape(dims))
    free_d = jax.device_put(fixed == 0)
    b_d = jax.device_put(b.astype(npdt))
    if warm:
        u0h = np.asarray(u0, dtype=np.float32)
        if u0h.shape[0] != Nd:
            raise ValueError(f"u0 has {u0h.shape[0]} DOFs, expected {Nd}.")
        u0_d = jax.device_put(u0h.reshape(nshape).astype(npdt))
    else:
        u0_d = b_d     # unused placeholder of the right shape/dtype
    stop2_d = jax.device_put(np.float64(stop_tol) ** 2)
    jax.block_until_ready((E_d, free_d, b_d, u0_d, stop2_d))
    h2d_s = pc() - t_h2d0
    t_s0 = pc()
    p = solver.setup_c(E_d, free_d)
    jax.block_until_ready(p)
    setup_s = pc() - t_s0
    if PHASE_HOOK[0] is not None:
        PHASE_HOOK[0]()
    t_l0 = pc()
    u_d, k_d, rn_d, hist_d, ok_d, bail_d = solver.solve_c(p, b_d, u0_d, stop2_d)
    jax.block_until_ready(u_d)
    cg_loop_s = pc() - t_l0
    t_d0 = pc()
    u_h = np.asarray(u_d).reshape(-1)
    k = int(k_d)
    history = [float(v) for v in np.asarray(hist_d[:k])] if k > 0 else [float(rn_d)]
    ok_flag, bailed = bool(ok_d), bool(bail_d)
    d2h_s = pc() - t_d0
    info = {"warm_started": warm, "bailed_to_cold": bailed, "bailout_iter": (int(warm_start_probe_iters) - 1) if bailed else None,
            "preconditioner_used": "gmg", "pcg_engine_used": "jax"}
    info["gmg"] = dict(levels=list(solver.edims), setup_s=setup_s, reused=False, rel_change=None, cg_loop_s=cg_loop_s,
                       smoother_steps=int(gmg_smoother_steps), coarsest_solve="dense_inverse")
    info["jax"] = dict(variant=variant, form=form, compiled_now=compiled_now, compile_s=solver.compile_s if compiled_now else None,
                       h2d_s=h2d_s, d2h_s=d2h_s, total_s=pc() - t0, breakdown=(not ok_flag) and float(rn_d) > stop_tol, u_device=u_d,
                       mem_analysis=solver.mem_analysis)
    LAST.clear()
    LAST.update(setup_s=setup_s, cg_loop_s=cg_loop_s, h2d_s=h2d_s, d2h_s=d2h_s, iters=k, compiled_now=compiled_now)
    return u_h, history, info


# ======================================================================================================================
# === former engines.py ===
# ======================================================================================================================
"""
Capability probe and engine selection for the ConjugateGradient backend (docs/PLAN_engine_buckling.md, parts A, E1,
E2; formulas and tables in docs/README.md ("Solver engines")).

Two choices are made per process:

  elastic  (analysis.parameters.engine)  "classic" | "auto" | "gpu_gmg" | "gpu_gmg_lean" | "gpu_block_jacobi" |
                                          "cpu_direct"
  buckling (optimiser.buckling.engine)   "classic" | "auto" | "cudss64" | "cudss32" | "cudss64_hybrid" |
                                          "cudss32_hybrid" | "gmg_lobpcg32" | "cpu_cholesky_lobpcg" |
                                          "cpu_cholesky_eigsh"

"classic" is today's code path, untouched (the guard pins it). "auto" probes the machine once -- CUDA device,
compute capability, FP64:FP32 rate, total/free VRAM, importable packages, OS/arch, cores, host RAM -- and picks
the fastest engine whose memory model fits. A selected engine that fails at run time (e.g. cuDSS ALLOC_FAILED,
an import error, a non-finite float32 factor) is replaced by the next one in its chain; one printed line, and
the event is kept in RECORD["fallbacks"] (run_info.json "engines").

Nothing here initialises CUDA unless an "auto" selection needs the device facts, and only numpy is imported at
module level, so this imports on Windows, aarch64 and CPU-only machines.

Emulated VRAM cap (plan E6): PCLS_VRAM_CAP_GB=<GB> makes the probe report at most that much total/free device
memory and caps CuPy's memory pool, so selection and the CuPy/nvmath allocations behave as on a smaller card.
"""
import json
import math
import os
import platform
import sys
import time
from pathlib import Path

import numpy as np

ELASTIC_ENGINES = ("classic", "auto", "gpu_gmg", "gpu_gmg_lean", "gpu_block_jacobi", "cpu_direct")
LEVELSET_ENGINES = ("classic", "auto", "gpu")      # optimiser.levelset.engine (the former gpu_levelset.py section of levelset_optimiser.py)
BUCKLING_ENGINES = ("classic", "auto", "cudss64", "cudss32", "cudss64_hybrid", "cudss32_hybrid", "gmg_lobpcg32",
                    "cpu_cholesky_lobpcg", "cpu_cholesky_eigsh", "cudss64_eigsh")

GiB = float(2 ** 30)
VRAM_FIT_FRACTION = 0.85       # an engine fits when its estimate <= this x free VRAM
HOST_FIT_FRACTION = 0.7        # ... or this x available host RAM (host-side factors)
SAFETY = 1.15                  # multiplier on every memory estimate (model calibrated on GH200, LOG 2026-10-22)
LEAN_VRAM_GB = 4.0             # cards at or below this total VRAM get the memory-lean elastic mode

# FP64:FP32 throughput by compute capability (vendor specifications). Datacentre parts are 1:2; consumer parts
# 1:32 (Turing sm_75) or 1:64 (Ampere sm_86, Ada sm_89, Blackwell consumer sm_120). Unknown -> measured.
FP64_RATIO = {(6, 0): 1 / 2, (7, 0): 1 / 2, (8, 0): 1 / 2, (9, 0): 1 / 2, (10, 0): 1 / 2,
              (6, 1): 1 / 32, (7, 5): 1 / 32, (8, 6): 1 / 64, (8, 9): 1 / 64, (12, 0): 1 / 64}
DATACENTRE_FP64 = 0.25         # ratio at or above this = "fast FP64"

# Process-wide record written to run_info.json["engines"] (empty while only "classic" ran, so classic runs carry
# no new key and the guard stays bit-exact).
RECORD = {}
FORCE_CLASSIC = False         # True: every engine option reads as "classic" (tools/guard.py pins it)
FAILED = set()                 # engines that failed at run time in this process
_CAP = {}
_DEFAULTS = {"buckling_engine": "auto", "lobpcg_tol": 1e-3}   # "auto" since tests 1-4 passed (LOG 2026-10-23)


# ============================================================================ configuration (optimiser -> analysis)
def configure_buckling(engine=None, lobpcg_tol=None):
    """Process default for ConjugateGradient.run_buckling(engine=None, lobpcg_tol=None); set by the optimiser from
    optimiser.buckling.engine / lobpcg_tol, so every run_buckling call site (inner loop, exported design, threshold
    sweep, SIMP / level set) uses the same engine."""
    if engine is not None:
        if engine not in BUCKLING_ENGINES:
            raise ValueError(f"buckling engine must be one of {BUCKLING_ENGINES}, got {engine!r}")
        _DEFAULTS["buckling_engine"] = engine
    if lobpcg_tol is not None:
        if not (0.0 < float(lobpcg_tol) < 1.0):
            raise ValueError("buckling lobpcg_tol must be in (0, 1)")
        _DEFAULTS["lobpcg_tol"] = float(lobpcg_tol)


def buckling_defaults():
    return dict(_DEFAULTS, buckling_engine="classic") if FORCE_CLASSIC else dict(_DEFAULTS)


# ============================================================================ probe
def _vram_cap_bytes():
    v = os.environ.get("PCLS_VRAM_CAP_GB")
    if not v:
        return None
    try:
        return float(v) * GiB
    except ValueError:
        return None


def _host_ram():
    """(total, available) bytes, or (None, None)."""
    try:
        import psutil
        vm = psutil.virtual_memory()
        return float(vm.total), float(vm.available)
    except Exception:
        pass
    try:
        pages, avail = os.sysconf("SC_PHYS_PAGES"), os.sysconf("SC_AVPHYS_PAGES")
        ps = os.sysconf("SC_PAGE_SIZE")
        return float(pages * ps), float(avail * ps)
    except Exception:
        pass
    try:                         # Windows has no os.sysconf: GlobalMemoryStatusEx (kernel32)
        import ctypes

        class _MemoryStatusEx(ctypes.Structure):
            _fields_ = [("dwLength", ctypes.c_ulong), ("dwMemoryLoad", ctypes.c_ulong),
                        ("ullTotalPhys", ctypes.c_ulonglong), ("ullAvailPhys", ctypes.c_ulonglong),
                        ("ullTotalPageFile", ctypes.c_ulonglong), ("ullAvailPageFile", ctypes.c_ulonglong),
                        ("ullTotalVirtual", ctypes.c_ulonglong), ("ullAvailVirtual", ctypes.c_ulonglong),
                        ("ullAvailExtendedVirtual", ctypes.c_ulonglong)]
        ms = _MemoryStatusEx()
        ms.dwLength = ctypes.sizeof(ms)
        if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(ms)):
            return float(ms.ullTotalPhys), float(ms.ullAvailPhys)
    except Exception:
        pass
    return None, None


def _importable(name):
    try:
        __import__(name)
        return True
    except Exception:
        return False


def _cudss_available():
    try:
        from nvmath.sparse.advanced import DirectSolver  # noqa: F401
        from nvmath.bindings import cudss                # noqa: F401  (loads libcudss)
        cudss.get_property(0)                            # loads libcudss (raises if missing)
        return True
    except Exception:
        return False


def _fp64_cache_path():
    return Path.home() / ".pcls_engine_cache.json"


def measure_fp64_ratio():
    """FP64:FP32 FMA throughput ratio by a ~10 ms numba-CUDA microbenchmark (cached per device name)."""
    from numba import cuda

    @cuda.jit
    def fma64(a, n):
        i = cuda.grid(1)
        x = a[i]; y = 1.0000001
        for _ in range(n):
            x = x * y + 1e-9
        a[i] = x

    @cuda.jit
    def fma32(a, n):
        i = cuda.grid(1)
        x = a[i]; y = np.float32(1.0000001)
        for _ in range(n):
            x = x * y + np.float32(1e-9)
        a[i] = x

    nthr = 256 * 1024
    out = {}
    for name, kern, dt in (("f64", fma64, np.float64), ("f32", fma32, np.float32)):
        a = cuda.to_device(np.ones(nthr, dtype=dt))
        kern[nthr // 256, 256](a, 16)                # compile + warm
        cuda.synchronize()
        t0 = time.perf_counter()
        kern[nthr // 256, 256](a, 2048)
        cuda.synchronize()
        out[name] = time.perf_counter() - t0
    return float(min(1.0, out["f32"] / max(out["f64"], 1e-9)))


def probe(gpu=True, refresh=False):
    """Capability summary (dict, cached per process). gpu=False never touches CUDA."""
    if _CAP and not refresh and (_CAP.get("gpu_probed") or not gpu):
        return dict(_CAP)
    cap = dict(os=platform.system(), arch=platform.machine(), python=sys.version.split()[0],
               cpu_cores=os.cpu_count())
    tot, avail = _host_ram()
    cap.update(host_ram_GB=None if tot is None else tot / GiB, host_ram_avail_GB=None if avail is None else avail / GiB)
    pk = {m: _importable(m) for m in ("numba", "cupy", "nvmath", "pypardiso", "sksparse")}
    cap["packages"] = pk
    cap["host_cholesky"] = "pardiso" if pk["pypardiso"] else ("cholmod" if pk["sksparse"] else "splu")
    cap["cuda"] = False
    cap["gpu_probed"] = bool(gpu)
    if gpu and pk["numba"]:
        try:
            from numba import cuda
            if cuda.is_available():
                dev = cuda.get_current_device()
                cc = tuple(int(v) for v in dev.compute_capability)
                free, total = cuda.current_context().get_memory_info()
                name = dev.name.decode() if isinstance(dev.name, bytes) else str(dev.name)
                cap.update(cuda=True, gpu_name=name, cc=list(cc), vram_total_GB=total / GiB, vram_free_GB=free / GiB)
                capb = _vram_cap_bytes()
                if capb is not None:
                    used = total - free
                    cap["vram_cap_GB"] = capb / GiB
                    cap["vram_total_GB"] = min(total, capb) / GiB
                    cap["vram_free_GB"] = max(0.0, min(free, capb - used)) / GiB
                ratio = FP64_RATIO.get(cc)
                src = "table"
                if ratio is None:
                    cache = {}
                    try:
                        cache = json.loads(_fp64_cache_path().read_text())
                    except Exception:
                        pass
                    ratio = cache.get(name)
                    src = "cache"
                    if ratio is None:
                        ratio = measure_fp64_ratio()
                        src = "measured"
                        cache[name] = ratio
                        try:
                            _fp64_cache_path().write_text(json.dumps(cache))
                        except Exception:
                            pass
                cap.update(fp64_ratio=float(ratio), fp64_ratio_source=src)
        except Exception as e:  # noqa: BLE001
            cap["cuda_error"] = f"{type(e).__name__}: {e}"[:200]
    cap["cupy_kernels"] = False
    if cap["cuda"] and pk["cupy"]:
        try:                                    # NVRTC compile of the engine kernels (needs the CUDA headers)
            import analysis
            analysis._cuda_kernels()
            cap["cupy_kernels"] = True
        except Exception as e:  # noqa: BLE001
            cap["cupy_kernels_error"] = f"{type(e).__name__}: {e}"[:300]
    cap["cudss"] = bool(cap["cupy_kernels"] and pk["nvmath"] and _cudss_available())
    _CAP.clear(); _CAP.update(cap)
    apply_vram_cap()
    return dict(cap)


def apply_vram_cap():
    """Cap CuPy's memory pool at PCLS_VRAM_CAP_GB (emulation of a smaller card)."""
    capb = _vram_cap_bytes()
    if capb is None:
        return
    try:
        import cupy as cp
        cp.get_default_memory_pool().set_limit(size=int(capb))
    except Exception:
        pass


# ============================================================================ memory model
def grid_sizes(nx, ny, nz, n_free=None):
    """Element / DOF counts from node counts."""
    Ne = (nx - 1) * (ny - 1) * (nz - 1)
    Nd = 3 * nx * ny * nz
    return dict(nx=nx, ny=ny, nz=nz, Ne=Ne, Nd=Nd, n=Nd if n_free is None else int(n_free))


def cholesky_nnz(nx, ny, nz, n):
    """Estimated nnz of the sparse Cholesky factor of the free-DOF K (nested dissection).
    min(separator model, volume model): 0.45 n S with S = 3 d2 d3 (the two smallest node counts: the cross-section
    separator of a slender grid), and 30 n^(4/3) (cube-like grids). Calibrated on Column_Buckling (240x30x30,
    n = 680 388: CHOLMOD nnz(L) = 8.9e8; the separator model gives 8.8e8)."""
    d = sorted((nx, ny, nz))
    sep = 0.45 * n * 3.0 * d[0] * d[1]
    vol = 30.0 * n ** (4.0 / 3.0)
    return float(min(sep, vol, 0.5 * n * n))


def buckling_memory(engine, nx, ny, nz, n_free, num_modes=6):
    """(device_bytes, host_bytes) estimate of one buckling evaluation (x SAFETY). Terms:
      nnzK = 81 n (27-node stencil x 3), nnzL from cholesky_nnz;
      element data  E = 1536 Ne   (U_e 24, stresses 2 x 48, S 64 doubles + scalars per element);
      LOBPCG blocks B = 8 n nb x 12 (X, AX, BX, W, AW, BW, P, AP, BP + temporaries), nb = num_modes + 4;
      matrix-free scratch F = 2 x 8 Nd nb;  stencil assembly scratch A = 8 x 243 Nn (transient);
      cudss64: 1.5 x 8 nnzL + 12 nnzK + E + B + F           (1.5 = cuDSS workspace)
      cudss32: 7.3 nnzL + E + max(A + 12 nnzK, 20 nnzK, 8 nnzK + B + F)   (float64 assembly, float32 copy;
               7.3 B per factor entry measured: float32 values + cuDSS index / supernode data)
      *_hybrid: device = limit + the non-factor terms above; host = 12 or 7.3 nnzL
      gmg_lobpcg32: E + B + F + 60 Nd (float32 V-cycle hierarchy and vectors)
      cpu_*: host = 1.5 x 8 nnzL + 12 (nnzK + nnzK/3) + B, device 0
    Calibration (GH200, Column_Buckling, tools/engine_memdiag.py job 7644301; peak minus the 0.9 GB of CUDA
    context + cuSOLVER/cuBLAS handles): cudss64 11.5 GB measured vs 11.5 model; cudss32 7.4 vs 7.4;
    gmg_lobpcg32 0.9 GB transient vs 1.05 model (+1.1 GB resident on GH200 from numba kernel local-memory
    reservation, which scales with the SM count and is not modelled)."""
    s = grid_sizes(nx, ny, nz, n_free)
    n, Ne, Nd, Nn = s["n"], s["Ne"], s["Nd"], nx * ny * nz
    nb = num_modes + 4
    nnzK = 81.0 * n
    nnzL = cholesky_nnz(nx, ny, nz, n)
    E = 1536.0 * Ne
    B = 8.0 * n * nb * 12
    F = 2 * 8.0 * Nd * nb
    A = 8.0 * 243 * Nn
    if engine in ("cudss64", "cudss64_eigsh"):
        dev, host = 1.5 * 8 * nnzL + 12 * nnzK + E + B + F, 0.0
    elif engine == "cudss32":
        dev, host = 7.3 * nnzL + E + max(A + 12 * nnzK, 20 * nnzK, 8 * nnzK + B + F), 0.0
    elif engine == "cudss64_hybrid":
        dev, host = 12 * nnzK + E + B + F, 1.5 * 8 * nnzL
    elif engine == "cudss32_hybrid":
        dev, host = E + max(A + 12 * nnzK, 20 * nnzK, 8 * nnzK + B + F), 7.3 * nnzL
    elif engine == "gmg_lobpcg32":
        dev, host = E + B + F + 60.0 * Nd, 0.0
    elif engine in ("cpu_cholesky_lobpcg", "cpu_cholesky_eigsh"):
        dev, host = 0.0, 1.5 * 8 * nnzL + 12 * (nnzK + nnzK / 3) + B + E
    elif engine == "classic":
        dev, host = 0.0, 1.5 * 8 * nnzL * 2 + 24 * 81 * Nd * 3     # LU of COO-assembled K (twice the Cholesky fill)
    else:
        raise ValueError(engine)
    return SAFETY * dev, SAFETY * host


def elastic_memory(engine, nx, ny, nz):
    """Device bytes of one elastic solve: float32 PCG vectors + element data (+ the GMG hierarchy, ~1/7 of the
    fine level per coarser level). Measured transient on GH200 (engine_bench elastic, Column_Buckling): 79 B/DOF
    block-Jacobi, 206 B/DOF GMG (hierarchy rebuilt per solve), 75 B/DOF GMG with reuse (+ the resident hierarchy):
    80 / 210 / 150 B/DOF (x SAFETY)."""
    s = grid_sizes(nx, ny, nz)
    per = {"gpu_block_jacobi": 80.0, "gpu_gmg": 210.0, "gpu_gmg_lean": 150.0, "cpu_direct": 0.0, "classic": 210.0}[engine]
    return SAFETY * (per * s["Nd"] + 4.0 * s["Ne"])


# ============================================================================ selection
def _fits_device(cap, nbytes):
    return cap.get("cuda") and nbytes <= VRAM_FIT_FRACTION * cap.get("vram_free_GB", 0.0) * GiB


def _fits_host(cap, nbytes):
    avail = cap.get("host_ram_avail_GB")
    return nbytes <= HOST_FIT_FRACTION * (avail if avail is not None else 8.0) * GiB


def _gmg_coarsenable(nx, ny, nz):
    try:
        return len(gmg_level_dims(nx, ny, nz)) >= 2
    except ValueError:
        return False
    except Exception:            # gpu_CG not importable here: let the engine try (run-time fallback)
        return True


def hybrid_limit_bytes(cap):
    """cuDSS hybrid-memory device limit: half of the free VRAM (not a fixed 4 GiB, which failed on a 2 GB card)."""
    return int(max(0.25 * GiB, 0.5 * cap.get("vram_free_GB", 0.0) * GiB))


def buckling_chain(cap, nx, ny, nz, n_free, num_modes=6, requested="auto"):
    """Ordered list of (engine, reason) to try; the first is the selection, the rest the run-time fallbacks.
    'classic' ends every chain where it can work (pypardiso present, or a small problem for SuperLU)."""
    fast64 = cap.get("fp64_ratio", 0.0) >= DATACENTRE_FP64
    if cap.get("cuda") and cap.get("packages", {}).get("cupy") and cap.get("cupy_kernels", True):
        if fast64:
            gpu_order = ["cudss64", "cudss32", "cudss64_hybrid", "gmg_lobpcg32"]
        else:
            gpu_order = ["cudss32", "gmg_lobpcg32", "cudss32_hybrid"]
    else:
        gpu_order = []
    order = gpu_order + ["cpu_cholesky_lobpcg", "classic"]
    out = []
    why = {}
    for e in order:
        dev, host = buckling_memory(e, nx, ny, nz, n_free, num_modes)
        ok = True
        if e.startswith("cudss") and not cap.get("cudss"):
            ok, why[e] = False, "cuDSS / nvmath-python not available"
        elif e == "gmg_lobpcg32" and not cap.get("cuda"):
            ok, why[e] = False, "no CUDA device"
        elif e == "gmg_lobpcg32" and not _gmg_coarsenable(nx, ny, nz):
            ok, why[e] = False, "grid too small for a multigrid hierarchy"
        elif e.endswith("_hybrid"):
            if not _fits_host(cap, host):
                ok, why[e] = False, f"host factor {host / GiB:.1f} GB > {HOST_FIT_FRACTION:.0%} of free RAM"
            elif not _fits_device(cap, dev + hybrid_limit_bytes(cap)):
                ok, why[e] = False, f"device part {(dev + hybrid_limit_bytes(cap)) / GiB:.1f} GB does not fit"
        elif e in ("cpu_cholesky_lobpcg",):
            if cap.get("host_cholesky") == "splu" and n_free > 300_000:
                ok, why[e] = False, "no Cholesky (pypardiso / scikit-sparse) and SuperLU is too slow at this size"
            elif not _fits_host(cap, host):
                ok, why[e] = False, f"host factor {host / GiB:.1f} GB does not fit"
        elif e == "classic":
            if not cap.get("packages", {}).get("pypardiso") and n_free > 300_000:
                ok, why[e] = False, "classic uses SuperLU without pypardiso (segfaulted at 42.6 GB on 0.68 M DOF)"
        elif dev > 0 and not _fits_device(cap, dev):
            ok, why[e] = False, (f"needs ~{dev / GiB:.1f} GB > {VRAM_FIT_FRACTION:.0%} of "
                                 f"{cap.get('vram_free_GB', 0):.1f} GB free VRAM")
        if ok:
            out.append((e, f"est. {dev / GiB:.2f} GB device / {host / GiB:.2f} GB host; "
                           f"FP64:FP32 {cap.get('fp64_ratio', 'n/a')}"))
    if requested not in (None, "auto"):
        names = [e for e, _ in out]
        if requested in names:
            out = out[names.index(requested):]
        else:
            out = [(requested, "requested explicitly" + (f" (auto would skip it: {why[requested]})"
                                                         if requested in why else ""))] + \
                  [x for x in out if x[0] != requested]
    return out, why


def elastic_choice(cap, nx, ny, nz, requested="auto", refinement=1):
    """(engine, preconditioner, gmg_reuse, refinement, reason) for the elastic solve."""
    lean = cap.get("cuda") and cap.get("vram_total_GB", 99.0) <= LEAN_VRAM_GB
    if requested == "auto":
        if not cap.get("cuda"):
            eng = "cpu_direct"
        else:
            eng = "gpu_gmg_lean" if lean else "gpu_gmg"
    else:
        eng = requested
    r = int(refinement)
    reason = "requested" if requested != "auto" else ("no CUDA device" if eng == "cpu_direct" else
                                                     ("<= %.0f GB VRAM: memory-lean" % LEAN_VRAM_GB if lean else "CUDA device"))
    if eng.startswith("gpu"):
        while r > 1:
            fx, fy, fz = (nx - 1) * r + 1, (ny - 1) * r + 1, (nz - 1) * r + 1
            if _fits_device(cap, elastic_memory(eng, fx, fy, fz)):
                break
            r -= 1
        if r != int(refinement):
            reason += f"; analysis_refinement capped {refinement} -> {r} by the memory model"
    pre = {"gpu_gmg": "gmg", "gpu_gmg_lean": "gmg", "gpu_block_jacobi": "block_jacobi"}.get(eng)
    reuse = 0.05 if eng == "gpu_gmg_lean" else None
    return eng, pre, reuse, r, reason


def levelset_choice(requested, dims):
    """("gpu" | "classic", reason) for the level-set grid steps (curvature, velocity extension, upwind transport,
    re-initialisation; docs/theory.md §19). "auto" and "gpu" use the numba-CUDA kernels when a CUDA device is
    usable and memory_bytes fits its free memory; otherwise "classic" (the NumPy / SciPy code)."""
    if requested not in LEVELSET_ENGINES:
        raise ValueError(f"levelset engine must be one of {LEVELSET_ENGINES}, got {requested!r}")
    if FORCE_CLASSIC:
        return "classic", "engines.FORCE_CLASSIC"
    if requested == "classic":
        return "classic", "requested"
    cap = probe(gpu=True)
    if not cap.get("cuda"):
        return "classic", "no CUDA device" + (f" ({cap['cuda_error']})" if cap.get("cuda_error") else "")
    try:
        from levelset_optimiser import memory_bytes
    except Exception as e:  # noqa: BLE001
        return "classic", f"gpu_levelset not importable ({type(e).__name__}: {e})"
    need = SAFETY * memory_bytes(*dims)
    if not _fits_device(cap, need):
        return "classic", (f"needs ~{need / GiB:.2f} GB > {VRAM_FIT_FRACTION:.0%} of "
                           f"{cap.get('vram_free_GB', 0):.1f} GB free VRAM")
    return "gpu", (f"{cap.get('gpu_name', 'CUDA device')}; est. {need / GiB:.2f} GB device"
                   + ("" if requested == "gpu" else "; auto"))


# ============================================================================ JAX elastic PCG engine (PLAN_phase1 step 4)
# analysis.parameters.pcg_engine: "jax" (gpu_jax.py, requested explicitly) or "auto" (this block picks "jax" or "fast"). Docs: docs/theory.md §17.
# Everything JAX is imported lazily by gpu_jax; on a machine without JAX / a CUDA JAX the answer is "fast" (reason in the record).
JAX_PCG_VARIANT = "mixed"        # float64 CG recursion, float32 V-cycle (see gpu_jax.py); chosen by accuracy on the 12 systems
JAX_PCG_FORM = "slice"
JAX_PCG_AUTO_MIN_ELEMENTS = 0    # auto picks jax only at or above this many elements (measured faster per optimiser iteration)
JAX_PCG_AUTO_ENABLED = True      # measured faster per solve (1.9-3.4x) and per optimiser iteration (7-21 %, 9 windows): runs/jax
_JAXC = {}                       # per-process cache of the probe and of the per-grid choices
JAX_BYTES_PER_ELEMENT = {("fp64", "slice"): 7500.0, ("fp32", "slice"): 3800.0, ("mixed", "slice"): 6500.0,
                         ("fp64", "stencil"): 9800.0, ("fp32", "stencil"): 4900.0, ("mixed", "stencil"): 8500.0}


def jax_probe():
    """(ok, platform-or-reason): JAX importable with a CUDA device. Imports jax (initialises XLA) on first call, cached."""
    if "jax" not in _JAXC:
        try:
            ok, why = available(require_gpu=True)
            _JAXC["jax"] = (bool(ok), why or "cuda")
        except Exception as e:  # noqa: BLE001
            _JAXC["jax"] = (False, f"{type(e).__name__}: {e}"[:150])
    return _JAXC["jax"]


def pcg_jax_memory(nx, ny, nz, variant=None, form=None):
    """Device bytes of one JAX GMG-PCG solve (x SAFETY): measured peak per element on GH200 (runs/jax/RESULTS.md); the XLA pool
    is not returned to the driver between solves, so this is also the resident size after the first solve."""
    Ne = (nx - 1) * (ny - 1) * (nz - 1)
    return SAFETY * JAX_BYTES_PER_ELEMENT[(variant or JAX_PCG_VARIANT, form or JAX_PCG_FORM)] * Ne


def pcg_choice(requested, nx, ny, nz, hx, hy, hz, element_type, preconditioner, gmg_levels, coarse_max_nodes):
    """(engine, variant, form, reason) for one solve: engine "jax" or "fast". `requested` is "jax" (explicit: used whenever it can
    run) or "auto" (jax only where it is measured faster and its memory fits). A declined request reads ("fast", None, None, why)."""
    if FORCE_CLASSIC:
        return "fast", None, None, "classic forced"
    if ("pcg_jax", nx, ny, nz) in FAILED:
        return "fast", None, None, "jax failed earlier on this grid"
    variant = os.environ.get("PCLS_JAX_VARIANT") or JAX_PCG_VARIANT
    form = os.environ.get("PCLS_JAX_FORM") or JAX_PCG_FORM
    key = ("pcg", requested, nx, ny, nz, float(hx), float(hy), float(hz), element_type, preconditioner, gmg_levels, coarse_max_nodes, variant, form)
    if key in _JAXC:
        return _JAXC[key]
    out = None
    try:
        jax_ok, why = ok(nx, ny, nz, hx, hy, hz, element_type, preconditioner, gmg_levels, coarse_max_nodes, require_gpu=True)
        if not jax_ok:
            out = ("fast", None, None, why)
        elif requested == "jax":
            out = ("jax", variant, form, "requested")
        else:
            cap = probe(gpu=True)
            need = pcg_jax_memory(nx, ny, nz, variant, form)
            ne = (nx - 1) * (ny - 1) * (nz - 1)
            if not JAX_PCG_AUTO_ENABLED:
                out = ("fast", None, None, "auto: jax not enabled as the default (not measured faster per optimiser iteration)")
            elif cap.get("fp64_ratio", 0.0) < DATACENTRE_FP64:
                out = ("fast", None, None, f"auto: FP64:FP32 {cap.get('fp64_ratio')} < {DATACENTRE_FP64} (jax variants keep float64 vectors)")
            elif ne < JAX_PCG_AUTO_MIN_ELEMENTS:
                out = ("fast", None, None, f"auto: {ne} elements < {JAX_PCG_AUTO_MIN_ELEMENTS} (jax not faster below)")
            elif not _fits_device(cap, need):
                out = ("fast", None, None, f"auto: jax needs ~{need / GiB:.1f} GB > {VRAM_FIT_FRACTION:.0%} of "
                                           f"{cap.get('vram_free_GB', 0):.1f} GB free VRAM")
            else:
                out = ("jax", variant, form, f"auto: est. {need / GiB:.2f} GB device")
    except Exception as e:  # noqa: BLE001
        out = ("fast", None, None, f"jax unavailable ({type(e).__name__}: {str(e)[:100]})")
    _JAXC[key] = out
    if record("pcg", requested=requested, selected=out[0], variant=out[1], form=out[2], reason=out[3]):
        print(f"engines: pcg = {out[0]}{'' if out[1] is None else ' (' + out[1] + '/' + out[2] + ')'} ({out[3]}).", flush=True)
    return out


# ============================================================================ record
def record(kind, **kw):
    """Store the selection of `kind` ('elastic' | 'buckling') once; later calls only update counters."""
    ent = RECORD.setdefault(kind, {})
    if not ent:
        ent.update(kw)
        if "capability" not in RECORD and _CAP:
            RECORD["capability"] = {k: v for k, v in _CAP.items() if k != "gpu_probed"}
        return True
    return False


def record_use(kind, engine):
    used = RECORD.setdefault(kind, {}).setdefault("used", {})
    used[engine] = used.get(engine, 0) + 1


def record_fallback(kind, failed, nxt, err):
    RECORD.setdefault("fallbacks", []).append(dict(kind=kind, failed=failed, next=nxt, error=str(err)[:300],
                                                   time=time.strftime("%Y-%m-%dT%H:%M:%S")))


def summary_lines(cap):
    """Human-readable capability line."""
    if cap.get("cuda"):
        g = (f"{cap['gpu_name']} (sm_{cap['cc'][0]}{cap['cc'][1]}), {cap['vram_total_GB']:.1f} GB "
             f"({cap['vram_free_GB']:.1f} free), FP64:FP32 {cap.get('fp64_ratio', float('nan')):.3g}")
    else:
        g = "no CUDA device"
    return (f"{cap['os']} {cap['arch']}, {cap['cpu_cores']} cores, "
            f"{(cap.get('host_ram_GB') or float('nan')):.0f} GB RAM; GPU: {g}; cuDSS: {cap.get('cudss')}; "
            f"host Cholesky: {cap.get('host_cholesky')}")


# ============================================================================ per-iteration update engine (step 3B)
# One separate block (docs/theory.md §18): the choice between the host code and the former gpu_update.py sections for the SIMP / PCLS
# update and diagnostics. Nothing above is touched by it.
UPDATE_ENGINES = ("auto", "gpu", "cpu", "classic")      # optimiser.update_engine
UPDATE_VRAM_FRACTION = 0.30    # share of total VRAM the update engine may take ("auto"); the rest is the elastic solver's
_UPDATE = {"engine": "auto"}   # process default (since the step-3B acceptance, LOG 2026-10-24)


def configure_update(engine=None):
    """Process default of the per-iteration update/diagnostics engine ("cpu" | "gpu" | "auto" | "classic"); set by
    optimise.py from optimiser.update_engine (a process-wide setting, not an optimiser parameter, so parameters.csv
    is unchanged). "classic" is "cpu" (the host code, bit-exact)."""
    if engine is not None:
        if engine not in UPDATE_ENGINES:
            raise ValueError(f"update engine must be one of {UPDATE_ENGINES}, got {engine!r}")
        _UPDATE["engine"] = engine


def update_engine_choice(kind, n_voxels, unsupported=None):
    """("gpu" | "cpu", reason) for the SIMP / PCLS update and diagnostics (the former gpu_update.py sections). FORCE_CLASSIC (the regression
    guard) and "cpu" / "classic" read as the host code; "auto" picks the GPU when a CUDA device exists and the
    estimated device memory (update_memory_bytes) fits UPDATE_VRAM_FRACTION of it; "gpu" insists and raises
    without a device. unsupported: the reason the run's optimiser options have no GPU update (e.g. PCLS
    stiffness_interpolation="ramp"); "auto" and "gpu" then select the host code with that reason, printed and
    recorded like any other choice (plan D3: never a silent path). Recorded once in RECORD["update"] (not under
    FORCE_CLASSIC or "cpu"/"classic", so those runs carry no new run_info key)."""
    if FORCE_CLASSIC:
        return "cpu", "FORCE_CLASSIC"
    req = _UPDATE["engine"]
    if req in ("cpu", "classic"):
        return "cpu", "requested"
    if unsupported:
        eng, why = "cpu", str(unsupported) + "; host update"
        if record("update", requested=req, selected=eng, reason=why, optimiser=kind):
            print(f"engines: update = {eng} ({why}).", flush=True)
        return eng, why
    cap = probe(gpu=True)
    if req == "gpu" and not cap.get("cuda"):
        raise RuntimeError(f"update engine 'gpu' needs a CUDA device ({cap.get('cuda_error', 'none found')}).")
    if not cap.get("cuda") or not cap["packages"].get("numba"):
        eng, why = "cpu", "no CUDA device"
    else:
        from baseline_optimiser import update_memory_bytes
        need = update_memory_bytes("gpu", int(n_voxels), kind) * SAFETY
        if req == "auto" and need > UPDATE_VRAM_FRACTION * cap["vram_total_GB"] * GiB:
            eng, why = "cpu", "device memory model %.2f GB > %.0f%% of %.1f GB VRAM" % (
                need / GiB, 100 * UPDATE_VRAM_FRACTION, cap["vram_total_GB"])
        else:
            eng, why = "gpu", ("requested" if req == "gpu" else "CUDA device") + "; %.2f GB device memory" % (need / GiB)
    if record("update", requested=req, selected=eng, reason=why, optimiser=kind):
        print(f"engines: update = {eng} ({why}).", flush=True)
    return eng, why
