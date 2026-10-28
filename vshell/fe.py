"""
Voxel finite-element model (linear elasticity) on a regular grid of cubic cells.

Element: 8-node hexahedron with Wilson-Taylor incompatible modes (the same formulation as
CalculiX C3D8I / Abaqus C3D8I). On a cube the 9 internal modes are condensed out once, so every
element shares one 24x24 stiffness matrix, scaled per element by E_e * h.

Partially filled boundary cells use an iso-strain (Voigt) mixture: E_e = E * vf_e, where vf_e is
the cell's solid volume fraction from the signed-distance field. Strains at the cell centre are
exact for the incompatible-mode element (the bubble-mode gradients vanish at xi = 0), and the
stress in the *material* of a partial cell is E * D * eps (iso-strain), which is what strength
checks use; the *homogenised* stress E_e * D * eps integrates to the true section resultants.

Solver: smoothed-aggregation algebraic multigrid (pyamg) with the six rigid-body modes as the
near-null space, as a preconditioner for conjugate gradients.
"""
from __future__ import annotations

import time
import numpy as np
import scipy.sparse as sp
from scipy import ndimage

# Node order of a cell (i, j, k offsets) -- identical to CalculiX / Abaqus C3D8 ordering.
CELL_OFFSETS = np.array([[0, 0, 0], [1, 0, 0], [1, 1, 0], [0, 1, 0],
                         [0, 0, 1], [1, 0, 1], [1, 1, 1], [0, 1, 1]], dtype=np.int64)
_XI = 2.0 * CELL_OFFSETS - 1.0          # natural coordinates of the nodes


def elasticity_matrix(E: float, nu: float) -> np.ndarray:
    """Isotropic D for engineering strains [exx, eyy, ezz, gxy, gyz, gzx]."""
    lam = E * nu / ((1 + nu) * (1 - 2 * nu))
    mu = E / (2 * (1 + nu))
    D = np.zeros((6, 6))
    D[:3, :3] = lam
    D[np.arange(3), np.arange(3)] = lam + 2 * mu
    D[3, 3] = D[4, 4] = D[5, 5] = mu
    return D


def _b_matrix(dN: np.ndarray) -> np.ndarray:
    """Strain-displacement matrix from shape-function gradients dN (n, 3) -> (6, 3n)."""
    n = dN.shape[0]
    B = np.zeros((6, 3 * n))
    for i in range(n):
        dx, dy, dz = dN[i]
        c = 3 * i
        B[0, c] = dx
        B[1, c + 1] = dy
        B[2, c + 2] = dz
        B[3, c] = dy; B[3, c + 1] = dx
        B[4, c + 1] = dz; B[4, c + 2] = dy
        B[5, c] = dz; B[5, c + 2] = dx
    return B


def _grad_N(xi, eta, zeta, h=1.0):
    """Gradients (8, 3) of the trilinear shape functions on a cube of side h."""
    g = np.empty((8, 3))
    for i, (a, b, c) in enumerate(_XI):
        g[i, 0] = a * (1 + b * eta) * (1 + c * zeta) / 8
        g[i, 1] = b * (1 + a * xi) * (1 + c * zeta) / 8
        g[i, 2] = c * (1 + a * xi) * (1 + b * eta) / 8
    return g * (2.0 / h)


def _grad_bubbles(xi, eta, zeta, h=1.0):
    """Gradients of the three incompatible modes (1-xi^2), (1-eta^2), (1-zeta^2)."""
    return np.diag([-2 * xi, -2 * eta, -2 * zeta]) * (2.0 / h)


def unit_element_stiffness(nu: float, incompatible: bool = True) -> np.ndarray:
    """24x24 stiffness of a unit cube (E = 1, h = 1). Scale by E_e * h for a real cell."""
    D = elasticity_matrix(1.0, nu)
    gp = np.array([-1.0, 1.0]) / np.sqrt(3.0)
    detJ = 0.5 ** 3
    Kcc = np.zeros((24, 24)); Kci = np.zeros((24, 9)); Kii = np.zeros((9, 9))
    for xi in gp:
        for eta in gp:
            for zeta in gp:
                Bc = _b_matrix(_grad_N(xi, eta, zeta))
                Bi = _b_matrix(_grad_bubbles(xi, eta, zeta))
                Kcc += Bc.T @ D @ Bc * detJ
                Kci += Bc.T @ D @ Bi * detJ
                Kii += Bi.T @ D @ Bi * detJ
    if not incompatible:
        return Kcc
    return Kcc - Kci @ np.linalg.solve(Kii, Kci.T)


def centre_b_matrix(h: float) -> np.ndarray:
    """B at the cell centre (exact strain for the incompatible-mode element)."""
    return _b_matrix(_grad_N(0.0, 0.0, 0.0, h))


class VoxelMesh:
    """Cells, nodes and connectivity for the active cells of a grid.

    vf      : (nx, ny, nz) solid volume fraction per cell
    origin  : coordinates of node (0, 0, 0)
    h       : cell size
    vf_min  : cells below this fraction are dropped
    keep_component_of : optional (m, 3) node coordinates; only the face-connected component(s)
                        of cells touching these points are kept (removes floating debris).
    """

    def __init__(self, vf: np.ndarray, origin, h: float, vf_min: float = 0.05):
        self.h = float(h)
        self.origin = np.asarray(origin, float)
        mask = vf >= vf_min
        lab, n = ndimage.label(mask)                         # 6-connected (face) components
        if n > 1:
            sizes = ndimage.sum(mask, lab, range(1, n + 1))
            mask = lab == (1 + int(np.argmax(sizes)))
            self.dropped_cells = int(sizes.sum() - sizes.max())
        else:
            self.dropped_cells = 0
        self.cell_ijk = np.argwhere(mask)
        self.vf = vf[mask].astype(np.float64)
        self.grid_shape = vf.shape
        corners = self.cell_ijk[:, None, :] + CELL_OFFSETS[None]
        flat = corners.reshape(-1, 3)
        node_shape = np.array(vf.shape) + 1
        lin = np.ravel_multi_index(flat.T, node_shape)
        uniq, inv = np.unique(lin, return_inverse=True)
        self.node_ijk = np.stack(np.unravel_index(uniq, node_shape), 1)
        self.conn = inv.reshape(-1, 8)
        self.X = self.origin + self.h * self.node_ijk
        self.cell_centres = self.origin + self.h * (self.cell_ijk + 0.5)

    @property
    def n_nodes(self):
        return len(self.node_ijk)

    @property
    def n_cells(self):
        return len(self.cell_ijk)


def assemble(mesh: VoxelMesh, E: float, nu: float, chunk: int = 20000) -> sp.csr_matrix:
    Ke = unit_element_stiffness(nu)
    scale = E * mesh.vf * mesh.h
    ndof = 3 * mesh.n_nodes
    edof = (3 * mesh.conn[:, :, None] + np.arange(3)[None, None, :]).reshape(-1, 24).astype(np.int32)
    K = sp.csr_matrix((ndof, ndof))
    ii = np.repeat(np.arange(24), 24); jj = np.tile(np.arange(24), 24)
    kflat = Ke.ravel()
    for s in range(0, mesh.n_cells, chunk):
        e = edof[s:s + chunk]
        rows = e[:, ii].ravel(); cols = e[:, jj].ravel()
        vals = (scale[s:s + chunk, None] * kflat[None, :]).ravel()
        K = K + sp.csr_matrix((vals, (rows, cols)), shape=(ndof, ndof))
    K.sum_duplicates()
    return K


def rigid_body_modes(X: np.ndarray) -> np.ndarray:
    c = X - X.mean(0)
    n = len(X)
    B = np.zeros((3 * n, 6))
    for d in range(3):
        B[d::3, d] = 1.0
    x, y, z = c.T
    B[0::3, 3] = -y; B[1::3, 3] = x          # rotation about z
    B[1::3, 4] = -z; B[2::3, 4] = y          # rotation about x
    B[0::3, 5] = z;  B[2::3, 5] = -x         # rotation about y
    return B


def solve(mesh: VoxelMesh, E: float, nu: float, fixed_nodes: np.ndarray, load_nodes: np.ndarray,
          force: np.ndarray, tol: float = 1e-8, maxiter: int = 800, verbose: bool = False,
          fixed_dofs: np.ndarray | None = None) -> dict:
    """Static solve with all DOF of fixed_nodes clamped (or exactly the global `fixed_dofs`, 3*node + axis)
    and `force` (N, 3-vector) shared equally by load_nodes."""
    import pyamg
    t0 = time.time()
    K = assemble(mesh, E, nu)
    t_asm = time.time() - t0
    ndof = K.shape[0]
    f = np.zeros(ndof)
    load_nodes = np.asarray(load_nodes, int)
    for d in range(3):
        f[3 * load_nodes + d] += force[d] / len(load_nodes)
    fixed = np.zeros(ndof, bool)
    if fixed_dofs is not None:
        fixed[np.asarray(fixed_dofs, int)] = True
    else:
        for d in range(3):
            fixed[3 * np.asarray(fixed_nodes, int) + d] = True
    free = np.where(~fixed)[0]
    Kff = K[free][:, free].tocsr()
    del K
    Bn = rigid_body_modes(mesh.X)[free]
    t1 = time.time()
    ml = pyamg.smoothed_aggregation_solver(Kff, B=Bn, symmetry='hermitian', max_coarse=1000)
    t_setup = time.time() - t1
    res = []
    uf = ml.solve(f[free], tol=tol, accel='cg', maxiter=maxiter, residuals=res)
    u = np.zeros(ndof); u[free] = uf
    t_solve = time.time() - t1 - t_setup
    rel = res[-1] / res[0] if res and res[0] > 0 else 0.0
    out = dict(u=u.reshape(-1, 3), compliance=float(f @ u), iterations=len(res) - 1, rel_residual=float(rel),
               ndof=ndof, n_free=len(free), t_assemble=t_asm, t_amg_setup=t_setup, t_solve=t_solve,
               converged=bool(rel <= tol * 10))
    if verbose:
        print(f"    FE: {mesh.n_cells} cells, {ndof} DOF, asm {t_asm:.1f}s, AMG {t_setup:.1f}s, "
              f"CG {t_solve:.1f}s ({out['iterations']} it, rel res {rel:.1e})")
    return out


def cell_stresses(mesh: VoxelMesh, u: np.ndarray, E: float, nu: float):
    """(strain, material stress, homogenised stress) at cell centres, each (n_cells, 6)."""
    B0 = centre_b_matrix(mesh.h)
    ue = u[mesh.conn].reshape(-1, 24)
    eps = ue @ B0.T
    D = elasticity_matrix(E, nu)
    sig_mat = eps @ D.T
    sig_hom = sig_mat * mesh.vf[:, None]
    return eps, sig_mat, sig_hom


def von_mises(s: np.ndarray) -> np.ndarray:
    sx, sy, sz, txy, tyz, tzx = s.T
    return np.sqrt(0.5 * ((sx - sy) ** 2 + (sy - sz) ** 2 + (sz - sx) ** 2) + 3 * (txy ** 2 + tyz ** 2 + tzx ** 2))


def tensor(s: np.ndarray) -> np.ndarray:
    """Voigt (n, 6) [xx, yy, zz, xy, yz, zx] -> (n, 3, 3)."""
    T = np.empty((len(s), 3, 3))
    T[:, 0, 0], T[:, 1, 1], T[:, 2, 2] = s[:, 0], s[:, 1], s[:, 2]
    T[:, 0, 1] = T[:, 1, 0] = s[:, 3]
    T[:, 1, 2] = T[:, 2, 1] = s[:, 4]
    T[:, 0, 2] = T[:, 2, 0] = s[:, 5]
    return T


def strain_energy_density(sig_hom: np.ndarray, eps: np.ndarray) -> np.ndarray:
    """Per-cell strain energy density of the homogenised cell, 0.5 * sigma : eps (engineering shear)."""
    return 0.5 * np.einsum('ij,ij->i', sig_hom, eps)
