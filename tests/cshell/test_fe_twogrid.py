"""Two-grid PCG engine (cshell/fe_twogrid.py, PLAN_twogrid.md) against the direct engine (fe_sparse)."""
import numpy as np
import pytest
from scipy import ndimage

from cshell.fe_sparse import ActiveGrid, SparseFE, buckling
from cshell.fe_twogrid import LeanGrid, TwoGridFE

E, NU, RHO_MIN = 1760.0, 0.42, 1e-6
BACKENDS = [pytest.param("cpu", id="cpu"), pytest.param("gpu", id="gpu", marks=pytest.mark.gpu)]


def hollow_tube(shape=(21, 9, 11), seed=0, margin=True):
    """Square tube (wall 1-2 elements, odd box dims), clamped at x = 0, axial + transverse tip load; the active set is
    the wall dilated by one layer of ρ_min elements (as in the optimiser), with ρ ∈ [0.3, 1] on the wall."""
    i, j, k = np.indices(shape)
    wall = ~((j >= 2) & (j < shape[1] - 2) & (k >= 2) & (k < shape[2] - 3))
    wall &= (j >= 1) & (k >= 1)                                     # leave a void layer inside the box on two sides
    mask = ndimage.binary_dilation(wall, iterations=1) if margin else wall
    rng = np.random.default_rng(seed)
    rho = np.where(wall, rng.uniform(0.3, 1.0, shape), RHO_MIN)
    nn = tuple(n + 1 for n in shape)
    node_wall = np.zeros(nn, bool)
    for c in np.ndindex(2, 2, 2):
        node_wall[c[0]:c[0] + shape[0], c[1]:c[1] + shape[1], c[2]:c[2] + shape[2]] |= wall
    fixed = np.zeros(nn + (3,), bool)
    fixed[0] = node_wall[0][..., None]
    load = np.zeros(nn + (3,))
    tip = node_wall[-1]
    load[-1][tip] = np.array([-30.0, 4.0, -6.0]) / tip.sum()
    return mask, rho, fixed, load


def engines(mask, rho, fixed, load, backend, **kw):
    gd = ActiveGrid(mask)
    d = SparseFE(gd, fixed, load, 1.0, E, NU, backend=backend, precision="fp64").factor(gd.gather_elems(rho))
    xp = np if backend == "cpu" else __import__("cupy")
    gt = LeanGrid(mask, xp=xp)
    t = TwoGridFE(gt, fixed, load, 1.0, E, NU, backend=backend, **kw).factor(gt.gather_elems(rho))
    return gd, d, gt, t


def h(a):
    return a.get() if hasattr(a, "get") else np.asarray(a)


@pytest.mark.parametrize("backend", BACKENDS)
def test_grid_numbering_matches_activegrid(backend):
    mask, *_ = hollow_tube()
    a = ActiveGrid(mask)
    lg = LeanGrid(mask, xp=np if backend == "cpu" else __import__("cupy"))
    assert np.array_equal(a.node_dense, lg.node_dense) and np.array_equal(a.conn, h(lg.conn))


@pytest.mark.parametrize("backend", BACKENDS)
def test_matrix_free_operators_match_assembled(backend):
    mask, rho, fixed, load = hollow_tube()
    gd, d, gt, t = engines(mask, rho, fixed, load, backend)
    xp = d.xp
    X = xp.asarray(np.random.default_rng(1).standard_normal((gd.ndof, 3)))
    for V in (X[:, 0], X):
        ref = d._K @ V
        assert float(xp.abs(t._K @ V - ref).max() / xp.abs(ref).max()) < 1e-13
    assert float(xp.abs(t._diag - d._K.diagonal()).max() / xp.abs(t._diag).max()) < 1e-14
    # geometric stiffness operator: same element matrices through both engines
    H = xp.asarray(np.random.default_rng(2).standard_normal((gd.Ne, 8, 8)))
    H = H + H.transpose(0, 2, 1)
    Gd, Gt = d._geometric_op(H), t._geometric_op(H)
    for V in (X[:, 0], X):
        ref = Gd @ V
        assert float(xp.abs(Gt @ V - ref).max() / xp.abs(ref).max()) < 1e-13


def test_galerkin_coarse_operator():
    """K_c = Pᵀ K P on free coarse DOF (their P columns never touch a clamped fine DOF); identity on fixed coarse DOF,
    which are exactly those interpolating to a clamped fine DOF."""
    import scipy.sparse as sps
    mask, rho, fixed, load = hollow_tube()
    gd, d, gt, t = engines(mask, rho, fixed, load, "cpu")
    P = sps.kron(t._Pn, sps.identity(3), format="csr")
    fc = t._fixed_c
    assert np.array_equal(fc, np.asarray((P.T @ d.fixed.astype(float)) > 0))
    assert fc.any() and not (abs(P[d.fixed][:, ~fc]).sum() > 0)
    ref = (P.T @ d._K @ P).toarray()
    pat = t._pat
    Kc = sps.csr_matrix((t._coarse_values(), pat.indices, pat.indptr), shape=(pat.ndof, pat.ndof)).toarray()
    ff = np.ix_(~fc, ~fc)
    assert np.abs(Kc[ff] - ref[ff]).max() < 1e-12 * np.abs(ref).max()
    assert np.allclose(Kc[fc][:, fc], np.eye(fc.sum())) and not Kc[fc][:, ~fc].any()
    assert np.linalg.eigvalsh(Kc).min() > 0


@pytest.mark.parametrize("backend", BACKENDS)
def test_twogrid_matches_direct(backend):
    """Plan acceptance 1: C ≤ 1e-8, u ≤ 1e-7, ∂C/∂ρ̂ ≤ 1e-7, λ (6 modes) ≤ 1e-6 relative to the direct engine."""
    mask, rho, fixed, load = hollow_tube()
    gd, d, gt, t = engines(mask, rho, fixed, load, backend)
    pd, pt = d.primal(), t.primal()
    its = t.pcg_summary()["its_exact"]
    print("PCG iterations", its, "coarse DOF", t.stats["coarse_ndof"], "of", t.stats["ndof"])
    assert max(its) < 40
    assert abs(pt.compliance / pd.compliance - 1) < 1e-8
    ud, ut = h(pd.u), h(pt.u)
    assert np.abs(ut - ud).max() < 1e-7 * np.abs(ud).max()
    gdc, gtc = d.dC_drho_hat(pd), t.dC_drho_hat(pt)
    assert np.abs(gtc - gdc).max() < 1e-7 * np.abs(gdc).max()
    bd = buckling(d, pd, num_modes=6, tol=1e-8, maxiter=200)
    bt = buckling(t, pt, num_modes=6, tol=1e-8, maxiter=200)
    n = min(len(bd.mu), 6)
    assert bd.converged and bt.converged and n >= 3 and len(bt.mu) == len(bd.mu)
    np.testing.assert_allclose(bt.mu[:n], bd.mu[:n], rtol=1e-6)
    assert abs(bt.mu_tilde / bd.mu_tilde - 1) < 1e-6
    assert np.abs(bt.dmu_tilde - bd.dmu_tilde).max() < 1e-4 * np.abs(bd.dmu_tilde).max()


@pytest.mark.parametrize("backend", BACKENDS)
def test_warm_start_and_nonconvergence(backend):
    mask, rho, fixed, load = hollow_tube()
    gt = LeanGrid(mask, xp=np if backend == "cpu" else __import__("cupy"))
    t = TwoGridFE(gt, fixed, load, 1.0, E, NU, backend=backend).factor(gt.gather_elems(rho))
    C0 = t.primal().compliance
    cold = t.pcg_summary()["its_exact"][0]
    t.factor(gt.gather_elems(rho * 1.001))                         # nearby design: warm start from the last u
    C1 = t.primal().compliance
    warm = t.pcg_summary()["its_exact"][0]
    assert warm < cold and abs(C1 * 1.001 / C0 - 1) < 1e-8           # K scales by 1.001 exactly (ρ̂ = ρ, all elements)
    t.maxiter, t._u_prev = 2, None
    with pytest.raises(RuntimeError, match="did not converge"):
        t.primal()
