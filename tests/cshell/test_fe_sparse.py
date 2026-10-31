"""Active-set cuDSS/host-Cholesky engine against the reference ShellCG (STATUS.md §2.2, 0.3) and §6.6(3)."""
import numpy as np
import pytest

from cshell.fe_sparse import ActiveGrid, SparseFE, centre_B, unit_D
from cshell.fea import ShellCG

E, NU, RHO_MIN = 1760.0, 0.42, 1e-6
BACKENDS = [pytest.param("cpu", id="cpu"), pytest.param("gpu", id="gpu", marks=pytest.mark.gpu)]
ENGINES = [pytest.param("direct", "cpu", id="direct-cpu"), pytest.param("direct", "gpu", id="direct-gpu", marks=pytest.mark.gpu),
           pytest.param("twogrid", "cpu", id="twogrid-cpu"),
           pytest.param("twogrid", "gpu", id="twogrid-gpu", marks=pytest.mark.gpu)]


def make_fe(engine, mask, fixed, load, backend):
    """(grid, engine) for the direct cuDSS/host-Cholesky engine or the two-grid PCG engine (fe_twogrid)."""
    g = ActiveGrid(mask)
    if engine == "direct":
        return g, SparseFE(g, fixed, load, 1.0, E, NU, backend=backend)
    from cshell.fe_twogrid import TwoGridFE
    return g, TwoGridFE(g, fixed, load, 1.0, E, NU, backend=backend)


def column(shape_e, P=(-3.0, 2.0, -10.0)):
    nn = tuple(n + 1 for n in shape_e)
    fixed = np.zeros(nn + (3,), bool); fixed[0] = True
    load = np.zeros(nn + (3,)); load[-1] = np.asarray(P) / (nn[1] * nn[2])
    return fixed, load


def ref(rho, fixed, load):
    return ShellCG(fixed, load, 1.0, E, NU, rel_tol=1e-12, engine="cpu_direct", warm_start=False).solve(rho)


@pytest.mark.parametrize("backend", BACKENDS)
def test_full_set_matches_shellcg(backend):
    rng = np.random.default_rng(0)
    rho = rng.uniform(0.3, 1.0, (20, 4, 5))
    fixed, load = column(rho.shape)
    g = ActiveGrid(np.ones(rho.shape, bool))
    fe = SparseFE(g, fixed, load, 1.0, E, NU, backend=backend).factor(g.gather_elems(rho))
    p = fe.primal()
    r = ref(rho, fixed, load)
    assert abs(p.compliance / r.compliance - 1) < 2e-6
    np.testing.assert_allclose(g.scatter_elems(p.U_e), r.U_e, rtol=1e-4, atol=1e-9 * r.U_e.max())
    assert abs(p.U_e.sum() / (0.5 * p.compliance) - 1) < 1e-10


@pytest.mark.parametrize("backend", BACKENDS)
def test_partial_set_matches_rho_min_box(backend):
    """A hollow-ish column: elements outside the set are ρ_min in the dense reference; difference O(ρ_min)."""
    shape = (24, 7, 7)
    i, j, k = np.indices(shape)
    solid = (np.abs(j - 3) <= 1) & (np.abs(k - 3) <= 3) | (np.abs(k - 3) <= 1) & (np.abs(j - 3) <= 3)   # cross
    rho = np.where(solid, 1.0, RHO_MIN)
    fixed, load = column(shape, P=(0.0, 0.0, -10.0))
    load[-1] = 0.0
    load[-1, 2:6, 2:6] = np.array([0.0, 0.0, -10.0]) / 16       # load on the cross only
    g = ActiveGrid(solid)
    fe = SparseFE(g, fixed, load, 1.0, E, NU, backend=backend).factor(g.gather_elems(rho))
    p = fe.primal()
    r = ref(rho, fixed, load)
    assert abs(p.compliance / r.compliance - 1) < 1e-5


@pytest.mark.parametrize("engine,backend", ENGINES)
def test_compliance_gradient_fd(engine, backend):
    rng = np.random.default_rng(2)
    rho = rng.uniform(0.5, 1.0, (30, 4, 4))
    fixed, load = column(rho.shape)
    g, fe = make_fe(engine, np.ones(rho.shape, bool), fixed, load, backend)
    r = g.gather_elems(rho)
    grad = fe.dC_drho_hat(fe.factor(r).primal())
    for e in [5, 200, 470]:
        h = 2e-3                     # solver round-off ~1e-10·C dominates below ~1e-3
        rp, rm = r.copy(), r.copy(); rp[e] += h; rm[e] -= h
        fd = (fe.factor(rp).primal().compliance - fe.factor(rm).primal().compliance) / (2 * h)
        assert abs(fd / grad[e] - 1) < 1e-4, (e, fd, grad[e])        # float64; spec asks < 1e-3


def test_centre_stress_voigt_order():
    """Uniform strain field u = ε x: centre stress equals D ε in (xx, yy, zz, yz, xz, xy) order."""
    eps = np.array([1e-3, -2e-3, 0.5e-3, 3e-3, -1e-3, 2e-3])        # engineering shears
    Ex = np.array([[eps[0], eps[5] / 2, eps[4] / 2], [eps[5] / 2, eps[1], eps[3] / 2], [eps[4] / 2, eps[3] / 2, eps[2]]])
    h = 0.4
    from cshell.fe_sparse import CORNERS
    X = CORNERS * h
    ue = (X @ Ex.T).reshape(-1)
    np.testing.assert_allclose(centre_B(h) @ ue, eps, rtol=1e-12, atol=1e-15)
    sig = E * unit_D(NU) @ eps
    lam = E * NU / ((1 + NU) * (1 - 2 * NU)); mu = E / (2 * (1 + NU))
    assert np.isclose(sig[0], lam * eps[:3].sum() + 2 * mu * eps[0]) and np.isclose(sig[3], mu * eps[3])


@pytest.mark.gpu
def test_fp32_preconditioned_matches_fp64():
    rng = np.random.default_rng(9)
    rho = rng.uniform(0.3, 1.0, (24, 5, 6))
    fixed, load = column(rho.shape)
    g = ActiveGrid(np.ones(rho.shape, bool))
    p64 = SparseFE(g, fixed, load, 1.0, E, NU, backend="gpu", precision="fp64").factor(g.gather_elems(rho)).primal()
    fe32 = SparseFE(g, fixed, load, 1.0, E, NU, backend="gpu", precision="fp32").factor(g.gather_elems(rho))
    p32 = fe32.primal()
    assert abs(p32.compliance / p64.compliance - 1) < 1e-10
    np.testing.assert_allclose(p32.U_e, p64.U_e, rtol=1e-8, atol=1e-14)
