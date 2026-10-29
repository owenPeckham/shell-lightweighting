"""Stress p-norm adjoint (THEORY §6.6(4)) and the J gradient through capacity soft-min (§6.5)."""
import numpy as np
import pytest

from cshell.fe_sparse import ActiveGrid, SparseFE
from cshell.objective import Settings, capacity, evaluate, stress_pnorm

E, NU = 1760.0, 0.42
BACKENDS = [pytest.param("cpu", id="cpu"), pytest.param("gpu", id="gpu", marks=pytest.mark.gpu)]


ENGINES = [pytest.param("direct", "cpu", id="direct-cpu"), pytest.param("direct", "gpu", id="direct-gpu", marks=pytest.mark.gpu),
           pytest.param("twogrid", "cpu", id="twogrid-cpu"),
           pytest.param("twogrid", "gpu", id="twogrid-gpu", marks=pytest.mark.gpu)]


def notched_bar(backend, engine="direct"):
    """40×6×6 bar (h = 1 mm), clamped at x = 0, tip shear + axial; a notch of empty elements at mid-length."""
    shape = (40, 6, 6)
    mask = np.ones(shape, bool)
    mask[19:21, :, 4:] = False
    nn = tuple(n + 1 for n in shape)
    fixed = np.zeros(nn + (3,), bool); fixed[0] = True
    load = np.zeros(nn + (3,)); load[-1] = np.array([5.0, 0.0, -10.0]) / (nn[1] * nn[2])
    g = ActiveGrid(mask)
    if engine == "twogrid":
        from cshell.fe_twogrid import TwoGridFE
        fe = TwoGridFE(g, fixed, load, 1.0, E, NU, backend=backend)
    else:
        fe = SparseFE(g, fixed, load, 1.0, E, NU, backend=backend)
    m = np.ones(g.Ne); m[g.gather_elems(np.indices(shape)[0]) < 3] = 0.0     # exclusion zone at the clamp
    return g, fe, m


@pytest.mark.parametrize("engine,backend", ENGINES)
def test_stress_adjoint_directional(engine, backend):
    g, fe, m = notched_bar(backend, engine)
    rng = np.random.default_rng(5)
    rho = rng.uniform(0.6, 1.0, g.Ne)
    x = g.gather_elems(np.indices(g.elem_shape)[0]) / 40
    d = 0.05 * np.sin(np.pi * x) * (1 + 0.3 * rng.standard_normal(g.Ne))

    def f(r):
        fe.factor(r)
        return stress_pnorm(fe, fe.primal(), r, m, p=8.0, q=2.0)       # ρ̂ ≈ ρ (ρ_min negligible here)

    s0 = f(rho)
    pred = float(s0.grad @ d)
    h = 1e-3
    fd = (f(rho + h * d).value - f(rho - h * d).value) / (2 * h)
    assert abs(fd / pred - 1) < 1e-3, (fd, pred)


def test_capacity_softmin_limits():
    lam, w = capacity(10.0, 0.0, 0.85, 12.0)
    assert np.isclose(lam, 10.0) and np.allclose(w, [1.0])
    lam, w = capacity(10.0, 1 / 2.0, 0.85, 12.0)                         # λ_b = 1.7 governs
    assert 1.69 < lam <= 1.7 and w[1] > 0.99


@pytest.mark.parametrize("engine,backend", ENGINES)
def test_J_gradient_directional(engine, backend):
    """∂J/∂ρ̂ along a smooth direction, with buckling active (axial compression); < 1e-3."""
    g, fe, m = notched_bar(backend, engine)
    fe.f[:] = 0
    nn = g.node_shape
    rng = np.random.default_rng(6)
    rho = rng.uniform(0.6, 1.0, g.Ne)
    st = Settings(knockdown=0.85, beta=12.0, eig_tol=1e-7, num_modes=4)
    # compressive tip load so buckling and strength both matter
    load = np.zeros(nn + (3,)); load[-1] = np.array([-400.0, 0.0, -20.0]) / (nn[1] * nn[2])
    fe.f = fe.xp.asarray(g.gather_nodes(load).reshape(-1))
    base = evaluate(fe, rho, rho, m, st)
    base_d = dict(C0=base.C * 1.1, lam0=base.lam * 0.9)
    e0 = evaluate(fe, rho, rho, m, st, base_d)
    x = g.gather_elems(np.indices(g.elem_shape)[0]) / 40
    d = 0.05 * np.sin(np.pi * x) * (1 + 0.3 * rng.standard_normal(g.Ne))
    pred = float(e0.dJ @ d)
    h = 1e-2
    Jp = evaluate(fe, rho + h * d, rho + h * d, m, st, base_d).J
    Jm = evaluate(fe, rho - h * d, rho - h * d, m, st, base_d).J
    fd = (Jp - Jm) / (2 * h)
    print("J", e0.J, "shares", e0.parts["shares"], "fd", fd, "pred", pred)
    assert abs(fd / pred - 1) < 1e-3, (fd, pred)
