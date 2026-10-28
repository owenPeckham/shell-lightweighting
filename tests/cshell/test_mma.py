"""Task B tests: V5-B4 (Svanberg toy problems), V4-7 (mass restoration), resume, projected L-BFGS (THEORY §7, §11)."""
import numpy as np
import pytest
from scipy.optimize import minimize

from cshell.mma import MMA, ProjectedLBFGSB, restore_mass


def toy(x):
    x = np.asarray(x, float)
    p1, p2 = np.array([5., 2, 1]), np.array([3., 4, 3])
    g = np.array([((x - p1) ** 2).sum() - 9, ((x - p2) ** 2).sum() - 9])
    dg = np.vstack([2 * (x - p1), 2 * (x - p2)])
    return (x ** 2).sum(), 2 * x, g, dg


def beam(x):
    a = np.array([61., 37, 19, 7, 1])
    return 0.0624 * x.sum(), 0.0624 * np.ones(5), np.array([(a / x ** 3).sum() - 1]), (-3 * a / x ** 4)[None, :]


def run(fun, x0, lo, hi, m, iters=11, **kw):
    opt = MMA(len(x0), m, lo, hi, move=1.0, **kw)
    x = np.array(x0, float)
    for _ in range(iters):
        x = opt.step(x, *fun(x))
    return x, fun(x)


def test_svanberg_toy_problem():
    # Svanberg (2007) toy problem; published optimum x = (2.0175, 1.7800, 1.2375), f = 8.7702
    x, (f, _, g, _) = run(toy, [4, 3, 2], 0.0, 5.0, 2)
    assert np.allclose(x, [2.0175, 1.7800, 1.2375], atol=5e-5)
    assert abs(f - 8.7702) < 5e-5 and g.max() < 1e-5
    ref = minimize(lambda z: (z ** 2).sum(), [4, 3, 2], jac=lambda z: 2 * z, bounds=[(0, 5)] * 3, method="SLSQP",
                   constraints=[{"type": "ineq", "fun": lambda z, i=i: -toy(z)[2][i]} for i in range(2)], tol=1e-12)
    assert np.allclose(x, ref.x, atol=1e-4)


def test_svanberg_beam_problem():
    # Svanberg (1987) 5-element cantilever beam; published optimum x = (6.016, 5.309, 4.494, 3.502, 2.153), f = 1.340
    x, (f, _, g, _) = run(beam, [5.0] * 5, 1.0, 10.0, 1)
    assert np.allclose(x, [6.0160, 5.3092, 4.4943, 3.5015, 2.1527], atol=5e-4)
    assert abs(f - 1.3400) < 5e-4 and abs(g[0]) < 1e-4


def test_state_roundtrip_bit_exact_resume():
    x0 = np.array([4., 3, 2])
    a = MMA(3, 2, 0.0, 5.0, move=1.0)
    x = x0.copy()
    for _ in range(4):
        x = a.step(x, *toy(x))
    saved = {k: v.copy() for k, v in a.state_dict().items()}
    xa = x.copy()
    for _ in range(4):
        xa = a.step(xa, *toy(xa))
    b = MMA(3, 2, 0.0, 5.0, move=1.0)
    b.load_state_dict(saved)
    xb = x.copy()
    for _ in range(4):
        xb = b.step(xb, *toy(xb))
    assert np.array_equal(xa, xb)
    a.reset_asymptotes()
    assert a.state.iteration == 0


def test_oscillation_halves_move():
    o = MMA(2, 0, 0.0, 1.0, move=0.1)
    o.step(np.zeros(2), 0.0, np.ones(2))
    flags = [o.note_objective(J) for J in [1, 2, 1, 2, 1, 2, 1, 2, 1, 2]]
    assert any(flags) and o.state.move == 0.05


def make_volume(n=400, seed=0):
    rng = np.random.default_rng(seed)
    idx = np.arange(n)
    F = np.exp(-0.5 * ((idx[:, None] - idx[None, :]) / 3.0) ** 2)      # positive Gaussian filter
    F /= F.sum(1, keepdims=True)
    w = rng.uniform(0.5, 2.0, n)
    s = lambda z: 1 / (1 + np.exp(-8 * (z - 0.5)))
    ds = lambda z: 8 * s(z) * (1 - s(z))
    vol = lambda xi: float(w @ s(F @ xi))
    grad = lambda xi: F.T @ (w * ds(F @ xi))
    return vol, grad, n, rng


@pytest.mark.parametrize("seed", range(5))
def test_restore_mass_V47(seed):
    vol, grad, n, rng = make_volume(seed=seed)
    xi = rng.uniform(0.2, 0.8, n)
    V0 = vol(xi)
    xi_plus = np.clip(xi + 0.2 * rng.uniform(-1, 1, n), 0, 1)     # MMA-like step breaks the mass
    assert abs(vol(xi_plus) / V0 - 1) > 1e-3
    active = np.ones(n, bool); active[:20] = False
    xr, eta, nev = restore_mass(xi_plus, vol, grad, V0, active)
    assert abs(vol(xr) / V0 - 1) < 1e-4 and nev <= 20
    assert 0 <= xr.min() and xr.max() <= 1


def test_restore_mass_no_bracket():
    vol = lambda xi: float(xi.sum())
    with pytest.raises(RuntimeError):
        restore_mass(np.full(5, 0.5), vol, lambda xi: np.ones(5), V0=10.0)


def test_projected_lbfgsb_box_and_mass():
    rng = np.random.default_rng(1)
    n = 40
    t = rng.uniform(-0.3, 1.3, n)
    Q = np.diag(np.linspace(1, 20, n))
    J = lambda x: 0.5 * (x - t) @ Q @ (x - t)
    dJ = lambda x: Q @ (x - t)
    opt = ProjectedLBFGSB(n, xmin=0.0, xmax=1.0, move=0.2)
    x = np.full(n, 0.5)
    for _ in range(80):
        x = opt.step(x, J(x), dJ(x))
    assert np.allclose(x, np.clip(t, 0, 1), atol=1e-6)
    # with constant mass sum(x) = V0 held by restore_mass
    vol = lambda z: float(z.sum()); gr = lambda z: np.ones(n)
    V0 = 20.0
    ref = minimize(J, np.full(n, 0.5), jac=dJ, bounds=[(0, 1)] * n, method="SLSQP",
                   constraints=[{"type": "eq", "fun": lambda z: z.sum() - V0}], tol=1e-14)
    opt = ProjectedLBFGSB(n, xmin=0.0, xmax=1.0, move=0.2)
    x = np.full(n, 0.5)
    for _ in range(150):
        x = opt.step(x, J(x), dJ(x), dg=np.ones((1, n)))
        x, _, _ = restore_mass(x, vol, gr, V0, tol=1e-10)
    assert abs(x.sum() - V0) < 1e-8
    assert J(x) < J(ref.x) * 1.01 + 1e-8
