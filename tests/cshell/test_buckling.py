"""Sparse-engine buckling (THEORY §4.4(c), §6.4, §6.6(5)): reference eigenvalues and directional sensitivity."""
import numpy as np
import pytest

from cshell.fe_sparse import ActiveGrid, SparseFE, buckling
from cshell.fea import ShellCG

E, NU = 1760.0, 0.42
BACKENDS = [pytest.param("cpu", id="cpu"), pytest.param("gpu", id="gpu", marks=pytest.mark.gpu)]


ENGINES = [pytest.param("direct", "cpu", id="direct-cpu"), pytest.param("direct", "gpu", id="direct-gpu", marks=pytest.mark.gpu),
           pytest.param("twogrid", "cpu", id="twogrid-cpu"),
           pytest.param("twogrid", "gpu", id="twogrid-gpu", marks=pytest.mark.gpu)]


def setup(rho, backend, element_type="q1_im", engine="direct"):
    nn = tuple(n + 1 for n in rho.shape)
    fixed = np.zeros(nn + (3,), bool); fixed[0] = True
    load = np.zeros(nn + (3,)); load[-1] = np.array([-40.0, 0.0, -3.0]) / (nn[1] * nn[2])   # axial + transverse
    g = ActiveGrid(np.ones(rho.shape, bool))
    if engine == "twogrid":
        from cshell.fe_twogrid import TwoGridFE
        return g, TwoGridFE(g, fixed, load, 1.0, E, NU, backend=backend, element_type=element_type), fixed, load
    return g, SparseFE(g, fixed, load, 1.0, E, NU, backend=backend, element_type=element_type), fixed, load


def run(fe, g, rho, **kw):
    fe.factor(g.gather_elems(rho))
    return buckling(fe, fe.primal(), num_modes=4, p=8.0, tol=1e-7, maxiter=100, **kw)


def test_eigenvalues_match_reference():
    """Same discrete model as analysis.ConjugateGradient.run_buckling (q1_im, host engine)."""
    rng = np.random.default_rng(3)
    rho = rng.uniform(0.5, 1.0, (30, 4, 5))
    g, fe, fixed, load = setup(rho, "cpu")
    mine = run(fe, g, rho)
    cg = ShellCG(fixed, load, 1.0, E, NU, rel_tol=1e-12, engine="cpu_direct", warm_start=False)
    cg.solve(rho.astype(np.float32))
    ref = cg.run_buckling(num_modes=4, aggregation_p=8.0, engine="cpu_cholesky_lobpcg", lobpcg_tol=1e-9)
    np.testing.assert_allclose(np.sort(mine.mu)[::-1][:3], np.sort(ref.mu_values)[::-1][:3], rtol=2e-4)


@pytest.mark.parametrize("engine,backend", ENGINES)
def test_directional_sensitivity_q1im(engine, backend):
    """THEORY §6.4 set-up: smooth δρ̂ = 0.05 sin(πx)(1 + 0.3 noise), central differences, < 1e-3."""
    rng = np.random.default_rng(4)
    rho = rng.uniform(0.5, 1.0, (30, 4, 5))
    g, fe, _, _ = setup(rho, backend, engine=engine)
    base = run(fe, g, rho)
    x = (np.arange(30) + 0.5) / 30
    d = 0.05 * np.sin(np.pi * x)[:, None, None] * (1 + 0.3 * rng.standard_normal(rho.shape))
    pred = float(base.dmu_tilde @ g.gather_elems(d))
    for h, tol in [(0.3, 2e-2), (0.1, 1e-3)]:
        fd = (run(fe, g, rho + h * d, X0=base.modes).mu_tilde - run(fe, g, rho - h * d, X0=base.modes).mu_tilde) / (2 * h)
        err = abs(fd / pred - 1)
        print(f"h={h}: fd={fd:.6e} pred={pred:.6e} rel={err:.2e}")
    assert err < tol


@pytest.mark.parametrize("engine,backend", ENGINES)
def test_directional_sensitivity_ksig_mask(engine, backend):
    """K_σ excluded near the clamp (first 3 element layers): μ̃ changes, and ∂μ̃ stays exact (< 1e-3 directional)."""
    rng = np.random.default_rng(5)
    rho = rng.uniform(0.5, 1.0, (30, 4, 5))
    g, fe, _, _ = setup(rho, backend, engine=engine)
    mask = np.ones(rho.shape); mask[:3] = 0.0
    km = g.gather_elems(mask)
    full = run(fe, g, rho)
    base = run(fe, g, rho, ksig_mask=km)
    assert base.mu_tilde < full.mu_tilde                      # less geometric stiffness -> lower μ (higher λ_b)
    x = (np.arange(30) + 0.5) / 30
    d = 0.05 * np.sin(np.pi * x)[:, None, None] * (1 + 0.3 * rng.standard_normal(rho.shape))
    pred = float(base.dmu_tilde @ g.gather_elems(d))
    h = 0.1
    fd = (run(fe, g, rho + h * d, X0=base.modes, ksig_mask=km).mu_tilde
          - run(fe, g, rho - h * d, X0=base.modes, ksig_mask=km).mu_tilde) / (2 * h)
    err = abs(fd / pred - 1)
    print(f"masked: mu {base.mu_tilde:.6e} (unmasked {full.mu_tilde:.6e}) fd={fd:.6e} pred={pred:.6e} rel={err:.2e}")
    assert err < 1e-3


@pytest.mark.parametrize("fix", [False, True])
def test_reference_q1im_fix(fix):
    """THEORY §6.4 reproducer on the reference code (fast host engine): bias ≈ 5.8 % unfixed, < 1e-3 with the fix."""
    rng = np.random.default_rng(7)
    rho = rng.uniform(0.5, 1.0, (30, 4, 5))
    nn = tuple(n + 1 for n in rho.shape)
    fixed = np.zeros(nn + (3,), bool); fixed[0] = True
    load = np.zeros(nn + (3,)); load[-1] = np.array([-40.0, 0.0, -3.0]) / (nn[1] * nn[2])
    cg = ShellCG(fixed, load, 1.0, E, NU, rel_tol=1e-12, engine="cpu_direct", warm_start=False)

    def mu(r, sens=False):
        cg.solve(r.astype(np.float32))
        b = cg.run_buckling(num_modes=4, aggregation_p=8.0, engine="cpu_cholesky_lobpcg", lobpcg_tol=1e-10,
                            fix_q1im_sensitivity=fix)
        return (b.mu_tilde, np.asarray(b.sensitivity)) if sens else b.mu_tilde

    m0, s0 = mu(rho, True)
    x = (np.arange(30) + 0.5) / 30
    d = 0.05 * np.sin(np.pi * x)[:, None, None] * (1 + 0.3 * rng.standard_normal(rho.shape))
    pred = float((s0 * d).sum())
    h = 0.1
    fd = (mu(rho + h * d) - mu(rho - h * d)) / (2 * h)
    err = abs(fd / pred - 1)
    print(f"fix={fix}: fd={fd:.6e} pred={pred:.6e} rel={err:.2e}")
    if fix:
        assert err < 1e-3
    else:
        assert err > 1e-2          # documents the reference bias (THEORY §6.4: ≈ 5.8 %)
