"""THEORY §6.6(6): dJ/dξ along a smooth direction on a small rod (B3 column, h_e = 0.8 mm), including filter, D1
bounds, as-built step, P, ψ → χ → ρ, FE (compliance, stress p-norm, buckling) and the volume gradient. < 2 %."""
from pathlib import Path

import numpy as np
import pytest

ROOT = Path(__file__).resolve().parents[2]


@pytest.fixture(scope="module")
def model():
    from cshell import config
    from cshell.optimise import Context, PhaseModel
    cfg = config.load(str(ROOT / "examples/cshell/bench/B3_column.json"),
                      {"grid": {"phases": [{"h_e": 0.8, "h_g": 0.2}]}, "fe": {"backend": "auto"},
                       "capacity": {"eig_tol": 1e-7}})
    ctx = Context(cfg)
    return ctx, PhaseModel(ctx, 0, 4.0)


def test_dJ_dxi_directional(model):
    ctx, m = model
    nv = ctx.surf.n_vertices
    dm = m.dm
    st = m.settings(0.5)
    base = m.baseline(st)
    c = base["c"]
    rng = np.random.default_rng(0)
    xi = np.concatenate([np.full(nv, 0.4), np.full(nv, 0.3)])
    filt = m.ds.filt
    d = np.concatenate([filt.apply(rng.standard_normal(nv)), filt.apply(rng.standard_normal(nv))])
    d *= 0.05 / np.abs(d).max()

    def J(x, grad=False):
        s = dm.forward(x[:nv], x[nv:], as_built=True)
        rho = dm.rho(s)
        fe = m.fe_eval(rho, st, base, c)
        if not grad:
            return fe["J"]
        g = np.zeros(m.disc.n_elem); g[m.grid_fe.elem_dense] = fe["dJ"]
        gd, ga = dm.pull_rho_hat(s, g)
        return fe["J"], np.concatenate([gd, ga])

    m.ensure_active(dm.rho(dm.forward((xi + d)[:nv], (xi + d)[nv:])) + dm.rho(dm.forward((xi - d)[:nv], (xi - d)[nv:])))
    J0, g = J(xi, True)
    pred = float(g @ d)
    errs = []
    for h in (1.0, 0.5):
        fd = (J(xi + h * d) - J(xi - h * d)) / (2 * h)
        errs.append(abs(fd / pred - 1))
        print(f"h={h}: J0={J0:.6f} fd={fd:.6e} pred={pred:.6e} rel={errs[-1]:.2e}")
    assert min(errs) < 0.02
    # volume gradient along the same direction (nominal geometry)
    V = lambda x: dm.volume(dm.forward(x[:nv], x[nv:], as_built=False))
    gV = np.concatenate(dm.volume_grad(dm.forward(xi[:nv], xi[nv:], as_built=False)))
    fdV = (V(xi + 1e-4 * d) - V(xi - 1e-4 * d)) / 2e-4
    assert abs(fdV / (gV @ d) - 1) < 1e-4
