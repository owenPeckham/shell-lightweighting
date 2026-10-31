"""GPU geometry chain (design.DeviceDesignMap, THEORY §3.6-3.10, §6.1) against the CPU chain on the B2 rod:
same V, ρ, pull-backs (to rounding) and the same 3-iteration optimiser history."""
import csv
from pathlib import Path

import numpy as np
import pytest

ROOT = Path(__file__).resolve().parents[2]
B2 = str(ROOT / "examples/cshell/bench/B2_slender.json")
PHASE = {"grid": {"phases": [{"h_e": 0.8, "h_g": 0.2}]}}


def _cfg(device, extra=None):
    from cshell import config
    o = {**PHASE, "grid": {**PHASE["grid"], "geometry_device": device}, "fe": {"backend": "auto"}}
    for k, v in (extra or {}).items():
        o[k] = {**o.get(k, {}), **v}
    return config.load(B2, o)


@pytest.fixture(scope="module")
def models():
    from cshell.optimise import Context, PhaseModel
    ctx = Context(_cfg("cpu"))
    m_cpu = PhaseModel(ctx, 0, 4.0)
    ctx.cfg = _cfg("gpu")
    m_gpu = PhaseModel(ctx, 0, 4.0)
    return ctx, m_cpu, m_gpu


def _rel(a, b):
    a, b = np.asarray(a, float), np.asarray(b, float)
    return float(np.abs(a - b).max() / max(np.abs(b).max(), 1e-300))


@pytest.mark.gpu
def test_setup_agrees(models):
    _, c, g = models
    assert g.geometry_device == "gpu" and c.geometry_device == "cpu"
    assert abs(g.V0 / c.V0 - 1) < 1e-10
    assert _rel(g.rho0, c.rho0) < 1e-12
    assert np.array_equal(g.stress_mask_dense, c.stress_mask_dense)


@pytest.mark.gpu
@pytest.mark.parametrize("as_built", [False, True])
def test_chain_agrees(models, as_built):
    """V to 1e-10 relative; ρ, ∂(gᵀρ̂)/∂ξ, ∂V/∂ξ and the sliver gradient to 1e-6 (max-norm relative)."""
    ctx, c, g = models
    nv = ctx.surf.n_vertices
    rng = np.random.default_rng(1)
    xd, xa = rng.uniform(0.1, 0.9, nv), rng.uniform(0.1, 0.9, nv)
    sc, sg = c.dm.forward(xd, xa, as_built), g.dm.forward(xd, xa, as_built)
    assert abs(g.dm.volume(sg) / c.dm.volume(sc) - 1) < 1e-10
    rc, rg = c.dm.rho(sc), g.dm.rho(sg)
    assert isinstance(rg, np.ndarray) and _rel(rg, rc) < 1e-12
    gE = rng.standard_normal(c.disc.n_elem)
    for a, b in zip(g.dm.pull_rho_hat(sg, gE), c.dm.pull_rho_hat(sc, gE)):
        assert _rel(a, b) < 1e-6
    for a, b in zip(g.dm.volume_grad(sg), c.dm.volume_grad(sc)):
        assert _rel(a, b) < 1e-6
    assert g.dm.sliver(sg)[0] == pytest.approx(c.dm.sliver(sc)[0], rel=1e-12)
    for a, b in zip(g.dm.sliver_grad(sg), c.dm.sliver_grad(sc)):
        assert _rel(a, b) < 1e-6


@pytest.mark.gpu
def test_offload_roundtrip(models):
    """release() frees the device band, upload() restores it bit-for-bit (volume unchanged)."""
    ctx, _, g = models
    nv = ctx.surf.n_vertices
    st = g.dm.forward(np.full(nv, 0.5), np.full(nv, 0.3), False)
    V1 = g.dm.volume(st)
    g.disc.release()
    assert not g.disc.resident
    with pytest.raises(RuntimeError):
        g.dm.volume(st)
    g.disc.upload()
    assert g.dm.volume(st) == pytest.approx(V1, rel=1e-14)
    # zero-copy fallback: kernels read the pinned host arrays; same results to atomic-order rounding
    sb = g.dm.forward(np.full(nv, 0.4), np.full(nv, 0.6), True)
    r1, gr1 = g.dm.rho(sb), g.dm.volume_grad(st)
    g.disc.use_mapped("test")
    assert g.disc.mode == "mapped" and g.disc.upload() == 0.0 and g.disc.mode == "mapped"
    assert g.dm.volume(st) == pytest.approx(V1, rel=1e-13)
    assert _rel(g.dm.rho(sb), r1) < 1e-12
    for a, b in zip(g.dm.volume_grad(st), gr1):
        assert _rel(a, b) < 1e-10
    g.disc.mapped = False
    g.disc.upload()
    assert g.disc.mode == "device"


@pytest.mark.gpu
@pytest.mark.parametrize("objective", ["compliance", "full"])
def test_history_agrees(tmp_path, objective):
    """3 MMA iterations on B2 with the CPU and the GPU chain: same J history to 1e-6."""
    from cshell.optimise import Context, Run
    extra = {"optimiser": {"max_iter": [3]}}
    if objective == "full":
        extra["objective"] = {"w_stiffness": 0.5, "w_capacity": 0.5}
    hist = {}
    ctx = None
    for dev in ("cpu", "gpu"):
        cfg = _cfg(dev, extra)
        if ctx is None:
            ctx = Context(cfg)
        ctx.cfg = cfg
        out = tmp_path / dev
        Run(cfg, "uniform", str(out)).run(ctx)
        with open(out / "history.csv") as f:
            hist[dev] = list(csv.DictReader(f))
    assert len(hist["cpu"]) == len(hist["gpu"]) == 3
    for rc, rg in zip(hist["cpu"], hist["gpu"]):
        assert abs(float(rg["J"]) - float(rc["J"])) <= 1e-6 * max(1.0, abs(float(rc["J"])))
        assert abs(float(rg["mass_err"]) - float(rc["mass_err"])) < 1e-8
