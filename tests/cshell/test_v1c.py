"""V1-c compliance half (THEORY §11): the all-solid design through the design fields (ξ̃ᵃ ≡ 1, ξ̃ᵈ ≡ 0 -> DesignMap ->
ρ) against the solid baseline ψ = φ (solid_rho), same FE pipeline, B1 stub rod at h_e 0.8 mm. C within 0.5 %."""
import dataclasses
from pathlib import Path

import numpy as np
import pytest

ROOT = Path(__file__).resolve().parents[2]


def _np(a):
    return np.asarray(a.get() if hasattr(a, "get") else a)


@pytest.mark.parametrize("backend", [pytest.param("gpu", marks=pytest.mark.gpu), "cpu_direct"])
def test_all_solid_compliance(backend):
    from cshell import config
    from cshell.optimise import Context, PhaseModel
    over = {"grid": {"phases": [{"h_e": 0.8, "h_g": 0.2}]},
            "fe": {"backend": "auto" if backend == "gpu" else backend,
                   "engine": "twogrid" if backend == "gpu" else "direct"}}
    ctx = Context(config.load(str(ROOT / "examples/cshell/bench/B1_stub.json"), over))
    m = PhaseModel(ctx, 0, 4.0)
    nv = ctx.surf.n_vertices
    st = dataclasses.replace(m.settings(0.5), w_P=0.0, buckling=False)
    rho = {"baseline_nominal": _np(m.dm.solid_rho(0.0)), "baseline_as_built": _np(m.rho0)}
    for name, ab in (("fields_nominal", False), ("fields_as_built", True)):
        rho[name] = _np(m.dm.rho(m.dm.forward(np.zeros(nv), np.ones(nv), as_built=ab)))
    C = {}
    for name, r in rho.items():
        C[name] = m.fe_eval(r, st, None, 1.0)["C"]
        print(f"[V1-c {backend}] C {name} = {C[name]:.8e}")
    rel = C["fields_nominal"] / C["baseline_nominal"] - 1
    rel_ab = C["fields_as_built"] / C["baseline_as_built"] - 1
    print(f"[V1-c {backend}] RESULT dC/C nominal {rel:.3e}, as-built {rel_ab:.3e}; max|drho| nominal "
          f"{np.abs(rho['fields_nominal'] - rho['baseline_nominal']).max():.2e}")
    assert abs(rel) < 5e-3
    assert abs(rel_ab) < 5e-3
