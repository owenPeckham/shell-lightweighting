"""Brute-force uniform-shell sweep with the same FE (V5-B2/B3 baseline): uniform ξ̃ᵃ = s, ξᵈ restored to constant
mass, scored by J, λ_b, C. python uniform_sweep.py <cfg> <phase> <out.json> [D_comp]"""
import json, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[3]))
import numpy as np
from cshell import config
from cshell.mma import restore_mass
from cshell.optimise import Context, PhaseModel
cfg = config.load(sys.argv[1]); ph = int(sys.argv[2]); out = sys.argv[3]
D = float(sys.argv[4]) if len(sys.argv) > 4 else cfg["design"]["outward_extent_initial_mm"]
ctx = Context(cfg); m = PhaseModel(ctx, ph, D); nv = ctx.surf.n_vertices; dm = m.dm
st = m.settings(1.0); base = m.baseline(st)
act = np.concatenate([~m.ds.passive, np.zeros(nv, bool)])          # restore with ξᵈ only: ξᵃ stays uniform
vol = lambda x: dm.volume(dm.forward(x[:nv], x[nv:], as_built=False))
vg = lambda x: np.concatenate(dm.volume_grad(dm.forward(x[:nv], x[nv:], as_built=False)))
rows = []
for s_ in np.linspace(0, 1, 11):
    xi = np.concatenate([np.full(nv, 0.3), np.full(nv, s_)])
    xi[:nv][m.ds.passive] = 0; xi[nv:][m.ds.passive] = 1
    try:
        xi, _, _ = restore_mass(xi, vol, vg, m.V0, act, tol=1e-4)
    except RuntimeError as e:
        rows.append(dict(s=float(s_), error=str(e))); continue
    rho = dm.rho(dm.forward(xi[:nv], xi[nv:], as_built=True))
    lam_s0 = cfg["material"]["strength_MPa"] / (base["c"] * base["sigma_pn0"])
    fe = m.fe_eval(rho, st, base, base["c"])
    r = dict(s=float(s_), J=fe["J"], C_C0=fe["C"] / base["C0"],
             lam_b_ratio=fe["lam_b"] / base["lam_b0"] if np.isfinite(base["lam_b0"]) else None,
             lam_ratio=fe["lam"] / base["lam0"], governing=fe["governing"])
    rows.append(r); print(json.dumps(r), flush=True)
Path(out).write_text(json.dumps(dict(base={k: v for k, v in base.items() if isinstance(v, (int, float, str))}, rows=rows), indent=1))
