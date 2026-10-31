"""Wall and offset profile along x for a rod benchmark: python wall_profile.py <cfg> <xi.npy> <phase> [D_comp]"""
import json, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[4]))
import numpy as np
from cshell import config
from cshell.optimise import Context, PhaseModel
cfg = config.load(sys.argv[1]); ctx = Context(cfg); ph = int(sys.argv[3])
D = float(sys.argv[4]) if len(sys.argv) > 4 else cfg["design"]["outward_extent_initial_mm"]
m = PhaseModel(ctx, ph, D)
nv = ctx.surf.n_vertices
xi = np.load(sys.argv[2])
_, xt_a, d, a, _, _ = m.dm.vertex_fields(xi[:nv], xi[nv:])
V = ctx.surf.V
side = np.abs(ctx.surf.N[:, 0]) < 0.3                     # curved surface only
act = side & ~m.ds.passive
L = V[:, 0].max()
edges = np.linspace(0, L, 9)
rows = []
for lo, hi in zip(edges[:-1], edges[1:]):
    s = act & (V[:, 0] >= lo) & (V[:, 0] < hi)
    if s.any():
        rows.append(dict(x=round(0.5 * (lo + hi), 1), n=int(s.sum()), d=float(np.median(d[s])), a=float(np.median(a[s])),
                         wall=float(np.median((a + d)[s])), outer_r=float(5 + np.median(d[s])),
                         inner_r=float(5 - np.median(a[s])), closed=float(np.mean(a[s] >= ctx.r_ma[s]))))
print(json.dumps(rows, indent=0))
