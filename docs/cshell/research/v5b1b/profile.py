"""V5-B1b regional design profile: python profile.py <cfg> <xi.npy> <phase> <D_comp> [x_fillet_lo x_fillet_hi]
Splits the active surface into slender (x < lo), fillet, and stub (x > hi) and reports area-weighted closed fraction
(a >= r_ma), median wall w = a + d, median d, median a."""
import json, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[4]))
import numpy as np
from cshell import config
from cshell.optimise import Context, PhaseModel
cfg = config.load(sys.argv[1]); ctx = Context(cfg)
m = PhaseModel(ctx, int(sys.argv[3]), float(sys.argv[4]))
lo, hi = (float(sys.argv[5]), float(sys.argv[6])) if len(sys.argv) > 6 else (76.0, 84.0)
nv = ctx.surf.n_vertices
xi = np.load(sys.argv[2])
_, xt_a, d, a, _, _ = m.dm.vertex_fields(xi[:nv], xi[nv:])
d, a, xt_a = [np.asarray(v.get() if hasattr(v, "get") else v) for v in (d, a, xt_a)]
V, A = ctx.surf.V, ctx.surf.A
act = ~np.asarray(m.ds.passive)
res = {}
for name, sel in (("slender", V[:, 0] < lo), ("fillet", (V[:, 0] >= lo) & (V[:, 0] <= hi)), ("stub", V[:, 0] > hi),
                  ("all", V[:, 0] > -1e9)):
    s = sel & act
    closed = a[s] >= ctx.r_ma[s]
    res[name] = dict(n=int(s.sum()), closed_area_frac=float((A[s] * closed).sum() / A[s].sum()),
                     snap_closed_098=float((A[s] * (xt_a[s] >= 0.98)).sum() / A[s].sum()),
                     median_d=float(np.median(d[s])), median_a=float(np.median(a[s])),
                     median_wall=float(np.median((a + d)[s])), d_max=float(d[s].max()),
                     a_over_rma_median=float(np.median(a[s] / ctx.r_ma[s])))
print(json.dumps(res, indent=1))
