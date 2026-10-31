import sys, time, numpy as np
sys.path.insert(0, '.')
from vshell import geometry as g
from cshell import surface as S
V, T, TF = g.tessellate(g.load_step('examples/parts/Solid Cantilever.step'), 0.02, 0.2)
for kw in [dict(), dict(tol=0.03)]:
    t = time.time()
    V1, T1, info = S.remesh_isotropic(V, T, ell=0.75, max_growth=2.2, **kw)
    h = info['hausdorff']
    print(kw, 'nv', len(V1), 'hd %.4f p999 %.4f frac %.5f' % (h['hausdorff'], h['p999'], h.get('frac_above_tol', -1)), 'hist', [(a, round(b, 3)) for a, b in info['history']], 'flips', info['flips'], 'time %.0f' % (time.time() - t), flush=True)
