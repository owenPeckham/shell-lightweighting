import sys, numpy as np, trimesh, igl
from OCP.BRepOffsetAPI import BRepOffsetAPI_MakeOffsetShape
from occ_util import DATA_DIR, OUT_DIR, load_step, volume, tessellate
name, t = sys.argv[1], float(sys.argv[2])
s = load_step(f'{DATA_DIR}/{name}.step')
m = BRepOffsetAPI_MakeOffsetShape(); m.PerformBySimple(s, t); o = m.Shape()
V0, T0 = tessellate(s, lin=0.05); V1, T1 = tessellate(o, lin=0.05)
m0 = trimesh.Trimesh(V0, T0); m0.merge_vertices(); m1 = trimesh.Trimesh(V1, T1); m1.merge_vertices()
print('orig mesh vol', round(m0.volume), 'offset mesh vol', round(m1.volume), 'offset watertight', m1.is_watertight, 'OCC vol', round(volume(o)))
P, _ = trimesh.sample.sample_surface(m1, 20000, seed=0)
S, I, C = igl.signed_distance(P, m0.vertices, m0.faces)[:3]
print('signed dist of offset surface to original: pct 1/5/50/95/99 =', np.round(np.percentile(S, [1, 5, 50, 95, 99]), 3), 'frac within +-10% of t:', round(np.mean(abs(S - t) < 0.1 * t), 3))
