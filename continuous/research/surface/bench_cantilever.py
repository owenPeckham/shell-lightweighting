"""Task A timing: Solid Cantilever end to end (tessellate -> Gamma_h -> operators -> filter -> depth -> clearance)."""
import sys, time, json, numpy as np
sys.path.insert(0, '.')
from vshell import geometry as g
from cshell import surface as S
stage = {}
t = time.time()
shape = g.load_step('examples/step/Solid Cantilever.step'); V, T, TF = g.tessellate(shape, 0.02, 0.2)
stage['step+tessellate'] = time.time() - t; t = time.time()
ds = S.DesignSurface.build(V, T, TF, ell=0.75, hausdorff_tol=0.02, tess_tol=0.02, verbose=True)
stage['design_surface(remesh+operators+face_id)'] = time.time() - t
print('nv', ds.n_vertices, 'nt', len(ds.T), 'hd', {k: v for k, v in ds.info['hausdorff'].items() if k != 'face_dev'}, 'ok', ds.info['hausdorff_ok'])
print('edge', ds.info['edge_length'], 'neg cot edges', ds.info['n_negative_cot_edges'], 'flips', ds.info['flips'], 'uniform nv', ds.info['n_uniform'])
t = time.time(); f = ds.filter(3.0); stage['filter_factorise'] = time.time() - t
x = np.random.default_rng(0).random(ds.n_vertices); y = f.apply(x)
print('filter range', y.min(), y.max(), 'x range', x.min(), x.max())
t = time.time(); N = ds.projected_normals(); stage['projected_normals'] = time.time() - t
t = time.time(); r, info = ds.depth(0.1, return_info=True, check_march=200); stage['medial_depth'] = time.time() - t
print('r_MA pct', np.percentile(r, [0, 5, 50, 95, 100]), {k: v for k, v in info.items() if k not in ('chord_half',)})
t = time.time(); ro = ds.clearance(); stage['outward_clearance'] = time.time() - t
print('r_out finite frac', np.isfinite(ro).mean(), 'pct', np.percentile(ro[np.isfinite(ro)], [0, 5, 50, 95]))
stage['total'] = sum(stage.values())
print(json.dumps(stage, indent=1))
np.savez('continuous/research/surface/cant_surface.npz', V=ds.V, T=ds.T, face_id=ds.face_id, r_ma=r, r_out=ro)
