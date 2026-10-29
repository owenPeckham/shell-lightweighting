import sys, json, time
import numpy as np, trimesh
from cshell import validate as vd
def shell():
    o = trimesh.creation.icosphere(subdivisions=5, radius=8.0); i = trimesh.creation.icosphere(subdivisions=5, radius=7.0); i.invert()
    m = trimesh.util.concatenate([o, i]); return np.asarray(m.vertices), np.asarray(m.faces)
def tube():
    m = trimesh.creation.annulus(r_min=4.0, r_max=5.0, height=20.0, sections=96); return np.asarray(m.vertices), np.asarray(m.faces)
for nm, geo in (("shell", shell), ("tube", tube)):
    V, T = geo()
    for h in (0.2, 0.3, 0.5):
        t0 = time.time()
        try:
            V1, T1, info = vd.remesh_surface(V, T, h)
        except Exception as e:
            print(nm, h, "REMESH FAIL", str(e)[:300], flush=True); continue
        nodes, tets = vd._gmsh_inside(V1, T1, h)
        n10, t10 = vd.to_quadratic(nodes, tets)
        q = vd.mesh_gate(n10, t10, V, T, h=h, enforce=False)
        q.pop("layers")
        print(nm, h, info, "tets", len(tets), "s", round(time.time() - t0), json.dumps(q, default=float), flush=True)
