"""Mesh-gate check on B2 solid + design at both levels, no solve. usage: python meshgate.py RUN_DIR"""
import sys, time, numpy as np, trimesh
sys.path.insert(0, "/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2")
from cshell import validate as v
rd = sys.argv[1]
for nm, f in (("solid", "original.stl"), ("design", "best_candidate.stl")):
    m = trimesh.load(f"{rd}/{f}", process=False)
    for h in (0.5, 0.5 / 2 ** 0.5):
        t0 = time.time()
        try:
            V1, T1, info = v.remesh_surface(np.asarray(m.vertices, float), np.asarray(m.faces), h)
            t1 = time.time()
            n, t = v._gmsh_inside(V1, T1, h)
            q = v.mesh_gate(n, v.to_quadratic(n, t)[1], V1, T1, h=h, enforce=False)
            print(nm, h, "surf", info, f"{t1-t0:.0f}s", "vol", f"{time.time()-t1:.0f}s", "tets", len(t), "gate",
                  {k: q[k] for k in ("dihedral_min", "frac_dihedral_lt_5", "n_dihedral_lt_5", "edge_max_over_h", "layers_p01", "passed", "failures")}, flush=True)
        except Exception as e:
            print(nm, h, "FAIL", repr(e)[:400], flush=True)
