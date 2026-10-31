"""Surface remesh + volume variants. usage: meshgate2.py RUN_DIR solid|design h mode  (mode: surf | vol:<algo>:<rounds>,...)"""
import sys, time, numpy as np, trimesh
sys.path.insert(0, "/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2")
from cshell import validate as v
from cshell.bodyfit_gpu import dihedral_min
def main():
    rd, nm, hs, mode = sys.argv[1], sys.argv[2], [float(x) for x in sys.argv[3].split(",")], sys.argv[4]
    if nm == "ctx":
        bc = np.load(f"{rd}/bc_faces.npz"); m = trimesh.Trimesh(bc["V"].astype(float), bc["T"].astype(np.int64), process=False)
    else:
        m = trimesh.load(f"{rd}/{'original.stl' if nm == 'solid' else 'best_candidate.stl'}", process=False)
    for h in hs:
        t0 = time.time()
        try:
            V1, T1, info = v.remesh_surface(np.asarray(m.vertices, float), np.asarray(m.faces), h)
        except Exception as e:
            print("SURF FAIL", repr(e)[:500], f"{time.time()-t0:.0f}s", flush=True); continue
        print("surf", info, f"{time.time()-t0:.0f}s", flush=True)
        for var in mode.split(",")[1:] if mode.startswith("vol") else []:
            a = var.split(":"); algo, rounds = int(a[0]), int(a[1]); thr = float(a[2]) if len(a) > 2 else 0.3
            t1 = time.time()
            try:
                n, t = v._gmsh_inside(V1, T1, h, opt_rounds=rounds, algo=algo, opt_thr=thr)
            except Exception as e:
                print("FAIL", var, repr(e)[:100], flush=True)
                continue
            dm = dihedral_min(n, t)
            d = np.sqrt(__import__("igl").point_mesh_squared_distance(n, V1, T1)[0])
            bnd = d < 1e-9
            bad = dm < 5
            print("vol", var, f"{time.time()-t1:.0f}s", "tets", len(t), "min dih", dm.min(), "n<5", int(bad.sum()),
                  "nbnd verts of bad tets", np.bincount(bnd[t[bad]].sum(1), minlength=5).tolist(), flush=True)


if __name__ == "__main__":
    main()
