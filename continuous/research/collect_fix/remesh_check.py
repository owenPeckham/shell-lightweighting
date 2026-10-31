"""Does the B2 body-fitted surface-remesh failure (CF_B2, job 7868889: edge_max 1.55 > 2.5 h at h = 0.5) come from the
STL loader (exact weld vs trimesh process=True) or from the re-exported geometry (P2 closure rule)? Each case twice
(remesh determinism). usage: python remesh_check.py   (CPU job, from the repo root)"""
import os
import sys

sys.path.insert(0, os.getcwd())
import numpy as np


def main():
    import trimesh
    from cshell import collect as C, validate as cval
    R = "runs_cshell"
    for tag, p in (("new STL, exact weld", f"{R}/CF_B2/best_candidate.stl"),
                   ("new STL, trimesh process", f"{R}/CF_B2/best_candidate.stl"),
                   ("old STL, exact weld", f"{R}/VF_B2/best_candidate.stl"),
                   ("old STL, trimesh process", f"{R}/VF_B2/best_candidate.stl")):
        if "trimesh" in tag:
            tm = trimesh.load(p, process=True)
            V, T = np.asarray(tm.vertices, float), np.asarray(tm.faces, np.int64)
        else:
            V, T, _ = C.load_mesh(p)
        for rep in range(2):
            try:
                out = cval.remesh_surface(V, T, 0.5)
                info = out[2] if isinstance(out, tuple) and len(out) > 2 else out
                print(f"RESULT {tag:<26} rep {rep} V {len(V)} T {len(T)}: ok {str(info)[:260]}", flush=True)
            except Exception as e:
                print(f"RESULT {tag:<26} rep {rep} V {len(V)} T {len(T)}: FAILED {str(e)[:260]}", flush=True)


if __name__ == "__main__":
    main()
