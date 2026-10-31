"""V3-b (THEORY §11): local shell buckling of a tube, R_m = 10, t = 1.5, L = 60, axial compression, clamped at z = 0
(x = 0 for the voxel model), load P shared over the nodes of the far end face, which is free to translate laterally.
  python v3b_run.py ccx   out.json 2 3 4      # body-fitted ccx *BUCKLE on C3D10, mesh layers through t
  python v3b_run.py voxel out.json 0.8 0.4 0.2 # cshell sparse engine (fe_sparse.buckling), element sizes h_e (mm)
Reported factor: lambda_1 (load factor for the reference P = 1000 N)."""
import json
import sys
import time
from pathlib import Path

import numpy as np

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "continuous" / "research" / "stage0"))

RM, T_W, L, P = 10.0, 1.5, 60.0, 1000.0
E, NU = 1760.0, 0.42


def run_ccx(out, layers_list, nproc, work, load="area"):
    import trimesh
    from cshell import validate as vd
    ro, ri = RM + T_W / 2, RM - T_W / 2
    m = trimesh.creation.annulus(r_min=ri, r_max=ro, height=L, sections=720)
    m.apply_translation([0, 0, L / 2])
    V, T = np.asarray(m.vertices), np.asarray(m.faces)
    zc = V[T].mean(1)[:, 2]
    TF = np.where(zc < 1e-6, 0, np.where(zc > L - 1e-6, 1, 2))
    rows = []
    for lay in layers_list:
        h = T_W / lay
        t0 = time.time()
        bf = vd.validate_design(V, T, V, T, TF, [0], [dict(faces=[1], force=[0, 0, -P])], Path(work) / f"ccx_l{lay}",
                                name=f"tube_l{lay}", h=h, layers=lay, h_min=h, h_bc=1e-3 / 0.75, n_modes=6, nproc=nproc,
                                want_stress=False, load_distribution=load)
        r = dict(case="ccx", load=load, layers=lay, h=h, n_nodes=bf.n_nodes, n_elems=bf.n_elems, dof=3 * bf.n_nodes,
                 lambdas=bf.buckling.tolist(), compliance=bf.compliance, volume=bf.volume, mesh_s=bf.seconds_mesh,
                 ccx_s=bf.seconds_ccx, total_s=time.time() - t0)
        rows.append(r); print(json.dumps(r), flush=True)
        Path(out).write_text(json.dumps(rows, indent=1))


def run_voxel(out, hs):
    from cshell.fe_sparse import ActiveGrid, SparseFE, buckling
    from tubes import section_fraction
    ro, ri = RM + T_W / 2, RM - T_W / 2
    rows = []
    for h in hs:
        t0 = time.time()
        n = 2 * int(np.ceil(ro / h))
        sec = section_fraction(ro, ri, h, -n / 2 * h, n, -n / 2 * h, n, sub=8 if h > 0.3 else 6)
        nx = int(round(L / h))
        vf = np.repeat(sec[None], nx, 0)
        g = ActiveGrid(vf > 0)
        nn = g.node_shape
        fixed = np.zeros(nn + (3,), bool); fixed[0] = True
        load = np.zeros(nn + (3,))
        act = np.zeros(int(np.prod(nn)), bool); act[g.node_dense] = True
        tip = act.reshape(nn)[-1]
        load[-1][tip, 0] = -P / tip.sum()
        fe = SparseFE(g, fixed, load, h, E, NU, backend="auto")
        fe.factor(g.gather_elems(vf))
        pr = fe.primal()
        b = buckling(fe, pr, num_modes=6, tol=1e-6, min_density=0.1)
        # modal strain-energy distribution along the axis (6 slabs of 10 mm; slab 0 = clamped end) of the first modes
        ex = np.unravel_index(g.elem_dense, g.elem_shape)[0] * h
        slab = np.minimum((ex / (L / 6)).astype(int), 5)
        edist = []
        for j in range(min(4, b.mu.size)):
            e = np.asarray(fe.element_bilinear(b.modes[:, j], b.modes[:, j], solid=False).get()
                           if hasattr(fe.element_bilinear(b.modes[:, j], b.modes[:, j], solid=False), "get")
                           else fe.element_bilinear(b.modes[:, j], b.modes[:, j], solid=False))
            edist.append((np.bincount(slab, e, 6) / e.sum()).round(4).tolist())
        r = dict(case="voxel", h=h, energy_by_slab=edist, dof=int(g.ndof), n_elems=int(g.Ne), lambdas=(1.0 / b.mu).tolist(), mu=b.mu.tolist(),
                 compliance=float(pr.compliance), converged=bool(b.converged), iters=int(b.iterations),
                 vol=float((vf * h ** 3).sum()), stats={k: v for k, v in fe.stats.items() if isinstance(v, (int, float, str))},
                 total_s=time.time() - t0)
        fe.free()
        rows.append(r); print(json.dumps(r), flush=True)
        Path(out).write_text(json.dumps(rows, indent=1))


if __name__ == "__main__":
    which, out = sys.argv[1], sys.argv[2]
    if which == "ccx":
        import os
        run_ccx(out, [float(a) for a in sys.argv[3:]] or [2, 3, 4], int(os.environ.get("SLURM_CPUS_PER_TASK", 8)),
                os.environ.get("TMPDIR", "/tmp"), os.environ.get("V3B_LOAD", "area"))
    else:
        run_voxel(out, [float(a) for a in sys.argv[3:]] or [0.8, 0.4, 0.2])
