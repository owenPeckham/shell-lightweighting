"""GPU (cuDSS) vs CalculiX on the SAME C3D10 deck of the B2 design (new chain: remesh_surface + gmsh inside + gate).
usage: python xcheck.py <stl> <h> <ccx_threads>   (run from the repo root)"""
import json, sys, time
sys.path.insert(0, "/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2")
from pathlib import Path
import numpy as np
import trimesh
from cshell import validate as vd, bodyfit_gpu as bg

stl, h, nt = sys.argv[1], float(sys.argv[2]), int(sys.argv[3])
out = Path("runs_cshell/VF_B2/xcheck"); out.mkdir(parents=True, exist_ok=True)
bc = np.load("runs_cshell/VF_B2/bc_faces.npz")
cV, cT = bc["V"].astype(float), bc["T"].astype(np.int64)
TF = np.arange(len(cT))
support, loads = TF[bc["support"]].tolist(), [dict(faces=TF[bc["load"] == 0].tolist(), force=[0.0, 0.0, -20.0])]
m = trimesh.load(stl, process=False); m.merge_vertices()
V, T = np.asarray(m.vertices, float), np.asarray(m.faces, np.int64)
t0 = time.time()
V1, T1, sinfo = vd.remesh_surface(V, T, h)
n4, t4 = vd._gmsh_inside(V1, T1, h, out)
nodes, t10 = vd.to_quadratic(n4, t4)
q = vd.mesh_gate(nodes, t10, V, T, h=h, enforce=False)
case = vd.build_case(nodes, t10, cV, cT, TF, support, loads, 0.75 * 0.4)
tm = time.time() - t0
t0 = time.time()
g = bg.solve(case.nodes, case.tets10, case.fixed, case.load_nodes, case.load_vec, vd.E_MPA, vd.NU, 4, exclude_bc_mm=2.0)
tg = time.time() - t0
g.pop("_phi_nodes")
inp = vd.write_inp(out / "design.inp", case, n_modes=4, want_stress=True)
t0 = time.time()
run = vd.run_ccx(inp, nt)
tc = time.time() - t0
res = vd.parse_dat(run.dat)
C, sq, vol = vd.responses_from_ccx(case, res)
R = dict(h=h, stl=stl, surface=sinfo, gate=q, nodes=len(nodes), tets=len(t10), seconds=dict(mesh=tm, gpu=tg, ccx=tc),
         C=dict(gpu=g["C"], ccx=C, rel=g["C"] / C - 1), sigma_q=dict(gpu=g["sigma_q"], ccx=sq, rel=g["sigma_q"] / sq - 1),
         lam_unexcl=dict(gpu=g["lam_unexcl"], ccx=res.buckling.tolist(),
                         rel=(np.asarray(g["lam_unexcl"])[:len(res.buckling)] / res.buckling - 1).tolist()),
         lam_excl_gpu=g["lam"], n_ks_excluded=g["n_ks_excluded"], factor_GiB=g["factor_GiB"])
json.dump(R, open(out / f"xcheck_h{h:g}.json", "w"), indent=1, default=float)
print(json.dumps(R, indent=1, default=float))
