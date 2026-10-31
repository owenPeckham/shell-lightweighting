"""Full body-fitted chain of collect (levelset surface + thickness-aware remesh + gmsh + peel + gate + GPU C3D10 solve)
for one model and one level: time, tets, ndof, factor size, peak GPU memory.
usage: python solve_check.py <run_dir> solid|design <h>"""
import json
import os
import sys
import time

sys.path.insert(0, os.getcwd())
import numpy as np


def main():
    run_dir, which, h = sys.argv[1], sys.argv[2], float(sys.argv[3])
    import cupy as cp
    from cshell import collect as C, config as ccfg, validate as cval
    from cshell.optimise import Context
    cfg = ccfg.load(json.load(open(os.path.join(run_dir, "config_used.json"))))
    ctx = Context(cfg, log=print)
    if which == "solid":
        V, T = ctx.V, ctx.T
    else:
        V, T, _ = C.load_mesh(os.path.join(run_dir, "best_candidate.stl"))
    t0 = time.time()
    s = 0.5 / 3
    Vs, Ts, ls = cval.levelset_surface(V, T, s, log=print)
    dev = cval.surface_deviation(Vs, Ts, V, T)
    print("DEV " + json.dumps(dev), flush=True)
    t1 = time.time()
    wd = os.path.join(os.environ.get("TMPDIR", "/tmp"), "bfs")
    bf = cval.validate_design(Vs, Ts, ctx.V, ctx.T, ctx.TF, ctx.support, ctx.loads, wd, f"{which}_h{h:.3f}", h=h,
                              h_bc=float(cfg["grid"]["phases"][-1]["h_e"]), E=cfg["material"]["E_MPa"],
                              nu=cfg["material"]["nu"], enforce_gate=False, remesh_kw=dict(thin=2.0, fine=s), peel=True)
    q = dict(bf.info.get("mesh_quality", {}))
    q.pop("layers", None)
    out = dict(part=os.path.basename(run_dir), which=which, h=h, C=bf.compliance, lam=bf.buckling.tolist(),
               lam_unexcl=bf.info.get("lam_unexcl"), sigma_q=bf.sigma_q, volume=bf.volume, n_nodes=bf.n_nodes,
               n_elems=bf.n_elems, seconds_mesh=bf.seconds_mesh, seconds_solve=bf.seconds_ccx,
               ndof=bf.info.get("ndof"), factor_GiB=bf.info.get("factor_GiB"), hybrid=bf.info.get("hybrid"),
               timings=bf.info.get("timings_s"), gate=q, t_surface=t1 - t0, levelset=ls,
               peak_gpu_GiB=cp.get_default_memory_pool().total_bytes() / 2 ** 30)
    print("RESULT " + json.dumps(out, default=float), flush=True)


if __name__ == "__main__":
    main()
