"""Isosurface-stuffing body-fitted mesh of one model at the two collect levels: tets, gate, deviation from the input
surface; with 'solve' also the GPU C3D10 solve (validate_design). usage:
python stuff_check.py <run_dir> solid|design [solve] [c=sqrt2]"""
import json
import os
import sys
import time

sys.path.insert(0, os.getcwd())
import numpy as np


def main():
    run_dir, which = sys.argv[1], sys.argv[2]
    solve = "solve" in sys.argv[3:]
    c = float(os.environ.get("STUFF_C", 2 ** 0.5))
    from cshell import collect as C, config as ccfg, validate as cval
    from cshell.optimise import Context
    cfg = ccfg.load(json.load(open(os.path.join(run_dir, "config_used.json"))))
    ctx = Context(cfg, log=print)
    if which == "solid":
        V, T = ctx.V, ctx.T
    else:
        V, T, _ = C.load_mesh(os.path.join(run_dir, "best_candidate.stl"))
    wd = os.path.join(os.environ.get("TMPDIR", "/tmp"), "bfs")
    for h in (0.5, 0.5 / 2 ** 0.5):
        out = dict(part=os.path.basename(run_dir), which=which, h=h, c=c)
        try:
            t0 = time.time()
            if solve:
                bf = cval.validate_design(V, T, ctx.V, ctx.T, ctx.TF, ctx.support, ctx.loads, wd, f"{which}_h{h:.3f}",
                                          h=h, h_bc=float(cfg["grid"]["phases"][-1]["h_e"]),
                                          E=cfg["material"]["E_MPa"], nu=cfg["material"]["nu"], enforce_gate=False,
                                          backend="stuff", remesh_kw=dict(stuff_c=c))
                q = dict(bf.info.get("mesh_quality", {}))
                out.update(C=bf.compliance, lam=bf.buckling.tolist(), lam_unexcl=bf.info.get("lam_unexcl"),
                           sigma_q=bf.sigma_q, volume=bf.volume, n_nodes=bf.n_nodes, n_elems=bf.n_elems,
                           seconds_mesh=bf.seconds_mesh, seconds_solve=bf.seconds_ccx, ndof=bf.info.get("ndof"),
                           factor_GiB=bf.info.get("factor_GiB"), hybrid=bf.info.get("hybrid"),
                           timings=bf.info.get("timings_s"))
            else:
                info = {}
                n4, t4 = cval.tetrahedralise(V, T, h, backend="stuff", remesh_kw=dict(stuff_c=c), info=info)
                n10, t10 = cval.to_quadratic(n4, t4)
                Tb = cval.boundary_faces(t4)
                q = cval.mesh_gate(n10, t10, n4, Tb, h=h, enforce=False)
                q.update(info, deviation=cval.surface_deviation(n4, Tb, V, T))
                out.update(n_elems=int(len(t4)), n_nodes10=int(len(n10)))
            q.pop("layers", None)
            out.update(gate=q, seconds=time.time() - t0)
        except Exception as e:
            import traceback
            traceback.print_exc()
            out["error"] = repr(e)[:800]
        print("RESULT " + json.dumps(out, default=float), flush=True)


if __name__ == "__main__":
    main()
