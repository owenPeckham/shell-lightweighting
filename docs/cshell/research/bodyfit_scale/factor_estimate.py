"""cuDSS memory estimate (symbolic analysis only, no factorisation) of the body-fitted C3D10 stiffness of a production
model at the two collect levels (stuffing mesh). usage: python factor_estimate.py <run_dir> solid|design [c]"""
import json
import os
import sys
import time

sys.path.insert(0, os.getcwd())
import numpy as np


def main():
    run_dir, which = sys.argv[1], sys.argv[2]
    c = float(sys.argv[3]) if len(sys.argv) > 3 else 2 ** 0.5
    import cupy as cp
    from nvmath.sparse.advanced import DirectSolver, DirectSolverMatrixType, DirectSolverOptions, ExecutionCUDA
    from cshell import collect as C, config as ccfg, validate as cval
    from cshell.bodyfit_gpu import P2Model
    from cshell.optimise import Context
    cfg = ccfg.load(json.load(open(os.path.join(run_dir, "config_used.json"))))
    ctx = Context(cfg, log=print)
    if which == "solid":
        V, T = ctx.V, ctx.T
    else:
        V, T, _ = C.load_mesh(os.path.join(run_dir, "best_candidate.stl"))
    for h in (0.5, 0.5 / 2 ** 0.5):
        out = dict(part=os.path.basename(run_dir), which=which, h=h, c=c)
        try:
            t0 = time.time()
            n4, t4, si = cval.stuff_tets(V, T, c * h)
            n10, t10 = cval.to_quadratic(n4, t4)
            case = cval.build_case(n10, t10, ctx.V, ctx.T, ctx.TF, ctx.support, ctx.loads,
                                   0.75 * float(cfg["grid"]["phases"][-1]["h_e"]))
            M = P2Model(case.nodes, case.tets10, case.fixed, 1760.0, 0.42)
            K = M.stiffness()
            b = cp.zeros((K.shape[0], 1))
            s = DirectSolver(K, b, options=DirectSolverOptions(sparse_system_type=DirectSolverMatrixType.SPD),
                             execution=ExecutionCUDA(hybrid_memory_mode_options=dict(hybrid_memory_mode=True)))
            s.plan()
            est = s.plan_info.memory_estimates
            out.update(n_tets=int(len(t4)), ndof=int(M.nf), nnz=int(K.nnz), seconds=time.time() - t0,
                       **{k: float(est[k]) / 2 ** 30 for k in est.dtype.names})
            s.free()
            del K, M, s
            cp.get_default_memory_pool().free_all_blocks()
        except Exception as e:
            out["error"] = repr(e)[:600]
        print("RESULT " + json.dumps(out, default=float), flush=True)


if __name__ == "__main__":
    main()
