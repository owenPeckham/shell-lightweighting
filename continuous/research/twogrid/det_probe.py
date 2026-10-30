"""Run-to-run bitwise reproducibility of every stage of the two-grid engine on the GPU (P6 of PLAN_collect_fixes.md).

Each stage is run `--reps` times on identical inputs (new engine instances for set-up stages) and the largest absolute
difference to the first run is printed (0 = bitwise equal). Run with both `deterministic` settings: the atomic kernels
(False) are expected to differ, the colour-ordered ones (True) not. cuSPARSE SpMM (prolongation), cuDSS (coarse factor
and solve), cuBLAS/cuSOLVER (Lanczos, eigen-solver Gram/eigh) are covered by the stages that use them.
Usage: python det_probe.py [--shape 81,33,41] [--reps 4] [--modes] [--det-only]
"""
import argparse
import sys
from pathlib import Path

import numpy as np

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT)); sys.path.insert(0, str(ROOT / "tests" / "cshell"))
import cupy as cp                                                    # noqa: E402
from cshell.fe_sparse import buckling                                # noqa: E402
from cshell.fe_twogrid import LeanGrid, TwoGridFE                    # noqa: E402
from test_fe_twogrid import E, NU, hollow_tube                       # noqa: E402


def dmax(runs):
    return max(float(np.abs(np.asarray(r) - np.asarray(runs[0])).max()) for r in runs[1:])


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--shape", default="81,33,41"); ap.add_argument("--reps", type=int, default=4)
    ap.add_argument("--modes", action="store_true", help="also the full buckling evaluation (slow)")
    ap.add_argument("--det-only", action="store_true", help="skip the atomic-kernel reference run")
    ap.add_argument("--mode", default="full", choices=["full", "kernels"], help="deterministic mode of the second run")
    a = ap.parse_args()
    shape = tuple(int(x) for x in a.shape.split(","))
    mask, rho, fixed, load = hollow_tube(shape)
    for det in ((a.mode,) if a.det_only else ("off", a.mode)):
        print(f"=== deterministic={det}  shape {shape}", flush=True)
        out = {}
        rng = np.random.default_rng(1)
        gt = None
        engines = []
        for r in range(a.reps):
            gt = LeanGrid(mask, xp=cp)
            t = TwoGridFE(gt, fixed, load, 1.0, E, NU, backend="gpu", deterministic=det).factor(gt.gather_elems(rho))
            engines.append(t)
            for name, v in (("diag", t._diag), ("lmax_DinvK", np.float64(t.stats["lmax_DinvK"])),
                            ("coarse_data", t._coarse_values())):
                out.setdefault(name, []).append(cp.asnumpy(v) if hasattr(v, "get") else v)
            X = cp.asarray(np.random.default_rng(2).standard_normal((t.g.ndof, 7)))
            out.setdefault("K_apply", []).append(cp.asnumpy(t._apply_K(X)))
            Hm = cp.asarray(np.random.default_rng(3).standard_normal((t.g.Ne, 8, 8)))
            out.setdefault("G_apply", []).append(cp.asnumpy(t._geometric_op(Hm) @ X))
            out.setdefault("spmm_PT", []).append(cp.asnumpy(t._PnT @ X.reshape(-1, 21)))
            Xc = cp.asarray(np.random.default_rng(4).standard_normal((t._pat.ndof, 7)))
            out.setdefault("spmm_P", []).append(cp.asnumpy(t._Pn @ Xc.reshape(-1, 21)))
            out.setdefault("coarse_solve", []).append(cp.asnumpy(t._co.solve(Xc)))
            out.setdefault("precondition", []).append(cp.asnumpy(t.precondition(X)))
            p = t.primal()
            out.setdefault("primal_u", []).append(cp.asnumpy(p.u))
            out.setdefault("compliance", []).append(np.float64(p.compliance))
            if a.modes:
                b = buckling(t, p, num_modes=6, tol=1e-6, maxiter=100)
                out.setdefault("mu", []).append(np.asarray(b.mu))
                out.setdefault("dmu", []).append(np.asarray(b.dmu_tilde))
                out.setdefault("modes", []).append(cp.asnumpy(b.modes))
        # repeated cuSPARSE / cuDSS calls on one factor
        t = engines[-1]
        X = cp.asarray(np.random.default_rng(2).standard_normal((t.g.ndof, 7)))
        Xc = cp.asarray(np.random.default_rng(4).standard_normal((t._pat.ndof, 7)))
        out["spmm_PT_same"] = [cp.asnumpy(t._PnT @ X.reshape(-1, 21)) for _ in range(a.reps * 4)]
        out["coarse_solve_same_factor"] = [cp.asnumpy(t._co.solve(Xc)) for _ in range(a.reps * 4)]
        for k, v in out.items():
            print(f"  {k:28s} max|diff| {dmax(v):.3e}", flush=True)
        for t in engines:
            t.free()


if __name__ == "__main__":
    main()
