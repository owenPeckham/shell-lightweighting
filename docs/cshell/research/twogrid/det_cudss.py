"""cuDSS run-to-run reproducibility of the coarse factor/solve (P6): block width (nrhs 7 or 1), DETERMINISTIC_MODE on/off,
fp32/fp64. For each variant the same K_c is factorised in two fresh solvers and the same right-hand sides are solved
several times; prints the largest difference to the first solution (0 = bitwise) and the time per right-hand side.
Usage: python det_cudss.py [--shape 81,33,41] [--reps 3]"""
import argparse
import sys
import time
from pathlib import Path

import numpy as np

ROOT = Path(__file__).resolve().parents[4]
sys.path.insert(0, str(ROOT)); sys.path.insert(0, str(ROOT / "tests" / "cshell"))
import cupy as cp                                                    # noqa: E402
from cshell.fe_twogrid import LeanGrid, TwoGridFE, _CoarseSolver     # noqa: E402
from test_fe_twogrid import E, NU, hollow_tube                       # noqa: E402

ap = argparse.ArgumentParser()
ap.add_argument("--shape", default="81,33,41"); ap.add_argument("--reps", type=int, default=3)
a = ap.parse_args()
mask, rho, fixed, load = hollow_tube(tuple(int(x) for x in a.shape.split(",")))
gt = LeanGrid(mask, xp=cp)
t = TwoGridFE(gt, fixed, load, 1.0, E, NU, backend="gpu").factor(gt.gather_elems(rho))
data = t._coarse_values()
pat = t._pat
R = cp.asarray(np.random.default_rng(4).standard_normal((pat.ndof, 7)))
print("coarse ndof", pat.ndof, flush=True)
for prec in ("fp32", "fp64"):
    for nrhs, mode in ((7, False), (1, False), (1, True), (7, True)):
        sols, tt = [], []
        try:
            for rep in range(a.reps):
                co = _CoarseSolver(cp, pat, prec, nrhs, deterministic=mode)
                co.factor(data)
                for _ in range(3):
                    cp.cuda.Device().synchronize(); t0 = time.time()
                    x = co.solve(R)
                    cp.cuda.Device().synchronize(); tt.append(time.time() - t0)
                    sols.append(cp.asnumpy(x))
                co.free()
            d = max(float(np.abs(s - sols[0]).max()) for s in sols[1:])
            print(f"{prec} nrhs {nrhs} mode {mode}: max|diff| {d:.3e}  {1e3 * np.median(tt) / 7:.2f} ms/rhs", flush=True)
        except Exception as e:
            print(f"{prec} nrhs {nrhs} mode {mode}: {type(e).__name__} {e}", flush=True)
