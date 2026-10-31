"""Two-grid PCG settings on an optimised (voided) design: iterations, time breakdown and device memory per setting.

Usage: python exp_pcg.py <run_dir> <out.json> [--phase 1 --ckpt 2] --configs '<json list of TwoGridFE kwargs>'
Each setting: fresh LeanGrid + TwoGridFE on the optimiser's active set, factor, primal (cold, tol 1e-10), a 14-column
block solve of G·X-like right-hand sides at tol 1e-4 (the buckling Krylov expansion), and timings of one matvec,
one V-cycle and one coarse solve (1 and 14 columns). Peak device memory from the CuPy pool (engine arrays only).
"""
import argparse
import json
import sys
import time
from pathlib import Path

import numpy as np

ROOT = Path(__file__).resolve().parents[4]
sys.path.insert(0, str(ROOT))


class Peak:
    """CuPy allocator wrapper tracking the peak bytes in use."""

    def __init__(self, cp):
        self.cp, self.pool = cp, cp.get_default_memory_pool()
        self.peak = 0

    def reset(self):
        self.pool.free_all_blocks()
        self.peak = self.pool.used_bytes()

    def poll(self):
        self.peak = max(self.peak, self.pool.used_bytes())
        return self.peak


def timed(cp, f, n=3):
    f(); cp.cuda.Device().synchronize()
    t = time.time()
    for _ in range(n):
        f()
    cp.cuda.Device().synchronize()
    return (time.time() - t) / n


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("run_dir"); ap.add_argument("out")
    ap.add_argument("--phase", type=int, default=1)
    ap.add_argument("--ckpt", type=int, default=2)
    ap.add_argument("--configs", required=True)
    a = ap.parse_args()
    import glob
    import cupy as cp
    from scipy import ndimage
    from cshell import config
    from cshell.optimise import Context, PhaseModel
    from cshell.fe_twogrid import LeanGrid, TwoGridFE
    rd = Path(a.run_dir)
    cfg = config.load(str(rd / "config_used.json"), {"fe": {"engine": "twogrid"}})
    ctx = Context(cfg)
    f = sorted(glob.glob(str(rd / "checkpoints" / f"*_{a.phase}_{a.ckpt:04d}.npz")))[0]
    z = np.load(f)
    model = PhaseModel(ctx, a.phase, float(z["D_comp"]))
    nv = ctx.surf.n_vertices
    xi = z["xi"]
    rho = np.asarray(model.dm.rho(model.dm.forward(xi[:nv], xi[nv:], as_built=True)))
    if model.geometry_device == "gpu":
        model.disc.release()
    margin = int(np.ceil(cfg["grid"]["active_margin_mm"] / model.h_e))
    mask = ndimage.binary_dilation((rho > 0).reshape(model.elem_shape), iterations=margin)
    E, nu = cfg["material"]["E_MPa"], cfg["material"]["nu"]
    pk = Peak(cp)
    res = []
    for kw in json.loads(a.configs):
        pk.reset()
        base = pk.pool.used_bytes()
        t0 = time.time()
        g = LeanGrid(mask, xp=cp)
        fe = TwoGridFE(g, model.fixed, model.load, model.h_e, E, nu, backend="gpu", **kw)
        r = g.gather_elems(rho)
        rho_hat = model.disc.rho_min + (1 - model.disc.rho_min) * r
        t1 = time.time()
        fe.factor(rho_hat)
        t2 = time.time()
        pk.poll()
        out = dict(cfg=kw, ndof=g.ndof, coarse_ndof=fe.stats["coarse_ndof"], setup_s=t1 - t0, factor_s=t2 - t1,
                   lmax=fe.stats["lmax_DinvK"], coarse_est_GiB=fe.stats.get("coarse_factor_estimate_GiB"))
        try:
            t = time.time(); p = fe.primal(); cp.cuda.Device().synchronize()
            out.update(primal_its=fe.stats_solves[-1][1], primal_s=time.time() - t, C=p.compliance)
            pk.poll()
            rng = np.random.default_rng(0)
            B = fe._apply_K(cp.asarray(rng.standard_normal((g.ndof, 14))))   # smooth-ish block: K·noise
            B = fe._cheb(B)
            t = time.time(); fe.solve(B, tol=1e-4); cp.cuda.Device().synchronize()
            out.update(block14_its_1e4=fe.stats_solves[-1][1], block14_s_1e4=time.time() - t)
            pk.poll()
            t = time.time(); fe.solve(B[:, :6]); cp.cuda.Device().synchronize()
            out.update(block6_its_1e10=fe.stats_solves[-1][1], block6_s_1e10=time.time() - t)
            pk.poll()
            x1, x14 = B[:, 0].copy(), cp.ascontiguousarray(B)
            out.update(matvec1_s=timed(cp, lambda: fe._apply_K(x1)), matvec14_s=timed(cp, lambda: fe._apply_K(x14)),
                       vcycle1_s=timed(cp, lambda: fe.precondition(x1[:, None])),
                       vcycle14_s=timed(cp, lambda: fe.precondition(x14)),
                       coarse1_s=timed(cp, lambda: fe._coarse_correction(x1[:, None].copy())),
                       coarse14_s=timed(cp, lambda: fe._coarse_correction(x14)))
            pk.poll()
            del B, x1, x14, p
        except RuntimeError as e:
            out["error"] = str(e)
        out["pool_peak_GiB"] = (pk.peak - base) / 2 ** 30
        out["device_used_GiB"] = (cp.cuda.runtime.memGetInfo()[1] - cp.cuda.runtime.memGetInfo()[0]) / 2 ** 30
        fe.free()
        del fe, g
        import gc
        gc.collect()
        cp.get_default_memory_pool().free_all_blocks()
        print(json.dumps(out, default=float), flush=True)
        res.append(out)
    Path(a.out).write_text(json.dumps(res, indent=1, default=float))


if __name__ == "__main__":
    main()
