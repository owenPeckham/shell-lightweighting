"""Export benchmarks (Task E): B2 rod from a hollow xi (end to end through Context) and cantilever solid/uniform-shell.

usage: python run_export_bench.py {b2|cantilever} H_X OUTDIR [WORKERS]
Writes <OUTDIR>/<case>_h<H_X>/... and prints the summary JSON path. Timings are in the *_export.json files.
"""
import json
import os
import sys
import time

import numpy as np

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../.."))
sys.path.insert(0, ROOT)
from cshell import export as ex  # noqa: E402


def main():
    case, hx, out = sys.argv[1], float(sys.argv[2]), sys.argv[3]
    workers = int(sys.argv[4]) if len(sys.argv) > 4 and sys.argv[4] else None
    out = os.path.join(out, f"{case}_h{hx}")
    os.makedirs(out, exist_ok=True)
    t0 = time.time()
    if case == "b2":
        from cshell import config
        from cshell.optimise import Context
        cfg = config.load(os.path.join(ROOT, "examples/cshell/bench/B2_slender.json"))
        ctx = Context(cfg)
        t_ctx = time.time() - t0
        nv = ctx.surf.n_vertices
        xi = np.concatenate([np.full(nv, 0.5), np.zeros(nv)])           # xi_d = 0.5, xi_a = 0
        s = ex.export_design(ctx, xi, out, h_x=hx, workers=workers)
        s["context_s"] = t_ctx
        res = {"design": s}
        base = ex.export_baselines(ctx, out, h_x=hx, workers=workers)
        res["baselines"] = base
    else:
        from cshell import config
        from vshell import geometry
        cfg = config.load(os.path.join(ROOT, "examples/cshell/cantilever.json"))
        shape = geometry.load_step(cfg["step_file"])
        V, T, TF = geometry.tessellate(shape, cfg["grid"]["tessellation_mm"])
        print(f"tessellation: {len(V)} vertices, {len(T)} triangles, {time.time() - t0:.1f} s", flush=True)
        c = ex.ExportConfig.from_run_cfg(cfg, h_x=hx, workers=workers)
        res = {"baselines": ex.baselines(V, T, out, c)}
    res["wall_s"] = time.time() - t0
    with open(os.path.join(out, "bench.json"), "w") as f:
        json.dump(res, f, indent=1, default=float)
    print("done", res["wall_s"], flush=True)


if __name__ == "__main__":
    main()
