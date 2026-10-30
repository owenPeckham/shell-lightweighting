"""Full objective (J, C, λ_s, λ_b, ∂J/∂ρ̂) on an optimised (voided) design with one FE engine, for PLAN_twogrid.md
acceptance 2. Evaluates the designs of consecutive checkpoints in order (so warm starts behave as in a run) and records
per-evaluation FE time, PCG statistics, peak host RSS; GPU peak comes from the job's nvidia-smi sampling.

Usage: python bench_design.py <run_dir> <out.npz> --engine direct|twogrid [--phase 1] [--ckpts 2,3 | --xi file.npy]
       [--coarse auto] [--krylov-tol 1e-2] [--degree 5] [--range 30] [--nrhs 14] [--eig-tol x]
       python bench_design.py --compare <ref.npz> <new.npz>
J uses a fixed base (C0 = 1, λ0 = 1): J differences between engines are w_K δln C + w_P δln λ, independent of it.
"""
import argparse
import dataclasses
import glob
import hashlib
import json
import os
import resource
import sys
import time
from pathlib import Path

import numpy as np

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT))


def rss_GiB() -> float:
    return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 2 ** 20      # ru_maxrss in KiB on Linux


def run(a):
    from cshell import config
    from cshell.optimise import Context, PhaseModel
    from cshell.fe_sparse import _host
    if a.det:                                           # two-grid determinism: full | kernels | off (atomic kernels)
        os.environ["CSHELL_TWOGRID_DET"] = a.det
    rd = Path(a.run_dir)
    fe_over = {"engine": a.engine}
    if a.engine == "twogrid":
        fe_over["twogrid"] = {"coarse_precision": a.coarse, "krylov_tol": a.krylov_tol, "nrhs": a.nrhs,
                              "degree": a.degree, "smoothing_range": a.range,
                              "deterministic": a.det or "off"}
    cfg = config.load(str(rd / "config_used.json"), {"fe": fe_over})
    t0 = time.time()
    ctx = Context(cfg)
    designs = []
    if a.xi:
        z = sorted(glob.glob(str(rd / "checkpoints" / "*.npz")))[-1]
        designs.append((Path(a.xi).stem, np.load(a.xi), float(np.load(z)["c_norm"]), float(np.load(z)["D_comp"])))
    else:
        for i in [int(s) for s in a.ckpts.split(",")]:
            f = sorted(glob.glob(str(rd / "checkpoints" / f"*_{a.phase}_{i:04d}.npz")))[0]
            z = np.load(f)
            designs.append((Path(f).stem, z["xi"], float(z["c_norm"]), float(z["D_comp"])))
    model = PhaseModel(ctx, a.phase, designs[0][3])
    rss_setup = rss_GiB()
    nv = ctx.surf.n_vertices
    st = model.settings(1.0)
    if a.eig_tol:
        st = dataclasses.replace(st, eig_tol=a.eig_tol)
    if a.no_buckling:                                   # dJ without the eigenvector-dependent term (engine accuracy)
        st = dataclasses.replace(st, buckling=False)
    base = dict(C0=1.0, lam0=1.0)
    out, rows, prev = {}, [], None
    for name, xi, c_norm, _ in designs:
        rho = np.asarray(_host(model.dm.rho(model.dm.forward(xi[:nv], xi[nv:], as_built=True))))
        rho_sha = hashlib.sha1(rho.tobytes()).hexdigest()[:12]            # identical across runs <=> design chain bitwise
        t1 = time.time()
        fe = model.fe_eval(rho, st, base, c_norm, prev)
        t_fe = time.time() - t1
        prev = fe["modes"]
        s = dict(model.fe.stats)
        pcg = model.fe.pcg_summary() if hasattr(model.fe, "pcg_summary") else {}
        row = dict(design=name, J=fe["J"], C=fe["C"], lam_s=fe["lam_s"], lam_b=fe["lam_b"], mu_tilde=fe["mu_tilde"],
                   eig_iters=fe["eig_iters"], t_fe=t_fe, ndof=int(model.grid_fe.ndof), engine=model.fe_engine,
                   pcg=pcg, rss_GiB=rss_GiB(), rho_sha=rho_sha,
                   stats={k: v for k, v in s.items() if isinstance(v, (int, float, str, bool))})
        print(json.dumps(row, default=float), flush=True)
        rows.append(row)
        out[f"{name}_dJ"] = np.asarray(fe["dJ"])
        out[f"{name}_elem"] = np.asarray(model.grid_fe.elem_dense)
        for rep_i in range(a.repeat):       # same rho again, cold start: FE-only run-to-run reproducibility
            if hasattr(model.fe, "_u_prev"):
                model.fe._u_prev = None
            fe2 = model.fe_eval(rho, st, base, c_norm, None)
            same = {k: bool(np.array_equal(np.asarray(fe2[k]), np.asarray(fe[k])))
                    for k in ("J", "C", "lam_s", "lam_b", "mu_tilde", "dJ")}
            print("REPEAT", name, rep_i + 1, json.dumps(same), "dJ max|diff|",
                  float(np.abs(np.asarray(fe2["dJ"]) - np.asarray(fe["dJ"])).max()), "J diff", fe2["J"] - fe["J"],
                  flush=True)
            row.setdefault("repeat_bitwise", []).append(all(same.values()))
    info = dict(run_dir=str(rd), phase=a.phase, engine=a.engine, rows=rows, rss_setup_GiB=rss_setup,
                rss_peak_GiB=rss_GiB(), total_s=time.time() - t0, eig_tol=st.eig_tol,
                coarse=a.coarse, krylov_tol=a.krylov_tol, nrhs=a.nrhs,
                degree=a.degree, range=a.range)
    np.savez(a.out, info=json.dumps(info, default=float), **out)
    print("INFO", json.dumps({k: v for k, v in info.items() if k != "rows"}, default=float), flush=True)


def compare(ref_path, new_path):
    r, n = np.load(ref_path), np.load(new_path)
    ri, ni = json.loads(str(r["info"])), json.loads(str(n["info"]))
    res = []
    for rr, nr in zip(ri["rows"], ni["rows"]):
        d = rr["design"]
        # dJ on the union of both active sets (elements outside an active set carry no gradient)
        e = np.union1d(r[f"{d}_elem"], n[f"{d}_elem"])
        gr, gn = np.zeros(e.size), np.zeros(e.size)
        gr[np.searchsorted(e, r[f"{d}_elem"])] = r[f"{d}_dJ"]
        gn[np.searchsorted(e, n[f"{d}_elem"])] = n[f"{d}_dJ"]
        rel = lambda k: abs(nr[k] - rr[k]) / abs(rr[k]) if rr[k] and np.isfinite(rr[k]) else abs(nr[k] - rr[k])
        c = dict(design=d, J_ref=rr["J"], J_rel=rel("J"), C_rel=rel("C"), lam_s_rel=rel("lam_s"),
                 lam_b_rel=rel("lam_b"), mu_rel=rel("mu_tilde"),
                 dJ_rel2=float(np.linalg.norm(gn - gr) / np.linalg.norm(gr)),
                 bitwise=bool(r[f"{d}_elem"].size == n[f"{d}_elem"].size and np.array_equal(gn, gr)
                              and rr["J"] == nr["J"] and rr["C"] == nr["C"] and rr["lam_b"] == nr["lam_b"]),
                 dJ_max_abs=float(np.abs(gn - gr).max()), rho_same=rr.get("rho_sha") == nr.get("rho_sha"),
                 t_fe_ref=rr["t_fe"], t_fe_new=nr["t_fe"], engines=(rr["engine"], nr["engine"]),
                 eig_iters=(rr["eig_iters"], nr["eig_iters"]))
        res.append(c)
        print(json.dumps(c, default=float))
    print("RSS peak GiB ref/new", ri["rss_peak_GiB"], ni["rss_peak_GiB"])
    return res


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--compare":
        compare(sys.argv[2], sys.argv[3])
        sys.exit(0)
    ap = argparse.ArgumentParser()
    ap.add_argument("run_dir"); ap.add_argument("out")
    ap.add_argument("--engine", default="twogrid")
    ap.add_argument("--phase", type=int, default=1)
    ap.add_argument("--ckpts", default="2,3")
    ap.add_argument("--xi", default=None)
    ap.add_argument("--coarse", default="auto")
    ap.add_argument("--krylov-tol", type=float, default=1e-2)
    ap.add_argument("--degree", type=int, default=5)
    ap.add_argument("--range", type=float, default=30.0)
    ap.add_argument("--nrhs", type=int, default=7)
    ap.add_argument("--eig-tol", type=float, default=None)
    ap.add_argument("--no-buckling", action="store_true")
    ap.add_argument("--repeat", type=int, default=0, help="re-evaluate each design this many times on the same rho")
    ap.add_argument("--det", default=None, choices=["full", "kernels", "off"],
                    help="two-grid determinism mode (default: the engine's, full)")
    run(ap.parse_args())
