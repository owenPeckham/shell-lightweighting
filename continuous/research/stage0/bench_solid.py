"""Stage 0 (STATUS §2.2): solid baselines (V2-a) and per-solve cost of the sparse cuDSS engine on the Study 1 parts,
plus the dense reference ShellCG at 0.8 mm for comparison. Geometry and volume fractions come from vshell's PartGrid,
so V2-a compares the FE alone. Usage: python bench_solid.py <part.json> <out.json> [h_e ...]"""
import json
import sys
import time
from pathlib import Path

import numpy as np

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT))
from vshell import config as vcfg, field, geometry  # noqa: E402
from cshell import fea  # noqa: E402
from cshell.fe_sparse import ActiveGrid, SparseFE, buckling  # noqa: E402

RHO_MIN = 1e-6


def gpu_mem():
    import cupy as cp
    free, total = cp.cuda.runtime.memGetInfo()
    return (total - free) / 2**30


def main(cfg_path, out_path, hs):
    cfg = vcfg.load(cfg_path)
    shape = geometry.load_step(cfg["step_file"])
    faces = geometry.face_table(shape)
    sup = sorted({i for s in cfg["supports"] for i in geometry.select_faces(faces, s)})
    loads = [dict(faces=geometry.select_faces(faces, L["faces"]), force=np.asarray(L["force_N"], float)) for L in cfg["loads"]]
    F = np.sum([L["force"] for L in loads], 0)
    V, T, TF = geometry.tessellate(shape, cfg["grid"]["tessellation_mm"])
    E, nu = cfg["material"]["E_MPa"], cfg["material"]["nu"]
    res = dict(part=cfg["name"], F=F.tolist(), runs=[])
    for h_e in hs:
        t0 = time.time()
        grid = field.PartGrid(V, T, h=h_e / 2, k_fe=2, band=4.0, pad=4.0)
        vf, origin, hfe = grid.fe_fractions(grid.phi)
        t_sdf = time.time() - t0
        nn = tuple(n + 1 for n in vf.shape)
        fixed, load, info = fea.build_bcs(origin, hfe, nn, V, T, TF, sup, loads)
        rho = RHO_MIN + (1 - RHO_MIN) * vf.astype(np.float64)
        mask = vf > 0
        r = dict(h_e=hfe, elem_box=list(vf.shape), dense_dof=3 * int(np.prod(nn)), sdf_s=t_sdf, **info)
        t0 = time.time(); g = ActiveGrid(mask); r["pattern_s"] = time.time() - t0
        fe = SparseFE(g, fixed, load, hfe, E, nu, backend="gpu")
        r.update(active_elems=g.Ne, sparse_dof=g.ndof, nnz=g.nnz)
        t0 = time.time(); fe.factor(g.gather_elems(rho)); r["first_factor_s"] = time.time() - t0
        r.update({k: fe.stats.get(k) for k in ("plan_s", "assemble_s", "factor_s")})
        r["gpu_mem_after_factor_GiB"] = gpu_mem()
        t0 = time.time(); fe.factor(g.gather_elems(rho)); r["refactor_s"] = time.time() - t0
        p = fe.primal()
        r.update(solve_s=p.seconds["solve"], compliance_Nmm=p.compliance, K0=float(F @ F / p.compliance))
        try:
            t0 = time.time()
            b = buckling(fe, p, num_modes=6, p=8.0, tol=1e-6, maxiter=200)
            r.update(buckling_s=time.time() - t0, mu=b.mu.tolist(), lambda1=float(1 / b.mu.max()) if b.mu.size else None,
                     P_buckle_N=float(np.linalg.norm(F) / b.mu.max()) if b.mu.size else None,
                     lobpcg_iters=b.iterations, buckling_converged=b.converged, energy_fraction=b.energy_fraction.tolist(),
                     gpu_mem_after_buckling_GiB=gpu_mem())
        except Exception as e:                                  # record, never hide
            r["buckling_error"] = repr(e)
        fe.free()
        if hfe >= 0.79:                                       # dense reference at 0.8 mm
            try:
                import gpu_analysis
                nx, ny, nz = nn
                r["dense_mem_model_GiB"] = gpu_analysis.elastic_memory("gpu_gmg", nx, ny, nz) / 2**30
                cg = fea.ShellCG(fixed, load, hfe, E, nu, rel_tol=1e-6, warm_start=False)
                t0 = time.time(); s = cg.solve(rho.astype(np.float32))
                r.update(dense_solve_s=time.time() - t0, dense_iters=s.iterations, dense_converged=s.converged,
                         dense_K0=float(F @ F / s.compliance), dense_engine=s.engine)
            except Exception as e:
                r["dense_error"] = repr(e)
        res["runs"].append(r)
        print(json.dumps(r), flush=True)
        Path(out_path).write_text(json.dumps(res, indent=1))


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2], [float(h) for h in sys.argv[3:]] or [0.8, 0.4])
