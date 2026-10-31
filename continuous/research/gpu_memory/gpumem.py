"""GPU-memory study of cshell's sparse FE engine (cshell/fe_sparse.py, read-only; imported, never edited).

Phases (select with --phases, default all):
  est     cuDSS plan-only memory estimates for reorderings (fp64/fp32), incl. hybrid minimum device memory
  fp64    today's engine: fp64 factor, primal, buckling with today's block Krylov and with a lean (no-copy) variant
  fp32    fp32 factor + fp64 PCG (engine's existing path): primal, buckling (lean eigensolver), gradients vs fp64
  hybrid  cuDSS hybrid memory mode (fp64 factor with a device-memory cap), factor/solve time
Usage: python gpumem.py <part.json> <out.json> --h 0.4 [--hollow 4] [--phases est,fp64,fp32,hybrid]
Units: memory GiB (2^30 B), time s, compliance N mm.
"""
import argparse
import os
import json
import sys
import time
from pathlib import Path

import numpy as np

ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(ROOT))
import cupy as cp  # noqa: E402
import cupyx.scipy.sparse as cpsp  # noqa: E402
from scipy import ndimage  # noqa: E402

from vshell import config as vcfg, field, geometry  # noqa: E402
from cshell import fea  # noqa: E402
import cshell.fe_sparse as fs  # noqa: E402
from cshell.fe_sparse import ActiveGrid, SparseFE  # noqa: E402

GiB = 2.0 ** 30
RHO_MIN = 1e-6
POOL = cp.get_default_memory_pool()
PEAK = [0]


def _alloc(n):                                    # CuPy allocator that records the pool's peak live bytes
    m = POOL.malloc(n)
    u = POOL.used_bytes()
    if u > PEAK[0]:
        PEAK[0] = u
    return m


cp.cuda.set_allocator(_alloc)


def dev_used():
    cp.cuda.Device().synchronize()
    f, t = cp.cuda.runtime.memGetInfo()
    return (t - f) / GiB


def mem_snapshot(tag, base):
    return dict(tag=tag, device_used=dev_used(), over_base=dev_used() - base, cupy_used=POOL.used_bytes() / GiB,
                cupy_held=POOL.total_bytes() / GiB)


def setup(cfg_path, h_e, hollow, grid_cls=ActiveGrid):
    cfg = vcfg.load(cfg_path)
    shape = geometry.load_step(cfg["step_file"])
    faces = geometry.face_table(shape)
    sup = sorted({i for s in cfg["supports"] for i in geometry.select_faces(faces, s)})
    loads = [dict(faces=geometry.select_faces(faces, L["faces"]), force=np.asarray(L["force_N"], float))
             for L in cfg["loads"]]
    V, T, TF = geometry.tessellate(shape, cfg["grid"]["tessellation_mm"])
    grid = field.PartGrid(V, T, h=h_e / 2, k_fe=2, band=4.0, pad=4.0)
    vf, origin, hfe = grid.fe_fractions(grid.phi)
    nn = tuple(n + 1 for n in vf.shape)
    fixed, load, _ = fea.build_bcs(origin, hfe, nn, V, T, TF, sup, loads)
    mask = vf > 0
    if hollow:                                    # skin design: keep a wall of `hollow` elements, remove the core
        core = ndimage.binary_erosion(mask, iterations=hollow)
        mask &= ~core
    rho = RHO_MIN + (1 - RHO_MIN) * vf.astype(np.float64)
    g = grid_cls(mask)
    return cfg, g, fixed, load, hfe, rho, float(np.linalg.norm(np.sum([L["force"] for L in loads], 0)))


def engine_bytes(fe):
    out = {}
    for k in ("_keep", "_perm", "_pos_blk", "_dofs", "_diag_fixed", "f"):
        a = getattr(fe, k, None)
        if a is not None:
            out[k] = a.nbytes / GiB
    if fe._K is not None:
        out["K"] = (fe._K.data.nbytes + fe._K.indices.nbytes + fe._K.indptr.nbytes) / GiB
        if fe._Kf is not None and fe._Kf is not fe._K:
            out["Kf32"] = fe._Kf.data.nbytes / GiB          # shares indices? (counted data only)
    return out


# ---------------------------------------------------------------------------------------------- lean eigensolver
def _chunks(n, c):
    return [slice(i, min(i + c, n)) for i in range(0, n, c)]


def _gram(Q, A, cols=8):
    """Qᵀ A Q, A applied to `cols` columns at a time (no n×k product stored)."""
    k = Q.shape[1]
    M = cp.empty((k, k))
    for s in _chunks(k, cols):
        M[:, s] = Q.T @ (A @ Q[:, s])
    return 0.5 * (M + M.T)


def _rmul_inplace(Q, W, rows=1 << 18):
    """Q[:, :W.shape[1]] ← Q W in row blocks (rows independent); returns the view."""
    k = W.shape[1]
    for s in _chunks(Q.shape[0], rows):
        Q[s, :k] = Q[s] @ W
    return Q[:, :k]


KRYLOV_SOLVE = [None]       # optional cheaper inner solve for the Krylov expansion and residual norm (not for η)
EIG = dict(extra=0.0)


def _eig_peak(fn):
    """Record the CuPy peak live bytes inside the eigensolver call, above its entry level."""
    def w(*a, **k):
        before, u0 = PEAK[0], POOL.used_bytes()
        PEAK[0] = u0
        try:
            return fn(*a, **k)
        finally:
            EIG["extra"] = (PEAK[0] - u0) / GiB
            PEAK[0] = max(PEAK[0], before)
    return w


def direct32(fe):
    """One fp32-factor solve per column (no fp64 refinement), fixed DOF zeroed."""
    fx = cp.asarray(fe.fixed)

    def s(B):
        one = B.ndim == 1
        B = B[:, None] if one else B
        X = cp.empty_like(B, dtype=cp.float64)
        for j in range(B.shape[1]):
            b = B[:, j].copy(); b[fx] = 0
            X[:, j] = fe._direct(b); X[fx, j] = 0
        return X[:, 0] if one else X
    return s


def pcg_loose(fe, tol):
    """fp64 CG preconditioned by the fp32 factor, loose tolerance (Krylov expansion only)."""
    def s(B):
        one = B.ndim == 1
        B = B[:, None] if one else B
        X = cp.empty_like(B, dtype=cp.float64)
        for j in range(B.shape[1]):
            b = B[:, j].copy(); b[cp.asarray(fe.fixed)] = 0
            X[:, j] = fe._pcg(b, tol=tol)
        return X[:, 0] if one else X
    return s


def lean_block_krylov(fe, G, X, nev, tol, maxiter, steps=3, n_neg=4):
    """Same iteration and Rayleigh-Ritz as fs._block_krylov_largest, but without the blocks list, KQ and GQ:
    one preallocated basis Q, Gram matrices by column chunks, in-place row-block products, and the residual from
    two matvecs on the nev Ritz vectors. Peak ≈ Q + (m + n_neg) restart vectors + solve buffers."""
    n, m = X.shape
    slv = KRYLOV_SOLVE[0] or fe.solve
    res = np.full(nev, np.inf)
    Xn, it, th = None, 0, None
    for it in range(1, maxiter + 1):
        X0 = X if Xn is None else cp.concatenate([X, Xn], 1)
        c0 = X0.shape[1]
        Q = cp.empty((n, c0 * (steps + 1)), order="F")       # column-major: column slices contiguous
        Q[:, :c0] = X0
        V = X0
        del X0, X, Xn
        for s in range(steps):
            V = slv(G @ V)
            V /= cp.linalg.norm(V, axis=0)[None, :]
            Q[:, (s + 1) * c0:(s + 2) * c0] = V
        del V
        for _ in range(2):                                   # K-orthonormalise twice (rank-revealing)
            w, U = cp.linalg.eigh(_gram(Q, fe._K))
            keep = w > 1e-10 * w.max()
            Q = _rmul_inplace(Q, U[:, keep] / cp.sqrt(w[keep])[None, :])
        theta, Y = cp.linalg.eigh(_gram(Q, G))
        Y_pos, th = Y[:, ::-1][:, :m], theta[::-1][:m]
        Xn = Q @ Y[:, :n_neg]
        X = Q @ Y_pos
        del Q
        POOL.free_all_blocks()
        Xv = X[:, :nev]
        R = G @ Xv - (fe._K @ Xv) * th[None, :nev]
        res = fs._host(cp.sqrt(cp.abs(cp.einsum("ij,ij->j", R, slv(R)))) / (cp.abs(th[:nev]) + 1e-300))
        del R
        if np.all(res[:nev] <= tol):
            break
    return fs._host(th), X, res, it


ORIG_KRYLOV = fs._block_krylov_largest


def run_buckling(fe, p, lean, tol, X0=None, inner=None):
    fs._block_krylov_largest = _eig_peak(lean_block_krylov if lean else ORIG_KRYLOV)
    KRYLOV_SOLVE[0] = inner
    POOL.free_all_blocks()
    u0 = POOL.used_bytes()
    PEAK[0] = u0
    d0 = dev_used()
    t0 = time.time()
    b = fs.buckling(fe, p, num_modes=6, p=8.0, tol=tol, maxiter=100, X0=X0)
    cp.cuda.Device().synchronize()
    dt = time.time() - t0
    out = dict(lean=lean, seconds=dt, iters=b.iterations, converged=b.converged, mu=b.mu.tolist(),
               mu_tilde=b.mu_tilde, peak_extra_GiB=(PEAK[0] - u0) / GiB, eig_peak_extra_GiB=EIG["extra"],
               device_before=d0, inner="direct32" if inner else "default")
    fs._block_krylov_largest = ORIG_KRYLOV
    KRYLOV_SOLVE[0] = None
    return out, b


def rel(a, b):
    a, b = np.asarray(a, float), np.asarray(b, float)
    return float(np.linalg.norm(a - b) / max(np.linalg.norm(b), 1e-300))


# --------------------------------------------------------------------------------------------------- phases
def phase_est(fe, rho_act, res, base):
    """Plan-only memory estimates (no factorisation) for reorderings and precisions."""
    from nvmath.sparse.advanced import (DirectSolver, DirectSolverMatrixType, DirectSolverOptions,
                                        DirectSolverReorderingAlg as RA, ExecutionCUDA, HybridMemoryModeOptions)
    data = fe._values(cp.asarray(rho_act))
    g = fe.g
    ind, ptr = cp.asarray(g.indices.astype(np.int32)), cp.asarray(g.indptr.astype(np.int32))
    out = []
    cases = [("fp64", "DEFAULT", {}), ("fp64", "AMD", {}), ("fp64", "NESTED_DISSECTION", {}),
             ("fp64", "DEFAULT", dict(use_superpanels=False)), ("fp32", "DEFAULT", {}), ("fp32", "AMD", {}),
             ("fp64", "DEFAULT", dict(hybrid=True))]
    for prec, alg, extra in cases:
        dt = np.float32 if prec == "fp32" else np.float64
        K = cpsp.csr_matrix((data.astype(dt), ind, ptr), shape=(g.ndof, g.ndof))
        b = cp.asfortranarray(cp.zeros((g.ndof, 1), dtype=dt))
        ex = ExecutionCUDA(hybrid_memory_mode_options=HybridMemoryModeOptions(hybrid_memory_mode=True)) \
            if extra.get("hybrid") else ExecutionCUDA()
        s = DirectSolver(K, b, options=DirectSolverOptions(sparse_system_type=DirectSolverMatrixType.SPD,
                                                           multithreading_lib=fs._cudss_mtlayer()), execution=ex)
        r = dict(prec=prec, reorder=alg, **{k: v for k, v in extra.items()})
        try:
            s.plan_config.reordering_algorithm = getattr(RA, alg)
            if "use_superpanels" in extra:
                s.plan_config.use_superpanels = extra["use_superpanels"]
            t0 = time.time()
            s.plan()
            r["plan_s"] = time.time() - t0
            e = s.plan_info.memory_estimates
            names = ["permanent_device_memory", "peak_device_memory", "permanent_host_memory", "peak_host_memory",
                     "hybrid_min_device_memory", "hybrid_max_device_memory"]
            r.update({n: float(e[n]) / GiB for n in names} if e.dtype.names else
                     {n: float(v) / GiB for n, v in zip(names, np.asarray(e).ravel()[:6])})
        except Exception as ex_:
            r["error"] = repr(ex_)[:400]
        s.free()
        del K, b, s
        POOL.free_all_blocks()
        print("est", json.dumps(r), flush=True)
        out.append(r)
    del data
    POOL.free_all_blocks()
    res["estimates"] = out


def phase_engine(fe, rho_act, res, base, prec, Fn, ref=None, tol=1e-3, orig_too=True):
    fe.precision = prec
    POOL.free_all_blocks()
    t0 = time.time()
    fe.factor(cp.asarray(rho_act))
    r = dict(precision=prec, first_factor_s=time.time() - t0, plan_s=fe.stats.get("plan_s"))
    r["mem_after_factor"] = mem_snapshot("after_factor", base)
    POOL.free_all_blocks()
    r["mem_after_factor_poolfreed"] = mem_snapshot("after_factor_poolfreed", base)
    r["engine_arrays_GiB"] = engine_bytes(fe)
    try:
        r["lu_nnz"] = int(fe._solver.factorization_info.lu_nnz)
    except Exception as e:
        r["lu_nnz_error"] = repr(e)[:200]
    t0 = time.time(); fe.factor(cp.asarray(rho_act)); r["refactor_s"] = time.time() - t0
    fe.stats["pcg_its"] = 0
    t0 = time.time(); p = fe.primal(); cp.cuda.Device().synchronize(); r["primal_s"] = time.time() - t0
    r["pcg_its_primal"] = fe.stats.get("pcg_its", 0)
    r["compliance"] = p.compliance
    r["resid_rel"] = float(cp.linalg.norm(fe._K @ p.u - fe.f * cp.asarray(~fe.fixed)) / cp.linalg.norm(fe.f))
    dC = fe.dC_drho_hat(p)
    # second (adjoint-like) RHS: a random load, to time a solve with a non-smooth RHS
    rng = np.random.default_rng(1)
    q = cp.asarray(rng.standard_normal(fe.g.ndof))
    fe.stats["pcg_its"] = 0
    t0 = time.time(); y = fe.solve(q); cp.cuda.Device().synchronize(); r["random_rhs_solve_s"] = time.time() - t0
    r["pcg_its_random"] = fe.stats.get("pcg_its", 0)
    qq = q.copy(); qq[cp.asarray(fe.fixed)] = 0
    r["random_resid_rel"] = float(cp.linalg.norm(fe._K @ y - qq) / cp.linalg.norm(qq))
    del y, q, qq
    bk = {}
    variants = [("orig", False, None), ("lean", True, None)] if orig_too else \
        [("lean", True, None), ("lean_pcg1e-4inner", True, pcg_loose(fe, 1e-4))]
    for key, lean, inner in variants:
        fe.stats["pcg_its"] = 0
        try:
            o, b = run_buckling(fe, p, lean, tol, inner=inner)
            o["pcg_its"] = fe.stats.get("pcg_its", 0)
            o["mem_after"] = dev_used()
            o["lambda1"] = float(1 / max(b.mu)) if len(b.mu) else None
            o["P_buckle_N"] = Fn / max(b.mu) if len(b.mu) else None
            bk[key] = o
            last = b
            if ref is not None and ref.get("mu") is not None and len(b.mu) >= len(ref["mu"]):
                o["vs_fp64"] = dict(mu_rel=rel(b.mu[:len(ref["mu"])], ref["mu"]),
                                    mu_tilde_rel=abs(b.mu_tilde - ref["mu_tilde"]) / ref["mu_tilde"],
                                    dmu_rel=rel(b.dmu_tilde, ref["dmu"]))
        except Exception as e:
            bk[key] = dict(error=repr(e)[:400])
            last = None
        print(prec, "buckling", key, json.dumps(bk[key]), flush=True)
        POOL.free_all_blocks()
    r["buckling"] = bk
    r["device_peak_cupy_GiB"] = PEAK[0] / GiB
    if ref is not None:
        r["vs_fp64"] = dict(compliance_rel=abs(p.compliance - ref["C"]) / ref["C"], dC_rel=rel(dC, ref["dC"]))
        if last is not None and ref.get("mu") is not None:
            r["vs_fp64"].update(mu_rel=rel(last.mu[:len(ref["mu"])], ref["mu"]) if len(last.mu) >= len(ref["mu"])
                                else None, mu_tilde_rel=abs(last.mu_tilde - ref["mu_tilde"]) / ref["mu_tilde"],
                                dmu_rel=rel(last.dmu_tilde, ref["dmu"]))
    out_ref = dict(C=p.compliance, dC=dC)
    adir = os.environ.get("GM_ARR_DIR")
    if adir:
        Path(adir).mkdir(parents=True, exist_ok=True)
        stem = Path(adir) / f"{res.get('_stem', 'run')}_{prec}"
        np.save(f"{stem}_dC.npy", dC)
        if last is not None:
            np.save(f"{stem}_dmu.npy", last.dmu_tilde)
    if last is not None:
        out_ref.update(mu=np.asarray(last.mu), mu_tilde=last.mu_tilde, dmu=last.dmu_tilde)
        if ref is None and "orig" in bk and "lean" in bk and "mu" in bk["orig"]:
            r["lean_vs_orig_mu_rel"] = rel(bk["lean"]["mu"][:len(bk["orig"]["mu"])], bk["orig"]["mu"])
    del p
    fe.free(); fe._K = None; fe._Kf = None
    POOL.free_all_blocks()
    r["mem_after_free"] = mem_snapshot("after_free", base)
    res[prec] = r
    print(prec, json.dumps({k: v for k, v in r.items() if k != "buckling"}), flush=True)
    return out_ref


NRHS = [1]


def phase_hybrid(fe, rho_act, res, base, limits_GiB, ref):
    from nvmath.sparse.advanced import (DirectSolver, DirectSolverMatrixType, DirectSolverOptions, ExecutionCUDA,
                                        HybridMemoryModeOptions)
    g = fe.g
    out = []
    data = fe._values(cp.asarray(rho_act))
    K = cpsp.csr_matrix((data, cp.asarray(g.indices.astype(np.int32)), cp.asarray(g.indptr.astype(np.int32))),
                             shape=(g.ndof, g.ndof))
    del data
    f = fe.f.copy()
    f[cp.asarray(fe.fixed)] = 0
    for lim in limits_GiB:
        for reg in (True, False):
            r = dict(limit_GiB=lim, register_cuda_memory=reg)
            b = cp.asfortranarray(cp.zeros((g.ndof, NRHS[0])))
            r["nrhs"] = NRHS[0]
            try:
                hm = HybridMemoryModeOptions(hybrid_memory_mode=True, register_cuda_memory=reg,
                                             hybrid_device_memory_limit=None if lim is None else int(lim * GiB))
                ex = ExecutionCUDA() if lim == 0 else ExecutionCUDA(hybrid_memory_mode_options=hm)
                s = DirectSolver(K, b, options=DirectSolverOptions(sparse_system_type=DirectSolverMatrixType.SPD,
                                                                   multithreading_lib=fs._cudss_mtlayer()),
                                 execution=ex)
                POOL.free_all_blocks()
                t0 = time.time(); s.plan(); r["plan_s"] = time.time() - t0
                e = s.plan_info.memory_estimates
                r["hybrid_min_device_GiB"] = float(e["hybrid_min_device_memory"]) / GiB
                t0 = time.time(); s.factorize(); cp.cuda.Device().synchronize(); r["factor_s"] = time.time() - t0
                r["device_used_after_factor"] = dev_used()
                r["cupy_held_after_factor"] = POOL.total_bytes() / GiB
                t0 = time.time(); s.factorize(); cp.cuda.Device().synchronize(); r["refactor_s"] = time.time() - t0
                ts = []
                for _ in range(3):
                    b[...] = f[:, None]
                    t0 = time.time(); x = cp.asarray(s.solve())[:, 0]; cp.cuda.Device().synchronize()
                    ts.append(time.time() - t0)
                r["solve_s"] = ts
                C = float(f @ x)
                r["compliance_rel_vs_fp64"] = abs(C - ref["C"]) / ref["C"] if ref else None
                r["compliance"] = C
                s.free()
            except Exception as ex_:
                r["error"] = repr(ex_)[:400]
                try:
                    s.free()
                except Exception:
                    pass
            del b
            POOL.free_all_blocks()
            print("hybrid", json.dumps(r), flush=True)
            out.append(r)
            if "error" not in r or lim == 0:
                break                                        # one registration setting is enough once it works
    res["hybrid"] = out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("part"); ap.add_argument("out")
    ap.add_argument("--h", type=float, default=0.4)
    ap.add_argument("--hollow", type=int, default=0)
    ap.add_argument("--phases", default="est,fp64,fp32,hybrid")
    ap.add_argument("--hybrid-limits", default="20,40")
    ap.add_argument("--nrhs", type=int, default=1)
    a = ap.parse_args()
    NRHS[0] = a.nrhs
    phases = a.phases.split(",")
    base = dev_used()
    t0 = time.time()
    cfg, g, fixed, load, hfe, rho, Fn = setup(a.part, a.h, a.hollow)
    res = dict(part=cfg["name"], h_e=hfe, hollow=a.hollow, Ne=g.Ne, ndof=g.ndof, nnz=g.nnz, setup_s=time.time() - t0,
               context_GiB=base, F_N=Fn)
    print(json.dumps(res), flush=True)
    fe = SparseFE(g, fixed, load, hfe, cfg["material"]["E_MPa"], cfg["material"]["nu"], backend="gpu",
                  precision="fp64")
    rho_act = g.gather_elems(rho)
    res["mem_after_engine_init"] = mem_snapshot("engine_init", base)
    res["_stem"] = Path(a.out).stem
    save = lambda: Path(a.out).write_text(json.dumps(res, indent=1, default=float))  # noqa: E731
    if "est" in phases:
        phase_est(fe, rho_act, res, base); save()
    ref = None
    if "fp64" in phases:
        ref = phase_engine(fe, rho_act, res, base, "fp64", Fn); save()
    if "fp32" in phases:
        phase_engine(fe, rho_act, res, base, "fp32", Fn, ref=ref, orig_too=False); save()
    if "hybrid" in phases:
        phase_hybrid(fe, rho_act, res, base, [float(x) for x in a.hybrid_limits.split(",")], ref); save()
    res["total_s"] = time.time() - t0
    save()


if __name__ == "__main__":
    main()
