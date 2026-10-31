"""Two-grid PCG for cshell's q1_im voxel FE: fine operator K_h (assembled CSR, or matrix-free), Chebyshev-Jacobi
smoother, exact Galerkin coarse operator K_2h = Pᵀ K_h P assembled element by element (a fine element's nodes
interpolate only from its parent coarse element's 8 nodes, so K_2h,E = Σ_children E ρ_c P_cᵀ K̂ P_c, eight fixed
24x24 matrices), factorised by cuDSS (fp32 by default). Measures iterations, time and memory; with --buckling it also
runs today's buckling() with the two-grid PCG as the inner solver. This is the 0.2 mm route tested one level down
(fine 0.4 mm, coarse 0.8 mm).
Usage: python twogrid.py <part.json> <out.json> --h 0.4 [--hollow 4] [--deg 3] [--coarse fp32] [--buckling]
"""
import argparse
import os
import json
import sys
import time
from pathlib import Path

import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parent))
import gpumem as gm  # noqa: E402  (sets the peak-tracking CuPy allocator; imports cshell from the worktree)
import cupy as cp  # noqa: E402
import cupyx.scipy.sparse as cpsp  # noqa: E402
import cupyx  # noqa: E402
import scipy.sparse as sps  # noqa: E402
from scipy import ndimage  # noqa: E402

import cshell.fe_sparse as fs  # noqa: E402
from cshell.fe_sparse import ActiveGrid, SparseFE  # noqa: E402

GiB = gm.GiB
C = fs.CORNERS


def child_P():
    """P_c (24x24) for the 8 children: fine node a of child c at fine offset c+C[a] in {0,1,2}^3 interpolates from the
    coarse element's node A with trilinear weight N_A(ξ), ξ = (c+C[a])/2."""
    out = np.zeros((8, 24, 24))
    for ci, c in enumerate(C):
        for a in range(8):
            xi = (c + C[a]) / 2.0
            for A in range(8):
                w = np.prod(np.where(C[A] == 1, xi, 1 - xi))
                for k in range(3):
                    out[ci, 3 * a + k, 3 * A + k] = w
    return out


class CoarseFE(SparseFE):
    """Galerkin coarse operator; rho_hat is (Nec, 8): fine ρ̂ of each child (0 if the child is inactive)."""

    def set_children(self, Mc):
        self._Mc = cp.asarray(Mc.reshape(8, 576))

    def _values(self, rho_hat):
        xp, g = self.xp, self.g
        blocks = xp.zeros((g.nblk, 3, 3))
        Ke = (self.E * rho_hat) @ self._Mc                                   # (Nec, 576)
        Ke = Ke.reshape(-1, 8, 3, 8, 3)
        for ab in range(64):
            a, b = divmod(ab, 8)
            cupyx.scatter_add(blocks, self._pos_blk[:, ab], Ke[:, a, :, b, :])
        data = xp.empty(g.nnz)
        data[self._perm] = blocks.reshape(-1)
        data *= self._keep
        data[self._diag_fixed] = 1.0
        # guard: coarse DOF with no stiffness (all supporting fine nodes outside material) -> identity
        diag = data[self._diag_all]
        z = diag <= 1e-14 * float(cp.abs(diag).max())
        self.stats["coarse_zero_diag"] = int(z.sum())
        data[self._diag_all[z]] = 1.0
        return data


def build_twogrid(g, fixed_dense, rho_dense, h, E, nu, coarse_prec, Khat_fine):
    t0 = time.time()
    es = np.array(g.elem_shape)
    esc = (es + 1) // 2
    mask = np.zeros(es, bool).ravel(); mask[g.elem_dense] = True; mask = mask.reshape(es)
    mpad = np.zeros(2 * esc, bool); mpad[:es[0], :es[1], :es[2]] = mask
    rpad = np.zeros(2 * esc); rpad[:es[0], :es[1], :es[2]] = np.where(mask, rho_dense, 0.0)
    cmask = mpad.reshape(esc[0], 2, esc[1], 2, esc[2], 2).any((1, 3, 5))
    gc = ActiveGrid(cmask)
    # child rho per coarse element in CORNERS order
    r6 = rpad.reshape(esc[0], 2, esc[1], 2, esc[2], 2)
    rho_c = np.stack([r6[:, c[0], :, c[1], :, c[2]] for c in C], -1).reshape(-1, 8)[gc.elem_dense]
    # coarse fixed DOF: any fine DOF within the 3x3x3 support fixed
    nnf = tuple(es + 1)
    fx = np.asarray(fixed_dense, bool).reshape(nnf + (3,))
    nnc = tuple(esc + 1)
    fpad = np.zeros(tuple(2 * np.array(nnc) - 1) + (3,), bool)
    fpad[:nnf[0], :nnf[1], :nnf[2]] = fx
    fxc = np.stack([ndimage.maximum_filter(fpad[..., k], size=3, mode="constant")[::2, ::2, ::2] for k in range(3)],
                   -1)
    fec = CoarseFE(gc, fxc, np.zeros(nnc + (3,)), 2 * h, E, nu, backend="gpu",
                   precision=coarse_prec)
    fec.set_children(child_P().transpose(0, 2, 1) @ Khat_fine[None] @ child_P())      # P_cᵀ K̂_h P_c (unit E, fine K̂)
    fec.Khat = None
    fec._diag_all = cp.asarray(np.flatnonzero(gc.row_of == gc.indices))
    # prolongation P (fine active DOF x coarse active DOF)
    nf = np.stack(np.unravel_index(g.node_dense, nnf), 1)                  # (Nf, 3) fine node ijk
    cidx = -np.ones(int(np.prod(nnc)), np.int64); cidx[gc.node_dense] = np.arange(gc.Nn)
    rows, cols, vals = [], [], []
    for corner in C:                                                        # up to 8 coarse parents per fine node
        odd = nf % 2
        take = np.all((corner == 0) | (odd == 1), 1)                       # use the +1 neighbour only on odd axes
        cn = nf[take] // 2 + corner * odd[take]
        w = np.prod(np.where(odd[take] == 1, 0.5, 1.0), 1)
        cd = cidx[(cn[:, 0] * nnc[1] + cn[:, 1]) * nnc[2] + cn[:, 2]]
        assert np.all(cd >= 0)
        rows.append(np.flatnonzero(take)); cols.append(cd); vals.append(w)
    Pn = sps.csr_matrix((np.concatenate(vals), (np.concatenate(rows), np.concatenate(cols))), shape=(g.Nn, gc.Nn))
    P = sps.kron(Pn, sps.identity(3), format="csr")
    Pd = cpsp.csr_matrix(P)
    PTd = cpsp.csr_matrix(P.T.tocsr())
    t_setup = time.time() - t0
    return fec, gc, rho_c, Pd, PTd, t_setup


class TwoGrid:
    def __init__(self, Kop, diag, fixed_mask, fec, P, PT, deg, lmax):
        self.K, self.Dinv, self.fx, self.fec, self.P, self.PT, self.deg = Kop, 1.0 / diag, fixed_mask, fec, P, PT, deg
        self.lmax, self.lmin = 1.1 * lmax, 0.1 * lmax                      # PETSc-style Chebyshev interval
        self.fxc = cp.asarray(fec.fixed)
        self.n_vcycles = 0

    def cheb(self, b):
        th, de = (self.lmax + self.lmin) / 2, (self.lmax - self.lmin) / 2
        sig = th / de; ro = 1 / sig
        Dinv = self.Dinv[:, None] if b.ndim == 2 else self.Dinv
        r = b.copy()
        d = Dinv * r / th
        x = d.copy()
        for _ in range(self.deg - 1):
            r -= self.K(d)
            rn = 1 / (2 * sig - ro)
            d = rn * ro * d + (2 * rn / de) * (Dinv * r)
            x += d
            ro = rn
        return x

    def coarse(self, r):
        rc = self.PT @ r
        if rc.ndim == 1:
            rc = rc[:, None]
        rc[self.fxc] = 0
        xc = cp.empty_like(rc)
        for j in range(rc.shape[1]):
            xc[:, j] = self.fec._direct(rc[:, j])
        xc[self.fxc] = 0
        return self.P @ xc

    def __call__(self, b):
        """Symmetric V(ν,ν) two-grid preconditioner, zero initial guess; b (n,) or (n,k)."""
        self.n_vcycles += 1 if b.ndim == 1 else b.shape[1]
        one = b.ndim == 1
        B = b[:, None] if one else b
        x = self.cheb(B)
        x += self.coarse(B - self.K(x))
        x += self.cheb(B - self.K(x))
        x[self.fx] = B[self.fx]                                            # identity rows of fixed DOF
        return x[:, 0] if one else x


def pcg(Kop, M, B, tol=1e-10, maxit=500):
    """Independent CG per column (batched SpMM); returns X and iterations."""
    one = B.ndim == 1
    B = B[:, None] if one else B
    nb = cp.linalg.norm(B, axis=0)
    X = cp.zeros_like(B)
    R = B.copy()
    Z = M(R)
    Pp = Z.copy()
    rz = cp.einsum("ij,ij->j", R, Z)
    for k in range(1, maxit + 1):
        AP = Kop(Pp)
        a = rz / cp.einsum("ij,ij->j", Pp, AP)
        X += a * Pp
        R -= a * AP
        rel = cp.linalg.norm(R, axis=0) / cp.maximum(nb, 1e-300)
        if bool(cp.all(rel <= tol)):
            break
        Z = M(R)
        rzn = cp.einsum("ij,ij->j", R, Z)
        Pp = Z + (rzn / rz) * Pp
        rz = rzn
    return (X[:, 0] if one else X), k, float(rel.max())


def matfree_op(fe, rho, chunk=1 << 20):
    """y = Σ_e E ρ_e K̂ u_e (q1_im, condensed K̂ identical for every voxel), identity on fixed DOF."""
    Kh = cp.asarray(fe.Khat)
    Er = fe.E * rho
    dofs = fe._dofs
    fx = cp.asarray(fe.fixed)

    def op(u):
        one = u.ndim == 1
        U = u[:, None] if one else u
        U0 = U.copy(); U0[fx] = 0
        Y = cp.zeros_like(U)
        for s in range(0, dofs.shape[0], chunk):
            d = dofs[s:s + chunk]
            ue = U0[d]                                                     # (c, 24, k)
            ye = cp.einsum("ij,ejk->eik", Kh, ue) * Er[s:s + chunk, None, None]
            cupyx.scatter_add(Y, d, ye)
        Y[fx] = U[fx]
        return Y[:, 0] if one else Y
    return op


class LeanGrid:
    """Connectivity only (no CSR pattern): enough for the matrix-free operator and the two-grid set-up at 0.2 mm."""

    def __init__(self, elem_mask):
        elem_mask = np.asarray(elem_mask, bool)
        self.elem_shape = tuple(int(n) for n in elem_mask.shape)
        self.node_shape = tuple(n + 1 for n in self.elem_shape)
        nyn, nzn = self.node_shape[1:]
        self.elem_dense = np.flatnonzero(elem_mask.ravel())
        eijk = np.stack(np.unravel_index(self.elem_dense, self.elem_shape), 1)
        nijk = eijk[:, None, :] + C[None]
        ndense = (nijk[..., 0] * nyn + nijk[..., 1]) * nzn + nijk[..., 2]
        del eijk, nijk
        self.node_dense, conn = np.unique(ndense.ravel(), return_inverse=True)
        self.conn = conn.reshape(-1, 8)
        self.Ne, self.Nn = self.conn.shape[0], self.node_dense.size
        self.ndof = 3 * self.Nn
        self.nnz = None

    def gather_nodes(self, a):
        a = np.asarray(a)
        return a.reshape((-1,) + a.shape[3:])[self.node_dense]

    def gather_elems(self, a):
        return np.asarray(a).ravel()[self.elem_dense]


class MFE:
    """Matrix-free fine level: what matfree_op and the two-grid need from SparseFE."""

    def __init__(self, g, fixed, load, h, E, nu):
        import gpu_analysis
        self.g, self.E, self.h = g, float(E), float(h)
        self.Khat = gpu_analysis.element_khat("q1_im", (h,) * 3, nu)
        self.fixed = g.gather_nodes(np.asarray(fixed, bool)).reshape(-1)
        f = g.gather_nodes(np.asarray(load, float)).reshape(-1)
        f[self.fixed] = 0
        self.f = cp.asarray(f)
        p = np.arange(3)
        self._dofs = cp.asarray((3 * g.conn[:, :, None] + p[None, None, :]).reshape(g.Ne, 24).astype(np.int64))

    def diag(self, rho):
        d = cp.zeros(self.g.ndof)
        cupyx.scatter_add(d, self._dofs, (self.E * rho)[:, None] * cp.asarray(np.diag(self.Khat))[None])
        d[cp.asarray(self.fixed)] = 1.0
        return d


def main_matfree(a):
    """0.2 mm route: matrix-free fine operator + Galerkin coarse (fp32 cuDSS) two-grid PCG, compliance only."""
    base = gm.dev_used()
    t0 = time.time()
    cfg, g, fixed, load, hfe, rho, Fn = gm.setup(a.part, a.h, a.hollow, grid_cls=LeanGrid)
    E, nu = cfg["material"]["E_MPa"], cfg["material"]["nu"]
    res = dict(part=cfg["name"], h_e=hfe, hollow=a.hollow, Ne=g.Ne, ndof=g.ndof, setup_s=time.time() - t0, deg=a.deg,
               coarse_prec=a.coarse, mode="matfree_only")
    print(json.dumps(res), flush=True)
    fe = MFE(g, fixed, load, hfe, E, nu)
    rho_act = g.gather_elems(rho)
    fe.rho_hat = cp.asarray(rho_act)
    rho_dense = np.zeros(int(np.prod(g.elem_shape))); rho_dense[g.elem_dense] = rho_act
    del rho
    fec, gc, rho_c, P, PT, t_tg = build_twogrid(g, fixed, rho_dense.reshape(g.elem_shape), hfe, E, nu, a.coarse,
                                                fe.Khat)
    del rho_dense
    res.update(coarse_ndof=gc.ndof, coarse_nnz=gc.nnz, twogrid_host_setup_s=t_tg)
    t1 = time.time(); fec.factor(rho_c); cp.cuda.Device().synchronize()
    res.update(coarse_factor_s=time.time() - t1, coarse_plan_s=fec.stats.get("plan_s"))
    gm.POOL.free_all_blocks()
    res["mem_after_coarse_factor"] = gm.mem_snapshot("coarse_factor", base)
    Kmf = matfree_op(fe, fe.rho_hat)
    x = cp.random.default_rng(1).standard_normal(g.ndof)
    Kmf(x); cp.cuda.Device().synchronize(); t1 = time.time()
    for _ in range(3):
        Kmf(x)
    cp.cuda.Device().synchronize(); res["matvec_matfree_1col_s"] = (time.time() - t1) / 3
    diag = fe.diag(fe.rho_hat)
    v = x
    for _ in range(20):
        v = Kmf(v) / diag; lam = float(cp.linalg.norm(v)); v /= lam
    del v, x
    res["lmax_DinvK"] = lam
    fmask = cp.asarray(fe.fixed)
    M = TwoGrid(Kmf, diag, fmask, fec, P, PT, a.deg, lam)
    cp.cuda.Device().synchronize(); t1 = time.time()
    u, its, rr = pcg(Kmf, M, fe.f, tol=1e-10)
    cp.cuda.Device().synchronize()
    res["pcg_matfree"] = dict(its=its, rel_res=rr, seconds=time.time() - t1, compliance=float(fe.f @ u))
    res["peak_cupy_GiB"] = gm.PEAK[0] / GiB
    res["mem_after_pcg"] = gm.mem_snapshot("after_pcg", base)
    res["total_s"] = time.time() - t0
    print(json.dumps(res), flush=True)
    Path(a.out).write_text(json.dumps(res, indent=1, default=float))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("part"); ap.add_argument("out")
    ap.add_argument("--h", type=float, default=0.4)
    ap.add_argument("--hollow", type=int, default=0)
    ap.add_argument("--deg", type=int, default=3)
    ap.add_argument("--coarse", default="fp32")
    ap.add_argument("--buckling", action="store_true")
    ap.add_argument("--ref", default=None, help="gpumem json with fp64 compliance / mu for comparison")
    ap.add_argument("--inner-tol", type=float, default=1e-4, help="PCG tol for Krylov expansion solves")
    ap.add_argument("--matfree-only", action="store_true")
    a = ap.parse_args()
    if a.matfree_only:
        return main_matfree(a)
    base = gm.dev_used()
    t0 = time.time()
    cfg, g, fixed, load, hfe, rho, Fn = gm.setup(a.part, a.h, a.hollow)
    E, nu = cfg["material"]["E_MPa"], cfg["material"]["nu"]
    res = dict(part=cfg["name"], h_e=hfe, hollow=a.hollow, ndof=g.ndof, nnz=g.nnz, setup_s=time.time() - t0, deg=a.deg,
               coarse_prec=a.coarse)
    fe = SparseFE(g, fixed, load, hfe, E, nu, backend="gpu", precision="fp64")
    rho_act = g.gather_elems(rho)
    fe.rho_hat = cp.asarray(rho_act)
    data = fe._values(fe.rho_hat)
    fe._K = cpsp.csr_matrix((data, cp.asarray(g.indices.astype(np.int32)), cp.asarray(g.indptr.astype(np.int32))),
                                 shape=(g.ndof, g.ndof))
    del data
    gm.POOL.free_all_blocks()
    res["mem_fine_K"] = gm.mem_snapshot("fine_K", base)
    # coarse level
    rho_dense = np.zeros(int(np.prod(g.elem_shape))); rho_dense[g.elem_dense] = rho_act
    fec, gc, rho_c, P, PT, t_tg = build_twogrid(g, fixed, rho_dense.reshape(g.elem_shape), hfe, E, nu, a.coarse, fe.Khat)
    res.update(coarse_ndof=gc.ndof, coarse_nnz=gc.nnz, twogrid_host_setup_s=t_tg)
    t1 = time.time(); fec.factor(rho_c); cp.cuda.Device().synchronize()
    res.update(coarse_factor_s=time.time() - t1, coarse_zero_diag=fec.stats.get("coarse_zero_diag"))
    gm.POOL.free_all_blocks()
    res["mem_after_coarse_factor"] = gm.mem_snapshot("coarse_factor", base)
    # Galerkin check: Pᵀ K P x = K_c x on free coarse DOF
    xc = cp.random.default_rng(0).standard_normal(gc.ndof); fxc = cp.asarray(fec.fixed); xc[fxc] = 0
    lhs = PT @ (fe._K @ (P @ xc)); rhs = fec._K @ xc
    res["galerkin_rel_err"] = float(cp.linalg.norm((lhs - rhs)[~fxc]) / cp.linalg.norm(rhs[~fxc]))
    del lhs, rhs, xc
    Kcsr = lambda x: fe._K @ x  # noqa: E731
    Kmf = matfree_op(fe, fe.rho_hat)
    # matrix-free vs CSR: equality and time
    x = cp.random.default_rng(1).standard_normal(g.ndof)
    res["matfree_rel_diff"] = float(cp.linalg.norm(Kmf(x) - Kcsr(x)) / cp.linalg.norm(Kcsr(x)))
    for name, op in (("csr", Kcsr), ("matfree", Kmf)):
        for k in (1, 10):
            X = cp.random.default_rng(2).standard_normal((g.ndof, k)) if k > 1 else x
            op(X); cp.cuda.Device().synchronize()
            t1 = time.time()
            for _ in range(5):
                op(X)
            cp.cuda.Device().synchronize()
            res[f"matvec_{name}_{k}col_s"] = (time.time() - t1) / 5
    del x, X
    diag = fe._K.diagonal()
    # λmax(D⁻¹K) by power iteration
    v = cp.random.default_rng(3).standard_normal(g.ndof)
    for _ in range(20):
        v = (Kcsr(v) / diag); lam = float(cp.linalg.norm(v)); v /= lam
    res["lmax_DinvK"] = lam
    fmask = cp.asarray(fe.fixed)
    f = fe.f.copy(); f[fmask] = 0
    for opname, op in (("csr", Kcsr), ("matfree", Kmf)):
        M = TwoGrid(op, diag, fmask, fec, P, PT, a.deg, lam)
        cp.cuda.Device().synchronize(); t1 = time.time()
        u, its, rr = pcg(op, M, f, tol=1e-10)
        cp.cuda.Device().synchronize()
        res[f"pcg_{opname}"] = dict(its=its, rel_res=rr, seconds=time.time() - t1, compliance=float(f @ u),
                                    vcycles=M.n_vcycles)
        print(opname, json.dumps(res[f"pcg_{opname}"]), flush=True)
    res["peak_cupy_GiB"] = gm.PEAK[0] / GiB
    res["mem_after_pcg"] = gm.mem_snapshot("after_pcg", base)
    if a.ref:
        ref = json.loads(Path(a.ref).read_text())
        Cref = ref["fp64"]["compliance"]
        res["compliance_rel_vs_direct"] = abs(res["pcg_csr"]["compliance"] - Cref) / Cref
    Path(a.out).write_text(json.dumps(res, indent=1, default=float))
    if a.buckling:
        M = TwoGrid(Kcsr, diag, fmask, fec, P, PT, a.deg, lam)
        stats = dict(its=0, calls=0)

        def solve(rhs):
            b = cp.array(rhs, dtype=cp.float64)
            b[fmask] = 0
            X, its, _ = pcg(Kcsr, M, b, tol=1e-10)
            stats["its"] += its; stats["calls"] += 1
            return X
        fe.solve = solve
        t1 = time.time(); p = fe.primal()
        def inner(B):
            b = cp.array(B, dtype=cp.float64); b[fmask] = 0
            X, its, _ = pcg(Kcsr, M, b, tol=a.inner_tol)
            stats["inner_its"] = stats.get("inner_its", 0) + its
            return X
        o, b = gm.run_buckling(fe, p, lean=True, tol=1e-3, inner=inner)
        o.update(total_s=time.time() - t1, exact_pcg_its=stats["its"], exact_calls=stats["calls"],
                 krylov_pcg_its=stats.get("inner_its", 0), inner_tol=a.inner_tol,
                 vcycles=M.n_vcycles, lambda1=float(1 / max(b.mu)) if len(b.mu) else None)
        if a.ref and ref["fp64"]["buckling"].get("lean", {}).get("mu"):
            mr = ref["fp64"]["buckling"]["lean"]["mu"]
            o["mu_rel_vs_direct"] = gm.rel(b.mu[:len(mr)], mr)
        adir = os.environ.get("GM_ARR_DIR")
        if adir:
            Path(adir).mkdir(parents=True, exist_ok=True)
            np.save(Path(adir) / f"{Path(a.out).stem}_dmu.npy", b.dmu_tilde)
            np.save(Path(adir) / f"{Path(a.out).stem}_dC.npy", fe.dC_drho_hat(p))
        res["buckling_twogrid"] = o
        print("buckling", json.dumps(o), flush=True)
    res["total_s"] = time.time() - t0
    Path(a.out).write_text(json.dumps(res, indent=1, default=float))


if __name__ == "__main__":
    main()
