"""GPU body-fitted check (cshell.validate engine "gpu"): quadratic tets (C3D10 node order, straight edges), cuDSS Cholesky, linear buckling by block
Krylov on K^-1 G with G = -K_sigma. Same discrete model as CalculiX C3D10 (4-point rule for K, sigma and K_sigma).
Units mm, N, MPa.

    (K + lambda K_sigma) phi = 0   <=>   G phi = mu K phi,  G = -K_sigma,  mu = 1/lambda  (largest mu first)

API: solve(nodes, tets10, fixed, load_nodes, load_vec, E, nu, n_modes, exclude_bc_mm) -> dict(C, lam, lam_unexcl,
     sigma_q, timings); dihedral_min / layers: mesh-quality measures for the validation gate.
K_sigma exclusion: elements with centroid within exclude_bc_mm of a clamped node carry no geometric stiffness in `lam`
(`lam_unexcl` is the plain factor); sigma_q excludes the same distance from clamped and loaded nodes.
CLI: python bf_gpu.py deck.inp [n_modes]   (re-solves a cshell.validate CalculiX deck with identical mesh and BCs)
"""
from __future__ import annotations

import json
import re
import sys
import time

import numpy as np

A4, B4 = 0.5854101966249685, 0.1381966011250105
GP = np.array([[A4, B4, B4, B4], [B4, A4, B4, B4], [B4, B4, A4, B4], [B4, B4, B4, A4]])   # barycentric (L1..L4)
EDGES = np.array([[0, 1], [1, 2], [2, 0], [0, 3], [1, 3], [2, 3]])                            # C3D10 mid-edge 5..10


def read_deck(path):
    txt = open(path).read()

    def block(start):
        i = txt.index(start); i = txt.index("\n", i) + 1; j = txt.index("*", i)
        return txt[i:j]
    nodes = np.array(block("*NODE, NSET=NALL").replace(",", " ").split(), float).reshape(-1, 4)[:, 1:]
    el = np.array(block("*ELEMENT").replace(",", " ").split(), np.int64).reshape(-1, 11)[:, 1:] - 1
    fixed = np.array(block("*NSET, NSET=NFIX").replace(",", " ").split(), np.int64) - 1
    cl = np.array(txt[txt.index("*CLOAD") + 7: txt.index("*NODE PRINT")].replace(",", " ").split(), float).reshape(-1, 3)
    ln = np.unique(cl[:, 0].astype(np.int64))
    lv = np.zeros((len(ln), 3))
    pos = np.searchsorted(ln, cl[:, 0].astype(np.int64))
    np.add.at(lv, (pos, cl[:, 1].astype(int) - 1), cl[:, 2])
    E, nu = [float(x) for x in txt[txt.index("*ELASTIC") + 9:].split("\n")[0].split(",")[:2]]
    return nodes, el, fixed, ln - 1, lv, E, nu


def _grads(cp, X, T):
    """Constant corner-gradient matrix per tet and volume: gL (Ne, 4, 3), vol (Ne,)."""
    p = X[T[:, :4]]
    J = cp.stack([p[:, 1] - p[:, 0], p[:, 2] - p[:, 0], p[:, 3] - p[:, 0]], 1)        # rows = edge vectors
    vol = cp.linalg.det(J) / 6.0
    Jinv = cp.linalg.inv(J)                                                             # columns = grad L2..L4
    g234 = cp.transpose(Jinv, (0, 2, 1))
    gL = cp.concatenate([-g234.sum(1, keepdims=True), g234], 1)
    return gL, vol


def _shape_grads(cp, gL):
    """grad N (Ne, 4 gp, 10, 3) of the P2 basis at the 4 Gauss points."""
    out = []
    for g in range(4):
        L = GP[g]
        corner = cp.asarray(4 * L - 1)[None, :, None] * gL
        mids = [4 * (float(L[b]) * gL[:, a] + float(L[a]) * gL[:, b]) for a, b in EDGES]
        out.append(cp.concatenate([corner, cp.stack(mids, 1)], 1))
    return cp.stack(out, 1)


def _Bmat(cp, dN):
    """Voigt B (Ne, 4, 6, 30) in ccx order (xx, yy, zz, xy, xz, yz), engineering shear."""
    Ne = dN.shape[0]
    B = cp.zeros((Ne, 4, 6, 30))
    for a in range(10):
        dx, dy, dz = dN[:, :, a, 0], dN[:, :, a, 1], dN[:, :, a, 2]
        c = 3 * a
        B[:, :, 0, c] = dx; B[:, :, 1, c + 1] = dy; B[:, :, 2, c + 2] = dz
        B[:, :, 3, c] = dy; B[:, :, 3, c + 1] = dx
        B[:, :, 4, c] = dz; B[:, :, 4, c + 2] = dx
        B[:, :, 5, c + 1] = dz; B[:, :, 5, c + 2] = dy
    return B


def _D(E, nu):
    lam, mu = E * nu / ((1 + nu) * (1 - nu * 2)), E / (2 * (1 + nu))
    D = np.zeros((6, 6))
    D[:3, :3] = lam
    D[np.arange(3), np.arange(3)] += 2 * mu
    D[3:, 3:] = np.eye(3) * mu
    return D


class P2Model:
    def __init__(self, nodes, tets10, fixed, E, nu, batch=200000):
        import cupy as cp
        import cupyx.scipy.sparse as csp
        self.cp, self.csp = cp, csp
        t0 = time.time()
        self.N, self.Ne = len(nodes), len(tets10)
        self.X = cp.asarray(nodes, float)
        self.T = cp.asarray(tets10, np.int64)
        self.D = cp.asarray(_D(E, nu))
        fixed_dof = (3 * np.asarray(fixed)[:, None] + np.arange(3)).ravel()
        free = np.ones(3 * self.N, bool); free[fixed_dof] = False
        rmap = -np.ones(3 * self.N, np.int64); rmap[free] = np.arange(free.sum())
        self.free, self.nf = cp.asarray(free), int(free.sum())
        self.rmap = cp.asarray(rmap)
        self.batch = batch
        # sparsity pattern (reduced DOF), position of every element entry in CSR data
        rows, cols = [], []
        for s in range(0, self.Ne, batch):
            dofs = self._dofs(slice(s, s + batch))
            r = cp.broadcast_to(dofs[:, :, None], (dofs.shape[0], 30, 30)).ravel()
            c = cp.broadcast_to(dofs[:, None, :], (dofs.shape[0], 30, 30)).ravel()
            rows.append(r); cols.append(c)
        r, c = cp.concatenate(rows), cp.concatenate(cols)
        del rows, cols
        keep = (r >= 0) & (c >= 0)
        key = r * self.nf + c
        key = cp.where(keep, key, -1)
        ukey, inv = cp.unique(key, return_inverse=True)
        off = 1 if int(ukey[0]) < 0 else 0
        self.pos = (inv.ravel() - off).astype(cp.int64)          # −1 for dropped entries
        ukey = ukey[off:]
        ur, uc = ukey // self.nf, ukey % self.nf
        self.indptr = cp.concatenate([cp.zeros(1, cp.int64), cp.cumsum(cp.bincount(ur, minlength=self.nf))]).astype(cp.int32)
        self.indices = uc.astype(cp.int32)
        self.nnz = int(ukey.size)
        del r, c, key, ukey, inv, ur, uc, keep
        self.t_pattern = time.time() - t0

    def _dofs(self, sl):
        T = self.T[sl]
        d = (3 * T[:, :, None] + self.cp.arange(3)).reshape(-1, 30)
        return self.rmap[d]

    def _assemble(self, elem_fn):
        cp = self.cp
        data = cp.zeros(self.nnz)
        for s in range(0, self.Ne, self.batch):
            ke = elem_fn(slice(s, s + self.batch)).ravel()
            p = self.pos[s * 900:(s * 900 + ke.size)]
            m = p >= 0
            data += cp.bincount(p[m], weights=ke[m], minlength=self.nnz)
        return self.csp.csr_matrix((data, self.indices, self.indptr), shape=(self.nf, self.nf))

    def _geom(self, sl):
        gL, vol = _grads(self.cp, self.X, self.T[sl][:, :4])
        dN = _shape_grads(self.cp, gL)
        return dN, vol

    def stiffness(self):
        cp = self.cp

        def fn(sl):
            dN, vol = self._geom(sl)
            B = _Bmat(cp, dN)
            return cp.einsum("e,egia,ij,egjb->eab", vol / 4.0, B, self.D, B, optimize=True)
        t0 = time.time()
        self.K = self._assemble(fn)
        self.t_K = time.time() - t0
        return self.K

    def stress(self, u_full):
        """Gauss-point stress (Ne, 4, 6) and volumes."""
        cp = self.cp
        S, V = [], []
        for s in range(0, self.Ne, self.batch):
            sl = slice(s, s + self.batch)
            dN, vol = self._geom(sl)
            B = _Bmat(cp, dN)
            ue = u_full[(3 * self.T[sl][:, :, None] + cp.arange(3)).reshape(-1, 30)]
            S.append(cp.einsum("ij,egjb,eb->egi", self.D, B, ue)); V.append(vol)
        return cp.concatenate(S), cp.concatenate(V)

    def geometric(self, sig):
        cp = self.cp

        def fn(sl):
            dN, vol = self._geom(sl)
            s = sig[sl]
            St = cp.stack([cp.stack([s[..., 0], s[..., 3], s[..., 4]], -1),
                           cp.stack([s[..., 3], s[..., 1], s[..., 5]], -1),
                           cp.stack([s[..., 4], s[..., 5], s[..., 2]], -1)], -2)           # (ne, 4, 3, 3)
            H = cp.einsum("e,egai,egij,egbj->eab", vol / 4.0, dN, St, dN, optimize=True)  # (ne, 10, 10)
            ke = cp.zeros((H.shape[0], 10, 3, 10, 3))
            for d in range(3):
                ke[:, :, d, :, d] = H
            return ke.reshape(-1, 30, 30)
        t0 = time.time()
        Ks = self._assemble(fn)
        self.t_Ks = time.time() - t0
        return Ks


class Cholesky:
    def __init__(self, K, nrhs=1):
        import cupy as cp
        from nvmath.sparse.advanced import DirectSolver, DirectSolverMatrixType, DirectSolverOptions, ExecutionCUDA
        t0 = time.time()
        self.cp = cp
        self.b = cp.asfortranarray(cp.zeros((K.shape[0], nrhs)))
        opts = DirectSolverOptions(sparse_system_type=DirectSolverMatrixType.SPD)
        self.s = DirectSolver(K, self.b, options=opts, execution=ExecutionCUDA())
        self.s.plan()
        est = self.s.plan_info.memory_estimates
        self.est_GiB = float(est["peak_device_memory"]) / 2 ** 30
        # Production parts (research/bodyfit_scale): if the factor does not fit next to the eigen-solver block vectors
        # (~100 vectors of ndof), keep the factor in host memory (cuDSS hybrid mode, same arithmetic) rather than a
        # second GPU. Logged through self.hybrid / self.est_host_GiB.
        cp.get_default_memory_pool().free_all_blocks()        # cached pool blocks are free for cuDSS too
        free = cp.cuda.Device().mem_info[0]
        reserve = 100 * 8 * K.shape[0] + (4 << 30)
        self.hybrid, self.est_host_GiB = False, 0.0
        if float(est["peak_device_memory"]) > free - reserve:
            self.s.free()
            lim = int(max(free - reserve, float(est["hybrid_min_device_memory"]) * 1.05))
            self.s = DirectSolver(K, self.b, options=opts, execution=ExecutionCUDA(hybrid_memory_mode_options=dict(
                hybrid_memory_mode=True, hybrid_device_memory_limit=lim)))
            self.s.plan()
            self.hybrid = True
            self.est_host_GiB = float(self.s.plan_info.memory_estimates["peak_host_memory"]) / 2 ** 30
            print(f"[bodyfit_gpu] factor {self.est_GiB:.1f} GiB > free {free / 2 ** 30:.1f} GiB - reserve "
                  f"{reserve / 2 ** 30:.1f} GiB: cuDSS hybrid memory mode, device limit {lim / 2 ** 30:.1f} GiB, "
                  f"host {self.est_host_GiB:.1f} GiB", flush=True)
        self.t_plan = time.time() - t0
        t0 = time.time()
        self.s.factorize()
        cp.cuda.Device().synchronize()
        self.t_factor = time.time() - t0

    def solve(self, R):
        cp = self.cp
        one = R.ndim == 1
        R = R[:, None] if one else R
        out = cp.empty_like(R)
        k = self.b.shape[1]
        for j0 in range(0, R.shape[1], k):
            j1 = min(j0 + k, R.shape[1])
            self.b[:, :] = 0.0
            self.b[:, : j1 - j0] = R[:, j0:j1]
            out[:, j0:j1] = cp.asarray(self.s.solve())[:, : j1 - j0]
        return out[:, 0] if one else out


def buckling(cp, K, G, chol, n_modes=4, m=None, tol=1e-7, maxit=200, seed=0):
    """Largest mu of G phi = mu K phi: block Krylov, basis [X, K^-1 G X], K-orthonormal Rayleigh-Ritz."""
    m = m or 2 * n_modes + 4
    rng = np.random.default_rng(seed)
    X = cp.asarray(rng.standard_normal((K.shape[0], m)))
    X = chol.solve(G @ chol.solve(X))
    hist = []
    for it in range(maxit):
        Y = chol.solve(G @ X)
        S = cp.concatenate([X, Y], 1)
        KS = K @ S
        M = S.T @ KS; M = 0.5 * (M + M.T)
        w, Q = cp.linalg.eigh(M)
        keep = w > w.max() * 1e-12
        Q = Q[:, keep] / cp.sqrt(w[keep])[None]
        S, KS = S @ Q, KS @ Q
        A = S.T @ (G @ S); A = 0.5 * (A + A.T)
        mu, Z = cp.linalg.eigh(A)
        order = cp.argsort(-mu)[:m]
        mu, Z = mu[order], Z[:, order]
        X = S @ Z
        KX = KS @ Z
        R = G @ X[:, :n_modes] - KX[:, :n_modes] * mu[None, :n_modes]
        res = cp.linalg.norm(R, axis=0) / (cp.abs(mu[:n_modes]) * cp.linalg.norm(KX[:, :n_modes], axis=0))
        hist.append(float(res.max()))
        if hist[-1] < tol:
            break
    return cp.asnumpy(mu), X, hist


def weighted_quantile(x, w, q):
    i = np.argsort(x)
    c = np.cumsum(w[i]); c /= c[-1]
    return float(x[i][np.searchsorted(c, q)])


class ScaleLimit(RuntimeError):
    """The body-fitted model exceeds the solver's capacity (`MAX_DOF`); raised before assembly."""


# Capacity of the one-GPU direct solve (set by collect from `validation.max_bodyfit_dof`; None = no limit). B2 at
# h/√2 (3.2 M DOF) needs a 51 GiB fp64 factor; production parts (19–42 M DOF) need 0.15–1.5 TB and the GPU stiffness
# assembly runs out of device memory at 3.5–4.5 M tets (research/bodyfit_scale/notes.md).
MAX_DOF = None


def solve(*a, **k):
    """_solve, then release the cuDSS factor and the CuPy pool (two B2 levels in one process ran out of device memory)."""
    if MAX_DOF and 3 * len(a[0]) > MAX_DOF:
        raise ScaleLimit(f"{3 * len(a[0])} DOF > validation.max_bodyfit_dof = {MAX_DOF:.3g}")
    import gc
    import cupy as cp
    try:
        return _solve(*a, **k)
    finally:
        gc.collect()
        cp.get_default_memory_pool().free_all_blocks()


def _solve(nodes, tets10, fixed, load_nodes, load_vec, E, nu, n_modes=4, quantile=0.995, exclude_bc_mm=2.0, nrhs=12,
           tol=1e-7, unexcl=True):
    import cupy as cp
    t = {}
    t0 = time.time()
    M = P2Model(nodes, tets10, fixed, E, nu)
    t["pattern"] = M.t_pattern
    K = M.stiffness(); t["assemble_K"] = M.t_K
    ch = Cholesky(K, nrhs); t["plan"] = ch.t_plan; t["factor"] = ch.t_factor
    f = np.zeros(3 * len(nodes))
    f[(3 * np.asarray(load_nodes)[:, None] + np.arange(3)).ravel()] = np.asarray(load_vec).ravel()
    fr = cp.asarray(f)[M.free]
    t1 = time.time()
    u = ch.solve(fr)
    t["solve"] = time.time() - t1
    C = float(fr @ u)
    uf = cp.zeros(3 * len(nodes)); uf[M.free] = u
    res_static = float(cp.linalg.norm(K @ u - fr) / cp.linalg.norm(fr))
    t1 = time.time()
    sig, vol = M.stress(uf)
    G = M.geometric(sig); G.data *= -1.0
    t["assemble_G"] = time.time() - t1
    from scipy.spatial import cKDTree
    cen = nodes[tets10[:, :4]].mean(1)
    t1 = time.time()
    if exclude_bc_mm and len(fixed):
        near = cKDTree(nodes[np.asarray(fixed)]).query(cen)[0] <= exclude_bc_mm
        sig_x = sig * cp.asarray(~near)[:, None, None]
        Gx = M.geometric(sig_x); Gx.data *= -1.0
    else:
        near, Gx = np.zeros(len(tets10), bool), G
    t["assemble_G"] += time.time() - t1
    t1 = time.time()
    mu, Xm, hist = buckling(cp, K, Gx, ch, n_modes, tol=tol)
    lam = (1.0 / mu[mu > 0])[:n_modes]
    if Gx is not G and unexcl:
        mu0, _, _ = buckling(cp, K, G, ch, n_modes, tol=tol)
        lam0 = (1.0 / mu0[mu0 > 0])[:n_modes]
    else:
        lam0 = lam
    t["eigen"] = time.time() - t1
    # stress quantile as cshell.validate.responses_from_ccx (exclusion 2 mm from BC nodes, V_e/4 weights)
    s = cp.asnumpy(sig); v = cp.asnumpy(vol)
    vm = np.sqrt(0.5 * ((s[..., 0] - s[..., 1]) ** 2 + (s[..., 1] - s[..., 2]) ** 2 + (s[..., 2] - s[..., 0]) ** 2)
                 + 3 * (s[..., 3] ** 2 + s[..., 4] ** 2 + s[..., 5] ** 2))
    bc = np.union1d(fixed, load_nodes)
    keepe = cKDTree(nodes[bc]).query(cen)[0] > exclude_bc_mm
    sq = weighted_quantile(vm.ravel(), np.repeat((v / 4 * keepe)[:, None], 4, 1).ravel(), quantile)
    # mode-1 location (nodal |phi|)
    ph = cp.zeros((3 * len(nodes), min(n_modes, Xm.shape[1]))); ph[M.free] = Xm[:, : ph.shape[1]]
    a = cp.asnumpy(cp.linalg.norm(ph.reshape(len(nodes), 3, -1), axis=1))
    loc = []
    for k in range(a.shape[1]):
        top = np.argsort(a[:, k])[-50:]
        w = a[:, k] ** 2
        loc.append(dict(x_top50=[float(nodes[top, 0].min()), float(nodes[top, 0].max())],
                        x_centroid_w=float((w * nodes[:, 0]).sum() / w.sum()),
                        frac_energy_top1pct=float(np.sort(w)[-len(w) // 100:].sum() / w.sum())))
    t["total"] = time.time() - t0
    out = dict(C=C, lam=lam.tolist(), lam_unexcl=np.asarray(lam0).tolist(), n_ks_excluded=int(near.sum()), sigma_q=sq, volume=float(v.sum()), neg_vol=int((v <= 0).sum()), ndof=M.nf,
               nnz=M.nnz, factor_GiB=ch.est_GiB, hybrid=ch.hybrid, factor_host_GiB=ch.est_host_GiB, static_rel_residual=res_static, eig_iters=len(hist),
               eig_res=hist[-1], timings_s=t, modes=loc)
    out["_phi_nodes"] = a
    try:
        ch.s.free()
    except Exception:
        pass
    return out


# ---- mesh-quality measures (CPU; numpy / libigl) ----
def dihedral_min(X, T):
    """Minimum dihedral angle (deg) per tet from outward face normals: theta_ij = pi - angle(n_i, n_j)."""
    p = X[T[:, :4]]
    n = []
    for i in range(4):
        a, b, c = [k for k in range(4) if k != i]
        v = np.cross(p[:, b] - p[:, a], p[:, c] - p[:, a])
        v /= np.linalg.norm(v, axis=1, keepdims=True)
        sgn = np.sign(np.einsum("ij,ij->i", v, p[:, a] - p[:, i]))   # point away from the opposite vertex
        n.append(v * sgn[:, None])
    dmin = np.full(len(T), 180.0)
    for i in range(4):
        for j in range(i + 1, 4):
            ang = 180 - np.degrees(np.arccos(np.clip(np.einsum("ij,ij->i", n[i], n[j]), -1, 1)))
            dmin = np.minimum(dmin, ang)
    return dmin


def layers(V, F, nodes, tets, n=20000, seed=0):
    """Tet layers through the wall: ray thickness t at surface samples (opposite wall only) / rms edge of the
    8 tets nearest the mid-wall point."""
    import igl
    from scipy.spatial import cKDTree
    tri = V[F]; nrm = np.cross(tri[:, 1] - tri[:, 0], tri[:, 2] - tri[:, 0]); A = .5 * np.linalg.norm(nrm, axis=1)
    nrm /= 2 * A[:, None]
    rng = np.random.default_rng(seed); pick = rng.choice(len(F), n, p=A / A.sum()); u = rng.random((n, 2))
    fl = u.sum(1) > 1; u[fl] = 1 - u[fl]
    P = tri[pick, 0] + u[:, :1] * (tri[pick, 1] - tri[pick, 0]) + u[:, 1:] * (tri[pick, 2] - tri[pick, 0]); N = nrm[pick]
    aabb = igl.AABB(); aabb.init(np.ascontiguousarray(V), np.ascontiguousarray(F))
    hit, t = aabb.intersect_ray_first(V, F, np.ascontiguousarray(P - 1e-4 * N), np.ascontiguousarray(-N))[:2]
    hit, t = np.asarray(hit), np.asarray(t) + 1e-4
    ok = hit >= 0
    cos = np.full(n, 0.0); cos[ok] = np.einsum("ij,ij->i", nrm[hit[ok]], N[ok])
    ok &= cos < -0.7
    T4 = tets[:, :4]
    L = np.stack([np.linalg.norm(nodes[T4[:, i]] - nodes[T4[:, j]], axis=1) for i, j in ((0, 1), (1, 2), (0, 2), (0, 3), (1, 3), (2, 3))], 1)
    Lr = np.sqrt((L ** 2).mean(1)); cen = nodes[T4].mean(1)
    mid = P[ok] - 0.5 * t[ok, None] * N[ok]
    jj = cKDTree(cen).query(mid, k=8)[1]
    nl = t[ok] / Lr[jj].mean(1)
    thin = t[ok] < 2.0
    return dict(samples=int(ok.sum()), layers_q=np.quantile(nl, [0, .01, .1, .5]).round(3).tolist(),
                frac_lt_2=float((nl < 2).mean()),
                layers_q_walls_lt_2mm=np.quantile(nl[thin], [0, .01, .5]).round(3).tolist() if thin.any() else [],
                wall_q=np.quantile(t[ok], [0, .01, .5]).round(4).tolist())


if __name__ == "__main__":
    deck = sys.argv[1]
    nm = int(sys.argv[2]) if len(sys.argv) > 2 else 4
    t0 = time.time()
    nodes, tets, fixed, ln, lv, E, nu = read_deck(deck)
    tr = time.time() - t0
    out = solve(nodes, tets, fixed, ln, lv, E, nu, nm)
    out.pop("_phi_nodes")
    out["read_s"] = tr
    out["deck"] = deck
    print(json.dumps(out, indent=1))
