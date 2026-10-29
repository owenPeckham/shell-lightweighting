"""Diagnose a non-finite two-grid set-up (MBB phase 0): diagonals, λmax, coarse diagonal, connectivity, probe solves."""
import glob
import sys
from pathlib import Path

import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parents[3]))


def main(run_dir):
    import cupy as cp
    import scipy.sparse as sps
    from scipy.sparse.csgraph import connected_components
    from scipy import ndimage
    from cshell import config
    from cshell.optimise import Context, PhaseModel
    from cshell.fe_twogrid import LeanGrid, TwoGridFE
    rd = Path(run_dir)
    cfg = config.load(str(rd / "config_used.json"), {"fe": {"engine": "twogrid"}})
    ctx = Context(cfg)
    z = np.load(sorted(glob.glob(str(rd / "checkpoints" / "*.npz")))[-1])
    model = PhaseModel(ctx, 0, float(z["D_comp"]))
    nv = ctx.surf.n_vertices
    xi = np.load(rd / "xi_uniform_phase0.npy")
    rho = np.asarray(model.dm.rho(model.dm.forward(xi[:nv], xi[nv:], as_built=True)))
    margin = int(np.ceil(cfg["grid"]["active_margin_mm"] / model.h_e))
    mask = ndimage.binary_dilation((rho > 0).reshape(model.elem_shape), iterations=margin)
    g = LeanGrid(mask, xp=cp)
    conn = g.conn.get()
    # element-node graph components
    A = sps.coo_matrix((np.ones(conn.size), (np.repeat(np.arange(g.Ne), 8), conn.ravel())), shape=(g.Ne, g.Nn)).tocsr()
    nc, lab = connected_components(A.T @ A, directed=False)
    fx = g.gather_nodes(model.fixed).reshape(-1, 3).any(1)
    clamped = set(np.unique(lab[fx]).tolist())
    sizes = np.bincount(lab)
    print("node components", nc, "sizes", sorted(sizes.tolist())[-5:], "unclamped components",
          [int(sizes[c]) for c in range(nc) if c not in clamped][:20], flush=True)
    fe = TwoGridFE(g, model.fixed, model.load, model.h_e, cfg["material"]["E_MPa"], cfg["material"]["nu"],
                   backend="gpu", coarse_precision="fp64")
    r = g.gather_elems(rho)
    rho_hat = model.disc.rho_min + (1 - model.disc.rho_min) * r
    print("rho_hat min/max", rho_hat.min(), rho_hat.max(), "finite", np.isfinite(rho_hat).all(), flush=True)
    data = None
    try:
        fe.factor(rho_hat)
    except Exception as e:
        print("factor raised", repr(e), flush=True)
    print("diag min", float(fe._diag.min()), "finite", bool(cp.isfinite(fe._diag).all()), "lmax",
          fe.stats.get("lmax_DinvK"), flush=True)
    data = fe._coarse_values()
    pat = fe._pat
    d = data[pat.diag_pos]
    print("coarse diag min", float(d.min()), "n<=0", int((d <= 0).sum()), "fixed_c", int(fe._fixed_c.sum()),
          "finite", bool(cp.isfinite(data).all()), flush=True)
    Kc = sps.csr_matrix((data.get(), pat.indices.get(), pat.indptr.get()), shape=(pat.ndof, pat.ndof))
    print("coarse symmetric err", abs(Kc - Kc.T).max(), flush=True)
    ncc, labc = connected_components(Kc, directed=False)
    print("coarse components", ncc, "sizes", sorted(np.bincount(labc).tolist())[-5:], flush=True)
    x = cp.asarray(np.random.default_rng(0).standard_normal((pat.ndof, 1)))
    y = fe._co.solve(x)
    print("coarse probe finite", bool(cp.isfinite(y).all()), flush=True)
    v = cp.asarray(np.random.default_rng(1).standard_normal(g.ndof))
    print("K v finite", bool(cp.isfinite(fe._apply_K(v)).all()), "cheb finite",
          bool(cp.isfinite(fe._cheb(v[:, None])).all()), "precond finite",
          bool(cp.isfinite(fe.precondition(v[:, None].copy())).all()), flush=True)


if __name__ == "__main__":
    main(sys.argv[1])
