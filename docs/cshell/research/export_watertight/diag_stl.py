"""Where are the duplicated directed edges / extra bodies of an exported STL?  usage: python diag_stl.py <stl> [h]
Welds by exact float32 position (the exporter writes V[F] of its welded mesh, so this reproduces its topology).
Lattice: vertex coords u = (x - origin)/h - 1/2; a lattice-edge vertex has exactly 2 integral coordinates; the origin's
fractional offset is estimated per axis as the mode of frac(x/h)."""
import sys
import numpy as np
from scipy import sparse
from scipy.sparse.csgraph import connected_components


def load(p):
    dt = np.dtype([("n", "<f4", 3), ("v", "<f4", (3, 3)), ("attr", "<u2")])
    a = np.fromfile(p, dtype=dt, offset=84)
    P = np.ascontiguousarray(a["v"].reshape(-1, 3))
    del a
    k = P.view(np.dtype((np.void, 12))).ravel()
    uk, first, inv = np.unique(k, return_index=True, return_inverse=True)
    return P[first].astype(np.float64), inv.reshape(-1, 3).astype(np.int64)


def main(p, h=0.05):
    V, F = load(p)
    nv, nf = len(V), len(F)
    print(f"{p}: V {nv} F {nf}", flush=True)
    off = []
    for a in range(3):
        fr = np.mod(V[:, a] / h, 1.0)
        hist, edges = np.histogram(fr, bins=2000)
        off.append(edges[hist.argmax()] + 0.5 / 2000)
    U = V / h - np.asarray(off)
    fr = np.abs(U - np.rint(U))
    integral = (fr < 2e-3).sum(1)
    print("vertices by # integral coords:", np.bincount(integral, minlength=4).tolist())
    d = np.concatenate([F[:, 0] * nv + F[:, 1], F[:, 1] * nv + F[:, 2], F[:, 2] * nv + F[:, 0]])
    ds = np.sort(d)
    dupm = ds[1:] == ds[:-1]
    dd = np.unique(ds[1:][dupm])
    print(f"duplicated directed edges: {int(dupm.sum())} (distinct {len(dd)})")
    a_, b_ = dd // nv, dd % nv
    va = np.unique(np.concatenate([a_, b_]))
    print(" their vertices by # integral coords:", np.bincount(integral[va], minlength=4).tolist())
    Ui = np.rint(U[va]).astype(np.int64) % 128
    print(" vertices near a tile face (rint index%128 in {0,127,1}):",
          int(((Ui == 0) | (Ui == 127) | (Ui == 1)).any(1).sum()), "/", len(va))
    for e in dd[:12]:
        i, j = e // nv, e % nv
        fa = np.nonzero(d == e)[0] % nf
        print(f"  edge {V[i].round(4).tolist()} -> {V[j].round(4).tolist()}  ints {integral[i]},{integral[j]}"
              f"  faces on it: {len(fa)} rev: {int((d == (j * nv + i)).sum())}  U_i {U[i].round(3).tolist()}"
              f" U_j {U[j].round(3).tolist()}")
        for f_ in fa:
            print("     face", [U[q].round(3).tolist() for q in F[f_]])
    g = sparse.coo_matrix((np.ones(3 * nf), (np.concatenate([F[:, 0], F[:, 1], F[:, 2]]),
                                              np.concatenate([F[:, 1], F[:, 2], F[:, 0]]))), shape=(nv, nv))
    nc, lab = connected_components(g, directed=False)
    fl = lab[F[:, 0]]
    vol = np.bincount(fl, weights=np.einsum("ij,ij->i", V[F[:, 0]], np.cross(V[F[:, 1]], V[F[:, 2]])) / 6.0,
                      minlength=nc)
    cnt = np.bincount(fl, minlength=nc)
    order = np.argsort(-np.abs(vol))
    print(f"bodies: {nc}")
    hasdup = np.zeros(nc, bool)
    hasdup[lab[va]] = True
    for c in order[:40]:
        sel = F[fl == c].ravel()
        lo, hi = V[sel].min(0), V[sel].max(0)
        print(f"  body {c}: vol {vol[c]:+.5g} mm3 faces {cnt[c]} bbox {lo.round(2).tolist()}..{hi.round(2).tolist()}"
              f" dup-edges {bool(hasdup[c])}")


if __name__ == "__main__":
    main(sys.argv[1], float(sys.argv[2]) if len(sys.argv) > 2 else 0.05)
