"""Stage 0 gate on the sparse engine: V2-b (Study 2 tubes vs Abaqus), V2-c (1.7 mm-wall tube convergence 0.8 → 0.2 mm),
V3-a (clamped-free Euler rod). Same 3-point-bending set-up as tests/validate_tubes.py (line supports, line load,
span 288, length 345.6, E 1760, ν 0.42). Tubes use the quarter model (symmetry x = 0 and y = 0), checked against the
full model at 0.8 mm. Usage: python tubes.py out.json [h ...]"""
import json
import sys
import time
from pathlib import Path

import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parents[3]))
from cshell.fe_sparse import ActiveGrid, SparseFE, buckling  # noqa: E402

E, NU = 1760.0, 0.42
SIZES = {"A": (15.0, 0.0), "B": (15.8, 5.0), "C": (18.0, 10.0), "D": (21.2, 15.0), "E": (25.0, 20.0),
         "W1.7": (25.0, 21.6)}
ABAQUS = {"A": 0.0086, "B": 0.0104, "C": 0.0158, "D": 0.0245, "E": 0.0351}
SPAN, LENGTH, F = 288.0, 345.6, 100.0


def section_fraction(ro, ri, h, y0, ny, z0, nz, sub=6):
    o = (np.arange(sub) + 0.5) / sub
    Y = (y0 + h * (np.arange(ny)[:, None] + o[None, :])).ravel()
    Z = (z0 + h * (np.arange(nz)[:, None] + o[None, :])).ravel()
    r = np.sqrt(Y[:, None] ** 2 + Z[None, :] ** 2)
    return ((r <= ro) & (r >= ri)).reshape(ny, sub, nz, sub).mean((1, 3))


def tube(size, h, quarter=True, pw=4.0):
    od, idd = SIZES[size]
    ro, ri = od / 2, idd / 2
    nz = 2 * int(np.ceil(ro / h))
    z0 = -nz / 2 * h
    if quarter:
        ny, y0, nx, x0 = nz // 2, 0.0, int(round(LENGTH / 2 / h)), 0.0
    else:
        ny, y0, nx, x0 = nz, z0, int(round(LENGTH / h)), -LENGTH / 2
    sec = section_fraction(ro, ri, h, y0, ny, z0, nz)
    vf = np.repeat(sec[None], nx, 0)
    g = ActiveGrid(vf > 0)
    nn = g.node_shape
    xs = x0 + h * np.arange(nn[0]); ys = y0 + h * np.arange(nn[1]); zs = z0 + h * np.arange(nn[2])
    X = np.stack(np.meshgrid(xs, ys, zs, indexing="ij"), -1)
    act = np.zeros(int(np.prod(nn)), bool); act[g.node_dense] = True
    act = act.reshape(nn)
    zmin, zmax = zs[np.nonzero(act.any((0, 1)))[0][[0, -1]]]
    fixed = np.zeros(nn + (3,), bool); load = np.zeros(nn + (3,))
    on = lambda c, v: np.isclose(X[..., c], v, atol=1e-9)
    # load pin as a fixed-width patch, identical at every h (a single node line is a singular line load whose local
    # indentation compliance grows as h -> 0): 4 mm wide, 0.5 mm deep
    patch = lambda x0, w, zsel: act & (np.abs(X[..., 0] - x0) <= w / 2 + 1e-9) & zsel
    # supports: one node plane (rotation-free, a 3 mm-wide uz = 0 patch acts as a partial clamp), 0.5 mm deep band
    sup = patch(SPAN / 2, 1e-6, X[..., 2] <= zmin + 0.5)
    fixed[sup, 2] = True
    top = patch(0.0, pw, X[..., 2] >= zmax - 0.5)
    if quarter:
        sup_l = None
        fixed[on(0, 0.0), 0] = True                       # symmetry x = 0
        fixed[on(1, 0.0), 1] = True                       # symmetry y = 0
        w = np.where(on(1, 0.0), 0.5, 1.0) * np.where(on(0, 0.0), 0.5, 1.0)    # mirror multiplicity 1/(copies)
        load[top, 2] = -F * w[top] / (4 * float(w[top].sum()))    # full-model node count = 4 Σ w; total F/4
    else:
        supl = patch(-SPAN / 2, 1e-6, X[..., 2] <= zmin + 0.5)
        fixed[supl, 2] = True
        fixed[supl, 0] = True
        for s in (sup, supl):
            idx = np.argwhere(s); j = idx[np.argmin(np.abs(X[s][:, 1]))]; fixed[tuple(j) + (1,)] = True
        load[top, 2] = -F / top.sum()
    fe = SparseFE(g, fixed, load, h, E, NU, backend="auto")
    t0 = time.time()
    fe.factor(1e-6 + (1 - 1e-6) * g.gather_elems(vf))
    p = fe.primal()
    C = p.compliance * (4 if quarter else 1)
    k = F ** 2 / C / 1000.0
    I = np.pi / 64 * (od ** 4 - idd ** 4)
    out = dict(case="tube", size=size, h=h, pw=pw, quarter=quarter, wall=(od - idd) / 2, dof=g.ndof, k=k,
               k_EB=48 * E * I / SPAN ** 3 / 1000.0, k_abaqus=ABAQUS.get(size), seconds=time.time() - t0,
               factor_s=fe.stats.get("factor_s"))
    fe.free()
    return out


def euler_rod(h, r=5.0, L=100.0):
    """Clamped-free rod, axial tip load: λ₁·P vs π²EI/(4L²) (V3-a)."""
    n = 2 * int(np.ceil(r / h))
    sec = section_fraction(r, 0.0, h, -n / 2 * h, n, -n / 2 * h, n)
    nx = int(round(L / h))
    vf = np.repeat(sec[None], nx, 0)
    g = ActiveGrid(vf > 0)
    nn = g.node_shape
    fixed = np.zeros(nn + (3,), bool); fixed[0] = True
    load = np.zeros(nn + (3,))
    act = np.zeros(int(np.prod(nn)), bool); act[g.node_dense] = True
    tip = act.reshape(nn)[-1]
    P = 100.0
    load[-1][tip, 0] = -P / tip.sum()
    fe = SparseFE(g, fixed, load, h, E, NU, backend="auto")
    t0 = time.time()
    fe.factor(g.gather_elems(vf))
    pr = fe.primal()
    b = buckling(fe, pr, num_modes=4, tol=1e-6, min_density=0.1)
    Pcr = np.pi ** 2 * E * (np.pi * r ** 4 / 4) / (4 * L ** 2)
    out = dict(case="euler", h=h, dof=g.ndof, P_fe=float(P / b.mu.max()), P_cr=Pcr, err=float(P / b.mu.max() / Pcr - 1),
               mu=b.mu.tolist(), iters=b.iterations, converged=b.converged, seconds=time.time() - t0)
    fe.free()
    return out


if __name__ == "__main__":
    out = sys.argv[1]; rows = []
    def emit(r):
        rows.append(r); print(json.dumps(r), flush=True); Path(out).write_text(json.dumps(rows, indent=1))
    for size, h, pw in [("E", .8, 4), ("E", .4, 2), ("E", .4, 4), ("E", .4, 6), ("E", .4, 10), ("E", .2, 4), ("E", .2, 2),
                        ("W1.7", .8, 4), ("W1.7", .4, 4), ("W1.7", .2, 4)]:
        try: emit(tube(size, h, pw=pw))
        except Exception as e: emit(dict(size=size, h=h, pw=pw, error=repr(e)))
