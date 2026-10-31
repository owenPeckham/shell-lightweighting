"""Residual non-manifold edges of skimage Lewiner MC on tie-free (gaussian) fields: which configurations?
usage: python mc_resid.py"""
import numpy as np
from skimage.measure import marching_cubes


def dups(v, f):
    nv = len(v)
    d = np.concatenate([f[:, 0] * nv + f[:, 1], f[:, 1] * nv + f[:, 2], f[:, 2] * nv + f[:, 0]]).astype(np.int64)
    s = np.sort(d)
    dd = np.unique(s[1:][s[1:] == s[:-1]])
    return dd // nv, dd % nv


def ambiguous_faces(p):
    """count lattice faces with a (+,-,+,-) diagonal sign pattern"""
    s = p > 0
    c = 0
    for ax in range(3):
        a, b = [x for x in range(3) if x != ax]
        q = np.moveaxis(s, (a, b), (0, 1))
        c += int(((q[:-1, :-1] == q[1:, 1:]) & (q[1:, :-1] == q[:-1, 1:]) & (q[:-1, :-1] != q[1:, :-1])).sum())
    return c


rng = np.random.default_rng(1)
n = 48
tot, shown = 0, 0
for rep in range(40):
    p = rng.normal(0, 0.05, (n,) * 3).astype(np.float32)
    v, f, _, _ = marching_cubes(p, 0.0, allow_degenerate=False)
    a, b = dups(v, f)
    tot += len(a)
    for i, j in zip(a[:2], b[:2]):
        if shown >= 6:
            break
        shown += 1
        cub = np.floor(np.minimum(v[i], v[j])).astype(int)
        mid = 0.5 * (v[i] + v[j])
        cub = np.floor(mid).astype(int)
        print(f"rep {rep} edge {v[i].round(3)} -> {v[j].round(3)}  cube {cub}")
        for c0 in [cub + np.array(o) for o in [(0, 0, 0), (-1, 0, 0), (0, -1, 0), (0, 0, -1)]]:
            if (c0 < 0).any() or (c0 + 1 >= n).any():
                continue
            blk = p[c0[0]:c0[0] + 2, c0[1]:c0[1] + 2, c0[2]:c0[2] + 2]
            print(f"   cube {c0.tolist()} signs {(blk > 0).astype(int).ravel().tolist()} amb faces {ambiguous_faces(blk)}")
print("gaussian, 40 fields of 48^3: duplicated directed edges", tot)
# same fields with no ambiguous face: flip negative nodes of ambiguous faces to +, iterate
tot2 = 0
rng = np.random.default_rng(1)
for rep in range(40):
    p = rng.normal(0, 0.05, (n,) * 3).astype(np.float32)
    for it in range(50):
        s = p > 0
        flip = np.zeros_like(s)
        for ax in range(3):
            a_, b_ = [x for x in range(3) if x != ax]
            q = np.moveaxis(s, (a_, b_), (0, 1))
            amb = (q[:-1, :-1] == q[1:, 1:]) & (q[1:, :-1] == q[:-1, 1:]) & (q[:-1, :-1] != q[1:, :-1])
            fl = np.zeros_like(q)
            fl[:-1, :-1] |= amb & ~q[:-1, :-1]
            fl[1:, :-1] |= amb & ~q[1:, :-1]
            flip |= np.moveaxis(fl, (0, 1), (a_, b_))
        if not flip.any():
            break
        p[flip] = 0.01
    v, f, _, _ = marching_cubes(p, 0.0, allow_degenerate=False)
    tot2 += len(dups(v, f)[0])
print("same fields, ambiguous faces removed: duplicated directed edges", tot2)
