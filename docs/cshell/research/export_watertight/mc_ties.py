"""Is skimage's Lewiner MC non-manifold only at exact ties of the face (asymptotic) decider? Random fields on a 48³
lattice, counted duplicated directed edges (skimage welds within one call). usage: python mc_ties.py"""
import numpy as np
from skimage.measure import marching_cubes


def dup(f):
    nv = f.max() + 1
    d = np.concatenate([f[:, 0] * nv + f[:, 1], f[:, 1] * nv + f[:, 2], f[:, 2] * nv + f[:, 0]]).astype(np.int64)
    d.sort()
    return int((d[1:] == d[:-1]).sum())


def run(psi):
    v, f, _, _ = marching_cubes(psi.astype(np.float32), 0.0, allow_degenerate=False)
    return dup(f)


rng = np.random.default_rng(0)
B, h = np.float32(0.15), 0.05
n = 48
for name, gen in [
    ("binary ±B", lambda: np.where(rng.random((n,) * 3) < 0.5, B, -B)),
    ("binary ±B, 10% material", lambda: np.where(rng.random((n,) * 3) < 0.9, B, -B)),
    ("gaussian", lambda: rng.normal(0, 0.05, (n,) * 3)),
    ("gaussian clipped ±B/3", lambda: np.clip(rng.normal(0, 0.1, (n,) * 3), -B / 3, B / 3)),
    ("ints ±1,±2", lambda: rng.choice([-2.0, -1.0, 1.0, 2.0], (n,) * 3)),
]:
    tot = {0.0: 0, 1e-4 * h: 0, -1e-4 * h: 0}
    for rep in range(4):
        p = gen().astype(np.float32)
        for s in tot:
            q = (p + np.float32(s)).astype(np.float32)
            eps = np.float32(1e-3 * h)
            q = np.where(np.abs(q) < eps, np.where(q < 0, -eps, eps), q).astype(np.float32)
            tot[s] += run(q)
    print(f"{name:<28} dup directed edges (4 fields): " + ", ".join(f"shift {s:+.0e}: {c}" for s, c in tot.items()),
          flush=True)
