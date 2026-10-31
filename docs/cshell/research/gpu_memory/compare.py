"""Gradient agreement of alternative solvers with the fp64 cuDSS reference (arrays saved under $SCRATCH).
Usage: python compare.py <ref_stem> <test_stem> ...   (stems without _dC.npy / _dmu.npy)"""
import json, os, sys
import numpy as np
D = os.path.join(os.environ["SCRATCH"], "gpu_memory_arrays")
ref = sys.argv[1]
out = {}
for t in sys.argv[2:]:
    r = {}
    for q in ("dC", "dmu"):
        a, b = np.load(f"{D}/{t}_{q}.npy"), np.load(f"{D}/{ref}_{q}.npy")
        d = np.random.default_rng(0).standard_normal(a.size)
        r[q] = dict(rel2=float(np.linalg.norm(a - b) / np.linalg.norm(b)),
                    relmax=float(np.abs(a - b).max() / np.abs(b).max()),
                    directional=float(abs(d @ (a - b)) / abs(d @ b)))
    out[t] = r
print(json.dumps(out, indent=1))
