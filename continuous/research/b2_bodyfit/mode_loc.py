"""Mode-1 location for the corrected meshes: top nodes, share of |phi|^2 within 3 mm of each escape hole / of the root."""
import json, sys
from pathlib import Path
import numpy as np
W = Path(__file__).resolve().parent / "work"
H = np.genfromtxt("runs_cshell/C_bench/B2_slender/escape_holes.csv", delimiter=",", names=True)
Hp = np.column_stack([H["x"], H["y"], H["z"]])
out = {}
for h in ("0.7", "0.5", "0.35"):
    d = np.load(W / f"m2_design_h{h}_mode.npz"); X, a = d["X"], d["a"][:, 0]
    w = a.astype(float) ** 2; w /= w.sum()
    top = np.argsort(a)[-200:]
    dh = np.linalg.norm(X[:, None] - Hp[None], axis=2)
    out[h] = dict(top200_x=[float(X[top, 0].min()), float(np.median(X[top, 0])), float(X[top, 0].max())],
                  top200_r=float(np.median(np.hypot(X[top, 1], X[top, 2]))),
                  top200_theta=float(np.median(np.degrees(np.arctan2(X[top, 2], X[top, 1])))),
                  top200_min_dist_hole=[float(dh[top].min(1).min()), float(np.median(dh[top].min(1)))],
                  share_within_3mm_hole=[float(w[dh[:, k] < 3].sum()) for k in range(len(Hp))],
                  share_x_lt_3=float(w[X[:, 0] < 3].sum()), share_x_gt_60=float(w[X[:, 0] > 60].sum()),
                  tip_share_top1pct=float(np.sort(w)[-len(w) // 100:].sum()))
json.dump(out, open(W / "mode_loc.json", "w"), indent=1); print(json.dumps(out, indent=1))
