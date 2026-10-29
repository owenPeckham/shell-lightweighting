"""Are the sub-0.1 mm walls of best_candidate.stl real films? (units mm)
1. rays along each escape-hole axis from outside (grid of offsets within 0.8 r): list of hits -> open or film
2. real-wall thin samples (t < 0.1, opposite normal cos < -0.7) from work/thin_samples.npz: cluster and locate
   relative to the holes (axial coordinate s along n from the hole site, radial distance rho from the axis)
3. connected components of the surface and of the 'thin' faces
"""
import json, sys
from pathlib import Path
import numpy as np, trimesh, igl

RUN = Path(sys.argv[1]); OUT = Path(__file__).resolve().parent / "work"
m = trimesh.load(RUN / "best_candidate.stl", process=False); m.merge_vertices()
V, F = np.asarray(m.vertices, float), np.asarray(m.faces, np.int64)
H = np.genfromtxt(RUN / "escape_holes.csv", delimiter=",", names=True)
cfg = json.load(open(RUN / "config_used.json"))
R = dict(hole_d=None)
try:
    R["hole_d"] = cfg["manufacturing"]["escape_hole_diameter_mm"]
except Exception:
    pass
r = 0.5 * (R["hole_d"] or 3.0)
fn = m.face_normals
res = []
for h in H:
    p = np.array([h["x"], h["y"], h["z"]]); n = np.array([h["nx"], h["ny"], h["nz"]]); n /= np.linalg.norm(n)
    e1 = np.cross(n, [1, 0, 0]); e1 = e1 if np.linalg.norm(e1) > .1 else np.cross(n, [0, 1, 0]); e1 /= np.linalg.norm(e1)
    e2 = np.cross(n, e1)
    rows = []
    for rho in (0.0, 0.4 * r, 0.8 * r):
        for ang in (np.linspace(0, 2 * np.pi, 8, endpoint=False) if rho > 0 else [0.0]):
            o = p + rho * (np.cos(ang) * e1 + np.sin(ang) * e2) + (h["wall_mm"] + 3.0) * n
            loc, ri, ti = m.ray.intersects_location([o], [-n], multiple_hits=True)
            s = np.sort((loc - p) @ n)[::-1]          # axial coordinate of hits, outside first
            rows.append(dict(rho=float(rho), hits_s=[round(float(x), 4) for x in s[:6]]))
    res.append(dict(hole=int(h["hole"]), p=p.tolist(), n=n.tolist(), wall=float(h["wall_mm"]), rays=rows))
R["axis_rays"] = res
T = np.load(OUT / "thin_samples.npz")
sel = (T["t"] < 0.1) & (T["cos"] < -0.7)
P = T["P"][sel]
R["n_real_thin_lt_0p1"] = int(sel.sum())
loc = []
for h in H:
    p = np.array([h["x"], h["y"], h["z"]]); n = np.array([h["nx"], h["ny"], h["nz"]]); n /= np.linalg.norm(n)
    d = P - p; s = d @ n; rho = np.linalg.norm(d - s[:, None] * n, axis=1)
    loc.append(dict(hole=int(h["hole"]), near=int((np.hypot(s, rho - r) < 1.0).sum()),
                    s_q=np.quantile(s, [0, .5, 1]).round(3).tolist() if len(s) else [],
                    rho_q=np.quantile(rho, [0, .5, 1]).round(3).tolist() if len(s) else []))
R["thin_vs_holes"] = loc
# all thin (t<0.5) samples: how many are at the hole bores (rho ~ r, s within the wall)?
Pa, ta = T["P"], T["t"]
for h in H:
    p = np.array([h["x"], h["y"], h["z"]]); n = np.array([h["nx"], h["ny"], h["nz"]]); n /= np.linalg.norm(n)
    d = Pa - p; s = d @ n; rho = np.linalg.norm(d - s[:, None] * n, axis=1)
    R.setdefault("thin_lt_0p5_on_bore", []).append(int(((np.abs(rho - r) < 0.15) & (s > -0.6) & (s < h["wall_mm"] + 0.6)).sum()))
R["n_thin_lt_0p5"] = int(len(ta))
cc = trimesh.graph.connected_components(m.face_adjacency, min_len=1)
R["surface_components"] = [len(c) for c in cc][:10]
json.dump(R, open(OUT / "holes_check.json", "w"), indent=1)
print(json.dumps(R, indent=1))
