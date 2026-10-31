"""B2 body-fitted diagnosis, part 1 (units mm, N, MPa).

A  thin-wall samples on best_candidate.stl (full MC mesh): where t < 0.5 mm, real sliver or ray artefact
B  the decimated validation surface (collect.decimate_for_validation, same call): self-intersections, Hausdorff, thin walls
C  the ccx design deck actually run (validation/design.inp): element type, tet quality, local h vs wall thickness
D  BC node sets in the design and solid decks: counts, x/r ranges, load sum, clamped/loaded area vs the STL end faces
Writes work/diag1.json and work/thin_samples.npz.
"""
import json, sys, time, re
from pathlib import Path
import numpy as np

RUN = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("runs_cshell/C_bench/B2_slender")
OUT = Path(__file__).resolve().parent / "work"
OUT.mkdir(exist_ok=True)
R = {}
t00 = time.time()
import trimesh, igl


def load_stl(p):
    m = trimesh.load(p, process=False)
    m.merge_vertices()
    return np.asarray(m.vertices, float), np.asarray(m.faces, np.int64)


def rays(V, F, n, seed=0, all_faces=False):
    tri = V[F]
    nrm = np.cross(tri[:, 1] - tri[:, 0], tri[:, 2] - tri[:, 0])
    area = 0.5 * np.linalg.norm(nrm, axis=1)
    nrm /= np.maximum(2 * area[:, None], 1e-30)
    rng = np.random.default_rng(seed)
    pick = np.arange(len(F)) if all_faces else rng.choice(len(F), size=n, p=area / area.sum())
    u = np.full((len(pick), 2), 1 / 3) if all_faces else rng.random((len(pick), 2))
    fl = u.sum(1) > 1
    u[fl] = 1 - u[fl]
    P = tri[pick, 0] + u[:, :1] * (tri[pick, 1] - tri[pick, 0]) + u[:, 1:] * (tri[pick, 2] - tri[pick, 0])
    N = nrm[pick]
    aabb = igl.AABB(); aabb.init(V, F)
    res = aabb.intersect_ray_first(V, F, np.ascontiguousarray(P - 1e-4 * N), np.ascontiguousarray(-N))
    hit, t = np.asarray(res[0]), np.asarray(res[1], float) + 1e-4
    ok = hit >= 0
    cosh = np.full(len(P), np.nan)
    cosh[ok] = np.einsum("ij,ij->i", nrm[hit[ok]], N[ok])     # ≈ −1 for a real opposite wall
    return P, N, t, hit, ok, cosh, pick


def cyl(P):
    return np.column_stack([P[:, 0], np.hypot(P[:, 1], P[:, 2]), np.degrees(np.arctan2(P[:, 2], P[:, 1]))])


holes = np.genfromtxt(RUN / "escape_holes.csv", delimiter=",", names=True)
H = np.column_stack([holes["x"], holes["y"], holes["z"]])
Hn = np.column_stack([holes["nx"], holes["ny"], holes["nz"]])

# ---------------- A: full MC mesh -----------------
V, F = load_stl(RUN / "best_candidate.stl")
R["A_mesh"] = dict(V=len(V), F=len(F), bbox=[V.min(0).tolist(), V.max(0).tolist()])
P, N, t, hit, ok, cosh, pick = rays(V, F, 0, all_faces=True)       # one ray per face centroid (5.8 M)
t = np.where(ok, t, np.nan)
thin = ok & (t < 0.5)
C = cyl(P[thin])
dh = np.min(np.linalg.norm(P[thin][:, None] - H[None], axis=2), axis=1) if thin.any() else np.zeros(0)
R["A_wall_all_faces"] = dict(n=int(ok.sum()), miss=int((~ok).sum()), min=float(np.nanmin(t)),
                             q001=float(np.nanquantile(t, 1e-3)), q01=float(np.nanquantile(t, 0.01)),
                             median=float(np.nanmedian(t)), n_lt_0p5=int(thin.sum()), n_lt_1p0=int((ok & (t < 1.0)).sum()),
                             n_lt_1p4=int((ok & (t < 1.4)).sum()), frac_lt_1p5=float(np.nanmean(t < 1.5)))
if thin.any():
    R["A_thin"] = dict(x_range=[float(C[:, 0].min()), float(C[:, 0].max())], r_range=[float(C[:, 1].min()), float(C[:, 1].max())],
                       cos_hit_normal_quantiles=np.nanquantile(cosh[thin], [0, .1, .5, .9, 1]).tolist(),
                       dist_to_hole_quantiles=np.quantile(dh, [0, .1, .5, .9, 1]).tolist(),
                       frac_within_2mm_of_hole=float((dh < 2.0).mean()),
                       frac_within_1mm_of_ends=float(((C[:, 0] < V[:, 0].min() + 1) | (C[:, 0] > V[:, 0].max() - 1)).mean()),
                       frac_real_wall=float((cosh[thin] < -0.7).mean()))
    hx, _ = np.histogram(C[:, 0], bins=np.arange(-1, 82, 2.0))
    R["A_thin"]["x_hist_2mm"] = hx.tolist()
    order = np.argsort(t[thin])[:30]
    R["A_thin"]["worst30"] = [dict(t=float(t[thin][i]), x=float(C[i, 0]), r=float(C[i, 1]), th=float(C[i, 2]),
                                   cos=float(cosh[thin][i]), d_hole=float(dh[i])) for i in order]
np.savez_compressed(OUT / "thin_samples.npz", P=P[thin], N=N[thin], t=t[thin], cos=cosh[thin])
# wall-thickness profile along x (outer-surface samples only: those that hit a real opposite wall)
real = ok & (cosh < -0.7)
xb = np.arange(0, 81, 4.0)
prof = []
for a, b in zip(xb[:-1], xb[1:]):
    s = real & (P[:, 0] >= a) & (P[:, 0] < b)
    prof.append([float(a), float(b), float(np.nanquantile(t[s], 0.01)) if s.any() else None,
                 float(np.nanmedian(t[s])) if s.any() else None, float(np.hypot(P[s, 1], P[s, 2]).max()) if s.any() else None])
R["A_profile_x0_x1_t01_tmed_rmax"] = prof
print("A done", time.time() - t00, flush=True)

# ---------------- B: decimated surface (exact collect call) -----------------
sys.path.insert(0, str(Path.cwd()))
from cshell.collect import decimate_for_validation
V2, F2, info = decimate_for_validation(V, F, 0.25)
R["B_decimation"] = {k: (float(v) if isinstance(v, (float, np.floating)) else v) for k, v in info.items()}
np.savez_compressed(OUT / "decimated.npz", V=V2, F=F2)
import pymeshlab
ms = pymeshlab.MeshSet(); ms.add_mesh(pymeshlab.Mesh(V2, F2.astype(np.int32)))
ms.compute_selection_by_self_intersections_per_face()
nsi = int(ms.current_mesh().selected_face_number())
R["B_self_intersecting_faces"] = nsi
e = np.linalg.norm(V2[F2[:, [1, 2, 0]]] - V2[F2], axis=2)
tri = V2[F2]
A2 = 0.5 * np.linalg.norm(np.cross(tri[:, 1] - tri[:, 0], tri[:, 2] - tri[:, 0]), axis=1)
q = 4 * np.sqrt(3) * A2 / (e ** 2).sum(1)                      # 1 equilateral, 0 degenerate
R["B_tri_quality"] = dict(edge_q=[float(x) for x in np.quantile(e, [0, .01, .5, .99, 1])],
                          q_min=float(q.min()), n_q_lt_0p1=int((q < .1).sum()), n_q_lt_0p3=int((q < .3).sum()))
d2 = igl.point_mesh_squared_distance(V2, V, F)[0]
d1 = igl.point_mesh_squared_distance(V[:: 7], V2, F2)[0]
R["B_hausdorff"] = dict(dec_to_mc_max=float(np.sqrt(d2.max())), mc_to_dec_max=float(np.sqrt(d1.max())),
                        mc_to_dec_q999=float(np.sqrt(np.quantile(d1, .999))))
P2, N2, t2, h2, ok2, c2, _ = rays(V2, F2, 0, all_faces=True)
t2 = np.where(ok2, t2, np.nan)
R["B_wall"] = dict(min=float(np.nanmin(t2)), q001=float(np.nanquantile(t2, 1e-3)), n_lt_0p5=int((t2 < .5).sum()),
                   n_lt_0p5_realwall=int(((t2 < .5) & (c2 < -.7)).sum()))
print("B done", time.time() - t00, flush=True)


# ---------------- C: the tet deck that was run -----------------
def read_deck(p):
    txt = open(p).read()
    def block(start):
        i = txt.index(start); i = txt.index("\n", i) + 1; j = txt.index("*", i)
        return txt[i:j]
    nodes = np.array(block("*NODE, NSET=NALL").replace(",", " ").split(), float).reshape(-1, 4)
    m = re.search(r"\*ELEMENT, TYPE=(\w+)", txt)
    el = np.array(block("*ELEMENT").replace(",", " ").split(), np.int64)
    el = el.reshape(-1, 11 if m.group(1) == "C3D10" else 5)
    nfix = np.array(block("*NSET, NSET=NFIX").replace(",", " ").split(), np.int64)
    nload = np.array(block("*NSET, NSET=NLOAD").replace(",", " ").split(), np.int64)
    cl = np.array(block("*STATIC").split("*CLOAD")[1].replace(",", " ").split() if False else
                  txt[txt.index("*CLOAD") + 7: txt.index("*NODE PRINT")].replace(",", " ").split(), float).reshape(-1, 3)
    return m.group(1), nodes[:, 1:], el[:, 1:] - 1, nfix - 1, nload - 1, cl


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


def deck_stats(p, name):
    et, X, T, nfix, nload, cl = read_deck(p)
    vol = np.einsum("ei,ei->e", np.cross(X[T[:, 1]] - X[T[:, 0]], X[T[:, 2]] - X[T[:, 0]]), X[T[:, 3]] - X[T[:, 0]]) / 6
    dm = dihedral_min(X, T)
    L = np.stack([np.linalg.norm(X[T[:, i]] - X[T[:, j]], axis=1) for i, j in ((0, 1), (1, 2), (0, 2), (0, 3), (1, 3), (2, 3))], 1)
    Lrms = np.sqrt((L ** 2).mean(1))
    eta = 6 * np.sqrt(2) * vol / Lrms ** 3                      # volume-length quality, 1 regular
    cen = X[T[:, :4]].mean(1)
    Cc = cyl(cen)
    bad = dm < 5
    # mid-edge nodes straight?
    mid = X[T[:, 4]] - 0.5 * (X[T[:, 0]] + X[T[:, 1]])
    s = dict(element_type=et, nodes=len(X), elements=len(T), volume=float(vol.sum()), neg_vol=int((vol <= 0).sum()),
             dihedral_min_deg=float(dm.min()), n_dihedral_lt_5=int(bad.sum()), n_dihedral_lt_10=int((dm < 10).sum()),
             dihedral_q=[float(x) for x in np.quantile(dm, [.001, .01, .5])], eta_min=float(eta.min()),
             n_eta_lt_0p1=int((eta < .1).sum()), edge_q=[float(x) for x in np.quantile(L, [0, .01, .5, .99, 1])],
             midedge_offset_max=float(np.abs(mid).max()),
             bad_x_hist_4mm=np.histogram(Cc[bad, 0], bins=np.arange(0, 84, 4.0))[0].tolist() if bad.any() else [],
             n_fix=len(nfix), n_load=len(nload), fix_x=[float(X[nfix, 0].min()), float(X[nfix, 0].max())],
             fix_r=[float(np.hypot(X[nfix, 1], X[nfix, 2]).min()), float(np.hypot(X[nfix, 1], X[nfix, 2]).max())],
             load_x=[float(X[nload, 0].min()), float(X[nload, 0].max())],
             load_r=[float(np.hypot(X[nload, 1], X[nload, 2]).min()), float(np.hypot(X[nload, 1], X[nload, 2]).max())],
             load_sum=float(cl[:, 2].sum()), load_dofs=sorted(set(cl[:, 1].astype(int).tolist())))
    return s, X, T, cen, Lrms, dm


Rd, X, T, cen, Lrms, dm = deck_stats(RUN / "validation/design.inp", "design")
R["C_design_deck"] = Rd
Rs = deck_stats(RUN / "validation/solid.inp", "solid")[0]
R["C_solid_deck"] = Rs
np.savez_compressed(OUT / "design_deck_quality.npz", cen=cen, Lrms=Lrms, dmin=dm)
# elements through the wall: wall thickness at a decimated-surface sample / rms edge of the tets near the mid-wall point
from scipy.spatial import cKDTree
tree = cKDTree(cen)
rs = (ok2 & (c2 < -.7))
idx = np.flatnonzero(rs)[:: max(1, rs.sum() // 20000)]
mid = P2[idx] - 0.5 * t2[idx, None] * N2[idx]
dd, jj = tree.query(mid, k=8)
hl = Lrms[jj].mean(1)
nl = t2[idx] / hl
R["C_layers"] = dict(samples=len(idx), layers_q=[float(x) for x in np.quantile(nl, [0, .01, .1, .5])],
                     frac_lt_2=float((nl < 2).mean()), frac_lt_1p5=float((nl < 1.5).mean()),
                     h_local_q=[float(x) for x in np.quantile(hl, [.01, .5, .99])])
R["seconds"] = time.time() - t00
json.dump(R, open(OUT / "diag1.json", "w"), indent=1)
print(json.dumps(R, indent=1)[:6000])
