"""Corrected body-fitted mesh + GPU solve for one case and one level (units mm, N, MPa).

usage: python mesh2.py <case: design|solid> <h_mm> [--ccx-deck] [--stl path]
  surface: (design) MC STL -> quadric pre-decimation to edge h/3 -> pymeshlab isotropic remesh to edge h
           (solid) CAD tessellation from bc_faces.npz -> isotropic remesh to edge h
  checks : watertight, self-intersections, volume vs input, Hausdorff to input
  volume : gmsh Delaunay inside the given surface triangulation (no reparametrisation), Netgen optimisation,
           straight C3D10 (cshell.validate.to_quadratic); BCs by cshell.validate.build_case (area-weighted traction,
           bc_tol = 0.75 * 0.4 mm as collect)
  quality: min dihedral, slivers, tet layers through the wall
  solve  : bf_gpu.solve (cuDSS); optionally writes the CalculiX deck for a cross-check
Writes work/m2_<case>_h<h>.json (+ .inp).
"""
import json, sys, time, tempfile
from pathlib import Path
import numpy as np

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE)); sys.path.insert(0, str(Path.cwd()))
import bf_gpu
from cshell import validate as cval

case, h = sys.argv[1], float(sys.argv[2])
want_deck = "--ccx-deck" in sys.argv
stl = sys.argv[sys.argv.index("--stl") + 1] if "--stl" in sys.argv else None
tag = sys.argv[sys.argv.index("--tag") + 1] if "--tag" in sys.argv else ""
RUN = Path("runs_cshell/C_bench/B2_slender")
OUT = HERE / "work"
name = f"m2_{case}{tag}_h{h:g}"
R = dict(case=case, h=h, stl=stl)
t0 = time.time()
import trimesh, pymeshlab, igl

bc = np.load(RUN / "bc_faces.npz")
cV, cT = bc["V"].astype(float), bc["T"].astype(np.int64)
TF = np.arange(len(cT))
support = TF[bc["support"]].tolist()
loads = [dict(faces=TF[bc["load"] == 0].tolist(), force=[0.0, 0.0, -20.0])]

if case == "design":
    m = trimesh.load(stl or RUN / "best_candidate.stl", process=False); m.merge_vertices()
    V0, F0 = np.asarray(m.vertices, float), np.asarray(m.faces, np.int64)
else:
    V0, F0 = cV, cT
vol0 = float(trimesh.Trimesh(V0, F0, process=False).volume)
ms = pymeshlab.MeshSet(); ms.add_mesh(pymeshlab.Mesh(V0, F0.astype(np.int32)))
if len(F0) > 50 * trimesh.Trimesh(V0, F0, process=False).area / h ** 2:
    tgt = int(trimesh.Trimesh(V0, F0, process=False).area / (0.433 * (h / 3) ** 2))
    ms.meshing_decimation_quadric_edge_collapse(targetfacenum=tgt, preservenormal=True, preservetopology=True,
                                                qualitythr=0.5, planarquadric=True)
ms.meshing_isotropic_explicit_remeshing(targetlen=pymeshlab.PureValue(h), iterations=8, featuredeg=60,
                                        checksurfdist=True, maxsurfdist=pymeshlab.PureValue(0.02))
# clean-up of the escape-hole films: delete self-intersecting faces + 3 rings, close the holes, relax, re-check
R["repair"] = []
for it in range(4):
    ms.compute_selection_by_self_intersections_per_face()
    nsi = int(ms.current_mesh().selected_face_number())
    if nsi == 0:
        break
    fsel = ms.current_mesh().face_selection_array()
    cm = ms.current_mesh()
    loc = (cm.vertex_matrix()[cm.face_matrix()[fsel]].mean(1)).round(2).tolist()
    R["repair"].append(dict(iter=it, n_self_int=nsi, where=loc[:8]))
    for _ in range(3):
        ms.compute_selection_transfer_face_to_vertex(inclusive=False)
        ms.compute_selection_transfer_vertex_to_face(inclusive=False)
        ms.apply_selection_dilatation()
    ms.meshing_remove_selected_faces()
    ms.meshing_remove_unreferenced_vertices()
    ms.meshing_close_holes(maxholesize=2000, newfaceselected=True)
    ms.meshing_isotropic_explicit_remeshing(targetlen=pymeshlab.PureValue(h), iterations=4, featuredeg=60,
                                            selectedonly=True, checksurfdist=False)
sm = ms.current_mesh()
V1, F1 = sm.vertex_matrix().astype(float), sm.face_matrix().astype(np.int64)
# project every vertex back onto the input surface (removes the flat hole-fill deviation); where that re-creates a
# self-intersection (sub-h films) keep the unprojected vertices of those faces
Vp = igl.point_mesh_squared_distance(V1, V0, F0)[2]
Vt = Vp.copy()
for it in range(5):
    ms2 = pymeshlab.MeshSet(); ms2.add_mesh(pymeshlab.Mesh(Vt, F1.astype(np.int32)))
    ms2.compute_selection_by_self_intersections_per_face()
    sel = ms2.current_mesh().face_selection_array()
    R["repair"].append(dict(project_iter=it, n_self_int=int(sel.sum())))
    if not sel.any():
        break
    vs = np.unique(F1[sel]); ring = np.unique(F1[np.isin(F1, vs).any(1)])
    Vt[ring] = V1[ring]
V1 = Vt
ms = pymeshlab.MeshSet(); ms.add_mesh(pymeshlab.Mesh(V1, F1.astype(np.int32)))
ms.compute_selection_by_self_intersections_per_face()
nsi = int(ms.current_mesh().selected_face_number())
print("repair", R["repair"], flush=True)
t1 = trimesh.Trimesh(V1, F1, process=False)
haus = float(np.sqrt(igl.point_mesh_squared_distance(V1, V0, F0)[0].max()))
haus2 = float(np.sqrt(igl.point_mesh_squared_distance(V0[:: max(1, len(V0) // 200000)], V1, F1)[0].max()))
e = np.linalg.norm(V1[F1[:, [1, 2, 0]]] - V1[F1], axis=2)
R["surface"] = dict(faces=len(F1), watertight=bool(t1.is_watertight), self_intersections=nsi,
                    vol_rel=float(t1.volume / vol0 - 1), hausdorff_out_in=haus, hausdorff_in_out=haus2,
                    edge_q=np.quantile(e, [0, .01, .5, .99]).round(4).tolist(), seconds=time.time() - t0)
print(R["surface"], flush=True)
if not t1.is_watertight or nsi:
    json.dump(R, open(OUT / f"{name}.json", "w"), indent=1)
    raise SystemExit("surface check failed")

# ---- volume mesh with gmsh inside the given triangulation ----
import gmsh
t2 = time.time()
tmp = Path(tempfile.mkdtemp(prefix="m2_", dir=str(OUT)))
t1.export(tmp / "s.stl")
gmsh.initialize(); gmsh.option.setNumber("General.Terminal", 0)
gmsh.option.setNumber("General.NumThreads", 16)
gmsh.merge(str(tmp / "s.stl"))
s = gmsh.model.getEntities(2)
l = gmsh.model.geo.addSurfaceLoop([x[1] for x in s])
gmsh.model.geo.addVolume([l]); gmsh.model.geo.synchronize()
gmsh.option.setNumber("Mesh.MeshSizeMax", h)
gmsh.option.setNumber("Mesh.Algorithm3D", 1)
gmsh.option.setNumber("Mesh.Optimize", 1)
gmsh.option.setNumber("Mesh.OptimizeNetgen", 1)
gmsh.model.mesh.generate(3)
tags, xyz, _ = gmsh.model.mesh.getNodes()
types, _, en = gmsh.model.mesh.getElements(3)
X = xyz.reshape(-1, 3)
idx = np.zeros(int(tags.max()) + 1, np.int64); idx[tags.astype(np.int64)] = np.arange(len(tags))
T4 = idx[en[list(types).index(4)].astype(np.int64).reshape(-1, 4)]
gmsh.finalize()
used = np.unique(T4); remap = -np.ones(len(X), np.int64); remap[used] = np.arange(len(used)); X, T4 = X[used], remap[T4]
T4 = cval._orient(X, T4)
nodes, tets10 = cval.to_quadratic(X, T4)
R["volume_mesh"] = dict(nodes=len(nodes), tets=len(tets10), seconds=time.time() - t2)
# quality
import diag1_lib as dl
dm = dl.dihedral_min(nodes, tets10)
vol = cval.tet_volumes(nodes, tets10)
R["quality"] = dict(dihedral_min=float(dm.min()), n_lt_5=int((dm < 5).sum()), n_lt_10=int((dm < 10).sum()),
                    dihedral_q001=float(np.quantile(dm, 1e-3)), neg_vol=int((vol <= 0).sum()),
                    vol_rel=float(vol.sum() / vol0 - 1))
R["layers"] = dl.layers(V1, F1, nodes, tets10)
print(R["quality"], R["layers"], flush=True)
caseo = cval.build_case(nodes, tets10, cV, cT, TF, support, loads, 0.75 * 0.4)
if "--clamp-plane" in sys.argv:      # diagnostic: clamp the whole root cross-section (incl. the lip beyond r0 = 5 mm)
    caseo.fixed = np.flatnonzero(nodes[:, 0] <= 0.3)
    caseo.info["n_fixed"] = int(caseo.fixed.size)
    caseo.info["clamp"] = "plane x<=0.3"
R["bc"] = caseo.info
if want_deck:
    cval.write_inp(OUT / f"{name}.inp", caseo, n_modes=4, want_stress=False)
out = bf_gpu.solve(caseo.nodes, caseo.tets10, caseo.fixed, caseo.load_nodes, caseo.load_vec, cval.E_MPA, cval.NU, 6)
phi = out.pop("_phi_nodes")
np.savez_compressed(OUT / f"{name}_mode.npz", X=nodes[: len(X)].astype(np.float32), a=phi[: len(X)].astype(np.float32),
                    F=F1.astype(np.int32), V=V1.astype(np.float32))
R["solve"] = out
R["seconds"] = time.time() - t0
json.dump(R, open(OUT / f"{name}.json", "w"), indent=1, default=float)
print(json.dumps(R, indent=1, default=float))
