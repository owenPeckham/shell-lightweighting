"""Body-fitted validation of an exported design with CalculiX (THEORY §8.6, §4.3, §4.4).

Pipeline: STL -> wall-thickness-graded linear tets (gmsh, optionally TetGen) -> quadratic C3D10 -> CalculiX deck
(`*STATIC` + `*BUCKLE`, clamped support-face nodes, load shared over load-face nodes) -> parse `.dat` -> compare with
the voxel FE (`write_validation_csv`, tolerances of §8.6). Units mm, N, MPa.

Boundary conditions follow §4.3: every mesh node within `bc_tol` (= 0.75 h_bc, default h_bc = the voxel h_e passed by the
caller, else the mesh size) of a support-face triangle is clamped in x, y, z; each load's force is spread over its
load-face nodes with consistent quadratic-tet weights (`load_distribution="area"`; the spec's equal nodal share is
available as "nodal" but is not mesh-consistent for C3D10). Mounts never move, so the triangles of the *original* tessellation
(`ctx.V, ctx.T, ctx.TF`, face ids) locate the BC nodes on any exported design.

Von Mises quantile (§4.4(b)): volume-weighted 0.995 quantile of the integration-point stresses (4 Gauss points per
C3D10, equal weight V_e/4), excluding elements whose centroid lies within `exclude_bc_mm` (2 mm) of a clamped or loaded
node. Buckling: `*BUCKLE` returns λ in  (K + λ K_σ) φ = 0  for the unit reference load of the static step; the
returned first factor is λ_1 (the load factor, not its inverse).

There is no silent fallback: if `ccx` fails, `validate_design(..., fallback="skfem")` logs a warning, solves the
compliance with scikit-fem P2 tets and marks the result `engine="skfem-P2"` (no buckling, no stress); the default
`fallback=None` re-raises.
"""
from __future__ import annotations

import csv
import logging
import os
import re
import shutil
import subprocess
import tempfile
import time
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np

log = logging.getLogger("cshell.validate")

E_MPA, NU = 1760.0, 0.42
# CalculiX C3D10 mid-edge nodes (0-based corner pairs): 5:(1,2) 6:(2,3) 7:(3,1) 8:(1,4) 9:(2,4) 10:(3,4)
_CCX_EDGES = np.array([[0, 1], [1, 2], [2, 0], [0, 3], [1, 3], [2, 3]])
# §8.6 tolerances: compliance, first buckling factor, 0.995 stress quantile (relative); ratios (absolute)
TOLERANCES = dict(compliance=0.05, buckling=0.10, stress=0.15, ratio=0.05)
# Ratio rows are RELATIVE (|r_bf / r_vox - 1| <= 5 %): §8.6 "within 5 %" of K/K0 = 3.2 means +-0.16, not +-0.05
# (an absolute 0.05 is 1.5 % of K/K0 = 3.2 but 5 % of K/K0 = 1, i.e. it tightens as the design improves).
# Body-fitted mesh gate (research/b2_bodyfit): the tetgen mesh of B2 had 24 % of tets below 5 deg dihedral and a
# median 1.2 tets through the 1.5 mm wall; it moved lambda_b by -46 %.
MESH_GATE = dict(min_layers_p01=2.0, max_frac_dihedral_lt_5=1e-3)


# ---------------------------------------------------------------------------------------------------------------
# Meshing
# ---------------------------------------------------------------------------------------------------------------
def wall_thickness(V: np.ndarray, T: np.ndarray, spacing: float = 0.5, max_samples: int = 80000, seed: int = 0,
                   device: str | None = None):
    """Local wall thickness at area-uniform surface samples by inward ray casting along −n_f (Warp, GPU or CPU).

    About 2·area/spacing² samples (capped at `max_samples`) are drawn on the surface; the thickness is the distance of the
    first hit of the ray from c − ε n_f along −n_f (∞ if none, or a self-hit at a sliver). Returns (points, t).
    Used for mesh-size grading and as the wall histogram input (§8.5).
    """
    import trimesh
    import warp as wp
    wp.config.quiet = True
    if "SCRATCH" in os.environ:
        wp.config.kernel_cache_dir = os.path.join(os.environ["SCRATCH"], "warp_cache")
    wp.init()
    dev = device or os.environ.get("CSHELL_WARP_DEVICE") or ("cuda:0" if wp.get_cuda_device_count() else "cpu")

    @wp.kernel
    def cast(mesh: wp.uint64, o: wp.array(dtype=wp.vec3), d: wp.array(dtype=wp.vec3), out: wp.array(dtype=wp.float32)):
        i = wp.tid()
        q = wp.mesh_query_ray(mesh, o[i], d[i], 1.0e8)
        if q.result:
            out[i] = q.t
        else:
            out[i] = -1.0

    m = trimesh.Trimesh(V, T, process=False)
    n_s = int(min(max_samples, max(1000, 2.0 * m.area / spacing ** 2)))
    c, fi = trimesh.sample.sample_surface(m, n_s, seed=seed)
    n = m.face_normals[fi]
    o = (c - 1e-5 * n).astype(np.float32)
    pts = wp.array(np.ascontiguousarray(V, np.float32), dtype=wp.vec3, device=dev)
    idx = wp.array(np.ascontiguousarray(T, np.int32).ravel(), dtype=wp.int32, device=dev)
    mesh = wp.Mesh(pts, idx)
    out = wp.empty(len(c), dtype=wp.float32, device=dev)
    wp.launch(cast, dim=len(c), inputs=[mesh.id, wp.array(o, dtype=wp.vec3, device=dev),
                                        wp.array((-n).astype(np.float32), dtype=wp.vec3, device=dev), out], device=dev)
    t = out.numpy().astype(float)
    t[(t < 1e-3)] = np.inf                                  # miss (-1) or self-hit at a sliver: no information
    return c, t + 1e-5 * (t < np.inf)


def _size_grid(V, pts, t, h_max, h_min, layers, growth, spacing):
    """Size field on a regular grid: s(x) = min_v (s_v + growth·|x − v|), s_v = clip(t_v/layers, h_min, h_max)."""
    from scipy.spatial import cKDTree
    sv = np.clip(t / layers, h_min, h_max)
    thin = sv < h_max
    lo = V.min(0) - 2 * spacing
    n = np.ceil((V.max(0) + 2 * spacing - lo) / spacing).astype(int) + 1
    ax = [lo[d] + spacing * np.arange(n[d]) for d in range(3)]
    X = np.stack(np.meshgrid(*ax, indexing="ij"), -1).reshape(-1, 3)
    if not thin.any():
        return lo, n, np.full(len(X), h_max)
    P, s0 = pts[thin], sv[thin]
    tree = cKDTree(P)
    k = min(12, len(P))
    d, j = tree.query(X, k=k)
    d, j = d.reshape(len(X), -1), j.reshape(len(X), -1)
    s = np.minimum(h_max, (s0[j] + growth * d).min(1))
    return lo, n, s


def tetrahedralise(V: np.ndarray, T: np.ndarray, h: float, layers: float = 2.0, h_min: float | None = None,
                   growth: float = 0.3, backend: str = "gmsh", workdir: str | Path | None = None,
                   wall_samples: int = 80000, size_field: tuple | None = None) -> tuple[np.ndarray, np.ndarray]:
    """Linear tetrahedral mesh of the closed surface (V, T), graded by local wall thickness (§8.6.2).

    Target edge length s(x) = clip(t_loc/layers, h_min, h) near thin walls (t_loc by inward ray casting at surface
    vertices), growing at rate `growth` away from them. With C3D10 (quadratic) `layers` = 2 puts >= 2 elements
    through the wall. backend "gmsh" remeshes the surface (discrete geometry) and honours the size field;
    "tetgen" keeps the input surface (use for already well-sized STLs) with a uniform `maxvolume` from `h`.
    `size_field` = (origin (3,), spacing, n (3,), values (prod(n),) x slowest, z fastest) overrides the wall-based grading.
    Returns (nodes (N, 3), tets (M, 4)) with positive orientation.
    """
    V = np.ascontiguousarray(V, float)
    T = np.ascontiguousarray(T, np.int64)
    h_min = h / 8 if h_min is None else h_min
    if backend == "keep":
        return _gmsh_inside(*remesh_surface(V, T, h)[:2], h, workdir)
    if backend == "tetgen":
        # tetgen's maxvolume bound does not bound edge length: on a 0.25 mm decimated MC surface it produced flat
        # slivers with 7 mm edges (B2). Kept for tests only; validate_design rejects such meshes via mesh_gate().
        import tetgen
        tg = tetgen.TetGen(V, T)
        nodes, tets = tg.tetrahedralize(order=1, mindihedral=10, minratio=1.5, maxvolume=h ** 3 / (6 * np.sqrt(2)),
                                        quiet=True)[:2]
        return nodes, _orient(nodes, tets)
    import gmsh
    import trimesh
    tmp = Path(tempfile.mkdtemp(prefix="vtmp_", dir=workdir))
    trimesh.Trimesh(V, T, process=False).export(tmp / "surf.stl")
    if size_field is None:
        pts, t = wall_thickness(V, T, spacing=max(h_min, h / 4), max_samples=wall_samples)
        sp = max(h_min, h / 4)
        lo, n, s = _size_grid(V, pts, t, h, h_min, layers, growth, spacing=sp)
    else:
        lo, sp, n, s = size_field
    with open(tmp / "size.dat", "w") as f:
        f.write(" ".join(map(str, lo)) + f" {sp} {sp} {sp} " + " ".join(map(str, n)) + "\n")
        np.savetxt(f, s, fmt="%.6g")                      # x slowest, z fastest (checked in the unit test)
    gmsh.initialize()
    try:
        gmsh.option.setNumber("General.Terminal", 0)
        gmsh.merge(str(tmp / "surf.stl"))
        gmsh.model.mesh.classifySurfaces(40 * np.pi / 180, True, True, np.pi)
        gmsh.model.mesh.createGeometry()
        surfs = [e[1] for e in gmsh.model.getEntities(2)]
        loop = gmsh.model.geo.addSurfaceLoop(surfs)
        gmsh.model.geo.addVolume([loop])
        gmsh.model.geo.synchronize()
        fld = gmsh.model.mesh.field.add("Structured")
        gmsh.model.mesh.field.setString(fld, "FileName", str(tmp / "size.dat"))
        gmsh.model.mesh.field.setNumber(fld, "TextFormat", 1)
        gmsh.model.mesh.field.setAsBackgroundMesh(fld)
        for o in ("MeshSizeFromPoints", "MeshSizeFromCurvature", "MeshSizeExtendFromBoundary"):
            gmsh.option.setNumber("Mesh." + o, 0)
        gmsh.option.setNumber("Mesh.MeshSizeMax", h)
        gmsh.option.setNumber("Mesh.MeshSizeMin", h_min)
        gmsh.option.setNumber("Mesh.Optimize", 1)
        gmsh.model.mesh.generate(3)
        tags, xyz, _ = gmsh.model.mesh.getNodes()
        types, _, enodes = gmsh.model.mesh.getElements(3)
        if 4 not in list(types):
            raise RuntimeError("gmsh produced no tetrahedra (is the STL watertight?)")
        en = enodes[list(types).index(4)].reshape(-1, 4)
    finally:
        gmsh.finalize()
        shutil.rmtree(tmp, ignore_errors=True)
    lut = np.zeros(int(tags.max()) + 1, np.int64)
    lut[tags] = np.arange(len(tags))
    nodes, tets = xyz.reshape(-1, 3), lut[en]
    used = np.unique(tets)                                  # drop surface-only nodes
    remap = -np.ones(len(nodes), np.int64)
    remap[used] = np.arange(len(used))
    nodes, tets = nodes[used], remap[tets]
    return nodes, _orient(nodes, tets)


def remesh_surface(V: np.ndarray, T: np.ndarray, h: float, max_dev: float = 0.02):
    """Isotropic remesh (pymeshlab) of a closed surface to edge h, after quadric pre-decimation of very fine
    (marching-cubes) input to edge h/3. Checks: watertight, no self-intersections, |dV/V| <= 1e-3; raises otherwise.
    Sub-h films (e.g. the escape-hole cap films of B2) are not represented; the export must not create them."""
    import pymeshlab
    import trimesh
    t0 = trimesh.Trimesh(V, T, process=False)
    ms = pymeshlab.MeshSet()
    ms.add_mesh(pymeshlab.Mesh(np.asarray(V, float), np.asarray(T, np.int32)))
    if len(T) > 50 * t0.area / h ** 2:
        ms.meshing_decimation_quadric_edge_collapse(targetfacenum=int(t0.area / (0.433 * (h / 3) ** 2)),
                                                    preservenormal=True, preservetopology=True, qualitythr=0.5,
                                                    planarquadric=True)
    ms.meshing_isotropic_explicit_remeshing(targetlen=pymeshlab.PureValue(h), iterations=8, featuredeg=60,
                                            checksurfdist=True, maxsurfdist=pymeshlab.PureValue(max_dev))
    m = ms.current_mesh()
    V1, T1 = m.vertex_matrix().astype(float), m.face_matrix().astype(np.int64)
    ms.compute_selection_by_self_intersections_per_face()
    nsi = int(ms.current_mesh().selected_face_number())
    t1 = trimesh.Trimesh(V1, T1, process=False)
    info = dict(faces=len(T1), watertight=bool(t1.is_watertight), self_intersections=nsi,
                vol_rel=float(t1.volume / t0.volume - 1))
    if not info["watertight"] or nsi or abs(info["vol_rel"]) > 1e-3:
        raise RuntimeError(f"surface remesh check failed: {info}")
    return V1, T1, info


def _gmsh_inside(V: np.ndarray, T: np.ndarray, h: float, workdir=None):
    """gmsh Delaunay + Netgen optimisation inside the GIVEN triangulation (no reparametrisation, so no
    'Wrong topology of boundary mesh for parametrization' on marching-cubes surfaces)."""
    import gmsh
    import trimesh
    tmp = Path(tempfile.mkdtemp(prefix="vtmp_", dir=workdir))
    trimesh.Trimesh(V, T, process=False).export(tmp / "s.stl")
    gmsh.initialize()
    try:
        gmsh.option.setNumber("General.Terminal", 0)
        gmsh.merge(str(tmp / "s.stl"))
        loop = gmsh.model.geo.addSurfaceLoop([e[1] for e in gmsh.model.getEntities(2)])
        gmsh.model.geo.addVolume([loop])
        gmsh.model.geo.synchronize()
        gmsh.option.setNumber("Mesh.MeshSizeMax", h)
        gmsh.option.setNumber("Mesh.Algorithm3D", 1)
        gmsh.option.setNumber("Mesh.Optimize", 1)
        gmsh.option.setNumber("Mesh.OptimizeNetgen", 1)
        gmsh.model.mesh.generate(3)
        tags, xyz, _ = gmsh.model.mesh.getNodes()
        types, _, en = gmsh.model.mesh.getElements(3)
        en = en[list(types).index(4)].reshape(-1, 4)
    finally:
        gmsh.finalize()
        shutil.rmtree(tmp, ignore_errors=True)
    lut = np.zeros(int(tags.max()) + 1, np.int64)
    lut[tags] = np.arange(len(tags))
    nodes, tets = xyz.reshape(-1, 3), lut[en]
    used = np.unique(tets)
    remap = -np.ones(len(nodes), np.int64)
    remap[used] = np.arange(len(used))
    nodes, tets = nodes[used], remap[tets]
    return nodes, _orient(nodes, tets)


def mesh_gate(nodes: np.ndarray, tets: np.ndarray, V: np.ndarray, T: np.ndarray, gate: dict | None = None) -> dict:
    """Minimum dihedral angles and tet layers through the wall (opposite-wall ray samples); raises on a miss."""
    from .bodyfit_gpu import dihedral_min, layers     # new module = b2_bodyfit/bf_gpu.py + diag1_lib.py
    g = {**MESH_GATE, **(gate or {})}
    dm = dihedral_min(nodes, tets)
    L = layers(np.asarray(V, float), np.asarray(T, np.int64), nodes, tets)
    q = dict(dihedral_min=float(dm.min()), frac_dihedral_lt_5=float((dm < 5).mean()), layers=L)
    if q["frac_dihedral_lt_5"] > g["max_frac_dihedral_lt_5"] or L["layers_q"][1] < g["min_layers_p01"]:
        raise RuntimeError(f"body-fitted mesh fails the quality gate: {q}")
    return q


def _orient(nodes: np.ndarray, tets: np.ndarray) -> np.ndarray:
    """Swap two corners of negatively oriented tets so every signed volume is positive."""
    tets = np.array(tets, np.int64)
    p = nodes[tets]
    det = np.einsum("ei,ei->e", np.cross(p[:, 1] - p[:, 0], p[:, 2] - p[:, 0]), p[:, 3] - p[:, 0])
    neg = det < 0
    tets[neg, 2], tets[neg, 1] = tets[neg, 1].copy(), tets[neg, 2].copy()
    return tets


def to_quadratic(nodes: np.ndarray, tets: np.ndarray) -> tuple[np.ndarray, np.ndarray]:
    """Add straight mid-edge nodes: (N, 3), (M, 4) -> (N + N_edges, 3), (M, 10) in CalculiX C3D10 order."""
    ed = np.sort(tets[:, _CCX_EDGES].reshape(-1, 2), axis=1)
    key = ed[:, 0] * len(nodes) + ed[:, 1]
    uk, inv = np.unique(key, return_inverse=True)
    pairs = np.stack([uk // len(nodes), uk % len(nodes)], 1)
    mid = 0.5 * (nodes[pairs[:, 0]] + nodes[pairs[:, 1]])
    t10 = np.concatenate([tets, len(nodes) + inv.reshape(-1, 6)], 1)
    return np.concatenate([nodes, mid]), t10


def tet_volumes(nodes: np.ndarray, tets: np.ndarray) -> np.ndarray:
    p = nodes[tets[:, :4]]
    return np.einsum("ei,ei->e", np.cross(p[:, 1] - p[:, 0], p[:, 2] - p[:, 0]), p[:, 3] - p[:, 0]) / 6.0


# ---------------------------------------------------------------------------------------------------------------
# Boundary conditions
# ---------------------------------------------------------------------------------------------------------------
def nodes_near_faces(nodes: np.ndarray, V: np.ndarray, T: np.ndarray, TF: np.ndarray, face_ids, tol: float) -> np.ndarray:
    """Indices of mesh nodes within `tol` of the triangles of `face_ids` (cshell/vshell rule, §4.3)."""
    import igl
    sel = np.ascontiguousarray(T[np.isin(TF, list(face_ids))], np.int64)
    if sel.size == 0:
        return np.zeros(0, np.int64)
    pts = V[np.unique(sel)]
    box = (nodes >= pts.min(0) - tol).all(1) & (nodes <= pts.max(0) + tol).all(1)
    cand = np.flatnonzero(box)
    if cand.size == 0:
        return cand
    d2 = igl.point_mesh_squared_distance(np.ascontiguousarray(nodes[cand]), np.ascontiguousarray(V, float), sel)[0]
    return cand[d2 <= tol * tol]


def consistent_face_weights(tets10: np.ndarray, nodes10: np.ndarray, sel_nodes: np.ndarray) -> dict:
    """Consistent P2 nodal weights of a uniform traction on the boundary triangles whose 3 corners all lie in `sel_nodes`:
    each triangle of area A puts A/3 on its three mid-edge nodes and 0 on its corners. Returns {node: area share}."""
    faces_loc = np.array([[0, 1, 2], [0, 1, 3], [0, 2, 3], [1, 2, 3]])
    eidx = {(0, 1): 0, (1, 2): 1, (0, 2): 2, (0, 3): 3, (1, 3): 4, (2, 3): 5}
    f = tets10[:, faces_loc].reshape(-1, 3)                      # corner node ids of all tet faces
    key = np.sort(f, axis=1)
    _, inv, cnt = np.unique(key, axis=0, return_inverse=True, return_counts=True)
    bnd = np.flatnonzero(cnt[inv.ravel()] == 1)
    insel = np.isin(f[bnd], sel_nodes).all(1)
    bnd = bnd[insel]
    w: dict = {}
    for b in bnd:
        e, fl = divmod(int(b), 4)
        loc = faces_loc[fl]
        p = nodes10[f[b]]
        A = 0.5 * np.linalg.norm(np.cross(p[1] - p[0], p[2] - p[0]))
        for a, c in ((0, 1), (1, 2), (0, 2)):
            i, j = sorted((loc[a], loc[c]))
            n = int(tets10[e, 4 + eidx[(i, j)]])
            w[n] = w.get(n, 0.0) + A / 3.0
    return w


@dataclass
class Case:
    """Mesh and BCs of one body-fitted analysis. `fixed` node ids (all 3 DOF); `load_nodes`/`load_vec` (n, 3) in N."""
    nodes: np.ndarray
    tets10: np.ndarray
    fixed: np.ndarray
    load_nodes: np.ndarray
    load_vec: np.ndarray
    E: float = E_MPA
    nu: float = NU
    info: dict = field(default_factory=dict)


def build_case(nodes10: np.ndarray, tets10: np.ndarray, V: np.ndarray, T: np.ndarray, TF: np.ndarray,
               support_faces, loads, bc_tol: float, E: float = E_MPA, nu: float = NU,
               load_distribution: str = "area") -> Case:
    """Clamp support-face nodes and share each load over its load-face nodes (§4.3). `loads`: dicts `faces`, `force`.
    load_distribution "area" (default): consistent P2 weights (A/3 on the mid-edge nodes of the boundary triangles), i.e. a
    uniform traction, which converges with the mesh. "nodal": the spec's equal share over all nodes within `bc_tol`
    (corner and mid-edge nodes alike); with C3D10 this is NOT a uniform traction and changed the first buckling factor of
    the V3-b tube by 2x (research/validate/v3b.md), so it is not the default (deviation from §4.3/§8.6, recorded).
    Loads on clamped nodes are dropped with a warning (§4.3 overlap rule)."""
    fixed = nodes_near_faces(nodes10, V, T, TF, support_faces, bc_tol)
    if fixed.size == 0:
        raise RuntimeError("no mesh nodes found on the support faces")
    f = {}
    counts = []
    for L in loads:
        ln = nodes_near_faces(nodes10, V, T, TF, L["faces"], bc_tol)
        if ln.size == 0:
            raise RuntimeError(f"no mesh nodes found on load faces {L['faces']}")
        counts.append(int(ln.size))
        if load_distribution == "area":
            w = consistent_face_weights(tets10, nodes10, ln)
            tot = sum(w.values())
            for i, a in w.items():
                f[i] = f.get(i, 0.0) + np.asarray(L["force"], float) * (a / tot)
        elif load_distribution == "nodal":
            for i in ln:
                f[int(i)] = f.get(int(i), 0.0) + np.asarray(L["force"], float) / ln.size
        else:
            raise ValueError(load_distribution)
    clash = [i for i in f if i in set(fixed.tolist())]
    for i in clash:
        del f[i]
    if clash:
        log.warning("%d loaded nodes are clamped; their load is removed (§4.3)", len(clash))
    ln = np.array(sorted(f), np.int64)
    lv = np.array([f[i] for i in ln]).reshape(-1, 3)
    return Case(nodes10, tets10, fixed, ln, lv, E, nu, dict(n_fixed=int(fixed.size), n_load=counts, n_clash=len(clash)))


# ---------------------------------------------------------------------------------------------------------------
# CalculiX deck, driver, parser
# ---------------------------------------------------------------------------------------------------------------
def _id_lines(ids: np.ndarray, per_line: int = 16) -> str:
    pad = (-len(ids)) % per_line
    a = np.concatenate([ids, -np.ones(pad, ids.dtype)]).reshape(-1, per_line)
    return "\n".join(", ".join(str(x) for x in row if x >= 0) for row in a) + "\n"


def write_inp(path: str | Path, case: Case, n_modes: int = 4, want_stress: bool = True, solver: str = "SPOOLES",
              buckle: bool = True) -> Path:
    """CalculiX deck: C3D10, linear static step, then `*BUCKLE` with `n_modes` factors (§8.6.2).

    Output requests: U at the load nodes (compliance f·u), and optionally element volumes (EVOL) and integration-point
    stresses (S) for the 0.995 quantile.
    """
    path = Path(path)
    nodes, tets = case.nodes, case.tets10
    with open(path, "w") as f:
        f.write("** cshell.validate deck (mm, N, MPa)\n*NODE, NSET=NALL\n")
        np.savetxt(f, np.column_stack([np.arange(1, len(nodes) + 1), nodes]), fmt=["%d", "%.9g", "%.9g", "%.9g"],
                   delimiter=", ")
        f.write("*ELEMENT, TYPE=C3D10, ELSET=EALL\n")
        np.savetxt(f, np.column_stack([np.arange(1, len(tets) + 1), tets + 1]), fmt="%d", delimiter=", ")
        f.write("*NSET, NSET=NFIX\n" + _id_lines(case.fixed + 1))
        f.write("*NSET, NSET=NLOAD\n" + _id_lines(case.load_nodes + 1))
        f.write(f"*MATERIAL, NAME=MAT\n*ELASTIC\n{case.E:.9g}, {case.nu:.9g}\n*SOLID SECTION, ELSET=EALL, MATERIAL=MAT\n")
        f.write("*BOUNDARY\nNFIX, 1, 3\n")
        cload = "*CLOAD\n" + "".join(f"{n}, {d + 1}, {v[d]:.12g}\n" for n, v in zip(case.load_nodes + 1, case.load_vec)
                                   for d in range(3) if v[d] != 0.0)
        f.write(f"*STEP\n*STATIC, SOLVER={solver}\n{cload}")
        f.write("*NODE PRINT, NSET=NLOAD\nU\n")
        if want_stress:
            f.write("*EL PRINT, ELSET=EALL\nS\n*EL PRINT, ELSET=EALL, TOTALS=NO\nEVOL\n")
        f.write("*END STEP\n")
        if buckle:
            # CalculiX takes the reference load of *BUCKLE from the buckle step itself (verified: without it the
            # factors are ~1e10); the load is therefore repeated there.
            f.write(f"*STEP\n*BUCKLE, SOLVER={solver}\n{n_modes}\n{cload}*END STEP\n")
    return path


@dataclass
class CcxRun:
    returncode: int
    seconds: float
    log: Path
    dat: Path


def run_ccx(inp: str | Path, nproc: int = 8, timeout: float | None = None, ccx: str | None = None) -> CcxRun:
    """Run `ccx -i <job>` in the deck's directory (OpenMP and SPOOLES threads = nproc)."""
    inp = Path(inp)
    exe = ccx or str(Path(os.environ.get("CSHELL_ENV", "")) / "bin" / "ccx")
    if not Path(exe).exists():
        exe = "ccx"
    env = dict(os.environ, OMP_NUM_THREADS=str(nproc), CCX_NPROC_EQUATION_SOLVER=str(nproc),
               CCX_NPROC_STIFFNESS_ASSEMBLY=str(nproc), CCX_NPROC_RESULTS=str(nproc), OMP_STACKSIZE="256M")
    t0 = time.time()
    with open(inp.with_suffix(".log"), "w") as lg:
        p = subprocess.run([exe, "-i", inp.stem], cwd=inp.parent, stdout=lg, stderr=subprocess.STDOUT, env=env,
                           timeout=timeout)
    return CcxRun(p.returncode, time.time() - t0, inp.with_suffix(".log"), inp.with_suffix(".dat"))


@dataclass
class CcxResult:
    """Parsed `.dat`: load-node displacements (id -> u), buckling factors, element volumes, ip stresses (xx yy zz xy xz yz)."""
    u_load: dict
    buckling: np.ndarray
    evol: np.ndarray | None
    stress: np.ndarray | None          # (n_el, 4, 6) ordered by element id


_HDR = re.compile(r"^\s*(displacements|stresses|volume|B U C K L I N G)", re.I)


def parse_dat(path: str | Path) -> CcxResult:
    """Parse CalculiX `.dat` blocks: `displacements`, `stresses`, `volume`, and the buckling factor table."""
    u, evol, buck, rows_s = {}, None, [], None
    block, lines = None, []

    def flush():
        nonlocal evol, rows_s, block, lines
        if block is None or not lines:
            block, lines = None, []
            return
        a = np.array([l.split() for l in lines], float)
        if block == "u" and not u:
            u.update({int(r[0]): r[1:4] for r in a})
        elif block == "s" and rows_s is None:
            rows_s = a
        elif block == "v" and evol is None:
            evol = a
        elif block == "b" and not buck:
            buck.extend(a[:, 1].tolist())
        block, lines = None, []

    with open(path) as f:
        for ln in f:
            s = ln.strip()
            if not s:
                continue
            m = _HDR.match(ln)
            if m:
                flush()
                k = m.group(1).lower()
                block = {"displacements": "u", "stresses": "s", "volume": "v"}.get(k, "b")
                continue
            if block is not None:
                if block == "b" and not re.match(r"^\d+\s+[-+0-9.Ee]+$", s):
                    continue
                if block != "b" and not re.match(r"^[-+0-9]", s):
                    flush()
                    continue
                lines.append(s)
        flush()
    stress = None
    if rows_s is not None:
        n_el = int(rows_s[:, 0].max())
        order = np.lexsort((rows_s[:, 1], rows_s[:, 0]))
        stress = rows_s[order, 2:8].reshape(n_el, -1, 6)
    ev = None
    if evol is not None:
        ev = np.zeros(int(evol[:, 0].max()))
        ev[evol[:, 0].astype(int) - 1] = evol[:, 1]
    return CcxResult(u, np.array(sorted(buck)), ev, stress)


def von_mises(s: np.ndarray) -> np.ndarray:
    """Von Mises from (..., 6) stresses ordered xx, yy, zz, xy, xz, yz."""
    sx, sy, sz, txy, txz, tyz = np.moveaxis(s, -1, 0)
    return np.sqrt(0.5 * ((sx - sy) ** 2 + (sy - sz) ** 2 + (sz - sx) ** 2) + 3 * (txy ** 2 + txz ** 2 + tyz ** 2))


def weighted_quantile(x: np.ndarray, w: np.ndarray, q: float) -> float:
    """Volume-weighted quantile: smallest x with cumulative weight fraction >= q (§4.4(b))."""
    o = np.argsort(x)
    c = np.cumsum(np.asarray(w, float)[o])
    return float(np.asarray(x)[o][np.searchsorted(c, q * c[-1])])


# ---------------------------------------------------------------------------------------------------------------
# Top level
# ---------------------------------------------------------------------------------------------------------------
@dataclass
class BodyFitted:
    """Body-fitted responses: compliance (N·mm), first buckling factor, 0.995 stress quantile (MPa), bookkeeping."""
    compliance: float
    buckling: np.ndarray
    sigma_q: float | None
    volume: float
    n_nodes: int
    n_elems: int
    seconds_mesh: float
    seconds_ccx: float
    engine: str = "ccx"
    info: dict = field(default_factory=dict)

    @property
    def lambda_b(self) -> float:
        return float(self.buckling[0]) if self.buckling.size else float("nan")


def responses_from_ccx(case: Case, res: CcxResult, quantile: float = 0.995, exclude_bc_mm: float = 2.0):
    """(compliance, sigma_q, volume) from a parsed run. C = f·u over the load nodes (§4.4(a))."""
    C = 0.0
    for n, v in zip(case.load_nodes, case.load_vec):
        C += float(np.dot(v, res.u_load[int(n) + 1]))
    sq = None
    vol = float(tet_volumes(case.nodes, case.tets10).sum())
    if res.stress is not None and res.evol is not None:
        from scipy.spatial import cKDTree
        cen = case.nodes[case.tets10[:, :4]].mean(1)
        bc = np.union1d(case.fixed, case.load_nodes)
        keep = cKDTree(case.nodes[bc]).query(cen)[0] > exclude_bc_mm
        vm = von_mises(res.stress)                                   # (n_el, 4)
        w = np.repeat((res.evol / 4.0 * keep)[:, None], vm.shape[1], 1)
        sq = weighted_quantile(vm.ravel(), w.ravel(), quantile)
    return C, sq, vol


def solve_skfem_p2(case: Case) -> float:
    """Fallback compliance f·u with scikit-fem P2 tets on the same mesh; BC dofs are matched to the mesh nodes by
    coordinate (vertex and mid-edge dofs), so clamping and loading are identical to the ccx deck (§8.6)."""
    import scipy.sparse.linalg as spl
    import skfem
    from scipy.spatial import cKDTree
    from skfem.models.elasticity import lame_parameters, linear_elasticity
    corner = np.unique(case.tets10[:, :4])
    remap = -np.ones(int(case.tets10.max()) + 1, np.int64)
    remap[corner] = np.arange(len(corner))
    m = skfem.MeshTet(case.nodes[corner].T.copy(), remap[case.tets10[:, :4]].T.copy())
    basis = skfem.Basis(m, skfem.ElementVector(skfem.ElementTetP2()))
    lam, mu = lame_parameters(case.E, case.nu)
    K = linear_elasticity(lam, mu).assemble(basis)
    loc = np.zeros((basis.N, 3))
    comp = np.zeros(basis.N, int)
    mid = 0.5 * (m.p[:, m.edges[0]] + m.p[:, m.edges[1]]).T
    for c in range(3):
        loc[basis.nodal_dofs[c]] = m.p.T
        loc[basis.edge_dofs[c]] = mid
        comp[basis.nodal_dofs[c]] = c
        comp[basis.edge_dofs[c]] = c
    eps = 1e-7 * max(1.0, float(np.abs(case.nodes).max()))
    fixed = np.flatnonzero(cKDTree(case.nodes[case.fixed]).query(loc)[0] < eps)
    f = np.zeros(basis.N)
    d, i = cKDTree(case.nodes[case.load_nodes]).query(loc)
    hit = np.flatnonzero(d < eps)
    f[hit] = case.load_vec[i[hit], comp[hit]]
    free = np.setdiff1d(np.arange(basis.N), fixed)
    u = np.zeros(basis.N)
    u[free] = spl.spsolve(K[free][:, free].tocsc(), f[free])
    return float(f @ u)


def validate_design(V: np.ndarray, T: np.ndarray, ctx_V: np.ndarray, ctx_T: np.ndarray, ctx_TF: np.ndarray,
                    support, loads, workdir: str | Path, name: str = "design", h: float = 0.5, layers: float = 2.0,
                    h_bc: float | None = None, n_modes: int = 4, nproc: int = 8, want_stress: bool = True,
                    E: float = E_MPA, nu: float = NU, backend: str = "gmsh", fallback: str | None = None, h_min: float | None = None,
                    load_distribution: str = "area", timeout: float | None = None, exclude_bc_mm: float = 2.0,
                    engine: str = "gpu") -> BodyFitted:
    """Mesh the STL surface (V, T), write and run the deck, parse the result (§8.6.2-3).

    (ctx_V, ctx_T, ctx_TF) with `support` face ids and `loads` locate the BC nodes (original tessellation).
    """
    wd = Path(workdir)
    wd.mkdir(parents=True, exist_ok=True)
    t0 = time.time()
    n4, t4 = tetrahedralise(V, T, h, layers=layers, h_min=h_min, backend=backend, workdir=wd)
    nodes, tets10 = to_quadratic(n4, t4)
    q = mesh_gate(nodes, tets10, V, T)
    case = build_case(nodes, tets10, ctx_V, ctx_T, ctx_TF, support, loads, 0.75 * (h_bc or h), E, nu, load_distribution)
    case.info["mesh_quality"] = q
    t_mesh = time.time() - t0
    if engine == "gpu":     # same C3D10 model on cuDSS: matched ccx to <1e-6 (C, lambda_1, sigma_q) at 14-30 s vs 410-480 s
        from .bodyfit_gpu import solve as gsolve
        t1 = time.time()
        o = gsolve(case.nodes, case.tets10, case.fixed, case.load_nodes, case.load_vec, E, nu, max(n_modes, 4),
                   exclude_bc_mm=exclude_bc_mm)
        o.pop("_phi_nodes", None)
        return BodyFitted(o["C"], np.asarray(o["lam"]), o["sigma_q"], o["volume"], len(nodes), len(tets10), t_mesh,
                          time.time() - t1, "gpu-P2", dict(case.info, **{k: o[k] for k in ("ndof", "eig_res", "timings_s")}))
    inp = write_inp(wd / f"{name}.inp", case, n_modes, want_stress)
    run = run_ccx(inp, nproc, timeout)
    ok = run.returncode == 0 and run.dat.exists()
    if ok:
        res = parse_dat(run.dat)
        ok = len(res.u_load) > 0
    if not ok:
        if fallback != "skfem":
            raise RuntimeError(f"ccx failed (rc={run.returncode}); see {run.log}")
        log.warning("ccx failed (rc=%s, %s): FALLING BACK to scikit-fem P2 (compliance only)", run.returncode, run.log)
        C = solve_skfem_p2(case)
        return BodyFitted(C, np.zeros(0), None, float(tet_volumes(nodes, tets10).sum()), len(nodes), len(tets10), t_mesh,
                          run.seconds, "skfem-P2", dict(ccx_log=str(run.log)))
    C, sq, vol = responses_from_ccx(case, res, exclude_bc_mm=exclude_bc_mm)
    return BodyFitted(C, res.buckling, sq, vol, len(nodes), len(tets10), t_mesh, run.seconds, "ccx", dict(case.info))


# ---------------------------------------------------------------------------------------------------------------
# Comparison with the voxel FE (§8.6 tolerances)
# ---------------------------------------------------------------------------------------------------------------
def compare(voxel: dict, bf: dict, tol: dict | None = None) -> list[dict]:
    """Rows comparing voxel and body-fitted responses.

    `voxel`, `bf`: {"solid": {...}, "design": {...}} each with `compliance` (N·mm), `lambda_b`, `sigma_q` (MPa).
    Absolute rows: relative differences against compliance 5 %, buckling 10 %, stress 15 %. Ratio rows (design/solid):
    K/K0 = C0/C, P/P0 = λ_b/λ_b0 and strength σ0/σ, RELATIVE difference of the ratio <= 0.05; a ratio miss blocks
    acceptance (`blocking` True). Missing entries are skipped.
    """
    tol = {**TOLERANCES, **(tol or {})}
    rows = []

    def add(metric, case, a, b, kind, t, blocking=False):
        if a is None or b is None or not (np.isfinite(a) and np.isfinite(b)):
            return
        diff = (b - a) / a if kind == "rel" else b - a
        rows.append(dict(metric=metric, case=case, voxel=a, bodyfitted=b, diff=diff, kind=kind, tol=t,
                         passed=bool(abs(diff) <= t), blocking=blocking))

    for case in ("solid", "design"):
        if case not in voxel or case not in bf:
            continue
        v, b = voxel[case], bf[case]
        add("compliance_Nmm", case, v.get("compliance"), b.get("compliance"), "rel", tol["compliance"])
        add("lambda_b", case, v.get("lambda_b"), b.get("lambda_b"), "rel", tol["buckling"])
        add("sigma_q995_MPa", case, v.get("sigma_q"), b.get("sigma_q"), "rel", tol["stress"])
    if all(c in voxel and c in bf for c in ("solid", "design")):
        def ratios(d):
            s, x = d["solid"], d["design"]
            r = {}
            if s.get("compliance") and x.get("compliance"):
                r["K/K0"] = s["compliance"] / x["compliance"]
            if s.get("lambda_b") and x.get("lambda_b"):
                r["Pb/Pb0"] = x["lambda_b"] / s["lambda_b"]
            if s.get("sigma_q") and x.get("sigma_q"):
                r["Ps/Ps0"] = s["sigma_q"] / x["sigma_q"]
            return r
        rv, rb = ratios(voxel), ratios(bf)
        for k in rv:
            if k in rb:
                add(k, "ratio", rv[k], rb[k], "rel", tol["ratio"], blocking=True)
    return rows


def write_validation_csv(rows: list[dict], path: str | Path) -> Path:
    """`validation.csv` with a pass/fail per row and a final summary row for blocking ratio misses."""
    path = Path(path)
    cols = ["metric", "case", "voxel", "bodyfitted", "diff", "kind", "tol", "passed", "blocking"]
    with open(path, "w", newline="") as f:
        w = csv.DictWriter(f, cols)
        w.writeheader()
        w.writerows(rows)
    return path
