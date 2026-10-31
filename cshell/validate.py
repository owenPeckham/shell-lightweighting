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

Engines: `engine="gpu"` (default; cshell.bodyfit_gpu, cuDSS, same C3D10 model, ccx-identical to 1e-6) or `"ccx"`.
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
# Ratio rows are RELATIVE: §8.6 "within 5 %" of K/K0 = 3.2 means +-0.16, not +-0.05 (an absolute 0.05 is 1.5 % of
# K/K0 = 3.2 but 5 % of K/K0 = 1, i.e. it tightens as the design improves). They are ONE-SIDED (Owen 2026-10-30, D3):
# a ratio blocks only when the voxel model over-predicts it, r_vox / r_bf - 1 > 5 %; a voxel under-prediction is
# conservative and only logged (B2: P_b/P_b0 voxel 1.79-1.81 at h_e 0.2-0.4 against the mesh-converged body-fitted
# 1.932, research/vf_b2/voxel_conv.py).
# Body-fitted mesh gate (research/b2_bodyfit): the tetgen mesh of B2 had 24 % of tets below 5 deg dihedral and a
# median 1.2 tets through the 1.5 mm wall; it moved lambda_b by -46 %.
# Gate numbers: the corrected B2 meshes have 0.4-1.5 deg minimum and <= 8e-5 of tets below 5 deg; the sliver mesh had
# 0.06 deg and 24 %. 0.25 deg rejects degenerate tets, 1e-3 allows isolated surface-corner slivers (a few of 6e5 tets
# moved nothing at the 1e-6 level), 2 layers is the C3D10 minimum for bending, 2.5 h bounds the edge (gmsh MeshSizeMax).
MESH_GATE = dict(min_layers_p01=2.0, max_frac_dihedral_lt_5=1e-3, min_dihedral_deg=0.25, max_edge_over_h=2.5)


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
                   wall_samples: int = 80000, size_field: tuple | None = None, remesh_kw: dict | None = None,
                   peel: bool = False, info: dict | None = None) -> tuple[np.ndarray, np.ndarray]:
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
    if backend == "stuff":                  # isosurface stuffing at L = stuff_c·h (remesh_kw: {"stuff_c": ...})
        nodes, tets, sinfo = stuff_tets(V, T, float((remesh_kw or {}).get("stuff_c", 2 ** 0.5)) * h)
        if info is not None:
            info["stuff"] = sinfo
        return nodes, tets
    if backend == "keep":
        V1, T1, rinfo = remesh_surface(V, T, h, **(remesh_kw or {}))
        nodes, tets = _gmsh_inside_robust(V1, T1, h, workdir)
        if info is not None:
            info["remesh"] = {k: v for k, v in rinfo.items() if k not in ("stages",)}
        if peel:
            nodes, tets, pinfo = peel_boundary_slivers(nodes, tets, h)
            if info is not None:
                info["peel"] = pinfo
        return nodes, tets
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


def _close(ms):
    """After deleting faces: drop non-manifold edges/vertices (they stop close_holes), then close every hole."""
    ms.meshing_remove_unreferenced_vertices()
    for _ in range(3):
        ms.meshing_repair_non_manifold_edges(method=0)
        ms.meshing_repair_non_manifold_vertices(vertdispratio=0)
        ms.meshing_remove_unreferenced_vertices()
    ms.meshing_close_holes(maxholesize=100000, newfaceselected=True)


def _self_int(V, T) -> int:
    import pymeshlab
    m = pymeshlab.MeshSet()
    m.add_mesh(pymeshlab.Mesh(V, np.asarray(T, np.int32)))
    m.compute_selection_by_self_intersections_per_face()
    return int(m.current_mesh().selected_face_number())


def _sdf_backend(V: np.ndarray, T: np.ndarray, c0: np.ndarray, device: str = "auto"):
    """Signed distance (negative inside, winding-number sign) of the closed mesh: Warp BVH on the GPU, igl otherwise
    (cshell.band backends; coordinates centred by c0 for float32)."""
    from .band import make_backend
    return make_backend(np.asarray(V, float) - c0, T, {"auto": "auto", "gpu": "warp", "cpu": "igl"}[device])


def levelset_surface(V: np.ndarray, T: np.ndarray, s: float, block: int = 16, pad: int = 3,
                     min_body_mm3: float = 1.0, device: str = "auto", log=None):
    """Watertight re-extraction of the closed surface (V, T) at lattice spacing `s` (body-fitted input, §8.6.2).

    φ = signed distance to (V, T) (winding-number sign) on the lattice nodes of every 16³-cell block within 2 cells of
    the surface (the only blocks that can hold a crossing), +3s elsewhere; |φ| ≥ 1e-3 s (crossings strictly inside the
    lattice edges); ambiguous lattice faces removed (`export.resolve_ambiguous_faces`, the export's rule); marching
    cubes per block (skimage Lewiner), vertices welded by lattice-edge key and placed in float64 at the linear
    crossing of φ (identical on both sides of a block face). The result is closed and 2-manifold by construction and
    does not self-intersect; it smooths features below ~s (sub-s slits and knife-edge tips). Bodies (connected
    components) with |V| < min_body_mm3 are dropped and reported; orientation is outward.
    Cost: one SDF query per node of the active blocks (~3 A/(s²·16) blocks · 17³ nodes). Returns (V1, T1, info)."""
    from skimage.measure import marching_cubes
    from .export import resolve_ambiguous_faces
    log = log or (lambda m: None)
    t0 = time.time()
    V = np.asarray(V, float)
    T = np.asarray(T, np.int64)
    vmin, vmax = V.min(0), V.max(0)
    c0 = 0.5 * (vmin + vmax)
    lo = vmin - pad * s + s * np.array([0.1234567, 0.2345671, 0.3456712])   # generic offset: no node on a mesh plane
    n = np.ceil((vmax + pad * s - lo) / s).astype(np.int64) + 1          # nodes per axis
    nb = -(-(n - 1) // block)                                            # blocks per axis
    # surface samples: every vertex, plus the vertices of long triangles bisected down to edges <= s/2
    cells = [np.unique(np.floor((V - lo) / s).astype(np.int64) @ np.array([n[1] * n[2], n[2], 1]))]
    for a in range(0, len(T), 1 << 22):
        P3 = V[T[a:a + (1 << 22)]]
        P3 = P3[np.linalg.norm(P3[:, [1, 2, 0]] - P3, axis=2).max(1) > 0.5 * s]
        while len(P3):
            e = np.linalg.norm(P3[:, [1, 2, 0]] - P3, axis=2)
            big = e.max(1) > 0.5 * s
            P3, e = P3[big], e[big]
            if not len(P3):
                break
            i = e.argmax(1)
            r = np.arange(len(P3))
            p, q, o = P3[r, i], P3[r, (i + 1) % 3], P3[r, (i + 2) % 3]
            m = 0.5 * (p + q)
            cells.append(np.unique(np.floor((m - lo) / s).astype(np.int64) @ np.array([n[1] * n[2], n[2], 1])))
            P3 = np.concatenate([np.stack([p, m, o], 1), np.stack([m, q, o], 1)])
    c = np.unique(np.concatenate(cells))
    c = np.stack(np.unravel_index(c, tuple(n)), 1)
    blk = set()
    for off in np.array(np.meshgrid([-2, 2], [-2, 2], [-2, 2], indexing="ij")).reshape(3, -1).T:
        b = np.clip((c + off) // block, 0, nb - 1)
        blk.update(np.unique(b @ np.array([nb[1] * nb[2], nb[2], 1])).tolist())
    blk = np.array(sorted(blk), np.int64)
    bijk = np.stack(np.unravel_index(blk, tuple(nb)), 1)
    del c, cells
    log(f"[levelset] s = {s:.4f} mm, lattice {tuple(n)}, {len(blk)} active blocks of {block}^3 cells "
        f"({len(T)} input faces), {time.time() - t0:.1f} s")
    # signed distance on the nodes of the active blocks
    be = _sdf_backend(V, T, c0, device)
    BIG = np.float32(3 * s)
    phi = np.full(tuple(n), BIG, np.float32)
    loc = np.stack(np.meshgrid(*[np.arange(block + 1)] * 3, indexing="ij"), -1).reshape(-1, 3)
    per = max(1, (1 << 23) // len(loc))
    nq = 0
    for a in range(0, len(blk), per):
        g = (bijk[a:a + per, None, :] * block + loc[None]).reshape(-1, 3)
        g = g[(g < n).all(1)]
        d = be.query(lo + g * s - c0, signed=True)[0]
        phi[g[:, 0], g[:, 1], g[:, 2]] = d
        nq += len(g)
    del be
    t_sdf = time.time() - t0
    eps = np.float32(1e-3 * s)
    small = np.abs(phi) < eps
    phi[small] = np.where(phi[small] < 0, -eps, eps)
    n_flip, n_it = resolve_ambiguous_faces(phi, float(2 * eps))
    # marching cubes per active block, weld by lattice-edge key
    PV, PF, PK, nv = [], [], [], 0
    for b in bijk:
        o = b * block
        sub = phi[o[0]:o[0] + block + 1, o[1]:o[1] + block + 1, o[2]:o[2] + block + 1]
        if not (sub.min() < 0 < sub.max()) or min(sub.shape) < 2:
            continue
        v, f, _, _ = marching_cubes(sub, 0.0, allow_degenerate=False)
        v = v.astype(np.float64)
        fr = np.abs(v - np.rint(v))
        axis = fr.argmax(1)
        lat = np.rint(v).astype(np.int64)
        r = np.arange(len(v))
        lat[r, axis] = np.floor(v[r, axis]).astype(np.int64)
        lat += o
        on_edge = (fr > 1e-6).sum(1) == 1
        key = ((lat[:, 0] * n[1] + lat[:, 1]) * n[2] + lat[:, 2]) * 3 + axis
        key[~on_edge] = -1 - (nv + r[~on_edge])                     # MC33 cube-centre vertex: never welded
        pos = lo + (v + o) * s
        PV.append(pos), PF.append(f.astype(np.int64) + nv), PK.append(key)
        nv += len(v)
    if not PV:
        raise RuntimeError("levelset_surface: no zero crossing (empty or inside-out mesh?)")
    pos, F, key = np.concatenate(PV), np.concatenate(PF), np.concatenate(PK)
    del PV, PF, PK
    uk, first, inv = np.unique(key, return_index=True, return_inverse=True)
    V1, T1 = pos[first], inv[F]
    # exact float64 crossing on lattice edges (both blocks then give the same point)
    e = uk >= 0
    k = uk[e]
    ax = k % 3
    nid = k // 3
    i0 = np.stack(np.unravel_index(nid, tuple(n)), 1)
    i1 = i0.copy()
    i1[np.arange(len(i1)), ax] += 1
    f0 = phi[i0[:, 0], i0[:, 1], i0[:, 2]].astype(np.float64)
    f1 = phi[i1[:, 0], i1[:, 1], i1[:, 2]].astype(np.float64)
    t = f0 / (f0 - f1)
    x = (lo + i0 * s).astype(np.float64)
    x[np.arange(len(x)), ax] += t * s
    V1[e] = x
    del phi
    T1 = T1[(T1[:, 0] != T1[:, 1]) & (T1[:, 1] != T1[:, 2]) & (T1[:, 0] != T1[:, 2])]
    # bodies: drop components below min_body_mm3, orientation outward
    import scipy.sparse as sps
    from scipy.sparse.csgraph import connected_components
    nV = len(V1)
    A = sps.coo_matrix((np.ones(2 * len(T1)), (T1[:, [0, 1]].ravel(), T1[:, [1, 2]].ravel())), shape=(nV, nV))
    ncomp, lab = connected_components(A, directed=False)
    a, b, cc = V1[T1[:, 0]], V1[T1[:, 1]], V1[T1[:, 2]]
    vol_f = np.einsum("ij,ij->i", a, np.cross(b, cc)) / 6.0
    flab = lab[T1[:, 0]]
    vol_c = np.bincount(flab, vol_f, ncomp)
    sgn = 1.0 if vol_c.sum() > 0 else -1.0
    keep_c = np.abs(vol_c) >= min_body_mm3
    dropped = vol_c[~keep_c & (np.bincount(flab, minlength=ncomp) > 0)] * sgn
    T1 = T1[keep_c[flab]]
    if sgn < 0:
        T1 = T1[:, [0, 2, 1]]
    used = np.unique(T1)
    remap = -np.ones(nV, np.int64)
    remap[used] = np.arange(len(used))
    V1, T1 = V1[used], remap[T1]
    vol_c = vol_c[keep_c] * sgn
    # closed 2-manifold check: every directed edge exactly once and its reverse present
    E = np.concatenate([T1[:, [0, 1]], T1[:, [1, 2]], T1[:, [2, 0]]])
    ek = E[:, 0] * len(V1) + E[:, 1]
    rk = E[:, 1] * len(V1) + E[:, 0]
    us = np.unique(ek)
    info = dict(spacing=float(s), lattice=[int(x) for x in n], active_blocks=int(len(blk)), sdf_queries=int(nq),
                faces=int(len(T1)), vertices=int(len(V1)), bodies=int(keep_c.sum()),
                body_volumes=sorted([float(x) for x in vol_c], key=abs, reverse=True)[:8],
                dropped_bodies=int(len(dropped)), dropped_mm3=float(np.abs(dropped).sum()),
                volume=float(vol_c.sum()), ambiguous_flipped=int(n_flip), ambiguous_iters=int(n_it),
                duplicate_directed_edges=int(len(ek) - len(us)),
                missing_reverse_edges=int((~np.isin(rk, us)).sum()), seconds_sdf=t_sdf, seconds=time.time() - t0)
    info["closed_manifold"] = info["duplicate_directed_edges"] == 0 and info["missing_reverse_edges"] == 0
    log(f"[levelset] {info}")
    return V1, T1, info


# Isosurface stuffing (Labelle & Shewchuk 2007, uniform BCC lattice): warp thresholds of the long (L) and short
# (L·√3/2) lattice edges; with these the output dihedral angles are bounded in [10.7°, 164.8°] in exact arithmetic
_ALPHA_LONG, _ALPHA_SHORT = 0.24697, 0.21918
# prism vertex permutations bringing vertex m (0-2 bottom, 3-5 top; vertical edges i - i+3) to position 0
_PRISM_PERM = np.array([[m % 3, (m + 1) % 3, (m + 2) % 3, m % 3 + 3, (m + 1) % 3 + 3, (m + 2) % 3 + 3] if m < 3 else
                        [m, (m - 2) % 3 + 3, (m - 1) % 3 + 3, m - 3, (m - 2) % 3, (m - 1) % 3] for m in range(6)])


def stuff_tets(V: np.ndarray, T: np.ndarray, L: float, device: str = "auto", pad: int = 2, chunk: int = 1 << 23,
               log=None):
    """Tetrahedral mesh of the solid bounded by the closed surface (V, T) by isosurface stuffing on a BCC lattice of
    spacing L (Labelle & Shewchuk 2007; body-fitted input of §8.6.2 for the production parts, research/bodyfit_scale).

    φ = signed distance to (V, T) (winding-number sign; Warp on the GPU) at the lattice nodes (cube corners and
    centres). Nodes with a cut point closer than α·|e| on an incident edge are snapped onto it (φ ← 0); every lattice
    tet with a node inside is kept whole (no node outside) or clipped by the linear zero set of φ on its edges, the
    clipped prisms/pyramids being split with diagonals from the lowest-index vertex of each quadrilateral face (the same
    choice on both sides of a lattice face: conforming; a prism always triangulates, Dompierre et al. 1999).
    The boundary interpolates the surface at lattice resolution (vertices on it, chord error O(L² κ)); features
    thinner than ~L are not resolved. No surface remesh, no Delaunay: closed by construction, dihedral angles bounded.
    Returns (nodes (N, 3), tets (M, 4) positively oriented, info)."""
    log = log or (lambda m: None)
    t0 = time.time()
    V = np.asarray(V, float)
    T = np.asarray(T, np.int64)
    vmin, vmax = V.min(0), V.max(0)
    c0 = 0.5 * (vmin + vmax)
    lo = vmin - pad * L + L * np.array([0.1234567, 0.2345671, 0.3456712])
    n = tuple(int(x) for x in np.ceil((vmax + pad * L - lo) / L).astype(int) + 1)   # black (corner) nodes per axis
    m = tuple(x - 1 for x in n)                                                        # red (centre) nodes per axis
    Nb, Nr = int(np.prod(n)), int(np.prod(m))
    be = _sdf_backend(V, T, c0, device)

    def query(shape, off):
        out = np.empty(int(np.prod(shape)), np.float32)
        for a in range(0, len(out), chunk):
            idx = np.stack(np.unravel_index(np.arange(a, min(a + chunk, len(out))), shape), 1)
            out[a:a + len(idx)] = be.query(lo + (idx + off) * L - c0, signed=True)[0]
        return out.reshape(shape)

    pb, pr = query(n, 0.0), query(m, 0.5)
    del be
    t_sdf = time.time() - t0
    eps = np.float32(1e-6 * L)
    for p in (pb, pr):
        z = np.abs(p) < eps
        p[z] = np.where(p[z] < 0, -eps, eps)
    # --- warp: nearest violating cut point per node -----------------------------------------------------------------
    db = np.full(n, np.inf, np.float32)
    dr = np.full(m, np.inf, np.float32)
    vb = np.zeros(n + (3,), np.float32)
    vr = np.zeros(m + (3,), np.float32)

    def edge(pa, pc, sa, sc, e, alpha, da, dc, va, vc):
        """Edges from nodes pa[sa] to pc[sc] (vector e): update the nearest violating cut of both ends in place."""
        fa, fc = pa[sa], pc[sc]
        cut = (fa < 0) != (fc < 0)
        t = np.where(cut, fa / np.where(cut, fa - fc, 1), 2.0)
        ln = float(np.linalg.norm(e))
        for tt, d, vv, sl, sgn in ((t, da, va, sa, 1.0), (1 - t, dc, vc, sc, -1.0)):
            ok = cut & (tt < alpha) & (tt * ln < d[sl])
            d[sl] = np.where(ok, tt * ln, d[sl])
            vv[sl] = np.where(ok[..., None], (sgn * tt)[..., None] * e, vv[sl])

    S = slice(None)
    for ax in range(3):
        a0 = tuple(slice(0, -1) if i == ax else S for i in range(3))
        a1 = tuple(slice(1, None) if i == ax else S for i in range(3))
        e = np.eye(3)[ax] * L
        edge(pb, pb, a0, a1, e, _ALPHA_LONG, db, db, vb, vb)
        edge(pr, pr, a0, a1, e, _ALPHA_LONG, dr, dr, vr, vr)
    for c in np.ndindex(2, 2, 2):
        sb = tuple(slice(c[i], c[i] + m[i]) for i in range(3))
        e = (np.array(c) - 0.5) * L                                 # red centre -> black corner
        edge(pr, pb, (S, S, S), sb, e, _ALPHA_SHORT, dr, db, vr, vb)
    snap_b, snap_r = np.isfinite(db), np.isfinite(dr)
    pb[snap_b] = 0.0
    pr[snap_r] = 0.0
    del db, dr
    sb_, sr_ = np.sign(pb).astype(np.int8), np.sign(pr).astype(np.int8)
    # --- lattice tets around every black edge with an inside node ---------------------------------------------------
    ring = ((0, 0), (-1, 0), (-1, -1), (0, -1))
    tets = []
    for ax in range(3):
        b_, c_ = [i for i in range(3) if i != ax]
        # black edge (i) -> (i + e_ax); valid where the 4 surrounding cells exist: index_b, index_c in 1..n-2
        sl0 = [slice(None)] * 3
        sl0[ax] = slice(0, n[ax] - 1)
        sl0[b_] = slice(1, n[b_] - 1)
        sl0[c_] = slice(1, n[c_] - 1)
        sl1 = list(sl0)
        sl1[ax] = slice(1, n[ax])
        neg = (sb_[tuple(sl0)] < 0) | (sb_[tuple(sl1)] < 0)
        cells = []
        for d in ring:
            slr = [None] * 3
            slr[ax] = slice(0, m[ax])
            slr[b_] = slice(1 + d[0], n[b_] - 1 + d[0])
            slr[c_] = slice(1 + d[1], n[c_] - 1 + d[1])
            cells.append(tuple(slr))
            neg |= sr_[tuple(slr)] < 0
        ii = np.stack(np.nonzero(neg), 1)
        base = ii.copy()
        base[:, b_] += 1
        base[:, c_] += 1
        e1 = base.copy()
        e1[:, ax] += 1
        idb0 = np.ravel_multi_index(base.T, n)
        idb1 = np.ravel_multi_index(e1.T, n)
        rid = []
        for d in ring:
            r = base.copy()
            r[:, b_] += d[0]
            r[:, c_] += d[1]
            rid.append(Nb + np.ravel_multi_index(r.T, m))
        for k in range(4):
            tets.append(np.stack([idb0, idb1, rid[k], rid[(k + 1) % 4]], 1))
    tets = np.concatenate(tets)
    # --- node arrays (warped positions, post-warp φ) ------------------------------------------------------------------
    used = np.unique(tets)
    isb = used < Nb
    pos = np.empty((len(used), 3))
    phi = np.empty(len(used))
    ib = np.stack(np.unravel_index(used[isb], n), 1)
    ir = np.stack(np.unravel_index(used[~isb] - Nb, m), 1)
    pos[isb] = lo + ib * L + vb[ib[:, 0], ib[:, 1], ib[:, 2]]
    pos[~isb] = lo + (ir + 0.5) * L + vr[ir[:, 0], ir[:, 1], ir[:, 2]]
    phi[isb] = pb[ib[:, 0], ib[:, 1], ib[:, 2]]
    phi[~isb] = pr[ir[:, 0], ir[:, 1], ir[:, 2]]
    del pb, pr, vb, vr, sb_, sr_
    lut = np.searchsorted(used, tets)                 # local node ids (sorted = global order kept)
    sg = np.sign(phi).astype(np.int8)[lut]
    nm, nz, npos = (sg < 0).sum(1), (sg == 0).sum(1), (sg > 0).sum(1)
    keep = nm > 0
    lut, sg, nm, nz, npos = lut[keep], sg[keep], nm[keep], nz[keep], npos[keep]
    order = np.argsort(sg, axis=1, kind="stable")      # − first, then 0, then +
    q = np.take_along_axis(lut, order, 1)
    nv = len(used)
    # cut points on (−, +) edges, one per lattice edge
    pairs = []
    for i in range(4):
        for j in range(4):
            if i < j:
                pairs.append((i, j))
    ends = []
    for i, j in pairs:
        mm = (np.take_along_axis(sg, order, 1)[:, i] < 0) & (np.take_along_axis(sg, order, 1)[:, j] > 0)
        ends.append(np.stack([q[mm, i], q[mm, j]], 1))
    ends = np.concatenate(ends)
    ckey = np.unique(ends[:, 0] * nv + ends[:, 1])
    ca, cb = ckey // nv, ckey % nv
    tc = phi[ca] / (phi[ca] - phi[cb])
    cpos = pos[ca] + tc[:, None] * (pos[cb] - pos[ca])

    def cut(a, b):
        return nv + np.searchsorted(ckey, a * nv + b)

    out = [q[npos == 0]]
    sel = (nm == 1) & (nz == 0)
    a = q[sel]
    out.append(np.stack([a[:, 0], cut(a[:, 0], a[:, 1]), cut(a[:, 0], a[:, 2]), cut(a[:, 0], a[:, 3])], 1))
    sel = (nm == 1) & (nz == 1) & (npos == 2)
    a = q[sel]
    out.append(np.stack([a[:, 0], a[:, 1], cut(a[:, 0], a[:, 2]), cut(a[:, 0], a[:, 3])], 1))
    sel = (nm == 1) & (nz == 2) & (npos == 1)
    a = q[sel]
    out.append(np.stack([a[:, 0], a[:, 1], a[:, 2], cut(a[:, 0], a[:, 3])], 1))
    prisms = []
    sel = (nm == 2) & (npos == 2)
    a = q[sel]
    prisms.append(np.stack([a[:, 0], cut(a[:, 0], a[:, 2]), cut(a[:, 0], a[:, 3]),
                            a[:, 1], cut(a[:, 1], a[:, 2]), cut(a[:, 1], a[:, 3])], 1))
    sel = (nm == 3) & (npos == 1)
    a = q[sel]
    prisms.append(np.stack([a[:, 0], a[:, 1], a[:, 2],
                            cut(a[:, 0], a[:, 3]), cut(a[:, 1], a[:, 3]), cut(a[:, 2], a[:, 3])], 1))
    P = np.concatenate(prisms)
    if len(P):
        P = np.take_along_axis(P, _PRISM_PERM[P.argmin(1)], 1)
        d15 = np.minimum(P[:, 1], P[:, 5]) < np.minimum(P[:, 2], P[:, 4])
        t1 = np.where(d15[:, None], P[:, [0, 1, 2, 5]], P[:, [0, 1, 2, 4]])
        t2 = np.where(d15[:, None], P[:, [0, 1, 5, 4]], P[:, [0, 4, 2, 5]])
        out += [t1, t2, P[:, [0, 4, 5, 3]]]
    sel = (nm == 2) & (nz == 1) & (npos == 1)
    a = q[sel]
    B = np.stack([a[:, 0], a[:, 1], cut(a[:, 1], a[:, 3]), cut(a[:, 0], a[:, 3])], 1)    # quad (a, b, c_b, c_a)
    z = a[:, 2]
    even = np.isin(B.argmin(1), (0, 2))
    out.append(np.where(even[:, None], np.stack([B[:, 0], B[:, 1], B[:, 2], z], 1),
                        np.stack([B[:, 1], B[:, 2], B[:, 3], z], 1)))
    out.append(np.where(even[:, None], np.stack([B[:, 0], B[:, 2], B[:, 3], z], 1),
                        np.stack([B[:, 1], B[:, 3], B[:, 0], z], 1)))
    tets = np.concatenate(out)
    nodes = np.concatenate([pos, cpos])
    u2 = np.unique(tets)
    remap = -np.ones(len(nodes), np.int64)
    remap[u2] = np.arange(len(u2))
    nodes, tets = nodes[u2], remap[tets]
    tets = _orient(nodes, tets)
    vol = tet_volumes(nodes, tets)
    info = dict(L=float(L), lattice=list(n), sdf_queries=Nb + Nr, snapped=int(snap_b.sum() + snap_r.sum()),
                cut_points=int(len(ckey)), n_tets=int(len(tets)), n_nodes=int(len(nodes)), volume=float(vol.sum()),
                min_tet_volume=float(vol.min()), seconds_sdf=t_sdf, seconds=time.time() - t0)
    log(f"[stuff] {info}")
    return nodes, tets, info


def boundary_faces(tets: np.ndarray) -> np.ndarray:
    """Faces of a tet mesh used by one tet, outward for positively oriented tets."""
    F = tets[:, [[1, 2, 3], [0, 3, 2], [0, 1, 3], [0, 2, 1]]].reshape(-1, 3)
    _, inv, cnt = np.unique(np.sort(F, 1), axis=0, return_inverse=True, return_counts=True)
    return F[cnt[inv.ravel()] == 1]


def surface_deviation(V: np.ndarray, T: np.ndarray, Vr: np.ndarray, Tr: np.ndarray, device: str = "auto",
                      max_ref_points: int = 50_000_000) -> dict:
    """Two-sided distance between a body-fitted surface (V, T) and the reference (exported) surface (Vr, Tr):
    `to_ref` from the vertices and face centroids of (V, T) to the reference, `from_ref` from the reference vertices
    (all, or a uniform subsample of max_ref_points) to (V, T); max, 99.9/99 % quantiles, mean (mm), and the volume
    difference V/V_ref − 1. Hausdorff = max of the two maxima."""
    def vol(X, F):
        return float(np.einsum("ij,ij->", X[F[:, 0]], np.cross(X[F[:, 1]], X[F[:, 2]]))) / 6.0

    V, Vr = np.asarray(V, float), np.asarray(Vr, float)
    c0 = 0.5 * (Vr.min(0) + Vr.max(0))
    out = dict(volume_rel=vol(V, T) / vol(Vr, Tr) - 1.0)
    be = _sdf_backend(Vr, Tr, c0, device)
    P = np.concatenate([V[np.unique(T)], V[T].mean(1)])
    d = np.abs(be.query(P - c0, signed=False)[0].astype(float))
    del be
    out["to_ref"] = dict(max=float(d.max()), q999=float(np.quantile(d, .999)), q99=float(np.quantile(d, .99)),
                         mean=float(d.mean()), argmax=[round(float(x), 3) for x in P[d.argmax()]])
    be = _sdf_backend(V, T, c0, device)
    P = Vr if len(Vr) <= max_ref_points else Vr[np.random.default_rng(0).choice(len(Vr), max_ref_points, False)]
    d = np.abs(be.query(P - c0, signed=False)[0].astype(float))
    del be
    out["from_ref"] = dict(max=float(d.max()), q999=float(np.quantile(d, .999)), q99=float(np.quantile(d, .99)),
                           mean=float(d.mean()), argmax=[round(float(x), 3) for x in P[d.argmax()]], points=int(len(P)))
    out["hausdorff"] = max(out["to_ref"]["max"], out["from_ref"]["max"])
    return out


def _thin_faces(V: np.ndarray, T: np.ndarray, d: float, rings: int = 2, device: str | None = None) -> np.ndarray:
    """Faces whose inward wall (ray along −n) or outward gap (ray along +n) from the centroid is shorter than d,
    dilated by `rings` vertex rings (Warp ray casts, GPU or CPU)."""
    import warp as wp
    import scipy.sparse as sps
    wp.config.quiet = True
    if "SCRATCH" in os.environ:
        wp.config.kernel_cache_dir = os.path.join(os.environ["SCRATCH"], "warp_cache")
    wp.init()
    dev = device or os.environ.get("CSHELL_WARP_DEVICE") or ("cuda:0" if wp.get_cuda_device_count() else "cpu")

    @wp.kernel
    def cast(mesh: wp.uint64, o: wp.array(dtype=wp.vec3), dr: wp.array(dtype=wp.vec3), out: wp.array(dtype=wp.float32)):
        i = wp.tid()
        q = wp.mesh_query_ray(mesh, o[i], dr[i], 1.0e8)
        if q.result:
            out[i] = q.t
        else:
            out[i] = 1.0e8

    c0 = V.mean(0)
    X = V[T] - c0
    n = np.cross(X[:, 1] - X[:, 0], X[:, 2] - X[:, 0])
    n /= np.maximum(np.linalg.norm(n, axis=1, keepdims=True), 1e-300)
    c = X.mean(1)
    mesh = wp.Mesh(wp.array(np.ascontiguousarray(V - c0, np.float32), dtype=wp.vec3, device=dev),
                   wp.array(np.ascontiguousarray(T, np.int32).ravel(), dtype=wp.int32, device=dev))
    t = []
    for sgn in (-1.0, 1.0):
        out = wp.empty(len(T), dtype=wp.float32, device=dev)
        wp.launch(cast, dim=len(T), inputs=[mesh.id, wp.array((c + sgn * 1e-5 * n).astype(np.float32), dtype=wp.vec3,
                                                              device=dev),
                                            wp.array((sgn * n).astype(np.float32), dtype=wp.vec3, device=dev), out],
                  device=dev)
        t.append(out.numpy())
    mask = np.minimum(*t) < d
    if rings and mask.any():
        nV = len(V)
        A = sps.csr_matrix((np.ones(3 * len(T)), (np.repeat(np.arange(len(T)), 3), T.ravel())), shape=(len(T), nV))
        for _ in range(rings):
            vm = (A.T @ mask.astype(float)) > 0
            mask = (A @ vm.astype(float)) > 0
    return mask


def peel_boundary_slivers(nodes: np.ndarray, tets: np.ndarray, h: float, min_deg: float = 5.0,
                          max_edge: float = 2.5, rounds: int = 5):
    """Remove boundary tets with >= 2 faces on the boundary whose minimum dihedral is below min_deg or whose longest
    edge exceeds max_edge·h (flat caps left by the Delaunay pass where four surface nodes are nearly coplanar or span a
    concave fold). Removing such a tet only replaces two boundary faces by the other two (the surface moves by the
    cap height); the volume removed is reported. Returns (nodes, tets, info)."""
    from .bodyfit_gpu import dihedral_min
    removed, vol = 0, 0.0
    for _ in range(rounds):
        F = np.sort(tets[:, [[1, 2, 3], [0, 2, 3], [0, 1, 3], [0, 1, 2]]], axis=2).reshape(-1, 3)
        _, inv, cnt = np.unique(F, axis=0, return_inverse=True, return_counts=True)
        nb = (cnt[inv.ravel()] == 1).reshape(-1, 4).sum(1)
        X = nodes[tets]
        L = np.max([np.linalg.norm(X[:, i] - X[:, j], axis=1) for i in range(4) for j in range(i + 1, 4)], axis=0)
        bad = (nb >= 2) & ((dihedral_min(nodes, tets) < min_deg) | (L > max_edge * h))
        if not bad.any():
            break
        removed += int(bad.sum())
        vol += float(tet_volumes(nodes, tets[bad]).sum())
        tets = tets[~bad]
    used = np.unique(tets)
    remap = -np.ones(len(nodes), np.int64)
    remap[used] = np.arange(len(used))
    return nodes[used], remap[tets], dict(peeled=removed, peeled_mm3=vol)


def remesh_surface(V: np.ndarray, T: np.ndarray, h: float, max_dev: float = 0.02, repair_iters: int = 4,
                   max_vol_rel: float = 5e-3, isolate: bool = True, thin: float | None = None, fine: float | None = None):
    """`_remesh_surface` in a fresh (spawned) process. In one process the second call of pymeshlab's isotropic remesher
    silently does nothing (B2 design: 929k faces instead of 107k at h = 0.354 after a first call at 0.5; same input
    arrays, copies made; research/vf_b2/meshgate2.py), so the two mesh levels of a collect run would be garbage."""
    if not isolate:
        return _remesh_surface(V, T, h, max_dev, repair_iters, max_vol_rel, thin, fine)
    import concurrent.futures as cf
    import multiprocessing as mp
    with cf.ProcessPoolExecutor(1, mp_context=mp.get_context("spawn")) as ex:
        return ex.submit(_remesh_surface, np.asarray(V, float), np.asarray(T, np.int64), h, max_dev, repair_iters,
                         max_vol_rel, thin, fine).result()


def _remesh_surface(V: np.ndarray, T: np.ndarray, h: float, max_dev: float = 0.02, repair_iters: int = 4,
                    max_vol_rel: float = 5e-3, thin: float | None = None, fine: float | None = None):
    """Isotropic remesh (pymeshlab) of a closed surface to edge h, after quadric pre-decimation of very fine
    (marching-cubes) input to edge h/3, then repair of self-intersections and projection back onto the input.

    Repair (B2, research/b2_bodyfit/mesh2.py): faces that self-intersect (sub-h films) are deleted with 3 rings, the hole
    is closed and re-remeshed; every vertex is then projected onto the INPUT surface (removes the flat hole-fill
    deviation), and where the projection re-creates a self-intersection the unprojected vertices of those faces are kept.
    Checks: watertight, no self-intersections, |dV/V| <= max_vol_rel (5e-3: remeshing shrinks curved surfaces by
    ~h^2 kappa^2, B2 solid -1.2e-3 at h = 0.5 and -2.8e-3 at 0.35, the same bias in solid and design), longest edge <= 2.5 h; raises otherwise (a failed gate, never a silent fallback).
    Returns (V1, T1, info)."""
    import igl
    import pymeshlab
    import trimesh
    # COPIES: pymeshlab/trimesh work in place on the buffers they are given; a second call on the same input arrays
    # (two mesh levels in one process) then started from the first call's modified vertices (B2: 925k-face garbage)
    t_in = trimesh.Trimesh(np.array(V, float), np.array(T, np.int64), process=False)
    t_in.merge_vertices()                    # duplicated vertices (unwelded STL) make the remesher a silent no-op
    V, T = np.array(t_in.vertices, float, order='C'), np.array(t_in.faces, np.int64, order='C')
    t0 = trimesh.Trimesh(V.copy(), T.copy(), process=False)
    ms = pymeshlab.MeshSet()
    ms.add_mesh(pymeshlab.Mesh(V.copy(), T.astype(np.int32)))
    if len(T) > 50 * t0.area / h ** 2:
        ms.meshing_decimation_quadric_edge_collapse(targetfacenum=int(t0.area / (0.433 * (h / 3) ** 2)),
                                                    preservenormal=True, preservetopology=True, qualitythr=0.5,
                                                    planarquadric=True)
    e0 = np.linalg.norm(V[T[:, [1, 2, 0]]] - V[T], axis=2).max()
    if e0 > 1.5 * h:      # coarse input (STEP tessellation, 248 faces on B2): split long edges first, the remesher then
        ms.meshing_surface_subdivision_midpoint(iterations=12, threshold=pymeshlab.PureValue(h))   # starts from a resolved surface
    n_thin, repairs_pre = 0, None
    if thin:
        # Thickness-aware remesh (level-set input, research/bodyfit_scale): faces whose inward wall or outward gap is
        # below thin*h (knife-edge rims, thin walls, narrow slits) are NOT remeshed at h - triangles of edge h there
        # cross the opposite sheet (the self-intersections of the production designs, whose delete/close repair then
        # broke watertightness). They are remeshed at `fine` (the lattice spacing), the rest at h.
        # Adaptive: where the h-pass still self-intersects (oblique gaps the normal rays miss, region borders), the
        # input faces within 2h of the intersecting faces join the fine set and the remesh restarts from the input
        # (up to 4 attempts; the delete/close repair below stays as the last resort).
        from scipy.spatial import cKDTree
        mask = _thin_faces(V, T, thin * h)
        cen_in = V[T].mean(1)
        thin_attempts = []
        for attempt in range(4):
            n_thin = int(mask.sum())
            cen_thin = cen_in[mask]
            ms = pymeshlab.MeshSet()
            ms.add_mesh(pymeshlab.Mesh(V.copy(), T.astype(np.int32), f_scalar_array=mask.astype(float)))
            ms.compute_selection_by_condition_per_face(condselect="fq < 0.5")
            ms.meshing_isotropic_explicit_remeshing(targetlen=pymeshlab.PureValue(h), iterations=8, featuredeg=60,
                                                    selectedonly=True, checksurfdist=True,
                                                    maxsurfdist=pymeshlab.PureValue(max_dev))
            if n_thin:
                m = ms.current_mesh()
                Vc, Tc = m.vertex_matrix().astype(float), m.face_matrix().astype(np.int64)
                sel = cKDTree(cen_thin).query(Vc[Tc].mean(1))[0] < 0.5 * (fine or h / 3)
                ms = pymeshlab.MeshSet()
                ms.add_mesh(pymeshlab.Mesh(Vc, Tc.astype(np.int32), f_scalar_array=sel.astype(float)))
                ms.compute_selection_by_condition_per_face(condselect="fq > 0.5")
                ms.meshing_isotropic_explicit_remeshing(targetlen=pymeshlab.PureValue(fine or h / 3), iterations=4,
                                                        featuredeg=60, selectedonly=True, checksurfdist=True,
                                                        maxsurfdist=pymeshlab.PureValue(max_dev))
            ms.compute_selection_by_self_intersections_per_face()
            m = ms.current_mesh()
            si = m.face_selection_array()
            thin_attempts.append(dict(thin_faces=n_thin, self_intersections=int(si.sum())))
            if not si.any():
                break
            Fm = m.face_matrix()
            Vm = m.vertex_matrix()
            near = cKDTree(Vm[Fm[si]].mean(1)).query(cen_in, distance_upper_bound=2 * h)[0] < np.inf
            mask = mask | near
        repairs_pre = thin_attempts
    else:
        ms.meshing_isotropic_explicit_remeshing(targetlen=pymeshlab.PureValue(h), iterations=8, featuredeg=60,
                                                checksurfdist=True, maxsurfdist=pymeshlab.PureValue(max_dev))
    stages = [ms.current_mesh().face_number()]
    repairs = []
    for it in range(repair_iters):
        ms.compute_selection_by_self_intersections_per_face()
        nsi = int(ms.current_mesh().selected_face_number())
        if nsi == 0:
            break
        repairs.append(nsi)
        for _ in range(3):
            ms.compute_selection_transfer_face_to_vertex(inclusive=False)
            ms.compute_selection_transfer_vertex_to_face(inclusive=False)
            ms.apply_selection_dilatation()
        ms.meshing_remove_selected_faces()
        _close(ms)
        ms.meshing_isotropic_explicit_remeshing(targetlen=pymeshlab.PureValue(h), iterations=4, featuredeg=60,
                                                selectedonly=True, checksurfdist=False)
        stages.append(ms.current_mesh().face_number())
    m = ms.current_mesh()
    V1, T1 = m.vertex_matrix().astype(float), m.face_matrix().astype(np.int64)
    if repairs:
        Vt = igl.point_mesh_squared_distance(V1, V, T)[2]
        for it in range(5):
            m2 = pymeshlab.MeshSet()
            m2.add_mesh(pymeshlab.Mesh(Vt, T1.astype(np.int32)))
            m2.compute_selection_by_self_intersections_per_face()
            sel = m2.current_mesh().face_selection_array()
            if not sel.any():
                break
            vs = np.unique(T1[sel])
            ring = np.unique(T1[np.isin(T1, vs).any(1)])
            Vt[ring] = V1[ring]
        V1 = Vt
    nsi = _self_int(V1, T1)
    for it in range(6):                      # residual self-intersections after projection: wider repair, no projection
        if nsi == 0:
            break
        ms = pymeshlab.MeshSet()
        ms.add_mesh(pymeshlab.Mesh(V1, T1.astype(np.int32)))
        ms.compute_selection_by_self_intersections_per_face()
        repairs.append(int(ms.current_mesh().selected_face_number()))
        for _ in range(4 + 2 * it):
            ms.compute_selection_transfer_face_to_vertex(inclusive=False)
            ms.compute_selection_transfer_vertex_to_face(inclusive=False)
            ms.apply_selection_dilatation()
        ms.meshing_remove_selected_faces()
        _close(ms)
        ms.meshing_isotropic_explicit_remeshing(targetlen=pymeshlab.PureValue(h), iterations=4, featuredeg=60,
                                                selectedonly=True, checksurfdist=False)
        V1, T1 = ms.current_mesh().vertex_matrix().astype(float), ms.current_mesh().face_matrix().astype(np.int64)
        nsi = _self_int(V1, T1)
    t1 = trimesh.Trimesh(V1, T1, process=False)
    e_max = float(np.linalg.norm(V1[T1[:, [1, 2, 0]]] - V1[T1], axis=2).max())
    info = dict(faces=len(T1), faces_in=len(T), thin_faces_in=n_thin, thin_attempts=repairs_pre, area=float(t0.area), stages=stages, watertight=bool(t1.is_watertight), self_intersections=nsi, repairs=repairs, edge_max=e_max,
                vol_rel=float(t1.volume / t0.volume - 1),
                hausdorff=float(np.sqrt(igl.point_mesh_squared_distance(V1, V, T)[0].max())))
    if not info["watertight"] or nsi or abs(info["vol_rel"]) > max_vol_rel or e_max > 2.5 * h:
        raise RuntimeError(f"surface remesh check failed: {info}")
    return V1, T1, info


def _gmsh_inside(V: np.ndarray, T: np.ndarray, h: float, workdir=None, opt_rounds: int = 3, algo: int = 1, opt_thr: float = 0.3):
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
        gmsh.option.setNumber("Mesh.Algorithm3D", algo)
        gmsh.option.setNumber("Mesh.Optimize", 1)
        gmsh.option.setNumber("Mesh.OptimizeNetgen", 1)
        gmsh.option.setNumber("Mesh.OptimizeThreshold", opt_thr)
        gmsh.model.mesh.generate(3)
        # extra volume optimisation: Netgen / Relocate3D remove the isolated sliver tets (min dihedral ~0.1 deg) left by
        # the Delaunay pass; boundary nodes stay fixed, so the surface and the volume are unchanged
        for _ in range(opt_rounds):
            for method in ("Netgen", "Relocate3D", ""):
                gmsh.model.mesh.optimize(method, True)
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


# (Mesh.Algorithm3D, optimisation rounds, OptimizeThreshold): tried in order until the minimum dihedral is >= 1 deg
# (4x the 0.25 deg gate); isolated Delaunay slivers (B2 solid h/sqrt2: 0.14 deg, B2 design: 1.1 deg) differ between
# variants (research/vf_b2/meshgate2.py: 14-20 deg and 1.1-1.6 deg with different variants).
_MESH_VARIANTS = ((1, 3, 0.3), (1, 3, 1.0), (4, 1, 0.3), (1, 8, 0.5))


def _gmsh_inside_robust(V, T, h, workdir=None, target_deg: float = 1.0):
    from .bodyfit_gpu import dihedral_min
    best = None
    for algo, rounds, thr in _MESH_VARIANTS:
        try:
            nodes, tets = _gmsh_inside(V, T, h, workdir, opt_rounds=rounds, algo=algo, opt_thr=thr)
        except Exception as e:                 # a variant may fail (e.g. HXT): try the next
            log.warning("gmsh variant %s failed: %s", (algo, rounds, thr), e)
            continue
        dm = float(dihedral_min(nodes, tets).min())
        log.info("gmsh variant %s: %d tets, min dihedral %.2f deg", (algo, rounds, thr), len(tets), dm)
        if best is None or dm > best[0]:
            best = (dm, nodes, tets)
        if dm >= target_deg:
            break
    if best is None:
        raise RuntimeError("gmsh failed for every variant")
    return best[1], best[2]


def mesh_gate(nodes: np.ndarray, tets: np.ndarray, V: np.ndarray, T: np.ndarray, gate: dict | None = None,
              h: float | None = None, enforce: bool = True) -> dict:
    """Quality gate of a body-fitted mesh (§8.6.2). Measures: minimum dihedral angle, fraction of tets below 5 deg,
    longest edge / h, and tet layers through the wall (opposite-wall ray samples, 1 % quantile). Raises RuntimeError on a
    miss when `enforce`; always returns the numbers (`passed`, `failures`)."""
    from .bodyfit_gpu import dihedral_min, layers
    g = {**MESH_GATE, **(gate or {})}
    dm = dihedral_min(nodes, tets)
    L = layers(np.asarray(V, float), np.asarray(T, np.int64), nodes, tets)
    X4 = nodes[tets[:, :4]]
    emax = float(max(np.linalg.norm(X4[:, i] - X4[:, j], axis=1).max() for i in range(4) for j in range(i + 1, 4)))
    q = dict(dihedral_min=float(dm.min()), frac_dihedral_lt_5=float((dm < 5).mean()), n_dihedral_lt_5=int((dm < 5).sum()),
             edge_max=emax, edge_max_over_h=None if h is None else emax / h, layers=L, layers_p01=float(L["layers_q"][1]))
    fails = []
    if q["dihedral_min"] < g["min_dihedral_deg"]:
        fails.append(f"min dihedral {q['dihedral_min']:.3f} < {g['min_dihedral_deg']} deg")
    if q["frac_dihedral_lt_5"] > g["max_frac_dihedral_lt_5"]:
        fails.append(f"fraction of tets below 5 deg {q['frac_dihedral_lt_5']:.2e} > {g['max_frac_dihedral_lt_5']}")
    if q["layers_p01"] < g["min_layers_p01"]:
        fails.append(f"tets through the wall (1 % quantile) {q['layers_p01']:.2f} < {g['min_layers_p01']}")
    if h is not None and q["edge_max_over_h"] > g["max_edge_over_h"]:
        fails.append(f"longest edge {q['edge_max_over_h']:.2f} h > {g['max_edge_over_h']} h")
    q["passed"], q["failures"] = not fails, fails
    if fails and enforce:
        raise RuntimeError(f"body-fitted mesh fails the quality gate: {fails}; {q}")
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
                    E: float = E_MPA, nu: float = NU, backend: str = "keep", fallback: str | None = None, h_min: float | None = None,
                    load_distribution: str = "area", timeout: float | None = None, exclude_bc_mm: float = 2.0,
                    engine: str = "gpu", enforce_gate: bool = True, gate: dict | None = None,
                    remesh_kw: dict | None = None, peel: bool = False) -> BodyFitted:
    """Mesh the STL surface (V, T), write and run the deck, parse the result (§8.6.2-3).

    (ctx_V, ctx_T, ctx_TF) with `support` face ids and `loads` locate the BC nodes (original tessellation).
    remesh_kw: passed to `remesh_surface` (backend "keep"; e.g. thin/fine for a `levelset_surface` input);
    peel: remove flat boundary caps after the volume mesh (`peel_boundary_slivers`).
    """
    wd = Path(workdir)
    wd.mkdir(parents=True, exist_ok=True)
    t0 = time.time()
    minfo = {}
    n4, t4 = tetrahedralise(V, T, h, layers=layers, h_min=h_min, backend=backend, workdir=wd, remesh_kw=remesh_kw,
                            peel=peel, info=minfo)
    nodes, tets10 = to_quadratic(n4, t4)
    Vg, Tg = (n4, boundary_faces(t4)) if backend == "stuff" else (V, T)     # wall rays on the meshed boundary
    q = mesh_gate(nodes, tets10, Vg, Tg, gate, h=h, enforce=enforce_gate)
    q.update(minfo)
    if backend == "stuff":
        q["deviation"] = surface_deviation(n4, Tg, V, T)
    del Vg, Tg
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
                          time.time() - t1, "gpu-P2", dict(case.info, lam_unexcl=o["lam_unexcl"], n_ks_excluded=o["n_ks_excluded"],
                                                       **{k: o[k] for k in ("ndof", "eig_res", "timings_s", "factor_GiB", "hybrid", "factor_host_GiB")}))
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
    K/K0 = C0/C, P/P0 = λ_b/λ_b0 and strength σ0/σ, one-sided (D3): `over` = r_vox/r_bf − 1 ≤ 0.05 passes; a ratio
    miss blocks acceptance (`blocking` True); `under` flags a voxel under-prediction beyond the tolerance (logged, not
    blocking). `diff` = (b − a)/a for every row. Missing entries are skipped.
    """
    tol = {**TOLERANCES, **(tol or {})}
    rows = []

    def add(metric, case, a, b, kind, t, blocking=False):
        if a is None or b is None or not (np.isfinite(a) and np.isfinite(b)):
            return
        diff = (b - a) / a if kind in ("rel", "ratio_over") else b - a
        if kind == "ratio_over":                                        # one-sided: voxel over-prediction only
            over = a / b - 1.0
            rows.append(dict(metric=metric, case=case, voxel=a, bodyfitted=b, diff=diff, kind=kind, tol=t,
                             passed=bool(over <= t), blocking=blocking, over=over, under=bool(over < -t)))
            return
        rows.append(dict(metric=metric, case=case, voxel=a, bodyfitted=b, diff=diff, kind=kind, tol=t,
                         passed=bool(abs(diff) <= t), blocking=blocking, over=None, under=False))

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
                add(k, "ratio", rv[k], rb[k], "ratio_over", tol["ratio"], blocking=True)
    return rows


def write_validation_csv(rows: list[dict], path: str | Path) -> Path:
    """`validation.csv` with a pass/fail per row and a final summary row for blocking ratio misses."""
    path = Path(path)
    cols = ["metric", "case", "voxel", "bodyfitted", "diff", "kind", "tol", "passed", "blocking", "over", "under"]
    with open(path, "w", newline="") as f:
        w = csv.DictWriter(f, cols)
        w.writeheader()
        w.writerows(rows)
    return path
