"""
Automatic segmentation of the part into N control regions on the coarse grid.

1. Mounts: interior voxels whose nearest boundary triangle belongs to an analytic B-rep face
   (plane, cylinder, ...) form 'preserve' blobs. In generative-design outcomes the generated
   geometry is freeform (BSpline) and the preserve/interface geometry is analytic, so these blobs are
   the lugs and pads that were previously split off by hand. Faces carrying supports or loads (and
   any `mounts.extra_faces`) are always mounts. Mounts are never hollowed.
2. Curve skeleton (3D thinning, scikit-image) -> branch graph (skan). Junction voxels whose
   balls of radius joint_k * r overlap are clustered into one joint; the rest of each branch is a strut.
3. Every interior voxel takes the label of its nearest seed (Euclidean distance transform).
Per region: volume, typical radius r, depth Rin, length L and slenderness lambda = L / r_g with
r_g = r / 2 (radius of gyration of a solid circle).
"""
from __future__ import annotations

from dataclasses import dataclass, field

import numpy as np
from scipy import ndimage
from scipy.spatial import cKDTree


@dataclass
class Segmentation:
    label_c: np.ndarray            # (coarse grid) region id inside the part, 0 outside
    label_ext_c: np.ndarray        # region id everywhere (nearest region outside the part)
    regions: list                  # dicts: id, kind, vol_mm3, r_mm, rin_mm, L_mm, lam, centroid, tangent
    skel_pts: np.ndarray           # (n, 3) skeleton voxel coordinates
    skel_tan: np.ndarray           # (n, 3) local unit tangent
    skel_region: np.ndarray        # (n,) region id of each skeleton voxel
    info: dict = field(default_factory=dict)

    def by_id(self):
        return {r["id"]: r for r in self.regions}

    def ids(self, kind=None):
        return [r["id"] for r in self.regions if kind is None or r["kind"] == kind]


def _local_tangents(pts, radius=3.0):
    kt = cKDTree(pts)
    tan = np.zeros_like(pts)
    for i, p in enumerate(pts):
        nb = pts[kt.query_ball_point(p, radius)]
        if len(nb) > 2:
            tan[i] = np.linalg.svd(nb - nb.mean(0), full_matrices=False)[2][0]
        else:
            tan[i] = (1.0, 0.0, 0.0)
    return tan


def segment(grid, V, T, tri_face, analytic_face_ids, forced_face_ids, cfg) -> Segmentation:
    import igl
    from skimage.morphology import skeletonize
    from skan import Skeleton

    h = grid.h_seg
    lo = grid.lo
    inside = grid.inside_c
    depth = grid.depth_c
    seg_cfg, m_cfg = cfg["segmentation"], cfg["mounts"]

    # ---- 1. mounts ---------------------------------------------------------------------------
    idx = np.argwhere(inside)
    P = lo + h * idx
    _, I, _ = igl.signed_distance(P, np.ascontiguousarray(V, np.float64), np.ascontiguousarray(T, np.int64))[:3]
    nearest_face = tri_face[I]
    analytic = np.isin(nearest_face, list(analytic_face_ids)) if m_cfg["auto_analytic"] else np.zeros(len(I), bool)
    forced = np.isin(nearest_face, list(forced_face_ids))
    if len(forced_face_ids):
        Tf = np.ascontiguousarray(T[np.isin(tri_face, list(forced_face_ids))], np.int64)
        d2 = igl.point_mesh_squared_distance(P, np.ascontiguousarray(V, np.float64), Tf)[0]
        forced |= d2 <= max(m_cfg["dilate_mm"], h) ** 2      # every voxel touching a support/load face
    mraw = np.zeros(inside.shape, bool)
    mraw[tuple(idx.T)] = analytic | forced
    forced_grid = np.zeros(inside.shape, bool)
    forced_grid[tuple(idx.T)] = forced
    lab, n = ndimage.label(mraw)
    sizes = ndimage.sum(mraw, lab, range(1, n + 1))
    has_forced = ndimage.maximum(forced_grid, lab, range(1, n + 1)) if n else []
    keep = [i + 1 for i in range(n) if sizes[i] > m_cfg["min_blob_fraction"] * inside.sum() or has_forced[i]]
    mount_lab = np.where(np.isin(lab, keep), lab, 0)
    it = max(int(round(m_cfg["dilate_mm"] / h)), 0)
    if it:
        grown = ndimage.grey_dilation(mount_lab, size=(3, 3, 3))
        for _ in range(it - 1):
            grown = np.where(grown > 0, grown, ndimage.grey_dilation(grown, size=(3, 3, 3)))
        mount_lab = np.where(mount_lab > 0, mount_lab, grown) * inside

    # ---- 2. skeleton -> joints and struts ----------------------------------------------------
    sk = skeletonize(inside)
    skel = Skeleton(sk.astype(np.uint8), spacing=h)
    coords = skel.coordinates * h                   # skan returns voxel indices; convert to mm from the origin
    deg = np.asarray((skel.graph > 0).sum(axis=1)).ravel()
    junc = np.where(deg >= 3)[0]
    rj = depth[tuple(np.round(coords[junc] / h).astype(int).T)] if len(junc) else np.zeros(0)
    k = float(seg_cfg["joint_k"])
    jl = -np.ones(len(junc), int)
    c = 0
    for a in range(len(junc)):
        if jl[a] >= 0:
            continue
        stack = [a]
        jl[a] = c
        while stack:
            u = stack.pop()
            d = np.linalg.norm(coords[junc] - coords[junc[u]], axis=1)
            for v in np.where((d < k * (rj + rj[u])) & (jl < 0))[0]:
                jl[v] = c
                stack.append(v)
        c += 1
    seeds = np.zeros(inside.shape, np.int32)
    regions = []
    Pin = idx * h                                   # interior voxel coords relative to origin
    for j in range(c):
        mem = jl == j
        cen, rad = coords[junc[mem]], k * rj[mem]
        dmin = np.full(len(Pin), np.inf)
        for q, rr in zip(cen, rad):
            dmin = np.minimum(dmin, np.linalg.norm(Pin - q, axis=1) - rr)
        sel = idx[dmin <= 0]
        if len(sel) == 0:
            continue
        rid = len(regions) + 1
        seeds[tuple(sel.T)] = rid
        regions.append(dict(id=rid, kind="joint", r_mm=float(np.mean(rj[mem])), L_mm=None, lam=None))
    for bi in range(skel.n_paths):
        path = skel.path_coordinates(bi) * h
        pv = np.round(path / h).astype(int)
        free = seeds[tuple(pv.T)] == 0
        if free.sum() < 2:
            continue
        rid = len(regions) + 1
        seeds[tuple(pv[free].T)] = rid
        r = depth[tuple(pv[free].T)]
        L = float(np.sum(np.linalg.norm(np.diff(path[free], axis=0), axis=1)))
        rmed = float(np.median(r))
        tang = np.linalg.svd(path[free] - path[free].mean(0), full_matrices=False)[2][0] if free.sum() > 2 else np.array([1.0, 0, 0])
        regions.append(dict(id=rid, kind="strut", r_mm=rmed, L_mm=L, lam=L / max(rmed / 2, 1e-6),
                            tangent=[float(v) for v in tang]))
    if not regions:
        raise RuntimeError("segmentation found no skeleton branches; check the STEP solid and grid.seg_mm")
    _, ind = ndimage.distance_transform_edt(seeds == 0, return_indices=True)
    label = seeds[tuple(ind)] * inside
    # mounts overwrite, one region per mount blob
    for mb in sorted(set(np.unique(mount_lab)) - {0}):
        rid = len(regions) + 1
        label[mount_lab == mb] = rid
        regions.append(dict(id=rid, kind="mount", r_mm=None, L_mm=None, lam=None))
    # drop regions left with almost no volume (merge into neighbours)
    vols = ndimage.sum(inside, label, [r["id"] for r in regions]) * h ** 3
    total = inside.sum() * h ** 3
    tiny = [r["id"] for r, v in zip(regions, vols) if v < 0.002 * total and r["kind"] != "mount"]
    if tiny:
        good = inside & ~np.isin(label, tiny)
        _, ind = ndimage.distance_transform_edt(~good, return_indices=True)
        label = np.where(np.isin(label, tiny), label[tuple(ind)], label) * inside
        regions = [r for r in regions if r["id"] not in tiny]
    # renumber 1..N
    old = [r["id"] for r in regions]
    remap = np.zeros(max(old) + 1, np.int32)
    for new, o in enumerate(old, start=1):
        remap[o] = new
    label = remap[label]
    for new, r in enumerate(regions, start=1):
        r["id"] = new
    # per-region statistics
    ids = [r["id"] for r in regions]
    vols = ndimage.sum(inside, label, ids) * h ** 3
    rin = ndimage.maximum(depth, label, ids)
    cents = ndimage.center_of_mass(inside, label, ids)
    for r, v, ri, ce in zip(regions, vols, rin, cents):
        r["vol_mm3"] = float(v)
        r["rin_mm"] = float(ri) + 0.5 * h + 0.5       # inset that closes the void everywhere in the region
        r["centroid"] = [float(lo[i] + h * ce[i]) for i in range(3)]
        if r["kind"] == "mount":
            r["r_mm"] = float(ri)
    _, ind = ndimage.distance_transform_edt(label == 0, return_indices=True)
    label_ext = label[tuple(ind)]
    # skeleton points with tangents and region ids (used for R, Euler buckling and section inertia)
    skel_pts = lo + coords
    skel_tan = _local_tangents(skel_pts, radius=max(3.0, 3 * h))
    skel_reg = label[tuple(np.round(coords / h).astype(int).T)]
    info = dict(n_joint=sum(r["kind"] == "joint" for r in regions), n_strut=sum(r["kind"] == "strut" for r in regions),
                n_mount=sum(r["kind"] == "mount" for r in regions), n_skeleton_voxels=int(sk.sum()),
                n_paths=int(skel.n_paths))
    return Segmentation(label.astype(np.int32), label_ext.astype(np.int32), regions, skel_pts, skel_tan, skel_reg, info)
