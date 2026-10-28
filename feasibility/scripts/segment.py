"""Automatic region segmentation of a GD solid into N control regions.

1. B-rep faces -> analytic (plane/cylinder/...) vs freeform (BSpline). In Fusion GD outcomes the
   preserve / obstacle-interface geometry is analytic, the generated geometry is BSpline. Voxels whose
   nearest boundary face is analytic become 'mount' regions (kept solid, never shelled) -> replaces the
   manual split.
2. Voxel skeleton (1 mm grid) -> branch graph (skan). Junction clusters -> 'joint' regions,
   branches -> 'strut' regions. Per region: length L, mean radius r, slenderness lambda = L / (r/2).
Outputs region label grid + per-region table (npz/json) and a picture.
"""
import sys, json, numpy as np, trimesh, igl
from scipy import ndimage as ndi
from skimage.morphology import skeletonize
from skan import Skeleton, summarize
from occ_util import DATA_DIR, OUT_DIR, load_step, tessellate, faces, surf_type

H = 1.0  # mm, segmentation grid


def voxel_inside(m, h, pad=3.0):
    lo = m.bounds[0] - pad; hi = m.bounds[1] + pad
    shape = np.ceil((hi - lo) / h).astype(int) + 1
    axes = [lo[i] + h * np.arange(shape[i]) for i in range(3)]
    G = np.stack(np.meshgrid(*axes, indexing='ij'), -1).reshape(-1, 3)
    w = igl.fast_winding_number(m.vertices.astype(np.float64), m.faces.astype(np.int64), G)
    return (w > 0.5).reshape(shape), lo, shape


def segment(name, mount_dilate=1.5, joint_k=1.6, out_dir=None):
    out_dir = out_dir or OUT_DIR
    shp = load_step(f'{DATA_DIR}/{name}.step')
    V, T, F = tessellate(shp, lin=0.05, per_face=True)
    ftype = np.array([surf_type(f) for f in faces(shp)])
    analytic_tri = np.isin(ftype[F], ['Plane', 'Cylinder', 'Cone', 'Sphere', 'Torus'])
    m = trimesh.Trimesh(V, T, process=False)
    inside, lo, shape = voxel_inside(m, H)
    idx = np.argwhere(inside); P = lo + H * idx
    S, I, C = igl.signed_distance(P, V, T)[:3]
    # -- mount detection: interior voxels whose nearest surface patch is analytic, then a little dilation
    mount = np.zeros(shape, bool); mount[tuple(idx.T)] = analytic_tri[I]
    # analytic faces also appear as tiny planar caps elsewhere: keep only sizeable connected blobs
    lab, n = ndi.label(mount); sizes = ndi.sum(mount, lab, range(1, n + 1))
    keep = np.isin(lab, 1 + np.where(sizes > 0.02 * inside.sum())[0])
    mount = ndi.binary_dilation(keep, iterations=int(round(mount_dilate / H))) & inside
    # -- skeleton of the non-mount part (plus a margin so branches reach into the mounts)
    dist_in = ndi.distance_transform_edt(inside) * H
    sk = skeletonize(inside)
    skel = Skeleton(sk.astype(np.uint8), spacing=H)
    br = summarize(skel, separator='_')
    # node degree -> junction nodes
    deg = np.asarray((skel.graph > 0).sum(axis=1)).ravel()
    coords = skel.coordinates
    junc = np.where(deg >= 3)[0]
    rj = dist_in[tuple(np.round(coords[junc] / H).astype(int).T)] if len(junc) else np.array([])
    # cluster junctions whose balls overlap -> one joint each
    jl = -np.ones(len(junc), int); c = 0
    for a in range(len(junc)):
        if jl[a] >= 0: continue
        stack = [a]; jl[a] = c
        while stack:
            u = stack.pop()
            d = np.linalg.norm(coords[junc] - coords[junc[u]], axis=1)
            for v in np.where((d < joint_k * (rj + rj[u])) & (jl < 0))[0]:
                jl[v] = c; stack.append(v)
        c += 1
    n_joint = c
    # region seeds on a label grid: joints first (balls), then strut paths
    seeds = np.zeros(shape, np.int32)
    zz = np.argwhere(inside)
    regions = []
    for j in range(n_joint):
        mem = junc[jl == j]; cen = coords[mem]; rad = joint_k * rj[jl == j]
        rid = len(regions) + 1
        d = np.min(np.linalg.norm((lo + H * zz)[:, None, :] - (cen * 1.0 + lo)[None], axis=2) - rad[None], axis=1)
        sel = zz[d <= 0]; seeds[tuple(sel.T)] = rid
        regions.append(dict(id=rid, kind='joint', r_mm=float(np.mean(rj[jl == j]))))
    for bi in range(len(br)):
        path = skel.path_coordinates(bi)
        pv = np.round(path / H).astype(int)
        free = seeds[tuple(pv.T)] == 0
        if free.sum() < 2: continue
        rid = len(regions) + 1
        seeds[tuple(pv[free].T)] = rid
        r = dist_in[tuple(pv[free].T)]
        L = float(np.sum(np.linalg.norm(np.diff(path[free], axis=0), axis=1)))
        tang = np.linalg.svd(path[free] - path[free].mean(0))[2][0] if free.sum() > 2 else np.array([1., 0, 0])
        regions.append(dict(id=rid, kind='strut', L_mm=L, r_mm=float(np.median(r)), lam=L / max(np.median(r) / 2, 1e-6), tangent=tang.tolist()))
    # every inside voxel -> nearest seed (EDT indices)
    _, ind = ndi.distance_transform_edt(seeds == 0, return_indices=True)
    label = seeds[tuple(ind)] * inside
    MOUNT = len(regions) + 1
    label[mount] = MOUNT
    regions.append(dict(id=MOUNT, kind='mount'))
    # volumes; drop tiny struts into neighbouring regions
    vol = ndi.sum(inside, label, [r['id'] for r in regions]) * H ** 3
    for r, v in zip(regions, vol): r['vol_mm3'] = float(v)
    np.savez_compressed(f'{out_dir}/seg_{name}.npz', label=label.astype(np.int16), lo=lo, h=H, inside=inside, dist=dist_in.astype(np.float32), skel=sk)
    json.dump(regions, open(f'{out_dir}/seg_{name}.json', 'w'), indent=1)
    return regions, label, lo


if __name__ == '__main__':
    for name in sys.argv[1:]:
        regions, label, lo = segment(name)
        k = {}
        for r in regions: k[r['kind']] = k.get(r['kind'], 0) + 1
        print(name, k, 'struts lambda:', sorted([round(r['lam'], 1) for r in regions if r['kind'] == 'strut'], reverse=True)[:12])
