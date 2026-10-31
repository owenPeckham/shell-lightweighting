"""Mesh-quality helpers shared by diag1/mesh2 (units mm)."""
import numpy as np


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


def layers(V, F, nodes, tets, n=20000, seed=0):
    """Tet layers through the wall: ray thickness t at surface samples (opposite wall only) / rms edge of the
    8 tets nearest the mid-wall point."""
    import igl
    from scipy.spatial import cKDTree
    tri = V[F]; nrm = np.cross(tri[:, 1] - tri[:, 0], tri[:, 2] - tri[:, 0]); A = .5 * np.linalg.norm(nrm, axis=1)
    nrm /= 2 * A[:, None]
    rng = np.random.default_rng(seed); pick = rng.choice(len(F), n, p=A / A.sum()); u = rng.random((n, 2))
    fl = u.sum(1) > 1; u[fl] = 1 - u[fl]
    P = tri[pick, 0] + u[:, :1] * (tri[pick, 1] - tri[pick, 0]) + u[:, 1:] * (tri[pick, 2] - tri[pick, 0]); N = nrm[pick]
    aabb = igl.AABB(); aabb.init(np.ascontiguousarray(V), np.ascontiguousarray(F))
    hit, t = aabb.intersect_ray_first(V, F, np.ascontiguousarray(P - 1e-4 * N), np.ascontiguousarray(-N))[:2]
    hit, t = np.asarray(hit), np.asarray(t) + 1e-4
    ok = hit >= 0
    cos = np.full(n, 0.0); cos[ok] = np.einsum("ij,ij->i", nrm[hit[ok]], N[ok])
    ok &= cos < -0.7
    T4 = tets[:, :4]
    L = np.stack([np.linalg.norm(nodes[T4[:, i]] - nodes[T4[:, j]], axis=1) for i, j in ((0, 1), (1, 2), (0, 2), (0, 3), (1, 3), (2, 3))], 1)
    Lr = np.sqrt((L ** 2).mean(1)); cen = nodes[T4].mean(1)
    mid = P[ok] - 0.5 * t[ok, None] * N[ok]
    jj = cKDTree(cen).query(mid, k=8)[1]
    nl = t[ok] / Lr[jj].mean(1)
    thin = t[ok] < 2.0
    return dict(samples=int(ok.sum()), layers_q=np.quantile(nl, [0, .01, .1, .5]).round(3).tolist(),
                frac_lt_2=float((nl < 2).mean()),
                layers_q_walls_lt_2mm=np.quantile(nl[thin], [0, .01, .5]).round(3).tolist() if thin.any() else [],
                wall_q=np.quantile(t[ok], [0, .01, .5]).round(4).tolist())
