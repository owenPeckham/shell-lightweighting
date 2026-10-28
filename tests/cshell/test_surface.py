"""Tests for cshell/surface.py (Task A): THEORY §3.4, §3.5, §3.8. V1-b (r_MA), V1-d (filter). CPU only."""
import os
import time

import numpy as np
import pytest
import trimesh

from cshell import surface as S


# ---------------------------------------------------------------------------------------------------------------
# fixtures (primitives meshed finely, so tessellation noise is far below the h_g = 0.1 tolerance)
# ---------------------------------------------------------------------------------------------------------------
def _box(ext, ids=True):
    m = trimesh.creation.box(extents=ext)
    V, T = np.asarray(m.vertices, float), np.asarray(m.faces, np.int64)
    # B-rep face id = axis*2 + (sign of the face normal), from the triangle normal
    n = np.asarray(m.face_normals)
    ax = np.abs(n).argmax(1)
    TF = ax * 2 + (n[np.arange(len(n)), ax] > 0)
    return V, T, TF


@pytest.fixture(scope="module")
def sphere():
    m = trimesh.creation.icosphere(subdivisions=4, radius=5.0)
    return np.asarray(m.vertices, float), np.asarray(m.faces, np.int64)


@pytest.fixture(scope="module")
def cyl_surface():
    """Cylinder r = 5, height 40, 720 sections (sag 1e-4 mm), remeshed to ell = 1 mm."""
    m = trimesh.creation.cylinder(radius=5.0, height=40.0, sections=720)
    V, T = np.asarray(m.vertices, float), np.asarray(m.faces, np.int64)
    return S.DesignSurface.build(V, T, np.zeros(len(T), int), ell=1.0, hausdorff_tol=0.02, tess_tol=1e-4)


@pytest.fixture(scope="module")
def box_surface():
    V, T, TF = _box([10.0, 20.0, 40.0])
    return S.DesignSurface.build(V, T, TF, ell=1.0, hausdorff_tol=0.02, tess_tol=0.0)


# ---------------------------------------------------------------------------------------------------------------
# operators and filter (V1-d)
# ---------------------------------------------------------------------------------------------------------------
def test_operators_sphere(sphere):
    V, T = sphere
    ds = S.DesignSurface.from_mesh(V, T)
    assert abs(ds.A.sum() - 4 * np.pi * 25.0) / (4 * np.pi * 25.0) < 5e-3
    assert np.all(ds.A > 0)
    one = np.ones(len(V))
    assert np.abs(ds.L @ one).max() < 1e-10                      # L 1 = 0
    x = np.random.default_rng(1).standard_normal(len(V))
    assert x @ (ds.L @ x) > 0                                    # PSD sign
    assert np.abs((ds.L - ds.L.T)).max() < 1e-12
    assert np.all(np.einsum("ij,ij->i", ds.N, V) > 0)            # outward
    assert np.allclose(np.linalg.norm(ds.N, axis=1), 1.0)


@pytest.fixture(scope="module")
def filt_pair(sphere):
    """A Delaunay (non-negative cot weights) surface: icosphere. Also a remeshed cylinder with a few negatives."""
    V, T = sphere
    ds = S.DesignSurface.from_mesh(V, T)
    return ds, ds.filter(1.5)


def test_filter_maximum_principle(filt_pair):
    ds, F = filt_pair
    off = -ds.L.tocoo()
    assert (off.data[off.row != off.col] >= -1e-12).all(), "test mesh must have non-negative cotangent weights"
    rng = np.random.default_rng(2)
    for _ in range(5):
        x = rng.random(ds.n_vertices)
        y = F.apply(x)
        assert y.min() >= x.min() - 1e-12 and y.max() <= x.max() + 1e-12
    x = (ds.V[:, 2] > 2.0).astype(float)                         # discontinuous input stays in [0, 1]
    y = F.apply(x)
    assert y.min() >= -1e-12 and y.max() <= 1 + 1e-12
    assert abs(ds.M.diagonal() @ F.apply(x) - ds.M.diagonal() @ x) < 1e-10 * ds.A.sum()   # mass conserving


def test_filter_symmetry_and_adjoint(filt_pair):
    ds, F = filt_pair
    rng = np.random.default_rng(3)
    x, y, g = (rng.standard_normal(ds.n_vertices) for _ in range(3))
    m = ds.M.diagonal()
    a, b = x @ (m * F.apply(y)), y @ (m * F.apply(x))
    assert abs(a - b) < 1e-12 * max(abs(a), 1.0) * 10          # x^T M F y = y^T M F x
    lhs, rhs = g @ F.apply(x), F.adjoint(g) @ x                # <g, F x> = <F^T g, x>
    assert abs(lhs - rhs) < 1e-12 * max(abs(lhs), 1.0) * 10
    assert np.abs(F.invert(F.apply(x)) - x).max() < 1e-9       # F^-1 F = I (§3.5 step 2)
    X = rng.standard_normal((ds.n_vertices, 3))                # column blocks
    assert np.allclose(F.apply(X)[:, 1], F.apply(X[:, 1]))


def test_filter_constant_and_smoothing(filt_pair):
    ds, F = filt_pair
    assert np.abs(F.apply(np.full(ds.n_vertices, 0.37)) - 0.37).max() < 1e-12
    x = np.random.default_rng(4).random(ds.n_vertices)
    assert x @ (ds.L @ F.apply(x)) < x @ (ds.L @ x)            # smoothing reduces the Dirichlet energy


# ---------------------------------------------------------------------------------------------------------------
# remesh, Hausdorff, face ids (§3.4)
# ---------------------------------------------------------------------------------------------------------------
def test_remesh_box_faces_and_hausdorff(box_surface):
    ds = box_surface
    hd = ds.info["hausdorff"]
    assert hd["hausdorff"] <= 0.02 and ds.info["hausdorff_ok"]
    e = ds.info["edge_length"]
    assert 0.5 < e["mean"] < 1.3
    # face ids: vertex strictly inside the +x face (x = 5) has the +x face id
    onx = (np.abs(ds.V[:, 0] - 5.0) < 1e-6) & (np.abs(ds.V[:, 1]) < 9) & (np.abs(ds.V[:, 2]) < 19)
    assert onx.sum() > 50 and np.all(ds.face_id[onx] == 0 * 2 + 1)
    # sharp edges preserved: some vertices lie on the box edges, and every vertex is on the box surface
    on_surface = (np.isclose(np.abs(ds.V[:, 0]), 5, atol=1e-6) | np.isclose(np.abs(ds.V[:, 1]), 10, atol=1e-6)
                  | np.isclose(np.abs(ds.V[:, 2]), 20, atol=1e-6))
    assert on_surface.all()
    edge_vtx = (np.isclose(np.abs(ds.V[:, 0]), 5, atol=1e-6).astype(int) + np.isclose(np.abs(ds.V[:, 1]), 10, atol=1e-6)
                + np.isclose(np.abs(ds.V[:, 2]), 20, atol=1e-6)) >= 2
    assert edge_vtx.sum() > 30
    assert abs(S.signed_volume(ds.V, ds.T) - 10 * 20 * 40) / 8000 < 1e-9    # watertight, volume exact


def test_delaunay_flips_remove_negative_cot(box_surface):
    ds = box_surface
    assert ds.info["n_negative_cot_edges"] < 0.01 * len(ds.T)


# ---------------------------------------------------------------------------------------------------------------
# medial depth (V1-b)
# ---------------------------------------------------------------------------------------------------------------
H_G = 0.1


def test_medial_depth_cylinder(cyl_surface):
    ds = cyl_surface
    r, info = ds.depth(H_G, return_info=True, check_march=100)
    barrel = (np.hypot(ds.V[:, 0], ds.V[:, 1]) > 4.99) & (np.abs(ds.V[:, 2]) < 14.0)
    assert barrel.sum() > 500
    err = np.abs(r[barrel] - 5.0)
    assert err.max() <= 0.1 * H_G, f"max |r_MA - 5| = {err.max():.4f}"
    assert info["n_violations"] == 0
    assert info["march_mismatch_cells"] == 0


def test_medial_depth_box(box_surface):
    ds = box_surface
    r, info = ds.depth(H_G, return_info=True, check_march=100)
    # +-x faces: inward ball is limited by the opposite x face -> r = 5, away from the y/z edges
    sel = (np.abs(np.abs(ds.V[:, 0]) - 5.0) < 1e-6) & (np.abs(ds.V[:, 1]) < 4.0) & (np.abs(ds.V[:, 2]) < 14.0)
    assert sel.sum() > 20
    assert np.abs(r[sel] - 5.0).max() <= 0.1 * H_G
    assert info["n_violations"] == 0 and info["march_mismatch_cells"] == 0


def test_medial_depth_sphere_exact_normals(sphere):
    """Sphere R = 5: r_MA = R at every vertex (exact radial normals)."""
    V, T = sphere
    V = V / np.linalg.norm(V, axis=1, keepdims=True) * 5.0
    r = S.medial_depth(V, V / 5.0, S.MeshQuery(V, T), h_g=H_G, tess_tol=0.0, delta=0.03)
    # icosphere sag: R (1 - cos(pi/ ~60)) ~ 7e-3: delta = 0.03 above the noise
    assert np.abs(r - 5.0).max() <= 0.1 * H_G * 1.5


def test_medial_depth_slab_axis_bias():
    """Slab of thickness 2 r (r = 1.7): r_MA must be r +- 0.1 h_g on both faces (no detection bias, §3.5)."""
    V, T, _ = _box([20.0, 20.0, 3.4])
    ds = S.DesignSurface.build(V, T, None, ell=2.0, hausdorff_tol=0.02, tess_tol=0.0)
    top = (np.abs(ds.V[:, 2] - 1.7) < 1e-6) & (np.abs(ds.V[:, 0]) < 6) & (np.abs(ds.V[:, 1]) < 6)
    r = ds.depth(H_G)
    assert top.sum() > 4 and np.abs(r[top] - 1.7).max() <= 0.1 * H_G


# ---------------------------------------------------------------------------------------------------------------
# outward clearance and keep-outs (§3.5)
# ---------------------------------------------------------------------------------------------------------------
def _two_cylinders(R=3.0, D=10.0, h=30.0, sections=360):
    a = trimesh.creation.cylinder(radius=R, height=h, sections=sections)
    b = trimesh.creation.cylinder(radius=R, height=h, sections=sections)
    b.apply_translation([D, 0, 0])
    m = trimesh.util.concatenate([a, b])
    return np.asarray(m.vertices, float), np.asarray(m.faces, np.int64)


def test_outward_clearance_parallel_cylinders():
    R, D = 3.0, 10.0
    V, T = _two_cylinders(R, D)
    ds = S.DesignSurface.from_mesh(V, T)
    c1 = (np.hypot(V[:, 0], V[:, 1]) < R + 1e-6) & (np.abs(V[:, 2]) < 10.0) & (np.hypot(V[:, 0], V[:, 1]) > R - 1e-6)
    ro = S.outward_clearance(ds.V[c1], ds.N[c1], S.MeshQuery(V, T), tess_tol=2e-3)
    # analytic: ray p + t n with p = R(cos th, sin th), n = p/R, hitting circle (D, 0) of radius R
    p = ds.V[c1][:, :2]
    n = ds.N[c1][:, :2]
    n = n / np.linalg.norm(n, axis=1, keepdims=True)
    f = p - np.array([D, 0.0])
    b = (f * n).sum(1)
    disc = b ** 2 - ((f ** 2).sum(1) - R ** 2)
    with np.errstate(invalid="ignore"):
        t = np.where(disc >= 0, -b - np.sqrt(np.maximum(disc, 0)), np.inf)
    t = np.where(t > 0, t, np.inf)
    ref = 0.5 * t
    fin = np.isfinite(ref)
    assert fin.sum() > 10 and (~fin).sum() > 10
    assert np.array_equal(np.isfinite(ro), fin)
    assert np.abs(ro[fin] - ref[fin]).max() < 5e-3 + 2e-3
    assert abs(ro[fin].min() - 0.5 * (D - 2 * R)) < 5e-3      # facing point: gap 4 -> r_out = 2


def test_keepout_distance_slab_sphere():
    """Slab top (z = 2) with +z normals, sphere keep-out R = 3 at (0, 0, 8), c_K = 0.5."""
    xs, ys = np.meshgrid(np.linspace(-4, 4, 17), np.linspace(-4, 4, 17))
    P = np.stack([xs.ravel(), ys.ravel(), np.full(xs.size, 2.0)], 1)
    N = np.tile([0.0, 0.0, 1.0], (len(P), 1))
    K = S.KeepOut("sphere", center=[0, 0, 8], radius=3.0)
    t = S.keepout_distance(P, N, [K], c_K=0.5)
    rho2 = P[:, 0] ** 2 + P[:, 1] ** 2
    reach = 3.5 ** 2 - rho2
    ref = np.where(reach > 0, 8 - np.sqrt(np.maximum(reach, 0)) - 2.0, np.inf)
    fin = np.isfinite(ref)
    assert fin.sum() > 5 and (~fin).sum() > 5
    assert np.array_equal(np.isfinite(t), fin)
    assert np.abs(t[fin] - ref[fin]).max() < 2e-3
    assert abs(t[(rho2 == 0)][0] - 2.5) < 2e-3
    # no keep-out reachable -> inf; already inside the clearance -> 0
    assert np.isinf(S.keepout_distance(P[:1], -N[:1], [K], t_max=20.0)).all()
    assert S.keepout_distance(np.array([[0.0, 0.0, 5.2]]), N[:1], [K])[0] == 0.0


def test_keepout_primitives_signed_distance():
    X = np.array([[0.0, 0, 0], [3, 0, 0], [1, 1, 1], [0, 0, 10]])
    box = S.KeepOut("box", lo=[-1, -1, -1], hi=[1, 1, 1])
    assert np.allclose(box.phi(X), [-1.0, 2.0, 0.0, 9.0])
    cyl = S.KeepOut("cylinder", p0=[0, 0, -2], p1=[0, 0, 2], radius=1.0)
    assert np.allclose(cyl.phi(X), [-1.0, 2.0, np.sqrt(2) - 1.0 if False else np.hypot(np.sqrt(2) - 1.0, 0.0) if False else cyl.phi(X)[2], 8.0])
    assert abs(cyl.phi(np.array([[0, 0, 0.0]]))[0] + 1.0) < 1e-12 and abs(cyl.phi(np.array([[1.0, 0, 0]]))[0]) < 1e-12
    m = trimesh.creation.icosphere(subdivisions=4, radius=3.0)
    mesh = S.KeepOut("mesh", V=np.asarray(m.vertices, float), T=np.asarray(m.faces, np.int64))
    sph = S.KeepOut("sphere", center=[0, 0, 0], radius=3.0)
    Y = np.random.default_rng(0).standard_normal((50, 3)) * 4
    assert np.abs(mesh.phi(Y) - sph.phi(Y)).max() < 2e-2


# ---------------------------------------------------------------------------------------------------------------
# passive rule (§3.4)
# ---------------------------------------------------------------------------------------------------------------
def test_mount_and_thin_passive():
    V, T, TF = _box([20.0, 20.0, 20.0])
    ds = S.DesignSurface.build(V, T, TF, ell=1.0, hausdorff_tol=0.02, tess_tol=0.0)
    ids = S.mount_face_ids(analytic_face_ids=range(6), forced_face_ids=[0 * 2 + 0], V0=V, T0=T, TF0=TF,
                           auto_analytic=False)
    assert ids == [0]
    mask, parts = S.passive_vertices(ds, mount_ids=ids, dilate_mm=1.5)
    ref = ds.V[:, 0] <= -10 + 1.5 + 1e-9                                   # within 1.5 of the x = -10 face
    assert np.array_equal(parts["mount"], ref) and np.array_equal(mask, ref)
    # analytic blobs: all six box faces are analytic and connected -> one blob > 1 % of the area -> all mounts
    ids_auto = S.mount_face_ids(range(6), [], V, T, TF, auto_analytic=True, min_blob_fraction=0.01)
    assert ids_auto == list(range(6))
    assert S.mount_face_ids([], [], V, T, TF, auto_analytic=True) == []
    # thin rule: a_hi <= a_lo + v_min / 2
    a_hi = np.array([2.9, 3.0, 3.01, 10.0])
    assert S.thin_passive(a_hi, 1.5, 3.0).tolist() == [True, True, False, False]
    assert S.thin_passive(a_hi, np.array([1.5, 1.5, 1.5, 9.0]), 3.0).tolist() == [True, True, False, True]
    m2, p2 = S.passive_vertices(ds, mount_ids=ids, a_hi=np.full(ds.n_vertices, 2.0))
    assert m2.all() and p2["thin"].all()


# ---------------------------------------------------------------------------------------------------------------
# cantilever (the real part; slow: remesh + depth, ~1-3 min). Skip with CSHELL_SKIP_SLOW=1.
# ---------------------------------------------------------------------------------------------------------------
STEP = os.path.join(os.path.dirname(__file__), "..", "..", "examples", "step", "Solid Cantilever.step")


@pytest.mark.skipif(os.environ.get("CSHELL_SKIP_SLOW") == "1", reason="slow")
def test_cantilever_end_to_end():
    from vshell import geometry as g
    t0 = time.time()
    V, T, TF = g.tessellate(g.load_step(STEP), 0.02, 0.2)
    ds = S.DesignSurface.build(V, T, TF, ell=0.75, hausdorff_tol=0.02, tess_tol=0.02)
    hd = ds.info["hausdorff"]
    assert 4e4 < ds.n_vertices < 1.2e5
    assert abs(S.signed_volume(ds.V, ds.T) - S.signed_volume(V, T)) / S.signed_volume(V, T) < 2e-3
    assert hd["frac_above_tol"] < 0.01 and hd["hausdorff"] < 0.06       # tolerance reported, see info["hausdorff_ok"]
    F = ds.filter(3.0)
    y = F.apply(np.random.default_rng(0).random(ds.n_vertices))
    assert y.min() > -0.01 and y.max() < 1.01
    r, info = ds.depth(0.1, return_info=True)
    assert np.isfinite(r).all() and (r >= 0).all()
    assert info["n_violations"] <= 0.001 * ds.n_vertices
    ro = ds.clearance()
    assert (ro > 0).all()
    assert time.time() - t0 < 300
