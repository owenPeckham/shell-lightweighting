"""Tests for cshell/band.py (THEORY §3.1, §3.6, §3.7, §3.10; verification V1-a). CPU by default (libigl backend and
Warp on its CPU device); the Warp-on-GPU comparison is marked @pytest.mark.gpu."""
import math

import numpy as np
import pytest
import trimesh

from cshell import band as B
from cshell.band import Band, heaviside_quintic


# ---- meshes and analytic SDFs ---------------------------------------------------------------------------------
def _mesh(m):
    return np.asarray(m.vertices, float), np.asarray(m.faces, np.int32)


@pytest.fixture(scope="module")
def sphere():
    V, T = _mesh(trimesh.creation.icosphere(subdivisions=7, radius=5.0))
    return V, T, lambda x: np.linalg.norm(x, axis=1) - 5.0


@pytest.fixture(scope="module")
def cylinder():
    V, T = _mesh(trimesh.creation.cylinder(radius=5.0, height=20.0, sections=1440))   # axis z, centred

    def sdf(x):
        q = np.stack([np.hypot(x[:, 0], x[:, 1]) - 5.0, np.abs(x[:, 2]) - 10.0], 1)
        return np.minimum(q.max(1), 0.0) + np.linalg.norm(np.maximum(q, 0.0), axis=1)
    return V, T, sdf


@pytest.fixture(scope="module")
def box():
    V, T = _mesh(trimesh.creation.box(extents=(10.0, 14.0, 20.0)))

    def sdf(x):
        q = np.abs(x) - np.array([5.0, 7.0, 10.0])
        return np.minimum(q.max(1), 0.0) + np.linalg.norm(np.maximum(q, 0.0), axis=1)
    return V, T, sdf


BACKENDS = [pytest.param(dict(backend="igl"), id="igl"),
            pytest.param(dict(backend="warp", device="cpu"), id="warp-cpu")]


@pytest.fixture(params=BACKENDS)
def kw(request):
    return request.param


# ---- φ against analytic SDFs ----------------------------------------------------------------------------------
@pytest.mark.parametrize("name", ["sphere", "cylinder", "box"])
def test_phi_matches_analytic_sdf(name, request, kw):
    V, T, sdf = request.getfixturevalue(name)
    b = Band.build(V, T, h_g=0.5, D_comp=1.5, **kw)
    err = np.abs(b.phi.astype(float) - sdf(b.centres()))
    assert b.n > 1000
    assert err.max() < 1e-4, err.max()


def test_band_membership_is_exact(box, kw):
    """The coarse-to-fine search returns exactly the cells with -r_in-h <= φ <= D+h (dense check, small grid)."""
    V, T, sdf = box
    for r_in in (math.inf, 3.0):
        b = Band.build(V, T, h_g=0.5, D_comp=1.0, r_in=r_in, levels=(8, 4), **kw)
        g = b.grid
        ax = [g.origin[a] + g.h * (np.arange(g.shape[a]) + 0.5) for a in range(3)]
        X = np.stack(np.meshgrid(*ax, indexing="ij"), -1).reshape(-1, 3)
        ref = sdf(X)
        lo = -math.inf if math.isinf(r_in) else -r_in - g.h
        inb = (ref >= lo) & (ref <= 1.0 + g.h)
        margin = np.minimum(np.abs(ref - lo), np.abs(ref - (1.0 + g.h)))   # skip cells within float32 noise of a threshold
        got = np.zeros(len(X), bool)
        got[b.flat] = True
        bad = (got != inb) & (margin > 1e-4)
        assert not bad.any()
        assert b.flat.tolist() == sorted(b.flat.tolist()) and len(np.unique(b.flat)) == b.n
        if not math.isinf(r_in):
            assert abs(b.n_deep - int((ref < lo - 1e-4).sum())) <= 2


def test_chunking_does_not_change_result(cylinder):
    V, T, _ = cylinder
    a = Band.build(V, T, 0.5, 1.0, backend="igl")
    c = Band.build(V, T, 0.5, 1.0, backend="igl", chunk=3000, levels=(16, 4))
    assert np.array_equal(a.flat, c.flat) and np.array_equal(a.phi, c.phi)


def test_extend_equals_direct_build(cylinder, kw):
    V, T, _ = cylinder
    b1 = Band.build(V, T, 0.5, 1.0, **kw)
    b2 = b1.extend(2.5)
    d = Band.build(V, T, 0.5, 2.5, **kw)
    assert b2.n >= b1.n and b2.D_comp == 2.5
    assert b2.grid.shape == d.grid.shape and np.allclose(b2.grid.origin, d.grid.origin, atol=1e-9)
    assert np.array_equal(b2.flat, d.flat)
    assert np.allclose(b2.phi, d.phi, atol=1e-6)
    # existing cells are kept unchanged (re-based indices)
    m = b2.grid.n_pad - b1.grid.n_pad
    i, j, k = b1.grid.unravel(b1.flat)
    pos = np.searchsorted(b2.flat, b2.grid.flat(i + m, j + m, k + m))
    assert np.array_equal(b2.flat[pos], b2.grid.flat(i + m, j + m, k + m))
    assert np.allclose(b2.centres(pos), b1.centres(), atol=1e-9) and np.array_equal(b2.phi[pos], b1.phi)


# ---- barycentrics, closest points --------------------------------------------------------------------------------
def test_closest_point_barycentrics(sphere, kw):
    V, T, _ = sphere
    b = Band.build(V, T, 0.5, 1.0, **kw)
    pi = b.closest_points()
    x = b.centres()
    # |x - π| = |φ| and the closest point lies on the surface (to within the facet curvature)
    assert np.abs(np.linalg.norm(x - pi, axis=1) - np.abs(b.phi)).max() < 2e-4
    assert np.abs(np.linalg.norm(pi, axis=1) - 5.0).max() < 1e-3
    assert np.all(b.bary > -1e-4)


# ---- §3.6 P -----------------------------------------------------------------------------------------------------
def test_P_rows_sum_to_one_and_reproduce_linear_field(kw):
    # fine tessellation = subdivided box; Γ_h = the 12-triangle box with a different triangulation: the surfaces
    # coincide, so π(x) lies on Γ_h and a vertex-linear field must be reproduced exactly at π(x)
    mb = trimesh.creation.box(extents=(10.0, 14.0, 20.0))
    Vf, Tf = trimesh.remesh.subdivide_to_size(mb.vertices, mb.faces, max_edge=1.3)
    Vf, Tf = np.asarray(Vf), np.asarray(Tf, np.int32)
    Vh, Th = _mesh(trimesh.creation.box(extents=(10.0, 14.0, 20.0)))
    Th = Th[:, ::-1].copy()                                    # any orientation is fine for an unsigned closest-point query
    b = Band.build(Vf, Tf, 0.5, 1.0, **kw)
    P = b.P(Vh, Th, **kw)
    assert P.shape == (b.n, len(Vh)) and np.all(np.diff(P.indptr) == 3)
    assert np.abs(np.asarray(P.sum(1)).ravel() - 1.0).max() < 1e-12
    g, c = np.array([0.3, -1.1, 0.7]), 2.0
    f = Vh @ g + c
    got = P @ f
    want = b.closest_points() @ g + c
    assert np.abs(got - want).max() < 2e-4, np.abs(got - want).max()
    # transpose product is the adjoint: <P f, w> = <f, P^T w>
    w = np.random.default_rng(1).standard_normal(b.n)
    assert abs((P @ f) @ w - f @ (P.T @ w)) < 1e-9 * abs(w).sum()


# ---- §3.7 / §3.10 volume ------------------------------------------------------------------------------------------
def test_heaviside_quintic_properties():
    s = np.linspace(-1.5, 1.5, 30001)
    H = heaviside_quintic(s, 1.0)
    assert H[0] == 0 and H[-1] == 1 and abs(heaviside_quintic(np.array([0.0]), 1.0)[0] - 0.5) < 1e-15
    assert np.all(np.diff(H) >= -1e-15)
    dH = np.gradient(H, s)
    assert abs(np.trapezoid(dH, s) - 1.0) < 1e-6
    # C²: H' and H'' vanish at s = ±ε
    assert abs(np.gradient(H, s)[np.argmin(abs(s - 1))]) < 1e-3
    assert abs(heaviside_quintic(np.array([1.0 - 1e-3]), 1.0)[0] - 1) < 1e-8


def _offset_volume_cyl(R, L, d):
    """Volume of {φ <= d}, d >= 0 for a solid cylinder (Steiner-type decomposition)."""
    return (math.pi * (R + d) ** 2 * L + 2 * math.pi * R ** 2 * d + math.pi ** 2 * R * d ** 2 + 4 * math.pi * d ** 3 / 3)


def test_volume_solid_cylinder_and_tube():
    """V1-a at h_g = 0.1: solid cylinder, and a tube ψ = max(φ - d, -φ - a) with constant d, a (§3.2)."""
    R, L, d, a = 5.0, 20.0, 0.5, 1.0
    V, T = _mesh(trimesh.creation.cylinder(radius=R, height=L, sections=720))
    b = Band.build(V, T, h_g=0.1, D_comp=1.0, backend="igl")
    v0 = math.pi * R ** 2 * L
    assert abs(b.V0() / v0 - 1) < 2e-3, b.V0() / v0 - 1
    phi = b.phi.astype(np.float64)
    psi = np.maximum(phi - d, -phi - a)
    v_tube = _offset_volume_cyl(R, L, d) - math.pi * (R - a) ** 2 * (L - 2 * a)
    assert abs(b.volume(psi) / v_tube - 1) < 2e-3, b.volume(psi) / v_tube - 1


def test_volume_with_finite_inner_band_counts_deep_cells():
    R, L = 5.0, 20.0
    V, T = _mesh(trimesh.creation.cylinder(radius=R, height=L, sections=360))
    full = Band.build(V, T, 0.25, 1.0, backend="igl")
    thin = Band.build(V, T, 0.25, 1.0, r_in=0.5, backend="igl")
    assert thin.n < 0.9 * full.n and thin.n_deep > 0
    assert abs(thin.V0() / full.V0() - 1) < 1e-9          # same cells counted once, in band or as deep


# ---- cache ---------------------------------------------------------------------------------------------------------
def test_cache_roundtrip(tmp_path, sphere):
    V, T, _ = sphere
    b = Band.build(V, T, 0.5, 1.0, backend="igl")
    p = tmp_path / "b.npz"
    b.save(p, dict(sha256="x"))
    c = Band.load(p)
    assert np.array_equal(b.flat, c.flat) and np.array_equal(b.phi, c.phi) and np.array_equal(b.bary, c.bary)
    assert c.grid.shape == b.grid.shape and c.D_comp == b.D_comp and np.array_equal(c.V, b.V)
    assert c.V0() == b.V0()


def test_build_from_step_cache(tmp_path, monkeypatch):
    from pathlib import Path
    step = Path(__file__).resolve().parents[2] / "examples" / "step" / "Solid Cantilever.step"
    monkeypatch.setenv("CSHELL_CACHE", str(tmp_path))
    b1 = B.build_from_step(step, h_g=1.0, D_comp=2.0, backend="igl")
    assert not b1.info.get("cached") and any(tmp_path.glob("band_*.npz"))
    b2 = B.build_from_step(step, h_g=1.0, D_comp=2.0, backend="igl")
    assert b2.info["cached"] and np.array_equal(b1.flat, b2.flat)
    assert abs(b1.V0() / 60942.0 - 1) < 0.02                                  # THEORY §4.2 table: V₀ = 60,942 mm³


# ---- GPU --------------------------------------------------------------------------------------------------------------
@pytest.mark.gpu
def test_warp_gpu_equals_igl(cylinder):
    V, T, sdf = cylinder
    g = Band.build(V, T, 0.25, 1.0, backend="warp", device="cuda:0", levels=(16, 4))
    c = Band.build(V, T, 0.25, 1.0, backend="igl")
    assert g.info["backend"] == "warp"
    assert np.array_equal(g.flat, c.flat)
    assert np.abs(g.phi - c.phi).max() < 1e-5
    # closest points agree up to float32 ties between neighbouring near-coplanar facets (facet width 0.022 mm here)
    # and, on the medial axis, between distant facets (a measure-zero set of cells)
    dpi = np.abs(g.closest_points() - c.closest_points()).max(1)
    assert (dpi > 0.03).mean() < 1e-3 and (dpi > 1e-4).mean() < 0.1, ((dpi > 0.03).mean(), (dpi > 1e-4).mean())
    x = g.centres()
    assert np.abs(np.linalg.norm(x - g.closest_points(), axis=1) - np.abs(g.phi)).max() < 2e-4
    assert np.abs(g.phi - sdf(g.centres())).max() < 1e-4
