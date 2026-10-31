"""Design chain (THEORY §3.7-3.10, §6.1, §6.6(1)): directional derivatives, Taylor test, volume, closure (V1-c, volume)."""
import numpy as np
import pytest
import scipy.sparse as sp
import scipy.sparse.linalg as spla

from cshell.design import DesignMap, DesignSpace, Discretisation, heaviside, passive_thin

R0 = 5.0                      # solid sphere of radius 5: φ = |x| − R0, r_MA = R0 everywhere


class _Filter:
    def __init__(self, V, F, r):
        import igl
        L = -igl.cotmatrix(V, F)
        M = igl.massmatrix(V, F, igl.MASSMATRIX_TYPE_VORONOI)
        self.M = M
        self.lu = spla.factorized((M + r * r * L).tocsc())

    def apply(self, x):
        return self.lu(self.M @ x)

    def adjoint(self, g):
        return self.M @ self.lu(g)


def sphere_problem(h_g=0.25, k=2, sub=4):
    import trimesh
    S = trimesh.creation.icosphere(subdivisions=sub, radius=R0)
    V, F = np.asarray(S.vertices), np.asarray(S.faces)
    n = int(np.ceil(2 * 8.0 / (k * h_g))) * k
    c = -8.0 + h_g * (np.arange(n) + 0.5)
    X = np.stack(np.meshgrid(c, c, c, indexing="ij"), -1).reshape(-1, 3)
    r = np.linalg.norm(X, axis=1)
    keep = r < R0 + 3.0
    X, r = X[keep], r[keep]
    phi = r - R0
    ijk = np.stack(np.meshgrid(*(np.arange(n),) * 3, indexing="ij"), -1).reshape(-1, 3)[keep]
    ne = n // k
    elem = ((ijk[:, 0] // k) * ne + ijk[:, 1] // k) * ne + ijk[:, 2] // k
    cp, _, tri = trimesh.proximity.closest_point(S, X * (R0 / np.maximum(r, 1e-9))[:, None])
    bary = trimesh.triangles.points_to_barycentric(V[F[tri]], cp)
    P = sp.csr_matrix((bary.ravel(), F[tri].ravel(), 3 * np.arange(len(X) + 1)), shape=(len(X), len(V)))
    disc = Discretisation(phi=phi, P=P, elem=elem, n_elem=ne ** 3, k=k, h_g=h_g)
    nv = len(V)
    area = np.asarray(_Filter(V, F, 1.0).M.diagonal())
    ds = DesignSpace(r_ma=np.full(nv, R0), a_hi=np.full(nv, R0 + 0.75 * h_g), d_hi=np.full(nv, 4.0), area=area,
                     passive=np.zeros(nv, bool), filt=_Filter(V, F, 1.5))
    return ds, disc, V


def test_heaviside_integral():
    s = np.linspace(-2, 2, 400001)
    H, dH = heaviside(np, s, 1.0)
    assert abs(np.trapezoid(dH, s) - 1) < 1e-8 and H[0] == 0 and H[-1] == 1
    np.testing.assert_allclose(np.gradient(H, s)[1000:-1000], dH[1000:-1000], atol=1e-6)


@pytest.mark.parametrize("as_built", [False, True])
def test_directional_and_taylor(as_built):
    """§6.6(1): V and Σ g_e ρ̂_e along δξ = F·random; rel error < 1e-6 and O(h²) Taylor remainder."""
    ds, disc, V = sphere_problem()
    dm = DesignMap(ds, disc)
    rng = np.random.default_rng(0)
    nv = len(V)
    xd, xa = rng.uniform(0.2, 0.6, nv), rng.uniform(0.1, 0.5, nv)
    dd, da = ds.filt.apply(rng.standard_normal(nv)), ds.filt.apply(rng.standard_normal(nv))
    gE = rng.standard_normal(disc.n_elem)

    def fun(x1, x2):
        st = dm.forward(x1, x2, as_built)
        return np.array([dm.volume(st), float(gE @ dm.rho_hat(st))]), st

    f0, st = fun(xd, xa)
    gv = dm.volume_grad(st)
    gr = dm.pull_rho_hat(st, gE)
    pred = np.array([gv[0] @ dd + gv[1] @ da, gr[0] @ dd + gr[1] @ da])
    h = 1e-6
    fd = (fun(xd + h * dd, xa + h * da)[0] - fun(xd - h * dd, xa - h * da)[0]) / (2 * h)
    assert np.all(np.abs(fd / pred - 1) < 1e-6), (fd, pred)
    rem = [np.abs(fun(xd + t * dd, xa + t * da)[0] - f0 - t * pred) for t in (1e-2, 1e-3, 1e-4)]
    slope = np.log10(rem[0] / rem[2]) / 2
    assert np.all(slope > 1.8), slope


def test_all_solid_volume_and_sliver():
    """V1-c (volume half): ξ̃ᵃ ≡ 1, ξ̃ᵈ ≡ 0 reproduces V₀ (ψ = φ) within 5e-4; sliver zero when closed."""
    ds, disc, V = sphere_problem(h_g=0.1, sub=5)
    dm = DesignMap(ds, disc)
    nv = len(V)
    st = dm.forward(np.zeros(nv), np.ones(nv), as_built=False)
    chi0, _ = heaviside(np, -disc.phi, disc.eps)
    V0 = float(chi0.sum()) * disc.h_g ** 3
    assert abs(dm.volume(st) / V0 - 1) < 5e-4
    assert abs(V0 / (4 / 3 * np.pi * R0 ** 3) - 1) < 2e-3                  # V1-a style analytic check
    assert dm.sliver(st)[0] == 0.0


def test_d1_wall_bound_and_passive():
    """D1: a ≥ a_lo(d) ≈ (t_min − d)⁺; w = a + d ≥ t_min − tiny; passive-thin rule."""
    ds, disc, V = sphere_problem()
    dm = DesignMap(ds, disc)
    nv = len(V)
    for xd in (0.0, 0.2, 0.5, 1.0):
        _, _, d, a, a_lo, _ = dm.vertex_fields(np.full(nv, xd), np.zeros(nv))
        assert np.all(a + d >= ds.t_min - 1e-9)                       # softplus(x) ≥ max(x, 0)
        assert np.all(a >= 0) and np.all(np.abs(a - np.maximum(ds.t_min - d, 0)) <= ds.soft_w * np.log(2) + 1e-12)
    assert passive_thin(np.array([1.4, 1.6, 3.0]), 1.5, 3.0).tolist() == [True, False, False]


# ---- PLAN_single_run §1, §2, §4: F², tapers, closure projection, solid start ---------------------------------------
def _helmholtz(V, F, r):
    import igl
    from cshell.surface import HelmholtzFilter
    L = -igl.cotmatrix(V, F)
    M = igl.massmatrix(V, F, igl.MASSMATRIX_TYPE_VORONOI)
    return HelmholtzFilter(M, L, r)


def _sphere_faces(sub=4):
    import trimesh
    S = trimesh.creation.icosphere(subdivisions=sub, radius=R0)
    return np.asarray(S.vertices), np.asarray(S.faces)


def test_repeated_filter_adjoint_and_inverse():
    """F²: ⟨F²x, y⟩ = ⟨x, (Fᵀ)²y⟩ to rounding; FD of the linear map; invert(apply(x)) = x; spike suppression."""
    from cshell.surface import RepeatedFilter
    V, T = _sphere_faces()
    F1 = _helmholtz(V, T, 1.5)
    F2 = RepeatedFilter(F1, 2)
    rng = np.random.default_rng(1)
    x, y = rng.standard_normal(len(V)), rng.standard_normal(len(V))
    lhs, rhs = F2.apply(x) @ y, x @ F2.adjoint(y)
    assert abs(lhs - rhs) <= 1e-12 * max(abs(lhs), 1.0)
    assert np.allclose(F2.apply(x), F1.apply(F1.apply(x)))
    i = 7                                                      # FD of y·F²x along e_i equals (F²)ᵀy at i
    for h in (1e-2, 1e-4):
        e = np.zeros(len(V)); e[i] = h
        fd = (F2.apply(x + e) @ y - F2.apply(x - e) @ y) / (2 * h)
        assert abs(fd - F2.adjoint(y)[i]) <= 1e-8 * max(1.0, abs(fd))
    assert np.allclose(F2.invert(F2.apply(x)), x, atol=1e-8)
    spike = np.zeros(len(V)); spike[0] = 1.0                 # one-vertex spike: peak/1-ring ratio drops with F²
    nb = np.unique(T[np.any(T == 0, axis=1)]); nb = nb[nb != 0]
    r1 = F1.apply(spike); r2 = F2.apply(spike)
    assert r2[0] / r2[nb].mean() < r1[0] / r1[nb].mean()


def _tapered_space(closure=0.9):
    from cshell.design import taper_weight
    from cshell.surface import RepeatedFilter, surface_distance
    ds, disc, V = sphere_problem()
    _, T = _sphere_faces(4)
    ds.filt = RepeatedFilter(_helmholtz(V, T, 1.5), 2)
    ds.passive = V[:, 2] > 0.8 * R0                               # a passive cap
    H = V[:, 0] < 2.0                                              # hollow-allowed set
    ds.taper_d = taper_weight(surface_distance(V, T, ds.passive), 1.5)
    ds.taper_a = taper_weight(surface_distance(V, T, ds.passive | ~H), 1.5)
    ds.closure_xi = closure
    ds.a_cl = ds.a_hi + 0.3 * (1 + np.sin(V[:, 0]))               # closed inset above a^hi (raw-depth blend)
    ds.a_top = ds.r_ma - 0.5 * ds.v_min                            # open voids ≥ v_min across (sr-fix 2)
    return ds, disc, V, T, H


@pytest.mark.parametrize("as_built", [False, True])
def test_directional_with_taper_and_projection(as_built):
    """∂/∂ξ through F², τ_d, τ_a and the closure projection: FD agreement < 1e-6, Taylor slope ≈ 2."""
    ds, disc, V, T, H = _tapered_space()
    dm = DesignMap(ds, disc)
    rng = np.random.default_rng(3)
    nv = len(V)
    xd = rng.uniform(0.2, 0.6, nv)
    xa = 0.5 + 0.48 * np.sin(0.6 * V[:, 1])                       # smooth: spans open, band and closed after F²
    dd, da = ds.filt.apply(rng.standard_normal(nv)), ds.filt.apply(rng.standard_normal(nv))
    gE = rng.standard_normal(disc.n_elem)

    def fun(x1, x2):
        st = dm.forward(x1, x2, as_built)
        return np.array([dm.volume(st), float(gE @ dm.rho_hat(st))]), st

    f0, st = fun(xd, xa)
    xc, w = ds.closure_xi, ds.closure_band                        # the closure band is exercised
    assert np.any((st.xit_a > xc - w) & (st.xit_a < xc))
    gv, gr = dm.volume_grad(st), dm.pull_rho_hat(st, gE)
    pred = np.array([gv[0] @ dd + gv[1] @ da, gr[0] @ dd + gr[1] @ da])
    h = 1e-6
    fd = (fun(xd + h * dd, xa + h * da)[0] - fun(xd - h * dd, xa - h * da)[0]) / (2 * h)
    assert np.all(np.abs(fd / pred - 1) < 1e-6), (fd, pred)
    rem = [np.abs(fun(xd + t * dd, xa + t * da)[0] - f0 - t * pred) for t in (1e-2, 1e-3, 1e-4)]
    assert np.all(np.log10(rem[0] / rem[2]) / 2 > 1.8)


def test_taper_and_hollow_closure():
    """τ_d: d = 0 on passive vertices and no cliff at their border; τ_a: closed (a = a^hi) outside H and on passive."""
    ds, disc, V, T, H = _tapered_space()
    dm = DesignMap(ds, disc)
    nv = len(V)
    _, xt_a, d, a, _, _ = dm.vertex_fields(np.full(nv, 0.8), np.full(nv, 0.2))
    assert np.all(d[ds.passive] == 0) and np.allclose(a[ds.passive | ~H], ds.a_cl[ds.passive | ~H])
    E = np.unique(np.sort(np.concatenate([T[:, [0, 1]], T[:, [1, 2]], T[:, [2, 0]]]), 1), axis=0)
    jump_taper = np.abs(d[E[:, 0]] - d[E[:, 1]]).max()
    ds_hard = DesignSpace(**{**ds.__dict__, "taper_d": None, "taper_a": None})
    d_hard = DesignMap(ds_hard, disc).vertex_fields(np.full(nv, 0.8), np.full(nv, 0.2))[2]
    jump_hard = np.abs(d_hard[E[:, 0]] - d_hard[E[:, 1]]).max()
    print("max edge jump of d: taper", jump_taper, "hard", jump_hard)
    assert jump_taper < 0.5 * jump_hard


def test_closure_projection():
    """ξ̂ = g(min(ξ̃/0.9, 1)): g(0) = 0, ξ̂ = 1 and dξ̂ = 0 for ξ̃ ≥ 0.9, C¹ at 0.9, monotone; inverse; a ≥ r_MA."""
    from cshell.design import closure_projection, closure_projection_inverse
    x = np.linspace(0, 1, 100001)
    y, dy = closure_projection(np, x, 0.9)
    assert y[0] == 0 and np.all(y[x >= 0.9] == 1) and np.all(dy[x >= 0.9] == 0)
    assert np.all(np.diff(y) >= 0) and abs(dy[np.searchsorted(x, 0.9) - 1]) < 1e-3
    np.testing.assert_allclose(np.gradient(y, x)[10:-10], dy[10:-10], atol=1e-3)
    assert np.allclose(closure_projection(np, closure_projection_inverse(y, 0.9), 0.9)[0], y)
    ds, disc, V, T, H = _tapered_space()
    nv = len(V)
    ds.passive = np.zeros(nv, bool)
    ds.taper_a = ds.taper_d = None
    _, xt_a, d, a, _, _ = DesignMap(ds, disc).vertex_fields(np.zeros(nv), np.full(nv, 0.95))
    assert np.all(xt_a >= 0.9) and np.all(a >= ds.r_ma)                 # closure reachable at ξ̃ᵃ ≥ 0.9


def test_inset_map_matches_export_closure():
    """sr-fix 2: ξ̃ᵃ ≥ ξ_c (export.closed_thr) gives a = a_cl exactly (closed as in the export); below the band
    (ξ̃ᵃ ≤ ξ_c − w) the void is at least v_min across (a ≤ r_MA − v_min/2); a is monotone in ξ̃ᵃ."""
    from cshell.design import closure_projection, inset_map, softplus
    ds, disc, V, T, H = _tapered_space()
    nv = len(V)
    a_lo = softplus(np, ds.t_min - np.zeros(nv), ds.soft_w)[0]
    prev = None
    for x in np.linspace(0, 1, 201):
        xt = np.full(nv, x)
        xh, dxh = closure_projection(np, xt, ds.closure_xi)
        a = inset_map(np, ds, xt, xh, dxh, a_lo)[0]
        if x >= ds.closure_xi:
            assert np.allclose(a, ds.a_cl)
        if x <= ds.closure_xi - ds.closure_band:
            assert np.all(2 * (ds.r_ma - a) >= ds.v_min - 1e-9)
        if prev is not None:
            assert np.all(a >= prev - 1e-12)
        prev = a


def test_solid_start_gradient_nonzero():
    """§1: the solid start (ξᵈ = 0, a = r_MA) is solid to < 1e-3 in volume and has ∂V/∂ξᵃ ≠ 0; the old closed start
    (ξᵃ = 1, a = a^hi) has ∂V/∂ξᵃ = 0."""
    from cshell.design import solid_start_xi_a
    ds, disc, V = sphere_problem()
    ds.closure_xi = 0.9
    dm = DesignMap(ds, disc)
    nv = len(V)
    xa = solid_start_xi_a(ds)
    _, _, _, a, _, _ = dm.vertex_fields(np.zeros(nv), xa)
    assert np.allclose(a, ds.r_ma, atol=1e-6)
    ds2, _, _, _, _ = _tapered_space()                             # with the sr-fix closure band: a = r_MA too
    ds2.taper_a = ds2.taper_d = None
    ds2.passive = np.zeros(nv, bool)
    a2 = DesignMap(ds2, disc).vertex_fields(np.zeros(nv), solid_start_xi_a(ds2))[3]
    assert np.abs(a2 - ds2.r_ma).max() < 0.05                     # F² ξ ≈ target (projected Richardson)
    st = dm.forward(np.zeros(nv), xa, as_built=False)
    chi0, _ = heaviside(np, -disc.phi, disc.eps)
    V0 = float(chi0.sum()) * disc.h_g ** 3
    assert abs(dm.volume(st) / V0 - 1) < 1e-3
    assert np.abs(dm.volume_grad(st)[1]).max() > 0
    st1 = dm.forward(np.zeros(nv), np.ones(nv), as_built=False)
    assert np.abs(dm.volume_grad(st1)[1]).max() == 0


def test_surface_distance_on_sphere():
    """Edge-path distance from the north pole vs the geodesic R θ: within [−1 %, +20 %]; no source → ∞."""
    from cshell.surface import surface_distance
    V, T = _sphere_faces(4)
    i = int(np.argmax(V[:, 2]))
    d = surface_distance(V, T, np.array([i]))
    geo = R0 * np.arccos(np.clip(V @ V[i] / R0 ** 2, -1, 1))
    m = geo > 1.0
    r = d[m] / geo[m]
    assert r.min() > 0.99 and r.max() < 1.20 and np.median(r) < 1.10, (r.min(), r.max())
    assert np.all(d[i] == 0)
    assert np.all(np.isinf(surface_distance(V, T, np.zeros(len(V), bool))))
