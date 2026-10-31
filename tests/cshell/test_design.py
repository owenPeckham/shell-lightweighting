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
