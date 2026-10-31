"""Shear/flexure ratio R_v (PLAN_single_run §4, §6) against beam theory on a rectangular cantilever with a tip load:
U_shear/U_flex = (3/5)·6EI/(GAL²) = (3/5)(1 + ν)(h/L)² for L/h = 2, 5, 20, within 10 %.

(1) the wall-energy quadrature on the exact Euler-Bernoulli/Jourawski stress field (σ_xx = Mz/I, parabolic τ_xz);
(2) the full chain used by the optimiser: voxel FE (q1_im, CPU direct), element-centre stresses, ElementStressSampler,
wall_energies, summed over the top and bottom faces (lines from each face to the mid-plane cover the volume once)."""
import numpy as np
import pytest

from cshell.objective import ElementStressSampler, shear_flexure_ratio, wall_energies, wall_frames

E, NU = 1760.0, 0.42
H, B = 4.0, 2.0                       # depth (z) and width (y), mm; narrow section: plane stress, parabolic τ


def beam_ratio(L):
    return 0.6 * (1 + NU) * (H / L) ** 2


def face_points(L, nx, ny):
    """Top and bottom face points (cell midpoints) with outward normals ±z and equal areas."""
    x = (np.arange(nx) + 0.5) * L / nx
    y = (np.arange(ny) + 0.5) * B / ny - B / 2
    X, Y = np.meshgrid(x, y, indexing="ij")
    top = np.stack([X.ravel(), Y.ravel(), np.full(X.size, H / 2)], 1)
    bot = top * np.array([1, 1, -1])
    P = np.concatenate([top, bot])
    N = np.concatenate([np.tile([0, 0, 1.0], (len(top), 1)), np.tile([0, 0, -1.0], (len(bot), 1))])
    return P, N


def test_wall_frames_orthonormal():
    rng = np.random.default_rng(0)
    N = rng.standard_normal((50, 3))
    t1, t2, n = wall_frames(N)
    for a, b in ((t1, t2), (t1, n), (t2, n)):
        assert np.abs((a * b).sum(1)).max() < 1e-12
    assert np.allclose(np.linalg.det(np.stack([t1, t2, n], 1)), 1.0)


@pytest.mark.parametrize("L_over_h", [2, 5, 20])
def test_ratio_analytic_beam_field(L_over_h):
    L = L_over_h * H
    P_tip, I = 1.0, B * H ** 3 / 12

    def sampler(X):
        s = np.zeros((len(X), 6))
        s[:, 0] = P_tip * (L - X[:, 0]) * X[:, 2] / I
        s[:, 4] = P_tip / (2 * I) * (H ** 2 / 4 - X[:, 2] ** 2)
        return s, np.ones(len(X))

    P, N = face_points(L, 200, 4)
    Uf, Us = wall_energies(sampler, P, N, np.zeros(len(P)), np.full(len(P), H / 2), E, NU, n_samples=16)
    R = Us.sum() / Uf.sum()
    print(f"L/h {L_over_h}: R {R:.5g} beam {beam_ratio(L):.5g}")
    assert abs(R / beam_ratio(L) - 1) < 0.01


def cantilever_fe(L, h_e=0.25):
    """Voxel cantilever (x ∈ [0, L], y ∈ [−B/2, B/2], z ∈ [−H/2, H/2]), tip shear −1 N in z with a parabolic nodal
    distribution over the end face. Beam-theory root at x = 0: u_x = 0 on the face (plane section, no rotation),
    u_z = 0 on the line z = 0 and u_y = 0 at its middle node (rigid modes only), so the root section contracts freely
    (Poisson). A full clamp adds root shear ∝ ν·(PL) that beam theory leaves out (+38 % in R at L/h = 20, job
    7902770). Returns the fe, primal, grid and the element box origin."""
    from cshell.fe_sparse import ActiveGrid, SparseFE
    shape = (int(round(L / h_e)), int(round(B / h_e)), int(round(H / h_e)))
    g = ActiveGrid(np.ones(shape, bool))
    nn = tuple(n + 1 for n in shape)
    fixed = np.zeros(nn + (3,), bool)
    fixed[0, :, :, 0] = True
    fixed[0, :, nn[2] // 2, 2] = True
    fixed[0, nn[1] // 2, nn[2] // 2, 1] = True
    z = np.arange(nn[2]) * h_e - H / 2
    wz = np.clip(H ** 2 / 4 - z ** 2, 0, None)
    wy = np.ones(nn[1]); wy[[0, -1]] = 0.5
    W = wy[:, None] * wz[None, :]
    load = np.zeros(nn + (3,))
    load[-1, :, :, 2] = -W / W.sum()
    fe = SparseFE(g, fixed, load, h_e, E, NU, backend="cpu")
    fe.factor(np.ones(g.Ne))
    return fe, fe.primal(), g, np.array([0.0, -B / 2, -H / 2])


@pytest.mark.parametrize("L_over_h", [2, 5, 20])
def test_ratio_fe_cantilever(L_over_h):
    L = L_over_h * H
    h_e = 0.25
    fe, pr, g, origin = cantilever_fe(L, h_e)
    sig = fe.centre_stress(pr.u, solid=True)
    smp = ElementStressSampler(sig, np.ones(g.Ne), g.elem_dense, g.elem_shape, origin, h_e)
    P, N = face_points(L, int(round(L / h_e)), int(round(B / h_e)))
    Uf, Us = wall_energies(smp, P, N, np.zeros(len(P)), np.full(len(P), H / 2), E, NU, n_samples=16)
    R = Us.sum() / Uf.sum()
    # per-vertex R is monotone in the distance from the tip region: bending grows towards the root
    Rv = shear_flexure_ratio(Uf, Us)
    print(f"L/h {L_over_h}: R {R:.5g} beam {beam_ratio(L):.5g} ratio {R / beam_ratio(L):.4f}; "
          f"R_v median {np.median(Rv):.3g}")
    assert abs(R / beam_ratio(L) - 1) < 0.10


def test_shear_dominated_core_vs_bending_skin():
    """Behaviour (§4): pure bending gives R → 0; a pure transverse-shear field gives R → ∞ (keep the core)."""
    P, N = face_points(10.0, 10, 2)
    def bend(X):
        s = np.zeros((len(X), 6)); s[:, 0] = X[:, 2]; return s, np.ones(len(X))
    def shear(X):
        s = np.zeros((len(X), 6)); s[:, 4] = 1 - X[:, 2] ** 2; return s, np.ones(len(X))
    z, r = np.zeros(len(P)), np.full(len(P), H / 2)
    Uf, Us = wall_energies(bend, P, N, z, r, E, NU)
    assert Us.max() == 0 and Uf.min() > 0
    Uf, Us = wall_energies(shear, P, N, z, r, E, NU)
    assert Uf.max() == 0 and np.all(shear_flexure_ratio(Uf + 0, Us) > 1e6)
