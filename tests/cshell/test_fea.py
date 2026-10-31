"""ShellCG adapter tests (THEORY §4.3, §6.2, §6.6 item 3). Run on CPU (cpu_direct) and GPU (marked)."""
import numpy as np
import pytest

from cshell import fea
from cshell.fea import ShellCG

E, NU = 1760.0, 0.42


def column(nx_e, ny_e, nz_e, P=(0.0, 0.0, -10.0)):
    """Clamp the x = 0 node face; share P over the x = L node face (h = 1 mm)."""
    shape = (nx_e + 1, ny_e + 1, nz_e + 1, 3)
    fixed = np.zeros(shape, bool)
    fixed[0] = True
    load = np.zeros(shape)
    load[-1, ...] = np.asarray(P) / ((ny_e + 1) * (nz_e + 1))
    return fixed, load


def test_volume_mesh_matches_parent_contract():
    """Same 11-tuple, dtypes and values as the parent when given the parent's own smeared BCs."""
    ne = (6, 3, 4)
    csg_f = np.zeros(ne + (3,), bool); csg_f[0] = True
    csg_l = np.zeros(ne + (3,), np.float32); csg_l[-1, :, :, 2] = -1.0
    parent = fea.analysis.ConjugateGradient("test", printouts="Silent", youngs_modulus=E * 1e6, poisson_ratio=NU,
                                            voxel_size_mm=0.5, csg_domain_data={"fixed_nodes": csg_f, "load_vectors": csg_l},
                                            element_type="q1_im")
    rho = np.random.default_rng(0).uniform(0.1, 1.0, ne).astype(np.float32)
    ref = parent._volume_mesh(rho)
    nn = tuple(n + 1 for n in ne)
    fixed = ref[9].reshape(nn + (3,)).astype(bool)
    load = ref[8].reshape(nn + (3,)).astype(np.float64)
    mine = ShellCG(fixed, load, 0.5, E, NU, warm_start=False)._volume_mesh(rho)
    assert len(mine) == 11
    for a, b in zip(ref, mine):
        if isinstance(a, np.ndarray):
            assert a.dtype == b.dtype and a.shape == b.shape
            np.testing.assert_allclose(a, b, rtol=1e-6, atol=0)
        else:
            assert np.isclose(a, b)
    assert mine[9].dtype == np.int32 and mine[9].size == 3 * np.prod(nn)


def test_cantilever_beam_and_energy_normalisation(engine):
    """40x4x4 clamped column, tip shear: Timoshenko beam within 3 %; sum of U_e = C/2; mm/N units."""
    L, b = 40.0, 4.0
    fixed, load = column(40, 4, 4)
    rel_tol = 1e-9 if engine == "cpu_direct" else 1e-6          # the reference GPU PCG is float32
    cg = ShellCG(fixed, load, 1.0, E, NU, rel_tol=rel_tol, engine=engine, warm_start=False)
    r = cg.solve(np.ones((40, 4, 4), np.float32))
    assert r.converged
    I, A, G = b ** 4 / 12, b * b, E / (2 * (1 + NU))
    kappa = 10 * (1 + NU) / (12 + 11 * NU)                       # rectangle (Cowper)
    C_beam = 10.0 ** 2 * (L ** 3 / (3 * E * I) + L / (kappa * G * A))
    assert abs(r.compliance / C_beam - 1) < 0.03, (r.compliance, C_beam)
    assert abs(r.U_e.sum() / (0.5 * r.compliance) - 1) < (1e-4 if engine == "cpu_direct" else 3e-3)   # GPU: float32 PCG and kernels
    tip = r.u[-1, :, :, 2].mean()
    assert abs(-10.0 * tip / r.compliance - 1) < 1e-4            # C = f·u with u in mm


def test_compliance_sensitivity_fd():
    """THEORY §6.6(3): element-wise central differences of C vs −2U_e/ρ̂ on a 30x4x4 column, < 1e-3. Host float64
    path only: the reference GPU PCG is float32, whose noise swamps element-wise differences (THEORY §13)."""
    engine = "cpu_direct"
    rng = np.random.default_rng(1)
    rho = rng.uniform(0.5, 1.0, (30, 4, 4))
    fixed, load = column(30, 4, 4, P=(-3.0, 2.0, -10.0))
    cg = ShellCG(fixed, load, 1.0, E, NU, rel_tol=1e-10, engine=engine, warm_start=False)
    r = cg.solve(rho)
    g = r.dC_drho_hat(rho)
    for e in [(2, 1, 1), (15, 0, 3), (28, 2, 2)]:
        dh = 2e-2 * rho[e]
        rp, rm = rho.copy(), rho.copy()
        rp[e] += dh; rm[e] -= dh
        fd = (cg.solve(rp).compliance - cg.solve(rm).compliance) / (2 * dh)
        assert abs(fd / g[e] - 1) < 1e-3, (e, fd, g[e])
