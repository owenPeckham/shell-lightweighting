"""Body-fitted CalculiX validation (THEORY §8.6): mesh, deck, parse, comparison. CPU; ccx is in the cshell env."""
import numpy as np
import pytest
import trimesh

from cshell import validate as vd

E, NU = 1760.0, 0.42


def rod_stl(r, L, sections=96):
    """Capped cylinder along z from 0 to L; TF: 0 = base (support), 1 = tip (load), 2 = side."""
    m = trimesh.creation.cylinder(radius=r, height=L, sections=sections)
    m.apply_translation([0, 0, L / 2])
    V, T = np.asarray(m.vertices), np.asarray(m.faces)
    zc = V[T].mean(1)[:, 2]
    TF = np.where(zc < 1e-6, 0, np.where(zc > L - 1e-6, 1, 2))
    return V, T, TF


def rod_case(r, L, h, force, layers=2.0):
    V, T, TF = rod_stl(r, L)
    n4, t4 = vd.tetrahedralise(V, T, h, layers=layers, h_min=h)
    nodes, t10 = vd.to_quadratic(n4, t4)
    case = vd.build_case(nodes, t10, V, T, TF, [0], [dict(faces=[1], force=force)], bc_tol=1e-3)
    return V, T, TF, case, vd.tet_volumes(n4, t4).sum()


def test_c3d10_ordering_and_orientation():
    """Mid-edge nodes in CalculiX order: 5:(1,2) 6:(2,3) 7:(3,1) 8:(1,4) 9:(2,4) 10:(3,4); shared edges merged."""
    nodes = np.array([[0, 0, 0], [1, 0, 0], [0, 1, 0], [0, 0, 1], [1, 1, 1.0]])
    tets = vd._orient(nodes, np.array([[0, 2, 1, 3], [1, 2, 3, 4]]))        # first is negatively oriented
    assert np.all(vd.tet_volumes(nodes, tets) > 0)
    n10, t10 = vd.to_quadratic(nodes, tets)
    assert t10.shape == (2, 10) and len(n10) == 5 + 9                          # 6 + 6 edges, 3 shared (face 1,2,3)
    for e in range(2):
        c = n10[t10[e, :4]]
        for k, (a, b) in enumerate([(0, 1), (1, 2), (2, 0), (0, 3), (1, 3), (2, 3)]):
            np.testing.assert_allclose(n10[t10[e, 4 + k]], 0.5 * (c[a] + c[b]))


def test_size_field_ordering_and_wall_grading():
    """gmsh `Structured` field: values with x slowest, z fastest. Edge size grows with z on a box; thin plate gets >= 2 layers."""
    box = trimesh.creation.box((12, 8, 12))
    lo = np.array([-8.0, -6.0, -8.0])
    n = np.array([5, 4, 9])
    sp = 2.0
    z = lo[2] + sp * np.arange(n[2])
    s = np.broadcast_to(np.clip(0.8 + 0.25 * (z + 6), 0.8, 3.0)[None, None, :], tuple(n)).ravel()
    nodes, tets = vd.tetrahedralise(np.asarray(box.vertices), np.asarray(box.faces), 3.0, size_field=(lo, sp, n, s))
    c = nodes[tets].mean(1)
    vol = vd.tet_volumes(nodes, tets)
    assert abs(vol.sum() - 12 * 8 * 12) < 1e-6 * 1152
    low, high = vol[c[:, 2] < -3].mean(), vol[c[:, 2] > 3].mean()
    assert high > 4 * low, (low, high)
    plate = trimesh.creation.box((20, 20, 1.0))
    pn, pt = vd.tetrahedralise(np.asarray(plate.vertices), np.asarray(plate.faces), 4.0, layers=2.0, h_min=0.2)
    assert len(np.unique(np.round(pn[:, 2], 4))) >= 3                          # >= 2 elements through the wall
    pts, t = vd.wall_thickness(np.asarray(plate.vertices), np.asarray(plate.faces), spacing=2.0)
    assert abs(np.median(t[np.isfinite(t)]) - 20.0) < 20.0 and np.isclose(t[np.isfinite(t)].min(), 1.0, atol=1e-3)


def test_cantilever_rod_vs_timoshenko_and_skfem_roundtrip(tmp_path):
    """r = 5, L = 40, tip shear 100 N: ccx C3D10 compliance vs Timoshenko (2 %), mesh volume, EVOL sum, deck/parse
    round-trip, and agreement with the scikit-fem P2 fallback on the same mesh (independent P2 implementation)."""
    r, L, P = 5.0, 40.0, 100.0
    V, T, TF, case, vol = rod_case(r, L, 1.2, [P, 0, 0])
    assert abs(vol / (np.pi * r * r * L) - 1) < 0.01                           # 96-gon + tet boundary
    inp = vd.write_inp(tmp_path / "cant.inp", case, want_stress=True, buckle=False)
    run = vd.run_ccx(inp, nproc=8)
    assert run.returncode == 0, run.log.read_text()[-500:]
    res = vd.parse_dat(run.dat)
    C, sq, vol_m = vd.responses_from_ccx(case, res)
    assert res.evol is not None and abs(res.evol.sum() / vol_m - 1) < 1e-6      # EVOL parse and C3D10 volumes
    assert res.stress.shape == (len(case.tets10), 4, 6) and sq > 0
    A = np.pi * r * r
    I = np.pi * r ** 4 / 4
    G = E / (2 * (1 + NU))
    k = 6 * (1 + NU) / (7 + 6 * NU)                                             # Cowper, circle
    C_tim = P * P * (L ** 3 / (3 * E * I) + L / (k * G * A))
    assert abs(C / C_tim - 1) < 0.02, (C, C_tim)
    C_fb = vd.solve_skfem_p2(case)
    assert abs(C_fb / C - 1) < 1e-4, (C_fb, C)
    # bending stress at the root: sigma = P L r / I (on the surface); 0.995 quantile is of the same order (excl. 2 mm)
    assert 0.3 * P * L * r / I < sq < 1.6 * P * L * r / I


def test_area_consistent_load_weights():
    """'area' load distribution: total force preserved, only mid-edge nodes loaded, uniform traction (equal per area)."""
    V, T, TF = rod_stl(5.0, 20.0, 48)
    n4, t4 = vd.tetrahedralise(V, T, 2.5, h_min=2.5)
    nodes, t10 = vd.to_quadratic(n4, t4)
    c = vd.build_case(nodes, t10, V, T, TF, [0], [dict(faces=[1], force=[3.0, 0, -7.0])], 1e-3, load_distribution="area")
    np.testing.assert_allclose(c.load_vec.sum(0), [3.0, 0, -7.0], atol=1e-12)
    assert c.load_nodes.min() >= len(n4)
    c2 = vd.build_case(nodes, t10, V, T, TF, [0], [dict(faces=[1], force=[3.0, 0, -7.0])], 1e-3, load_distribution="nodal")
    np.testing.assert_allclose(c2.load_vec.sum(0), [3.0, 0, -7.0], atol=1e-12)


def test_clamped_free_column_buckling(tmp_path):
    """Clamped-free rod r = 5, L = 100, axial 100 N: first factor vs pi^2 EI/(4 L^2) (+ shear correction), 3 %."""
    r, L, P = 5.0, 100.0, 100.0
    V, T, TF, case, _ = rod_case(r, L, 1.4, [0, 0, -P])
    inp = vd.write_inp(tmp_path / "col.inp", case, n_modes=3, want_stress=False)
    run = vd.run_ccx(inp, nproc=8)
    assert run.returncode == 0, run.log.read_text()[-500:]
    res = vd.parse_dat(run.dat)
    I, A, G = np.pi * r ** 4 / 4, np.pi * r * r, E / (2 * (1 + NU))
    Pe = np.pi ** 2 * E * I / (4 * L ** 2)
    Pcr = Pe / (1 + Pe / (0.9 * G * A))
    lam = res.buckling
    assert lam.size == 3 and abs(lam[0] * P / Pcr - 1) < 0.03, (lam * P, Pcr)
    assert abs(lam[1] / lam[0] - 1) < 0.02                                      # double (x, y) bending mode


def test_ccx_failure_is_not_silent(tmp_path, monkeypatch, caplog):
    """Default: ccx failure raises. fallback='skfem': logs a warning, engine tag set, compliance only."""
    V, T, TF = rod_stl(5.0, 20.0, 48)
    loads = [dict(faces=[1], force=[10.0, 0, 0])]
    bad = vd.CcxRun(1, 0.0, tmp_path / "x.log", tmp_path / "x.dat")
    monkeypatch.setattr(vd, "run_ccx", lambda *a, **k: bad)
    with pytest.raises(RuntimeError, match="ccx failed"):
        vd.validate_design(V, T, V, T, TF, [0], loads, tmp_path / "a", h=2.0, want_stress=False)
    with caplog.at_level("WARNING", logger="cshell.validate"):
        out = vd.validate_design(V, T, V, T, TF, [0], loads, tmp_path / "b", h=2.0, h_bc=1e-3 / 0.75, fallback="skfem")
    assert out.engine == "skfem-P2" and out.compliance > 0 and "FALLING BACK" in caplog.text


def test_compare_tolerances_and_csv(tmp_path):
    vox = {"solid": dict(compliance=10.0, lambda_b=5.0, sigma_q=30.0), "design": dict(compliance=20.0, lambda_b=3.0, sigma_q=36.0)}
    bf = {"solid": dict(compliance=10.2, lambda_b=5.3, sigma_q=31.0), "design": dict(compliance=20.8, lambda_b=3.1, sigma_q=44.0)}
    rows = vd.compare(vox, bf)
    by = {(r["metric"], r["case"]): r for r in rows}
    assert by[("compliance_Nmm", "solid")]["passed"] and by[("lambda_b", "solid")]["passed"]       # 2 %, 6 %
    assert not by[("sigma_q995_MPa", "design")]["passed"]                                          # +22 %
    assert abs(by[("K/K0", "ratio")]["bodyfitted"] - 10.2 / 20.8) < 1e-12
    assert by[("K/K0", "ratio")]["blocking"] and abs(by[("K/K0", "ratio")]["diff"] - (10.2 / 20.8 - 0.5)) < 1e-12
    p = vd.write_validation_csv(rows, tmp_path / "validation.csv")
    txt = p.read_text().splitlines()
    assert txt[0].startswith("metric,case") and len(txt) == len(rows) + 1
