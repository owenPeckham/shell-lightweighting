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
        vd.validate_design(V, T, V, T, TF, [0], loads, tmp_path / "a", h=2.0, want_stress=False, engine="ccx",
                           backend="gmsh", enforce_gate=False)
    with caplog.at_level("WARNING", logger="cshell.validate"):
        out = vd.validate_design(V, T, V, T, TF, [0], loads, tmp_path / "b", h=2.0, h_bc=1e-3 / 0.75, fallback="skfem",
                                engine="ccx", backend="gmsh", enforce_gate=False)
    assert out.engine == "skfem-P2" and out.compliance > 0 and "FALLING BACK" in caplog.text


def test_compare_tolerances_and_csv(tmp_path):
    vox = {"solid": dict(compliance=10.0, lambda_b=5.0, sigma_q=30.0), "design": dict(compliance=20.0, lambda_b=3.0, sigma_q=36.0)}
    bf = {"solid": dict(compliance=10.2, lambda_b=5.3, sigma_q=31.0), "design": dict(compliance=20.8, lambda_b=3.1, sigma_q=44.0)}
    rows = vd.compare(vox, bf)
    by = {(r["metric"], r["case"]): r for r in rows}
    assert by[("compliance_Nmm", "solid")]["passed"] and by[("lambda_b", "solid")]["passed"]       # 2 %, 6 %
    assert not by[("sigma_q995_MPa", "design")]["passed"]                                          # +22 %
    assert abs(by[("K/K0", "ratio")]["bodyfitted"] - 10.2 / 20.8) < 1e-12
    assert by[("K/K0", "ratio")]["blocking"] and abs(by[("K/K0", "ratio")]["diff"] - (10.2 / 20.8 / 0.5 - 1)) < 1e-12 \
        and by[("K/K0", "ratio")]["kind"] == "rel"
    p = vd.write_validation_csv(rows, tmp_path / "validation.csv")
    txt = p.read_text().splitlines()
    assert txt[0].startswith("metric,case") and len(txt) == len(rows) + 1


def test_ratio_tolerance_is_relative():
    """§8.6: K/K0 = 3.2 -> 3.32 (+3.8 %) passes; the old absolute 0.05 rule would fail it. +6 % fails."""
    vox = {"solid": dict(compliance=3.2, lambda_b=5.0, sigma_q=30.0), "design": dict(compliance=1.0, lambda_b=5.0, sigma_q=30.0)}
    ok = {"solid": dict(compliance=3.32, lambda_b=5.0, sigma_q=30.0), "design": dict(compliance=1.0, lambda_b=5.0, sigma_q=30.0)}
    bad = {"solid": dict(compliance=3.2 * 1.06, lambda_b=5.0, sigma_q=30.0), "design": dict(compliance=1.0, lambda_b=5.0, sigma_q=30.0)}
    r = {x["metric"]: x for x in vd.compare(vox, ok) if x["case"] == "ratio"}["K/K0"]
    assert r["passed"] and abs(r["voxel"] - 3.2) < 1e-12 and abs(r["diff"] - 0.0375) < 1e-9
    assert not {x["metric"]: x for x in vd.compare(vox, bad) if x["case"] == "ratio"}["K/K0"]["passed"]


# -- remesh, quality gate (CPU) ------------------------------------------------------------------------------------------
def tube_stl(ro, ri, L, sections=96):
    m = trimesh.creation.annulus(r_min=ri, r_max=ro, height=L, sections=sections)
    return np.asarray(m.vertices), np.asarray(m.faces)


def shell_stl(ro=8.0, ri=7.0, sub=5):
    o = trimesh.creation.icosphere(subdivisions=sub, radius=ro)
    i = trimesh.creation.icosphere(subdivisions=sub, radius=ri)
    i.invert()
    m = trimesh.util.concatenate([o, i])
    return np.asarray(m.vertices), np.asarray(m.faces)


def _mesh(V, T, h):
    V1, T1, info = vd.remesh_surface(V, T, h)
    nodes, tets = vd._gmsh_inside(V1, T1, h)
    return info, nodes, tets


def test_remesh_and_mesh_gate_thin_wall():
    """Spherical shell, wall 1.0 mm: h = 0.3 passes (>= 2 tets through the wall, dihedral > 0.25 deg, edge <= 2.5 h);
    h = 0.5 fails on the layers criterion (RuntimeError); enforce=False returns the numbers without raising.
    (Measured: layers_p01 2.18 and 1.13, min dihedral 21 / 11.6 deg.)"""
    V, T = shell_stl()
    for h, expect in ((0.3, True), (0.5, False)):
        info, nodes, tets = _mesh(V, T, h)
        assert info["watertight"] and info["self_intersections"] == 0 and abs(info["vol_rel"]) < 1e-3
        n10, t10 = vd.to_quadratic(nodes, tets)
        q = vd.mesh_gate(n10, t10, V, T, h=h, enforce=False)
        assert q["passed"] == expect, q
        if not expect:
            assert any("through the wall" in f for f in q["failures"])
            with pytest.raises(RuntimeError, match="quality gate"):
                vd.mesh_gate(n10, t10, V, T, h=h)
        else:
            assert q["layers_p01"] >= 2.0 and q["dihedral_min"] > 0.25 and q["edge_max_over_h"] <= 2.5


def test_mesh_gate_rejects_sliver():
    """One near-flat tet (apex pushed onto the base plane) fails the minimum dihedral criterion."""
    V, T = shell_stl()
    _, nodes, tets = _mesh(V, T, 0.3)
    nodes = nodes.copy()
    a, b, c, d = tets[0]
    nodes[d] = (nodes[a] + nodes[b] + nodes[c]) / 3 + 1e-4 * np.cross(nodes[b] - nodes[a], nodes[c] - nodes[a])
    n10, t10 = vd.to_quadratic(nodes, vd._orient(nodes, tets))
    q = vd.mesh_gate(n10, t10, V, T, h=0.3, enforce=False)
    assert not q["passed"] and any("min dihedral" in f for f in q["failures"]), q


def test_remesh_unwelded_stl_and_long_edges():
    """An STL with duplicated vertices is welded first (else the remesher is a silent no-op leaving 20 mm edges)."""
    m = trimesh.creation.annulus(r_min=4.0, r_max=5.0, height=20.0, sections=96)
    V, T = np.asarray(m.vertices), np.asarray(m.faces)
    try:
        info = vd.remesh_surface(V, T, 0.5)[2]
    except RuntimeError as e:                    # acceptable: a loud failure, never a silent coarse surface
        assert "surface remesh check failed" in str(e)
    else:
        assert info["edge_max"] <= 2.5 * 0.5 and info["watertight"]


# -- GPU engine (needs a GPU: cuDSS) -----------------------------------------------------------------------------------
def _has_gpu():
    try:
        import cupy as cp
        return cp.cuda.runtime.getDeviceCount() > 0
    except Exception:
        return False


gpu = pytest.mark.skipif(not _has_gpu(), reason="no GPU")


@gpu
def test_gpu_engine_equals_ccx_same_deck(tmp_path):
    """Same C3D10 mesh and BCs: GPU (cuDSS) vs CalculiX. Axial compression of a rod r = 5, L = 60: C, lambda_1..3, sigma_q
    (2 mm exclusion) agree to 1e-5 relative (lambda without K_sigma exclusion, as ccx has none)."""
    from cshell import bodyfit_gpu as bg
    P = 100.0
    V, T, TF, case, _ = rod_case(5.0, 60.0, 1.6, [0, 0, -P])
    inp = vd.write_inp(tmp_path / "x.inp", case, n_modes=3, want_stress=True)
    run = vd.run_ccx(inp, nproc=8)
    assert run.returncode == 0, run.log.read_text()[-500:]
    res = vd.parse_dat(run.dat)
    C, sq, vol = vd.responses_from_ccx(case, res)
    o = bg.solve(case.nodes, case.tets10, case.fixed, case.load_nodes, case.load_vec, vd.E_MPA, vd.NU, 3, exclude_bc_mm=2.0)
    assert abs(o["C"] / C - 1) < 1e-5, (o["C"], C)
    assert abs(o["sigma_q"] / sq - 1) < 1e-5, (o["sigma_q"], sq)
    np.testing.assert_allclose(o["lam_unexcl"], res.buckling[:3], rtol=1e-5)
    assert abs(o["volume"] / vol - 1) < 1e-9


@gpu
def test_gpu_column_buckling_and_ks_exclusion():
    """Clamped-free rod: lambda vs pi^2 EI/(4L^2) with shear correction (3 %); excluding K_sigma within 2 mm of the clamp
    changes lambda by < 2 % on a uniform column (nothing singular there) and never lowers it; n_ks_excluded > 0."""
    from cshell import bodyfit_gpu as bg
    r, L, P = 5.0, 100.0, 100.0
    V, T, TF, case, _ = rod_case(r, L, 1.4, [0, 0, -P])
    o = bg.solve(case.nodes, case.tets10, case.fixed, case.load_nodes, case.load_vec, vd.E_MPA, vd.NU, 3, exclude_bc_mm=2.0)
    I, A, G = np.pi * r ** 4 / 4, np.pi * r * r, E / (2 * (1 + NU))
    Pe = np.pi ** 2 * E * I / (4 * L ** 2)
    Pcr = Pe / (1 + Pe / (0.9 * G * A))
    assert abs(o["lam"][0] * P / Pcr - 1) < 0.03, (o["lam"][0] * P, Pcr)
    assert o["n_ks_excluded"] > 0 and o["lam"][0] >= o["lam_unexcl"][0] * (1 - 1e-9)
    assert abs(o["lam"][0] / o["lam_unexcl"][0] - 1) < 0.02
    o0 = bg.solve(case.nodes, case.tets10, case.fixed, case.load_nodes, case.load_vec, vd.E_MPA, vd.NU, 3, exclude_bc_mm=0.0)
    assert o0["n_ks_excluded"] == 0 and abs(o0["lam"][0] / o["lam_unexcl"][0] - 1) < 1e-8


@gpu
def test_validate_design_gpu_default_chain(tmp_path):
    """validate_design defaults (remesh + gmsh inside + gate + GPU engine): cantilever rod r = 5, L = 40, tip shear
    100 N vs Timoshenko (3 %); gate numbers, both lambda variants and the engine tag recorded."""
    r, L, P = 5.0, 40.0, 100.0
    V, T, _ = rod_stl(r, L)
    m = trimesh.Trimesh(V, T, process=True).subdivide_to_size(max_edge=1.2)     # fan caps -> sensibly sized triangles
    V, T = np.asarray(m.vertices), np.asarray(m.faces)
    zc = V[T].mean(1)[:, 2]
    TF = np.where(zc < 1e-6, 0, np.where(zc > L - 1e-6, 1, 2))
    out = vd.validate_design(V, T, V, T, TF, [0], [dict(faces=[1], force=[P, 0, 0])], tmp_path, h=0.8, h_bc=0.2)
    assert out.engine == "gpu-P2" and out.info["mesh_quality"]["passed"]
    A, I, G = np.pi * r * r, np.pi * r ** 4 / 4, E / (2 * (1 + NU))
    k = 6 * (1 + NU) / (7 + 6 * NU)
    C_tim = P * P * (L ** 3 / (3 * E * I) + L / (k * G * A))
    assert abs(out.compliance / C_tim - 1) < 0.03, (out.compliance, C_tim)
    assert np.isfinite(out.lambda_b) and "lam_unexcl" in out.info and out.sigma_q > 0
