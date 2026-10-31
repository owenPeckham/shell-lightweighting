"""cshell.collect (THEORY §5.4, §7.4 S2, §8): decision rule, uniform-shell thickness, re-scoring identities, the
dshell start mapping and an end-to-end run + collect on the B2 rod (CPU, coarse grids)."""
import json
import math
import os
from pathlib import Path

import numpy as np
import pytest

ROOT = Path(__file__).resolve().parents[2]
CRIT = dict(g_min=0.02, f_K=0.85, f_P=0.85)
LN = math.log(1.02)


# ---- decision rule ---------------------------------------------------------------------------------------------
@pytest.mark.parametrize("J,K,P,export_ok,blocking,expect", [
    (LN + 1e-3, 0.9, 0.9, True, False, True),
    (LN, 0.9, 0.9, True, False, False),            # strict: J must exceed ln(1 + g_min)
    (0.2, 0.84, 1.5, True, False, False),
    (0.2, 0.85, 0.85, True, False, True),          # floors are inclusive
    (0.2, 1.2, 0.849, True, False, False),
    (0.2, 1.2, 1.2, False, False, False),
    (0.2, 1.2, 1.2, True, True, False),
])
def test_decision_truth_table(J, K, P, export_ok, blocking, expect):
    from cshell.collect import decide
    acc, reason, out = decide([dict(res="opt", J=J, K_K0=K, P_P0=P)], CRIT, export_ok, blocking)
    assert acc is expect
    assert reason.startswith("cshell design accepted" if expect else "original solid kept")


def test_decision_every_resolution():
    from cshell.collect import decide
    ok = dict(J=0.1, K_K0=1.0, P_P0=1.0)
    acc, reason, out = decide([dict(res="opt", **ok), dict(res="shift", **ok), dict(res="coarse", J=0.01, K_K0=1, P_P0=1)],
                              CRIT)
    assert not acc and "coarse" in reason and out[2]["pass_J"] is False and out[0]["pass_J"] is True
    assert not decide([], CRIT)[0]


# ---- score: identities and the no-FE sensitivity scalings -----------------------------------------------------
def test_score_identities():
    from cshell.collect import score
    from cshell.objective import Settings
    st = Settings()
    raw0 = dict(C=2.0, sigma_pn=10.0, sigma_q=12.0, mu_tilde=0.2)
    s0 = score(raw0, raw0, st)
    assert abs(s0["J"]) < 1e-15 and s0["K_K0"] == 1.0 and abs(s0["P_P0"] - 1) < 1e-15
    raw = dict(C=1.0, sigma_pn=20.0, sigma_q=0.0, mu_tilde=1e-9)     # strength-governed, stiffer
    s = score(raw, raw0, st)
    assert abs(s["K_K0"] - 2.0) < 1e-15 and s["governing"] == "strength"
    # λ_s ∝ σ_f: with both strength-governed in the limit, P/P0 is σ_f independent
    a = score(dict(raw, mu_tilde=0.0), dict(raw0, mu_tilde=0.0), st, sigma_f=20.0)
    b = score(dict(raw, mu_tilde=0.0), dict(raw0, mu_tilde=0.0), st, sigma_f=40.0)
    assert abs(a["P_P0"] - 0.5) < 1e-12 and abs(b["P_P0"] - 0.5) < 1e-12
    assert abs(score(raw0, raw0, st, kappa=0.6)["lam_b"] - 0.6 / 0.2) < 1e-12


# ---- uniform shell thickness (§7.4 S1): sphere, Vol{0 ≤ φ ≤ t} = Vol{φ ≤ 0} → t = R(2^{1/3} − 1) -----------
def test_uniform_shell_thickness_sphere():
    import trimesh
    from cshell.export import solve_shell_thickness
    R = 10.0
    m = trimesh.creation.icosphere(subdivisions=5, radius=R)
    t = solve_shell_thickness(np.asarray(m.vertices), np.asarray(m.faces), h_c=0.1, workers=8)
    exact = R * (2 ** (1 / 3) - 1)
    assert abs(t / exact - 1) < 5e-3, (t, exact)


# ---- B2 rod: re-scoring identities -----------------------------------------------------------------------------
B2_SET = {"grid": {"phases": [{"h_e": 0.8, "h_g": 0.2}], "export_h_mm": 0.2}, "fe": {"backend": "cpu_direct"},
          "optimiser": {"starts": ["uniform", "neutral"], "max_iter": [2]},
          "validation": {"body_fitted": False, "coarse_grid": {"h_e": 1.0, "h_g": 0.25}}}


@pytest.fixture(scope="module")
def b2(tmp_path_factory):
    from cshell import config
    from cshell.optimise import Context
    out = tmp_path_factory.mktemp("b2")
    cfg = config.load(str(ROOT / "examples/cshell/bench/B2_slender.json"), {**B2_SET, "output_dir": str(out)})
    return cfg, Context(cfg)


def test_rescore_solid_equals_baseline(b2):
    from cshell import collect as C
    cfg, ctx = b2
    r = C.resolutions(cfg)
    assert [x["res"] for x in r] == ["opt", "shift", "coarse"]
    m = C.make_model(ctx, r[0], 4.0)
    st = C.scoring_settings(m)
    b_o = cfg["manufacturing"]["outer_bias_mm"]
    rho, V = C.cell_rescore(m, None, None, b_o, solid=True)
    assert np.allclose(rho, C._h(m.rho0), rtol=0, atol=1e-12) and abs(V / float(C._h(m.V0)) - 1) < 1e-12
    raw = C.raw_responses(m, rho, st)
    base = m.baseline(st)
    assert abs(raw["C"] / base["C0"] - 1) < 1e-10
    if st.w_P > 0:                                       # B2 is compliance-only: the baseline has no c then
        assert abs(raw["sigma_q"] / raw["sigma_pn"] - base["c"]) < 1e-8 * base["c"]
    sc = C.score(raw, raw, st)
    assert abs(sc["J"]) < 1e-14
    # the mesh-SDF path (unwelded triangle soup) reproduces φ on the original tessellation
    phi = C._h(m.disc.phi).astype(np.float64)
    soup = (ctx.V[ctx.T].reshape(-1, 3) - m.band.c0).astype(np.float32)
    phiE = C.mesh_phi(m, soup)
    near = np.abs(phi) < 2.0
    assert np.max(np.abs(C._h(phiE)[near] - phi[near])) < 1e-3
    # Δ: a design with a = 1e4 everywhere (void branch never active) and b_o = 0 gives Δ = 0: ρ(φ_E + Δ) = ρ(φ_E)
    nv = ctx.surf.n_vertices
    f = C.design_fields(np.zeros(nv), np.full(nv, 1e4), np.zeros(nv, bool), ctx.r_ma, 0.2)
    rD, VD = C.cell_rescore(m, phiE, f, 0.0)
    r0, V0E = C.cell_rescore(m, phiE, None, 0.0, solid=True)
    assert np.allclose(rD, r0, rtol=0, atol=1e-12) and abs(VD / V0E - 1) < 1e-12
    # shifted grid: origin moved by ½h_e, solid K within a few % of the unshifted value
    ms = C.make_model(ctx, r[1], 4.0)
    raws = C.raw_responses(ms, ms.rho0, C.scoring_settings(ms))
    assert abs(raws["C"] / raw["C"] - 1) < 0.05


def test_design_bcs_islands_and_isolated_loads(b2):
    """P1 / P4(i): material-aware BCs. The solid keeps the PhaseModel BCs exactly; a design with an island and
    with the load face emptied drops the island, moves the load to nodes that touch material (force conserved) and
    leaves no loaded node without material."""
    from cshell import collect as C
    cfg, ctx = b2
    m = C.make_model(ctx, C.resolutions(cfg)[0], 4.0)
    bcs = C.DesignBCs(m, ctx)
    assert bcs.mask_mismatch == 0
    rho0 = C._h(m.rho0).astype(np.float64)
    r2, fixed, load, info = bcs.adjust(rho0)
    assert r2 is rho0 and np.array_equal(fixed, m.fixed) and np.array_equal(load, m.load) and info["islands"] == 0
    es = m.elem_shape
    rho_isl = rho0.reshape(es).copy()
    # an island: a 2³ material block in an empty corner of the element box (no material within one element)
    assert not np.any(rho_isl[:3, :3, :3] > 0)
    rho_isl[:2, :2, :2] = 1.0
    # empty every element next to every 4th load node (scattered: the part stays connected, the rest still touch)
    ln = bcs.loads[0][0]
    half = np.zeros(int(np.prod(m.node_shape)), bool)
    half[ln[::4]] = True
    near = C._elem_touch(half.reshape(m.node_shape))
    rho_isl[near] = 0.0
    r2, fixed, load, info = bcs.adjust(rho_isl.ravel())
    print("DesignBCs info", info)
    assert info["isolated_load_nodes"] == 0
    assert np.allclose(load.reshape(-1, 3).sum(0), ctx.F, rtol=1e-12, atol=1e-12)          # force conserved
    assert info["load_share_redistributed"] > 0.1 and sum(info["load_nodes_dropped"]) >= len(ln[::4])
    assert info["islands"] == 1 and abs(info["island_mm3"] - 8 * m.h_e ** 3) < 1e-9
    assert np.all(r2.reshape(es)[:2, :2, :2] == 0) and r2 is not rho_isl
    bcs.apply(fixed, load)
    assert m.fe is None and m.grid_fe is None and m.load is load
    bcs.apply(bcs.fixed0, bcs.load0)
    assert m.load is bcs.load0 and m.stress_mask_dense is bcs.mask0


def test_stl_soup_weld_roundtrip(tmp_path):
    """Binary STL reader (no trimesh processing) and the exact-coordinate weld reproduce trimesh's welded mesh."""
    import trimesh
    from cshell import collect as C
    from cshell.export import write_stl
    m = trimesh.creation.icosphere(subdivisions=3, radius=5.0)
    p = tmp_path / "s.stl"
    write_stl(str(p), np.asarray(m.vertices), np.asarray(m.faces))
    soup = C.stl_soup(str(p))
    assert soup.shape == (3 * len(m.faces), 3) and soup.dtype == np.float32
    V, T, vol = C.load_mesh(str(p))
    tm = trimesh.load(str(p), process=True)
    assert len(V) == len(tm.vertices) and len(T) == len(tm.faces)
    assert abs(vol / float(tm.volume) - 1) < 1e-12
    c0 = np.array([0.5, -0.25, 1.0])
    sc = C.stl_soup(str(p), c0)
    assert np.array_equal(sc, (soup.astype(np.float64) - c0).astype(np.float32))


def test_export_verdict_and_early_reject():
    """D2 / P3: a failed export (MBB smoke numbers, F3) is the decision reason; nothing is re-scored."""
    from cshell.collect import decide, export_verdict
    ex = dict(ok=False, V0_mm3=85230.7, cleanup=dict(opening_removed_mm3=9486.55),
              mass_correction=dict(ok=False, fail=False, rel_before=0.14496, delta=0.0,
                                   bracket_failed=dict(limit=0.35)),
              mesh=dict(watertight_manifold=False, bodies=39, duplicate_directed_edges=1890, volume_rel=0.14546))
    ok, why = export_verdict(ex, G_s=0.05)
    assert not ok and "did not bracket" in why and "non-convergence" in why and "+14.55%" in why
    acc, reason, out = decide([], CRIT, False, False, export_why=why)
    assert not acc and reason.startswith("original solid kept: export failed: mass correction")
    assert "no feasible" not in reason
    good = dict(ok=True, V0_mm3=1.0, mass_correction=dict(ok=True, fail=False), cleanup={},
                mesh=dict(watertight_manifold=True, volume_rel=1e-4))
    assert export_verdict(good) == (True, "")
    assert not export_verdict(dict(good, mesh=dict(watertight_manifold=True, volume_rel=3e-3)))[0]


def test_cell_rescore_gpu_matches_cpu(b2):
    """P5a: the device re-scoring cell map equals the host one (ρ to 1e-12, V to 1e-12, J to 1e-10)."""
    cp = pytest.importorskip("cupy")
    try:
        cp.cuda.runtime.getDeviceCount()
    except Exception:
        pytest.skip("no CUDA device")
    import copy
    from cshell import collect as C
    cfg, ctx = b2
    out = {}
    for dev in ("cpu", "gpu"):
        c2 = copy.copy(ctx)
        c2.cfg = copy.deepcopy(cfg)
        c2.cfg["grid"]["geometry_device"] = dev
        m = C.make_model(c2, C.resolutions(cfg)[0], 4.0)
        assert m.geometry_device == dev
        nv = ctx.surf.n_vertices
        rng = np.random.default_rng(0)
        f = C.design_fields(rng.uniform(0, 2, nv), rng.uniform(0, 3, nv), rng.uniform(0, 1, nv) > 0.7, ctx.r_ma, 0.2)
        soup = (ctx.V[ctx.T].reshape(-1, 3) - m.band.c0).astype(np.float32)
        phiE = C.mesh_phi(m, soup)
        rho_s, V_s = C.cell_rescore(m, None, None, 0.1, solid=True)
        rho_p, V_p = C.cell_rescore(m, None, f, 0.1)
        rho_e, V_e = C.cell_rescore(m, phiE, f, 0.1)
        if dev == "gpu":
            m.disc.release()
        st = C.scoring_settings(m)
        r0 = C.raw_responses(m, rho_s, st)
        re = C.raw_responses(m, rho_e, st)
        out[dev] = dict(rho=(rho_s, rho_p, rho_e), V=(V_s, V_p, V_e), J=C.score(re, r0, st)["J"])
        C.free_model(m)
    for k in range(3):
        assert np.max(np.abs(out["cpu"]["rho"][k] - out["gpu"]["rho"][k])) < 1e-12, k
        assert abs(out["cpu"]["V"][k] / out["gpu"]["V"][k] - 1) < 1e-12, k
    assert abs(out["cpu"]["J"] - out["gpu"]["J"]) < 1e-10, (out["cpu"]["J"], out["gpu"]["J"])


def test_end_to_end_run_collect(b2, tmp_path):
    from cshell import cli
    from cshell.collect import read_vtp_point_data
    cfg, _ = b2
    cfg_path = tmp_path / "b2.json"
    json.dump({**cfg, "output_dir": str(tmp_path)}, open(cfg_path, "w"))
    for s in ("uniform", "neutral"):
        assert cli.main(["run", str(cfg_path), "--start", s]) == 0
    rd = tmp_path / cfg["name"]
    assert (rd / "done_uniform.json").exists()
    assert cli.main(["collect", str(rd)]) == 0
    S = json.load(open(rd / "summary.json"))
    for k in ("schema", "name", "accepted", "outcome", "reason", "criteria", "mass", "best", "starts", "comparison",
              "resolutions", "checks", "load_cases", "Lambda", "export", "validation", "sensitivity", "dshell",
              "resources", "timings_s", "files"):
        assert k in S, k
    assert S["schema"] == "cshell.summary/1"
    print("e2e reason:", S["reason"], "export:", S["export"]["best_ok"], S["export"]["best_reject_reason"],
          S["export"]["closure_rule_band"])
    assert S["export"]["best_ok"] is True, S["export"]["best_reject_reason"]
    assert S["export"]["uniform_shell_mounts_solid"] is True             # D1
    # P2: the export closure rule reproduces the optimiser's band volume
    assert abs(S["export"]["closure_rule_band"]["rule_vs_opt"]) < 2e-3
    for r in S["resolutions"]:                                           # P1: no loaded node without material
        assert r["bc_isolated_load_nodes"] == 0 and r["error"] is None, r
    assert {r["res"] for r in S["resolutions"]} == {"opt", "shift", "coarse"}
    assert {r["design"] for r in S["resolutions"]} == {"solid", "uniform_shell", "best"}
    assert [c["label"] for c in S["comparison"]][:2] == ["solid", "uniform shell"]
    assert len(S["checks"]) == 3 and S["best"]["start"] in ("uniform", "neutral")
    solid = [r for r in S["resolutions"] if r["design"] == "solid"]
    assert all(abs(r["J"]) < 1e-12 for r in solid)
    us = [r for r in S["resolutions"] if r["design"] == "uniform_shell" and r["res"] == "opt"][0]
    assert abs(us["mass_err"]) < 0.03                                   # constant mass on the FE grid
    for f in ("final_design.stl", "best_candidate.stl", "original.stl", "uniform_shell.stl", "fields.vtp",
              "sensitivity.csv", "buckling_modes.npz", "bc_faces.npz", "summary.txt"):
        assert (rd / f).exists(), f
    pd = read_vtp_point_data(rd / "fields.vtp")
    for k in ("xi", "d", "a", "w", "r_MA", "d_hi", "passive", "stress", "stress_solid", "W_bar"):
        assert k in pd, k
    import csv
    rows = list(csv.DictReader(open(rd / "sensitivity.csv")))
    assert len(rows) == 20 and {"kappa_b", "sigma_f", "J_star", "accepted"} <= set(rows[0])
    fin = (rd / "final_design.stl").read_bytes()
    assert fin == (rd / ("best_candidate.stl" if S["accepted"] else "original.stl")).read_bytes()


# ---- S2 dshell start on the cantilever (§7.4) --------------------------------------------------------------------
def test_dshell_start_cantilever():
    from cshell import config
    from cshell.collect import load_dshell, dshell_run_dir, dshell_start_xi
    from cshell.optimise import Context
    cfg = config.load(str(ROOT / "examples/cshell/cantilever.json"))
    vs = load_dshell(dshell_run_dir(cfg))
    assert vs["hollow"] == {7, 8} and abs(vs["s"] - 1.9127) < 1e-3
    ctx = Context(cfg)
    lo, hi = ctx.V.min(0), ctx.V.max(0)
    assert np.all(vs["xyz"].min(0) >= lo - 1) and np.all(vs["xyz"].max(0) <= hi + 1)   # same frame as the STEP
    xi = dshell_start_xi(ctx, cfg)
    nv = ctx.surf.n_vertices
    assert xi.shape == (2 * nv,) and np.all((xi >= 0) & (xi <= 1))
    from cshell import config as ccfg
    filt = ctx.surf.filter(ccfg.phase_value(cfg["filter"]["radius_mm"], 0))
    xt_a = np.clip(filt.apply(xi[nv:]), 0, 1)
    A = ctx.surf.A
    open_share = float((A * (xt_a < 0.5)).sum() / A.sum())
    print("dshell start: open area share", open_share)
    assert 0.02 < open_share < 0.8                       # two hollow struts of 14 regions
