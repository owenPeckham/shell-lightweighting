"""cshell.collect (THEORY §5.4, §7.4 S2, §8): decision rule, uniform-shell thickness, re-scoring identities, the
vshell start mapping and an end-to-end run + collect on the B2 rod (CPU, coarse grids)."""
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
    cd = C.CellData(m, ctx)
    _, rho = C.chi_rho(cd, cd.phi + cfg["manufacturing"]["outer_bias_mm"])
    raw = C.raw_responses(m, rho, st)
    base = m.baseline(st)
    assert abs(raw["C"] / base["C0"] - 1) < 1e-10
    if st.w_P > 0:                                       # B2 is compliance-only: the baseline has no c then
        assert abs(raw["sigma_q"] / raw["sigma_pn"] - base["c"]) < 1e-8 * base["c"]
    sc = C.score(raw, raw, st)
    assert abs(sc["J"]) < 1e-14
    # the mesh-SDF path reproduces φ on the original tessellation
    phiE = C.mesh_phi(cd, ctx.V, ctx.T)
    near = np.abs(cd.phi) < 2.0
    assert np.max(np.abs(phiE[near] - cd.phi[near])) < 1e-3
    # Δ: solid design (a = ∞) gives b_o everywhere
    nv = ctx.surf.n_vertices
    D = C.asbuilt_increment(cd, np.zeros(nv), np.full(nv, 1e4), ctx.r_ma, 0.2, 0.0)
    assert np.all(D == 0.0)
    # shifted grid: origin moved by ½h_e, solid K within a few % of the unshifted value
    ms = C.make_model(ctx, r[1], 4.0)
    raws = C.raw_responses(ms, ms.rho0, C.scoring_settings(ms))
    assert abs(raws["C"] / raw["C"] - 1) < 0.05


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
              "resolutions", "checks", "load_cases", "Lambda", "export", "validation", "sensitivity", "vshell",
              "resources", "timings_s", "files"):
        assert k in S, k
    assert S["schema"] == "cshell.summary/1"
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


# ---- S2 vshell start on the cantilever (§7.4) --------------------------------------------------------------------
def test_vshell_start_cantilever():
    from cshell import config
    from cshell.collect import load_vshell, vshell_run_dir, vshell_start_xi
    from cshell.optimise import Context
    cfg = config.load(str(ROOT / "examples/cshell/cantilever.json"))
    vs = load_vshell(vshell_run_dir(cfg))
    assert vs["hollow"] == {7, 8} and abs(vs["s"] - 1.9127) < 1e-3
    ctx = Context(cfg)
    lo, hi = ctx.V.min(0), ctx.V.max(0)
    assert np.all(vs["xyz"].min(0) >= lo - 1) and np.all(vs["xyz"].max(0) <= hi + 1)   # same frame as the STEP
    xi = vshell_start_xi(ctx, cfg)
    nv = ctx.surf.n_vertices
    assert xi.shape == (2 * nv,) and np.all((xi >= 0) & (xi <= 1))
    from cshell import config as ccfg
    filt = ctx.surf.filter(ccfg.phase_value(cfg["filter"]["radius_mm"], 0))
    xt_a = np.clip(filt.apply(xi[nv:]), 0, 1)
    A = ctx.surf.A
    open_share = float((A * (xt_a < 0.5)).sum() / A.sum())
    print("vshell start: open area share", open_share)
    assert 0.02 < open_share < 0.8                       # two hollow struts of 14 regions
