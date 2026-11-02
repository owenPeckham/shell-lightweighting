"""Growth loop on the tiny case (PLAN_single_run §4, §6) and the C2 contract (§7): B3 column (rod r = 5 mm, L = 100 mm,
axial compression), one phase at h_e = 0.8 mm, solid start, three growth steps of two MMA iterations.
GPU only (one module-scoped run, ~minutes)."""
import csv
import json
from pathlib import Path

import numpy as np
import pytest

ROOT = Path(__file__).resolve().parents[2]
pytestmark = pytest.mark.gpu

C2_COLS = ["it", "phase", "growth_step", "R_k", "J", "K_K0", "lamB_ratio", "sig_ratio", "mass_err", "closed_frac",
           "hollow_allowed_frac"]


def _cfg(extra):
    from cshell import config
    # the C2 compliance contract (§4) is tested here; the §8 weighted default in test_weighted_floors_run
    base = {"grid": {"phases": [{"h_e": 0.8, "h_g": 0.2}]}, "optimiser": {"starts": ["solid"]},
            "objective": {"type": "compliance", "constraints": ["buckling", "strength"]}}
    return config.load(str(ROOT / "examples/cshell/bench/B3_column.json"), {**base, **extra})


def test_weighted_floors_run(tmp_path):
    """§8 defaults: weighted objective with the K/K0 and P/P0 floors as MMA constraints; every quantile runs
    (patience None), then the best boundary design; evaluate() scores the run's own objective."""
    from cshell import config
    from cshell.optimise import Context, Run, evaluate, load_design
    cfg = config.load(str(ROOT / "examples/cshell/bench/B3_column.json"), {
        "grid": {"phases": [{"h_e": 0.8, "h_g": 0.2}]},
        "growth": {"quantiles": [0.3, 0.6, 1.0], "max_iters_per_step": 3, "tol": 10.0}})   # every gain "low"
    assert cfg["objective"]["type"] == "weighted" and cfg["growth"]["patience"] is None
    ctx = Context(cfg)
    res = Run(cfg, "solid", str(tmp_path)).run(ctx)
    g = res["phase_0"]["growth"]
    assert len(g["J_b"]) == 4 and g["steps"] == 3                      # all quantiles despite low gains
    with open(tmp_path / "history.csv") as f:
        hist = list(csv.DictReader(f))
    assert all(np.isfinite(float(r["P_P0"])) and np.isfinite(float(r["K_K0"])) for r in hist)
    dsg = load_design(str(tmp_path))
    r = evaluate(ctx, dsg["xi"], 0, hollow=dsg["hollow"], D_comp=dsg["D_comp"])
    assert r["objective"] == "weighted" and set(r["floors"]) == {"K", "P"}
    assert r["J"] == pytest.approx(0.5 * np.log(r["K_K0"]) + 0.5 * np.log(r["P_P0"]), abs=1e-6)
    # the run's stress normalisation c drifts (Le et al.); evaluate uses the solid's c: J agrees to the c effect
    assert r["J"] == pytest.approx(g["best_J"], abs=5e-3) and r["J_K"] == pytest.approx(np.log(r["K_K0"]))
    assert r["feasible"] == (r["K_K0"] >= 0.85 * 0.99 and r["P_P0"] >= 0.85 * 0.99)    # feasibility_tol 1 %


@pytest.fixture(scope="module")
def grown(tmp_path_factory):
    from cshell.optimise import Context, Run
    # step_tol ≫ any ΔJ: each step ends at min_iters_per_step = 2 (exercises the in-step convergence rule)
    cfg = _cfg({"growth": {"quantiles": [0.3, 0.6, 1.0], "max_iters_per_step": 25, "min_iters_per_step": 2,
                           "step_tol": 1.0, "tol": -1e9}})
    ctx = Context(cfg)
    out = tmp_path_factory.mktemp("grow")
    run = Run(cfg, "solid", str(out))
    res = run.run(ctx)
    with open(out / "history.csv") as f:
        hist = list(csv.DictReader(f))
    return ctx, cfg, out, res, hist


def test_history_columns_and_growth(grown):
    ctx, cfg, out, res, hist = grown
    assert res["status"] == "done"
    assert all(c in hist[0] for c in C2_COLS)
    assert len(hist) == 6                                      # 3 steps × 2 iterations
    steps = [int(r["growth_step"]) for r in hist]
    assert steps == [0, 0, 1, 1, 2, 2]
    H = [float(r["hollow_allowed_frac"]) for r in hist]
    assert all(b >= a - 1e-12 for a, b in zip(H, H[1:])) and H[-1] > H[0] and H[-1] == pytest.approx(1.0)
    assert all(np.isfinite(float(r["R_k"])) and float(r["R_k"]) >= 0 for r in hist)   # R is recomputed per step
    assert all(abs(float(r["mass_err"])) <= cfg["optimiser"]["mass_tol"] * 1.0001 for r in hist)
    for r in hist:                                             # constraints evaluated every iteration
        assert np.isfinite(float(r["sig_ratio"])) and float(r["lamB_ratio"]) > 0
    g = res["phase_0"]["growth"]
    assert g["steps"] == 3 and len(g["J_b"]) == 4


def test_solid_start_opens_voids(grown):
    """The first iteration is (nearly) the solid; the closed share falls once H admits vertices."""
    ctx, cfg, out, res, hist = grown
    c0 = float(hist[0]["closed_frac"])
    assert abs(float(hist[0]["J"])) < 0.05
    assert min(float(r["closed_frac"]) for r in hist) < c0 or res["phase_0"]["closed_area_frac"] < c0


def test_outputs_c2(grown):
    from cshell.optimise import load_design
    ctx, cfg, out, res, hist = grown
    nv = ctx.surf.n_vertices
    assert np.load(out / "xi_solid_phase0.npy").shape == (2 * nv,)
    dsg = load_design(str(out))
    assert dsg["xi"].shape == (2 * nv,) and dsg["hollow"].shape == (nv,) and dsg["phase"] == 0
    txt = (out / "fields.vtp").read_text()
    for k in ("d", "a", "w", "r_MA", "R", "closed", "hollow_allowed", "passive", "stress", "stress_solid"):
        assert f'Name="{k}"' in txt, k
    from cshell.collect import read_vtp_point_data
    pd = read_vtp_point_data(out / "fields.vtp")
    assert np.nanmax(pd["stress_solid"]) > 0 and np.isfinite(pd["stress"]).mean() > 0.9
    assert np.all(np.isfinite(pd["R"]))
    z = np.load(out / "bc_faces.npz")
    for k in ("support_faces", "load_faces", "force_N", "V", "T", "support", "load"):
        assert k in z.files, k
    assert np.allclose(z["force_N"], [[-50.0, 0, 0]])


def test_evaluate_matches_run(grown):
    """C2 evaluate(): J of the final design equals the run's last boundary J; shifted grid gives finite results."""
    from cshell.optimise import evaluate, load_design
    ctx, cfg, out, res, hist = grown
    dsg = load_design(str(out))
    r = evaluate(ctx, dsg["xi"], 0, hollow=dsg["hollow"], D_comp=dsg["D_comp"])
    for k in ("J", "K_K0", "lamB_ratio", "sig_ratio", "mass_err", "feasible", "governing"):
        assert k in r, k
    J_last = res["phase_0"]["growth"]["best_J"]                # growth ends on the best boundary design (sr-fix 1)
    assert abs(r["J"] - J_last) < 1e-6 * max(1.0, abs(J_last)), (r["J"], J_last)
    assert abs(r["mass_err"]) < 2e-4 and r["K_K0"] == pytest.approx(np.exp(r["J"]))
    rs = evaluate(ctx, dsg["xi"], 0, shift=(0.4, 0.4, 0.4), hollow=dsg["hollow"], D_comp=dsg["D_comp"])
    assert np.isfinite(rs["J"]) and isinstance(rs["feasible"], bool) and rs["origin"] != r["origin"]


def test_growth_stops_after_patience(tmp_path):
    """sr-fix 1: tol larger than any gain and patience 2: growth stops at boundary 2 (two low-gain steps in a row),
    not at boundary 1, and the design is the best boundary design."""
    from cshell.optimise import Context, Run
    cfg = _cfg({"growth": {"quantiles": [0.3, 0.6, 1.0], "max_iters_per_step": 1, "min_iters_per_step": 1,
                           "tol": 10.0, "patience": 2}})
    ctx = Context(cfg)
    res = Run(cfg, "solid", str(tmp_path)).run(ctx)
    g = res["phase_0"]["growth"]
    assert len(g["J_b"]) == 3 and g["steps"] == 2 and g["iters"] == [1, 1]
    assert g["best_J"] == pytest.approx(max(g["J_b"])) and g["J_b"][g["best_boundary"]] == pytest.approx(g["best_J"])
    log = (tmp_path / "run_log.txt").read_text()
    assert "growth stops at boundary 2: 2 consecutive steps" in log


def test_export_volume_matches_optimiser(grown, tmp_path):
    """sr-fix 2: the exported volume before the mass correction δ equals the optimiser's volume within 0.3 % (same
    closure at ξ̃ᵃ ≥ export.closed_thr; open voids at least void_min across)."""
    from cshell import export as cex
    from cshell.design import vertex_fields_final
    from cshell.optimise import evaluate, load_design
    ctx, cfg, out, res, hist = grown
    nv = ctx.surf.n_vertices
    # a design with real voids: open (ξᵃ = 0.2) on the middle third of the rod, closed elsewhere, d = 0.3
    x = ctx.surf.V[:, 0]
    mid = (x > x.min() + (x.max() - x.min()) / 3) & (x < x.max() - (x.max() - x.min()) / 3)
    xi = np.concatenate([np.full(nv, 0.3), np.where(mid, 0.2, 1.0)])
    r = evaluate(ctx, xi, 0)
    h_g = float(cfg["grid"]["phases"][-1]["h_g"])
    f = vertex_fields_final(ctx, xi, h_g, None, None)
    assert np.all(f["closed"][f["xt_a"] >= cfg["export"]["closed_thr"]])
    ex = cex.export_design_gpu(ctx, xi, str(tmp_path), cfg["design"]["outward_extent_initial_mm"], name="t", h_g=h_g,
                               fields=f, closed_thr=cfg["export"]["closed_thr"], h_x=0.2,
                               holes=False)      # escape holes are export-only material removal (~35 mm³ here)
    V_ex = ex["mass_correction"]["V_before_mm3"]
    print("V0", r["V0"], "V optimiser", r["V"], "V export before δ", V_ex, "rel", V_ex / r["V"] - 1)
    assert np.mean(~f["closed"][mid & ~f["passive"]]) > 0.5                              # voids are present
    assert abs(V_ex / r["V"] - 1) < 3e-3
