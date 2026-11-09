"""PLAN_lightweight: objective.type = "lightweight" (default) minimises V/V0 subject to K/K0 ≥ m_K, P/P0 ≥ m_P.

CPU: settings and floors, config validation, growth best-boundary rule, decision rule, export target, start table,
finalise wiring (mocked evaluate/export). GPU: gradient of f0 = V/V0 against central differences, a short lightweight
run on B3 (no mass restoration, V/V0 falls, evaluate scores the lightweight objective), export δ targeting V_opt.
"""
import copy
import csv
import json
import math
from pathlib import Path
from types import SimpleNamespace

import numpy as np
import pytest

from cshell import collect as col
from cshell import config
from cshell import export as cex
from cshell import finalise as fin

ROOT = Path(__file__).resolve().parents[2]
B3 = str(ROOT / "examples/cshell/bench/B3_column.json")


# ---- CPU ----------------------------------------------------------------------------------------------------------
def test_settings_lightweight_constraints():
    from cshell.optimise import PhaseModel
    c = copy.deepcopy(config.DEFAULTS)
    assert config.is_lightweight(c) and config.floors(c) == (1.02, 1.02)
    c["objective"].update(w_capacity=0.0, floors_as_constraints=False)   # weighted-only keys do not matter
    st = PhaseModel.settings(SimpleNamespace(ctx=SimpleNamespace(cfg=c), phase=0), 0.0)
    assert st.objective == "weighted" and st.constraints == ("floor_K", "floor_P")
    assert st.w_P > 0 and st.buckling                                    # capacity (strength + buckling) evaluated
    c["objective"].update(type="weighted", w_capacity=0.5, floors_as_constraints=True)
    assert config.floors(c) == (0.85, 0.85) and not config.is_lightweight(c)


def test_config_validation():
    c = copy.deepcopy(config.DEFAULTS)
    c.update(step_file="x.step", supports=[{"type": "plane"}], loads=[{"faces": {}, "force_N": [0, 0, 1]}])
    config.validate(c)
    c["objective"]["lw_margin_capacity"] = 0.0
    with pytest.raises(ValueError):
        config.validate(c)
    c["objective"].update(lw_margin_capacity=1.02, type="lightest")
    with pytest.raises(ValueError):
        config.validate(c)


def test_boundary_better():
    from cshell.optimise import boundary_better as bb
    tol = config.feasibility_log_tol(config.DEFAULTS)                   # 1 %: g ≤ 0.01005
    assert bb(0.1, 0.5, None, None, True, tol)                          # first boundary
    assert bb(0.05, 0.0, 0.0, 0.0198, True, tol)                         # feasible beats infeasible (the solid)
    assert not bb(0.40, 0.02, 0.05, 0.0, True, tol)                      # infeasible never beats feasible
    assert bb(0.30, -0.01, 0.20, 0.0, True, tol) and not bb(0.10, -0.01, 0.20, 0.0, True, tol)   # lowest V
    assert bb(0.10, 0.05, 0.30, 0.20, True, tol) and not bb(0.30, 0.25, 0.10, 0.20, True, tol)   # least infeasible
    assert bb(0.30, 9.0, 0.20, -1.0, False, tol)                         # weighted: higher J only
    assert bb(0.05, 0.005, 0.0, np.nan, True, tol)                       # nan viol counts as infeasible


def test_decide_lightweight():
    op = dict(objective="lightweight", J=0.3, K_K0=1.01, P_P0=1.005, feasible=True)
    sh = dict(objective="lightweight", J=0.29, K_K0=1.003, P_P0=1.001, feasible=True)
    acc, why = fin.decide(op, sh, True, "", 0.0)
    assert acc and "mass saving 30.0%" in why
    assert not fin.decide(op, dict(sh, P_P0=0.99, feasible=False), True, "", 0.0)[0]
    assert not fin.decide(dict(op, feasible=False), sh, True, "", 0.0)[0]
    acc, why = fin.decide(dict(op, J=-0.01), sh, True, "", 0.0)
    assert not acc and "no mass saving" in why
    assert not fin.decide(op, sh, False, "not watertight", 0.0)[0]


def test_export_target_fraction():
    assert fin.export_target_fraction(dict(V_V0=0.62, mass_err=-0.4)) == pytest.approx(0.62)
    assert fin.export_target_fraction(dict(V_V0=None, mass_err=-0.4)) == pytest.approx(0.6)
    assert fin.export_target_fraction(dict(V_V0=float("nan"), mass_err=None)) is None


def _lw_run_dir(tmp_path, nv=50):
    cfg = config.load(str(ROOT / "examples/cshell/mbb.json"))
    assert cfg["objective"]["type"] == "lightweight"                   # the example configs use the default
    cfg["output_dir"] = str(tmp_path)
    rd = tmp_path / "MBB"
    (rd / "checkpoints").mkdir(parents=True)
    json.dump(cfg, open(rd / "config_used.json", "w"), indent=1)
    cols = ["start", "phase", "growth_step", "it", "J", "K_K0", "P_P0", "lamB_ratio", "sig_ratio", "mass_err",
            "V_V0", "g_max", "gpu_peak_GiB", "t_fe"]
    rows = [["solid", p, 0, i, 0.1 * i, 1.03, 1.02, 1.1, 0.9, -0.1 * i, 1 - 0.1 * i, -0.01, 9.0, 4.0]
            for p in (0, 1) for i in range(3)]
    with open(rd / "history.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(cols)
        w.writerows(rows)
    for p in (0, 1):
        np.save(rd / f"xi_solid_phase{p}.npy", np.zeros(2 * nv))
    np.savez(rd / "checkpoints" / "solid_1_0003.npz", xi=np.zeros(2 * nv), D_comp=np.array(6.0),
             hours_used=np.array(0.5))
    json.dump({"start": "solid", "status": "done"}, open(rd / "done_solid.json", "w"))
    return cfg, rd


def test_start_table_lightweight_mass_free(tmp_path):
    cfg, rd = _lw_run_dir(tmp_path)
    r = col.start_table(str(rd), cfg, merge=False)[0]
    assert r["status"] == "done" and r["feasible"] and r["mass_err"] == pytest.approx(-0.2)   # V ≠ V0 is fine
    cfg["objective"]["type"] = "weighted"
    assert not col.start_table(str(rd), cfg, merge=False)[0]["feasible"]                     # constant mass


def test_finalise_lightweight_targets_v_opt(tmp_path, monkeypatch):
    nv = 50
    cfg, rd = _lw_run_dir(tmp_path, nv)
    calls = []

    def vertex_fields(ctx, xi, D_comp=4.0, h_g=None, **kw):
        xt = np.linspace(0.0, 1.0, nv)
        return dict(d=np.full(nv, 0.5), a=np.where(xt > 0.95, 2.5, 1.0), d_hi=np.full(nv, 3.0), xt_a=xt,
                    passive=np.arange(nv) < 5, r_ma=ctx.r_ma)

    def export_design_gpu(ctx, xi, out_dir, D_comp=4.0, name="final_design", **kw):
        calls.append(kw)
        Path(out_dir, f"{name}.stl").write_bytes(b"\0" * 84)
        Vt = 1000.0 * kw["target_fraction"]
        return dict(ok=True, V0_mm3=Vt, V_solid_mm3=1000.0, target_fraction=kw["target_fraction"],
                    mesh=dict(faces=10, watertight_manifold=True, bodies=1, volume=Vt * (1 + 1e-4), volume_rel=1e-4),
                    mass_correction=dict(ok=True, fail=False, delta=0.01, mesh_steps=[]), voids=dict(enclosed=0),
                    wall=dict(min_nominal=1.0, median_nominal=2.0, max_nominal=4.0), cleanup={}, timing_s={})

    def ev(ctx, xi, phase, shift=None, D_comp=None, **kw):
        return dict(objective="lightweight", J=0.35, K_K0=1.015, P_P0=1.004, lamB_ratio=1.2, sig_ratio=0.95,
                    mass_err=-0.35, V_V0=0.65, feasible=True, floors=dict(K=True, P=True), floor_K=1.0,
                    floor_P=1.0, governing="strength")

    monkeypatch.setattr(cex, "vertex_fields", vertex_fields)
    monkeypatch.setattr(cex, "export_design_gpu", export_design_gpu)
    ctx = SimpleNamespace(cfg=cfg, surf=SimpleNamespace(n_vertices=nv, A=np.ones(nv)), r_ma=np.full(nv, 2.0))
    S = fin.finalise(str(rd), ctx=ctx, evaluate=ev, report=False)
    assert calls[0]["target_fraction"] == pytest.approx(0.65)         # δ targets V_opt/V0, not the solid's V0
    assert S["decision"]["accepted"], S["decision"]
    assert S["mass"]["saving"] == pytest.approx(1 - 650.065 / 1000.0) and S["mass"]["solid_g"] > S["mass"]["final_g"]
    assert "mass saving" in (rd / "summary.txt").read_text()


# ---- GPU ----------------------------------------------------------------------------------------------------------
def _lw_cfg(extra=None):
    return config.load(B3, {"grid": {"phases": [{"h_e": 0.8, "h_g": 0.2}]}, "objective": {"type": "lightweight"},
                            **(extra or {})})


@pytest.mark.gpu
def test_f0_gradient_fd():
    """∂(V/V0)/∂ξ (the lightweight f0 = −J, J = 1 − V/V0) against central differences along a smooth direction."""
    from cshell.optimise import Context, PhaseModel
    ctx = Context(_lw_cfg())
    m = PhaseModel(ctx, 0, 4.0)
    try:
        nv = ctx.surf.n_vertices
        dm, filt = m.dm, m.ds.filt
        rng = np.random.default_rng(1)
        xi = np.concatenate([np.full(nv, 0.4), np.full(nv, 0.3)])
        d = np.concatenate([filt.apply(rng.standard_normal(nv)), filt.apply(rng.standard_normal(nv))])
        d /= np.abs(d).max()

        def f0(x):
            return dm.volume(dm.forward(x[:nv], x[nv:], as_built=False)) / m.V0
        gd, ga = dm.volume_grad(dm.forward(xi[:nv], xi[nv:], as_built=False))
        an = float(np.concatenate([gd, ga]) @ d) / m.V0
        errs = []
        for h in (1e-2, 3e-3, 1e-3, 3e-4):
            fd = (f0(xi + h * d) - f0(xi - h * d)) / (2 * h)
            errs.append(abs(fd - an) / abs(an))
        print("f0 gradient rel. errors", errs, "analytic", an)
        assert abs(an) > 0 and min(errs) < 1e-3
    finally:
        m.release()


@pytest.fixture(scope="module")
def lw_run(tmp_path_factory):
    from cshell.optimise import Context, Run
    cfg = _lw_cfg({"growth": {"quantiles": [0.5, 1.0], "max_iters_per_step": 5, "min_iters_per_step": 2}})
    ctx = Context(cfg)
    out = tmp_path_factory.mktemp("lw")
    res = Run(cfg, "solid", str(out)).run(ctx)
    with open(out / "history.csv") as f:
        hist = list(csv.DictReader(f))
    return ctx, cfg, out, res, hist


@pytest.mark.gpu
def test_lightweight_run_no_restoration(lw_run):
    ctx, cfg, out, res, hist = lw_run
    vv = np.array([float(r["V_V0"]) for r in hist])
    print("V/V0", vv.round(4).tolist(), "g_max", [round(float(r["g_max"]), 4) for r in hist])
    assert all(int(float(r["n_restore"])) == 0 for r in hist)            # no mass restoration
    assert abs(vv[0] - 1.0) < 1e-3 and np.abs(vv - 1.0).max() > 1e-3     # starts at the solid; V is not held
    assert np.allclose([float(r["J"]) for r in hist], 1 - vv)             # J = mass saving
    assert np.isfinite(float(hist[0]["g_max"])) and float(hist[0]["g_max"]) > 0   # the solid is infeasible (1 < 1.02)
    assert res["phase_0"]["growth"]["best_boundary"] >= 0


@pytest.mark.gpu
def test_lightweight_evaluate_and_export_target(lw_run, tmp_path):
    from cshell.design import vertex_fields_final
    from cshell.optimise import evaluate, load_design
    ctx, cfg, out, res, hist = lw_run
    dsg = load_design(str(out))
    r = evaluate(ctx, dsg["xi"], 0, hollow=dsg["hollow"], D_comp=dsg["D_comp"])
    assert r["objective"] == "lightweight" and set(r["floors"]) == {"K", "P"}
    assert r["J"] == pytest.approx(1 - r["V_V0"]) and r["V_V0"] == pytest.approx(1 + r["mass_err"])
    assert r["feasible"] == (r["K_K0"] >= 1.0 and r["P_P0"] >= 1.0)
    h_g = float(cfg["grid"]["phases"][-1]["h_g"])
    f = vertex_fields_final(ctx, dsg["xi"], h_g, dsg["D_comp"], dsg["hollow"])
    ex = cex.export_design_gpu(ctx, dsg["xi"], str(tmp_path), dsg["D_comp"], name="t", h_g=h_g, fields=f,
                               closed_thr=cfg["export"]["closed_thr"], h_x=0.2, holes=False,
                               target_fraction=r["V_V0"])
    print("V_solid", ex["V_solid_mm3"], "target", ex["V0_mm3"], "mesh", ex["mesh"]["volume"])
    assert ex["V0_mm3"] == pytest.approx(r["V_V0"] * ex["V_solid_mm3"], rel=1e-12)
    assert abs(ex["mesh"]["volume"] / ex["V0_mm3"] - 1) < 2e-3                 # δ hits V_opt, not the solid's V0
