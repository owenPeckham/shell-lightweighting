"""PLAN_single_run §3 / contract C3: finalise with a mocked C2 evaluate, a mocked GPU export and a mocked report (CPU).

The run directory is synthetic (config_used.json from examples/cshell/mbb.json, history.csv with the C2 columns, xi
files, a checkpoint); the context is a stub with the attributes finalise reads. What is checked: the C3 summary keys,
the shifted-grid call (shift = ½ h_e of the final phase in x, y, z), the decision rule, the guarded report, the legacy
start table (C_C0, lam_ratio columns) and that --out leaves the run directory untouched."""
import csv
import json
import math
import os
from pathlib import Path
from types import SimpleNamespace

import numpy as np
import pytest

from cshell import collect as col
from cshell import config
from cshell import export as cex
from cshell import finalise as fin

ROOT = Path(__file__).resolve().parents[2]
NV = 50
C3_KEYS = ("name", "decision", "mass", "opt", "shift", "export", "closed_area_frac", "hollow_area_frac", "timing",
           "peak")


def _run_dir(tmp_path, legacy=False):
    cfg = config.load(str(ROOT / "examples/cshell/mbb.json"))
    cfg["output_dir"] = str(tmp_path)
    start = "neutral" if legacy else "solid"
    cfg["optimiser"]["starts"] = [start]
    rd = tmp_path / "MBB"
    (rd / "checkpoints").mkdir(parents=True)
    json.dump(cfg, open(rd / "config_used.json", "w"), indent=1)
    if legacy:
        cols = ["start", "phase", "iter", "J", "C_C0", "lam_ratio", "mass_err", "gpu_mem_GiB", "t_geometry", "t_fe",
                "t_mma", "t_restore"]
        rows = [[start, p, i, 0.4 + 0.01 * i, 0.9, 2.4, 1e-6, 12.0, 0.1, 5.0, 0.3, 0.1] for p in (0, 1) for i in range(3)]
    else:
        cols = ["start", "phase", "growth_step", "it", "R_k", "J", "K_K0", "lamB_ratio", "sig_ratio", "mass_err",
                "closed_frac", "hollow_allowed_frac", "gpu_peak_GiB", "t_geometry", "t_fe", "t_mma", "t_restore"]
        rows = [[start, p, 0, i, 0.1, 0.1 + 0.05 * i, 1.1, 1.02, 0.97, 2e-6, 0.3, 0.5, 9.5, 0.2, 4.0, 0.5, 0.0]
                for p in (0, 1) for i in range(3)]
    with open(rd / "history.csv", "w", newline="") as f:
        w = csv.writer(f)
        w.writerow(cols)
        w.writerows(rows)
    for p in (0, 1):
        np.save(rd / f"xi_{start}_phase{p}.npy", np.zeros(2 * NV))
    np.savez(rd / "checkpoints" / f"{start}_1_0003.npz", xi=np.zeros(2 * NV), D_comp=np.array(6.0),
             hours_used=np.array(0.5))
    json.dump({"start": start, "status": "done"}, open(rd / f"done_{start}.json", "w"))
    return cfg, rd


def _ctx(cfg):
    return SimpleNamespace(cfg=cfg, surf=SimpleNamespace(n_vertices=NV, A=np.ones(NV)), r_ma=np.full(NV, 2.0))


@pytest.fixture
def mocks(monkeypatch):
    calls = dict(export=[], evaluate=[], report=[])

    def vertex_fields(ctx, xi, D_comp=4.0, h_g=None):
        xt = np.linspace(0.0, 1.0, NV)
        return dict(d=np.full(NV, 0.5), a=np.where(xt > 0.95, 2.5, 1.0), d_hi=np.full(NV, 3.0), xt_a=xt,
                    passive=np.arange(NV) < 5, r_ma=ctx.r_ma)

    def export_design_gpu(ctx, xi, out_dir, D_comp=4.0, stress=None, name="final_design", h_g=None, fields=None,
                          closed_thr=None, **kw):
        calls["export"].append(dict(D_comp=D_comp, h_x=kw.get("h_x"), closed_thr=closed_thr, out_dir=out_dir))
        Path(out_dir, f"{name}.stl").write_bytes(b"\0" * 84)
        return dict(ok=True, V0_mm3=1000.0, mesh=dict(faces=1234, watertight_manifold=True, bodies=1, volume=1000.2,
                                                      volume_rel=2e-4),
                    mass_correction=dict(ok=True, fail=False, delta=-0.05, mesh_steps=[]), voids=dict(enclosed=2, holes=4),
                    wall=dict(min_nominal=1.2, p05_nominal=1.5, median_nominal=2.5, max_nominal=6.0, min_as_built=1.0),
                    cleanup=dict(opening_removed_mm3=3.0), timing_s={})

    def write_report(run_dir):
        calls["report"].append(run_dir)
        assert os.path.exists(os.path.join(run_dir, "summary.json"))      # the report reads the summary
        return os.path.join(run_dir, "report.pdf")

    monkeypatch.setattr(cex, "vertex_fields", vertex_fields)
    monkeypatch.setattr(cex, "export_design_gpu", export_design_gpu)
    import cshell.report as rep
    monkeypatch.setattr(rep, "write_report", write_report, raising=False)
    return calls


def _evaluate(calls, J=0.3, feasible=True):
    def ev(ctx, xi, phase, shift=None, D_comp=None):
        calls["evaluate"].append(dict(phase=phase, shift=shift, D_comp=D_comp))
        return dict(J=J, K_K0=1.09, lamB_ratio=1.01, sig_ratio=0.98, mass_err=1e-6, feasible=feasible,
                    governing="buckling")
    return ev


def test_finalise_accepts_and_writes_c3_summary(tmp_path, mocks):
    cfg, rd = _run_dir(tmp_path)
    S = fin.finalise(str(rd), ctx=_ctx(cfg), evaluate=_evaluate(mocks))
    for k in C3_KEYS:
        assert k in S, k
    assert S["decision"]["accepted"] is True, S["decision"]
    h_e = cfg["grid"]["phases"][-1]["h_e"]
    ev = mocks["evaluate"][0]
    assert ev["phase"] == len(cfg["grid"]["phases"]) - 1 and ev["D_comp"] == 6.0
    assert np.allclose(ev["shift"], [0.5 * h_e] * 3)
    assert mocks["export"][0]["h_x"] == 0.2 and mocks["export"][0]["D_comp"] == 6.0
    assert S["opt"]["J"] == pytest.approx(0.2) and S["opt"]["K_K0"] == pytest.approx(1.1)
    assert S["opt"]["lamB_ratio"] == pytest.approx(1.02) and S["opt"]["sig_ratio"] == pytest.approx(0.97)
    for k in ("J", "K_K0", "lamB_ratio", "sig_ratio", "feasible"):
        assert k in S["shift"], k
    for k in ("h_mm", "faces", "watertight", "volume_mm3", "holes", "voids", "wall"):
        assert k in S["export"], k
    assert set(S["export"]["wall"]) >= {"min", "p05", "median", "max"}
    assert set(S["mass"]) >= {"target_g", "final_g"}
    rho = cfg["material"]["density_g_cm3"] * 1e-3
    assert S["mass"]["target_g"] == pytest.approx(1000.0 * rho) and S["mass"]["final_g"] == pytest.approx(1000.2 * rho)
    assert set(S["timing"]) >= {"phase_s", "finalise_s", "export_s"} and len(S["timing"]["phase_s"]) == 2
    assert S["timing"]["phase_s"][0] == pytest.approx(3 * 4.7)
    assert set(S["peak"]) == {"gpu_gib", "host_gib"} and S["peak"]["gpu_gib"] >= 9.5
    assert 0 <= S["closed_area_frac"] <= 1 and 0 <= S["hollow_area_frac"] <= 1
    assert S["report"]["ok"] is True and mocks["report"] == [str(rd)]
    on_disk = json.load(open(rd / "summary.json"))
    assert on_disk["decision"] == S["decision"] and on_disk["report"]["ok"] is True
    assert (rd / "summary.txt").read_text().startswith("cshell summary: MBB")


@pytest.mark.parametrize("case", ["infeasible", "low_J", "evaluate_raises", "no_evaluate"])
def test_finalise_rejects(tmp_path, mocks, monkeypatch, case):
    cfg, rd = _run_dir(tmp_path)
    if case == "infeasible":
        ev = _evaluate(mocks, feasible=False)
    elif case == "low_J":
        ev = _evaluate(mocks, J=0.01)
    elif case == "evaluate_raises":
        def ev(*a, **k):
            raise RuntimeError("PCG not converged")
    else:
        import cshell.optimise as opt
        monkeypatch.delattr(opt, "evaluate", raising=False)
        ev = None
    S = fin.finalise(str(rd), ctx=_ctx(cfg), evaluate=ev)
    assert S["decision"]["accepted"] is False
    assert S["decision"]["reason"].startswith("original SOLID part kept")
    if case in ("evaluate_raises", "no_evaluate"):
        assert S["shift"]["error"] and S["shift"]["feasible"] is False
    assert (rd / "final_design.stl").exists()                           # export still written (D2: kept to inspect)


def test_report_failure_does_not_fail_the_run(tmp_path, mocks, monkeypatch):
    import cshell.report as rep

    def boom(run_dir):
        raise ValueError("matplotlib exploded")
    monkeypatch.setattr(rep, "write_report", boom, raising=False)
    cfg, rd = _run_dir(tmp_path)
    S = fin.finalise(str(rd), ctx=_ctx(cfg), evaluate=_evaluate(mocks))
    assert S["report"]["ok"] is False and "matplotlib exploded" in S["report"]["error"]
    assert json.load(open(rd / "summary.json"))["report"]["ok"] is False


def test_out_dir_leaves_run_dir_untouched_legacy_run(tmp_path, mocks):
    cfg, rd = _run_dir(tmp_path, legacy=True)
    before = {p: p.stat().st_mtime_ns for p in rd.rglob("*")}
    out = tmp_path / "reexport"
    S = fin.finalise(str(rd), ctx=_ctx(cfg), out_dir=str(out), evaluate=_evaluate(mocks))
    assert {p: p.stat().st_mtime_ns for p in rd.rglob("*")} == before
    assert (out / "config_used.json").exists() and (out / "history.csv").exists() and (out / "summary.json").exists()
    assert S["opt"]["start"] == "neutral" and S["opt"]["K_K0"] == pytest.approx(1 / 0.9)
    assert S["opt"]["lamB_ratio"] is None                                  # not in a Stage 2 history
    assert mocks["export"][0]["out_dir"] == str(out)
    assert mocks["export"][0]["closed_thr"] == 0.98                      # Stage 2 map: legacy closure threshold


def test_decide_truth_table():
    crit = math.log(1.02)
    good = dict(J=0.1, feasible=True, lamB_ratio=1.0, sig_ratio=1.0)
    assert fin.decide(dict(J=0.1), good, True, "", crit)[0] is True
    assert fin.decide(dict(J=0.01), good, True, "", crit)[0] is False
    assert fin.decide(dict(J=0.1), dict(good, J=0.01), True, "", crit)[0] is False
    assert fin.decide(dict(J=0.1), dict(good, feasible=False), True, "", crit)[0] is False
    assert fin.decide(dict(J=0.1), good, False, "not watertight", crit)[0] is False
    assert fin.decide(dict(J=0.1), None, True, "", crit)[0] is False


def test_start_table_reads_c2_and_legacy_columns(tmp_path):
    cfg, rd = _run_dir(tmp_path)
    r = col.start_table(str(rd), cfg, merge=False)[0]
    assert r["status"] == "done" and r["feasible"] and r["K_K0_opt"] == pytest.approx(1.1)
    assert r["lamB_ratio_opt"] == pytest.approx(1.02) and r["xi_file"] == "xi_solid_phase1.npy"
