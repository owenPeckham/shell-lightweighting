"""Start-name alias after the vshell -> dshell rename (2026-10-31): "vshell" is accepted on input, and a run directory
written before the rename (history_vshell.csv, xi_vshell_phase*.npy, checkpoints/vshell_*.npz, done_vshell.json,
start column "vshell") is still read by resume, collect's start table, history merge and the report (CPU, no FE)."""
import copy
import csv
import json
import os

import numpy as np

from cshell import cli, config
from cshell.collect import latest_checkpoint, read_history, start_table
from cshell.optimise import Run

COLS = ["start", "phase", "iter", "J", "C_C0", "lam_ratio", "mass_err"]


def _cfg(starts):
    cfg = copy.deepcopy(config.DEFAULTS)
    cfg["optimiser"]["starts"] = list(starts)
    cfg["optimiser"]["vshell_run"] = "examples/dshell/runs/X"
    return config.normalise_starts(cfg)


def _legacy_run(d, n_ph):
    """Fake run dir with the pre-rename file names: start uniform (new names) and start vshell (old names)."""
    os.makedirs(os.path.join(d, "checkpoints"))
    for start, J in (("uniform", 0.40), ("vshell", 0.50)):
        with open(os.path.join(d, f"history_{start}.csv"), "w", newline="") as f:
            w = csv.writer(f)
            w.writerow(COLS)
            for ph in range(n_ph):
                for it in range(3):
                    w.writerow([start, ph, it, J + 0.01 * it, 1.0, 1.0, 0.0])
        for ph in range(n_ph):
            np.save(os.path.join(d, f"xi_{start}_phase{ph}.npy"), np.full(8, 0.5))
        np.savez(os.path.join(d, "checkpoints", f"{start}_{n_ph - 1}_0002.npz"), xi=np.full(8, 0.5), phase=n_ph - 1,
                 it=2, c_norm=1.0, s_J=np.nan, D_comp=3.0, hours_used=1.5 if start == "vshell" else 1.0)
        json.dump({"start": start, "status": "done"}, open(os.path.join(d, f"done_{start}.json"), "w"))
    cfg = _cfg(["uniform", "vshell"])
    json.dump(dict(cfg, optimiser=dict(cfg["optimiser"], starts=["uniform", "vshell"])),
              open(os.path.join(d, "config_used.json"), "w"))
    return cfg


def test_alias_names():
    assert config.canonical_start("vshell") == "dshell"
    assert config.canonical_start("uniform") == "uniform"
    assert config.start_names("vshell") == ("dshell", "vshell")
    assert config.start_names("neutral") == ("neutral",)
    cfg = _cfg(["uniform", "vshell", "dshell", "neutral"])
    assert cfg["optimiser"]["starts"] == ["uniform", "dshell", "neutral"]
    assert cfg["optimiser"]["dshell_run"] == "examples/dshell/runs/X"
    assert config.DEFAULTS["optimiser"]["starts"] == ["uniform", "dshell", "neutral"]


def test_start_file_prefers_new_name(tmp_path):
    d = str(tmp_path)
    assert config.start_file(d, "done_{}.json", "vshell").endswith("done_dshell.json")   # none exists: new name
    open(os.path.join(d, "done_vshell.json"), "w").close()
    assert config.start_file(d, "done_{}.json", "dshell").endswith("done_vshell.json")   # only legacy exists
    open(os.path.join(d, "done_dshell.json"), "w").close()
    assert config.start_file(d, "done_{}.json", "dshell").endswith("done_dshell.json")   # both: new name


def test_legacy_run_dir_collect_side(tmp_path):
    d = str(tmp_path)
    n_ph = len(config.DEFAULTS["grid"]["phases"])
    cfg = _legacy_run(d, n_ph)
    hist = read_history(d)                                       # merges history_*.csv, canonical start names
    assert {r["start"] for r in hist} == {"uniform", "dshell"}
    with open(os.path.join(d, "history.csv")) as f:
        assert "vshell" not in {r["start"] for r in csv.DictReader(f)}
    for s in ("dshell", "vshell"):
        z = latest_checkpoint(d, s)
        assert z is not None and float(z["hours_used"]) == 1.5
    rows = {r["start"]: r for r in start_table(d, cfg)}
    assert set(rows) == {"uniform", "dshell"}
    r = rows["dshell"]
    assert r["status"] == "done" and r["phases_done"] == n_ph and r["iterations"] == 3 * n_ph
    assert r["xi_file"] == f"xi_vshell_phase{n_ph - 1}.npy" and os.path.exists(os.path.join(d, r["xi_file"]))
    assert r["hours"] == 1.5 and abs(r["J_opt"] - 0.52) < 1e-12


def test_legacy_run_dir_resume_side(tmp_path):
    d = str(tmp_path)
    n_ph = len(config.DEFAULTS["grid"]["phases"])
    cfg = _legacy_run(d, n_ph)
    run = Run(cfg, "vshell", d)
    assert run.start == "dshell"
    rs = run.load_latest()                                       # legacy checkpoints/vshell_*.npz
    assert rs is not None and rs.phase == n_ph - 1 and rs.it == 2 and rs.hours_used == 1.5
    run._logf.close()
    # --resume with every start done (done_vshell.json counts for dshell): nothing to run, no FE context built
    assert cli.main(["run", "--resume", d]) == 0
    assert cli.main(["run", "--resume", d, "--start", "vshell"]) == 0
    assert cli.main(["run", "--resume", d, "--start", "dshell"]) == 0


def test_report_reader_canonicalises(tmp_path):
    from cshell.report import Run as Report
    d = str(tmp_path)
    _legacy_run(d, len(config.DEFAULTS["grid"]["phases"]))
    read_history(d)
    R = Report(d)
    assert {r["start"] for r in R.hist} == {"uniform", "dshell"}
