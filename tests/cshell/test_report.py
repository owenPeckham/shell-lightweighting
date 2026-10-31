"""cshell.report.build (THEORY 9): synthetic full run dir, partial run dir, real partial run, empty dir. CPU only."""
import json
import os
import shutil

import numpy as np
import pytest

from cshell import report

REPO = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _npages(pdf):
    import re
    return len(re.findall(rb"/Type\s*/Page\b", open(pdf, "rb").read()))


def _check(pdf):
    assert os.path.isfile(pdf) and os.path.getsize(pdf) < 5e6
    assert _npages(pdf) == 15


def _synthetic(d):
    import pyvista as pv
    import trimesh
    os.makedirs(d, exist_ok=True)
    box = trimesh.creation.box(extents=(40, 10, 10))
    box.export(os.path.join(d, "original.stl"))
    trimesh.creation.box(extents=(40, 10, 8)).export(os.path.join(d, "final_design.stl"))
    m = pv.PolyData(np.asarray(box.vertices), np.hstack([np.full((len(box.faces), 1), 3), box.faces]).ravel())
    rng = np.random.default_rng(0)
    for k in ["xi", "d", "a", "w", "r_MA", "d_hi", "passive", "W_bar", "stress"]:
        m.point_data[k] = rng.random(m.n_points)
    m.save(os.path.join(d, "fields.vtp"))
    json.dump({"name": "Syn", "step_file": os.path.join(d, "x.step"), "loads": [{"force_N": [0, 0, -1]}],
               "supports": [{"type": "cylinder"}], "optimiser": {"starts": ["uniform"]}},
              open(os.path.join(d, "config_used.json"), "w"))
    json.dump({"allocations": [{"host": "h", "slurm_job": "1"}]}, open(os.path.join(d, "run_info.json"), "w"))
    json.dump({"name": "Syn", "accepted": True, "reason": "J* > ln(1+g_min)", "mass": {"error": 1e-5},
               "comparison": [{"label": "solid", "K_K0": 1.0, "P_P0": 1.0, "J": 0.0, "governing": "-"},
                              {"label": "final", "K_K0": 0.9, "P_P0": 1.2, "J": 0.1, "governing": "stiffness"}],
               "resolutions": [{"label": "opt", "J": 0.1, "K_K0": 0.9, "P_P0": 1.2},
                               {"label": "shift", "J": 0.09, "K_K0": 0.91, "P_P0": 1.19}],
               "load_cases": [{"case": "L1", "lam_s": 2.0, "lam_b": 3.0, "softmin": 1.9}],
               "resources": {"gpu_model": "GH200", "gpu_hours": 1.0}}, open(os.path.join(d, "summary.json"), "w"))
    with open(os.path.join(d, "history.csv"), "w") as f:
        f.write("start,phase,iter,J,C_C0,lam_s_ratio,lam_b_ratio,mass_err,G_s,max_dxi,t_geometry,t_fe,t_mma,t_restore,gpu_mem_GiB\n")
        for i in range(6):
            f.write(f"uniform,{i // 3},{i},{-0.1 * i},1,1,1,1e-5,{0.1 ** i},0.01,1,2,3,4,10\n")
    with open(os.path.join(d, "sensitivity.csv"), "w") as f:
        f.write("kappa_b,sigma_f,J_star\n")
        for kb in (0.6, 0.7, 0.85, 1.0):
            for sf in (-20, -10, 0, 10, 20):
                f.write(f"{kb},{sf},{kb * (1 + sf / 100)}\n")
    open(os.path.join(d, "validation.csv"), "w").write("quantity,voxel,body_fitted\nK,1.0,1.01\n")
    open(os.path.join(d, "escape_holes.csv"), "w").write("void,x,y,z\n0,1,2,3\n")
    open(os.path.join(d, "wall_thickness.csv"), "w").write("w_nominal,w_asbuilt\n1,1.1\n2,1.9\n")


def test_synthetic_full(tmp_path):
    d = str(tmp_path / "Syn")
    _synthetic(d)
    pdf = report.build(d, verbose=False)
    _check(pdf)
    log = os.path.join(d, "plots", "report_log.txt")
    assert not os.path.exists(log) or "failed" not in open(log).read(), open(log).read()


def test_empty_dir(tmp_path):
    _check(report.build(str(tmp_path), verbose=False))


def test_partial_real_run(tmp_path):
    src = os.path.join(REPO, "runs_cshell", "timing", "Cantilever")
    if not os.path.isdir(src):
        pytest.skip("runs_cshell/timing/Cantilever not present")
    d = str(tmp_path / "Cantilever")
    os.makedirs(d)
    for f in os.listdir(src):
        p = os.path.join(src, f)
        if os.path.isfile(p) and f != "report.pdf":
            shutil.copy(p, d)
    _check(report.build(d, verbose=False))
