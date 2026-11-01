"""cshell.report (PLAN_single_run section 5, contract C4): BC arrows on the analysed faces, streamed STL decimation and
sections, synthetic C2/C3 run dir, Stage 2 adapter, empty dir.  CPU only."""
import json
import os
import re

import numpy as np
import pytest

from cshell import report

REPO = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
NPAGES = len(report.PAGES)


def _npages(pdf):
    return len(re.findall(rb"/Type\s*/Page\b", open(pdf, "rb").read()))


def _check(pdf):
    assert os.path.isfile(pdf) and os.path.getsize(pdf) < 8e6
    assert _npages(pdf) == NPAGES == 11


def _box_bc(layout="stage2"):
    import trimesh
    m = trimesh.creation.box(extents=(40, 10, 10)).subdivide().subdivide().subdivide()
    V, T = np.asarray(m.vertices, float), np.asarray(m.faces)
    c = V[T].mean(1)
    top = np.all(np.abs(V[T][:, :, 2] - 5) < 1e-9, axis=1)
    left = top & (c[:, 0] > -16) & (c[:, 0] < -4)
    right = top & (c[:, 0] > 4) & (c[:, 0] < 16)
    sup = np.all(V[T][:, :, 0] < -19.99, axis=1)
    load = np.full(len(T), -1)
    load[left | right] = 0
    return V, T, sup, load, left, right


def test_load_arrows_on_loaded_faces(tmp_path):
    V, T, sup, load, left, right = _box_bc()
    F = np.array([[0, 0, -200.0]])
    bc = report.load_bc({"V": V, "T": T, "support": sup, "load": load, "force_N": F})
    arr = report.load_arrows(bc)
    assert len(arr) == 2                                               # one arrow per connected loaded patch
    assert np.allclose(sum(a["vec"] for a in arr), F[0])               # shares sum to F, along F
    A = report.tri_areas(V, T)
    for a in arr:
        assert abs(np.linalg.norm(a["vec"]) - 200 * A[a["tris"]].sum() / A[load == 0].sum()) < 1e-9
        assert a["vec"][0] == 0 and a["vec"][1] == 0 and a["vec"][2] < 0
        # independent check: dense barycentric samples of the loaded triangles
        rng = np.random.default_rng(0)
        tri = V[T[load == 0]]
        u, v = rng.random((2, 200000))
        flip = u + v > 1
        u, v = np.where(flip, 1 - u, u), np.where(flip, 1 - v, v)
        k = rng.integers(0, len(tri), 200000)
        S = tri[k, 0] + u[:, None] * (tri[k, 1] - tri[k, 0]) + v[:, None] * (tri[k, 2] - tri[k, 0])
        assert np.linalg.norm(S - a["base"], axis=1).min() < 1.0
        assert a["snap_mm"] < 1e-6                                     # planar patch: centroid already on the face


def test_arrow_base_snaps_for_nonconvex_patch():
    V = np.array([[0, 0, 0], [10, 0, 0], [0, 10, 0], [0, 0, 1], [1, 0, 1], [0, 1, 1]], float)   # two separate triangles
    tri = V[[0, 1, 2]][None]
    p, d = report.closest_on_triangles(np.array([20.0, 20.0, 3.0]), tri)
    assert d == pytest.approx(np.linalg.norm(p - [20, 20, 3]))
    assert np.allclose(p, [5, 5, 0])                                    # hypotenuse edge, z projected to 0


@pytest.mark.parametrize("layout", ["stage2", "c2_labels", "c2_indices"])
def test_bc_layouts(layout):
    V, T, sup, load, *_ = _box_bc()
    z = {"V": V, "T": T, "force_N": np.array([[0, 0, -1.0]])}
    if layout == "stage2":
        z.update(support=sup, load=load)
    elif layout == "c2_labels":
        z.update(support_faces=sup, load_faces=load)
    else:
        z.update(support_faces=np.where(sup)[0], load_faces=np.where(load == 0)[0])
    bc = report.load_bc(z)
    assert bc["support"].sum() == sup.sum() and (bc["load"] == 0).sum() == (load == 0).sum()
    assert len(report.load_arrows(bc)) == 2


def _write_stl(path, m):
    tri = np.asarray(m.vertices, "<f4")[m.faces]
    rec = np.zeros(len(tri), np.dtype([("n", "<f4", 3), ("v", "<f4", (3, 3)), ("a", "<u2")]))
    rec["v"] = tri
    with open(path, "wb") as f:
        f.write(b"\0" * 80)
        f.write(np.uint32(len(tri)).tobytes())
        f.write(rec.tobytes())


def test_stream_stl_decimation_and_section(tmp_path):
    import trimesh
    m = trimesh.creation.icosphere(subdivisions=5, radius=10.0)
    p = str(tmp_path / "s.stl")
    _write_stl(p, m)
    V, T, N, secs = report.stream_stl(p, target=2000, planes=[(2, 0.0)], chunk=7000)
    assert N == len(m.faces) and 800 < len(T) < 6000
    assert np.abs(np.linalg.norm(V, axis=1).mean() - 10) < 0.3          # decimated surface stays on the sphere
    assert np.abs(V.max(0) - 10).max() < 1.0
    seg = secs[0]
    length = np.linalg.norm(seg[:, 1] - seg[:, 0], axis=1).sum()
    assert length == pytest.approx(2 * np.pi * 10, rel=0.01)             # section of a sphere at the equator


def _synthetic(d):
    import pyvista as pv
    import trimesh
    os.makedirs(d, exist_ok=True)
    V, T, sup, load, *_ = _box_bc()
    np.savez(os.path.join(d, "bc_faces.npz"), V=V, T=T, support_faces=sup, load_faces=load, force_N=np.array([[0, 0, -200.0]]))
    m = trimesh.creation.icosphere(subdivisions=4, radius=5.0)
    m.apply_scale([4, 1, 1])
    _write_stl(os.path.join(d, "final_design.stl"), m)
    P = np.asarray(m.vertices)
    pm = pv.PolyData(P, np.hstack([np.full((len(m.faces), 1), 3), m.faces]).ravel())
    rng = np.random.default_rng(0)
    w = 1.5 + 3 * rng.random(len(P))
    passive = (P[:, 0] > 15).astype(float)
    w[passive > 0] = np.nan
    w[(P[:, 0] < -15)] = np.nan
    pm.point_data["w"], pm.point_data["passive"] = w, passive
    pm.point_data["stress"], pm.point_data["stress_solid"] = 5 + rng.random(len(P)), 6 + rng.random(len(P))
    pm.save(os.path.join(d, "fields.vtp"))
    json.dump({"name": "Syn", "loads": [{"force_N": [0, 0, -200]}], "manufacturing": {"t_min_mm": 1.5},
               "optimiser": {"starts": ["solid"]}}, open(os.path.join(d, "config_used.json"), "w"))
    json.dump({"name": "Syn", "decision": {"accepted": True, "reason": "J > ln 1.02 on both grids"},
               "mass": {"target_g": 10.0, "final_g": 10.001},
               "opt": {"J": 0.12, "K_K0": 1.13, "lamB_ratio": 1.05, "sig_ratio": 0.9},
               "shift": {"J": 0.11, "K_K0": 1.12, "lamB_ratio": 1.04, "sig_ratio": 0.92, "feasible": True},
               "export": {"h_mm": 0.2, "faces": len(m.faces), "watertight": True, "volume_mm3": 100.0, "holes": 2,
                          "voids": {"total": 1, "without_hole": 0, "per_void": [{"id": 0, "volume_mm3": 50.0, "holes": 2}]},
                          "wall": {"min": 1.5, "p05": 1.6, "median": 2.5, "max": 4.5}},
               "closed_area_frac": 0.1, "hollow_area_frac": 0.7,
               "timing": {"phase_s": [10, 20], "finalise_s": 5, "export_s": 3}, "peak": {"gpu_gib": 3.0, "host_gib": 4.0}},
              open(os.path.join(d, "summary.json"), "w"))
    with open(os.path.join(d, "history.csv"), "w") as f:
        f.write("it,phase,growth_step,R_k,J,K_K0,lamB_ratio,sig_ratio,mass_err,closed_frac,hollow_allowed_frac\n")
        for i in range(12):
            f.write(f"{i},{i // 6},{i // 3},{0.1 * (i // 3)},{0.01 * i},{1 + 0.01 * i},1.1,0.9,1e-5,0.1,{0.1 * (i // 3)}\n")
    with open(os.path.join(d, "wall_thickness.csv"), "w") as f:
        f.write("bin_lo_mm,bin_hi_mm,count_nominal,count_as_built\n")
        for i in range(10):
            f.write(f"{0.5 * i},{0.5 * (i + 1)},{10 * i},{9 * i}\n")
    open(os.path.join(d, "escape_holes.csv"), "w").write("void_id,hole,x,y,z,wall_mm,score_kind\n0,0,1,2,3,2.5,stress\n0,1,-1,2,3,2.5,stress\n")


def test_synthetic_full(tmp_path):
    d = str(tmp_path / "Syn")
    _synthetic(d)
    pdf = report.write_report(d, verbose=False)
    _check(pdf)
    log = os.path.join(d, "plots", "report_log.txt")
    assert not os.path.exists(log), open(log).read()                    # no page failed and none fell back to 'not available'


def test_empty_dir(tmp_path):
    _check(report.write_report(str(tmp_path), verbose=False))


@pytest.mark.parametrize("part", ["MBB", "Cantilever", "L-bracket"])
def test_stage2_adapter(tmp_path, part):
    src = os.path.join(REPO, "examples", "cshell", "runs", part)
    if not os.path.isfile(os.path.join(src, "summary.json")):
        pytest.skip("Stage 2 example run not present")
    out = str(tmp_path / part)
    os.makedirs(out)
    pdf = report.write_report(src, out_dir=out, verbose=False)         # never writes into examples/
    _check(pdf)
    R = report.Run(src, out)
    assert R.S["opt"]["lamB_ratio"] and R.S["decision"]["accepted"] is not None
    bc = R.bc()
    arr = report.load_arrows(bc)
    assert len(arr) == {"MBB": 2, "Cantilever": 1, "L-bracket": 1}[part]
    assert len(report.patches(bc["T"], bc["support"])) == {"MBB": 4, "Cantilever": 4, "L-bracket": 3}[part]
    for a in arr:
        assert a["snap_mm"] < 1.0
