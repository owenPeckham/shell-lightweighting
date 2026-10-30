"""THEORY §8.1-8.5: tiled export. Hollow sphere with analytic wall, opening rule, escape holes, mass correction, mesh
checks, uniform shell and the B2 rod end to end (CPU)."""
import math
import os
from pathlib import Path

import numpy as np
import pytest
import trimesh

from cshell import export as ex

ROOT = Path(__file__).resolve().parents[2]
R0 = 8.0


@pytest.fixture(scope="module")
def sphere():
    m = trimesh.creation.icosphere(subdivisions=4, radius=R0)
    V, T = np.asarray(m.vertices), np.asarray(m.faces)
    return V, T, V / np.linalg.norm(V, axis=1, keepdims=True), float(m.volume)


def _design(sphere, d, a):
    V, T, _, _ = sphere
    nv = len(V)
    return ex.Design(V, T, V, T, d=np.full(nv, d), a=np.full(nv, a), d_hi=np.full(nv, 3.0), active=np.ones(nv, bool),
                     closed=np.zeros(nv, bool))


def _workers():
    return min(8, len(os.sched_getaffinity(0)))


@pytest.fixture(scope="module")
def hollow(sphere, tmp_path_factory):
    """a = 2 (void radius 6), d = 0.9 → mass correction δ ≈ 0.1; two escape holes."""
    out = tmp_path_factory.mktemp("hollow")
    cfg = ex.ExportConfig(h_x=0.1, tile=64, block=16, workers=_workers(), wall_samples=20000, log=lambda *_: None)
    e = ex.Exporter(_design(sphere, 0.9, 2.0), cfg, normals=sphere[2])
    s = e.run(str(out))
    return s, out, e


def test_hollow_sphere_volume_wall_voids(hollow, sphere):
    s, out, e = hollow
    V0 = sphere[3]
    # V0 on the export grid = solid volume of the faceted sphere to the ramp accuracy
    assert abs(s["V0_mm3"] / V0 - 1) < 1e-3
    mc = s["mass_correction"]
    # analytic: (R + d)^3 - (R - a)^3 = R^3 plus the volume of the two holes: d = 1.0 without holes
    assert mc["ok"] and abs(mc["rel_after"]) < 2e-5
    assert 0.09 < mc["delta"] < 0.30                                      # d: 0.9 → ≈ 1.0 + holes
    assert s["mesh"]["watertight_manifold"] and s["mesh"]["trimesh"]["watertight"]
    assert abs(s["mesh"]["volume"] / s["V0_mm3"] - 1) <= 2e-3             # R1
    # one enclosed void of volume 4/3 π 6³ = 904.8 mm³ (cells, before holes), kept by the opening
    v = s["voids"]
    assert v["total"] == 1 and v["enclosed"] == 1
    assert abs(v["per_void"][0]["volume_mm3"] / (4 / 3 * math.pi * 6 ** 3) - 1) < 0.01
    # wall thickness a + d ≈ 3.0 (median); as-built = −b_v on the void side
    w = s["wall"]
    assert abs(w["median_nominal"] - 3.0) < 0.06
    assert w["min_as_built"] <= w["min_nominal"]


def test_escape_holes(hollow, sphere):
    s, out, e = hollow
    rows = np.genfromtxt(out / "escape_holes.csv", delimiter=",", skip_header=1, usecols=range(0, 10))
    assert rows.shape[0] == 2                                              # two holes in the single void
    assert s["voids"]["without_hole"] == 0 and s["voids"]["holes"] == 2
    p = rows[:, 2:5]
    assert np.linalg.norm(p[0] - p[1]) > 1.5 * 6.0                         # far apart along the principal axis
    # void connected to the exterior after cutting: one body only (an enclosed void would be a second shell)
    assert s["mesh"]["bodies"] == 1 and s["mesh"]["trimesh"]["bodies"] == 1
    # hole diameter: mesh crosses a disc of radius 1.5 around the hole axis; the axis is empty through the wall
    V, F = e.mesh
    tm = trimesh.Trimesh(V.astype(float), F, process=False)
    n = rows[0, 5:8]
    p0 = rows[0, 2:5]
    side = np.cross(n, [1.0, 0.3, 0.2]); side /= np.linalg.norm(side)
    q = np.array([p0 + n * 1.0, p0 + n * 2.0, p0 + n * 1.5 + 1.0 * side,      # inside the D = 3 mm hole
                  p0 + n * 1.5 + 2.5 * side])                                 # control: wall material beside the hole
    inside = tm.contains(q)
    assert not inside[0] and not inside[1] and not inside[2] and inside[3]
    # the genus of the exported surface is 1 (a sphere shell with two tunnels joining the void to the outside)
    assert s["mesh"]["euler"] == 0


def test_escape_hole_caps_clear_walls(hollow, sphere):
    """Each hole's inner cap lies wholly in the void and its outer cap wholly outside the part (no grazing cut of the
    curved void wall: B2 had films down to 0.0015 mm there); wall statistics have no sub-(t_min - b_v) sample off the rims."""
    s, out, e = hollow
    H = e.holes
    assert len(H["p"]) == 2
    r = H["r"][0]
    for p, n, s0, s1 in zip(H["p"], H["n"], H["s0"], H["s1"]):
        e1 = np.cross(n, [1.0, 0.3, 0.2]); e1 /= np.linalg.norm(e1)
        e2 = np.cross(n, e1)
        ang = np.linspace(0, 2 * np.pi, 24, endpoint=False)
        ring = r * (np.cos(ang)[:, None] * e1 + np.sin(ang)[:, None] * e2)
        assert np.all(np.linalg.norm(p + s0 * n + ring, axis=1) < 6.0)          # inner cap in the void (radius 6)
        assert np.all(np.linalg.norm(p + s1 * n + ring, axis=1) > 6.0 + 3.0)    # outer cap beyond the outer skin
    w = s["wall_off_rims"]
    assert w["n_below_t_min_minus_b_v"] == 0 and w["min_as_built"] > e.cfg.t_min - e.cfg.b_v


def test_b2_validation_export_has_no_films():
    """Re-exported B2 design (runs_cshell/VF_B2, written by the validation-fix collect): no wall below t_min - b_v away
    from the flagged hole rims, and every void is connected to the exterior (one body)."""
    import json
    f = ROOT / "runs_cshell/VF_B2/best_candidate_export.json"
    if not f.exists():
        pytest.skip("runs_cshell/VF_B2 not generated")
    s = json.load(open(f))
    assert s["voids"]["without_hole"] == 0 and s["mesh"]["bodies"] == 1
    w = s["wall_off_rims"]
    # walls sit exactly at the t_min bound by design (research/vf_b2/wall_where.py, job 7864770): count only samples
    # below t_min - b_v by more than the export voxel discretisation (2.5 h_x), PLAN_collect_fixes P7
    cfg = json.load(open(f.parent / "config_used.json"))
    thr = cfg["manufacturing"]["t_min_mm"] - cfg["manufacturing"]["wall_bias_mm"] - 2.5 * s["h_x"]
    if "n_below_t_min_minus_b_v_disc" in w:
        assert w["n_below_t_min_minus_b_v_disc"] == 0, w
    assert w["min_as_built"] >= thr, (w, thr)


def test_closure_rule_interpolates_the_optimiser_a():
    """§8.2 closure rule (P2): a cell is closed (A_CLOSED) only when all its positive-weight vertices are closed; a
    mixed triangle interpolates the optimiser's a linearly (it used to interpolate A_CLOSED = 1e4 and close the whole
    ring)."""
    V = np.array([[0, 0, 0], [1, 0, 0], [0, 1, 0], [1, 1, 0.0]])
    T = np.array([[0, 1, 2], [1, 3, 2]])
    a = np.array([2.0, 3.0, 5.0, 7.0])
    closed = np.array([False, True, True, True])
    dz = ex.Design(V, T, V, T, d=np.zeros(4), a=a, d_hi=np.ones(4), active=np.ones(4, bool), closed=closed)
    tri = np.array([0, 0, 1, 0], np.int32)
    b1 = np.array([0.2, 0.5, 0.3, 1.0], np.float32)          # vertex-1 weight
    b2 = np.array([0.3, 0.5, 0.3, 0.0], np.float32)          # vertex-2 weight
    out = dz.a_at(tri, b1, b2)
    w0 = 1.0 - b1.astype(float) - b2.astype(float)
    lin = w0 * a[0] + b1 * a[1] + b2 * a[2]
    assert abs(out[0] - lin[0]) < 1e-12                       # mixed: linear in the optimiser's a
    assert out[1] == ex.A_CLOSED                             # on edge 1-2: vertex 0 has zero weight -> closed
    assert out[2] == ex.A_CLOSED                             # triangle 1 is all closed
    assert out[3] == ex.A_CLOSED                             # at vertex 1
    assert np.all(out[:1] < 10.0)


def test_mass_correction_clip_and_bracket(sphere, tmp_path):
    """d + δ is clipped at d_hi: a too-small d_hi cannot restore V; the result is flagged, never silently wrong."""
    V, T, N, V0 = sphere
    nv = len(V)
    dz = ex.Design(V, T, V, T, d=np.full(nv, 0.9), a=np.full(nv, 2.0), d_hi=np.full(nv, 0.95), active=np.ones(nv, bool),
                   closed=np.zeros(nv, bool))
    cfg = ex.ExportConfig(h_x=0.25, tile=32, block=8, workers=_workers(), holes=False, wall_hist=False,
                          write_stl=False, log=lambda *_: None)
    s = ex.Exporter(dz, cfg, normals=N).run(str(tmp_path))
    mc = s["mass_correction"]
    assert not mc["ok"] and "bracket_failed" in mc and not s["ok"]


# -- opening rule --------------------------------------------------------------------------------------------------
def _ball(shape, c, r):
    g = np.indices(shape).astype(float)
    return sum((g[i] - c[i]) ** 2 for i in range(3)) <= r * r


def test_opening_sliver_and_rounded_ends():
    h = 0.1
    R2 = ex.ball_r2(1.5, h)                                               # r = 15 cells
    shape = (200, 200, 200)
    m = np.zeros(shape, bool)
    m[20:30, 20:120, 20:120] = True                                       # sliver 1 mm thick plate
    ball = _ball(shape, (150, 60, 60), 20)                                # r = 2 mm ball
    m |= ball
    cyl = np.zeros(shape, bool)                                           # flat-ended cylinder r = 2 mm, 8 mm long
    g = np.indices(shape)
    cyl[100:180, ...] = ((g[1, 100:180] - 140) ** 2 + (g[2, 100:180] - 140) ** 2) <= 400
    m |= cyl
    o = ex.morph_open(m, R2)
    assert not o[20:30, 20:120, 20:120].any()                             # sliver gone
    from scipy import ndimage
    lost = ball & ~o                                                      # ball kept (≥ v_min) up to its boundary layer
    assert ndimage.distance_transform_edt(ball)[lost].max() <= 1.5 and (o & ball).sum() > 0.95 * ball.sum()
    assert not (o & ~m).any()                                             # anti-extensive
    assert np.array_equal(ex.morph_open(o, R2), o)                        # idempotent (re-opening changes 0 cells)
    # flat end of the cylinder becomes rounded: corner cell removed, axis cell kept
    assert o[100 + 5, 140, 140] and not o[100, 140 + 19, 140]
    # every kept cell lies in a ball of radius 1.5 mm contained in the original
    from scipy import ndimage
    d = ndimage.distance_transform_edt(~ex.morph_open(m, R2) if False else ~(ndimage.distance_transform_edt(m) > 15))
    assert (d[o] <= 15 + 1e-9).all()


def test_tiled_opening_equals_monolithic():
    n, nt = 32, 3
    h = 0.25
    R2 = ex.ball_r2(1.5, h)                                               # r = 6 cells
    shape = (n * nt,) * 3
    rng = np.random.default_rng(1)
    m = np.zeros(shape, bool)
    for _ in range(6):
        m |= _ball(shape, rng.uniform(15, 80, 3), rng.uniform(3, 12))
    m[40:44, 10:90, 10:90] = True
    grid = ex.Grid(np.zeros(3), np.full(3, (n * nt - 2) * h), h, n)
    assert tuple(grid.nt) == (nt,) * 3
    st = ex.TileStore(grid)
    for t in grid.tile_ids():
        sl = tuple(slice(t[a] * n, (t[a] + 1) * n) for a in range(3))
        st.set(t, np.where(m[sl], 1, 2).astype(np.uint8))                 # void = 1, everything else outside = 2
    ex._G.update(store=st, R2=R2)
    res = ex._pmap(ex._open_tile, grid.tile_ids(), 2)
    out = m.copy()
    for tid, idx, _ in res:
        if idx is None:
            continue
        sl = tuple(slice(tid[a] * n, (tid[a] + 1) * n) for a in range(3))
        blk = out[sl].reshape(-1)
        blk[idx] = False
        out[sl] = blk.reshape(n, n, n)
    assert np.array_equal(out, ex.morph_open(m, R2))


def test_hole_cylinder_psi():
    grid = ex.Grid(np.zeros(3), np.full(3, 6.0), 0.1, 64)
    psi = np.full((64, 64, 64), -0.3, np.float32)
    holes = dict(p=np.array([[3.0, 3.0, 1.0]]), n=np.array([[0.0, 0.0, 1.0]]), s0=np.array([-0.5]), s1=np.array([3.0]),
                 r=np.array([0.75]))
    ex._hole_psi_into(psi, grid, (0, 0, 0), holes, 0.1)
    c = lambda x, y, z: psi[tuple(int(round((v - o) / 0.1 - 0.5)) for v, o in zip((x, y, z), grid.origin))]
    assert c(3.0, 3.0, 2.0) > 0 and c(3.5, 3.0, 2.0) > 0                  # inside the cylinder
    assert c(3.9, 3.0, 2.0) < 0 and c(3.0, 3.0, 4.5) < 0                  # outside radius / beyond the end cap
    assert abs(psi.max() - 0.75) < 0.08


def test_mesh_checks_vs_trimesh(sphere):
    m = trimesh.creation.icosphere(subdivisions=3, radius=5.0)
    c = ex.mesh_checks(np.asarray(m.vertices), np.asarray(m.faces))
    assert c["watertight_manifold"] and c["bodies"] == 1 and c["euler"] == 2
    assert abs(c["volume"] - m.volume) < 1e-9 * m.volume
    f = np.asarray(m.faces)[1:]                                           # delete one face: open mesh
    c2 = ex.mesh_checks(np.asarray(m.vertices), f)
    assert not c2["watertight_manifold"] and not trimesh.Trimesh(m.vertices, f, process=False).is_watertight
    g = np.asarray(m.faces).copy(); g[3] = g[3][::-1]                     # one flipped face: inconsistent winding
    assert not ex.mesh_checks(np.asarray(m.vertices), g)["watertight_manifold"]


def test_drop_material_islands():
    big = trimesh.creation.icosphere(subdivisions=2, radius=5.0)
    spk = trimesh.creation.icosphere(subdivisions=1, radius=0.05); spk.apply_translation([20, 0, 0])
    void = trimesh.creation.icosphere(subdivisions=2, radius=2.0); void.invert()
    m = trimesh.util.concatenate([big, spk, void])
    F, info = ex.drop_material_islands(np.asarray(m.vertices), np.asarray(m.faces), 1.0)
    assert info["bodies_before"] == 3 and info["islands_dropped"] == 1        # speck dropped, void (negative volume) kept
    assert len(F) == len(big.faces) + len(void.faces)


def test_axis_aligned_faces_single_body(tmp_path):
    """Faces at 'nice' coordinates must not leave single-node bubbles (grid is offset from the CAD lattice)."""
    m = trimesh.creation.box(extents=(10.0, 10.0, 10.0))
    m.apply_translation([5.0, 5.0, 5.0])
    cfg = ex.ExportConfig(h_x=0.1, tile=64, block=16, workers=_workers(), wall_hist=False, write_stl=False,
                          log=lambda *_: None)
    out = ex.baselines(np.asarray(m.vertices), np.asarray(m.faces), str(tmp_path), cfg)
    assert out["original"]["mesh"]["bodies"] == 1 and out["uniform_shell"]["mesh"]["bodies"] == 2
    assert abs(out["original"]["mesh"]["volume"] / 1000.0 - 1) < 1e-3


def test_uniform_shell_and_original(sphere, tmp_path):
    V, T, N, V0 = sphere
    cfg = ex.ExportConfig(h_x=0.1, tile=64, block=16, workers=_workers(), wall_samples=5000, log=lambda *_: None)
    out = ex.baselines(V, T, str(tmp_path), cfg)
    t_an = (2 ** (1 / 3) - 1) * R0                                        # (R + t)^3 = 2 R^3
    s = out["uniform_shell"]
    assert abs(s["t_final_mm"] - t_an) < 0.02
    assert s["ok"] and abs(s["mesh"]["volume"] / s["V0_mm3"] - 1) < 2e-3
    assert s["mesh"]["bodies"] == 2                                       # outer skin and the enclosed original volume
    assert abs(s["wall"]["median_nominal"] - t_an) < 0.04
    o = out["original"]
    assert o["ok"] and abs(o["mesh"]["volume"] / V0 - 1) < 2e-3 and o["mass_correction"]["delta"] == 0.0


def test_closed_design_is_the_original(sphere, tmp_path):
    """All vertices closed and d = 0: ψ = φ, V = V₀ without correction, no Γ_h queries needed."""
    V, T, N, V0 = sphere
    nv = len(V)
    dz = ex.Design(V, T, V, T, d=np.zeros(nv), a=np.zeros(nv), d_hi=np.full(nv, 3.0), active=np.ones(nv, bool),
                   closed=np.ones(nv, bool))
    assert dz.trivial_solid and dz.Vh is None
    cfg = ex.ExportConfig(h_x=0.2, tile=32, block=8, workers=_workers(), wall_hist=False, write_stl=False,
                          log=lambda *_: None)
    s = ex.Exporter(dz, cfg, normals=N).run(str(tmp_path))
    assert s["ok"] and abs(s["mesh"]["volume"] / V0 - 1) < 2e-3 and s["mass_correction"]["delta"] == 0.0


# -- B2 rod through Context, hollow xi, end to end ----------------------------------------------------------------------
def test_b2_rod_end_to_end(tmp_path):
    from cshell import config
    from cshell.optimise import Context
    cfg = config.load(str(ROOT / "examples/cshell/bench/B2_slender.json"))
    ctx = Context(cfg, log=lambda *_: None)
    nv = ctx.surf.n_vertices
    xi = np.concatenate([np.full(nv, 0.5), np.zeros(nv)])               # ξᵈ = 0.5, ξᵃ = 0
    s = ex.export_design(ctx, xi, str(tmp_path), h_x=0.1, tile=64, block=16, workers=_workers(), log=lambda *_: None)
    assert s["mesh"]["watertight_manifold"] and abs(s["mesh"]["volume"] / s["V0_mm3"] - 1) <= 2e-3
    assert s["voids"]["without_hole"] == 0 and s["mesh"]["bodies"] == 1
