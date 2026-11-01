"""PLAN_single_run §2-§3: GPU export (`export.GpuExporter`). Analytic checks:
  - convex polyhedron offset by d (closed design, no mass correction): mesh volume vs the Steiner formula
    V + A d + M d² + 4π/3 d³ (M = ½ Σ_edges ℓ θ), bounds vs the polyhedron's bounds ± d (axis order of the marching
    cubes), watertight single body, ψ (ramp) volume vs mesh volume;
  - hollow sphere with escape holes across several tiles: watertight, one body (the void is vented), δ as the shell
    formula predicts, |V_mesh/V_target − 1| ≤ 1e-3 after the mass correction, ψ volume vs mesh volume;
  - smooth opening kernel on an analytic void (ball ∪ thin slab): slab removed, ball unchanged, 1-Lipschitz (no
    terraces), idempotent;
  - ambiguous-face rule: CuPy and NumPy give the same field."""
import math

import numpy as np
import pytest
import trimesh

from cshell import export as ex

pytestmark = pytest.mark.gpu


def _steiner(V, T, d):
    m = trimesh.Trimesh(V, T, process=False)
    ang = m.face_adjacency_angles                                   # angle between the two face normals
    e = m.vertices[m.face_adjacency_edges]
    M = 0.5 * float((np.linalg.norm(e[:, 1] - e[:, 0], axis=1) * ang).sum())
    return float(m.volume) + float(m.area) * d + M * d * d + 4.0 / 3.0 * math.pi * d ** 3


def test_gpu_offset_polyhedron_steiner(tmp_path):
    m = trimesh.creation.icosphere(subdivisions=3, radius=1.0)
    V = np.asarray(m.vertices) * [9.0, 6.0, 4.5] + [13.3, -5.1, 2.7]     # convex, off-centre, anisotropic
    T = np.asarray(m.faces)
    nv, d = len(V), 0.6
    dz = ex.Design(V, T, V, T, d=np.full(nv, d), a=np.zeros(nv), d_hi=np.full(nv, 3.0), active=np.ones(nv, bool),
                   closed=np.ones(nv, bool))
    cfg = ex.ExportConfig(h_x=0.1, tile=64, block=16, mass_correct=False, wall_samples=5000, log=print)
    s = ex.GpuExporter(dz, cfg).run(str(tmp_path))
    mesh = s["mesh"]
    assert mesh["watertight_manifold"] and mesh["bodies"] == 1, mesh
    Vs = _steiner(V, T, d)
    print("steiner", Vs, "mesh", mesh["volume"], "ramp", s["mass_correction"]["V_ramp_final_mm3"])
    assert abs(mesh["volume"] / Vs - 1) < 1.5e-3
    assert abs(s["mass_correction"]["V_ramp_final_mm3"] / mesh["volume"] - 1) < 2e-3        # ψ volume vs mesh
    stl =trimesh.load(str(tmp_path / "final_design.stl"), process=False)
    lo, hi = stl.bounds
    assert np.allclose(lo, V.min(0) - d, atol=0.05) and np.allclose(hi, V.max(0) + d, atol=0.05), (lo, hi)


@pytest.fixture(scope="module")
def hollow(tmp_path_factory):
    m = trimesh.creation.icosphere(subdivisions=4, radius=8.0)
    V, T = np.asarray(m.vertices), np.asarray(m.faces)
    nv = len(V)
    dz = ex.Design(V, T, V, T, d=np.full(nv, 0.9), a=np.full(nv, 2.0), d_hi=np.full(nv, 3.0),
                   active=np.ones(nv, bool), closed=np.zeros(nv, bool))
    out = tmp_path_factory.mktemp("hollow_gpu")
    cfg = ex.ExportConfig(h_x=0.1, tile=64, block=16, wall_samples=5000, log=print)
    e = ex.GpuExporter(dz, cfg, normals=V / np.linalg.norm(V, axis=1, keepdims=True))
    return e.run(str(out)), out, float(m.volume)


def test_gpu_hollow_sphere(hollow):
    s, out, Vtess = hollow
    mesh, mc = s["mesh"], s["mass_correction"]
    print({k: mesh[k] for k in ("faces", "bodies", "volume_rel")}, mc["delta"], mc["mesh_steps"], s["voids"])
    assert s["ok"] and mesh["watertight_manifold"]
    assert mesh["bodies"] == 1                                         # void vented by its escape holes
    assert s["voids"]["enclosed"] == 1 and s["voids"]["holes"] == 2
    assert s["V0_mm3"] == pytest.approx(Vtess, rel=1e-12)
    assert abs(mesh["volume_rel"]) <= 1e-3
    assert abs(mc["V_ramp_final_mm3"] / mesh["volume"] - 1) < 2e-3
    # shell (8 + 0.9 + δ)³ − 6³ = 8³ (holes and facets aside): δ ≈ 0.096
    assert 0.05 < mc["delta"] < 0.15
    assert (out / "final_design.stl").exists() and (out / "escape_holes.csv").exists()
    assert (out / "wall_thickness.csv").exists() and s["wall"]["median_nominal"] > 2.0
    assert s["grid_shape"][0] // s["tile"] >= 3                         # several tiles per axis


def _open(s_region, H, n, r, h):
    import cupy as cp
    R = ex.opening_R(r, h)
    rng_ = np.arange(-R, R + 1)
    D = np.stack(np.meshgrid(rng_, rng_, rng_, indexing="ij"), -1).reshape(-1, 3)
    q2 = (D * D).sum(1)
    sel = q2 <= R * R
    o = np.argsort(q2[sel], kind="stable")
    off = cp.asarray(np.ascontiguousarray(D[sel][o], np.int32))
    offd = cp.asarray((np.sqrt(q2[sel][o]) * h).astype(np.float32))
    m = s_region.shape[0]
    out = cp.empty(n ** 3, cp.float32)
    k = ex._gpu_module().get_function("open_void")
    k(*ex._launch(n ** 3), (cp.asarray(s_region, dtype=cp.float32), np.int32(m), np.int32(m), np.int32(H), np.int32(n),
                            np.int32(n), np.int32(n), off, offd, np.int32(len(offd)), np.float32(r),
                            np.float32(r + 2 * h), np.float32(h), np.float32(-2 * h), out))
    return out.get().reshape(n, n, n)


def test_smooth_opening_ball_and_thin_slab():
    h, r = 0.1, 1.5
    R = ex.opening_R(r, h)
    H, n = 24, 96
    assert H >= R
    m = n + 4 * H
    x = (np.arange(m) - m / 2 + 0.5) * h
    X, Y, Z = np.meshgrid(x, x, x, indexing="ij")
    ball = 3.0 - np.sqrt(X * X + Y * Y + Z * Z)                       # void ball, radius 3 (≥ r: kept)
    slab = np.minimum(0.5 - np.abs(Z), X)                              # void slab x > 0, |z| < 0.5 (thin: removed)
    s = np.maximum(ball, slab).astype(np.float32)
    n1 = n + 2 * H
    p1 = _open(s, H, n1, r, h)                                         # first opening on a core with a halo H
    p2 = _open(p1, H, n, r, h)                                         # re-opening of the opened field
    core1 = (slice(H, H + n),) * 3
    a = p1[core1]
    s_c = s[(slice(2 * H, 2 * H + n),) * 3]
    xc = x[2 * H:2 * H + n]
    Xc, _, Zc = np.meshgrid(xc, xc, xc, indexing="ij")
    # idempotent (exactly: the centres {ψ ≥ r} do not change)
    assert np.array_equal(p2, a)
    # thin slab removed away from the ball, ball unchanged away from the slab
    far = (Xc > 3.6) & (np.abs(Zc) < 0.5)
    assert far.any() and (s_c[far] > 0).all() and (a[far] < 0).all()
    away = ((np.abs(Zc) > 1.2) | (Xc < -0.5)) & (s_c > -2 * h)
    assert np.max(np.abs(a[away] - s_c[away])) < 0.2 * h
    assert np.all(a <= s_c + 1e-6)                                     # opening ⊂ void
    # 1-Lipschitz: no terraces in the removed region
    for ax in range(3):
        assert np.max(np.abs(np.diff(a, axis=ax))) <= h * (1 + 1e-4)


def test_ambiguous_faces_cupy_equals_numpy():
    import cupy as cp
    rng = np.random.default_rng(3)
    psi = rng.standard_normal((40, 40, 40)).astype(np.float32)
    a = psi.copy()
    b = cp.asarray(psi)
    ra = ex.resolve_ambiguous_faces(a, np.float32(1e-3))
    rb = ex.resolve_ambiguous_faces(b, np.float32(1e-3))
    assert ra == rb and np.array_equal(a, b.get())


# ---------------------------------------------------------------------------------------------------------------------
# sr-q2: minimum wall (wall_cap + material opening) and voids without a clean escape-hole site
# ---------------------------------------------------------------------------------------------------------------------
def test_wall_cap_kernel_tilted_plane():
    """Outside {n·x > 0}, void branch p = −n·x − 1 (wall 1.0 along n): wall_cap moves the void wall to n·x = −t_min
    (true normal distance, any tilt) within h/2, and leaves cells beyond the search radius unchanged."""
    import cupy as cp
    h, tmin = 0.1, 1.5
    Rw = int(math.ceil((tmin + h) / h)) + 2
    rng_ = np.arange(-Rw, Rw + 1)
    D = np.stack(np.meshgrid(rng_, rng_, rng_, indexing="ij"), -1).reshape(-1, 3)
    q2 = (D * D).sum(1)
    sel = q2 <= Rw * Rw
    o = np.argsort(q2[sel], kind="stable")
    off = cp.asarray(np.ascontiguousarray(D[sel][o], np.int32))
    offd = cp.asarray((np.sqrt(q2[sel][o]) * h).astype(np.float32))
    k = ex._gpu_module().get_function("wall_cap")
    H, n = Rw, 48
    m = n + 2 * H
    x = (np.arange(m) - m / 2 + 0.5) * h
    X = np.stack(np.meshgrid(x, x, x, indexing="ij"), -1)
    for nv in ([1.0, 0, 0], [1.0, 1.0, 0], [1.0, 0.6, 0.3]):
        nv = np.asarray(nv) / np.linalg.norm(nv)
        s = X @ nv
        O = s.astype(np.float32)
        p = (-s - 1.0)[(slice(H, H + n),) * 3].astype(np.float32)
        out, dm = cp.empty(n ** 3, cp.float32), cp.empty(n ** 3, cp.float32)
        k(*ex._launch(n ** 3), (cp.asarray(O), cp.asarray(p), np.int32(m), np.int32(m), np.int32(H), np.int32(n),
                                np.int32(n), np.int32(n), off, offd, np.int32(len(offd)), np.float32(tmin),
                                np.float32(h), out, dm))
        out = out.get().reshape(n, n, n)
        sc = s[(slice(H, H + n),) * 3]
        near = np.abs(sc + tmin) < h                                    # around the new wall: distance-like
        assert np.max(np.abs(out[near] - (-sc[near] - tmin))) < 0.5 * h, nv
        clear = np.abs(sc + tmin) > 0.5 * h                             # void exactly beyond t_min (sign)
        assert np.array_equal((out > 0)[clear], (sc < -tmin)[clear]), nv
        deep = sc < -(tmin + 2 * h + Rw * h)                            # beyond the search radius: unchanged
        if deep.any():
            assert np.allclose(out[deep], p[deep])
        assert np.all(out <= p + 1e-6)                                  # only ever removes void


def _sphere(sub=4, R=8.0):
    m = trimesh.creation.icosphere(subdivisions=sub, radius=R)
    return np.asarray(m.vertices), np.asarray(m.faces)


def _run(dz, tmp_path, name, **kw):
    V = dz.Vh
    cfg = ex.ExportConfig(**{**dict(h_x=0.1, tile=64, block=16, wall_samples=40000, log=print), **kw})
    return ex.GpuExporter(dz, cfg, normals=V / np.linalg.norm(V, axis=1, keepdims=True)).run(str(tmp_path / name))


@pytest.fixture(scope="module")
def tilted(tmp_path_factory):
    """Sphere R 8 whose outer offset d steps from 0 to 1.4 mm across the equator over 1 mm (slope ≈ 1.4) with
    a = 1.6 − d: the wall a + d = 1.6 along the radius, ≈ 1.6/√(1 + 1.4²) ≈ 0.93 mm normal to the tilted band."""
    V, T = _sphere()
    nv = len(V)
    u = np.clip((V[:, 2] + 0.5) / 1.0, 0, 1)
    d = 1.4 * u * u * (3 - 2 * u)
    a = 1.6 - d
    dz = ex.Design(V, T, V, T, d=d, a=a, d_hi=np.full(nv, 3.0), active=np.ones(nv, bool), closed=np.zeros(nv, bool))
    out = tmp_path_factory.mktemp("tilted")
    kw = dict(t_min=1.2, b_v=0.2, void_min_mm3=5.0, mass_correct=False)   # V ≠ the solid's: no δ (geometry only)
    return _run(dz, out, "off", min_wall=False, **kw), _run(dz, out, "on", **kw)


def test_min_wall_tilted_band(tilted):
    off, on = tilted
    lim = 1.2 - 0.2                                                     # t_min − b_v
    print("off", off["wall_off_rims"], "on", on["wall_off_rims"], on["min_wall"])
    assert off["wall_off_rims"]["min_nominal"] < lim                   # discriminates: thin walls without the cap
    for s in (off, on):
        assert s["mesh"]["watertight_manifold"] and s["mesh"]["bodies"] == 1
    assert on["min_wall"]["void_to_material_mm3"] > 0
    assert on["wall_off_rims"]["min_nominal"] >= lim                    # every wall off the hole rims ≥ t_min − b_v
    assert on["wall_off_rims"]["n_nominal_below_t_min_minus_b_v"] == 0


@pytest.fixture(scope="module")
def fin(tmp_path_factory):
    """Hollow sphere (d 0.9, a 2) with one closed vertex 2-ring at the north pole: the closure rule (cells whose
    closest triangle is wholly closed have no void) leaves a cone of material, ~1 mm wide at the wall, through the
    void to the centre (the SR_VAL2 MBB/Cantilever void-void fins)."""
    V, T = _sphere()
    nv = len(V)
    top = int(np.argmax(V[:, 2]))
    closed = np.zeros(nv, bool)
    closed[top] = True
    for _ in range(2):
        closed |= np.isin(np.arange(nv), T[np.isin(T, np.flatnonzero(closed)).any(1)].ravel())
    dz = ex.Design(V, T, V, T, d=np.full(nv, 0.9), a=np.full(nv, 2.0), d_hi=np.full(nv, 3.0),
                   active=np.ones(nv, bool), closed=closed)
    out = tmp_path_factory.mktemp("fin")
    return _run(dz, out, "off", min_wall=False), _run(dz, out, "on")


def test_min_wall_removes_closure_fins(fin):
    off, on = fin
    print("off", off["wall_off_rims"], off["wall"]["min_nominal"], "on", on["wall_off_rims"], on["wall"]["min_nominal"],
          on["min_wall"])
    lim = 1.5 - 0.2
    assert off["wall_off_rims"]["min_nominal"] < lim                   # the fin is there without the fix
    assert on["min_wall"]["fins_removed_mm3"] > 0
    assert on["wall_off_rims"]["min_nominal"] >= lim
    for s in (off, on):
        assert s["ok"] and s["mesh"]["watertight_manifold"] and s["mesh"]["bodies"] == 1
        assert abs(s["mesh"]["volume_rel"]) <= 1e-3


def test_hollow_sphere_multi_shell(tmp_path):
    """§8 (escape holes optional, off by default): a hollow sphere without holes exports as 1 material body with
    2 shells (outer + inward-facing void shell), watertight, the signed volume = outer − void, and the mass correction
    converges; with holes on, a void too narrow for the hole's inner cap (3.6 < 2 (r + h) = 4.2, hole_d 4) is
    reported unvented, not filled (fill_unholed off); the sr-q2 fill stays available explicitly."""
    V, T = _sphere(sub=3, R=6.0)
    nv = len(V)
    dz = ex.Design(V, T, V, T, d=np.zeros(nv), a=np.full(nv, 4.2), d_hi=np.full(nv, 3.0),
                   active=np.ones(nv, bool), closed=np.zeros(nv, bool))
    kw = dict(void_min_mm3=5.0, wall_samples=5000, hole_d=4.0)
    off = _run(dz, tmp_path, "off", holes=False, **kw)
    m = off["mesh"]
    print("holes off:", m["shells"], m["bodies"], m["void_shells"], m["shell_volumes_mm3"], off["mass_correction"])
    assert m["watertight_manifold"] and m["shells"] == 2 and m["bodies"] == 1 and m["void_shells"] == 1
    assert abs(sum(m["shell_volumes_mm3"]) - m["volume"]) < 1e-6 * abs(m["volume"])
    assert m["shell_volumes_mm3"][1] < 0 and abs(m["volume_rel"]) <= 1e-3 and off["ok"]
    assert off["mass_correction"]["ok"] and not off["mass_correction"].get("fail")
    sealed = _run(dz, tmp_path, "sealed", holes=True, **kw)            # holes on, no site: reported, kept hollow
    assert sealed["voids"]["holes"] == 0 and sealed["voids"]["without_hole"] == 1
    assert sealed["mesh"]["shells"] == 2 and sealed["mesh"]["bodies"] == 1 and sealed["voids"]["enclosed"] == 1
    filled = _run(dz, tmp_path, "filled", holes=True, fill_unholed=True, **kw)
    print(sealed["voids"], sealed["mesh"]["bodies"], filled["voids"], filled["mesh"]["bodies"])
    assert filled["voids"]["filled_no_hole"] == 1 and filled["voids"]["enclosed"] == 0
    assert filled["mesh"]["bodies"] == 1 and filled["ok"]
    assert abs(filled["mesh"]["volume_rel"]) <= 1e-3                   # material conserved
    assert filled["mass_correction"]["delta"] < sealed["mass_correction"]["delta"]   # filled void paid by δ < 0
