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
    R = int(math.ceil((r + 2 * h) / h))
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
                            np.float32(r + 2 * h), np.float32(h), np.float32(-(R + 1) * h), out))
    return out.get().reshape(n, n, n)


def test_smooth_opening_ball_and_thin_slab():
    h, r = 0.1, 1.5
    R = int(math.ceil((r + 2 * h) / h))
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
