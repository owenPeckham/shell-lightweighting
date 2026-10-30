"""Everything after the optimiser for one part (THEORY §5.4, §7.4 S2, §8, §10.5): `python -m cshell collect <run_dir>`.

Pipeline (`collect`)
  1. Starts: one row per start from history.csv, checkpoints and xi_<start>_phase<k>.npy; the best *feasible* start
     (final phase complete, |V/V₀ − 1| ≤ mass_tol) by its final optimiser J (§7.4).
  2. Parametric FE of the best design on the optimisation grid (the optimiser's own as-built pipeline): material
     von Mises for hole siting, W̄ for the optimality page, buckling modes.
     The uniform-shell t (mount faces solid, §7.4 S1 D1) is solved for constant mass on the same band; the model
     is then released (host and device) before the export so the forked export workers do not inherit it.
  3. Export (§8, `export.export_design`): best_candidate.stl (+ escape_holes.csv, wall_thickness.csv,
     best_candidate_export.json); original.stl, uniform_shell.stl (`export.baselines`, t for constant mass).
     A best export that fails its checks (`export_verdict`) is kept for inspection, not re-scored, and the solid is
     kept with reason "export failed: …" (D2).
  4. Re-score on the export geometry (§5.4, §8.6.1) for the solid, the uniform shell and the best export, at
     (i) the optimisation grid, (ii) that grid shifted by ½h_e in x, y, z, (iii) h_e = 0.5 mm, h_g = 0.125 mm.
     Each exported mesh is turned into ψ on the phase band by its exact signed distance φ_E at the cell centres plus
     the as-built increment of the same design (§3.9):  ψ_ab = φ_E + Δ,  Δ = b_o on the outer branch of
     max(φ − d, −φ − a), Δ = a − a_ab = b_v S((r_MA − a)/b_v) on the void branch (interpolated with P, §3.6);
     χ = H_ε(−ψ_ab), ρ_e = mean χ, then the optimiser's FE (`objective.evaluate`). The solid uses φ itself (Δ = b_o),
     so its re-score is the optimiser's baseline exactly. The cell maps run on the device for the GPU chain
     (`DeviceDesignMap.rescore`), the STL is read unwelded (`stl_soup`) one mesh at a time, BCs act only on nodes
     that touch material and unanchored material islands are removed (`DesignBCs`). All designs are scored against the solid on the *same*
     grid with the solid's normalisation c = Q_q/σ_PN (§4.4(b)); responses are kept raw so that `sensitivity.csv`
     needs no new FE (λ_b ∝ κ_b, λ_s ∝ σ_f).
  5. Body-fitted check (`validate.validate_design`, CalculiX C3D10, §8.6.2) of the solid and the best export;
     a miss on a ratio blocks acceptance.
  6. Decision (§5.4): accept iff, at every resolution run, J* > ln(1 + g_min), K/K₀ ≥ f_K and P/P₀ ≥ f_P, the
     export is valid (mass correction, watertight, R1) and no blocking validation miss. final_design.stl is the best
     candidate if accepted, else a copy of original.stl, and `reason` says why.

Outputs in <run_dir>: summary.json, summary.txt, validation.csv, sensitivity.csv, original.stl, uniform_shell.stl,
best_candidate.stl, final_design.stl, fields.vtp, buckling_modes.npz, bc_faces.npz, collect_log.txt (+ the export's
escape_holes.csv, wall_thickness.csv, *_export.json).

summary.json schema ("cshell.summary/1"; units mm, N, MPa, g, s; every key present, null when not computed)
  schema            "cshell.summary/1"
  name              part name (config "name")
  accepted          bool: final_design.stl is the cshell design
  outcome           "cshell design accepted" | "original SOLID part kept"
  reason            human-readable reason (all failed criteria joined by "; ")
  criteria          {g_min, ln_1_g_min, f_K, f_P}
  mass              {target (g, solid V₀ on the export grid), final (g, final_design mesh), error (V/V₀ − 1 of the
                    best-candidate mesh), V0_mm3, density_g_mm3}
  best              {start, J_opt (optimiser's last J), D_comp, xi_file, phase} | null
  starts            [{start, status ("done"|"incomplete"|"missing"), feasible, phases_done, iterations, J_opt,
                      K_K0_opt, P_P0_opt, mass_err, hours, closed_area_frac, w_min, w_max}]
  comparison        [{label, source, J, K_K0, P_P0, governing, w_min, w_max}] in the order
                    solid, uniform shell, vshell best, cshell <start> (optimiser values) ..., cshell export
                    (re-scored at the optimisation grid); `source` ∈ {"rescored", "optimiser", "vshell"}
  resolutions       [{label ("<res>: <design>"), res ("opt"|"shift"|"coarse"), design ("solid"|"uniform_shell"|
                      "best"), h_e, h_g, shift_mm, J, K_K0, P_P0, C_Nmm, K_N_mm, lam_s, lam_b, lam, lam1,
                      governing, sigma_pn, sigma_q, V_mm3, mass_err, dof, seconds}]
  checks            [{res, J, K_K0, P_P0, pass_J, pass_K, pass_P}] (best export at each resolution)
  load_cases        [{case, lam_s, lam_b, softmin}] at the optimisation grid (one summed case, D3), solid and best
  Lambda            Λ/2 estimate for the optimality page: area-weighted median of W̄ over open active vertices (N/mm²)
  export            {best: {...best_candidate_export.json without timing}, uniform_shell_t_mm, delta_mm, voids,
                     holes, wall: {min_nominal, min_as_built, median_nominal}}
  validation        {ran, engine, blocking_miss, n_rows, n_failed, file} | {ran: false, why}
  sensitivity       {file, accepted_fraction, kappa_b: [...], sigma_f_pct: [...]}
  vshell            {run_dir, J, K_K0, P_P0, accepted, s_mm} | null
  resources         {gpu_model, peak_gpu_GiB, gpu_hours (Σ starts' hours_used + collect), collect_s}
  timings_s         {context, starts, param_fe, export_baselines, export_best, rescore_<res>, validation, total}
  files             {name: relative path} of every file written

sensitivity.csv: kappa_b, sigma_f (percent change), sigma_f_MPa, J_star (best export, optimisation grid), J_min,
  K_K0_min, P_P0_min (over the resolutions), accepted (0/1; validation and export verdicts held fixed).
fields.vtp (Γ_h, point data): xi (= ξᵃ), xi_d, xi_a, xit_a, d, a, w (= a + d where open, NaN where closed), r_MA,
  d_hi, passive, stress (best, material von Mises just inside the design surface), stress_solid, W_bar.
buckling_modes.npz: V, T (Γ_h); solid_mag (3, N_v), design_mag (3, N_v) = |mode| at the node nearest each surface
  point (max-normalised), solid_lambda, design_lambda (λ_j = 1/μ_j, most critical first).
bc_faces.npz: V, T (original tessellation), support (bool per triangle), load (int per triangle: load index, −1 none).

S2 vshell start (§7.4) is `vshell_start_xi`, called from `optimise.Run.initial_xi`.
"""
from __future__ import annotations

import copy
import csv
import dataclasses
import glob
import json
import math
import os
import shutil
import time
from contextlib import contextmanager
from pathlib import Path

import numpy as np

ROOT = Path(__file__).resolve().parents[1]
SCHEMA = "cshell.summary/1"
KAPPAS = (0.6, 0.7, 0.85, 1.0)
SIGMA_PCT = (-20, -10, 0, 10, 20)
COARSE = dict(h_e=0.5, h_g=0.125)


def _h(a):
    """Host numpy view of a NumPy/CuPy array."""
    return a.get() if hasattr(a, "get") else np.asarray(a)


def _get(cfg: dict, path: str, default):
    d = cfg
    for k in path.split("."):
        if not isinstance(d, dict) or k not in d:
            return default
        d = d[k]
    return d


# =====================================================================================================================
# S2 vshell start (THEORY §7.4)
# =====================================================================================================================
def vshell_run_dir(cfg: dict) -> Path:
    """`optimiser.vshell_run` (default examples/runs/<name>, relative paths against the repository root)."""
    p = _get(cfg, "optimiser.vshell_run", None) or f"examples/runs/{cfg['name']}"
    p = Path(p)
    return p if p.is_absolute() else ROOT / p


def load_vshell(run_dir) -> dict:
    """vshell's best design: segmentation cells (seg_xyz, seg_label from arrays.npz), θ per region (1 solid, 0 hollow,
    from summary.json 'best hollow candidate'), its outer offset s and its own scores."""
    run_dir = Path(run_dir)
    S = json.load(open(run_dir / "summary.json"))
    z = np.load(run_dir / "arrays.npz")
    if "seg_xyz" not in z.files or "seg_label" not in z.files:
        raise KeyError(f"{run_dir}/arrays.npz has no seg_xyz/seg_label: cannot map vshell regions")
    best = S.get("best hollow candidate") or {}
    hollow = set(int(i) for i in best.get("hollow regions", []))
    xyz = z["seg_xyz"].astype(np.float64)
    lab = z["seg_label"].astype(np.int64)
    u = np.unique(np.round(xyz[:, 0], 6))
    h_seg = float(np.min(np.diff(u))) if len(u) > 1 else 1.0
    return dict(xyz=xyz, label=lab, hollow=hollow, s=float(best.get("outer offset / strut wall s mm", 0.0)),
                h_seg=h_seg, summary=S, accepted="accepted" in str(S.get("decision", {}).get("outcome", "")).lower()
                and "hollow" in str(S.get("decision", {}).get("outcome", "")).lower(), best=best)


def vshell_start_xi(ctx, cfg: dict | None = None, log=print, vs: dict | None = None) -> np.ndarray:
    """ξ = [ξᵈ, ξᵃ] of the S2 start: θ_i = θ of the vshell region containing p_i − (r_MA,i/2) n_i (nearest
    segmentation cell); filtered targets ξ̃ᵃ = θ, ξ̃ᵈ = s(1 − θ)/d^hi (clipped to [0, 1]); raw ξ = F⁻¹ ξ̃ clipped
    (phase-0 filter and d^hi, as `optimise.PhaseModel`). Mass is restored by the caller."""
    from scipy.spatial import cKDTree
    from . import config as ccfg
    cfg = cfg or ctx.cfg
    vs = vs or load_vshell(vshell_run_dir(cfg))
    s = ctx.surf
    nv = s.n_vertices
    N = s.projected_normals() if hasattr(s, "projected_normals") else s.N
    q = s.V - 0.5 * ctx.r_ma[:, None] * N
    dist, j = cKDTree(vs["xyz"]).query(q)
    lab = vs["label"][j]
    theta = np.where(np.isin(lab, list(vs["hollow"])), 0.0, 1.0)
    far = dist > 2.0 * vs["h_seg"]
    if far.any():
        theta[far] = 1.0
        log(f"[cshell] vshell start: {int(far.sum())} of {nv} probe points are > 2 h_seg from any vshell cell: "
            "θ = 1 (solid) there")
    dsg = cfg["design"]
    D0 = dsg["outward_extent_initial_mm"]
    d_hi = np.minimum.reduce([ctx.r_out - dsg["gap_mm"], ctx.keepout - dsg["keep_out_clearance_mm"],
                              np.full(nv, D0)]).clip(0.0)
    if not dsg["allow_outward"]:
        d_hi[:] = 0.0
    xt_d = np.where(d_hi > 1e-9, vs["s"] * (1.0 - theta) / np.maximum(d_hi, 1e-9), 0.0).clip(0.0, 1.0)
    filt = s.filter(ccfg.phase_value(cfg["filter"]["radius_mm"], 0))
    xd = np.clip(filt.invert(xt_d), 0.0, 1.0)
    xa = np.clip(filt.invert(theta), 0.0, 1.0)
    A = s.A
    log(f"[cshell] vshell start: s = {vs['s']:.3f} mm, hollow regions {sorted(vs['hollow'])}, hollow area share "
        f"{float((A * (theta == 0)).sum() / A.sum()):.3f}")
    return np.concatenate([xd, xa])


# =====================================================================================================================
# Run-directory reading
# =====================================================================================================================
def read_history(run_dir) -> list[dict]:
    from .optimise import merge_history
    merge_history(run_dir)                               # per-start files -> history.csv (atomic)
    p = os.path.join(run_dir, "history.csv")
    if not os.path.exists(p):
        return []
    with open(p) as f:
        return list(csv.DictReader(f))


def _f(x):
    try:
        return float(x)
    except (TypeError, ValueError):
        return float("nan")


def latest_checkpoint(run_dir, start):
    files = sorted(glob.glob(os.path.join(run_dir, "checkpoints", f"{start}_*.npz")))
    return np.load(files[-1]) if files else None


def start_table(run_dir, cfg: dict) -> list[dict]:
    """One row per configured start (and any other start found in history.csv)."""
    hist = read_history(run_dir)
    n_ph = len(cfg["grid"]["phases"])
    starts = list(dict.fromkeys(list(cfg["optimiser"]["starts"]) + [r["start"] for r in hist]))
    rows = []
    for s in starts:
        h = [r for r in hist if r["start"] == s]
        z = latest_checkpoint(run_dir, s)
        done = [k for k in range(n_ph) if os.path.exists(os.path.join(run_dir, f"xi_{s}_phase{k}.npy"))]
        status = "missing" if not h else ("done" if (n_ph - 1) in done else "incomplete")
        last = h[-1] if h else {}
        # optimiser J of the final phase: last row of that phase (history may contain repeated rows after a resume)
        fin = [r for r in h if int(r["phase"]) == n_ph - 1]
        lr = fin[-1] if fin else last
        mass_err = _f(lr.get("mass_err"))
        feasible = status == "done" and abs(mass_err) <= cfg["optimiser"]["mass_tol"] * 1.0001
        rows.append(dict(start=s, status=status, feasible=bool(feasible), phases_done=len(done), iterations=len(h),
                         J_opt=_f(lr.get("J")), K_K0_opt=1.0 / _f(lr.get("C_C0")) if lr else float("nan"),
                         P_P0_opt=_f(lr.get("lam_ratio")), mass_err=mass_err,
                         hours=float(z["hours_used"]) if z is not None else float("nan"),
                         D_comp=float(z["D_comp"]) if z is not None else cfg["design"]["outward_extent_initial_mm"],
                         xi_file=f"xi_{s}_phase{n_ph - 1}.npy" if status == "done" else None))
    return rows


def pick_best(rows: list[dict]):
    c = [r for r in rows if r["feasible"] and np.isfinite(r["J_opt"])]
    return max(c, key=lambda r: r["J_opt"]) if c else None


# =====================================================================================================================
# Scoring grids
# =====================================================================================================================
@contextmanager
def grid_shift(s: float):
    """Shift the band/FE grid origin by −s in x, y, z (the geometry stays put): patches `band.make_grid` while a
    PhaseModel is built. Verified by the caller against model.origin."""
    from . import band
    if s == 0:
        yield
        return
    orig = band.make_grid

    def mg(V, h, D_comp):
        g = orig(V, h, D_comp)
        n = int(math.ceil(s / h)) + 1
        return band.Grid(g.origin - s, g.h, tuple(int(x) + n for x in g.shape), g.n_pad)

    band.make_grid = mg
    try:
        yield
    finally:
        band.make_grid = orig


def resolutions(cfg: dict) -> list[dict]:
    """(i) optimisation grid, (ii) shifted by ½h_e, (iii) h_e = 0.5 mm unless equal to (i) or disabled."""
    last = cfg["grid"]["phases"][-1]
    out = [dict(res="opt", h_e=float(last["h_e"]), h_g=float(last["h_g"]), shift=0.0),
           dict(res="shift", h_e=float(last["h_e"]), h_g=float(last["h_g"]), shift=0.5 * float(last["h_e"]))]
    co = _get(cfg, "validation.coarse_grid", COARSE)
    if cfg["validation"].get("second_resolution", True) and co and abs(co["h_e"] - last["h_e"]) > 1e-12:
        out.append(dict(res="coarse", h_e=float(co["h_e"]), h_g=float(co["h_g"]), shift=0.0))
    return out


def make_model(ctx, r: dict, D_comp: float, log=print):
    """PhaseModel of the last continuation phase on the grid `r` (h_e, h_g, origin shift)."""
    from .optimise import PhaseModel
    cfg = copy.deepcopy(ctx.cfg)
    ph = len(cfg["grid"]["phases"]) - 1
    cfg["grid"]["phases"][ph] = dict(cfg["grid"]["phases"][ph], h_e=r["h_e"], h_g=r["h_g"])
    # geometry chain: config grid.geometry_device (auto = GPU when present; F1: CPU and GPU chains agree to 1e-15)
    tg = cfg["fe"].setdefault("twogrid", {})
    tg["maxiter"] = max(int(tg.get("maxiter", 100)), 300)   # re-scoring PCG headroom (P4(ii), F4)
    c2 = copy.copy(ctx)
    c2.cfg = cfg
    with grid_shift(r["shift"]):
        m = PhaseModel(c2, ph, D_comp, log)
    # host memory (P5): the band's φ / closest-triangle copies live on in the Discretisation (or DeviceDisc); only the
    # cell index (centres for the mesh SDF) and the centring offset are kept here
    m.band.phi = m.band.tri = m.band.bary = None
    m.band._be = None
    if r["shift"]:
        from .band import make_grid
        o = make_grid(ctx.V, r["h_g"], D_comp).origin
        if not np.allclose(m.origin, o - r["shift"], atol=1e-9):
            raise RuntimeError("grid shift was not applied (PhaseModel no longer builds its grid via band.make_grid)")
    return m


class _ValidationFailed(Exception):
    pass


def _release(model) -> None:
    if hasattr(model, "release"):
        model.release()
    elif model.fe is not None:
        model.fe.free()


def scoring_settings(model):
    """Final-phase settings with stress and buckling always evaluated (floors need P even if w_P = 0)."""
    st = model.settings(1.0)
    return dataclasses.replace(st, buckling=bool(model.ctx.cfg["capacity"]["buckling"]))


# =====================================================================================================================
# Host memory (PLAN_collect_fixes P5)
# =====================================================================================================================
def rss_gib() -> float:
    """Resident set size of this process (GiB, /proc/self/status)."""
    try:
        for line in open("/proc/self/status"):
            if line.startswith("VmRSS:"):
                return int(line.split()[1]) / 2 ** 20
    except OSError:
        pass
    return float("nan")


def trim_host_memory() -> None:
    """Return freed host memory to the OS: cyclic GC, CuPy's device and pinned pools, glibc malloc_trim(0)."""
    import ctypes
    import gc
    gc.collect()
    try:
        import cupy as cp
        cp.get_default_memory_pool().free_all_blocks()
        cp.get_default_pinned_memory_pool().free_all_blocks()
    except Exception:
        pass
    try:
        ctypes.CDLL("libc.so.6").malloc_trim(0)
    except OSError:
        pass


def free_model(model) -> None:
    """Release a PhaseModel's FE engine and device band, then its host arrays (caller drops its references)."""
    _release(model)
    for k in ("band", "disc", "dm", "ds", "grid_fe", "fe", "fixed", "load", "rho0"):
        if hasattr(model, k):
            setattr(model, k, None)
    trim_host_memory()


def _grid_info(model):
    """What DenseLookup / mode_magnitudes need of a model, without holding the model."""
    from types import SimpleNamespace
    return SimpleNamespace(origin=np.asarray(model.origin, float), h_e=model.h_e, elem_shape=model.elem_shape,
                           node_shape=model.node_shape)


# =====================================================================================================================
# Re-scoring cell maps (THEORY §8.6.1): host for the CPU chain, `DeviceDesignMap.rescore` for the GPU chain
# =====================================================================================================================
def stl_soup(path, c0=None) -> np.ndarray:
    """Vertices (3n, 3) float32 of the STL's triangles, unwelded (binary STL through np.memmap, chunked; ASCII via
    trimesh). c0: subtract it in float64 before rounding to float32 (the Warp frame of `band.make_backend`).
    No welding is needed for the winding-number SDF (triangle-based) or the volume."""
    size = os.path.getsize(path)
    with open(path, "rb") as f:
        head = f.read(84)
    n = int(np.frombuffer(head[80:84], "<u4")[0]) if len(head) == 84 else -1
    if n >= 0 and 84 + 50 * n == size:
        dt = np.dtype([("n", "<f4", 3), ("v", "<f4", (3, 3)), ("attr", "<u2")])
        mm = np.memmap(path, dt, "r", offset=84, shape=(n,))
        out = np.empty((3 * n, 3), np.float32)
        ch = 1 << 22
        for s in range(0, n, ch):
            v = np.asarray(mm["v"][s:s + ch]).reshape(-1, 3)
            out[3 * s:3 * s + len(v)] = v if c0 is None else (v.astype(np.float64) - c0)
        del mm
    else:
        import trimesh
        tm = trimesh.load(path, process=False, force="mesh")
        v = np.asarray(tm.triangles, np.float64).reshape(-1, 3)
        out = (v if c0 is None else v - c0).astype(np.float32)
    out += np.float32(0.0)                                   # −0 → +0 (welding by bit pattern)
    return out


def weld(soup: np.ndarray):
    """Vertex weld of an unwelded (3n, 3) float32 soup by exact coordinates: (V float64, T int64)."""
    u = soup.view(np.uint32)
    order = np.lexsort((u[:, 2], u[:, 1], u[:, 0]))
    us = u[order]
    new = np.ones(len(us), bool)
    new[1:] = np.any(us[1:] != us[:-1], axis=1)
    del us
    idx = np.cumsum(new) - 1
    inv = np.empty(len(order), np.int64)
    inv[order] = idx
    V = soup[order[new]].astype(np.float64)
    return V, inv.reshape(-1, 3)


def load_mesh(path):
    """(V, T, volume) of an STL, welded (validation mesher input)."""
    V, T = weld(stl_soup(path))
    a, b, c = V[T[:, 0]], V[T[:, 1]], V[T[:, 2]]
    vol = float(np.einsum("ij,ij->", a, np.cross(b, c))) / 6.0
    return V, T, vol


def mesh_phi(model, soup_c: np.ndarray, chunk: int = 1 << 21):
    """Signed distance (negative inside) of the closed triangle soup `soup_c` (centred by model.band.c0, float32) at
    the band cell centres: float32, on the device for the GPU chain (CuPy), else NumPy."""
    from .band import make_backend
    band = model.band
    c0 = band.c0
    T = np.arange(len(soup_c), dtype=np.int32).reshape(-1, 3)
    be = make_backend(soup_c, T)
    del T
    gpu = model.geometry_device == "gpu"
    if gpu:
        import cupy as cp
        out = cp.empty(band.n, cp.float32)
    else:
        out = np.empty(band.n, np.float32)
    for s in range(0, band.n, chunk):
        sl = slice(s, min(band.n, s + chunk))
        q = be.query(band.centres(sl) - c0, signed=True)[0]
        out[sl] = cp.asarray(q) if gpu else q
    del be
    return out


def design_fields(d, a, closed, r_ma, b_v: float) -> dict:
    """Vertex fields of a re-scored design: d, a, cl (closed 0/1) and sv = b_v S((r_MA − a)/b_v) (THEORY §3.9)."""
    from .design import smoothstep
    S, _ = smoothstep(np, (r_ma - np.asarray(a, np.float64)) / b_v)
    return dict(d=np.asarray(d, np.float64), a=np.asarray(a, np.float64), cl=np.asarray(closed, np.float64),
                sv=b_v * S)


def cell_rescore(model, phiE, f: dict | None, b_o: float, solid: bool = False, want_rho: bool = True):
    """ρ (element box, host float64) of ψ_ab = ψ + Δ and the nominal volume of ψ (THEORY §3.9, §8.6.1).

    ψ = φ_E (phiE, the exported mesh) or, with phiE None, the parametric max(φ − Pd, −φ − Pa) of `f`;
    Δ = b_o on the outer branch of max(φ − Pd, −φ − Pa), P sv on the void branch; cells whose positive-weight
    vertices are all closed use a = A_CLOSED (export closure rule, §8.2). solid: ψ = φ_E (or φ), Δ = b_o."""
    from .export import A_CLOSED
    if model.geometry_device == "gpu":
        model.disc.upload()
        f = f or {}
        return model.dm.rescore(phiE, f.get("d"), f.get("a"), f.get("cl"), f.get("sv"), b_o, A_CLOSED, solid=solid,
                                want_rho=want_rho)
    from .design import heaviside
    D = model.disc
    phi = np.asarray(D.phi, np.float64)
    if solid:
        pn = phi if phiE is None else phiE.astype(np.float64)
        dl = b_o
    else:
        P = D.P
        dc, ac = P @ f["d"], P @ f["a"]
        ac = np.where(P @ f["cl"] >= 1.0 - 1e-9, A_CLOSED, ac)
        o, i = phi - dc, -phi - ac
        del dc, ac
        outer = o >= i
        pn = np.where(outer, o, i) if phiE is None else phiE.astype(np.float64)
        del o, i
        dl = np.where(outer, b_o, P @ f["sv"])
    V = float(heaviside(np, -pn, D.eps)[0].sum()) * D.h_g ** 3 + D.V_extra
    rho = None
    if want_rho:
        rho = np.bincount(D.elem, weights=heaviside(np, -(pn + dl), D.eps)[0], minlength=D.n_elem) / D.k ** 3
    return rho, V


# =====================================================================================================================
# Boundary conditions consistent with the design (PLAN_collect_fixes P1, P4(i))
# =====================================================================================================================
_SHIFTS = [(a, b, c) for a in (0, 1) for b in (0, 1) for c in (0, 1)]


def _node_touch(mat):
    """Nodes with at least one adjacent element in `mat` (element box bool)."""
    es = mat.shape
    out = np.zeros(tuple(n + 1 for n in es), bool)
    for a, b, c in _SHIFTS:
        out[a:a + es[0], b:b + es[1], c:c + es[2]] |= mat
    return out


def _elem_touch(nodes):
    """Elements with at least one corner node in `nodes` (node box bool)."""
    es = tuple(n - 1 for n in nodes.shape)
    out = np.zeros(es, bool)
    for a, b, c in _SHIFTS:
        out |= nodes[a:a + es[0], b:b + es[1], c:c + es[2]]
    return out


class DesignBCs:
    """Material-aware BCs for re-scoring one model (THEORY §4.3 node rule restricted to nodes that touch material).

    A node touches material when one of its ≤ 8 adjacent elements has ρ ≥ `thr`. Material components (26-connected
    elements with ρ ≥ thr) that contain no element at a clamped material node are dropped (ρ = 0, P4(i): islands held
    only by ρ_min stall the PCG). Supports clamp only nodes that touch the remaining material; each load is shared
    equally among its load-face nodes that touch it (the dropped share is redistributed and reported)."""

    def __init__(self, model, ctx, thr: float = 1e-3):
        from . import fea
        self.model, self.thr = model, thr
        self.fixed0, self.load0 = model.fixed, model.load
        self.fixed_n = model.fixed.any(-1)
        self.loads = [(fea.nodes_near_faces(model.origin, model.h_e, model.node_shape, ctx.V, ctx.T, ctx.TF,
                                            L["faces"]), np.asarray(L["force"], float)) for L in ctx.loads]
        cap = ctx.cfg["capacity"]
        self.r_ex = int(np.ceil(cap["exclude_bc_mm"] / model.h_e))
        self.mount_ok = None
        if not cap["include_mounts"]:
            if model.geometry_device == "gpu":
                model.disc.upload()
            self.mount_ok = model.dm.elem_mean(ctx.mount.astype(float)).reshape(model.elem_shape) < 0.5
        self.mask0, self.ksig0 = model.stress_mask_dense, model.ksig_mask_dense
        # same rule as PhaseModel._stress_mask (elem_mean on the GPU sums with atomics: a tie at 0.5 could differ)
        self.mask_mismatch = int(np.count_nonzero(self._stress_mask(self.fixed0, self.load0) != self.mask0))
        self.current = "original"

    def _stress_mask(self, fixed, load):
        from scipy import ndimage
        bc = fixed.any(-1) | (load != 0).any(-1)
        m = ~ndimage.binary_dilation(bc, iterations=self.r_ex)[:-1, :-1, :-1]
        if self.mount_ok is not None:
            m &= self.mount_ok
        return m.ravel()

    def adjust(self, rho: np.ndarray):
        """(ρ without islands, fixed, load, info) for the design ρ (element box)."""
        from scipy import ndimage
        m = self.model
        es = m.elem_shape
        mat = (rho >= self.thr).reshape(es)
        fx = self.fixed_n & _node_touch(mat)
        lab, nlab = ndimage.label(mat, structure=np.ones((3, 3, 3), bool))
        keep = np.zeros(nlab + 1, bool)
        keep[np.unique(lab[_elem_touch(fx) & mat])] = True
        keep[0] = False
        island = mat & ~keep[lab]
        del lab
        rho2 = rho
        info = dict(components=int(nlab), islands=int(nlab - keep[1:].sum()),
                    island_mm3=float(rho[island.ravel()].sum()) * m.h_e ** 3)
        if island.any():
            rho2 = rho.copy()
            rho2[island.ravel()] = 0.0
            mat &= ~island
        touch = _node_touch(mat).ravel()
        fixed = (self.fixed_n.ravel() & touch)
        info["fixed_nodes"], info["fixed_dropped"] = int(fixed.sum()), int(self.fixed_n.sum() - fixed.sum())
        load = np.zeros((touch.size, 3))
        Fsum = sum(float(np.linalg.norm(F)) for _, F in self.loads)
        dropped = 0.0
        nd = []
        for ln, F in self.loads:
            k = ln[touch[ln]]
            if k.size == 0:
                raise RuntimeError("re-scoring: no load-face node touches material")
            load[k] += F / k.size
            dropped += float(np.linalg.norm(F)) * (1.0 - k.size / ln.size)
            nd.append(int(ln.size - k.size))
        info["load_nodes_dropped"] = nd
        info["load_share_redistributed"] = dropped / Fsum if Fsum > 0 else 0.0
        fixed = np.repeat(fixed[:, None], 3, 1).reshape(self.fixed0.shape)
        load = load.reshape(self.load0.shape)
        # check: every loaded node touches material (all 8 adjacent ρ < thr counts as isolated)
        info["isolated_load_nodes"] = int(((load != 0).any(-1).ravel() & ~touch).sum())
        return rho2, fixed, load, info

    def apply(self, fixed, load) -> None:
        """Set the model's BCs (and the stress / K_σ masks that depend on them); the FE engine is rebuilt lazily."""
        m = self.model
        same = np.array_equal(fixed, self.fixed0) and np.array_equal(load, self.load0)
        key = "original" if same else id(load)
        if key == self.current:
            return
        m.fixed, m.load = (self.fixed0, self.load0) if same else (fixed, load)
        m.stress_mask_dense = self.mask0 if same else self._stress_mask(fixed, load)
        m.ksig_mask_dense = self.ksig0 if same else m._ksig_mask()
        if m.fe is not None:
            m.fe.free()
            m.fe = None
        m.grid_fe = None
        self.current = key


def raw_responses(model, rho: np.ndarray, st, keep_fields: bool = False) -> dict:
    """Responses without normalisation (c = 1, no base) through `objective.evaluate` (same FE as the optimiser)."""
    from .objective import evaluate
    t0 = time.time()
    rho = _h(rho).astype(np.float64)
    model.ensure_active(rho)
    g = model.grid_fe
    r = g.gather_elems(rho)
    rho_hat = model.disc.rho_min + (1 - model.disc.rho_min) * r
    m = g.gather_elems(model.stress_mask_dense).astype(float)
    ev = evaluate(model.fe, rho_hat, r, m, st, None, 1.0, None)
    sp, bk, pr = ev.parts["stress"], ev.parts["buckling"], ev.parts["primal"]
    mu = np.asarray(_h(bk.mu), float) if bk is not None else np.zeros(0)
    out = dict(C=float(ev.C), sigma_pn=float(ev.sigma_pn), sigma_q=float(sp.quantile), mu_tilde=float(ev.mu_tilde),
               lam1=float(1.0 / mu[0]) if mu.size and mu[0] > 0 else float("inf"), dof=int(g.ndof),
               seconds=time.time() - t0)
    if keep_fields:
        vm = np.zeros(model.disc.n_elem)
        vm[g.elem_dense] = _h(sp.vm)
        W = np.zeros(model.disc.n_elem)
        W[g.elem_dense] = _h(pr.U_e) / (_h(rho_hat) * model.h_e ** 3)
        out.update(vm_dense=vm, W_dense=W, mu=mu,
                   modes=None if bk is None else _h(bk.modes)[:, :3], node_dense=np.asarray(g.node_dense))
    return out


def score(raw: dict, raw0: dict, st, sigma_f: float | None = None, kappa: float | None = None) -> dict:
    """J, K/K₀, P/P₀ of `raw` against the solid `raw0` on the same grid (THEORY §5.1, §4.4(d)); c = Q_q/σ_PN of the
    solid; λ_s ∝ σ_f and λ_b ∝ κ_b, so other σ_f, κ_b need no new FE."""
    from .objective import capacity
    sf = st.sigma_f if sigma_f is None else sigma_f
    kb = st.knockdown if kappa is None else kappa
    c = raw0["sigma_q"] / raw0["sigma_pn"] if np.isfinite(raw0["sigma_q"]) and raw0["sigma_pn"] > 0 else 1.0

    def cap(r):
        lam_s = sf / (c * r["sigma_pn"])
        lam, w = capacity(lam_s, r["mu_tilde"], kb, st.beta)
        gov = "strength" if (r["mu_tilde"] == 0 or w[0] >= w[1]) else "buckling"
        return lam_s, (kb / r["mu_tilde"] if r["mu_tilde"] > 0 else float("inf")), lam, gov

    ls, lb, lam, gov = cap(raw)
    ls0, lb0, lam0, gov0 = cap(raw0)
    K, P = raw0["C"] / raw["C"], lam / lam0
    wsum = st.w_K + st.w_P
    J = (st.w_K * math.log(K) + st.w_P * math.log(P)) / wsum
    return dict(J=J, K_K0=K, P_P0=P, lam_s=ls, lam_b=lb, lam=lam, governing=gov)


def decide(checks: list[dict], crit: dict, export_ok: bool = True, blocking_miss: bool = False,
           extra_reasons: list[str] | None = None, export_why: str | None = None):
    """THEORY §5.4. checks: [{res, J, K_K0, P_P0}]. Returns (accepted, reason, checks with pass flags). A failed
    export (D2) is the reason by itself: the design is then not re-scored, so `checks` is empty."""
    lnG = math.log(1 + crit["g_min"])
    why = list(extra_reasons or [])
    out = []
    for c in checks:
        pj, pk, pp = c["J"] > lnG, c["K_K0"] >= crit["f_K"], c["P_P0"] >= crit["f_P"]
        out.append(dict(c, pass_J=bool(pj), pass_K=bool(pk), pass_P=bool(pp)))
        if not pj:
            why.append(f"{c['res']}: J = {c['J']:+.4f} ≤ ln(1 + {crit['g_min']:g}) = {lnG:.4f}")
        if not pk:
            why.append(f"{c['res']}: K/K0 = {c['K_K0']:.3f} < {crit['f_K']:g}")
        if not pp:
            why.append(f"{c['res']}: P/P0 = {c['P_P0']:.3f} < {crit['f_P']:g}")
    if not export_ok:
        why.insert(0, "export failed: " + (export_why or "mass correction / watertight / R1 checks"))
    elif not checks:
        why.append("no feasible cshell design")
    if blocking_miss:
        why.append("body-fitted check misses a ratio tolerance (blocking)")
    acc = not why
    if acc:
        o = min(out, key=lambda c: c["J"])
        reason = (f"cshell design accepted: J = {o['J']:+.4f} > ln(1 + {crit['g_min']:g}) with K/K0 ≥ "
                  f"{min(c['K_K0'] for c in out):.3f}, P/P0 ≥ {min(c['P_P0'] for c in out):.3f} at every resolution")
    else:
        reason = "original solid kept: " + "; ".join(why)
    return acc, reason, out


# =====================================================================================================================
# Small writers
# =====================================================================================================================
def write_vtp(path, V: np.ndarray, T: np.ndarray, point_data: dict) -> None:
    """ASCII VTK XML PolyData (ParaView); bool arrays as UInt8."""
    def arr(a):
        return " ".join(f"{x:.7g}" for x in np.asarray(a, float).ravel())
    with open(path, "w") as f:
        f.write('<?xml version="1.0"?>\n<VTKFile type="PolyData" version="0.1" byte_order="LittleEndian">\n<PolyData>\n')
        f.write(f'<Piece NumberOfPoints="{len(V)}" NumberOfPolys="{len(T)}">\n<PointData>\n')
        for k, v in point_data.items():
            v = np.asarray(v)
            typ = "UInt8" if v.dtype == bool else "Float64"
            f.write(f'<DataArray type="{typ}" Name="{k}" format="ascii">\n')
            f.write((" ".join(str(int(x)) for x in v) if v.dtype == bool else arr(v)) + "\n</DataArray>\n")
        f.write('</PointData>\n<Points>\n<DataArray type="Float64" NumberOfComponents="3" format="ascii">\n')
        f.write(arr(V) + "\n</DataArray>\n</Points>\n<Polys>\n")
        f.write('<DataArray type="Int64" Name="connectivity" format="ascii">\n'
                + " ".join(map(str, np.asarray(T, np.int64).ravel())) + "\n</DataArray>\n")
        f.write('<DataArray type="Int64" Name="offsets" format="ascii">\n'
                + " ".join(map(str, 3 * np.arange(1, len(T) + 1))) + "\n</DataArray>\n</Polys>\n</Piece>\n</PolyData>\n</VTKFile>\n")


def read_vtp_point_data(path) -> dict:
    """Inverse of `write_vtp` for its own files (tests, report)."""
    import re
    txt = open(path).read()
    out = {}
    for m in re.finditer(r'<DataArray type="(\w+)" Name="(\w+)" format="ascii">\n(.*?)\n</DataArray>', txt, re.S):
        if m.group(2) in ("connectivity", "offsets"):
            continue
        out[m.group(2)] = np.array(m.group(3).split(), float)
    return out


def _write_csv(path, rows: list[dict], cols=None):
    cols = cols or (list(rows[0].keys()) if rows else [])
    with open(path, "w", newline="") as f:
        w = csv.DictWriter(f, cols, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)


def _jsonable(o):
    if isinstance(o, dict):
        return {str(k): _jsonable(v) for k, v in o.items()}
    if isinstance(o, (list, tuple)):
        return [_jsonable(v) for v in o]
    if isinstance(o, (np.bool_,)):
        return bool(o)
    if isinstance(o, (np.integer,)):
        return int(o)
    if isinstance(o, (float, np.floating)):
        return float(o) if np.isfinite(o) else None
    if isinstance(o, np.ndarray):
        return _jsonable(o.tolist())
    return o


class DenseLookup:
    """Element field on the FE element box, sampled at points (element containing the point; 0 outside)."""

    def __init__(self, model, values):
        self.o, self.h, self.shape = np.asarray(model.origin, float), model.h_e, model.elem_shape
        self.v = np.asarray(values)

    def __call__(self, X):
        ijk = np.floor((np.asarray(X, float) - self.o) / self.h).astype(np.int64)
        ok = np.all((ijk >= 0) & (ijk < np.array(self.shape)), 1)
        ijk = np.clip(ijk, 0, np.array(self.shape) - 1)
        f = (ijk[:, 0] * self.shape[1] + ijk[:, 1]) * self.shape[2] + ijk[:, 2]
        return np.where(ok, self.v[f], 0.0)


def mode_magnitudes(model, raw: dict, X: np.ndarray) -> np.ndarray:
    """|mode| (first three) at the FE node nearest each point, max-normalised; NaN where the node is not active."""
    if raw.get("modes") is None:
        return np.full((3, len(X)), np.nan)
    ns = model.node_shape
    ijk = np.clip(np.rint((X - model.origin) / model.h_e).astype(np.int64), 0, np.array(ns) - 1)
    f = (ijk[:, 0] * ns[1] + ijk[:, 1]) * ns[2] + ijk[:, 2]
    nd = raw["node_dense"]
    pos = np.clip(np.searchsorted(nd, f), 0, len(nd) - 1)
    hit = nd[pos] == f
    out = np.full((3, len(X)), np.nan)
    M = raw["modes"]
    for j in range(min(3, M.shape[1])):
        u = M[:, j].reshape(-1, 3)
        mag = np.linalg.norm(u, axis=1)
        mag = mag / max(mag.max(), 1e-300)
        out[j, hit] = mag[pos[hit]]
    return out



def decimate_for_validation(V: np.ndarray, T: np.ndarray, edge: float, vol_tol: float = 1e-3, log=print):
    """Quadric edge-collapse (pymeshlab, topology preserved) of a marching-cubes surface to edge length ≈ `edge` for
    the body-fitted mesher (§8.5.1: decimation only with a volume check). Returns (V, T, info); the input is
    returned unchanged when it is already coarse or when the check fails (logged)."""
    import trimesh
    tm = trimesh.Trimesh(V, T, process=False)
    A, vol0 = float(tm.area), float(tm.volume)
    target = int(A / (0.433 * edge ** 2))
    info = dict(faces_in=int(len(T)), target=target, volume_in=vol0)
    if target >= len(T):
        return V, T, dict(info, decimated=False)
    import pymeshlab
    ms = pymeshlab.MeshSet()
    ms.add_mesh(pymeshlab.Mesh(np.asarray(V, float), np.asarray(T, np.int32)))
    ms.meshing_decimation_quadric_edge_collapse(targetfacenum=target, preservenormal=True, preservetopology=True,
                                                preserveboundary=True, qualitythr=0.5, planarquadric=True)
    m = ms.current_mesh()
    V2, T2 = m.vertex_matrix().astype(float), m.face_matrix().astype(np.int64)
    t2 = trimesh.Trimesh(V2, T2, process=False)
    rel = float(t2.volume) / vol0 - 1.0
    ok = bool(t2.is_watertight) and abs(rel) <= vol_tol
    info.update(faces_out=int(len(T2)), volume_rel=rel, watertight=bool(t2.is_watertight), decimated=ok)
    log(f"[collect] validation surface: {len(T)} -> {len(T2)} faces, V rel {rel:+.2e}, watertight "
        f"{t2.is_watertight}{'' if ok else ' -> check FAILED, using the full surface'}")
    return (V2, T2, info) if ok else (V, T, info)

# =====================================================================================================================
# collect
# =====================================================================================================================
def export_verdict(ex: dict | None, vol_tol: float = 2e-3, G_s: float | None = None):
    """(ok, reason) of an export summary (PLAN_collect_fixes P3, D2): mass correction bracketed and within δ_fail,
    watertight, |V_STL/V₀ − 1| ≤ vol_tol. A failed bracket after a large void_min opening fill is reported as the
    optimiser's non-convergence (§8.2/§8.4 behave as specified), not as a re-scoring error."""
    if ex is None:
        return True, ""
    why = []
    mc, mesh, cl = ex.get("mass_correction", {}), ex.get("mesh", {}), ex.get("cleanup", {}) or {}
    V0 = ex.get("V0_mm3") or float("nan")
    if not mc.get("ok", False):
        bf = mc.get("bracket_failed") or {}
        why.append(f"mass correction did not bracket (V/V0-1 before correction {mc.get('rel_before', float('nan')):+.2%}"
                   + (f", limit ±{bf['limit']:g} mm" if "limit" in bf else "") + ")")
        fill = cl.get("opening_removed_mm3")
        if fill and fill / V0 > 0.01:
            why.append(f"the void_min opening filled {fill / V0:+.1%} of V0: the design is not converged in the "
                       f"sliver measure" + (f" (final G_s = {G_s:.3g})" if G_s is not None and np.isfinite(G_s) else "")
                       + " - optimiser non-convergence, not a re-scoring error")
    elif mc.get("fail", False):
        why.append(f"mass correction |delta| = {abs(mc.get('delta', float('nan'))):.3f} mm above delta_fail")
    if not mesh.get("watertight_manifold", False):
        why.append(f"mesh not watertight ({mesh.get('bodies')} bodies, {mesh.get('duplicate_directed_edges')} "
                   "duplicated directed edges)")
    rel = mesh.get("volume_rel")
    if rel is None or not np.isfinite(rel) or abs(rel) > vol_tol:
        why.append(f"V_STL/V0 - 1 = {rel if rel is not None else float('nan'):+.2%} (tolerance {vol_tol:.1%})")
    ok = bool(ex.get("ok", False)) and not why
    if not ok and not why:
        why.append("export checks failed")
    return ok, "; ".join(why)


def collect(run_dir: str, allow_incomplete: bool = False, validate: bool | None = None, export_kw: dict | None = None,
            log=None, geometry_device: str | None = None) -> dict:
    """Merge the starts of `run_dir`, export, re-score, validate and decide (module docstring). Returns summary.
    geometry_device: "cpu" | "gpu" | "auto" for the re-scoring cell maps (default: config grid.geometry_device)."""
    from . import export as cex
    from .optimise import Context, PeakGPU
    run_dir = os.path.abspath(run_dir)
    logf = open(os.path.join(run_dir, "collect_log.txt"), "a")

    def _log(msg):
        print(msg, flush=True)
        logf.write(msg + "\n"); logf.flush()
    log = log or _log
    T0 = time.time()
    tim, mem = {}, {}
    pk_gpu = PeakGPU()
    try:
        pk_gpu.start()
    except Exception:
        pk_gpu = None

    def memo(stage):
        mem[stage] = round(rss_gib(), 2)
        log(f"[collect] host RSS {mem[stage]:.1f} GiB after {stage}")
    cfg = json.load(open(os.path.join(run_dir, "config_used.json")))
    from . import config as ccfg
    cfg = ccfg.load(cfg)                                  # re-apply defaults (keys added after the run)
    if geometry_device is not None:
        cfg["grid"]["geometry_device"] = geometry_device
    o, m = cfg["objective"], cfg["manufacturing"]
    crit = dict(g_min=o["min_gain"], ln_1_g_min=math.log(1 + o["min_gain"]), f_K=o["floor_stiffness"],
                f_P=o["floor_capacity"])
    files = {}
    ekw = dict(export_kw or {})

    # 1. starts ----------------------------------------------------------------------------------------------------
    t0 = time.time()
    rows = start_table(run_dir, cfg)
    best = pick_best(rows)
    if best is None and allow_incomplete:
        cand = [r for r in rows if r["status"] != "missing" and np.isfinite(r["J_opt"])]
        if cand:
            best = max(cand, key=lambda r: r["J_opt"])
            log(f"[collect] WARNING no complete feasible start; --allow-incomplete uses the latest checkpoint of "
                f"{best['start']} (status {best['status']})")
    tim["starts"] = time.time() - t0
    log("[collect] starts: " + "; ".join(f"{r['start']} {r['status']} J {r['J_opt']:+.4f} feasible {r['feasible']}"
                                         for r in rows))
    t0 = time.time()
    ctx = Context(cfg, log=log)
    tim["context"] = time.time() - t0
    memo("context")
    nv = ctx.surf.n_vertices
    xi = None
    if best is not None:
        if best["xi_file"]:
            xi = np.load(os.path.join(run_dir, best["xi_file"]))
        else:
            xi = np.asarray(latest_checkpoint(run_dir, best["start"])["xi"])
        if len(xi) != 2 * nv:
            raise ValueError(f"xi of {best['start']} has {len(xi)} entries, Γ_h has {nv} vertices (2N expected)")
    D_comp = best["D_comp"] if best else cfg["design"]["outward_extent_initial_mm"]
    h_g_last = float(cfg["grid"]["phases"][-1]["h_g"])
    for r in rows:                                        # wall range per start (vertex fields, final-phase map)
        r["w_min"] = r["w_max"] = r["closed_area_frac"] = None
        p = os.path.join(run_dir, r["xi_file"]) if r["xi_file"] else None
        if p and os.path.exists(p):
            f = cex.vertex_fields(ctx, np.load(p), r["D_comp"], h_g_last)
            act = ~f["passive"]
            op = act & (f["a"] < ctx.r_ma)
            w = (f["a"] + f["d"])[op]
            r["w_min"], r["w_max"] = (float(w.min()), float(w.max())) if w.size else (None, None)
            A = ctx.surf.A * act
            r["closed_area_frac"] = float((A * (f["a"] >= ctx.r_ma)).sum() / A.sum())
    b_v, b_o = m["wall_bias_mm"], m["outer_bias_mm"]
    vf = cex.vertex_fields(ctx, xi, D_comp, h_g_last) if xi is not None else None
    f_best = None
    if vf is not None:
        f_best = design_fields(vf["d"], vf["a"], vf["xt_a"] >= cex.ExportConfig.closed_thr, ctx.r_ma, b_v)

    # 2. parametric FE of the best design on the optimisation grid; uniform-shell t on its band ----------------------
    res_list = resolutions(cfg)
    t0 = time.time()
    model = make_model(ctx, res_list[0], D_comp, log)
    memo("model opt (parametric)")
    st = scoring_settings(model)
    V0_band = float(_h(model.V0))
    rule, rho_p = {}, None
    if xi is not None:                                    # geometry first (band on the device), then the FE
        dm = model.dm
        x = dm.xp.asarray(xi)
        V_opt = dm.volume(dm.forward(x[:nv], x[nv:], as_built=False))
        _, V_rule = cell_rescore(model, None, f_best, 0.0, want_rho=False)
        rule = dict(V_opt_rel=V_opt / V0_band - 1.0, V_rule_rel=V_rule / V0_band - 1.0, rule_vs_opt=V_rule / V_opt - 1.0)
        log(f"[collect] best ({best['start']}) band volume: optimiser V/V0-1 {rule['V_opt_rel']:+.3e}, export closure "
            f"rule {rule['V_rule_rel']:+.3e} (rule/optimiser - 1 = {rule['rule_vs_opt']:+.3e})")
        rho_p = _h(dm.rho(dm.forward(x[:nv], x[nv:], as_built=True))).astype(np.float64)
    # uniform shell (§7.4 S1, D1: mount faces solid): t with V_nominal(t) = V0 on this band
    has_mount = bool(np.asarray(ctx.mount).any())

    def shell_dV(t):
        fu = cex.uniform_shell_fields(ctx, t, h_g_last)
        return cell_rescore(model, None, design_fields(fu["d"], fu["a"], fu["closed"], ctx.r_ma, b_v), 0.0,
                            want_rho=False)[1] / V0_band - 1.0
    from scipy.optimize import brentq
    t_band = None
    try:
        t_band = brentq(shell_dV, 0.02, D_comp, xtol=1e-5)
        log(f"[collect] uniform shell{' (mounts solid)' if has_mount else ''}: t = {t_band:.4f} mm for constant mass "
            "on the optimisation band")
    except ValueError as e:
        log(f"[collect] uniform shell: no t in [0.02, {D_comp}] mm restores the mass on the band ({e})")
        if has_mount:
            raise RuntimeError(f"uniform shell: no t ≤ D_comp = {D_comp} mm restores the solid's mass") from e
    if model.geometry_device == "gpu":
        model.disc.release()                             # the FE gets the device memory (as fe_eval does)
    raw_solid0 = raw_responses(model, model.rho0, st, keep_fields=True)
    raw_param = None
    if rho_p is not None:
        raw_param = raw_responses(model, rho_p, st, keep_fields=True)
        rho_p = None
        log(f"[collect] parametric best ({best['start']}) on the optimisation grid: "
            f"{score(raw_param, raw_solid0, st)}")
    g_opt = _grid_info(model)
    tim["param_fe"] = time.time() - t0
    free_model(model)
    model = None
    memo("release model opt (parametric)")

    # 3. export (the PhaseModel is released: the forked export workers do not inherit it) ----------------------------
    t0 = time.time()
    bdir = os.path.join(run_dir, "export_baselines")
    base_ex = cex.export_baselines(ctx, bdir, t_coarse=t_band, h_g=h_g_last, log=log, **ekw)
    for nm in ("original", "uniform_shell"):
        shutil.copyfile(os.path.join(bdir, f"{nm}.stl"), os.path.join(run_dir, f"{nm}.stl"))
        files[f"{nm}.stl"] = f"{nm}.stl"
    t_shell = float(base_ex["uniform_shell"]["t_final_mm"])
    if t_shell > D_comp:
        raise RuntimeError(f"uniform shell t = {t_shell:.2f} mm exceeds the band's outward extent {D_comp} mm")
    tim["export_baselines"] = time.time() - t0
    trim_host_memory()
    memo("export baselines")
    best_ex, export_ok, export_why = None, True, ""
    if xi is not None:
        t0 = time.time()
        stress_fn = DenseLookup(g_opt, raw_param["vm_dense"])
        best_ex = cex.export_design(ctx, xi, run_dir, D_comp, stress=stress_fn, name="best_candidate", h_g=h_g_last,
                                    log=log, **ekw)
        files["best_candidate.stl"] = "best_candidate.stl"
        tim["export_best"] = time.time() - t0
        trim_host_memory()
        memo("export best")
        hist_b = [h_ for h_ in read_history(run_dir) if h_["start"] == best["start"]]
        G_s = _f(hist_b[-1].get("G_s")) if hist_b else None
        export_ok, export_why = export_verdict(best_ex, G_s=G_s)
        if not export_ok:
            log(f"[collect] best_candidate export FAILED its checks: {export_why}. The STL is kept for inspection and "
                "not re-scored (D2); the original solid is kept.")

    # 4. re-score ----------------------------------------------------------------------------------------------------
    stl = {"uniform_shell": os.path.join(run_dir, "uniform_shell.stl")}
    fields = {"uniform_shell": None}
    fu = cex.uniform_shell_fields(ctx, t_shell, h_g_last)
    fields["uniform_shell"] = design_fields(fu["d"], fu["a"], fu["closed"], ctx.r_ma, b_v)
    if best_ex is not None and export_ok:
        stl["best"] = os.path.join(run_dir, "best_candidate.stl")
        fields["best"] = f_best
    c0 = np.round(0.5 * (ctx.V.min(0) + ctx.V.max(0)), 3)            # band.c0 (Warp frame)
    res_rows, raws = [], {}
    fields_extra, modes = {}, {}
    for r in res_list:
        t0 = time.time()
        model = make_model(ctx, r, D_comp, log)
        if not np.allclose(model.band.c0, c0):
            raise RuntimeError("band centring offset differs from the mesh frame used for the STL soup")
        memo(f"model {r['res']}")
        st_r = scoring_settings(model)
        bcs = DesignBCs(model, ctx)
        V0 = float(_h(model.V0))
        rho_chk, _ = cell_rescore(model, None, None, b_o, solid=True)
        if not np.allclose(rho_chk, _h(model.rho0), rtol=0, atol=1e-12):
            raise RuntimeError("re-scoring ρ of the solid differs from the optimiser's baseline ρ0")
        del rho_chk
        rhos = {"solid": _h(model.rho0).astype(np.float64)}
        vols = {"solid": V0}
        for name, path in stl.items():
            soup = stl_soup(path, c0)
            phiE = mesh_phi(model, soup)
            del soup
            rhos[name], vols[name] = cell_rescore(model, phiE, fields[name], b_o)
            del phiE
            trim_host_memory()
        if model.geometry_device == "gpu":
            model.disc.release()                         # the FE gets the device memory (as fe_eval does)
        designs, bcinfo = {}, {}
        for name, rho in rhos.items():
            rho2, fixed, load, info = bcs.adjust(rho)
            bcinfo[name] = info
            if info["islands"] or info["fixed_dropped"] or any(info["load_nodes_dropped"]):
                log(f"[collect] {r['res']}: {name}: BCs on material: {info}")
            bcs.apply(fixed, load)
            keep = r["res"] == "opt" and name in ("solid", "best")
            try:
                designs[name] = raw_responses(model, rho2, st_r, keep_fields=keep)
            except RuntimeError as e:                        # e.g. PCG not converged: recorded, the row fails
                log(f"[collect] {r['res']}: {name}: FE FAILED: {e!r}")
                designs[name] = dict(C=float("nan"), sigma_pn=float("nan"), sigma_q=float("nan"),
                                     mu_tilde=float("nan"), lam1=float("nan"), dof=0, seconds=0.0, error=repr(e)[:300])
            fe = model.fe
            designs[name]["pcg"] = list(getattr(fe, "stats_solves", []) or [])[-3:] if fe is not None else []
            del rho2
        bcs.apply(bcs.fixed0, bcs.load0)
        raw0 = designs["solid"]
        if r["res"] == "opt":
            log(f"[collect] solid re-score vs the parametric pass on the same grid: C {raw0['C'] / raw_solid0['C'] - 1:+.2e}")
        for name, raw in designs.items():
            sc = score(raw, raw0, st_r) if np.isfinite(raw["C"]) else dict(
                J=float("nan"), K_K0=float("nan"), P_P0=float("nan"), lam_s=float("nan"), lam_b=float("nan"),
                lam=float("nan"), governing=None)
            bi = bcinfo[name]
            row = dict(label=f"{r['res']}: {name}", res=r["res"], design=name, h_e=r["h_e"], h_g=r["h_g"],
                       shift_mm=r["shift"], **sc, C_Nmm=raw["C"], K_N_mm=float(ctx.F @ ctx.F / raw["C"]),
                       lam1=raw["lam1"], sigma_pn=raw["sigma_pn"], sigma_q=raw["sigma_q"], V_mm3=vols[name],
                       mass_err=vols[name] / V0 - 1.0, dof=raw["dof"], seconds=raw["seconds"],
                       error=raw.get("error"), pcg=raw.get("pcg"),
                       bc_load_share_redistributed=bi["load_share_redistributed"],
                       bc_load_nodes_dropped=bi["load_nodes_dropped"], bc_fixed_dropped=bi["fixed_dropped"],
                       bc_isolated_load_nodes=bi["isolated_load_nodes"], islands=bi["islands"],
                       island_mm3=bi["island_mm3"])
            res_rows.append(row)
            log(f"[collect] {row['label']:<24} J {sc['J']:+.4f} K/K0 {sc['K_K0']:.4f} P/P0 {sc['P_P0']:.4f} "
                f"{sc['governing']} V/V0-1 {row['mass_err']:+.2e} load moved {bi['load_share_redistributed']:.1%}")
        raws[r["res"]] = (designs, st_r)
        if r["res"] == "opt":
            surfX = ctx.surf.V
            Ns = ctx.surf.projected_normals() if hasattr(ctx.surf, "projected_normals") else ctx.surf.N
            dd = vf["d"] if vf is not None else np.zeros(nv)
            pts_d = surfX + (dd - 0.5 * model.h_e)[:, None] * Ns
            pts_s = surfX - 0.5 * model.h_e * Ns
            fields_extra = dict(stress_solid=DenseLookup(model, raw0["vm_dense"])(pts_s))
            # design fields: the re-scored export when it passed its checks, else the parametric best (same grid)
            rb, gb = (designs["best"], model) if "best" in designs and "vm_dense" in designs["best"] else (raw_param, g_opt)
            if rb is not None:
                fields_extra["stress"] = DenseLookup(gb, rb["vm_dense"])(pts_d)
                fields_extra["W_bar"] = DenseLookup(gb, rb["W_dense"])(pts_d)
            modes = dict(solid_mag=mode_magnitudes(model, raw0, pts_s), solid_lambda=1.0 / raw0["mu"][raw0["mu"] > 0]
                         if raw0.get("mu") is not None and raw0["mu"].size else np.zeros(0))
            if rb is not None and rb.get("mu") is not None:
                modes.update(design_mag=mode_magnitudes(gb, rb, pts_d),
                             design_lambda=1.0 / rb["mu"][rb["mu"] > 0] if rb["mu"].size else np.zeros(0))
            for dsg in designs.values():                  # drop big arrays
                for k in ("vm_dense", "W_dense", "modes", "node_dense"):
                    dsg.pop(k, None)
        tim[f"rescore_{r['res']}"] = time.time() - t0
        rb = gb = None
        free_model(model)
        model = bcs = None
        trim_host_memory()
        memo(f"release model {r['res']}")
    raw_param = raw_solid0 = None

    # 5. body-fitted validation ---------------------------------------------------------------------------------
    from . import validate as cval
    do_val = cfg["validation"].get("body_fitted", True) if validate is None else validate
    val = dict(ran=False, why="disabled" if not do_val else ("export failed" if not export_ok else "no design"))
    blocking = val_failed = False
    if do_val and "best" in stl:
        try:
            t0 = time.time()
            vd = os.path.join(run_dir, "validation")
            h_val = _get(cfg, "validation.h_mm", 0.5)       # coarse level; the finer is h/sqrt(2) (>= 2 tets per wall)
            kw = dict(h=h_val, h_bc=res_list[0]["h_e"], n_modes=4, nproc=int(_get(cfg, "validation.nproc", min(
                32, len(os.sched_getaffinity(0))))), want_stress=True, E=cfg["material"]["E_MPa"], nu=cfg["material"]["nu"],
                fallback=None, timeout=_get(cfg, "validation.timeout_s", None))
            bf, dec, errs, used = {}, {}, {}, {}
            VE, TE, _ = load_mesh(stl["best"])
            # Same mesher for solid and design (ratios cancel mesher bias): isotropic surface remesh to h, gmsh volume
            # inside that triangulation, quality gate; solved on the GPU (cuDSS, C3D10-identical). Two levels
            # (h, h/sqrt(2)) give the mesh-convergence estimate; the finer level is compared. No tetgen fallback:
            # its sliver mesh moved lambda_b by -46 % on B2 (research/b2_bodyfit).
            kw.update(engine=_get(cfg, "validation.engine", "gpu"))
            for nm, (Vs, Ts) in (("solid", (ctx.V, ctx.T)), ("design", (VE, TE))):
                levels = [h_val, h_val / 2 ** 0.5]
                dec[nm] = dict(remesh="isotropic", levels=levels)
                try:
                    lv = []
                    for i_l, hh in enumerate(levels):
                        # the coarse level is reported (gate numbers included) but only the finer one must pass the gate;
                        # it only feeds the (non-blocking) convergence flag, so its failure is recorded, not blocking:
                        # pymeshlab's isotropic remesh at h = 0.5 passes or fails the edge check depending on the vertex
                        # order of the same B2 surface (research/collect_fix/remesh_check.py, job 7870945)
                        fine = i_l == len(levels) - 1
                        try:
                            lv.append(cval.validate_design(Vs, Ts, ctx.V, ctx.T, ctx.TF, ctx.support, ctx.loads, vd,
                                                           f"{nm}_h{hh:.3f}", backend="keep", enforce_gate=fine,
                                                           **dict(kw, h=hh)))
                        except Exception as e:
                            if fine:
                                raise
                            errs[f"{nm}/coarse_level"] = repr(e)[:600]
                            log(f"[collect] WARNING body-fitted {nm}: coarse level h = {hh:.3f} failed ({e!r:.300}); "
                                "no level-convergence estimate, the finer level is compared")
                    bf[nm] = lv[-1]
                    if len(lv) == 2:
                        dec[nm]["level_change"] = dict(C=lv[1].compliance / lv[0].compliance - 1,
                                                       lam=lv[1].lambda_b / lv[0].lambda_b - 1)
                        dec[nm]["converged_C"] = bool(abs(dec[nm]["level_change"]["C"]) <= 0.02)
                    else:
                        dec[nm]["level_change"], dec[nm]["converged_C"] = None, None
                    dec[nm]["gate"] = [x.info.get("mesh_quality") for x in lv]
                    dec[nm]["lambda_unexcl"] = [x.info.get("lam_unexcl") for x in lv]
                    dec[nm]["seconds"] = [dict(mesh=x.seconds_mesh, solve=x.seconds_ccx) for x in lv]
                    used[nm] = "keep"
                    if dec[nm]["converged_C"] is False:
                        errs[f"{nm}/convergence"] = f"compliance changes {dec[nm]['level_change']['C']:+.2%} between levels"
                        log(f"[collect] WARNING {nm}: {errs[f'{nm}/convergence']} (> 2 %)")
                except Exception as e:                        # recorded; a failed check blocks acceptance below
                    errs[f"{nm}/keep"] = repr(e)[:600]
                    log(f"[collect] body-fitted {nm} FAILED: {e!r}")
            del VE, TE
            if len(bf) < 2:
                raise _ValidationFailed(errs)
            d_opt = raws["opt"][0]
            vox = {k: dict(compliance=d_opt[n]["C"], lambda_b=d_opt[n]["lam1"], sigma_q=d_opt[n]["sigma_q"])
                   for k, n in (("solid", "solid"), ("design", "best"))}
            bfd = {k: dict(compliance=v.compliance, lambda_b=v.lambda_b, sigma_q=v.sigma_q) for k, v in bf.items()}
            vrows = cval.compare(vox, bfd)
            cval.write_validation_csv(vrows, os.path.join(run_dir, "validation.csv"))
            files["validation.csv"] = "validation.csv"
            blocking = any(r["blocking"] and not r["passed"] for r in vrows)
            for r in vrows:
                if r.get("under"):
                    log(f"[collect] body-fitted: voxel under-predicts {r['metric']} by {-r['over']:.1%} "
                        "(conservative; logged, not blocking, §8.6 D3)")
            val = dict(ran=True, engine=sorted({v.engine for v in bf.values()}), blocking_miss=blocking, surface=dec,
                       errors=errs, mesher=used,
                       n_rows=len(vrows), n_failed=sum(not r["passed"] for r in vrows), file="validation.csv",
                       under_predicted=[r["metric"] for r in vrows if r.get("under")],
                       seconds_mesh={k: v.seconds_mesh for k, v in bf.items()},
                       seconds_ccx={k: v.seconds_ccx for k, v in bf.items()})
            tim["validation"] = time.time() - t0
            log(f"[collect] body-fitted: {val}")
        except _ValidationFailed as e:
            val_failed = True
            val = dict(ran=False, why="body-fitted check failed (mesher/solver)", errors=e.args[0], blocking_miss=True)
            tim["validation"] = time.time() - t0
            log(f"[collect] body-fitted check FAILED: {e.args[0]}")

    # 6. decision ----------------------------------------------------------------------------------------------------
    checks = [dict(res=r["res"], J=r["J"], K_K0=r["K_K0"], P_P0=r["P_P0"]) for r in res_rows if r["design"] == "best"]
    extra = ["body-fitted check could not be run (see validation.errors)"] if val_failed else []
    extra += [f"{r['res']}: FE of the best export failed ({r['error']})" for r in res_rows
              if r["design"] == "best" and r.get("error")]
    acc, reason, checks = decide(checks, crit, export_ok, blocking, extra, export_why=export_why)
    final = os.path.join(run_dir, "final_design.stl")
    shutil.copyfile(os.path.join(run_dir, "best_candidate.stl" if acc else "original.stl"), final)
    files["final_design.stl"] = "final_design.stl"
    log(f"[collect] DECISION: {reason}")

    # sensitivity ------------------------------------------------------------------------------------------------
    sens = []
    for kb in KAPPAS:
        for sp in SIGMA_PCT:
            js, ks, ps, jopt = [], [], [], float("nan")
            for rk, (designs, st_r) in raws.items():
                if "best" not in designs or not np.isfinite(designs["best"]["C"]):
                    continue
                sc = score(designs["best"], designs["solid"], st_r, st_r.sigma_f * (1 + sp / 100), kb)
                js.append(sc["J"]); ks.append(sc["K_K0"]); ps.append(sc["P_P0"])
                if rk == "opt":
                    jopt = sc["J"]
            ok = bool(js) and decide([dict(res="s", J=a, K_K0=b, P_P0=c) for a, b, c in zip(js, ks, ps)], crit,
                                     export_ok, blocking, extra)[0]
            sens.append(dict(kappa_b=kb, sigma_f=sp, sigma_f_MPa=cfg["material"]["strength_MPa"] * (1 + sp / 100),
                             J_star=jopt, J_min=min(js) if js else float("nan"), K_K0_min=min(ks) if ks else float("nan"),
                             P_P0_min=min(ps) if ps else float("nan"), accepted=int(ok)))
    _write_csv(os.path.join(run_dir, "sensitivity.csv"), sens)
    files["sensitivity.csv"] = "sensitivity.csv"

    # fields.vtp, buckling_modes.npz, bc_faces.npz -----------------------------------------------------------------
    s = ctx.surf
    pdata = {}
    if vf is not None:
        op = vf["a"] < ctx.r_ma
        pdata = dict(xi=xi[nv:], xi_d=xi[:nv], xi_a=xi[nv:], xit_a=vf["xt_a"], d=vf["d"], a=vf["a"],
                     w=np.where(op, vf["a"] + vf["d"], np.nan), r_MA=ctx.r_ma, d_hi=vf["d_hi"],
                     passive=vf["passive"].astype(bool))
    else:
        pdata = dict(r_MA=ctx.r_ma)
    pdata.update(fields_extra)
    write_vtp(os.path.join(run_dir, "fields.vtp"), s.V, s.T, pdata)
    files["fields.vtp"] = "fields.vtp"
    np.savez_compressed(os.path.join(run_dir, "buckling_modes.npz"), V=s.V.astype(np.float32), T=s.T.astype(np.int32),
                        **{k: np.asarray(v, np.float32) for k, v in modes.items()})
    files["buckling_modes.npz"] = "buckling_modes.npz"
    load_id = np.full(len(ctx.T), -1, np.int32)
    for i, L in enumerate(ctx.loads):
        load_id[np.isin(ctx.TF, L["faces"])] = i
    np.savez_compressed(os.path.join(run_dir, "bc_faces.npz"), V=ctx.V.astype(np.float32), T=ctx.T.astype(np.int32),
                        support=np.isin(ctx.TF, ctx.support), load=load_id)
    files["bc_faces.npz"] = "bc_faces.npz"

    # summary ------------------------------------------------------------------------------------------------------
    rho_g = cfg["material"]["density_g_cm3"] * 1e-3
    bmesh = (best_ex or {}).get("mesh", {})
    V0x = (best_ex or base_ex["original"]).get("V0_mm3")
    vfin = bmesh.get("volume") if acc else base_ex["original"].get("mesh", {}).get("volume")
    mass = dict(target=V0x * rho_g if V0x else None, final=vfin * rho_g if vfin else None,
                error=bmesh.get("volume_rel"), V0_mm3=V0x, density_g_mm3=rho_g,
                mass_matched=bool(export_ok) if best_ex is not None else None)
    d_opt = {r["design"]: r for r in res_rows if r["res"] == "opt"}
    comparison = [dict(label="solid", source="rescored", J=0.0, K_K0=1.0, P_P0=1.0,
                       governing=d_opt["solid"]["governing"], w_min=None, w_max=None),
                  dict(label="uniform shell", source="rescored", **{k: d_opt["uniform_shell"][k] for k in
                                                                    ("J", "K_K0", "P_P0", "governing")},
                       w_min=t_shell, w_max=t_shell)]
    vrow = None
    try:
        vs = load_vshell(vshell_run_dir(cfg))
        b = vs["best"]
        vrow = dict(run_dir=str(vshell_run_dir(cfg)), J=b.get("J"), K_K0=b.get("K/K0"), P_P0=b.get("P/P0"),
                    accepted=vs["accepted"], s_mm=vs["s"])
        comparison.append(dict(label="vshell best", source="vshell", J=b.get("J"), K_K0=b.get("K/K0"),
                               P_P0=b.get("P/P0"), governing=b.get("governing mode"),
                               w_min=b.get("as-built wall mm"), w_max=b.get("outer offset / strut wall s mm")))
    except (FileNotFoundError, KeyError) as e:
        log(f"[collect] vshell row not available: {e}")
    for r in rows:
        comparison.append(dict(label=f"cshell {r['start']}", source="optimiser", J=r["J_opt"], K_K0=r["K_K0_opt"],
                               P_P0=r["P_P0_opt"], governing=None, w_min=r["w_min"], w_max=r["w_max"]))
    if "best" in d_opt:
        bw = (best_ex or {}).get("wall", {})
        comparison.append(dict(label="cshell export", source="rescored", **{k: d_opt["best"][k] for k in
                                                                            ("J", "K_K0", "P_P0", "governing")},
                               w_min=bw.get("min_nominal"), w_max=next((r["w_max"] for r in rows if best and
                                                                        r["start"] == best["start"]), None)))
    elif best_ex is not None:
        comparison.append(dict(label="cshell export", source="not mass-matched", J=None, K_K0=None, P_P0=None,
                               governing=None, w_min=None, w_max=None))
    load_cases = [dict(case=f"1 ({n})", lam_s=d_opt[k]["lam_s"], lam_b=d_opt[k]["lam_b"], softmin=d_opt[k]["lam"])
                  for k, n in (("solid", "solid"), ("best", "cshell")) if k in d_opt]
    Lam = None
    if vf is not None and "W_bar" in pdata:
        sel = ~vf["passive"] & (vf["a"] < ctx.r_ma) & np.isfinite(pdata["W_bar"]) & (pdata["W_bar"] > 0)
        if sel.any():
            w_, a_ = pdata["W_bar"][sel], ctx.surf.A[sel]
            o_ = np.argsort(w_)
            c_ = np.cumsum(a_[o_])
            Lam = float(w_[o_][np.searchsorted(c_, 0.5 * c_[-1])])
    hist = read_history(run_dir)
    peak = max([_f(h_["gpu_mem_GiB"]) for h_ in hist if np.isfinite(_f(h_["gpu_mem_GiB"]))] or [float("nan")])
    gpu_model = None
    try:
        import cupy as cp
        gpu_model = cp.cuda.runtime.getDeviceProperties(0)["name"].decode()
    except Exception:
        pass
    tim["total"] = time.time() - T0
    import resource
    mem["peak_rss_GiB"] = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 2 ** 20
    mem["peak_rss_children_GiB"] = resource.getrusage(resource.RUSAGE_CHILDREN).ru_maxrss / 2 ** 20
    peak_collect_gpu = pk_gpu.stop() if pk_gpu is not None else float("nan")
    ex_sum = None
    if best_ex is not None:
        ex_sum = {k: v for k, v in best_ex.items() if k not in ("timing_s",)}
    S = dict(schema=SCHEMA, name=cfg["name"], accepted=acc,
             outcome="cshell design accepted" if acc else "original SOLID part kept", reason=reason, criteria=crit,
             mass=mass,
             best=None if best is None else dict(start=best["start"], J_opt=best["J_opt"], D_comp=D_comp,
                                                 xi_file=best["xi_file"], phase=len(cfg["grid"]["phases"]) - 1),
             starts=rows, comparison=comparison, resolutions=res_rows, checks=checks, load_cases=load_cases,
             Lambda=Lam,
             export=dict(best=ex_sum, best_ok=export_ok if best_ex is not None else None,
                         best_reject_reason=export_why or None, closure_rule_band=rule or None,
                         uniform_shell_t_mm=t_shell, uniform_shell_t_band_mm=t_band,
                         uniform_shell_mounts_solid=bool(base_ex["uniform_shell"].get("mounts_solid", False)),
                         delta_mm=(best_ex or {}).get("mass_correction", {}).get("delta"),
                         voids=(best_ex or {}).get("voids"), holes=((best_ex or {}).get("voids") or {}).get("holes"),
                         wall=(best_ex or {}).get("wall")),
             validation=val,
             sensitivity=dict(file="sensitivity.csv", accepted_fraction=float(np.mean([r["accepted"] for r in sens])),
                              kappa_b=list(KAPPAS), sigma_f_pct=list(SIGMA_PCT)),
             vshell=vrow,
             resources=dict(gpu_model=gpu_model, peak_gpu_GiB=peak, peak_gpu_collect_GiB=peak_collect_gpu,
                            host_rss_GiB=mem,
                            gpu_hours=float(np.nansum([r["hours"] for r in rows])) + tim["total"] / 3600,
                            collect_s=tim["total"]),
             timings_s=tim, files=files)
    S["files"]["summary.json"] = "summary.json"
    S["files"]["summary.txt"] = "summary.txt"
    S = _jsonable(S)
    json.dump(S, open(os.path.join(run_dir, "summary.json"), "w"), indent=1)
    with open(os.path.join(run_dir, "summary.txt"), "w") as f:
        f.write(summary_text(S))
    log(f"[collect] done in {tim['total']:.0f} s (peak host RSS {mem['peak_rss_GiB']:.1f} GiB, export workers "
        f"{mem['peak_rss_children_GiB']:.1f} GiB, peak GPU {peak_collect_gpu:.1f} GiB)")
    return S


def summary_text(S: dict) -> str:
    def g(x, fmt="{:+.4f}"):
        return "-" if x is None else fmt.format(x)
    L = [f"cshell summary: {S['name']}", "", f"DECISION: {S['outcome']}", f"  {S['reason']}", ""]
    m = S["mass"]
    L.append(f"mass target {g(m['target'], '{:.2f}')} g, final {g(m['final'], '{:.2f}')} g, best-candidate mesh "
             f"V/V0-1 {g(m['error'], '{:+.2e}')}")
    ex = S.get("export") or {}
    if ex.get("best_ok") is False:
        L.append(f"best-candidate export NOT MASS-MATCHED / failed (not re-scored): {ex.get('best_reject_reason')}")
    L += ["", f"{'design':<22}{'source':<10}{'J':>9}{'K/K0':>8}{'P/P0':>8}  governing   wall mm"]
    for r in S["comparison"]:
        L.append(f"{r['label']:<22}{r['source']:<10}{g(r['J']):>9}{g(r['K_K0'], '{:.3f}'):>8}"
                 f"{g(r['P_P0'], '{:.3f}'):>8}  {str(r['governing'] or '-'):<11} {g(r['w_min'], '{:.2f}')}-"
                 f"{g(r['w_max'], '{:.2f}')}")
    L += ["", "resolution checks (best export):"]
    for c in S["checks"]:
        L.append(f"  {c['res']:<7} J {g(c['J'])} K/K0 {g(c['K_K0'], '{:.3f}')} P/P0 {g(c['P_P0'], '{:.3f}')}  "
                 f"pass J/K/P {c['pass_J']}/{c['pass_K']}/{c['pass_P']}")
    L += ["", f"validation: {S['validation']}", f"sensitivity: accepted in "
          f"{S['sensitivity']['accepted_fraction']:.0%} of the κ_b × σ_f grid", "",
          "starts: " + "; ".join(f"{r['start']} {r['status']} J {g(r['J_opt'])}" for r in S["starts"]),
          f"timings s: " + ", ".join(f"{k} {v:.0f}" for k, v in S["timings_s"].items() if v is not None)]
    return "\n".join(L) + "\n"
