"""In-job finalise (PLAN_single_run §3, contract C3): the last step of `cshell run`; `cshell export <run_dir>` re-runs it.

    finalise(run_dir, ctx=None) -> summary dict

Steps (one GPU, a few GiB):
  1. Design: the final-phase xi of the run's start (a legacy run with several starts: the best feasible one by its
     optimiser J, `collect.start_table` / `pick_best`), D_comp from its last checkpoint; the optimiser's J, K/K0,
     λ_B/λ_B0, σ_PN/σ_PN0 from the last history.csv row of the final phase.
  2. Shifted-grid re-score: `optimise.evaluate(ctx, xi, phase, shift=(s, s, s), D_comp)` with s = ½ h_e of the final
     phase (C2). Device memory is returned before the export.
  3. Export (`export.export_design_gpu`, h = export.h_mm, default 0.2 mm): ψ tile by tile on the GPU, smooth opening,
     clean-up, escape holes, mass correction δ (Brent on the ψ volume, then secant steps on the mesh volume against
     the solid's exact volume), Warp marching cubes, checks, final_design.stl, escape_holes.csv, wall_thickness.csv.
  4. Decision: accepted iff J > ln(1 + min_gain) on the optimisation grid and on the shifted grid, the design is
     feasible on both grids for the run's own objective (§8: weighted → floors K/K0, P/P0; compliance → the λ_B, σ_PN constraints) and the export passed
     (`collect.export_verdict`: mass matched, watertight, |V_STL/V_target − 1| ≤ 2e-3). final_design.stl is always
     the design; `decision.accepted` says whether it replaces the solid.
  5. summary.json / summary.txt (C3 keys below), then `report.write_report(out_dir)`; a report failure is logged and
     recorded in summary["report"], it does not fail the run.

Opt-in: `sensitivity` (C2 `evaluate` on the optimisation grid for κ_b × σ_f variations of the config, sensitivity.csv)
and `bodyfit` (body-fitted C3D10 check of final_design.stl and the solid, `validate.validate_design`, validation.csv).
Dropped from the default (Owen 2026-10-31): best-of-starts, uniform-shell and dshell baselines, 3-grid check,
body-fitted check.

summary.json ("cshell.summary/2", units mm, g, s, GiB; null when not computed)
  name, decision {accepted, reason}
  mass {target_g (solid: exact volume of the tessellation), final_g (STL), error (V_STL/V_target − 1), delta_mm,
        V_target_mm3, V_final_mm3, density_g_mm3}
  opt {J, K_K0, lamB_ratio, sig_ratio, mass_err, start, xi_file, phase, D_comp}
  shift {J, K_K0, lamB_ratio, sig_ratio, feasible, governing, mass_err, shift_mm, seconds, error}
  export {h_mm, faces, watertight, bodies, volume_mm3, volume_rel, holes, voids {total, enclosed, holes, without_hole, per_void}, wall {min, p05, median,
          max, min_as_built, min_off_rims} (nominal, mm, ray casting; min_off_rims: samples off the escape-hole bores
          and lips), min_wall (export min-wall step: void_to_material_mm3, fins_removed_mm3), delta_mm, opening_removed_mm3, ok, why, file}
  closed_area_frac (closed share of the active Γ_h area), hollow_area_frac (share of the whole Γ_h area over an open
  void), timing {phase_s [Σ t_geometry + t_fe + t_mma + t_restore per phase], finalise_s, export_s, rescore_s,
  context_s}, peak {gpu_gib, host_gib} (whole job: optimiser history and finalise), peak_export {gpu_gib, host_gib}
  (the export alone), report {ok, path | error}, files {name: relative path}
"""
from __future__ import annotations

import copy
import csv
import json
import math
import os
import shutil
import time
from pathlib import Path

import numpy as np

from . import collect as col

ROOT = Path(__file__).resolve().parents[1]
SCHEMA = "cshell.summary/2"
COPY = ("config_used.json", "history.csv", "bc_faces.npz", "fields.vtp")    # small run artefacts the report reads


# =====================================================================================================================
# helpers
# =====================================================================================================================
def load_config(run_dir) -> dict:
    """config_used.json with the current defaults re-applied; a STEP path that no longer exists (deleted worktree) is
    looked up by file name in examples/parts."""
    from . import config as ccfg
    cfg = ccfg.load(json.load(open(os.path.join(run_dir, "config_used.json"))))
    sp = cfg.get("step_file")
    if sp and not os.path.exists(sp):
        alt = ROOT / "examples" / "parts" / os.path.basename(sp)
        if alt.exists():
            cfg["step_file"] = str(alt)
    return cfg


def hwm_reset() -> None:
    """Reset the process's peak resident set (VmHWM) so that a later `hwm_gib` covers one stage."""
    try:
        with open("/proc/self/clear_refs", "w") as f:
            f.write("5")
    except OSError:
        pass


def hwm_gib() -> float:
    try:
        for line in open("/proc/self/status"):
            if line.startswith("VmHWM:"):
                return int(line.split()[1]) / 2 ** 20
    except OSError:
        pass
    return float("nan")


def _peak_gpu():
    try:
        from .optimise import PeakGPU
        return PeakGPU().start()
    except Exception:
        return None


def _free_device():
    try:
        from .optimise import free_device_memory
        free_device_memory()
    except Exception:
        pass
    col.trim_host_memory()


def phase_seconds(hist: list[dict], start: str, n_ph: int) -> list:
    """Per phase Σ (t_geometry + t_fe + t_mma + t_restore) over the start's history rows (None without timings)."""
    out = []
    for k in range(n_ph):
        rows = [r for r in hist if r.get("start") == start and str(r.get("phase")) == str(k)]
        tot = [sum(col._f(r.get(c)) for c in ("t_geometry", "t_fe", "t_mma", "t_restore") if c in r) for r in rows]
        tot = [t for t in tot if np.isfinite(t)]
        out.append(float(sum(tot)) if tot else None)
    return out


def area_fracs(ctx, fields: dict, closed: np.ndarray) -> tuple:
    A = np.asarray(ctx.surf.A, float)
    act = ~np.asarray(fields["passive"], bool)
    cl = np.asarray(closed, bool) | (fields["a"] >= ctx.r_ma)
    Aa = float(A[act].sum())
    closed_frac = float(A[act & cl].sum() / Aa) if Aa > 0 else None
    hollow_frac = float(A[act & ~cl].sum() / A.sum()) if A.sum() > 0 else None
    return closed_frac, hollow_frac


def closed_threshold(cfg: dict) -> float:
    """ξ̃ᵃ at and above which the export closes a vertex: config export.closed_thr (= design.closure_xi, 0.9). Every
    run, Stage 2 included, is mapped by the single chain design.vertex_fields_final, so there is no legacy value."""
    return float((cfg.get("export") or {}).get("closed_thr", 0.9))


def _inv(x):
    try:
        return 1.0 / float(x)
    except (TypeError, ValueError, ZeroDivisionError):
        return None


def _floors_line(o: dict) -> str:
    s = f"floors: K/K0 ≥ {o.get('floor_K')}, P/P0 ≥ {o.get('floor_P')}"
    if o.get("floor_S") is not None:
        s += f", σ_PN0/σ_PN ≥ {o.get('floor_S')}"
    if o.get("floor_B") is not None:
        s += f", λ_B/λ_B0 ≥ {o.get('floor_B')}"
    return s


def _infeasible_why(r: dict) -> str:
    if r.get("objective") == "weighted":
        return (f"floors violated (K/K0 {r.get('K_K0')}, P/P0 {r.get('P_P0')} against {r.get('floor_K')}, "
                f"{r.get('floor_P')}; σ_PN0/σ_PN {_inv(r.get('sig_ratio'))} against {r.get('floor_S')}, "
                f"λ_B/λ_B0 {r.get('lamB_ratio')} against {r.get('floor_B')}; failed {r.get('floors')}; "
                f"governing {r.get('governing')})")
    return (f"constraints violated (λ_B/λ_B0 {r.get('lamB_ratio')}, σ_PN/σ_PN0 {r.get('sig_ratio')}, governing "
            f"{r.get('governing')})")


def decide(opt: dict, shift: dict | None, export_ok: bool, export_why: str, crit: float) -> tuple:
    """(accepted, reason) on the run's own objective (§8): J > ln(1 + g) and feasible (weighted: floors K/K0, P/P0;
    compliance: the λ_B, σ_PN constraints) on the optimisation and the shifted grid, and the export passed."""
    why = []
    J = opt.get("J")
    if J is None or not np.isfinite(J) or not J > crit:
        why.append(f"optimiser J = {J if J is not None else float('nan'):+.4f} ≤ ln(1 + g_min) = {crit:.4f}")
    if opt.get("feasible") is False:
        why.append("optimisation grid: " + _infeasible_why(opt))
    if shift is None:
        why.append("shifted-grid re-score not run")
    elif shift.get("error"):
        why.append(f"shifted-grid re-score failed: {shift['error']}")
    else:
        Js = shift.get("J")
        if Js is None or not np.isfinite(Js) or not Js > crit:
            why.append(f"shifted grid: J = {Js if Js is not None else float('nan'):+.4f} ≤ {crit:.4f}")
        if not shift.get("feasible", False):
            why.append("shifted grid: " + _infeasible_why(shift))
    if not export_ok:
        why.append(f"export failed: {export_why}")
    if why:
        return False, "original SOLID part kept: " + "; ".join(why)
    return True, (f"cshell design accepted: J = {J:+.4f} (optimisation grid), {shift['J']:+.4f} (shifted grid) "
                  f"> ln(1 + g_min) = {crit:.4f}, feasible, export mass-matched and watertight")


def summary_text(S: dict) -> str:
    def g(x, fmt="{:+.4f}"):
        return "-" if x is None else fmt.format(x)
    m, o, s, e = S["mass"], S["opt"], S.get("shift") or {}, S.get("export") or {}
    w = e.get("wall") or {}
    L = [f"cshell summary: {S['name']}", "", f"DECISION: {'accepted' if S['decision']['accepted'] else 'solid kept'}",
         f"  {S['decision']['reason']}", "",
         f"mass target {g(m.get('target_g'), '{:.3f}')} g, final {g(m.get('final_g'), '{:.3f}')} g "
         f"(V/V_target - 1 {g(m.get('error'), '{:+.2e}')}, delta {g(m.get('delta_mm'), '{:+.4f}')} mm)", "",
         f"objective: {o.get('objective') or s.get('objective') or '-'} (J = the run's objective; ln K/K0 separate)",
         f"{'grid':<14}{'J':>9}{'lnK/K0':>9}{'K/K0':>8}{'P/P0':>8}{'lamB':>8}{'sigPN':>8}  {'governing':<10} feasible",
         *(f"{lab:<14}{g(r.get('J')):>9}{g(r.get('J_K')):>9}{g(r.get('K_K0'), '{:.3f}'):>8}"
           f"{g(r.get('P_P0'), '{:.3f}'):>8}{g(r.get('lamB_ratio'), '{:.3f}'):>8}{g(r.get('sig_ratio'), '{:.3f}'):>8}"
           f"  {str(r.get('governing') or '-'):<10} {r.get('feasible')}{'  ' + r['error'] if r.get('error') else ''}"
           for lab, r in (("optimisation", o), ("shifted", s))),
         *([_floors_line(o)] if o.get("objective") == "weighted"
           else []),
         "",
         f"export h {e.get('h_mm')} mm: {e.get('faces')} faces, watertight {e.get('watertight')}, material bodies "
         f"{e.get('bodies')}, shells {e.get('shells')}, enclosed voids {(e.get('voids') or {}).get('enclosed')}, "
         f"escape holes {e.get('holes')}" + (f", UNVENTED voids {(e.get('voids') or {}).get('unvented')} (trapped "
                                            f"powder: manufacturing warning)"
                                            if (e.get('voids') or {}).get('unvented') else ""),
         f"fins removed by the export {g((e.get('min_wall') or {}).get('fins_removed_mm3'), '{:.2f}')} mm³ (check, "
         f"expected ≈ 0), void → material {g((e.get('min_wall') or {}).get('void_to_material_mm3'), '{:.2f}')} mm³",
         f"wall mm (nominal): min {g(w.get('min'), '{:.3f}')}, p05 {g(w.get('p05'), '{:.3f}')}, median "
         f"{g(w.get('median'), '{:.3f}')}, max {g(w.get('max'), '{:.3f}')}; min off the hole rims "
         f"{g(w.get('min_off_rims'), '{:.3f}')}",
         f"closed area {g(S.get('closed_area_frac'), '{:.1%}')} of active, hollow {g(S.get('hollow_area_frac'), '{:.1%}')}"
         " of the surface",
         "timing s: " + ", ".join(f"{k} {v:.0f}" for k, v in S["timing"].items() if isinstance(v, (int, float))),
         f"peak: GPU {g(S['peak'].get('gpu_gib'), '{:.2f}')} GiB, host {g(S['peak'].get('host_gib'), '{:.2f}')} GiB; "
         f"export alone GPU {g(S['peak_export'].get('gpu_gib'), '{:.2f}')} GiB, host "
         f"{g(S['peak_export'].get('host_gib'), '{:.2f}')} GiB"]
    return "\n".join(L) + "\n"


def _hk(hollow) -> dict:
    """`hollow=` keyword only when the run has a hollow-allowed set (C2 runs)."""
    return {} if hollow is None else {"hollow": hollow}


def _evaluate_fn(evaluate):
    if evaluate is not None:
        return evaluate
    import inspect
    from . import optimise
    fn = getattr(optimise, "evaluate", None)
    try:                                                  # objective.evaluate is imported into optimise as well
        c2 = fn is not None and "shift" in inspect.signature(fn).parameters
    except (TypeError, ValueError):
        c2 = False
    if not c2:
        raise AttributeError("optimise.evaluate(ctx, xi, phase, shift, D_comp) (contract C2) is not available in "
                             "this build")
    return fn


# keys of a C2 score kept in summary.json opt/shift (§8: the run's own objective J; J_K = ln K/K0 kept separate)
SCORE_KEYS = ("objective", "J", "J_K", "K_K0", "P_P0", "lamB_ratio", "sig_ratio", "mass_err", "governing", "floor_K",
              "floor_P", "floor_S", "floor_B", "floors")


def score(ctx, xi, phase: int, D_comp: float, evaluate=None, log=print, hollow=None, shifted: bool = True) -> dict:
    """C2 evaluate on the final-phase grid (shifted by ½ h_e in x, y, z, or the optimisation grid itself). Errors are
    recorded, not raised."""
    h_e = float(ctx.cfg["grid"]["phases"][phase]["h_e"])
    s = 0.5 * h_e if shifted else 0.0
    label = f"shifted grid (+{s:g} mm)" if shifted else "optimisation grid"
    t0 = time.time()
    try:
        r = dict(_evaluate_fn(evaluate)(ctx, xi, phase, shift=(s, s, s) if shifted else None, D_comp=D_comp,
                                        **_hk(hollow)))
        out = {k: r.get(k) for k in SCORE_KEYS}
        out["feasible"] = bool(r.get("feasible", False))
        out["error"] = None
    except Exception as e:                                # recorded; the decision rejects
        log(f"[finalise] {label} re-score FAILED: {e!r}")
        out = dict({k: None for k in SCORE_KEYS}, feasible=False, error=repr(e)[:300])
    out.update(shift_mm=[s, s, s], seconds=time.time() - t0)
    log(f"[finalise] {label}: {out.get('objective')} J {out['J']} (ln K/K0 {out.get('J_K')}) K/K0 {out['K_K0']} P/P0 "
        f"{out.get('P_P0')} lamB {out['lamB_ratio']} sig {out['sig_ratio']} governing {out['governing']} feasible "
        f"{out['feasible']} ({out['seconds']:.0f} s)")
    return out


def rescore_shift(ctx, xi, phase: int, D_comp: float, evaluate=None, log=print, hollow=None) -> dict:
    """`score` on the shifted grid."""
    return score(ctx, xi, phase, D_comp, evaluate, log, hollow, shifted=True)


def sensitivity(ctx, xi, phase: int, D_comp: float, out_dir: str, evaluate=None, log=print, hollow=None) -> dict:
    """Opt-in: C2 evaluate on the optimisation grid for buckling knockdown κ_b × strength σ_f variations
    (collect.KAPPAS × collect.SIGMA_PCT); sensitivity.csv."""
    ev = _evaluate_fn(evaluate)
    rows = []
    base = ctx.cfg
    for kb in col.KAPPAS:
        for sp in col.SIGMA_PCT:
            c2 = copy.copy(ctx)
            c2.cfg = copy.deepcopy(base)
            c2.cfg["capacity"]["knockdown"] = kb
            c2.cfg["material"]["strength_MPa"] = base["material"]["strength_MPa"] * (1 + sp / 100)
            try:
                r = ev(c2, xi, phase, shift=None, D_comp=D_comp, **_hk(hollow))
                rows.append(dict(kappa_b=kb, sigma_f=sp, sigma_f_MPa=c2.cfg["material"]["strength_MPa"], J=r.get("J"),
                                 lamB_ratio=r.get("lamB_ratio"), sig_ratio=r.get("sig_ratio"),
                                 feasible=int(bool(r.get("feasible")))))
            except Exception as e:
                rows.append(dict(kappa_b=kb, sigma_f=sp, error=repr(e)[:200]))
            _free_device()
    col._write_csv(os.path.join(out_dir, "sensitivity.csv"), rows,
                   ["kappa_b", "sigma_f", "sigma_f_MPa", "J", "lamB_ratio", "sig_ratio", "feasible", "error"])
    ok = [r for r in rows if "error" not in r]
    log(f"[finalise] sensitivity: {len(ok)}/{len(rows)} evaluated")
    return dict(file="sensitivity.csv", n=len(rows), feasible_fraction=float(np.mean([r["feasible"] for r in ok]))
                if ok else None)


def bodyfit(ctx, stl: str, out_dir: str, opt_like: dict, log=print) -> dict:
    """Opt-in body-fitted check (collect's §8.6.2 chain at its finer level): C3D10 of the solid and final_design.stl
    (isosurface-stuffing mesher, GPU engine), ratios compared with the voxel ratios `opt_like` (K_K0, lamB_ratio)."""
    from . import bodyfit_gpu as cbf
    from . import validate as cval
    cfg = ctx.cfg
    h = float(col._get(cfg, "validation.h_mm", 0.5)) / 2 ** 0.5
    kw = dict(h=h, h_bc=float(cfg["grid"]["phases"][-1]["h_e"]), n_modes=4, want_stress=True,
              nproc=int(col._get(cfg, "validation.nproc", min(32, len(os.sched_getaffinity(0))))),
              E=cfg["material"]["E_MPa"], nu=cfg["material"]["nu"], fallback=None,
              timeout=col._get(cfg, "validation.timeout_s", None), engine=col._get(cfg, "validation.engine", "gpu"),
              backend="stuff", remesh_kw=dict(stuff_c=float(col._get(cfg, "validation.stuff_lattice_over_h", 2 ** 0.5))))
    cbf.MAX_DOF = float(col._get(cfg, "validation.max_bodyfit_dof", 4e6))
    vd = os.path.join(out_dir, "validation")
    VE, TE, _ = col.load_mesh(stl)
    try:
        bf = {nm: cval.validate_design(Vs, Ts, ctx.V, ctx.T, ctx.TF, ctx.support, ctx.loads, vd, f"{nm}_h{h:.3f}",
                                       enforce_gate=True, **kw)
              for nm, (Vs, Ts) in (("solid", (ctx.V, ctx.T)), ("design", (VE, TE)))}
    except cbf.ScaleLimit as e:
        log(f"[finalise] body-fitted check not run: {e}")
        return dict(ran=False, scale_limited=True, why=str(e))
    except Exception as e:
        log(f"[finalise] body-fitted check FAILED: {e!r}")
        return dict(ran=False, why=repr(e)[:600], blocking_miss=True)
    K, L = opt_like.get("K_K0"), opt_like.get("lamB_ratio")
    vox = dict(solid=dict(compliance=1.0, lambda_b=1.0),
               design=dict(compliance=1.0 / K if K else None, lambda_b=L))
    bfd = {k: dict(compliance=v.compliance / bf["solid"].compliance, lambda_b=v.lambda_b / bf["solid"].lambda_b)
           for k, v in bf.items()}
    rows = cval.compare(vox, bfd)
    cval.write_validation_csv(rows, os.path.join(out_dir, "validation.csv"))
    return dict(ran=True, file="validation.csv", n_rows=len(rows), n_failed=sum(not r["passed"] for r in rows),
                blocking_miss=any(r["blocking"] and not r["passed"] for r in rows))


def write_report(out_dir: str, log=print) -> dict:
    """report.write_report (contract C4), guarded: a failure is recorded, never raised."""
    try:
        from . import report as crep
        fn = getattr(crep, "write_report", None)
        if fn is None:
            raise AttributeError("report.write_report (contract C4) is not available in this build")
        p = fn(out_dir)
        log(f"[finalise] report: {p}")
        return dict(ok=True, path=os.path.relpath(str(p), out_dir) if p else None)
    except Exception as e:
        log(f"[finalise] report FAILED (the run is not failed): {e!r}")
        return dict(ok=False, error=repr(e)[:300])


# =====================================================================================================================
# finalise
# =====================================================================================================================
def finalise(run_dir: str, ctx=None, *, out_dir: str | None = None, rescore: bool = True, report: bool = True,
             do_sensitivity: bool = False, do_bodyfit: bool = False, h_mm: float | None = None, evaluate=None,
             export_kw: dict | None = None, log=None) -> dict:
    """Finalise the run in `run_dir` (module docstring); outputs go to `out_dir` (default run_dir; another directory
    leaves run_dir untouched). `evaluate`: C2 function (default optimise.evaluate; tests pass a mock)."""
    from . import export as cex
    run_dir = os.path.abspath(run_dir)
    out_dir = os.path.abspath(out_dir or run_dir)
    os.makedirs(out_dir, exist_ok=True)
    same = out_dir == run_dir
    logf = open(os.path.join(out_dir, "finalise_log.txt"), "a")

    def _log(msg):
        print(msg, flush=True)
        logf.write(msg + "\n")
        logf.flush()
    log = log or _log
    T0 = time.time()
    pk = _peak_gpu()
    tim, files = {}, {}
    try:
        cfg = load_config(run_dir)
        if not same:
            for nm in COPY:
                src = os.path.join(run_dir, nm)
                if os.path.exists(src):
                    shutil.copyfile(src, os.path.join(out_dir, nm))
                    files[nm] = nm
            json.dump(cfg, open(os.path.join(out_dir, "config_used.json"), "w"), indent=1)
        n_ph = len(cfg["grid"]["phases"])
        last = n_ph - 1
        h_g_last = float(cfg["grid"]["phases"][last]["h_g"])

        # 1. design ---------------------------------------------------------------------------------------------
        rows = col.start_table(run_dir, cfg, merge=same)
        best = col.pick_best(rows)
        if best is None:
            raise RuntimeError("no finished feasible start in " + run_dir + ": " + "; ".join(
                f"{r['start']} {r['status']} feasible {r['feasible']}" for r in rows))
        xi = np.load(os.path.join(run_dir, best["xi_file"]))
        D_comp = float(best["D_comp"])
        hollow = None
        try:                                              # C2 run: design_<start>_phase<k>.npz holds H and D_comp
            from .optimise import load_design
            dd = load_design(run_dir, best["start"], last)
            xi, hollow, D_comp = dd["xi"], dd["hollow"], float(dd["D_comp"])
        except (ImportError, FileNotFoundError, KeyError) as e:
            log(f"[finalise] no design npz ({e}); xi from {best['xi_file']}, hollowing allowed everywhere")
        opt = dict(J=best["J_opt"], K_K0=best["K_K0_opt"], lamB_ratio=best.get("lamB_ratio_opt"),
                   sig_ratio=best.get("sig_ratio_opt"), mass_err=best["mass_err"], start=best["start"],
                   xi_file=best["xi_file"], phase=last, D_comp=D_comp)
        tol = float(cfg["objective"].get("feasibility_tol", 1e-3))
        lb, sg = opt["lamB_ratio"], opt["sig_ratio"]
        opt["feasible"] = (bool(np.log(1.0 / lb) <= tol and np.log(sg) <= tol)        # C2 constraints from history
                           if lb is not None and sg is not None and np.isfinite(lb) and np.isfinite(sg) and lb > 0
                           and sg > 0 else None)
        log(f"[finalise] design: start {best['start']} ({best['xi_file']}, D_comp {D_comp:g} mm), optimiser J "
            f"{opt['J']:+.4f}" + (f" (best of {len(rows)} starts)" if len(rows) > 1 else ""))
        t0 = time.time()
        if ctx is None:
            from .optimise import Context
            ctx = Context(cfg, log=log)
        tim["context_s"] = time.time() - t0
        nv = ctx.surf.n_vertices
        if len(xi) != 2 * nv:
            raise ValueError(f"xi has {len(xi)} entries, Γ_h has {nv} vertices (2N expected)")

        # 2. shifted-grid re-score ------------------------------------------------------------------------------
        shift = None
        peak_rescore = dict(gpu_gib=None, host_gib=None)
        if rescore:
            # sr-fix 3: the optimiser's FE engine, band and cached CuPy/cuDSS blocks are released before the re-score,
            # which builds one PhaseModel with the optimiser's engine and settings (optimise.evaluate)
            _free_device()
            hwm_reset()
            pkr = _peak_gpu()
            # §8: the final design (the best growth boundary, possibly not the last history row) scored with the
            # run's own objective on the optimisation grid; history-derived values kept as opt_history
            o_s = score(ctx, xi, last, D_comp, evaluate, log, hollow, shifted=False)
            _free_device()
            if not o_s.get("error"):
                opt_hist = {k: opt.get(k) for k in ("J", "K_K0", "lamB_ratio", "sig_ratio", "feasible")}
                opt.update({k: o_s.get(k) for k in SCORE_KEYS}, feasible=o_s["feasible"], history=opt_hist)
            shift = rescore_shift(ctx, xi, last, D_comp, evaluate, log, hollow)
            peak_rescore = dict(gpu_gib=pkr.stop() if pkr is not None else None, host_gib=hwm_gib())
            shift["peak"] = peak_rescore
            tim["rescore_s"] = shift["seconds"]
            _free_device()
            log(f"[finalise] shifted-grid re-score peak: GPU {peak_rescore['gpu_gib']} GiB, host "
                f"{peak_rescore['host_gib']} GiB")
        sens = None
        if do_sensitivity:
            t0 = time.time()
            sens = sensitivity(ctx, xi, last, D_comp, out_dir, evaluate, log, hollow)
            files["sensitivity.csv"] = "sensitivity.csv"
            tim["sensitivity_s"] = time.time() - t0

        # 3. export ---------------------------------------------------------------------------------------------
        h = float(h_mm if h_mm is not None else (cfg.get("export") or {}).get("h_mm", 0.2))
        thr = closed_threshold(cfg)
        fields = cex.vertex_fields(ctx, xi, D_comp, h_g_last, **_hk(hollow))
        closed = fields["xt_a"] >= thr
        _free_device()
        hwm_reset()
        pke = _peak_gpu()
        t0 = time.time()
        ex = cex.export_design_gpu(ctx, xi, out_dir, D_comp, name="final_design", h_g=h_g_last, fields=fields,
                                   closed_thr=thr, h_x=h, log=log, **(export_kw or {}))
        tim["export_s"] = time.time() - t0
        peak_export = dict(gpu_gib=pke.stop() if pke is not None else None, host_gib=hwm_gib())
        export_ok, export_why = col.export_verdict(ex)
        for nm in ("final_design.stl", "final_design_export.json", "escape_holes.csv", "wall_thickness.csv"):
            if os.path.exists(os.path.join(out_dir, nm)):
                files[nm] = nm
        mesh, mc, w = ex.get("mesh", {}), ex.get("mass_correction", {}), ex.get("wall", {})
        log(f"[finalise] export: {mesh.get('faces')} faces, watertight {mesh.get('watertight_manifold')}, bodies "
            f"{mesh.get('bodies')}, V/V_target-1 {mesh.get('volume_rel', float('nan')):+.2e}, delta "
            f"{mc.get('delta', float('nan')):+.4f} mm, {tim['export_s']:.0f} s, peak GPU {peak_export['gpu_gib']} GiB "
            f"host {peak_export['host_gib']:.2f} GiB; {'ok' if export_ok else 'FAILED: ' + export_why}")

        # 4. optional body-fitted check, decision -----------------------------------------------------------------
        val = None
        if do_bodyfit:
            t0 = time.time()
            val = bodyfit(ctx, os.path.join(out_dir, "final_design.stl"), out_dir, shift or opt, log)
            tim["bodyfit_s"] = time.time() - t0
            if val.get("ran"):
                files["validation.csv"] = "validation.csv"
        crit = math.log(1.0 + float(cfg["objective"].get("min_gain", 0.02)))
        acc, reason = decide(opt, shift, export_ok, export_why, crit)
        if val is not None and val.get("blocking_miss"):
            acc, reason = False, "original SOLID part kept: body-fitted check missed a ratio (validation.csv)"
        log(f"[finalise] DECISION: {reason}")

        # 5. summary, report --------------------------------------------------------------------------------------
        hist = col.read_history(run_dir, merge=False)
        rho = float(cfg["material"]["density_g_cm3"]) * 1e-3
        V0, Vf = ex.get("V0_mm3"), mesh.get("volume")
        cfr, hfr = area_fracs(ctx, fields, closed)
        gpu_hist = [col._f(r.get("gpu_peak_GiB", r.get("gpu_mem_GiB"))) for r in hist if r.get("start") == best["start"]]
        gpu_hist = [x for x in gpu_hist if np.isfinite(x)]
        peak_fin = pk.stop() if pk is not None else float("nan")
        import resource
        host = max(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 2 ** 20, hwm_gib())   # VmHWM after a reset
        tim["finalise_s"] = time.time() - T0
        S = dict(
            schema=SCHEMA, name=cfg["name"], decision=dict(accepted=acc, reason=reason),
            mass=dict(target_g=V0 * rho if V0 else None, final_g=Vf * rho if Vf else None, error=mesh.get("volume_rel"),
                      delta_mm=mc.get("delta"), V_target_mm3=V0, V_final_mm3=Vf, density_g_mm3=rho),
            opt=opt, shift=shift,
            export=dict(h_mm=h, faces=mesh.get("faces"), watertight=mesh.get("watertight_manifold"),
                        bodies=mesh.get("bodies"), shells=mesh.get("shells"), void_shells=mesh.get("void_shells"),
                        shell_volumes_mm3=mesh.get("shell_volumes_mm3"), volume_mm3=Vf, volume_rel=mesh.get("volume_rel"),
                        holes=(ex.get("voids") or {}).get("holes", 0), voids=ex.get("voids") or dict(total=0, enclosed=0),
                        wall=dict(min=w.get("min_nominal"), p05=w.get("p05_nominal"), median=w.get("median_nominal"),
                                  max=w.get("max_nominal"), min_as_built=w.get("min_as_built"),
                                  min_off_rims=(ex.get("wall_off_rims") or {}).get("min_nominal")),
                        min_wall={k: v for k, v in (ex.get("min_wall") or {}).items() if k != "fin_points"},
                        delta_mm=mc.get("delta"), mesh_steps=mc.get("mesh_steps"),
                        opening_removed_mm3=(ex.get("cleanup") or {}).get("opening_removed_mm3"),
                        closed_thr=thr, ok=export_ok, why=export_why or None, file="final_design.stl",
                        timing_s=ex.get("timing_s")),
            closed_area_frac=cfr, hollow_area_frac=hfr,
            timing=dict(phase_s=phase_seconds(hist, best["start"], n_ph), **tim),
            peak=dict(gpu_gib=max(gpu_hist + ([peak_fin] if np.isfinite(peak_fin) else [])) if (
                gpu_hist or np.isfinite(peak_fin)) else None, host_gib=host),
            peak_export=peak_export, peak_rescore=peak_rescore,
            peak_optimiser=dict(gpu_gib=max(gpu_hist) if gpu_hist else None,
                                host_gib=max([x for x in (col._f(r.get("host_peak_GiB")) for r in hist
                                                          if r.get("start") == best["start"]) if np.isfinite(x)],
                                             default=None)),
            sensitivity=sens, validation=val, report=None, files=files)
        S["files"].update({"summary.json": "summary.json", "summary.txt": "summary.txt",
                           "finalise_log.txt": "finalise_log.txt"})

        def dump():
            json.dump(col._jsonable(S), open(os.path.join(out_dir, "summary.json"), "w"), indent=1)
            with open(os.path.join(out_dir, "summary.txt"), "w") as f:
                f.write(summary_text(col._jsonable(S)))
        dump()
        if report:
            S["report"] = write_report(out_dir, log)
            if S["report"].get("ok") and S["report"].get("path"):
                S["files"]["report"] = S["report"]["path"]
            dump()
        log(f"[finalise] done in {time.time() - T0:.0f} s")
        return col._jsonable(S)
    finally:
        if pk is not None:
            pk.stop()
        logf.close()
