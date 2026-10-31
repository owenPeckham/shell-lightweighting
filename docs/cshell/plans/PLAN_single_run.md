# PLAN: single run from the solid, in-job finalise, smooth fields, solid regions, dshell-style report (DRAFT)

Branch `worktree-cshell-solid-pockmarks`. Evidence: `docs/cshell/research/pockmarks/FINDINGS.md`.
Rules: nothing on the login node; optimiser and export use one GPU and minimum memory; slot cap 24.

## 1. One start, from the original solid
- New start `solid` (the default; `uniform`/`dshell`/`neutral` stay available as options).
  - ξᵈ = 0 and a = r_MA (inset on the medial axis), so the part is geometrically solid but a zero-thickness void surface exists to give a non-zero gradient.
  - The mass target already equals the solid's mass, so the start is feasible and no mass ramp is needed.
- `cshell run <part.json>` runs phases A and B, then finalise (§3), in one Slurm job. `launch_part.sh` and `run.sh` lose the 3-start fan-out and the collect step.

## 2. Smooth fields (dimples)
- Filter twice, ξ̃ = clip(F² ξ): one more solve with the existing factor, and the adjoint is Fᵀ Fᵀ.
- Smooth r_MA once with a ~1 mm Helmholtz filter before computing a_hi and the a ≥ r_MA closure.
- Taper d to 0 towards passive regions over r_f, replacing the hard override, so there is no cliff.
- Export opening: use a smooth distance-based ψ in removed void cells instead of the constant −h/2.

## 3. In-job finalise (replaces collect)
- Steps: mass correction δ, then one shifted-grid re-score (accept if J > ln 1.02 on both grids), then export, then the report.
- GPU export at h_x = 0.2 mm, replacing the 0.05 mm CPU export:
  - ψ is evaluated tile by tile with the existing fused kernels, then marching cubes on the GPU (Warp), streaming triangles to the STL.
  - Escape holes and island removal are kept.
  - Memory target: a few GB of GPU and host memory.
  - The STL is small enough to be its own preview.
- Dropped from the default run: best-of-starts, the uniform-shell and dshell baselines, the 3-grid check and the body-fitted check. Sensitivity and body-fitted stay available as flags.
- `cshell export <run>` re-exports a finished run.

## 4. Objective and shear/flexure-driven hollowing (Owen 2026-10-31: option iii)

### Objective and constraints (Owen, 2026-10-31)
- Objective: minimise compliance at constant mass (the solid's), J = ln K/K0.
- Constraints, all relative to the original solid:
  - mass V = V_solid;
  - buckling λ_B ≥ λ_B0 (lowest load factor, block-Krylov as now);
  - strength σ_PN ≤ σ_PN0, the stress p-norm outside the exclusion zones. This replaces the softmin capacity in the objective.
- Both buckling and strength enter MMA as constraints, normalised as ln(λ_B0/λ_B) ≤ 0 and ln(σ_PN/σ_PN0) ≤ 0.
- Accept when J > ln 1.02 and both constraints hold on the shifted grid.

### Exclusion zone
A cell is excluded within r_ex of any constrained or loaded face, with r_ex = max(2 mm, 2 r_MA at that face), roughly one section depth (Saint-Venant).

### Shear/flexure ratio R_v (dshell's R, without a skeleton)
1. At each Γ_h vertex, take the local wall frame (n = surface normal; t₁, t₂ in the tangent plane).
2. Along the through-thickness line from the surface to the medial axis, integrate:
   - U_flex = ∫ (σ_11² + σ_22²)/2E + σ_12²/2G — in-plane membrane and bending, plus in-plane shear (torsion, which hollow closed sections resist well);
   - U_shear = ∫ (σ_1n² + σ_2n²)/2G — transverse shear.
3. R_v = U_shear / U_flex, smoothed with the design filter.
- Behaviour: bending gives σ peaking at the surface with zero transverse shear (R → 0, core idle, hollow it). Shear-dominated joints give τ peaking at the core (R large, core working, keep it solid).

### Growth
- Start from the solid (§1).
- The hollow-allowed set H = {v : R_v ≤ R_k}, enforced as a bound ξᵃ_v ≥ 1 outside H, tapered over r_f.
- Growth step: R_k rises through the quantiles of R. After each rise, MMA re-optimises d and a inside H for a few iterations. R is recomputed from the current design's stress at each step, because the stress redistributes.
- Growth stops when a further step raises J by less than tol_grow (default 1e-3). There is no upper limit on R (Owen). The extent of hollowing is therefore decided by performance, and the order by flexural dominance, as in dshell's mass sweep.
- Closure is reachable: a vertex outside H is closed (a = a_hi), and the export closes at ξ̃ᵃ ≥ 0.9 with a smooth projection.
- Phase B (0.4 mm) refines with H frozen.

## 5. Report (dshell-style, `cshell/report.py` rewrite)
- Pages:
  1. summary
  2. geometry and BCs
  3. 3D renders of the final design: hollow coloured by wall thickness, solid light grey, mounts dark grey; 4 views
  4. sections through the thickest regions
  5. von Mises stress, solid vs final
  6. capacity by mode, solid vs final
  7. wall-thickness histogram
  8. convergence
  9. voids and escape holes
  10. resolution check
  11. run time
- Renders use a matplotlib Poly3DCollection with face shading on the 0.2 mm STL, decimated in the job.
- Loads and supports drawn from the faces the analysis uses (`bc_faces.npz`), not "indicative placement":
  - supports: the constrained faces highlighted;
  - loads: one arrow per connected loaded patch, at its area centroid, pointing along F and scaled by its share of |F|;
  - a test asserts that each arrow's base lies on the part surface, within 1 mm of a loaded face.
- Check the loads against manuscript Fig. 3/4 and p.12, before any run:
  - Cantilever: four bolt tabs fixed; load downward on the tip block.
  - L-bracket: two top bolt holes fixed; load downward on the end block.
  - MBB (mirrored full beam): end bolt tabs fixed; load downward on the two mid-span pads, one per beam. The current plot draws a single arrow at y = 0, between the beams.
  - Any mismatch between the config and the manuscript goes to Owen before running.
- Fixed defects: overlapping colourbars, the histogram axis, the occlusion-free scatter plots.

## 6. Tests and validation
- R check against beam theory, on a rectangular cantilever with a tip load: U_s/U_b = (3/5)·6EI/(G A L²) for L/h = 2, 5, 20, within 10 %.
- Exclusion-zone test.
- Growth-loop test on the tiny case.
- Unit tests: F² adjoint (finite-difference check), solid-start gradient non-zero, closure projection, the streaming exporter (watertight output, volume matching ψ), report smoke test.
- Validation: 3 parts × 1 solid start, full phases (3 GPU jobs). Compare J with Stage 2 (MBB 0.481, Cantilever 0.490, L-bracket 0.129), and measure the dimple RMS with the diagnosis script.
- Then update `examples/cshell/runs`, the docs (README, STATUS) and a PR.

## 7. Work split and contracts (three parallel developers)
Base branch: `worktree-cshell-solid-pockmarks`. Each developer has its own branch: `sr-core`, `sr-export` or `sr-report`.

**File ownership.** Edit only the files you own; anything else needs a request through the orchestrator.

| Developer | Owns |
|---|---|
| core | design.py, surface.py, objective.py, optimise.py, fea.py, fe_*.py, mma.py, band.py, config.py, tests/cshell/test_{design,objective,optimise,surface,shear_flexure,growth}*.py |
| export | export.py, collect.py, a new finalise.py, cli.py, cshell/cluster/*.sh, tests/cshell/test_{export,finalise,collect}*.py |
| report | report.py, tests/cshell/test_report*.py |

**C1. Config keys** (core adds them to `config.DEFAULTS`; other branches read them with `.get(..., default)`):
- `objective`: `{"type": "compliance", "constraints": ["buckling", "strength"]}`
- `optimiser.starts`: `["solid"]`
- `growth`: `{"tol": 1e-3, "quantiles": [0.1, 0.2, ..., 1.0], "iters_per_step": 8}`
- `capacity.exclude_bc_mm`: as now, with r_ex = max(exclude_bc_mm, 2 r_MA)
- `export`: `{"h_mm": 0.2, "device": "gpu"}`

**C2. Core API** (export calls it):
- `optimise.evaluate(ctx, xi, phase: int, shift: tuple | None = None, D_comp=None) -> dict`
  - Returns: J, K_K0, lamB_ratio (= λ_B/λ_B0), sig_ratio (= σ_PN/σ_PN0), mass_err, feasible (bool), governing.
  - Uses the same model as the optimiser's final phase; `shift` offsets the FE grid as `collect.grid_shift` does.
- `optimise.Run(cfg, "solid", out_dir).run(ctx)` writes:
  - `xi_solid_phase{k}.npy`;
  - `history.csv` with columns it, phase, growth_step, R_k, J, K_K0, lamB_ratio, sig_ratio, mass_err, closed_frac, hollow_allowed_frac;
  - `fields.vtp` point data: d, a, w, r_MA, R, closed, hollow_allowed;
  - `bc_faces.npz` (as now, with keys support_faces, load_faces, force_N).

**C3. Finalise** (export): `finalise.finalise(run_dir, ctx=None) -> dict`
- Run as the last step of `cli run`; `cshell export <run_dir>` re-runs it.
- Writes:
  - `final_design.stl` (0.2 mm, GPU, streamed);
  - `summary.json`;
  - `summary.txt`;
  - `escape_holes.csv`;
  - `wall_thickness.csv`;
  - then calls `report.write_report(run_dir)`.
- `summary.json` keys used by the report:
  - name, decision {accepted, reason};
  - mass {target_g, final_g};
  - opt {J, K_K0, lamB_ratio, sig_ratio};
  - shift {J, K_K0, lamB_ratio, sig_ratio, feasible};
  - export {h_mm, faces, watertight, volume_mm3, holes, voids, wall {min, p05, median, max}};
  - closed_area_frac, hollow_area_frac;
  - timing {phase_s: [...], finalise_s, export_s};
  - peak {gpu_gib, host_gib}.

**C4. Report** (report): `report.write_report(run_dir) -> path`
- Reads only the C2/C3 artefacts plus config_used.json.
- Must also work on the Stage 2 outputs in `examples/cshell/runs/*` through an adapter, which tolerates missing keys and writes a placeholder page for each.
