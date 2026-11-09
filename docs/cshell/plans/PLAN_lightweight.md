# PLAN_lightweight — lightweighting as the default cshell mode

Agreed with Owen 2026-11-09. Branch `worktree-cshell-lightweight` (worktree `.claude/worktrees/cshell-lightweight`).

## Goal
cshell's default becomes a **lightweighter**: from the original solid, find the lightest shelled design whose
stiffness and capacity are at least those of the solid. The current constant-mass weighted objective stays available
as a non-default option (`objective.type = "weighted"`), behaviour unchanged.

## Formulation (`objective.type = "lightweight"`, new default)
    minimise   f0 = V/V0                                  (no restore_mass; V free)
    subject to ln K/K0 ≥ ln m_K,  ln P/P0 ≥ ln m_P       (MMA constraints, reuse floor_K / floor_P plumbing)
with margins `objective.lw_margin_stiffness = objective.lw_margin_capacity = 1.02` (config). All other constraints
(t_min, v_min, as-built biases, passive/mount regions, closure skirt, growth through R quantiles) unchanged.
- The solid start is mildly infeasible (K/K0 = P/P0 = 1 < 1.02); MMA must recover. Check the first iterations.
- Scale f0 sensibly for MMA (e.g. f0 = V/V0 or ln V/V0, gradient vol_grad/V0); choose one, document it.
- `J` in history/logs for lightweight = mass saving 1 − V/V0 (record V/V0, K/K0, P/P0, max constraint violation).
- Growth schedule "best boundary": the boundary with the lowest V among those that are feasible (constraints ≤ feasibility
  tol); if none feasible, the least-infeasible. Stopping/patience in terms of the lightweight J.
- Convergence tolerances (tol_f etc.) must make sense for the new f0.

## Downstream
- Export (`export.py`) δ mass correction targets the optimised V (V_opt from the design), not V0, in lightweight mode.
- Finalise / decision rule / summary: accept iff K/K0 ≥ 1 and P/P0 ≥ 1 on the optimisation **and** shifted-grid
  re-score; report V/V0 (mass saving %), K/K0, P/P0, governing mode, walls. Mass check = export V vs V_opt.
- report.py: must not break; label mass saving rather than "constant mass" where cheap to do so.
- README, config comments, THEORY_continuous.md: short note on the new default and how to select `weighted`.

## Tests
- Existing CPU and GPU suites must pass (weighted path unchanged; tests that assume the default is weighted should pin
  `objective.type="weighted"` explicitly).
- New tests: lightweight settings build the right constraints; no mass restoration; gradient check of f0; growth
  best-boundary picks lowest feasible V; export target V_opt; decision rule ≥ 1.

## Study
1. Smoke: Cantilever lightweight with an iteration cap, check V/V0 decreases and constraints approach feasibility.
2. Full: Cantilever, L-bracket, MBB (`examples/cshell/*.json`), lightweight default, 1 GPU job each, in parallel.
   Output under `/projects/b6bl/owenpeckham.b6bl/home_data/shell-lightweighting/runs_cshell/LW1/<part>`.
3. Comparison baseline: existing round-4 constant-mass runs `runs_cshell/SR_VAL4/<part>` (not re-run).
Round-4 took 35–44 min per part (optimise ~30–37 min + finalise 5–8 min).

## Slurm (keep conservative so jobs backfill)
- Full runs: `--gpus=1`, `--time` ≤ 1:15:00, explicit `--mem` from the round-4 peak host RSS + ~30 %, minimal CPUs.
- Tests/smoke: ≤ 30 min, small mem. ≤ 4 slots in use at once. Nothing on the login node (not even imports).
