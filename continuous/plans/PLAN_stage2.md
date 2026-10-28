# PLAN: Stage 2 completion, built in parallel (agreed with Owen 2026-10-28: skip the Stage 1 compliance-only part runs, go to full Stage 2 runs; build in parallel)

Worktree `/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2` (= REPO below), branch
`cshell-stage2` (branched from `cshell`). Work and run everything from REPO. **Do not commit, switch branch or push**;
the orchestrator commits. Touch only your task's files. Spec: `continuous/THEORY_continuous.md` v1.2 (cite § in
docstrings). Common rules (env, sbatch, tests, style, units mm/N/MPa, float64 reductions, no silent fallbacks) as in
`continuous/plans/PLAN_parallel_modules.md` "Common rules". Env: `source continuous/cluster/env.sh`. Login node: edit
only; tests and runs via `sbatch continuous/cluster/{cpu_test,gpu_test,run}.sh` from REPO (logs in `REPO/logs/`).
`runs_cshell` and `.cache_cshell` are symlinks to projects storage shared with other checkouts: write new runs under
`runs_cshell/<your task letter>_*` only. Respect your slot budget. Test on the smallest case first (B-rods in
`examples/cshell/bench/`, then the cantilever with `--set optimiser.max_iter=[2,2]`).

Current state (read, don't redo): `cshell/optimise.py` (Context, PhaseModel, Run: phases, restoration, MMA,
checkpoints, `--resume`, `--max-hours`, §3.5 extension), `design.py` (DesignMap; NumPy/CuPy via `xp` from the cell
arrays), `band.py`, `fe_sparse.py` (cuDSS; fp32-factor-preconditioned fp64 CG when fp64 does not fit), `objective.py`,
`export.py`, `validate.py`. Timing run `runs_cshell/timing/Cantilever` (compliance only): phase B (0.4 mm, 220 M band
cells, 6.1 M DOF) per iteration t_geometry 5.8 s, t_fe 10.7 s, t_mma 14.5 s, t_restore 37 s (5 volume evaluations);
GPU 84 GiB of 95.6. **The CPU geometry chain is 80 % of the iteration.**

## Task G: geometry chain on the GPU (`cshell/optimise.py` PhaseModel and Run loop except `Run.initial_xi`, `cshell/design.py`, `cshell/band.py`, `cshell/mma.py`, `cshell/config.py`)
Goal: phase-B iteration ≲ 15 s (compliance part) without changing any result beyond float tolerance.
1. Hold the band data the chain needs (φ, P, cell→element map) on the GPU as CuPy when CUDA is present (config
   `grid.geometry_device`: "auto" | "cpu" | "gpu"; log the choice). Compact storage: φ float32 (chain arithmetic and all
   reductions — V, ∇V, pull-backs to vertices — in float64), elem int32, P as CSR with int32 indices / float32 data or
   (triangle id int32 + 3 float32 barycentrics) — whichever is faster and smaller. Optionally drop cells whose ψ cannot
   cross ±ε for any admissible (d, a) (constant χ: count them into V once).
2. Restoration (`mma.restore_mass`), volume, volume gradient, sliver, pull-backs and ρ all run on the device; only
   per-vertex vectors (≈10⁵) and per-element ρ for the FE cross to the host (or pass ρ to `SparseFE` on device).
   Also profile `t_mma` (14.5 s): split it into pull-back / gradients / `mmasub` and fix whatever dominates.
3. GPU memory budget: phase B FE uses up to 84 GiB (cantilever, compliance only; more with buckling, MBB larger).
   Measure the geometry footprint; if FE + geometry would exceed the HBM, free/shrink geometry buffers around the factor
   or fall back to CPU geometry **with a log line**. Record the numbers in `continuous/BENCHMARKS.md`.
4. Tests: existing `tests/cshell/test_design.py`, `test_mma.py`, `test_end_to_end.py` pass on CPU; new `@pytest.mark.gpu`
   test that GPU and CPU chains agree (V to 1e-10 relative, gradients to 1e-6) on the B2 rod, and that a 3-iteration B2
   run gives the same history (J to 1e-6) on both devices.
5. Measure: cantilever and MBB, full objective (default config: w_K = w_P = 0.5, buckling on), `max_iter=[2,3]`, uniform
   start: per-component time per iteration and peak GPU memory in phases A and B. Report whether the MBB with buckling
   fits at 0.4 mm (STATUS K11); if not, try in order: fewer eigensolver block vectors, the fp32-preconditioned path.
Slots: 3.

## Task C: collect, decision, vshell start, CLI (`cshell/collect.py` new, `cshell/cli.py`, `Run.initial_xi` in `optimise.py` only, `continuous/cluster/`)
1. **S2 vshell start (§7.4):** load `examples/runs/<Name>/arrays.npz` (+ `regions.csv`, `summary.json`; find the
   region-label grid and θ of vshell's best design there — read the vshell code that writes them) or run vshell if
   absent; ξᵃ_i = θ of the region containing p_i − (r_MA,i/2) n_i, ξᵈ_i = s(1 − θ)/d^{hi}_i (clip to [0,1]).
   Config key `optimiser.vshell_run` (default `examples/runs/<name>`). Test on the cantilever (CPU job).
2. **`python -m cshell collect <run_dir>` (§5.4, §8, §10.5):** best feasible start by final J; export via `export.py`
   (holes sited by final stress if available); **re-score on the export geometry** (voxel FE of the exported nominal ψ
   through the same as-built pipeline as the optimiser) at (i) the optimisation grid, (ii) grid shifted by ½h_e in x, y,
   z, (iii) h_e = 0.5 mm (h_g = 0.125); the same for the solid and for the uniform shell (§7.4 S1 note: ψ = max(φ − t,
   −φ), t for constant mass); body-fitted check via `validate.py` (solid and final design, tolerances §8.6; a ratio miss
   blocks acceptance); decision rule J* > ln(1+g_min), K/K₀ ≥ f_K, P/P₀ ≥ f_P at every resolution, else
   `final_design.stl` = original with the reason. Writes `summary.json`, `summary.txt`, `validation.csv`,
   `sensitivity.csv` (κ_b ∈ {0.6,0.7,0.85,1.0} × σ_f −20 … +20 % in 10 % steps, re-scored without new FE since λ_b ∝ κ_b,
   λ_s ∝ σ_f), `best_candidate.stl`, `final_design.stl`, `original.stl`, `uniform_shell.stl`, `fields.vtp` (Γ_h: ξ, d,
   a, w, r_MA, d^{hi}, passive). Per-start rows (all starts) and the vshell row from its `summary.json`. Document the
   `summary.json` schema in the module docstring early (Task R reads it).
3. CLI: `collect`, `setup` (geometry, band, P, caches; prints DOF and memory plan), `faces`/`init` (delegate to vshell's
   equivalents), `bench B1|B2|B3|all` (runs `examples/cshell/bench/*.json`), `report` (calls `cshell.report.build(run_dir)`
   — Task R). Update the module docstring.
4. `continuous/cluster/run.sh`: confirm it resubmits with `--resume` via `--dependency=afterany` and that one start per
   job works (`--start`); add `continuous/cluster/collect.sh`.
5. Tests (`tests/cshell/test_collect.py`, CPU, B-rod scale): decision-rule truth table; uniform-shell t solve; re-score of
   the solid equals the baseline; summary schema. One GPU end-to-end `run` + `collect` on the B2 rod.
Slots: 3.

## Task R: report (`cshell/report.py` new, `tests/cshell/test_report.py`)
`build(run_dir) -> path` writes `report.pdf` (§9, 15 pages) from files in the run dir only (rebuildable): follow
`vshell/report.py` (text pages first, one figure per page, each figure in its own try/except with failures in
`plots/report_log.txt`, rasterised 3D, PDF < 5 MB). Inputs: `summary.json` (schema in `cshell/collect.py`'s docstring
once it appears; until then use §9 page 1 and §10.5), `history.csv` (columns `optimise.HIST_COLS`), `config_used.json`,
`run_info.json`, `xi_<start>_phase<k>.npy`, `fields.vtp`, `validation.csv`, `sensitivity.csv`, `escape_holes.csv`,
`wall_thickness.csv`. Pages whose inputs are missing render a "not available" page rather than failing. Provenance
page: git commit + dirty flag, SHA-256 of STEP, config and every source file incl.
`continuous/reference/{analysis,gpu_analysis}.py`, package versions, CUDA driver/runtime, host, job id (reuse
`vshell/util.py`). Tests: build from a synthetic run dir (fabricated small files) and from `runs_cshell/timing/Cantilever`
(partial: no summary) without exceptions; page count; size < 5 MB. CPU only. Slots: 1.

## Return (each, under 250 words)
Files written; test results with job ids; measured timings; deviations from the spec and why; open issues.
