# cshell hand-over, 2026-10-29

For the next Claude Code session. Read this first, then `continuous/STATUS.md` (stage log), `continuous/VALIDATION.md`,
`continuous/BENCHMARKS.md`, and the two in-progress plans named below. The spec is `continuous/THEORY_continuous.md`
(v1.2 + §4.4(c) K_σ note), the original brief `continuous/PROMPT_claude_code.md`.

## Where the code is
- **Branch `cshell-stage2`**, worktree `/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2`
  (branched from `cshell`; pushed to origin). The last commit is a **WIP commit** containing two unfinished developer
  tasks (see "In progress"); everything before it is tested. `cshell` itself has not moved: when Stage 2 is done,
  `git checkout cshell && git merge --ff-only cshell-stage2` (Owen's call).
- Bulk data in projects storage, symlinked: `runs_cshell/`, `.cache_cshell/`, `continuous/research/tubeE_ccx/runs`.
- Env: `source continuous/cluster/env.sh` (conda env in projects; cd's to `$SLURM_SUBMIT_DIR`, so submit from the worktree
  and give tests absolute paths).

## Owen's rules (also in memory `cshell-fe-resource-rules`)
- **Nothing runs on the login node** — not tests, imports, `python -c`, mesh checks. Everything via sbatch. Pass this rule
  into every subagent brief.
- The optimiser's FE must use **minimum GPU and host memory**, be fast and accurate, on **one GPU** (multi-GPU only for
  validation solves). He declined the `--mem=240G` hybrid route for that reason.
- Slot cap 8 jobs (Isambard-AI). Minimum resources per job from a small timed run. Plan before implementing; developers
  work from an agreed `PLAN.md`.

## Decisions taken this session (Owen)
1. Skip the Stage 1 compliance-only part runs; go straight to full Stage 2 runs (3 parts × 3 starts).
2. **Build the two-grid PCG engine before the full runs** (`continuous/plans/PLAN_twogrid.md`).
3. Start the full runs once two-grid passes, **without** waiting for the body-fitted check; re-run `collect` later.
4. Body-fitted check (`continuous/plans/PLAN_validation_fixes.md`): GPU C3D10 + cuDSS solver as default (ccx optional),
   gmsh remesh + mesh-quality gate (no tetgen fallback), **relative** 5 % ratio tolerance, escape-hole film fix.
5. **K_σ excluded within `exclude_bc_mm` (2 mm) of clamped nodes** in voxel and body-fitted buckling (spec §4.4(c)).
6. FE memory option 1 adopted (done). V2-b tube E: investigated, no FE change (see VALIDATION.md).

## Stage 2 complete (2026-10-31)
All three parts accepted (STATUS §5, VALIDATION V6); usage guide `README_cshell.md`; `cshell` fast-forwarded to
`cshell-stage2` on GitHub. Open for Stage 3 (needs Owen's go-ahead, STATUS §2.5): scalable body-fitted solver
(iterative/multi-GPU) for full parts; export mass residual on passive mounts (+0.5 % before correction); geometry-chain
determinism (opt-in bitwise runs); V5-B1b load case; L-bracket capacity sign vs test.

## Decisions 2026-10-30 (Owen; overnight standing orders)
- Slot cap 24 for this session. Critical-path work only; ask before anything else.
- max_iter phase B 120 for the L-bracket (done: converged at 0.1662); fe.engine = twogrid default everywhere;
  eig_tol stays 1e-3; fe.twogrid.deterministic default "off" (opt-in "kernels"/"full"); V7 relaxed to final J ≤ 1e-3.
- Collect plan `plans/PLAN_collect_fixes.md` approved (mounts solid in the uniform-shell baseline; reject early when the
  export fails; one-sided 5 % ratio tolerance in §8.6). V5-B1b documented as is.
- Overnight: a part genuinely failing a criterion → accept and document (no criterion changes, no extra optimisation);
  a toolchain bug fix that changes optimiser results → re-run the 9 starts (runs_cshell/FULL configs) and collect;
  collect jobs may request up to ~460G (1 GPU) if measured need exceeds 230G; fast-forward `cshell` to
  `cshell-stage2` and push only if every Stage 2 gate item passes or is documented per these decisions.

## Done and committed (this session)
- Stage 2 code: `collect.py` (decision, re-scoring at opt/shifted/0.5 mm grids, export, body-fitted step, summary,
  sensitivity, fields.vtp, modes, bc faces), `report.py` (15 pages), CLI (`run/collect/report/setup/bench/faces/init`),
  vshell (S2) start, auto-resume + auto-collect cluster scripts (`continuous/cluster/README.md`, `launch_part.sh`).
- Geometry chain on GPU: phase-B chain 53 s → 0.6 s per iteration, identical J.
- FE memory: lean Krylov, node-level G, int32 indices, 14-RHS solves (MBB 0.4 mm peak 89 → 77 GiB, buckling 2× faster);
  **cuDSS memory leak fixed** (nvmath stream-ordered pool, `_release_device_memory`); factor mode chosen from cuDSS's
  plan-time estimate; hybrid only if host memory fits the cgroup limit.
- Concurrency: per-start `history_<start>.csv` merged atomically; `run_info.d/`.
- Validation: V2-c passes (0.4 within 0.04 % of 0.2 mm); V2-b tube E gap explained (voxel = CalculiX within 0.2 % on the
  same BCs; rollers + NLGEOM explain the trend, 2.5 % uniform residual vs Abaqus unexplained); B2 body-fitted diagnosis
  (`continuous/research/b2_bodyfit/diagnosis.md`).
- Studies: `continuous/research/gpu_memory/REPORT.md` (memory options), `research/tubeE*/`, `research/b2_bodyfit/`.

## In progress (in the WIP commit; check the running jobs' logs first)
| Task | Plan | Status note | Files |
| --- | --- | --- | --- |
| Two-grid PCG engine (matrix-free 0.4 mm, Galerkin 0.8 mm coarse cuDSS factor, Chebyshev smoother) | `plans/PLAN_twogrid.md` | `research/twogrid/STATUS_twogrid.md` (if the developer wrote it) | `cshell/fe_twogrid.py`, `fe_sparse.py`, `optimise.py` (`ensure_active`), `config.py`, tests |
| Validation fixes (GPU body-fitted, mesh gate, relative tolerance, hole films, K_σ exclusion body-fitted side) | `plans/PLAN_validation_fixes.md` | `research/b2_bodyfit/STATUS_validation_fixes.md` (if written) | `cshell/validate.py`, `export.py`, `collect.py`, `bodyfit_gpu.py`, tests |
| K_σ clamp mask, voxel side | spec §4.4(c) | **done and tested** (directional 4.0e-5, jobs 7808887/7808888); one line still needed in `ensure_active` to attach `fe.ksig_mask` (asked of the two-grid developer — check it is there) | `fe_sparse.buckling(ksig_mask)`, `PhaseModel._ksig_mask`, `tests/cshell/test_buckling.py` |

Jobs running at hand-over (results in `logs/` and the task dirs): 7809224, 7809226 (`tg_bench`), 7809227 (`tg_run3`),
7809375 (GPU tests), 7809124, 7808904 (CPU tests), 7809370 (`vf_b2_collect`, B2 collect with the new validation chain).

## Measured numbers that set the next steps
| Phase B 0.4 mm, full objective | DOF | Engine | FE s/iteration | Peak GPU |
| --- | --- | --- | --- | --- |
| Cantilever (7801144) | 6.1 M | fp32 factor + fp64 CG | 94 | 81 GiB |
| MBB (7801143) | 8.4 M | fp32 factor + fp64 CG | 129 | 94.8 GiB (limit) |
| MBB (7808291) | 8.4 M | hybrid fp64, 240 G host | 58 | 54 GiB |
| MBB solid (study) | 4.9 M | two-grid (prototype) | ≈ 23 (est.) | ≈ 20–30 GiB (est.) |
Two-grid targets (plan): MBB ≤ 30 s/iteration, ≤ 30 GiB GPU, default `--mem`; J to 1e-7 and dJ to 1e-5 vs direct.

## Remaining work, in order
1. **Finish two-grid** (`PLAN_twogrid.md` acceptance 1–4): read `STATUS_twogrid.md` and the tg_* logs; complete, test
   (CPU + GPU, all §6.6 gradient tests through both engines), benchmark on the voided cantilever/MBB designs
   (`runs_cshell/fix_cantilever`, `runs_cshell/fixhyb_mbb`), commit. Confirm `fe.ksig_mask` is attached in `ensure_active`.
2. **Launch the full runs**: 3 parts × 3 starts (uniform, vshell, neutral), one start per job via
   `continuous/cluster/launch_part.sh` / `run.sh` (auto-resume across 24 h, last start auto-submits collect). Size
   `--time/--mem` from step 1. Slot cap 8: 8 starts at once, the 9th after. Small smoke test first (`max_iter=[2,3]`).
3. **Finish the validation fixes** (`PLAN_validation_fixes.md`): read `STATUS_validation_fixes.md` and job 7809370; then
   re-run `collect` on B2 and on each finished part.
4. Then Stage 2 gate items: V5-B1b (stub + slender rod with fillet), V6 (per-part tables vs solid / uniform shell /
   vshell / manuscript sign agreement, §11.1), V7 (identical rerun to 1e-4), V1-c compliance half, reports for the three
   parts, `README_cshell.md`, `STATUS.md` §5 per-part table, spec updates (§8.3, §8.6, §4.2 engine).
5. Stage 3 (needs Owen's go-ahead): priority list in `STATUS.md` §2.5; 0.2 mm via two-grid (memory study: 25–35 M DOF in
   40–57 GiB) is now feasible.

## Open issues worth knowing
- Uncommitted-then-WIP files mix two developers' edits in `optimise.py`, `fe_sparse.py`, `config.py`, `collect.py`: review
  the WIP diff (`git show --stat HEAD`) before building on it; run the full test suite via sbatch.
- B2 body-fitted: λ_b near the clamp is mesh-dependent without the K_σ exclusion; with it, expect convergence — check.
- `examples/cshell/*.json` use relative `output_dir`; check where runs land.
- Two-grid scratch includes `research/twogrid/debug_nan.py`: a NaN was being chased — check `STATUS_twogrid.md` (may be missing if the developer stopped first) and the tg_* logs.
- STATUS.md log needs entries for 2026-10-29 (this hand-over summarises them).

## Late update (two-grid developer, after the WIP commit) — see `research/twogrid/STATUS_twogrid.md`
- MBB phase B (8.4 M DOF): two-grid FE 24.9 s/iteration vs 58 s direct hybrid; peak GPU 26.4 vs 52.8 GiB; host RSS
  29.6 vs 167 GB — targets met. Cantilever phase B 32 s vs 41.5 s, 19.6 vs 42.6 GiB. Phase A: same time, ⅓ memory.
- Accuracy: J ≤ 9e-9 vs direct on 3 of 4 checks; MBB design 3 differs by 3e-7 (plan: 1e-7) — suspected eigen-solver
  tolerance; jobs 7809224/7809226 (eig_tol 1e-5) decide it.
- Not done: dJ comparison, acceptance 3 (history of the MBB 3-iteration run 7809227 to compare), one CPU test hang
  (cause unknown), BENCHMARKS.md/notes.md, final test runs.
- Deviation: a coarse DOF is fixed if **any** fine DOF it interpolates to is clamped (plan said all; that made the MBB
  phase-A coarse matrix singular). `buckling()` now builds G via `fe._geometric_op(H)`; the `ksig_mask` line is in
  `ensure_active`.
