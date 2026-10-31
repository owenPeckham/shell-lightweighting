# Two-grid FE engine: status 2026-10-29 (PLAN_twogrid.md; nothing committed)

**Implemented**
- `cshell/fe_twogrid.py` (new): `LeanGrid` (connectivity built on the GPU, no CSR pattern) and `TwoGridFE(SparseFE)` with the same interface. It has a matrix-free K (CUDA kernel), matrix-free G (`_geometric_op`), a Galerkin coarse level at 2h assembled on the GPU (`_BlockPattern`), and a cuDSS coarse factor (fp32, or fp64 if the fp32 probe solve is non-finite). The smoother is Chebyshev-Jacobi (degree 5, range 30, Lanczos λmax), and PCG is fp64 Polak-Ribière, block-wise, with fused kernels. The primal warm-starts. `solve_inner` serves the buckling Krylov steps (tol min(1e-2, 0.1·res)).
- `cshell/fe_sparse.py`: `SparseFE._geometric_op(H)` (moved out of `buckling()`, same behaviour); `_block_krylov_largest` uses `fe.solve_inner` when it exists.
- `cshell/optimise.py`: `_build_engine` handles fe.engine auto/direct/twogrid. "auto" = direct fp64 if DOF ≤ 2 M and the trial factor is fp64 in device memory, else twogrid. It sets `fe.ksig_mask` and adds a `fe_engine` history column.
- `cshell/config.py`: `fe.engine` "auto", `auto_direct_max_dof`, `fe.twogrid{...}`.
- Tests: `tests/cshell/test_fe_twogrid.py` (new); twogrid parametrisation added to test_fe_sparse (FD), test_buckling (directional), test_objective (stress adjoint, J) and test_end_to_end.
- Research scripts in this directory: `bench_design.py`, `bench.sh`, `exp_pcg.py`, `run3.sh`, `debug_nan.py`.

**Deviation from the plan.** A coarse DOF is fixed if *any* fine DOF it interpolates to is clamped. The plan's rule (fixed only if *all* are clamped) gave a singular K_c on MBB phase A (NaN; job 7808953).

**Results.** Phase B uses ckpts 2 and 3; direct = hybrid fp64 with 240G; both engines have the ksig mask.
- GPU tests 7808907: 17 passed. That run used the old coarse rule, so it must be re-run.
- MBB 8.36 M DOF (direct 7808957, two-grid 7809111):
  - J rel diff 8e-9 (d2) and 3.0e-7 (d3). d3 is outside the 1e-7 criterion. Cause: the eigen-solver stopped after 2 outer iterations vs 3 for direct; eig_tol 1e-3 implies a μ error of about 1e-6.
  - FE time 24.9 s vs 58 s direct. Peak GPU 26.4 vs 52.8 GiB. Host RSS 29.6 (about 28 of it geometry set-up) vs 167 GB.
  - PCG iterations: exact solves 40–53, Krylov solves about 5.
- Cantilever 6.13 M DOF (direct 7808959, two-grid 7809112): J diff ≤ 3e-9; 32 s vs 41.5 s; 19.6 vs 42.6 GiB; RSS 23 vs 117 GB.
- Phase A (two-grid vs direct): MBB 7.7 s / 5.0 GiB vs 7.4 s / 15.0 GiB; cantilever 5.8 s / 4.0 GiB vs 5.5 s / 11.0 GiB.
- Smoother and coarse study (exp_pcg, 7808792): fp64 coarse gives no gain. Degree 5 and range 30 take primal iterations from 78 to 53. nrhs 7 beats 14 on both memory and time (7808882).
- Run3 MBB auto, 7809227: COMPLETED. Peak 26.5 GiB, RSS 35 GB, 5:17 wall. History in `runs_cshell/TG_run3_auto_mbb/MBB/history.csv`, not yet compared.

**Still running**
- 7809124: CPU tests (fe_twogrid, buckling, objective) with `-v --durations`. A previous CPU run hung for more than 14 min in test_buckling; the cause is not yet known.
- 7809224 / 7809226: MBB phase B at eig_tol 1e-5, direct vs two-grid. This shows whether the d3 J gap closes when the eigen-solve is tighter.

**Unfinished**
- dJ 2-norm comparison not yet run: `python bench_design.py --compare runs_cshell/TG_bench/mbb_directM_ph1.npz runs_cshell/TG_bench/mbb_tgG_ph1.npz` (CPU job; same for `cant_*`).
- Acceptance 3 still needs the direct runs (`PART=mbb|cantilever ENG=direct`, `--mem=240G`) and `cantilever ENG=auto`, then a comparison of history J to 1e-6.
- Full CPU and GPU test suites, including end-to-end, must be re-run on the final code.
- The CPU hang must be found and fixed.
- `continuous/BENCHMARKS.md` and `notes.md` are not written.

**Next steps**
1. Read 7809124, 7809224 and 7809226.
2. Run the compare jobs for dJ.
3. Run the remaining run3 jobs.
4. Run `gpu_test.sh pytest tests/cshell` and `cpu_test.sh`.
5. Write BENCHMARKS and notes, then hand the work to the orchestrator for commit.
