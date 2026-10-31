# Two-grid engine: acceptance notes (PLAN_twogrid.md), 2026-10-30

Numbers and job ids: `continuous/BENCHMARKS.md` "Two-grid engine". Development log: `STATUS_twogrid.md`.

## Acceptance
1. **Tests.** CPU (7859087: test_objective, test_buckling, test_fe_twogrid, test_fe_sparse) 22 passed, 17 GPU-only
   skipped; GPU (7859529: same files + test_end_to_end) 45 passed. Gradient tests (compliance FD, stress adjoint notched
   bar, buckling directional, J directional, end-to-end) are parametrised over both engines.
2. **Voided phase-B designs vs hybrid fp64 direct.** J ≤ 8e-9 on 5 of 6 designs; MBB ckpt 3: 3.0e-7 (plan 1e-7).
   dJ (2-norm rel): 6e-5 – 1.6e-3 with buckling (plan 1e-5); **≤ 5e-10 with buckling off** (7857876, both phase-0
   designs and the MBB phase-B design). Time and memory targets: MBB 24.9 s / 26.4 GiB on the bench design
   (≤ 30 s, ≤ 30 GiB); 25–50 s per iteration in the 3-iteration run (eigen-iteration count varies).
3. **3-iteration runs** reproduce the direct history: J ≤ 5.7e-8 (MBB), ≤ 2.7e-8 (cantilever); plan 1e-6.
4. This note and BENCHMARKS.md.

## Why the full-objective J/dJ gaps are not engine error
The block-Krylov eigen-solver stops at a dual-norm residual `eig_tol` = 1e-3: eigenvalues are accurate to ≈ tol²
(λ_b gaps 1e-9 – 7e-7), eigenvectors only to ≈ tol/gap. The buckling sensitivity is built from the eigenvectors, so
two correct linear solvers that take different Krylov paths (two-grid inner solves are inexact by design, tol
min(1e-2, 0.1·res)) give dJ differing at O(1e-4). Evidence: (a) with the buckling term off, dJ agrees to 5e-10;
(b) at eig_tol 1e-5 (7857131) neither engine converges in 100 outer iterations — the direct engine included — and dJ
still differs by 7–9e-5, so the 1e-5 dJ criterion is not attainable through this eigen-solver with either engine;
(c) MBB ckpt 3's J gap (3e-7) comes with 2 vs 3 outer eigen iterations. The buckling directional FD test (< 1e-3)
passes through the two-grid engine, so its eigen-gradient is consistent.

## Fixes made while finishing (2026-10-30)
- `optimise._build_engine`: `fe.backend = "cpu_direct"` was not recognised as CPU, so `auto` fell through to two-grid
  on CPU (slow tests; the CPU end-to-end test's SuperLU "exactly singular" coarse factor came from this route and did
  not reproduce with two-grid forced on B2 at h_e 0.8 on either backend, run + collect: 7858463/7858973/7858729).
- `TwoGridFE`: CPU backend assembles K as SciPy CSR per factor (the NumPy element-loop einsum made CPU buckling tests
  run > 10 min: the "CPU hang"); the GPU stays matrix-free. G's CPU apply uses batched matmul.
- `TwoGridFE`: connectivity copied to the host when the grid was built on the GPU but the FE runs on the CPU.

## Open
- Eigen-solver: eig_tol 1e-5 is not reached in 100 outer iterations on the voided MBB/cantilever designs by either
  engine. Not needed at the production tolerance (1e-3); worth a look before tightening it.
- Decided (Owen, 2026-10-30): `fe.engine = "twogrid"` is the default in every phase (phase A: same time as direct
  at 1/3 the GPU memory); `fe.backend = "cpu_direct"` still forces the SciPy direct solve. eig_tol stays 1e-3 (the
  ~1e-4 dJ noise is acceptable).

## Optional deterministic modes (P6 of PLAN_collect_fixes.md) and PCG stagnation (P4 ii), 2026-10-30
**Decision (Owen, 2026-10-30): default `off`** (the original atomic kernels, fastest); V7 relaxed to final J within 1e-3;
design.py determinism not pursued. Modes are opt-in through config `fe.twogrid.deterministic` (passed by
`optimise._build_engine`; `TwoGridFE(deterministic=...)`, env `CSHELL_TWOGRID_DET` when the argument is None):
`off` | `kernels` | `full`.
- Sources of run-to-run noise found (`det_probe.py`, 7868031; each stage run 4x, max |diff| to run 1): K apply 3.6e-12,
  G apply 1.4e-14, Jacobi diagonal 1.4e-12, coarse Galerkin values 2.7e-12 (all float64 atomics: now colour-ordered,
  0), cuDSS coarse factor+solve (fp32: O(1e-2) absolute; fp64 3e-11; the solve alone on one factor also differs).
  Bitwise already: cuSPARSE SpMM (P, Pᵀ), cuBLAS dot/gemm, cuSOLVER eigh, cupy reductions, Lanczos. Also atomic:
  `buckling()` adjoint load and `objective.stress_pnorm` adjoint rhs (both now `fe._scatter_dofs`, same result for SparseFE).
- Fix: 8 parity colours of (i, j, k) (fine and coarse elements); elements of one colour share no node, so plain adds, one
  launch per colour, fixed order. cuDSS: `CUDSS_CONFIG_DETERMINISTIC_MODE` (set through nvmath's `config_ptr`); it only
  supports ONE right-hand side per solve, so the coarse solve loops over the PCG block's columns (`det_cudss.py`:
  27k coarse DOF 0.11 -> 1.07 ms/rhs, 561k DOF 0.73 -> 9.4 ms/rhs fp32; 7868212, 7869144).
- Result on the real MBB voided design (bench.sh, ckpts 2 / 3, per evaluation; 7868216 / 7869145-8):

  | mode | t_fe ckpt 2 / 3 | peak GPU | run-to-run J |
  | --- | --- | --- | --- |
  | off (before) | 47.0 / 24.8 s | 25.8 GiB | ~1e-9 |
  | kernels | 54.3 / 28.3 s (+14-16 %) | 25.4 GiB | 7e-11 (cuDSS) |
  | full | 124 / 63.8 s (+150 %) | 26.6 GiB | 9e-13 (design chain, below) |

  FE alone, same rho evaluated twice (`bench_design.py --repeat`, 7870887, 7870417): J, C, λ_s, λ_b, μ̃ and dJ
  bitwise equal in `full` (not in `kernels`/`off`: 7870418/9). J/dJ vs direct unchanged: dJ without buckling 3.0e-10 /
  5.2e-10 / 1.8e-10 (7872172, plan <= 1e-9); full-objective J 1.0e-8 / 3.0e-7, dJ 1.2e-4 / 1.6e-3 (as before; eig-solver
  tolerance, see above).
- **Not bitwise end to end: `design.py` is still non-deterministic** (`geo_fwd`: `atomicAdd(rho+elem[c], H)` and the
  block-sum `atomicAdd(vol, ...)`; `geo_rev`: `atomicAdd(out+v[i], ...)`, plus the rescore kernels). `rho_sha` of the
  same design differs between runs. V7 cantilever uniform, `fe.engine=twogrid` in both phases (7870901/3 full, 7870904/5
  kernels): phase-0 J at it 59 differs already (0.52059 vs 0.52040; volume error -5.4e-5 vs -2.0e-5), final J
  0.519243 vs 0.518932 (3.1e-4, full) and 0.518998 vs 0.519191 (1.9e-4, kernels): FAIL (< 1e-4 needed). The FE is
  no longer the source; fixing needs fixed-order reductions in design.py (not a file of this work).
- P4(ii): `maxiter` is honoured (ctor arg, `fe.twogrid.maxiter`). A column whose recursive residual falls < 2x over
  `stall_window` = 20 iterations and is <= `stall_accept` = 1e-7 relative is accepted with a printed warning; counted in
  `stats["pcg_stall_accepted"/"pcg_stall_max_rel"]` and `pcg_summary()["stalled"/"stall_max_rel"]`; otherwise it raises at
  maxiter as before. Test: `test_stagnation_accepted_with_warning` (Jacobi preconditioner, window 2).
- Tests: GPU 7870886 (test_fe_twogrid, test_fe_sparse, test_buckling, test_objective, test_end_to_end) 48 passed; CPU
  7870422 (first four) 23 passed, 7870895 test_objective after the objective.py edit 5 passed.
- Final code, default `off` (2026-10-30): bench MBB phase B ckpts 2 / 3 7875876: 48.3 / 24.8 s, 25.9 GiB peak (pre-P6:
  47.0 / 24.8 s, 25.8 GiB: unchanged). GPU tests 7875874 (test_fe_twogrid, test_fe_sparse, test_buckling,
  test_objective, test_end_to_end) 48 passed; CPU 7875875 (first four) 23 passed. V7 rerun with the V7 tolerance 1e-3:
  default-mode final J differences are 2e-4 to 3e-4 in the runs above (7870901/3/4/5), so within 1e-3.
