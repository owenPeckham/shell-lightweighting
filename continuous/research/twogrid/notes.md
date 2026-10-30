# Two-grid engine: acceptance notes (PLAN_twogrid.md), 2026-10-30

Numbers and job ids: `continuous/BENCHMARKS.md` "Two-grid engine". Development log: `STATUS_twogrid.md`.

## Acceptance
1. **Tests.** CPU (7859087: test_objective, test_buckling, test_fe_twogrid, test_fe_sparse) 22 passed, 17 GPU-only
   skipped; GPU and remaining CPU suites: see the commit message. Gradient tests (compliance FD, stress adjoint notched
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
- Phase A stays on direct fp64 under `auto` (same time, 3× the GPU memory of two-grid: 15 vs 5 GiB MBB). Owen's call
  whether to use two-grid everywhere (`fe.engine = twogrid`).
