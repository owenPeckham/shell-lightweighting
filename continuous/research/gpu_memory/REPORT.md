# GPU memory of the FE engine: options measured (2026-10-28)

One GH200 (95.6 GiB), `cshell/fe_sparse.py` unchanged; (E) = estimate. Scripts `gpumem.py`, `twogrid.py`, `compare.py`;
raw results `*.json`; assumptions `ASSUMPTIONS.md`. Jobs 7800356, 7800358, 7800389, 7800391, 7800400, 7800403, 7800431,
7800570, 7800682.

**Today, MBB at 0.4 mm (4.86 M DOF): peak ≈ 89 GiB** = fp64 factor 58.6 + engine arrays 10.6 + buckling 18.9.

| Option | GiB saved (MBB) | Accuracy vs fp64 | Solve / buckling time vs today (0.052 s / 11.6 s) | Effort |
| --- | --- | --- | --- | --- |
| 1a Lean block-Krylov storage | 5.5 | μ 2e-12 | same | small |
| 1b Engine trims (`_keep`, `_perm` int32, G per node) | ~9 (E) | exact | same | small |
| 1c Multi-RHS cuDSS solves (14 columns) | 0 | exact | 0.18 s per 14 columns vs 0.73 s; buckling ~4 s (E) | small |
| 2 Hybrid memory, device cap 20 GiB | 37 | C 1e-13 | refactor 11.5 vs 7.8 s; solve 0.38 s; buckling ~13 s with 1c (E) | small |
| 3 fp32 factor + fp64 PCG | 24.6 | C 3e-11, dC 2e-10, λ 1.5e-8 | solve 0.6–0.8 s; buckling 34 s | small |
| 4 Two-grid PCG (Galerkin coarse 0.8 mm factor) | 55 (~65 matrix-free, E) | λ ≤ 4e-8, dμ ≤ 1.1e-4 | solve 0.45–0.51 s; buckling 22 s; no fine factor; ≈ 23 s/iteration vs 19.6 (E) | moderate–large |

Failed or no gain: fp32 factor used directly inside the Krylov steps (μ off 5–8 %); AMD reordering (factor 2.5× larger);
ND ≈ default.

**Recommendations.**
(a) Production 0.4 mm: implement 1a–1c (MBB peak ≈ 74 GiB (E), buckling ≈ 3× faster). Choose the factor mode from cuDSS's
plan-time memory estimate (it matched the measured factor), not a kB/DOF constant; if fp64 does not fit, hybrid, then fp32.
(b) 0.2 mm: two-grid PCG with a matrix-free fine level — cantilever 25.5 M DOF, 25 iterations, 2.35 s, 40 GiB; MBB 35.4 M
DOF, 24 iterations, 3.0 s, 57 GiB. Buckling at 0.2 mm not run (2–4 min (E)).

GPU headroom left for the geometry chain (Task G): ≈ 7 GiB today, ≈ 22 with option 1, ≈ 59 with 1 + hybrid, ≈ 70 two-grid.

Open: device memory not returned after `DirectSolver.free()`; two-grid not yet tried on an optimised (voided) design;
superpanels off saves 3.4 GiB on paper, speed untimed.
