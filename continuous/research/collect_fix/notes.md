# collect fixes (PLAN_collect_fixes P1, P2, P3, P4(i), P5, P7, D1-D3) - developer notes, 2026-10-30

Scripts (submit from the repo root):
- `memprof_cf.sh` / `memprof_cf.py`: per-stage host memory of collect (cgroup anon/current, RSS, PSS) on a copy of the
  MBB smoke with the SMOKE_mbb STLs re-used (`--skip-export`) and the best forced through the export gate
  (`--force-ok`, memory/PCG only, not a decision). Job 7870471.
- `closure_diag.py`: where the export closure rule still adds volume over the optimiser (job 7870594).
- `chain_check.py`: GPU vs CPU re-scoring chain on the MBB smoke opt grid (job 7872205).
- `remesh_check.py`: B2 body-fitted surface remesh at h = 0.5, loader x STL, twice each (job 7870945).

## Results
| item | measured | target |
|---|---|---|
| P1 MBB smoke uniform shell K/K0 opt / shift / coarse | 0.7750 / 0.7726 / 0.7826 (opt-shift 0.3 %) | [0.6, 0.8], <= 5 % |
| P1 isolated loaded nodes, every re-scored row (smoke, B2, memprof) | 0 | 0 |
| P2 band volume, closure rule vs optimiser: smoke best | +0.505 % (was +3.79 %) | 0.2 % (NOT met) |
| P2 B2 | +0.004 %; export V before correction 0.4591 % vs 0.4606 % | 0.2 % |
| P3/D2 smoke decision | "export failed: mesh not watertight (39 bodies, 372 duplicated directed edges)"; mass now matched (delta -0.078 mm, V_STL/V0-1 +1.4e-4) | export failure is the reason |
| P4(i) smoke all 3 resolutions incl. best (old 39-body STL, forced) | completed, PCG max 53 it at 1e-10 (shift best had stalled at 100 it before) | completes |
| P5 peak host outside export (memprof) | cgroup current 40.9 GiB, anon 29.8 GiB, RSS 39.5 GiB (PhaseModel build) | <= 60 GiB |
| P5 RSS returns after each model release | +1.7 / -0.3 / -0.5 GiB | within 5 GiB |
| P5 GPU vs CPU chain, MBB opt grid | max drho 9e-16, dV 1.5e-15, dC 1.4e-14, dJ 5.0e-13 | dJ <= 1e-10 |
| P5 full collect MBB smoke (job 7868888) | 35 min; parent peak RSS 51.8 GiB, Slurm summed MaxRSS 152 GiB (upper bound, CoW pages counted per worker) | fits 180G |
| P7 B2 export, walls below t_min - b_v - 2.5 h_x | 0 (min as built 1.203 mm) | 0 |

P2 residual (closure_diag): 0.46 % of V0 in mount cells and 0.05 % in thin-passive cells deeper than the
interpolated a^hi (r_MA under-estimates the depth there); the optimiser treats these cells as void, the export keeps
them solid. No active vertex is closed. The mass correction absorbs it (delta -0.078 mm).

B2 body-fitted: pymeshlab's isotropic remesh at h = 0.5 passes or fails its edge check depending only on the vertex
order of the same surface (old STL: trimesh-welded passes, exact-weld fails, deterministic over repeats). The coarse
level only feeds the non-blocking convergence flag, so collect now records its failure and compares the finer level.

B2 collect (job 7875773, --mem=110G; 64G OOM-killed in the export, job 7871594): accepted, J >= +1.1572,
K/K0 >= 3.181, P/P0 >= 2.184; Pb/Pb0 voxel under-predicts by 6.3 % (logged, D3); solid voxel C 77.8394673999612 vs
77.83946739986106 in VF_B2 (1.3e-12: converged rows unchanged). Slurm summed MaxRSS 94 GiB (72 export workers).
Tests: CPU 7872207 (43 passed, 4 skipped), GPU 7872208 (47 passed).
