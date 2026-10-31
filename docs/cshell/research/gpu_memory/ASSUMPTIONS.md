# GPU-memory study: inputs, units, assumptions

- Engine under test: `cshell/fe_sparse.py` at branch `cshell-stage2` (imported read-only; never edited). Prototypes
  live here: `gpumem.py` (memory/precision/eigensolver/hybrid phases), `twogrid.py` (two-grid PCG, matrix-free
  operator), `int64_test.py`, `job.sh` (sbatch wrapper: `sbatch [--export=ALL,SCRIPT=twogrid.py] job.sh <part> <tag> <h> [args]`).
- Parts: `examples/dshell/cantilever.json`, `examples/dshell/mbb.json` (Study 1 solids, same geometry/BC/load path as
  `docs/cshell/research/stage0/bench_solid.py`: vshell PartGrid at h/2 with k_fe=2, `fea.build_bcs`).
  Material and loads from those configs (E in MPa, forces in N, lengths in mm, compliance in N mm).
- rho_hat = 1e-6 + (1 - 1e-6) * volume fraction (solid design, fraction-weighted boundary voxels), as in Stage 0.
- "Hollow" proxy (`--hollow k`): the solid's active mask minus its k-fold binary erosion (a skin k voxels thick;
  k=4 at 0.4 mm = 1.6 mm wall). Not an optimised design; it stands in for the thin-wall fill pattern.
  The real hollow design (runs_cshell/timing/Cantilever) was not used: rho_hat from xi needs the optimise pipeline
  that Task G is editing.
- Active set: material elements only (no void margin). Production adds a void margin (dilation), which enlarges
  the DOF count; per-DOF figures are given so they can be scaled.
- Memory: GiB = 2^30 B. `device_used` = total - free from cudaMemGetInfo (includes CUDA context, CuPy pool cache and
  cuDSS); `cupy_used`/`cupy_held` from the CuPy pool; cuDSS = device_used - cupy_held - context. cuDSS plan-time
  estimates from `plan_info.memory_estimates`.
- Buckling settings as production (`cshell/config.py`): 6 modes, p = 8, eig_tol = 1e-3 (dual-norm residual;
  eigenvalue error ~ tol^2), block m = 10 (6 + 4 extra), 4 negative deflation vectors, 3 Krylov steps.
- Accuracy references: the fp64 cuDSS factor (direct) run in the same job. Relative errors are 2-norm relative.
- GPU: one GH200 (95.6 GiB usable), one socket's host memory (~115 GB) unless stated.
