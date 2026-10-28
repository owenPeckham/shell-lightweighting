# cshell: measured costs (THEORY §4.2)

All on one GH200 (95.6 GiB HBM, Isambard-AI), fp64 cuDSS (nvmath-python 1.0, cuDSS 0.7) unless stated. Solid parts,
geometry and volume fractions from vshell's PartGrid (h_g = h_e/2). Scripts: `continuous/research/stage0/`.
Raw numbers: `continuous/research/stage0/bench_*.json`, logs `logs/cshell_bench0_7794574_*.out`.

## Elastic solve, sparse active-set engine vs dense reference (h_e = 0.8 mm)

| Part | Sparse DOF | Dense-box DOF | Plan (s) | First factor (s) | Refactor (s) | Solve (s) | GPU (GiB) | K₀ sparse | K₀ dense ShellCG | K₀ vshell | Dense solve (s, PCG its) |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Cantilever | 0.53 M | 22.2 M | 2.4 | 2.7 | 0.29 | 0.34* | 5.1 | 257.8 | 263.9 | 256.8 | 7.3 (84) |
| L-bracket | 0.36 M | 13.1 M | 1.7 | 1.9 | 0.17 | 0.09 | 3.9 | 48.60 | 51.87 | 48.3 | 8.9 (45) |
| MBB | 0.72 M | 13.0 M | 3.0 | 3.4 | 0.35 | 0.08 | 6.9 | 473.5 | 482.6 | 473 | 6.6 (48) |

\* first solve includes one-off set-up. Dense `ShellCG` (reference GMG-PCG, float32, rel_tol 1e-6) reads 2–7 % stiffer
than the direct solve and vshell: OPEN_QUESTIONS B5.

## h_e = 0.4 mm (production resolution)

| Part | Sparse DOF | Active elements | nnz | Plan (s) | First factor (s) | Refactor (s) | Solve (s) | GPU after factor (GiB) | GPU with buckling (GiB)† | K₀ |
| --- | --- | --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Cantilever | 3.52 M | 1.05 M | 265 M | 17.6 | 21.9 | 4.3 | 0.038 | 54.0 | 66.1 | 252.3 |
| L-bracket | 2.41 M | 0.72 M | 182 M | 11.4 | 13.9 | 2.4 | 0.026 | 38.9 | 48.2 | 47.90 |
| MBB | 4.86 M | 1.46 M | 368 M | 24.8 | 32.8 | 7.8 | 0.056 | 77.0 | 93.9 | 467.2 |

† with the first (unconverged, 200-iteration) eigensolver; to be re-measured with the current one.

## Buckling (block-Krylov on the cuDSS factor, 6 modes, block 10, dual-norm tol 1e-3)

| Part | h_e | Outer iterations | Time (s) | λ₁ | P_cr = λ₁‖F‖ (N) |
| --- | --- | --- | --- | --- | --- |
| Cantilever | 0.8 | 3 | 2.0 | 3.364 | 673 |
| Cantilever | 0.4 | (re-measure) | | 3.287 | 657 |
| L-bracket | 0.4 | (re-measure) | | 9.942 | 1988 |
| MBB | 0.4 | (re-measure) | | 7.341 | 1468 |

## Geometry band (Warp, `cshell/band.py`; cantilever, D_comp = 4 mm) — `research/band_bench.md`

| h_g (mm) | Cells | Build (s) | Peak host RAM (GiB) | V₀ vs B-rep |
| --- | --- | --- | --- | --- |
| 0.4 | 3.7 M | 1.2 | 0.6 | +1.1e-3 |
| 0.2 | 28.2 M | 3.6 | 1.9 | +2.6e-4 |
| 0.1 | 220 M | 25.8 | 9.6 | +5e-5 |

## Per-iteration estimate at 0.4 mm (to be replaced by measured optimiser iterations)

Refactor 2.4–7.8 s + primal and stress-adjoint solves (< 0.2 s) + buckling (≈ 3 outer iterations × ~45 block solves) +
geometry (band ψ/χ on GPU, seconds). Expected ≈ 10–20 s per iteration, i.e. phase B (80 iterations) ≈ 15–30 min per start.

## FE memory option 1 (lean eigensolver, node-level G, int32 indices, 14-RHS solves; `research/gpu_memory/REPORT.md`)

MBB solid, h_e = 0.4 mm, 4.86 M DOF, same script and job type as the table above (job 7800841 vs 7794574-era run
`bench_mbb_04c.json`; tests 7800843 GPU 21 passed, 7800844 CPU 14 passed).

| Quantity | Before | Option 1 |
| --- | --- | --- |
| GPU used after factor (GiB) | 71.5 | 67.1 |
| GPU used after buckling (GiB) | 86.5 | 77.1 |
| Peak GPU (nvidia-smi, 0.5 s sampling, GiB) | ≈ 89 (study) | 77.4 |
| Buckling (s, 4 outer iterations) | 11.5 | 5.7 |
| Single-RHS solve (s) | 0.13 | 0.27 (solves the 14-column block) |
| Refactor (s) | 7.8 | 7.8 |
| Compliance | 85.617328601565 | 85.617328601560 |
| μ (6 modes), max relative difference | — | 3.4e-12 |

## Geometry chain on GPU (`design.DeviceDesignMap`, Task G; runs `runs_cshell/G_timing/*`)

The band chain (ψ → χ → ρ, V, ∇V, Pᵀ pull-backs, restoration volumes) runs as fused CuPy kernels over a compact band
(φ f32, closest Γ_h triangle i32, raw barycentrics 3×f32, element i32 = 24 B/cell; nothing per cell is stored between
passes). Per-vertex work (filter, D1 bounds, sliver) stays on the host. The band is held in pinned host memory and
released from the device during every FE evaluation (`grid.geometry_offload`, default on), so the FE sees the same free
memory as with the CPU chain. If the re-upload does not fit, the kernels read the pinned copy in place (zero-copy) for
the rest of the phase, with a log line. Config: `grid.geometry_device` = auto | cpu | gpu.

Phase B (h_e 0.4, h_g 0.1), seconds per iteration (median of the iterations shown), same code and FE, CPU vs GPU chain.
J agrees to ≤ 3e-10 in every iteration (all printed digits in the runs listed).

| Part (objective) | Band cells / device footprint | Chain | t_geometry | t_pull | t_mmasub | t_restore (5 V evals) | chain total | t_fe |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| Cantilever (full, buckling) jobs 7800362 / 7800683 | 220 M / 4.93 GiB | CPU | 5.24 | 13.67 | 0.43 | 33.7 | 53 | 95 |
| | | GPU | 0.03 | 0.04 | 0.44 | 0.07 | 0.58 (+0.01 upload) | 95 |
| MBB (compliance, fp32 factor) jobs 7800935 / 7800934 | 282 M / 6.30 GiB | CPU | 7.8 | 19.3 | 0.44 | 23.9 | 51 | 16 |
| | | GPU | 0.02 | 0.07 | 0.45 | 0.04 | 0.58 (+0.03 upload) | 16 |

Phase A (h_g 0.2, 28–36 M cells): chain 7.3 → 0.6 s (cantilever), 5.9 → 0.6 s (MBB). What is left of `t_mma` is
`mmasub` itself (0.45 s, mmapy on ~1.3·10⁵ variables, CPU). The phase-B iteration is now the FE (16–95 s).

GPU memory (driver view): the band costs 0 GiB during the FE and 4.9 / 6.3 GiB between FE evaluations; re-upload
0.01–0.03 s from pinned memory. Cantilever full objective peaks at 87 GiB (fp32-preconditioned factor, buckling).

**MBB with buckling at 0.4 mm does not fit (K11), independently of the geometry chain** (band off the device during the
FE): job 7800463 (6 modes, fp32 factor) OOM in the Krylov orthonormalisation; job 7800764 (4 modes) OOM in buckling
after the active-set rebuild. Cause: device memory of a freed cuDSS DirectSolver is not returned (measured: free memory
85.8 → 41.5 GiB across one rebuild at 6.3 → 8.4 M DOF, MBB job 7800934; 80.9 → 51.3 GiB cantilever), so the second
factor plus buckling no longer fits. With `fe.precision = auto` the new fe_sparse chooses fp64 hybrid at 6.3 M DOF,
which spills the factor to host memory and was killed by the 115 GB job memory limit (jobs 7800876/7800877, both
chains); the MBB timings above therefore use `fe.precision = fp32`.
