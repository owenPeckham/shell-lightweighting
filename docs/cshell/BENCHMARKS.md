> Renamed 2026-10-31: vshell → dshell (discrete optimiser, `dshell/`), continuous/ → docs/cshell/, shell_lightweighting/ → ushell/.

# cshell: measured costs (THEORY §4.2)

All on one GH200 (95.6 GiB HBM, Isambard-AI), fp64 cuDSS (nvmath-python 1.0, cuDSS 0.7) unless stated. Solid parts,
geometry and volume fractions from vshell's PartGrid (h_g = h_e/2). Scripts: `docs/cshell/research/stage0/`.
Raw numbers: `docs/cshell/research/stage0/bench_*.json`, logs `logs/cshell_bench0_7794574_*.out`.

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

## Two-grid engine (`cshell/fe_twogrid.py`, PLAN_twogrid.md; default `fe.engine = auto`)
Matrix-free q1_im K at h_e (CUDA kernel), Galerkin coarse level at 2h_e (cuDSS fp32 factor), Chebyshev-Jacobi
(degree 5, range 30), block PCG fp64 tol 1e-10; buckling inner solves to min(1e-2, 0.1·res). `auto`: direct fp64
when ≤ 2 M DOF and the fp64 factor fits in device memory (phase A), else two-grid (phase B). Reference "direct" =
hybrid fp64 cuDSS (`--mem=240G`). Voided designs of `fixhyb_mbb` / `fix_cantilever`, full objective, eig_tol 1e-3.

| Phase B 0.4 mm, full objective | DOF | Engine | FE s / evaluation | Peak GPU | Peak host RSS |
| --- | --- | --- | --- | --- | --- |
| MBB (bench 7808957 / 7809111) | 8.36 M | hybrid fp64 / two-grid | 58 / 24.9 | 52.8 / 26.4 GiB | 167 / 29.6 GB |
| Cantilever (7808959 / 7809112) | 6.13 M | hybrid fp64 / two-grid | 41.5 / 32 | 42.6 / 19.6 GiB | 117 / 23 GB |
| MBB 3-iteration run (7856981 / 7809227) | 8.36 M | hybrid fp64 / auto | 51–137 / 25–50 | 50.1 / 25.9 GiB | 161 / 34 GB (job) |
| Cantilever 3-iteration run (7856982 / 7856983) | 6.13 M | hybrid fp64 / auto | 46–108 / 33–34 | 82.9 / 19.6 GiB | 111 / 26 GB (job) |
| Phase A (MBB / cantilever, two-grid) | 1.3 / 0.9 M | two-grid vs direct | 7.7 vs 7.4 / 5.8 vs 5.5 | 5.0 vs 15.0 / 4.0 vs 11.0 GiB | |
PCG iterations: 40–53 per exact (primal/adjoint) solve, ≈ 5 per buckling inner solve (≈ 19 at eig_tol 1e-5).
Wall time per 3-iteration run: MBB 5:17 (two-grid) vs 9:23 (hybrid); cantilever 4:39 vs 6:22.

Accuracy vs direct (`research/twogrid/notes.md`): compliance-and-stress dJ (buckling off) ≤ 5e-10 rel 2-norm, J ≤ 7e-11
(7857876, MBB phase B 10 s vs 99 s); full objective J ≤ 8e-9 on 5 of 6 designs, 3e-7 on MBB ckpt 3; full dJ 6e-5–1.6e-3,
set by the eigen-solver tolerance, not the engine; 3-iteration run histories J ≤ 5.7e-8 (MBB) and ≤ 2.7e-8 (cantilever).

### Two-grid determinism, opt-in `fe.twogrid.deterministic` (default `off` = the atomic kernels; P6; `research/twogrid/notes.md`, jobs 7868216, 7869145-8, 7870417-9, 7870887)
MBB phase B, voided design, full objective, FE seconds per evaluation (ckpt 2 / 3) and peak GPU: atomic kernels (before)
47.0 / 24.8 s, 25.8 GiB; colour-ordered kernels + default cuDSS (`kernels`) 54.3 / 28.3 s (+14-16 %), 25.4 GiB; plus cuDSS
deterministic mode (`full`, default) 124 / 63.8 s (+150 %, cuDSS solves one right-hand side at a time: 9.4 vs 0.73 ms/rhs at
561 k coarse DOF), 26.6 GiB. Final code, default `off` (7875876): 48.3 / 24.8 s, 25.9 GiB (= before). FE-only re-evaluation of the same ρ is bitwise equal in `full` only. J/dJ vs direct unchanged.
