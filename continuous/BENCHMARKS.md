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
