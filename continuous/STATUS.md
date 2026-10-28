# cshell: status, plan and risks

Stage log for `cshell` (THEORY_continuous.md §12). Newest entries at the top of §0.

## 0. Log

| Date | Stage | Entry |
| --- | --- | --- |
| 2026-10-28 | 0.2 | FE-engine research done (`continuous/research/fe_engine/REPORT.md`). Measured on one GH200, thin-wall Q1 hex, fp64 cuDSS Cholesky: 3.3 M DOF factor 1.8 s, solve 0.024 s, 33 GiB; 7.6 M DOF factor 5.5 s, 75 GiB; 13.6 M DOF does not fit (≈ 9 kB/DOF). Symbolic `plan()` 13–32 s on CPU, only when the sparsity pattern changes. nvmath-python/cuDSS available (pip aarch64, and the `sa-gpu` container). Recommendation put to Owen: material-only active-set grid + assembled q1_im + cuDSS (§2.2, 0.3). `ShellCG` adapter and its tests written (`cshell/fea.py`, `tests/cshell/`); env `cshell` building. |
| 2026-10-28 | pre-0 | Owen: D1 adopted (a ≥ (t_min − d)⁺); plan re-ordered so fast/sparse FE is researched and built first (§2.2); MMA vs L-BFGS-B comparison added to Stage 1. |
| 2026-10-28 | pre-0 | Spec read in full. Reference code checked selectively (ElementData, stress kernel, memory models located). Environment survey done. Plan, risks and disagreements written (this file). No code yet. |

**Done:** reading of the spec; dependency survey (§2.1); location of the q1_im buckling-sensitivity bug (§4, item D6).
**Next:** Stage 0.1 (env, `ShellCG`, dense baseline costs) and 0.2 (fast/sparse FE research) in parallel.
**Blockers:** none.

---

## 1. Platform and working practice

- **Isambard-AI** (login node `login43`, aarch64, GH200). Slot cap 8 queued/running jobs. HBM per GPU to be confirmed with `nvidia-smi` in Stage 0; Isambard-AI GH200s are the 96 GB variant, so the spec's "144 GB only" items (0.3 mm dense cantilever) are assumed **not** to fit.
- Login node: editing, git, job submission only. Every test, benchmark and run goes through `sbatch` (CPU-only `pytest` on a compute node too, because numba JIT and OCC tessellation are not light).
- Storage: code, configs, docs and small reports in `$HOME/shell-lightweighting`; the conda env, SDF/band caches, checkpoints, run folders and VTK files in `/projects/b6bl/owenpeckham.b6bl/home_data/shell-lightweighting/…`, symlinked back (`runs_cshell/`, `.cache_cshell/`). Never symlink git-tracked paths.
- Git: work on feature branch `cshell` (never `main`), commit at each passing gate with the tests named in the message (§10.6).
- I implement the modules myself in the main thread (they are tightly coupled through the sensitivity chain). Subagents only for genuinely parallel, context-heavy side work (e.g. the manuscript extraction, the CalculiX validation driver), each with its own output directory and a share of the slot cap.

## 2. Plan

### 2.1 Environment (Stage 0)

New conda env `cshell` (Python 3.12) in the projects area. Survey of aarch64 availability on 2026-10-28:

| Package | aarch64 build | Note |
| --- | --- | --- |
| numpy, scipy, numba, numba-cuda, cupy | yes (conda-forge) | Owen's `pcls_buck` env already has numba + cupy working on GH200; copy its CUDA pins. |
| build123d / cadquery-ocp | yes (pip `cadquery_ocp-8.0.1.1` manylinux_2_28_aarch64, `build123d-0.13`) | vshell's STEP import. |
| libigl | yes (pip 2.6.3 aarch64) | signed distance, cotmatrix, massmatrix. |
| warp-lang | yes (pip aarch64) | GPU closest-point / winding number. |
| mmapy | yes (pure Python) | MMA. |
| trimesh, scikit-image, matplotlib, pytest | yes | |
| **gpytoolbox** | **no aarch64 wheel** | Remesh via pymeshlab (conda-forge, check aarch64) or a small Botsch–Kobbelt remesher in numpy/igl (≈150 lines, tested on Hausdorff ≤ 0.02 mm). |
| calculix (`ccx`) | conda-forge listing found; aarch64 to confirm | Fallback: build `ccx` from source, or scikit-fem P2 for compliance as §8.6 allows. |
| tetgen / gmsh / pyvista | to confirm | Needed only for §8.6. |

Write `continuous/environment.yml` and `continuous/cluster/` (sbatch templates: `cpu_test.sh`, `gpu_test.sh`, `run.sh` with `--max-hours 23.5` and automatic `--resume` resubmission via `--dependency=afterany`).

### 2.2 Stage 0: fast, accurate FE first (re-ordered by Owen, 2026-10-28)

Owen's instruction: research and settle the FE engine **before** everything else, so no later test or run is slowed by the dense-box solver. This brings spec Stage 3a (narrow-band sparse FE) forward into Stage 0, with Owen's go-ahead. The spec (§4.2, §12) is updated when the engine is chosen.

**0.1 Reference baseline (what has to be beaten).**
1. Env; import `analysis`, `gpu_analysis` from `continuous/reference/` in `cshell/fea.py`; on a GPU node run `probe()`, `verify_host_kernel_consistency()`, record HBM size.
2. `ShellCG(analysis.ConjugateGradient)` with node-level BCs (§4.3), mm↔SI layer, tests on `cpu_direct` and GPU (Voigt order, energy normalisation, 40×4×4 beam). It stays the **reference model**: every faster engine must reproduce its answers.
3. Measure dense cost on the three parts: elastic solve at 0.8 / 0.4 mm, buckling at 0.8 mm (time, PCG/LOBPCG iterations, peak memory vs `elastic_memory` / `buckling_memory`). Into `BENCHMARKS.md`.

**0.2 Research: methods for very rapid, accurate (sparse) voxel FE on GH200.** Literature and available-code survey (researcher agent, conclusions only), in parallel with 0.1. Questions, each to be answered with expected cost/accuracy for our sizes (material-only DOF ≈ 2–4 M at 0.4 mm, 16–32 M at 0.2 mm, 126–256 M at 0.1 mm):
- **Active-set (narrow-band) voxel grids:** keep only elements within the design envelope {φ < d^{hi} + h_e} (25–60× fewer DOF than the box); SPGrid-style bricks (Liu et al. 2018), giga-voxel PETSc GMG (Aage et al. 2017); handling of the growing outer skin (one layer of ρ_min elements around the active set so gradients exist where material can appear), floating fragments, and active-set rebuild cost.
- **Solvers on that set:** (i) matrix-free q1_im + GMG by brick agglomeration; (ii) assembled CSR + GPU AMG (AmgX, CuPy smoothed aggregation, PETSc GAMG); (iii) **GPU sparse direct (cuDSS via nvmath-python)**: one factorisation reused for the primal, the stress adjoint, and shift-invert Lanczos buckling, which would make 0.4 mm (and maybe 0.2 mm) buckling routine and remove the 0.8 mm buckling compromise (K2).
- **Buckling on the sparse operator:** shift-invert Lanczos/LOBPCG with the chosen solver; the sensitivity then has to be written in cshell (derived from the same discrete model; the D6 fix carries over).
- **Accuracy per DOF:** whether cut-cell / finite-cell integration or higher-order elements beat the ersatz-material q1_im at equal cost for 1.5–2 mm walls (F7), versus simply going to 0.2 mm.
- **What is already installed** on Isambard-AI (structural-analysis skill: CuPy/cuDSS/AmgX, MFEM-CUDA, PETSc GPU, Warp, torch-fem) and on aarch64.

**0.3 Choose and build.** I bring Owen one recommendation with the numbers, then build it as `cshell/fe_sparse.py` behind the same interface as `ShellCG` (compliance, element energies, stress, adjoint solve, buckling + sensitivity). Tests: identical answers to the dense `ShellCG` reference on small and medium cases (compliance to 10⁻⁵ relative at equal tolerance; λ_b to 10⁻⁴), all §6.6 gradient tests through it, CPU path for R8.

**0.4 Validation and costs on the new engine.** V2-a (solid baselines vs vshell, 0.8 mm), V2-b (tubes at 0.8 / 0.4 / 0.2 mm), V2-c (thin-wall convergence to 0.2 mm), V3-a, V3-b (including 0.8 mm, D5). `BENCHMARKS.md` updated with sparse vs dense per-iteration cost, and a GPU-hour estimate for Stages 1–2 put to Owen before full runs.

**0.5 Context reading** (delegated extraction): manuscript Study 1 Tables 2–5, Study 2, §6.1–6.2; vshell theory numbers for V2-b and §11.1.

Gate: sparse engine matches the dense reference; V2-a, V2-b, V2-c, V3 pass; measured per-iteration cost at 0.4 and 0.2 mm for the three parts.

### 2.3 Stage 1: geometry, mass, compliance

Modules in dependency order, each with its tests before the next one starts:

1. `surface.py`: Γ_h remesh (ℓ_Γ = 0.75 mm, Hausdorff check, B-rep face id per vertex); lumped M, cotangent L; Helmholtz filter F and Fᵀ with one factorisation per r_f (V1-d); r_MA by march + bisection on exact distances (V1-b, ray-chord sanity check); r_out by ray cast; keep-out bounds; passive vertices from vshell's mount detection.
2. `band.py`: sparse band (φ, closest triangle, barycentrics) on GPU with Warp, cached `.npz` keyed by STEP SHA-256 + grid parameters; P (cell → Γ_h vertex, 3 nnz/row).
3. `design.py`: ξ → ξ̃ → (d, a, a^ab) → ψ → χ (quintic H_ε) → ρ → ρ̂, volume, sliver measure, and the hand-written pull-backs of §6.1 (CuPy on GPU, NumPy float64 on CPU; same code via an `xp` switch). V1-a, V1-c, V4-1 (directional + Taylor).
4. `mma.py`: mmapy wrapper, scaling, restoration by Brent along ∇V (V4-7), Svanberg toy problems (V5-B4). Optimiser comparison asked for by Owen: MMA against projected L-BFGS-B (mass held by the same restoration step) on B2 and B3, reporting evaluations to converge and final J. MMA stays the default unless L-BFGS-B is clearly better.
5. `objective.py`, `optimise.py` (compliance only, w_P = 0): phases A/B with rebuild of band/P/grid between phases, adaptive D_comp with filter-inverting re-map, checkpoints every iteration, `--resume`, `--max-hours`, `history.csv`. V4-3, V4-6.
6. Benchmarks B1 (natural closure from S1), B2; then compliance-only runs of the three parts from S1 and S3, checking |V/V₀ − 1| ≤ 10⁻⁴ at every iterate.

Gate: V1, V4 (1–3, 6, 7), V5-B1/B2/B4, three compliance-only runs meeting R1.

### 2.4 Stage 2: capacity, robustness, manufacture, decision

1. Stress p-norm with my own centre-strain (B(0)) material-stress kernel and a matching bilinear ηᵀK⁰u kernel; adjoint via `matrix_free_pcg` with the primal's operator (`q1_im`, ν, `gmg`, `tol=0`, `rel_tol=1e-8`); host path for CPU tests. V4-4.
2. Buckling wrapper: separate 0.8 mm `ShellCG` instance fed the 2³-block-averaged ρ̂; corrected q1_im sensitivity (D6 below); chain back by /8. V4-5 (< 10⁻³ for q1 and corrected q1_im), V3-a, V3-b.
3. As-built erosion, soft-min capacity, Le normalisation c, sliver penalty with continuation, full J; S2 (vshell) start.
4. `export.py` (tiled ψ at 0.05 mm, closure snap, v_min opening, small-void fill, largest component, escape holes, δ mass correction, marching cubes, R1 mesh check, wall histogram), `validate.py` (TetGen/gmsh C3D10 + ccx static and `*BUCKLE`), resolution check (grid shift + 0.5 mm), decision rule, `report.py` (15 pages, rebuildable), `summary.json/.txt`, provenance.
5. Full runs of the three parts (3 starts each), V2-c, V5-B3, V6, V7; `VALIDATION.md`; honest comparison with solid / uniform shell / vshell / manuscript.

Gate: V2-c, V3, V4 (4–5), V5-B3, V6, V7; reports for all three parts.

### 2.5 Stage 3 (not started without Owen's go-ahead): proposed priority

1. (a) Narrow-band sparse FE: **moved to Stage 0** (§2.2). Remaining Stage 3 part: 0.1 mm validation runs if 0.2 mm is the production resolution.
2. **(e) Plastic limit-load capacity.** The only route to the MBB's measured +14.8 % (ductile redistribution).
3. **(d) Local/global mode classification** with separate knock-downs; cheap once modes are reported.
4. **(f) Kinematic BC bound,** to bracket the bolt-hole uncertainty.
5. (c) escape holes in the loop; (b) multi-GPU; (g) ribs/infill; (h) Fusion round-trip.

## 3. Risks

| # | Risk | Effect | Mitigation |
| --- | --- | --- | --- |
| K1 | **Dense box is wasteful.** Material fills 1.7–4 % of the box, so every solve at 0.4 mm handles 100–173 M DOF, most of it ρ_min void (contrast 10⁶). | PCG iteration counts and per-iteration time may be high; phase B (80 its × 3 starts × 3 parts) could need hundreds of GPU-hours. | Measure in Stage 0 before committing. Crop to {φ < d^{hi} + h_e} (§4.2). Warm starts. If too slow, propose bringing Stage 3a forward. |
| K2 | **Buckling at 0.8 mm only.** A 1.5–2 mm wall is 2 elements thick at 0.8 mm. | Local shell modes (F3) under-resolved; the optimiser could thin walls that would buckle locally. | Quantify: run V3-b also at 0.8 mm (not only 0.2–0.4); report the 0.8 vs 0.4 λ_b error on final designs where 0.4 fits (L-bracket/MBB may fit with `gmg_lobpcg32` if crop is tight — check `buckling_memory`). |
| K3 | Mixed resolutions per iteration: compliance/stress at 0.4 mm, buckling at 0.8 mm, so one extra elastic solve at 0.8 mm per iteration. | Cost, and J mixes two discretisations. | Accept (spec); log both grids in history and report. |
| K4 | `compute_element_stress` returns element-*averaged* stresses (it takes no `element_type`), not the q1_im centre strain the adjoint assumes. | Inconsistent stress gradient. | Own centre-strain kernel used for both the response and its adjoint; unit test against the reference kernel's output. |
| K5 | GH200 HBM is 96 GB here. | No 0.3 mm dense check for the cantilever; resolution check is grid-shift + 0.5 mm until Stage 3a. | Report it on the resolution page. |
| K6 | Band size at h_g = 0.1 mm: ≈ 1.6·10⁸ cells (cantilever, D_comp = 4 mm), ≈ 2.6·10⁸ after one D_comp doubling; ≈ 40 B/cell. | 6–10 GB GPU + host memory; one-off SDF time. | Warp BVH on GPU; cache to disk in the projects area; tiled processing. |
| K7 | Non-convexity / void closure is a near-topological switch. | Local optima, start dependence. | Multi-start (S1–S3, S4 optional), γ_c window, reopening probe, report all starts. |
| K8 | aarch64 gaps (gpytoolbox, maybe ccx/tetgen). | Missing remesher or validator. | Fallbacks in §2.1; any fallback logged and reported. |
| K9 | V7 (identical rerun to 10⁻⁴) with non-deterministic GPU reductions and LOBPCG starts. | Could fail by chance. | Fixed seeds, float64 reductions, deterministic LOBPCG init; state tolerance. |
| K10 | Full runs vs 24 h limit and slot cap 8: 3 parts × 3 starts = 9 runs. | Queueing. | `--max-hours` + dependent resubmission; run 8 at once, ninth after. |

## 4. Disagreements and inconsistencies with the spec

Each will be resolved by updating THEORY_continuous.md in the same change as the implementation, as the prompt requires.

| # | Where | Issue | Proposed resolution |
| --- | --- | --- | --- |
| D1 | §3.4 a^{lo} = t_min | **The design space excludes the add-in's uniform shell** (a = 0, d = t) and any wall built mostly outward. Every void wall must keep ≥ 1.5 mm *inside* Γ₀ even where d already supplies the wall, so material sits nearer the neutral axis than necessary (less I per gram), and members with r_MA < 3 mm are forced passive although the add-in hollows them. cshell could then lose to the uniform shell it is compared against. | Make the bound depend on d: a ≥ a^{lo}(d) = (t_min − d)⁺, smoothed (softplus, width ~0.1 mm), via a_i = a^{lo}(d_i) + (a^{hi}_i − a^{lo}(d_i)) ξ̃ᵃ_i. Still a box in ξ, still w ≥ t_min, adds one chain-rule term. Passive rule becomes a^{hi} ≤ v_min/2 where d^{hi} ≥ t_min. **Decided by Owen 2026-10-28: adopt a ≥ (t_min − d)⁺ (smoothed).** Spec §3.4, §6.1, §14 to be updated with the implementation. |
| D2 | §1.3 R1 vs prompt R1 vs §7.2 | Spec R1 says 10⁻³ per iterate; prompt and §7.2/§12 say 10⁻⁴. | Use 10⁻⁴ (restoration tolerance) everywhere; fix §1.3. |
| D3 | §5.1, §6.5 | ω_k in J although D3 (single load case) says no case weights. | Drop ω_k (sum over the force list as one case); keep `loads` a list. |
| D4 | §7.4 S2 | vshell's hollow regions have a = 0, d = s, not representable with a ≥ t_min. | Map as specified (approximate start) and say so; exact if D1 is adopted. |
| D5 | §11 V3-b | Production buckling runs at 0.8 mm but local-buckling validation only at 0.2–0.4 mm. | Also report V3-b at 0.8 mm (no pass criterion, quantifies K2). |
| D6 | §6.4 | Location confirmed: `ElementData.__init__` sets `self.K0 = lam_s Klam + mu_s Kmu` (analysis.py l. 5018), used by Term A (l. 6013) and the explicit adjoint term (l. 6047). | Fix without copying code: inside `ShellCG.run_buckling`, temporarily substitute `analysis.ElementData` with a subclass whose `__init__` resets `K0 = E_s·Khat` (E_s = 2μ_s(1+ν)) for q1/q1_im, restore afterwards. Prove with the §6.4 directional test; write up in OPEN_QUESTIONS.md. |
| D7 | §4.4(b) | `compute_element_stress` has no `element_type` argument and averages; spec asks to "confirm which point". | See K4: own centre-strain kernel. |
| D8 | §7.2 | Restoring along ∇V moves ξᵃ too, so the restoration itself can open/close voids slightly. | Keep as specified (it is the least-change correction and keeps V monotone); monitor in history (count of vertices crossing ξ̃ᵃ = 0.98). Revisit only if it fights MMA. |

## 5. Per-part results

Not started. Filled in at the Stage 2 gate: decision, J, K/K₀, P/P₀, governing mode, wall range, mass error, run time / GPU-hours, comparison of solid / uniform shell / vshell best / cshell, sign agreement with the manuscript.
