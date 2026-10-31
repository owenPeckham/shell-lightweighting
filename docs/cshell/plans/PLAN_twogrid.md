> Renamed 2026-10-31: vshell → dshell (discrete optimiser, `dshell/`), continuous/ → docs/cshell/, shell_lightweighting/ → ushell/.

# PLAN: two-grid PCG FE engine for the optimiser (agreed with Owen 2026-10-29)

Owen's requirements: the analysis used inside the optimiser uses the **minimum GPU memory** while staying **fast and
accurate**, on **one GPU only** (multi-GPU is allowed only for validation solves), and also the **minimum host (CPU)
memory**: no host-side copies of fine-level matrices or element arrays beyond what set-up needs (build on the GPU,
free host temporaries), measure peak host RSS (`/usr/bin/time -v` or `resource.getrusage`) alongside GPU memory, and
report both. **Nothing runs on the login node** — not even quick tests or import checks; everything via sbatch.
Full runs wait for this engine.

Worktree `/home/b6bl/owenpeckham.b6bl/shell-lightweighting/.claude/worktrees/cshell-stage2` (= REPO), branch
`cshell-stage2`. Do not commit, switch branch or push (the orchestrator commits). Common rules as in
`docs/cshell/plans/PLAN_parallel_modules.md` "Common rules"; env `source cshell/cluster/env.sh` (cd's to
SLURM_SUBMIT_DIR: submit from REPO, pass absolute test paths). Runs under `runs_cshell/TG_*` only.

## Why
Hollow designs at h_e = 0.4 mm reach 6.1 M (cantilever) – 8.4 M (MBB) active DOF. The fp64 cuDSS factor then needs
79–114 GiB and does not fit; today's fallbacks (`cshell/fe_sparse.py`):
- fp32 factor + fp64 CG: 94–129 s FE per iteration, peak 81–94.8 GiB (MBB at the limit) — jobs 7801143/7801144;
- hybrid fp64 (factor partly in host RAM, needs `--mem=240G`): 58 s, peak 54 GiB — job 7808291.
The memory study (`docs/cshell/research/gpu_memory/REPORT.md`, prototype `twogrid.py`, results `*tg*.json`) measured a
two-grid PCG on the MBB solid at 0.4 mm: 24 iterations, 0.45–0.51 s per solve, buckling 22 s, λ ≤ 4e-8 and dμ ≤ 1.1e-4
vs the direct solve, memory −55 GiB (−65 matrix-free); at 0.2 mm matrix-free 2.35 s (cantilever 25.5 M DOF, 40 GiB) and
3.0 s (MBB 35.4 M DOF, 57 GiB). Not yet tried on a voided design.

## Method (THEORY §4.1–4.2; same discrete model as `SparseFE`: q1_im, E ρ̂ K̂, clamped DOF as identity)
- **Fine level (h_e):** K u applied **matrix-free** element by element on the GPU (gather u_e, batched K̂ u_e scaled by
  E ρ̂_e, scatter-add), for k right-hand sides at once. No assembled fine K, no fine factor. Diagonal of K computed
  element-wise for the smoother. If matrix-free is > 1.5× slower per solve than an assembled fine CSR, report both
  and keep matrix-free as default unless the per-iteration total rises by > 20 %.
- **Coarse level (2 h_e):** trilinear prolongation P from coarse nodes (those touching active fine elements) to fine
  nodes; **Galerkin** coarse operator K_c = Pᵀ K P assembled element-wise without a fine matrix: every fine element is
  one of the 8 children of a coarse element, so K_c = Σ_e E ρ̂_e (P_childᵀ K̂ P_child) with 8 precomputed 24×24
  matrices. Clamped DOF: coarse DOF whose support is entirely clamped are fixed; P respects the fine mask. Factorise
  K_c with cuDSS (fp32 if that is enough for the PCG iteration count; record), one factor per design.
- **Preconditioner:** two-grid V-cycle — Chebyshev(Jacobi) pre/post smoothing (degree 3, eigen-bound by a few power
  iterations), coarse direct solve. Symmetric, so PCG is valid.
- **PCG:** fp64, relative residual tolerance chosen so that every §6.6 gradient test still passes at its current
  threshold (expect ~1e-10 for primal and adjoints); warm start from the previous design iterate's solution; block
  (multi-RHS) version for the buckling Krylov steps. Non-convergence raises (no silent fallback).
- **Buckling:** existing `buckling()` block-Krylov on K⁻¹G with K⁻¹ = this PCG; K products via the matrix-free operator
  (`fe._K @ X` must keep working: expose an operator object with `@` for vectors and blocks).

## Interface and files
- `cshell/fe_twogrid.py` (new): `TwoGridFE` with the **same public interface as `SparseFE`** (`factor(ρ̂)`, `solve`,
  `primal`, `element_bilinear`, `centre_stress`, `dC_drho_hat`, `free`, `stats`, `_K` operator, `_dofs`, `_Khat`,
  `fixed`, `f`, `rho_hat`, `xp`, `E`, `h`, …) so `objective.py`, `buckling()` and `optimise.py` use it unchanged.
- `cshell/fe_sparse.py`: only if a shared helper must move; keep `SparseFE` behaviour identical.
- `cshell/optimise.py` (`ensure_active` engine construction only) and `cshell/config.py`: `fe.engine` = "twogrid" |
  "direct" | "auto". Default **"auto" = direct fp64 when cuDSS's plan estimate says it fits with margin AND the
  problem is small (phase A, < ~2 M DOF), else two-grid** — report the phase-A memory of both so Owen can decide
  whether two-grid should be used everywhere. Log the engine per phase; record in history (`fe_engine` column) and
  `stats` (PCG iterations per solve, coarse DOF, memory).
- `cshell/collect.py`: re-scoring uses the same engine selection (no edit needed if it goes through `PhaseModel`).
- Tests: `tests/cshell/test_fe_twogrid.py` + parametrise existing `test_fe_sparse.py` / `test_buckling.py` /
  `test_objective.py` gradient tests over both engines where cheap.

## Acceptance (each with numbers in the return)
1. CPU + GPU tests pass; two-grid vs direct on test problems: compliance ≤ 1e-8 rel, u ≤ 1e-7, λ (6 modes) ≤ 1e-6,
   ∂C/∂ρ̂ ≤ 1e-7; all §6.6 gradient tests (compliance FD, stress adjoint 40×6×6 notched bar < 1e-3, buckling directional
   q1_im < 1e-3, end-to-end dJ/dξ < 2 %) pass with the two-grid engine.
2. On **voided phase-B designs** (use `runs_cshell/fixhyb_mbb` and `runs_cshell/fix_cantilever` checkpoints/ξ, MBB 8.4 M
   and cantilever 6.1 M DOF, full objective): J, C, λ_s, λ_b, dJ vs the hybrid-fp64 direct solve (J to 1e-7, dJ to
   1e-5 rel in the 2-norm); PCG iterations per solve; FE time per iteration (target ≤ 30 s MBB) and peak GPU memory
   (target ≤ 30 GiB MBB, measured by nvidia-smi sampling) with the default `--mem` (no 240G).
3. A 3-iteration phase-B run (cantilever and MBB, `--set optimiser.max_iter=[1,3]`) reproduces the direct history
   (J to 1e-6) with the default resource request.
4. Numbers into `docs/cshell/BENCHMARKS.md` ("Two-grid engine") and a short `docs/cshell/research/twogrid/notes.md`.

Slots: 4 (check `squeue -u $USER -h | wc -l` < 8 before each submission; chain with --dependency, don't poll).

## Return (under 300 words)
Files changed; tests and job ids; accuracy table vs direct; per-iteration time and peak memory on both parts (direct
fp32 / hybrid / two-grid); default engine chosen per phase and why; open issues.
