# cshell: open questions, assumptions and reference-code issues

Each item: status, the default implemented, and what would change it. Newest at the top of each section.

## A. Questions for Owen

| # | Question | Default implemented | Source |
| --- | --- | --- | --- |
| Q1 | Buckling knock-down κ_b. | 0.85; report re-scores 0.6–1.0. | THEORY §14 |
| Q2 | Escape holes: diameter and count. | 3 mm, 2 per void. | THEORY §14 |
| Q3 | Plastic limit-load capacity (Stage 3e) before a test campaign? | Not started. | THEORY §14 |
| Q4 | **V2-b tolerance vs vshell's own result.** vshell's tube validation at h = 0.8 mm is +4.5 % in E (+4.1 % in ratio) against Abaqus, already outside the 4 % gate. Only h = 0.8 mm results are stored. | Gate kept at 4 %; report the h = 0.8/0.4/0.2 mm numbers and the trend. | `continuous/research/manuscript_numbers.md` |
| Q5 | **V2-c criterion (0.4 within 3 % of 0.2).** The vshell doc reports −6.2 % at 0.4 mm and ±6 % thin-wall noise (with vshell's linear-ramp volume fractions). | Gate kept; the quintic Heaviside and smooth ρ should do better. Report honestly if not. | same |
| Q6 | **Abaqus is not ground truth for absolute stiffness.** Abaqus over-predicts measured PA11 tube stiffness by +35 to +60 %, and even the K(E)/K(A) ratio is 4.08 in Abaqus against 4.72 measured. V2-b validates the FE against another FE only. | Ratios between designs are what cshell uses; stated in the report. | same |
| Q7 | THEORY §11.1 quotes the vshell cantilever as K/K₀ 0.992, P/P₀ 2.63 (= `examples/runs/Cantilever/summary.txt`, 7 solid regions); the vshell theory doc's k = 0 row reads +1.2 %, +148 %. | Use the run summary (the repository's actual vshell output). | same |
| Q8 | The manuscript text says walls printed "up to 13.6 %" thinner; Table 5 gives 9.9–12.6 %. | b_v = 0.2 mm as specified. | same |
| Q9 | vshell's solid L-bracket K₀ = 48.3 N/mm lies above the manuscript's Abaqus bounds (25–38 N/mm). Different BC models (vshell clamps whole bolt-hole faces) are the likely cause. | V2-a compares against vshell (same BCs), as specified. | same |

## B. Reference-code issues (continuous/reference, not edited)

### B1. q1_im buckling sensitivity uses the Gauss solid stiffness (THEORY §6.4)

`analysis.ElementData.__init__` sets `self.K0 = lam_s·Klam + mu_s·Kmu` (analysis.py l. 5018), the trilinear Gauss
stiffness, and uses it in Term A (l. 6013) and the explicit adjoint term (l. 6047), while the eigenproblem is assembled
with `E_e·Khat` (q1_im). For q1_im the returned ∂μ̃/∂ρ̂ is therefore not the derivative of the solved model:
directional finite differences show a constant ≈ 5.8 % bias (spec reproducer, THEORY §6.4 table).

cshell's production engine (`cshell/fe_sparse.py`, `buckling`) does not use this code: its sensitivity uses
K⁰_e = E·K̂ (q1_im) throughout and is verified by `tests/cshell/test_buckling.py::test_directional_sensitivity_q1im`
(same set-up as THEORY §6.4). Suggested fix in the reference: for `element_type in ("q1", "q1_im")`, set
`self.K0 = 2·mu_s·(1 + nu_run)·Khat`. Status of the in-cshell wrapper test of that fix: pending.

### B2. Reference GPU path is float32

`matrix_free_pcg` and the energy kernels are float32: the identity Σ U_e = C/2 holds only to ≈ 1·10⁻³ at
`rel_tol = 1e-6`, and element-wise finite differences of C are not meaningful on GPU. Not a bug; recorded because
§6.6(3) is therefore checked on the float64 host path (and on the float64 cuDSS engine).

### B3. `compute_element_stress` averages over Gauss points

It takes no `element_type` and returns element-averaged compatible-B stress, not the q1_im centre strain THEORY §4.4(b)
needs. cshell computes centre stress itself (`fe_sparse.SparseFE.centre_stress`).

### B4. `ConjugateGradient._volume_mesh` docstring says `fixed_mask (Nnodes,)`; it is int32 per DOF (3·Nnodes).

## C. Assumptions

| # | Assumption |
| --- | --- |
| A1 | Elements outside the active set carry ρ̂ = ρ_min in the dense model; dropping them changes K by O(ρ_min) and no design derivative (STATUS §2.2, 0.3). Tested against the dense reference (`test_fe_sparse.py`). |
| A2 | Owen's D1: inset bound a ≥ (t_min − d)⁺, smoothed. |

### B5. Dense reference solve reads 2–7 % stiffer than an exact solve (h_e = 0.8 mm, solid Study 1 parts)

`ConjugateGradient.run_analysis` (GMG-PCG, float32, `pcg_rel_tol = 1e-6`) gave K₀ = 263.9 / 51.87 / 482.6 N/mm for the
cantilever / L-bracket / MBB, against 257.8 / 48.60 / 473.5 from the fp64 sparse direct solve on the same elements, BCs
and ρ̂ (and 256.8 / 48.3 / 473 from vshell's independent AMG solver). Compliance from an unconverged CG under-estimates
the energy, so K comes out high; with ρ_min = 1e-6 void the relative residual does not bound the energy error.
Reproducer: `continuous/research/stage0/bench_solid.py` (dense run at 0.8 mm). Not used by cshell; worth a look for
PCLS runs that rely on absolute compliance.
