> Renamed 2026-10-31: vshell → dshell (discrete optimiser, `dshell/`), continuous/ → docs/cshell/, shell_lightweighting/ → ushell/.

# Prompt for Claude Code: build `cshell`, the continuous constant-mass shelling optimiser

You are working in the `shell-lightweighting` repository. Your job is to design, implement, verify and document **`cshell`**, a headless Python toolchain that:

- takes a solid STEP part (a Fusion generative-design outcome), its supports and its load;
- returns a part of **exactly the same mass** whose wall thickness varies **continuously** over the surface, built both into the original volume (void wall moved inward) and out of it (outer skin moved outward);
- chooses the thickness by **gradient-based optimisation on GPU finite-element fields** to maximise stiffness and load capacity (strength and real eigen-buckling);
- closes the void locally where hollowing doesn't pay (solid joints), and returns the original solid with the reason if it doesn't pay anywhere.

It is the continuous successor to `vshell/`, which is already in the repo, working, and discrete (each region either solid or hollow). The owner is Owen Peckham (University of Bristol, DMF Lab). He cannot run physical tests at present, so validation is numerical, plus agreement with the trends measured in the manuscript.

## 1. Read first, in this order

1. **`docs/cshell/THEORY_continuous.md` is the specification.** It covers the full theory, equations, sensitivities, optimiser, outputs, software layout, test plan with pass criteria, staged roadmap with gates, risks, and Owen's decisions (§14). Follow it. Where you deviate, say why in `docs/cshell/STATUS.md` and update the spec in the same change.
2. `docs/cshell/reference/README.md`, then `docs/cshell/reference/vshell_theory_and_operation.md` and `docs/cshell/reference/feasibility_study.md`.
3. **The manuscript, `docs/cshell/reference/main.pdf`** (about 24 MB; read it by page ranges). Focus on Study 1 (Tables 2–5), Study 2 and Sections 6.1–6.2.
4. **The reference GPU FE stack:** `docs/cshell/reference/gpu_analysis.py` and `docs/cshell/reference/analysis.py`. These are large files, so read selectively:
   - the module docstrings;
   - `ConjugateGradient`: `__init__`, `_volume_mesh`, `run_analysis`, `run_buckling`;
   - `matrix_free_pcg`;
   - the memory models (`elastic_memory`, `buckling_memory`) and `probe`.

   Spec Appendix A summarises what matters and what has already been checked.
5. **The existing code:**
   - `vshell/`, especially `geometry.py`, `field.py`, `segment.py`, `report.py`, `util.py` and `config.py`;
   - `tests/`;
   - `examples/*.json`, `examples/parts/*.step` and `examples/dshell/runs/*/summary.txt`.

## 2. Non-negotiable requirements (spec §1.3 and §14)

- **R1, constant mass.** Every reported iterate has |V/V₀ − 1| ≤ 1e-4 after the restoration step (§7.2). The exported mesh is within 2e-3.
- **R2, stay solid unless hollowing earns it.** Apply vshell's acceptance rule plus the resolution check of §5.4. Otherwise output the original part with the reason. Never tune parameters to force a hollow result.
- **R3, supports.** Use the minimum-stiffness bolt-hole case: all DOF of the support faces are clamped.
- **R4, mounts never change.** Mounts are the analytic preserve faces plus the support and load faces.
- **R5, MJF manufacturability:**
  - t_min = 1.5 mm nominal, with a 0.2 mm as-built bias;
  - no voids thinner than v_min = 3 mm;
  - escape holes in every enclosed void;
  - **no ceiling on outward growth**, but keep clearance from other members of the part and from any keep-out bodies;
  - only a single load case is required.
- **Objective weights.** w_K and w_P are user-tunable hyper-parameters (config and `--set`), equal by default, and recorded in every summary and report.
- **R6, verified gradients.** Every gradient passes the finite-difference and directional tests of §6.6 before it is used in an optimisation.
- **R7, report and provenance.** Every run writes `report.pdf` and full provenance (§9), following `dshell/report.py`.
- **R8, headless and testable.** It runs headless on GH200 GPUs, and the test suite runs CPU-only at small sizes (the reference code's `cpu_direct` engine works without CUDA).
- **Runs survive the 24 h SLURM limit.** Runs must be resumable across allocations: provide `--resume` and `--max-hours` (§7.6).

## 3. Use the reference FE stack, don't reinvent it

- **Import, don't edit.** Import `analysis` and `gpu_analysis` from `docs/cshell/reference/` via `sys.path` in `cshell/fea.py`. **Do not modify or copy those files.** `interfaces.py` there is a stub: the original wasn't supplied.
- **Adapter.** Build `ShellCG(analysis.ConjugateGradient)` with node-level boundary conditions, exactly as specified in §4.3. Mind:
  - the dtypes, including the per-DOF int32 `fixed_mask`;
  - the element-shaped `csg_domain_data` it still needs;
  - SI units (Pa, m, N, J).
- **Element.** Use `q1_im` for production.
- **Known issue to fix inside cshell (§6.4).** For `q1_im`, the sensitivity returned by `run_buckling` uses the Gauss solid stiffness in Term A and in the adjoint term. That gives a 5.8 % directional bias; the same code is exact to 1e-4 for `q1`.
  1. Correct it in a cshell subclass or wrapper, using E_s·Khat.
  2. Prove the fix with the directional test.
  3. Write it up for Owen in `docs/cshell/OPEN_QUESTIONS.md`, with the reproducer.
- **Stress adjoint.** Call `matrix_free_pcg` with exactly the primal's operator: `element_type="q1_im"`, `nu`, `"gmg"`, `tol=0` and a `rel_tol` (§6.3). Its defaults differ from `ConjugateGradient`'s.
- **Memory.** Check memory with `gpu_analysis.elastic_memory` and `gpu_analysis.buckling_memory` before every phase, and whenever the outward computational extent grows (§3.5, §4.2). Dense 0.4 mm elastic solves need about 42 GB for the cantilever. Dense 0.4 mm buckling does not fit, so run buckling on the 0.8 mm grid and chain the sensitivity back. Measure rather than assume, and record the numbers in `docs/cshell/BENCHMARKS.md`.

## 4. Stages and gates (spec §12)

Work through the stages in order, and don't start a stage until the previous gate passes.

- **Stage 0: set-up.**
  - Environment and dependencies (§10.4; everything must have aarch64 builds for the GH200 Grace host).
  - Import the reference stack.
  - Run `gpu_analysis.probe()` and `verify_host_kernel_consistency()`.
  - Build the `ShellCG` adapter.
  - Measure solve time and memory at h_e = 0.8 and 0.4 mm for the three Study 1 parts.

  Gate: V2-a (solid baselines within 2 % of vshell at 0.8 mm) and V2-b (Study 2 tubes within 4 % of Abaqus) pass.
- **Stage 1: geometry, mass and compliance.**
  - Design surface mesh and Helmholtz filter.
  - r_MA, by exact-distance bisection.
  - Clearance and keep-out bounds, and the adaptive outward extent.
  - Sparse band and P.
  - ψ → χ → ρ chain with hand-written pull-backs.
  - Volume, mass restoration and MMA.
  - Compliance-only objective, starts S1 and S3, phases A and B.

  Gate: V1, V4 (1–3, 6, 7) and V5-B1/B2/B4 pass, and compliance-only runs of the three parts meet R1 at every iterate.
- **Stage 2: capacity, robustness, manufacture and decision.**
  - Stress p-norm and its adjoint.
  - Buckling, with the corrected sensitivity.
  - As-built erosion and soft-min capacity in J.
  - Sliver penalty and the S2 (vshell) start.
  - Export: clean-up, escape holes, mass correction and meshing.
  - Resolution check, body-fitted CalculiX validation and the full report.

  Gate: V2-c, V3, V4 (4–5), V5-B3, V6 and V7 pass, and there is a full report for each of the three parts. Compare honestly with vshell and with the manuscript's measured changes (§11.1).
- **Stage 3: extensions.** Propose a priority order to Owen in `STATUS.md`, and don't start without his go-ahead. Candidates:
  - narrow-band sparse FE at 0.2/0.1 mm;
  - multi-GPU;
  - escape holes inside the loop;
  - local/global buckling classification;
  - plastic limit-load capacity;
  - the kinematic BC bound;
  - interior ribs or infill.

## 5. How to approach the work

- **Plan before coding.** Read the material, then write `docs/cshell/STATUS.md` with your plan, the risks you see and any disagreement with the spec.
- **Small, tested increments.** Every module has unit tests: pytest, CPU by default, with `@pytest.mark.gpu` for GPU tests. No gradient is used before it passes §6.6.
- **No silent fallbacks.** Every engine, resolution or solver fallback is logged and shown in the report. A non-converged PCG or eigen-solve is never accepted silently.
- **Questions.** Owen's decisions are in spec §14. For anything still open, implement the stated default, record the question in `docs/cshell/OPEN_QUESTIONS.md` and carry on. Stop and ask only for decisions that would waste days of GPU time or can't be undone.
- **Report honestly.** If the continuous optimiser doesn't beat vshell or the solid on a part, that is a valid result: report it with the mechanism, don't tune it away. The linear model is expected to miss the MBB's ductile capacity gain (§4.5); say so.
- **Keep vshell working.** `cshell` imports from it and must not change its behaviour. After any edit to `vshell/`, run `python tests/dshell/test_basic.py`.

## 6. Definition of done (Stages 0–2)

1. **End-to-end runs.** `python -m cshell run examples/cshell/<part>.json` runs end to end for the Cantilever, L-bracket and MBB: every start, phase, export, validation and report. Runs resume correctly across 24 h allocations, and the outputs are as listed in spec §10.5.
2. **Validation.** All tests in spec §11 pass, and the results are tabulated with their numbers in `docs/cshell/VALIDATION.md`.
3. **Status summary.** `docs/cshell/STATUS.md` summarises each part:
   - decision, J, K/K₀, P/P₀ and governing mode;
   - wall range, mass error, and run time / GPU-hours;
   - the comparison of solid, uniform shell, vshell best and cshell;
   - whether the predicted sign of change agrees with the manuscript's measurements.
4. **Open questions.** `docs/cshell/OPEN_QUESTIONS.md` lists every assumption and every issue found in the reference code.
5. **README.** `cshell/README.md` covers install, commands, config keys and outputs, in the style of `dshell/README.md`.
6. **Spec kept current.** The spec is updated wherever the implementation deviated.

Begin with section 1: read the spec in full.
