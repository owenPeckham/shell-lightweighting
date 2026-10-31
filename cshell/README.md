# cshell — continuous constant-mass shelling optimiser

`cshell` turns a solid part (STEP) into a hollowed, variable-thickness shell part of the **same mass** that is
stiffer and stronger: it maximises J = w_K ln K/K₀ + w_P ln P/P₀ (stiffness and capacity — strength p-norm and
linear buckling, soft-min — relative to the solid), subject to constant mass, minimum wall t_min, minimum void
v_min, as-built biases and closed escape holes. Spec: `docs/cshell/THEORY_continuous.md`; validation record:
`docs/cshell/VALIDATION.md`; costs: `docs/cshell/BENCHMARKS.md`; state: `docs/cshell/STATUS.md`.

## Use (Isambard-AI; nothing on the login node)

1. Config: copy `examples/cshell/cantilever.json` — `name`, `step_file`, `output_dir`, `supports` (face selectors),
   `loads` (face selectors + `force_N`). Everything else defaults (`cshell/config.py`); override with `--set k=v`.
   `python -m cshell init` / `faces` print a template and the STEP faces (run them via `gpu_test.sh`, not on the login
   node).
2. Run the three starts (uniform, dshell, neutral; `vshell` is accepted as the old name of dshell), one GPU each, auto-resumed, collect submitted by the last one:

       CSHELL_TIME=04:00:00 SBATCH_MEM_PER_NODE=64G CSHELL_COLLECT_MEM=180G \
         cshell/cluster/launch_part.sh examples/cshell/<part>.json uniform dshell neutral

   or a single start: `sbatch cshell/cluster/run.sh <cfg> --start uniform`; collect by hand:
   `sbatch --mem=180G cshell/cluster/collect.sh <abs run dir>`.
3. Results in `<output_dir>/<name>/`: `summary.txt` / `summary.json` (decision, reason, per-resolution J, K/K₀,
   P/P₀, mass error, baselines: solid, uniform shell with solid mounts, dshell best), `final_design.stl` (the shell
   part if accepted, else the original), `best_candidate.stl`, `fields.vtp` (ParaView: thickness, offsets, stress),
   `buckling_modes.npz`, `sensitivity.csv`, `validation.csv` (body-fitted check where it ran), `report.pdf`.

Committed results for the three example parts (MBB, Cantilever, L-bracket; reports, summaries, histories, fields,
preview STLs): `examples/cshell/runs/<part>/`, described in `examples/cshell/runs/README.md`. The `dshell` start
reads the discrete optimiser's result from `examples/dshell/runs/<name>` (`optimiser.dshell_run`).

## What runs

- Optimiser (`cshell/optimise.py`): MMA on per-vertex thickness/offset fields (filtered, D1 bounds), geometry chain
  ψ → χ → ρ on the GPU (Warp), continuation phase A (h_e 0.8 mm) → phase B (0.4 mm).
- FE (`cshell/fe_twogrid.py`, default `fe.engine = "twogrid"`): active-set q1_im hexes, matrix-free fine operator,
  Galerkin coarse level factorised by cuDSS, Chebyshev–Jacobi, fp64 PCG; buckling by block-Krylov; one GPU,
  ≤ 26 GiB. `fe.engine = "direct"` (cuDSS fp64) is kept for validation. Reproducibility: runs differ at ~10⁻⁴–10⁻³
  in final J (GPU atomic sums amplified by MMA); `fe.twogrid.deterministic = "full"` makes the FE bitwise (+150 % time).
- Collect (`cshell/collect.py`): merge starts, watertight export with mass correction and escape holes
  (`cshell/export.py`), re-score at the optimisation grid, a ½h_e-shifted grid and h_e 0.5 mm, body-fitted C3D10
  check (`cshell/validate.py`, `bodyfit_gpu.py`; isosurface-stuffed tets, quality gate, two levels), decision (§5.4).
  The body-fitted check runs where the model fits one GPU (`validation.max_bodyfit_dof`, 4·10⁶); larger models are
  decided on the voxel criteria and flagged "body-fitted check not run: model exceeds solver capacity".

## Costs (GH200, h_e 0.4 mm)
A start: 34–44 min (phase-B 25–50 s per iteration, 6–9 M DOF), ≤ 26 GiB GPU, ≤ 35 GB host. Collect: 30–40 min,
64–80 GiB host, ≤ 57 GiB GPU. Three parts × three starts ≈ 6 GPU-hours.

## Validation status and limits
See `docs/cshell/VALIDATION.md`. Known limits: body-fitted validation of full production parts needs a scalable
(iterative / multi-GPU) solver (Stage 3); V5-B1b stub closure is not reproduced under a compliance-only load
(documented); export mass residual up to +0.5 % before correction on passive mount regions (absorbed by the mass
correction).
