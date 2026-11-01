# cshell — continuous constant-mass shelling optimiser

`cshell` turns a solid part (STEP) into a hollowed, variable-thickness shell part of the **same mass** that is
stiffer and/or stronger. One run starts from the original solid and maximises the weighted objective

    J = w_K ln K/K0 + w_P ln P/P0        (w_K = w_P = 0.5)

(K stiffness, P capacity = soft-min of strength p-norm and global linear buckling, both relative to the solid),
subject to constant mass, floors K/K0 ≥ 0.85 and P/P0 ≥ 0.85, minimum wall t_min, minimum void v_min and as-built
biases. Spec: `docs/cshell/THEORY_continuous.md`; plan as implemented: `docs/cshell/plans/PLAN_single_run.md`; validation:
`docs/cshell/VALIDATION.md`; costs: `docs/cshell/BENCHMARKS.md`; state and known issues: `docs/cshell/STATUS.md`.

## Use (Isambard-AI; nothing on the login node)

1. Config: copy `examples/cshell/cantilever.json` — `name`, `step_file`, `output_dir`, `supports` (face selectors),
   `loads` (face selectors + `force_N`). Everything else defaults (`cshell/config.py`); override with `--set k=v`.
   `python -m cshell init` / `faces` print a template and the STEP faces (run via `gpu_test.sh`).
2. Run one part (one Slurm job, one GPU, auto-resumed across 24 h allocations; optimise then finalise in the same job):

       cshell/cluster/launch_part.sh examples/cshell/<part>.json [--set k=v ...]

   Same thing without the wrapper: `sbatch cshell/cluster/run.sh <cfg> [run args]`; CLI `python -m cshell run <cfg>
   [--resume DIR] [--max-hours H] [--no-finalise]`.
3. Re-run finalise on a finished run (shifted-grid re-score, 0.2 mm GPU export, summary, report; the run directory is
   left untouched with `--out`): `sbatch cshell/cluster/export.sh <run_dir> [--out DIR] [--h MM] [--sensitivity]`
   (CLI `python -m cshell export <run_dir>`). A report alone: `python -m cshell report <run_dir>` in a GPU job.
4. Results in `<output_dir>/<name>/`: `summary.txt` / `summary.json` (decision, J, K/K0, P/P0, governing mode, mass
   check, export checks), `final_design.stl` (multi-shell: sealed voids are separate shells), `final_design_export.json`,
   `wall_thickness.csv`, `fields.vtp` (ParaView), `report.pdf` + `plots/`, `history.csv`, `xi_solid_phase*.npy`,
   `design_solid_phase*.npz`, `finalise_log.txt`.

Committed results for the three example parts (reports, summaries, fields, preview STLs):
`examples/cshell/runs/<part>/`, described in `examples/cshell/runs/README.md`.

## Key config keys

| Key | Default | Meaning |
| --- | --- | --- |
| `objective.type` | `"weighted"` | `"weighted"`: J above with floors as MMA constraints; `"compliance"`: J = ln K/K0 with `objective.constraints` (`buckling`, `strength`) |
| `objective.w_stiffness`, `w_capacity` | 0.5, 0.5 | Weights w_K, w_P |
| `objective.floor_stiffness`, `floor_capacity` | 0.85, 0.85 | K/K0 and P/P0 floors |
| `optimiser.starts` | `["solid"]` | Start designs (`uniform`, `dshell`, `neutral` remain available) |
| `growth.enabled`, `growth.quantiles`, `growth.patience` | on, 0.1…1.0, `None` | The hollow-allowed region grows in order of the shear/flexure ratio R through every quantile (patience `None`: all quantiles run) |
| `design.closure_skirt_mm` | 1.5 | Closure spreads over this skirt at closure borders (0 = off) |
| `export.h_mm` | 0.2 | GPU streaming export resolution |
| `export.escape_holes` | `false` | Escape (powder-removal) holes are optional; off: enclosed voids stay sealed |
| `manufacturing.t_min_mm`, `void_min_mm` | 1.5, 3.0 | Minimum wall and void width |

## What runs

- Optimiser (`cshell/optimise.py`): MMA on per-vertex thickness/offset fields (smoothed with an F² filter, smoothed
  r_MA, taper, closure skirt), geometry chain ψ → χ → ρ on the GPU (Warp), growth of the hollow-allowed region,
  continuation phases A (h_e 0.8 mm) → B (0.4 mm).
- FE (`cshell/fe_twogrid.py`, default `fe.engine = "twogrid"`): active-set q1_im hexes, matrix-free fine operator,
  Galerkin coarse level factorised by cuDSS, Chebyshev–Jacobi, fp64 PCG; buckling by block-Krylov; one GPU.
  `fe.engine = "direct"` (cuDSS fp64) is kept for validation. Runs differ at ~10⁻⁴–10⁻³ in final J (GPU atomic sums).
- Finalise (`cshell/finalise.py`, called by `run`/`export`): re-score on a ½h_e-shifted grid, streaming 0.2 mm GPU
  export (`cshell/export.py`) with mass correction, summary, report (`cshell/report.py`). There is no collect stage.
  The legacy multi-start `cshell collect` remains for Stage 2 run directories.

## Costs (GH200)
A full part is one GPU job: round-4 validation (`runs_cshell/SR_VAL4`) used ≤ 36 GiB GPU and ≤ 5.5 GiB host for
finalise (re-score 155–215 s, export 19–36 s). Per-iteration costs: `docs/cshell/BENCHMARKS.md`.

## Known limits
See `docs/cshell/STATUS.md`: residual export fins (54–126 mm³, ≈ 0.1 % V); body-fitted validation not run at
production scale (needs a scalable solver); the export mass residual is absorbed by the mass correction.
