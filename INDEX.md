# INDEX: map of the repository

Constant-mass structural shelling: a solid part becomes a hollow shell part of the same mass that is stiffer and/or
stronger. Three optimisers, from simplest to most general:

| Optimiser | Code | Design variables | Docs | Example results |
| --- | --- | --- | --- | --- |
| **ushell** (uniform shell) | `ushell/` (Fusion 360 add-in) | one wall thickness | `README.md` (install) | `ushell/logs/` (example log) |
| **dshell** (discrete regions; was `vshell`) | `dshell/` | which regions are hollow + their wall | `dshell/README.md` | `examples/dshell/runs/<part>/` |
| **cshell** (continuous) | `cshell/` | per-vertex wall thickness and offset fields | `cshell/README.md`, `docs/cshell/` | `examples/cshell/runs/<part>/` |

Renamed 2026-10-31: `vshell` → `dshell`, `continuous/` → `docs/cshell/`, `shell_lightweighting/` → `ushell/`,
`examples/step/` → `examples/parts/`, `examples/runs/` → `examples/dshell/runs/`. Older documents keep the old
names in their text.

## Top level

| Path | Contents |
| --- | --- |
| `README.md` | Overview, licence disclaimer, how to install the Fusion add-in |
| `INDEX.md` | This map |
| `LICENSE` | Licence |
| `ushell/` | Fusion 360 add-in: `ShellOptimisation.py`, `ShellOptimisation.manifest` (Fusion keys on these names), `resources/`, `logs/` |
| `dshell/` | Discrete-region optimiser (Python package): `README.md`, `requirements.txt`, `cli.py`, `config.py`, `geometry.py` (STEP, tessellation), `segment.py`, `field.py`, `fe.py` (voxel FE), `mechanics.py`, `optimise.py`, `report.py`, `util.py` |
| `cshell/` | Continuous optimiser (Python package): `README.md`, `cluster/` (Slurm scripts), `cli.py`, `config.py`, `surface.py`, `band.py`, `design.py`, `objective.py`, `fe_twogrid.py` / `fe_sparse.py` / `fea.py` (FE), `mma.py`, `optimise.py`, `collect.py`, `export.py`, `validate.py` / `bodyfit_gpu.py`, `report.py` |
| `examples/parts/` | The three Study 1 STEP parts: `Solid MBB.step`, `Solid Cantilever.step`, `Solid L-bracket.step` (shared by all optimisers) |
| `examples/dshell/` | dshell configs (`mbb.json`, `cantilever.json`, `l_bracket.json`) and `runs/<part>/` results |
| `examples/cshell/` | cshell configs (`mbb.json`, `cantilever.json`, `l_bracket.json`), `bench/` (B1–B3 benchmark parts + configs, `make_bench.py`), `runs/<part>/` results |
| `docs/cshell/` | cshell theory, status, validation, benchmarks, plans, research notes and Owen's reference FE code (details below) |
| `feasibility/` | Feasibility prototype for variable thickness (scripts, results JSON, figures; `feasibility/README.md`) |
| `tests/dshell/` | dshell tests (scripts: `test_basic.py`, `validate_tubes.py`, `validate_thin_wall.py`) |
| `tests/cshell/` | cshell pytest suite (CPU by default, `@pytest.mark.gpu` tests need CUDA) |

## Results and reports, per optimiser and part

| Part | dshell (`examples/dshell/runs/`) | cshell (`examples/cshell/runs/`) |
| --- | --- | --- |
| MBB | `MBB/report.pdf`, `summary.txt`: solid kept (J +0.0003) | `MBB/report.pdf`, `summary.txt`: accepted, J +0.482, K/K0 1.10, P/P0 2.39 |
| Cantilever | `Cantilever/report.pdf`: hollow design accepted (J +0.479) | `Cantilever/report.pdf`: accepted, J +0.490, K/K0 0.98, P/P0 2.71 |
| L-bracket | `L-bracket/report.pdf`: solid kept | `L-bracket/report.pdf`: accepted, J +0.129, K/K0 1.16, P/P0 1.12 |

- **dshell runs** (`examples/dshell/runs/<part>/`, from before the rename, so the texts say `vshell`): `report.pdf`,
  `summary.json/.txt`, `run_log.txt`, `run_info.json`, `config_used.json`, `regions.csv`, `evaluations.csv`,
  `mass_sweep.csv`, `escape_holes.csv`, `sensitivity.csv`, `arrays.npz` (segmentation; the cshell `dshell` start reads
  it). No STLs. 2–3 MB per part.
- **cshell runs** (`examples/cshell/runs/<part>/`): report, summaries, histories, fields (`fields.vtp` for ParaView),
  buckling modes, final design fields `xi_<start>_phase*.npy`, plots, and decimated preview STLs
  (`final_design_preview.stl`, `original_preview.stl`, `uniform_shell_preview.stl`). File list, the full-resolution
  data location on Isambard and how to regenerate: `examples/cshell/runs/README.md`. J = w_K ln K/K0 + w_P ln P/P0
  at constant mass, re-scored on the exported design; full comparison in `docs/cshell/VALIDATION.md` V6 and
  `docs/cshell/STATUS.md` §5.

## How to run

- **ushell**: install in Fusion 360 (see `README.md`), run from Utilities → Add-Ins.
- **dshell** (any PC): `pip install -r dshell/requirements.txt`, then from the repository root
  `python -m dshell run examples/dshell/l_bracket.json` (details `dshell/README.md`).
- **cshell** (Isambard-AI, one GH200 per start; never on the login node): environment in `cshell/cluster/env.sh`
  (conda env from `docs/cshell/environment.yml`); full part:
  `cshell/cluster/launch_part.sh examples/cshell/cantilever.json` (three starts, then collect + report);
  `python -m cshell --help` for `setup`, `run`, `collect`, `report`, `faces`, `init`. Outputs go to
  `runs_cshell/<name>/` (a symlink at the repository root to project storage; configs' `output_dir`). Details:
  `cshell/README.md`, `cshell/cluster/README.md`.

## Tests

- dshell: `python tests/dshell/test_basic.py` (~1 min), `tests/dshell/validate_tubes.py`, `tests/dshell/validate_thin_wall.py`.
- cshell: `sbatch cshell/cluster/cpu_test.sh` (CPU) and `sbatch cshell/cluster/gpu_test.sh pytest tests/cshell` (GPU).

## Documents (`docs/cshell/`)

| File | Contents |
| --- | --- |
| `THEORY_continuous.md` | cshell specification: geometry, mechanics, sensitivities, optimiser, export, validation, outputs |
| `STATUS.md` | Log, plan, risks, per-part results (§5) |
| `VALIDATION.md` | Validation record V1–V7 (V6: the three parts vs solid / uniform shell / dshell / manuscript) |
| `BENCHMARKS.md` | Costs and timings |
| `HANDOVER.md` | Hand-over at the end of Stage 2, Stage 3 candidates |
| `OPEN_QUESTIONS.md` | Questions and decisions |
| `PROMPT_*.md`, `README.md` | Session prompts and the hand-over pack description |
| `plans/` | Agreed plans per work package (`PLAN_*.md`, incl. this restructure) |
| `research/` | Studies behind the decisions (scripts, job logs, `REPORT.md` per study) |
| `reference/` | Owen's FE stack used by cshell (`analysis.py`, `gpu_analysis.py`), `baseline_optimiser.py`, manuscript `main.pdf`, exports of the vshell theory doc and the feasibility study |
| `environment.yml` | cshell conda environment (aarch64) |
