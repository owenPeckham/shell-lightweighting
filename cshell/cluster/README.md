# cshell/cluster: Slurm scripts for cshell (Isambard-AI, GH200)

Submit from the repository or worktree root; `env.sh` activates the conda env, puts `TMPDIR` on scratch and `cd`s to
`$SLURM_SUBMIT_DIR`, so a job runs the checkout it was submitted from. Logs go to `logs/`.

| Script | Use |
|---|---|
| `cpu_test.sh [pytest args]` | CPU pytest (72 cores, 16 GB) |
| `gpu_test.sh pytest [args]` / `gpu_test.sh script.py` / `gpu_test.sh -m cshell ...` | GPU tests and short runs |
| `run.sh <cfg> [run args]` | one part: optimise, then finalise in the same job (`cshell run`, PLAN_single_run §3); `--max-hours` = time limit − 0.4 h; on pause (exit 3) or SIGTERM it submits its own continuation (`--resume`, `afterany`), one queue slot per part |
| `export.sh <run_dir> [--out DIR] [--h MM] [--no-rescore] [--sensitivity] [--bodyfit]` | re-run finalise on a finished run (`cshell export`): shifted-grid re-score, GPU export (0.2 mm), summary.json/txt, report |
| `launch_part.sh <cfg> [run args]` | one 24 h `run.sh` for the part (no per-start fan-out, no collect job) |

Full part, e.g. the cantilever (1 slot):

    cshell/cluster/launch_part.sh examples/cshell/cantilever.json

Re-export a finished run into another directory (the run directory is left untouched):

    sbatch cshell/cluster/export.sh runs_cshell/FULL/MBB --out runs_cshell/MBB_reexport

Finalise outputs (contract C3): `final_design.stl`, `final_design_export.json`, `escape_holes.csv`,
`wall_thickness.csv`, `summary.json`, `summary.txt`, `finalise_log.txt`, then the report (`report.write_report`; a
report failure is logged in summary.json, it does not fail the job). The legacy multi-start `cshell collect` is kept
for Stage 2 run directories but has no cluster script any more.

Environment knobs: `CSHELL_TIME` (launch, default 24:00:00), `CSHELL_MARGIN_H` (0.4), `CSHELL_CHAIN_MAX` (8
allocations per part).
