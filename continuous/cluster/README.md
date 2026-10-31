# continuous/cluster: Slurm scripts for cshell (Isambard-AI, GH200)

Submit from the repository or worktree root; `env.sh` activates the conda env, puts `TMPDIR` on scratch and `cd`s to
`$SLURM_SUBMIT_DIR`, so a job runs the checkout it was submitted from. Logs go to `logs/`.

| Script | Use |
|---|---|
| `cpu_test.sh [pytest args]` | CPU pytest (72 cores, 16 GB; add `--mem=48G` for collect tests) |
| `gpu_test.sh pytest [args]` / `gpu_test.sh script.py` / `gpu_test.sh -m cshell ...` | GPU tests and short runs |
| `run.sh <cfg> --start S [run args]` | one start of one part; `--max-hours` = time limit − 0.4 h; on pause (exit 3) or SIGTERM it submits its own continuation (`--resume`, `afterany`), so one queue slot per start; the last start to finish submits `collect.sh` once |
| `collect.sh <run_dir>` | `python -m cshell collect` then `report` (GPU, 4 h default) |
| `launch_part.sh <cfg> [starts] [-- run args]` | one 24 h `run.sh` per start (default: the config's starts) |

Full part, e.g. the cantilever with three starts (3 slots, then 1 for collect):

    continuous/cluster/launch_part.sh examples/cshell/cantilever.json

Measured resources (2026-10-31, h_e 0.4 / h_g 0.1): a start runs 25–50 s per phase-B iteration, 34–44 min in all,
≤ 26 GiB GPU and ≤ 35 GB host — submit with `SBATCH_MEM_PER_NODE=64G CSHELL_TIME=04:00:00`; collect needs 64–80 GiB
host and ≤ 57 GiB GPU, 30–40 min — `CSHELL_COLLECT_MEM=180G` (run.sh passes it to the collect it submits).

Environment knobs: `CSHELL_COLLECT_MEM` (collect's --mem; default the socket's), `CSHELL_TIME` (launch, default 24:00:00), `CSHELL_MARGIN_H` (0.4), `CSHELL_CHAIN_MAX` (8
allocations per start), `CSHELL_AUTO_COLLECT` (1).
