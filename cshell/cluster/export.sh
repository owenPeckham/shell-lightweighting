#!/bin/bash
# Re-run finalise on a finished run (PLAN_single_run §3): shifted-grid re-score, GPU export at 0.2 mm, summary, report.
#   sbatch [--time=..] cshell/cluster/export.sh <run_dir> [--out DIR] [--h MM] [--no-rescore] [--sensitivity] [--bodyfit]
# One GPU. The export alone needs a few GiB of device and host memory; the re-score FE needs what the optimiser's
# final phase needs (hence the socket's default memory; with --no-rescore `--mem=16G` suffices).
#SBATCH --job-name=cshell_export
#SBATCH --gpus=1
#SBATCH --time=01:00:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
nvidia-smi --query-gpu=name,memory.total --format=csv,noheader
/usr/bin/time -v python -m cshell export "$@" 2> >(tee "$TMPDIR/time.txt" >&2)
grep -E "Maximum resident|Elapsed" "$TMPDIR/time.txt" || true
