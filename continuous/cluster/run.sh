#!/bin/bash
# cshell run. Usage: sbatch [--time=..] continuous/cluster/run.sh <cfg.json> [cshell run args]
# Resumable: resubmit with --resume <run_dir> (THEORY §7.6).
#SBATCH --job-name=cshell_run
#SBATCH --gpus=1
#SBATCH --time=02:00:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
nvidia-smi --query-gpu=name,memory.total --format=csv,noheader
python -m cshell run "$@"
