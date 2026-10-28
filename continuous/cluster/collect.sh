#!/bin/bash
# cshell collect (THEORY §5.4, §8): sbatch [--time=..] continuous/cluster/collect.sh <run_dir> [collect args]
# Runs re-scoring FE on the GPU and the export / CalculiX check on the socket's 72 cores.
#SBATCH --job-name=cshell_collect
#SBATCH --gpus=1
#SBATCH --time=04:00:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
nvidia-smi --query-gpu=name,memory.total --format=csv,noheader
python -m cshell collect "$@"
python -m cshell report "$1" || echo "[collect.sh] report failed (see plots/report_log.txt)"
