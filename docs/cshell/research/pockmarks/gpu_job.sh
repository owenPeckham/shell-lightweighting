#!/bin/bash
# Throwaway diag job (GPU). Usage: sbatch docs/cshell/research/pockmarks/gpu_job.sh script.py args...
#SBATCH --job-name=pockmarks_gpu
#SBATCH --gpus=1
#SBATCH --mem=64G
#SBATCH --time=01:00:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
nvidia-smi --query-gpu=name,memory.total --format=csv,noheader
python "$@"
