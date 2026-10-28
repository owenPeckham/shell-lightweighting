#!/bin/bash
# Usage: sbatch continuous/research/surface/run_py.sh script.py [args]   (CPU job, scratch work for Task A)
#SBATCH --job-name=surf_py
#SBATCH --cpus-per-task=72
#SBATCH --mem=16G
#SBATCH --time=00:30:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
export CUDA_VISIBLE_DEVICES=""
python "$@"
