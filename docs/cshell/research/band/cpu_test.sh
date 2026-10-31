#!/bin/bash
# Band CPU tests. Usage: sbatch docs/cshell/research/band/cpu_test.sh [pytest args]
#SBATCH --job-name=band_cpu
#SBATCH --cpus-per-task=72
#SBATCH --mem=16G
#SBATCH --time=00:25:00
#SBATCH --output=docs/cshell/research/band/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
export CUDA_VISIBLE_DEVICES=""
python -m pytest -q tests/cshell/test_band.py --durations=12 "$@"
