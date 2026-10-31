#!/bin/bash
# Band CPU tests. Usage: sbatch continuous/research/band/cpu_test.sh [pytest args]
#SBATCH --job-name=band_cpu
#SBATCH --cpus-per-task=72
#SBATCH --mem=16G
#SBATCH --time=00:25:00
#SBATCH --output=continuous/research/band/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
export CUDA_VISIBLE_DEVICES=""
python -m pytest -q tests/cshell/test_band.py --durations=12 "$@"
