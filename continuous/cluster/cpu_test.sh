#!/bin/bash
# CPU-only pytest (R8). Usage: sbatch continuous/cluster/cpu_test.sh [pytest args]
#SBATCH --job-name=cshell_cpu_test
#SBATCH --cpus-per-task=72
#SBATCH --mem=16G
#SBATCH --time=00:30:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
export CUDA_VISIBLE_DEVICES=""
python -m pytest -q "${@:-tests/cshell}"
