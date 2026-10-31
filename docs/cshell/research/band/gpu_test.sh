#!/bin/bash
# Band GPU tests / script. Usage: sbatch ...gpu_test.sh pytest [args] | python script.py [args]
#SBATCH --job-name=band_gpu
#SBATCH --gpus=1
#SBATCH --mem=64G
#SBATCH --time=00:40:00
#SBATCH --output=docs/cshell/research/band/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
nvidia-smi --query-gpu=name,memory.total --format=csv,noheader
if [ "${1:-pytest}" = pytest ]; then shift || true; python -m pytest -q tests/cshell/test_band.py --durations=12 "$@"; else [ "$1" = python ] && shift; python "$@"; fi
