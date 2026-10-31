#!/bin/bash
# GPU pytest / script. Usage: sbatch cshell/cluster/gpu_test.sh pytest [args] | python script.py [args]
#SBATCH --job-name=cshell_gpu_test
#SBATCH --gpus=1
#SBATCH --mem=32G
#SBATCH --time=00:30:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
nvidia-smi --query-gpu=name,memory.total --format=csv,noheader
if [ "${1:-pytest}" = pytest ]; then shift || true; python -m pytest -q "${@:-tests/cshell}"; else python "$@"; fi
