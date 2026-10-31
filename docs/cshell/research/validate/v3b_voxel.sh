#!/bin/bash
# V3-b voxel side: cshell sparse engine, h_e = 0.8, 0.4, 0.2 (cuDSS hybrid memory at 0.2).
#SBATCH --job-name=v3b_voxel
#SBATCH --gpus=1
#SBATCH --mem=220G
#SBATCH --time=01:30:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
python docs/cshell/research/validate/v3b_run.py voxel docs/cshell/research/validate/v3b_voxel${V3B_TAG:-}.json ${@:-0.8 0.4 0.2}
