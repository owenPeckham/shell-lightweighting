#!/bin/bash
# V3-b body-fitted side: ccx *BUCKLE on C3D10, mesh layers through t. sbatch continuous/research/validate/v3b_ccx.sh
#SBATCH --job-name=v3b_ccx
#SBATCH --cpus-per-task=72
#SBATCH --mem=160G   # layers 5 needs more (OOM at 160G)
#SBATCH --time=04:00:00
#SBATCH --output=logs/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
export V3B_LOAD=${V3B_LOAD:-area}
export CUDA_VISIBLE_DEVICES="" CSHELL_WARP_DEVICE=cpu
python continuous/research/validate/v3b_run.py ccx continuous/research/validate/v3b_ccx_${V3B_LOAD:-area}.json ${@:-2 3 4 5}
