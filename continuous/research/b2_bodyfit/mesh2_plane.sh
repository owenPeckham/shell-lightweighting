#!/bin/bash
#SBATCH --job-name=b2_plane
#SBATCH --gpus=1
#SBATCH --mem=16G
#SBATCH --time=00:20:00
#SBATCH --array=0-1%2
#SBATCH --output=continuous/research/b2_bodyfit/work/%x_%A_%a.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
H=(0.5 0.35)
python continuous/research/b2_bodyfit/mesh2.py design ${H[$SLURM_ARRAY_TASK_ID]} --clamp-plane --tag _plane
