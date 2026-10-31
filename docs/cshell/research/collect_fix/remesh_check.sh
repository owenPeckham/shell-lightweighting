#!/bin/bash
# sbatch continuous/research/collect_fix/remesh_check.sh   (from the repo root)
#SBATCH --job-name=cf_remesh
#SBATCH --cpus-per-task=16
#SBATCH --mem=48G
#SBATCH --time=00:45:00
#SBATCH --output=continuous/research/collect_fix/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
export CUDA_VISIBLE_DEVICES=""
python continuous/research/collect_fix/remesh_check.py
