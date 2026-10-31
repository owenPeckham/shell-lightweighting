#!/bin/bash
# Collect re-scoring diagnosis (CPU vs GPU chain, uniform shell BCs, shifted grid). Submit from the repo root.
#SBATCH --job-name=cm_diag
#SBATCH --gpus=1
#SBATCH --mem=160G
#SBATCH --time=01:30:00
#SBATCH --output=continuous/research/collect_mem/%x_%j.out
set -euo pipefail
source continuous/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
OUT=runs_cshell/CM_diag_${SLURM_JOB_ID}; mkdir -p $OUT
/usr/bin/time -v python continuous/research/collect_mem/diag.py runs_cshell/SMOKE_mbb/MBB $OUT --shift
