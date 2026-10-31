#!/bin/bash
#SBATCH --job-name=b2_mesh2
#SBATCH --gpus=1
#SBATCH --mem=8G
#SBATCH --time=00:50:00
#SBATCH --array=0-2,5%3
#SBATCH --output=docs/cshell/research/b2_bodyfit/work/%x_%A_%a.out
set -euo pipefail
source cshell/cluster/env.sh; trap 'rm -rf "$TMPDIR"' EXIT
C=(design design design solid solid solid); H=(0.7 0.5 0.35 0.7 0.5 0.35)
i=$SLURM_ARRAY_TASK_ID
X=""
python docs/cshell/research/b2_bodyfit/mesh2.py ${C[$i]} ${H[$i]} $X
